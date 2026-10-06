'    Copyright (C) 2018-2026 Hazel Ward
'
'    This file is a part of Winapp2ool
'
'    Winapp2ool is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    Winapp2ool is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with Winapp2ool.  If not, see <http://www.gnu.org/licenses/>.

Option Strict On

Imports System.Threading.Tasks

''' <summary>
''' Categorizes entries as added, removed, or modified between two versions of winapp2.ini.
''' Entry names compare ignoring case, so an entry renamed only in case counts as present in
''' both files. Normalizes deprecated values to suppress false-positive diffs, delegates rename
''' and merger detection to <see cref="MergeDetector"/>, and coordinates key-level analysis
''' via <see cref="KeyModificationAnalyzer"/>.
''' </summary>
Public Class EntryChangeDetector

    Private ReadOnly _state As DiffState
    Private ReadOnly _file1 As iniFile
    Private ReadOnly _file2 As iniFile
    Private ReadOnly _mergeDetector As MergeDetector
    Private ReadOnly _keyAnalyzer As KeyModificationAnalyzer
    Private ReadOnly _renderer As DiffOutputRenderer

    ''' <summary>Creates a new <c> EntryChangeDetector </c></summary>
    ''' 
    ''' <param name="state">
    ''' Shared diff state tracking all entry changes
    ''' </param>
    ''' 
    ''' <param name="file1">
    ''' The old version of winapp2.ini as an <c> iniFile </c>
    ''' </param>
    ''' 
    ''' <param name="file2">
    ''' The new version of winapp2.ini as an <c> iniFile </c>
    ''' </param>
    ''' 
    ''' <param name="mergeDetector">
    ''' Handles rename and merger detection for removed entries
    ''' </param>
    '''
    ''' <param name="keyAnalyzer">
    ''' Tracks key-level changes between entry versions
    ''' </param>
    '''
    ''' <param name="renderer">
    ''' Produces <c> MenuSection </c> output for entries removed without replacement
    ''' </param>
    Public Sub New(state As DiffState,
                   file1 As iniFile,
                   file2 As iniFile,
                   mergeDetector As MergeDetector,
                   keyAnalyzer As KeyModificationAnalyzer,
                   renderer As DiffOutputRenderer)

        _state = state
        _file1 = file1
        _file2 = file2
        _mergeDetector = mergeDetector
        _keyAnalyzer = keyAnalyzer
        _renderer = renderer

    End Sub

    ''' <summary>
    ''' Replaces each deprecated value listed in <c> PathReplacements </c> (environment variables,
    ''' and <c> *.* </c> becoming <c> * </c>) in every key of <paramref name="winapp"/>, so a change
    ''' to the newer spelling doesn't show up as a diff. The match is case-sensitive. We change the
    ''' keys in place, so everything that reads the file afterward, including the rendered
    ''' output, sees the replaced values.
    ''' </summary>
    '''
    ''' <param name="winapp">
    ''' The file whose key values will be normalized in place
    ''' </param>
    Public Sub SnuffNoisyChanges(winapp As iniFile)

        For Each section In winapp

            For Each key In section.Keys : CleanKeyValue(key) : Next

        Next

    End Sub

    ''' <summary>
    ''' Applies each <c> PathReplacements </c> pair to <paramref name="key"/>'s value in turn,
    ''' with a case-sensitive <c> Replace </c>, so a later pair sees the output of earlier ones
    ''' </summary>
    '''
    ''' <param name="key">
    ''' The key whose value is normalized against <c> PathReplacements </c>
    ''' </param>
    Private Sub CleanKeyValue(key As iniKey)

        For i = 0 To PathReplacements.Count - 1

            Dim oldPath = PathReplacements.Keys(i)

            If Not key.Value.Contains(oldPath) Then Continue For

            key.Value = key.Value.Replace(oldPath, PathReplacements.Values(i))

        Next

    End Sub

    ''' <summary>
    ''' Records each entry whose name (ignoring case) isn't in the old file as added, and adds
    ''' its section to <c> PotentialMatches </c> as a rename or merger candidate
    ''' </summary>
    Public Sub ProcessNewEntries()

        For Each section In _file2

            If _file1.Contains(section.Name) Then Continue For

            _state.ModifiedEntries.AddedEntryNames.Add(section.Name)
            _state.ModifiedEntries.PotentialMatches.Add(section)

        Next

    End Sub

    ''' <summary>
    ''' Records each old entry whose name (ignoring case) isn't in the new file as removed, and
    ''' compares every other old entry against its new version through
    ''' <see cref="KeyModificationAnalyzer.FindModifications"/>. Then adds the new version of each
    ''' entry that comparison marked as modified to <c> PotentialMatches </c>.
    ''' </summary>
    Public Sub ProcessOldEntries()

        For Each section In _file1

            If Not _file2.Contains(section.Name) Then

                _state.ModifiedEntries.RemovedEntryNames.Add(section.Name)
                Continue For

            End If

            _keyAnalyzer.FindModifications(section, _file2.GetSection(section.Name))

        Next

        For Each modifiedEntryName In _state.ModifiedEntries.ModifiedEntryNames
            _state.ModifiedEntries.PotentialMatches.Add(_file2.GetSection(modifiedEntryName))
        Next

    End Sub

    ''' <summary>
    ''' Sorts the removed entries, in parallel, into three bins. Only <c> FileKey </c> and
    ''' <c> RegKey </c> values decide the bin; other key types can differ freely. <br /><br />
    '''
    ''' Renamed: an added entry matches every old FileKey and RegKey, has the same number of each,
    ''' and raised neither the more-patterns nor the wildcard-reduction flag. Each pattern-by-pattern
    ''' FileKey match overwrites those flags, so they reflect only the last such match, and key
    ''' order can decide between a rename and a merger (see <see cref="MergeDetector.CountMatches"/>). <br />
    ''' Merged: at least one candidate matches at least one old FileKey or RegKey. <br />
    ''' Removed without replacement: the entry has no FileKey or RegKey, or none of the candidates
    ''' we found by name and content matched any of them. <br /><br />
    '''
    ''' Then converts any rename that a parallel merger claimed (see
    ''' <see cref="ReconcileRenamesAndMergers"/>).
    ''' </summary>
    '''
    ''' <returns>
    ''' A header with the removal counts, then one section per entry removed without replacement,
    ''' ordered by name ignoring case. Empty when no entries were removed.
    ''' </returns>
    Public Function ProcessRemovals() As List(Of MenuSection)

        Dim totalRemoved = _state.ModifiedEntries.RemovedEntryNames.Count
        Dim removedHeader = $"{totalRemoved} total entries removed"

        Dim out As New List(Of MenuSection)

        gLog(Nothing, leadr:=True)

        Dim noReplacementCount = 0

        Using gLogScope("Entry removals:")

            Dim results = New Concurrent.ConcurrentDictionary(Of String, MenuSection)(StringComparer.OrdinalIgnoreCase)
            Dim entryLogs = New Concurrent.ConcurrentDictionary(Of String, List(Of String))(StringComparer.OrdinalIgnoreCase)
            Dim potentialMatchesSnapshot = _state.ModifiedEntries.PotentialMatches.ToList()

            Dim snapshotTextMap As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            Dim oldEntryTextMap As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            PrePopulateCachesAndTextMaps(potentialMatchesSnapshot, snapshotTextMap, oldEntryTextMap)

            Dim contentIndexes = BuildContentIndexes(potentialMatchesSnapshot)
            Dim eligibleNames = BuildEligibleNameSet(potentialMatchesSnapshot)

            Parallel.ForEach(_state.ModifiedEntries.RemovedEntryNames,
                     Sub(entry)

                         Dim result As MenuSection = Nothing
                         Dim capturedLines As New List(Of String)

                         Using cap = gLogCapture()

                             result = ProcessSingleRemoval(entry, potentialMatchesSnapshot, snapshotTextMap, oldEntryTextMap, contentIndexes, eligibleNames)
                             capturedLines.AddRange(cap.Lines)

                         End Using

                         If result IsNot Nothing Then results(entry) = result
                         If capturedLines.Count > 0 Then entryLogs(entry) = capturedLines

                     End Sub)

            ReconcileRenamesAndMergers()

            Dim renamedCount = _state.MergedEntries.RenamedEntryNames.Count
            Dim mergedCount = _state.MergedEntries.OldToNewMergeDict.Count
            noReplacementCount = totalRemoved - renamedCount - mergedCount

            ' Nothing removed means nothing to announce: an "0 total entries removed" banner is
            ' noise the Diff Summary already covers. The global log still records the count below,
            ' so the saved changelog's shape is unchanged
            If totalRemoved > 0 Then

                Dim header As New MenuSection
                header.AddColoredLine(removedHeader, ConsoleColor.DarkRed, True)
                header.AddBlank()
                header.AddColoredLine($"  - {noReplacementCount} entries removed without replacement", ConsoleColor.Red, True)
                header.AddDivider(solid:=False)
                out.Add(header)

            End If

            For Each key In results.Keys.OrderBy(Function(k) k, StringComparer.OrdinalIgnoreCase)

                Dim logLines As List(Of String) = Nothing
                If entryLogs.TryGetValue(key, logLines) Then EmitCaptured(logLines)
                out.Add(results(key))

            Next

        End Using

        gLog($"- {noReplacementCount} removed without replacement", leadr:=True)

        Return out

    End Function

    ''' <summary>
    ''' Populates the new/old entry caches and pre-computes uppercased section text
    ''' for each potential match and removed entry
    ''' </summary>
    '''
    ''' <param name="potentialMatches">
    ''' Snapshot of added and modified entries to cache
    ''' </param>
    '''
    ''' <param name="snapshotTextMap">
    ''' Populated with uppercased text for each potential match, keyed by section name
    ''' </param>
    '''
    ''' <param name="oldEntryTextMap">
    ''' Populated with uppercased text for each removed entry, keyed by entry name
    ''' </param>
    Private Sub PrePopulateCachesAndTextMaps(potentialMatches As List(Of iniSection),
                                             snapshotTextMap As Dictionary(Of String, String),
                                             oldEntryTextMap As Dictionary(Of String, String))

        For Each section In potentialMatches

            _state.Caches.CachedNewEntries(section.Name) = section
            snapshotTextMap(section.Name) = section.ToString().ToUpperInvariant()

        Next

        For Each entryName In _state.ModifiedEntries.RemovedEntryNames

            Dim oldSection = _file1.GetSection(entryName)
            _state.Caches.CachedOldEntries(oldSection.Name) = oldSection
            oldEntryTextMap(entryName) = oldSection.ToString().ToUpperInvariant()

        Next

    End Sub

    ''' <summary>
    ''' Builds reverse indexes over the FileKey and RegKey values in <paramref name="potentialMatches"/>
    ''' for content-aware candidate lookup during removal processing. All three compare ignoring case.
    ''' <list type="bullet">
    '''   <item><term>KeyValueIndex</term>
    '''     <description>Whole key value → set of section names containing that value</description></item>
    '''   <item><term>PathRootIndex</term>
    '''     <description>Path root from <see cref="GetPathRoot"/> → set of section names, catching
    '''     pattern or flag changes where the path root stays the same</description></item>
    '''   <item><term>WildcardPrefixes</term>
    '''     <description>For roots with a <c> * </c> after the first character, the prefix before
    '''     the <c> * </c>, grouped by first path component</description></item>
    ''' </list>
    ''' </summary>
    '''
    ''' <param name="potentialMatches">
    ''' Snapshot of added and modified entries whose keys are indexed
    ''' </param>
    '''
    ''' <returns>
    ''' A <c> ContentIndexes </c> instance containing all three reverse indexes
    ''' </returns>
    Private Shared Function BuildContentIndexes(potentialMatches As List(Of iniSection)) As ContentIndexes

        Dim indexes As New ContentIndexes()

        For Each section In potentialMatches

            For Each key In section.Keys

                If Not key.KeyType.Equals("FileKey", StringComparison.OrdinalIgnoreCase) AndAlso
                   Not key.KeyType.Equals("RegKey", StringComparison.OrdinalIgnoreCase) Then Continue For

                ' Index exact value
                If Not indexes.KeyValueIndex.ContainsKey(key.Value) Then indexes.KeyValueIndex(key.Value) = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                indexes.KeyValueIndex(key.Value).Add(section.Name)

                ' Index path root (first two backslash components) for wildcard resilience
                Dim root = GetPathRoot(key.Value)

                If root Is Nothing Then Continue For

                If Not indexes.PathRootIndex.ContainsKey(root) Then indexes.PathRootIndex(root) = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                indexes.PathRootIndex(root).Add(section.Name)

                ' If root has wildcard, also index the prefix before * grouped by first component
                Dim starIdx = root.IndexOf("*"c)

                If Not starIdx > 0 Then Continue For

                Dim prefix = root.Substring(0, starIdx)
                Dim firstComp = GetFirstComponent(root)

                If firstComp IsNot Nothing Then

                    If Not indexes.WildcardPrefixes.ContainsKey(firstComp) Then indexes.WildcardPrefixes(firstComp) = New List(Of KeyValuePair(Of String, String))
                    indexes.WildcardPrefixes(firstComp).Add(New KeyValuePair(Of String, String)(prefix, section.Name))

                End If

            Next

        Next

        Return indexes

    End Function

    ''' <summary>
    ''' Returns the set of section names from <paramref name="potentialMatches"/> that are
    ''' eligible for candidacy (i.e. present in <c> AddedEntryNames </c> or <c> ModifiedEntryNames </c>)
    ''' </summary>
    '''
    ''' <param name="potentialMatches">
    ''' Snapshot of added and modified entries to filter
    ''' </param>
    '''
    ''' <returns>
    ''' A <c> HashSet </c> of section names eligible for rename/merger matching
    ''' </returns>
    Private Function BuildEligibleNameSet(potentialMatches As List(Of iniSection)) As HashSet(Of String)

        Dim eligibleNames As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        For Each section In potentialMatches

            If _state.ModifiedEntries.AddedEntryNames.Contains(section.Name) OrElse
               _state.ModifiedEntries.ModifiedEntryNames.Contains(section.Name) Then eligibleNames.Add(section.Name)

        Next

        Return eligibleNames

    End Function

    ''' <summary>
    ''' Processes a single removed entry: gathers rename/merger candidates from name heuristics
    ''' and content-aware index lookups, filters to eligible entries, and delegates to
    ''' <see cref="MergeDetector.AssessRenamesAndMergers"/>. An entry with no FileKey or RegKey
    ''' skips the search and counts as removed without replacement.
    ''' </summary>
    '''
    ''' <param name="entryName">
    ''' The name of the removed entry being processed
    ''' </param>
    '''
    ''' <param name="potentialMatches">
    ''' Snapshot of added and modified entries for name-based heuristic matching
    ''' </param>
    '''
    ''' <param name="snapshotTextMap">
    ''' Pre-computed uppercased text for each potential match section
    ''' </param>
    '''
    ''' <param name="oldEntryTextMap">
    ''' Pre-computed uppercased text for each removed entry
    ''' </param>
    '''
    ''' <param name="indexes">
    ''' Reverse content indexes built by <see cref="BuildContentIndexes"/>
    ''' </param>
    '''
    ''' <param name="eligibleNames">
    ''' Set of section names eligible for candidacy (added or modified)
    ''' </param>
    '''
    ''' <returns>
    ''' A <c> MenuSection </c> describing the removal if no rename/merger was found;
    ''' <c> Nothing </c> if a rename or merger was recorded in <c> DiffState </c>
    ''' </returns>
    Private Function ProcessSingleRemoval(entryName As String,
                                           potentialMatches As List(Of iniSection),
                                           snapshotTextMap As Dictionary(Of String, String),
                                           oldEntryTextMap As Dictionary(Of String, String),
                                           indexes As ContentIndexes,
                                           eligibleNames As HashSet(Of String)) As MenuSection

        Dim oldSection = _file1.GetSection(entryName)

        If oldSection.Keys.GetByType("FileKey").Count = 0 AndAlso
           oldSection.Keys.GetByType("RegKey").Count = 0 Then

            Return _renderer.MakeDiff(oldSection, 1)

        End If

        Dim allCandidates = GatherCandidateNames(entryName, oldSection, potentialMatches, snapshotTextMap, oldEntryTextMap, indexes)
        Dim combinedMatches = FilterToEligibleSections(allCandidates, eligibleNames)
        Dim changesRecorded = _mergeDetector.AssessRenamesAndMergers(combinedMatches, oldSection)

        Return If(changesRecorded, Nothing, _renderer.MakeDiff(oldSection, 1))

    End Function

    ''' <summary>
    ''' Gathers candidate section names for a removed entry using four heuristics: name and
    ''' browser LangSecRef matching (<see cref="FindProbableMatches"/>), then, for each old FileKey
    ''' and RegKey, a whole-value lookup, a path root lookup, and a wildcard prefix lookup. The
    ''' wildcard lookup adds a candidate when the old key's root, cut at its first <c> * </c>,
    ''' starts with a new key's wildcard prefix under the same first component (ignoring case).
    ''' </summary>
    '''
    ''' <param name="entryName">
    ''' The name of the removed entry
    ''' </param>
    '''
    ''' <param name="oldSection">
    ''' The removed entry's section from the old file
    ''' </param>
    '''
    ''' <param name="potentialMatches">
    ''' Snapshot of added and modified entries for name-based heuristic matching
    ''' </param>
    '''
    ''' <param name="snapshotTextMap">
    ''' Pre-computed uppercased text for each potential match section
    ''' </param>
    '''
    ''' <param name="oldEntryTextMap">
    ''' Pre-computed uppercased text for each removed entry
    ''' </param>
    '''
    ''' <param name="indexes">
    ''' Reverse content indexes for value, path root, and wildcard prefix lookups
    ''' </param>
    '''
    ''' <returns>
    ''' A set of all candidate section names found across all heuristics
    ''' </returns>
    Private Function GatherCandidateNames(entryName As String,
                                           oldSection As iniSection,
                                           potentialMatches As List(Of iniSection),
                                           snapshotTextMap As Dictionary(Of String, String),
                                           oldEntryTextMap As Dictionary(Of String, String),
                                           indexes As ContentIndexes) As HashSet(Of String)

        Dim allCandidates As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        Dim probableMatches = FindProbableMatches(entryName.Split(CChar(" ")), potentialMatches, snapshotTextMap, oldEntryTextMap(entryName))
        For Each section In probableMatches : allCandidates.Add(section.Name) : Next

        For Each key In oldSection.Keys

            If Not key.KeyType.Equals("FileKey", StringComparison.OrdinalIgnoreCase) AndAlso
               Not key.KeyType.Equals("RegKey", StringComparison.OrdinalIgnoreCase) Then Continue For

            Dim exactHits As HashSet(Of String) = Nothing
            If indexes.KeyValueIndex.TryGetValue(key.Value, exactHits) Then For Each hit In exactHits : allCandidates.Add(hit) : Next

            ' Path root lookup catches same-root changes (e.g. flag or pattern changes)
            Dim root = GetPathRoot(key.Value)

            If root Is Nothing Then Continue For

            Dim rootHits As HashSet(Of String) = Nothing
            If indexes.PathRootIndex.TryGetValue(root, rootHits) Then For Each hit In rootHits : allCandidates.Add(hit) : Next

            ' Wildcard prefix lookup: check if this old key's path root
            ' would be captured by a new entry's wildcard root.
            ' e.g. %AppData%\GetRightToGo starts with prefix %AppData%\GetRight
            '      from new entry GetRight * whose root is %AppData%\GetRight*
            ' Also handles wildcarded old roots: %AppData%\GetRight* is compared
            ' as %AppData%\GetRight so it matches a new prefix %AppData%\GetRigh
            ' (from %AppData%\GetRigh*, a wildcard reduction).
            Dim comparableRoot = root
            Dim rootStarIdx = root.IndexOf("*"c)
            If rootStarIdx > 0 Then comparableRoot = root.Substring(0, rootStarIdx)

            Dim firstComp = GetFirstComponent(root)

            If firstComp Is Nothing Then Continue For

            Dim prefixList As List(Of KeyValuePair(Of String, String)) = Nothing

            If Not indexes.WildcardPrefixes.TryGetValue(firstComp, prefixList) Then Continue For

            For Each wp In prefixList

                If comparableRoot.StartsWith(wp.Key, StringComparison.OrdinalIgnoreCase) Then allCandidates.Add(wp.Value)

            Next

        Next

        Return allCandidates

    End Function

    ''' <summary>
    ''' Filters a set of candidate section names to only those present in
    ''' <paramref name="eligibleNames"/> and resolves each to its <c> iniSection </c>
    ''' from the new file
    ''' </summary>
    '''
    ''' <param name="candidateNames">
    ''' All candidate section names gathered by the heuristics
    ''' </param>
    '''
    ''' <param name="eligibleNames">
    ''' Set of section names eligible for candidacy (added or modified)
    ''' </param>
    '''
    ''' <returns>
    ''' A list of <c> iniSection </c> instances from the new file for each eligible candidate
    ''' </returns>
    Private Function FilterToEligibleSections(candidateNames As HashSet(Of String),
                                               eligibleNames As HashSet(Of String)) As List(Of iniSection)

        Dim combinedMatches As New List(Of iniSection)

        For Each candidateName In candidateNames

            If Not eligibleNames.Contains(candidateName) Then Continue For

            Dim s2 = _file2.GetSection(candidateName)
            If s2 IsNot Nothing Then combinedMatches.Add(s2)

        Next

        Return combinedMatches

    End Function

    ''' <summary>
    ''' Converts any remaining rename whose target is also in <c> MergedEntryNames </c>
    ''' into a merger, adding the renamed entry to <c> MergeDict </c> and <c> OldToNewMergeDict </c>
    ''' and dropping the rename. In the <c> Parallel.ForEach </c> in <see cref="ProcessRemovals"/>,
    ''' a merger's <c> TrackMerger </c> can run before the competing rename's <c> ConfirmRename </c>
    ''' has registered the rename, so <c> TrackMerger </c> never sees the rename to fold it in.
    ''' This pass catches those cases after all parallel work is complete.
    ''' </summary>
    Private Sub ReconcileRenamesAndMergers()

        Dim renamesToConvert As New List(Of String)

        For Each renamedTarget In _state.MergedEntries.RenamedEntryNames
            If _state.MergedEntries.MergedEntryNames.Contains(renamedTarget) Then renamesToConvert.Add(renamedTarget)
        Next

        For Each target In renamesToConvert

            Dim renameHolder As String = Nothing
            If Not _state.MergedEntries.RenamedEntryPairs.TryGetValue(target, renameHolder) Then Continue For

            If Not _state.MergedEntries.MergeDict.ContainsKey(target) Then _state.MergedEntries.MergeDict.Add(target, New List(Of String))
            If Not _state.MergedEntries.MergeDict(target).Contains(renameHolder) Then _state.MergedEntries.MergeDict(target).Add(renameHolder)

            If Not _state.MergedEntries.OldToNewMergeDict.ContainsKey(renameHolder) Then _state.MergedEntries.OldToNewMergeDict.Add(renameHolder, New List(Of String))
            If Not _state.MergedEntries.OldToNewMergeDict(renameHolder).Contains(target) Then _state.MergedEntries.OldToNewMergeDict(renameHolder).Add(target)

            _state.MergedEntries.RenamedEntryPairs.Remove(target)
            _state.MergedEntries.RenamedEntryNames.Remove(target)

        Next

    End Sub

    ''' <summary>
    ''' Returns the directory-level path root of a key value for indexing, cutting the value
    ''' at its first pipe first. For paths with 3+ backslash components, returns the
    ''' first two (e.g. <c> %AppData%\SomeApp </c>). For paths with exactly 2 components,
    ''' returns the full directory path (e.g. <c> %AppData%\GetRight* </c> from
    ''' <c> %AppData%\GetRight*|GetRight.lst;*.data|RECURSE </c>).
    ''' </summary>
    '''
    ''' <param name="value">
    ''' The raw key value string, optionally containing pipe-delimited patterns and flags
    ''' </param>
    '''
    ''' <returns>
    ''' The first two backslash-delimited path components, or <c> Nothing </c> if the path has no backslash
    ''' </returns>
    Private Shared Function GetPathRoot(value As String) As String

        ' Strip pipe-delimited flags (e.g. "|GetRight.lst;*.data|RECURSE")
        Dim pipeIdx = value.IndexOf("|"c)
        Dim pathOnly = If(pipeIdx > 0, value.Substring(0, pipeIdx), value)

        Dim first = pathOnly.IndexOf("\"c)
        If first < 0 Then Return Nothing

        Dim second = pathOnly.IndexOf("\"c, first + 1)
        Return If(second > 0, pathOnly.Substring(0, second), pathOnly)

    End Function

    ''' <summary>
    ''' Returns the first backslash-delimited component of a path (typically the environment
    ''' variable or drive root). E.g. <c> %AppData%\GetRight* </c> → <c> %AppData% </c>
    ''' </summary>
    '''
    ''' <param name="pathRoot">
    ''' The path string to extract the first component from
    ''' </param>
    '''
    ''' <returns>
    ''' The substring before the first <c> \ </c>, or <c> Nothing </c> if the path has no backslash
    ''' or starts with one
    ''' </returns>
    Private Shared Function GetFirstComponent(pathRoot As String) As String

        Dim idx = pathRoot.IndexOf("\"c)
        Return If(idx > 0, pathRoot.Substring(0, idx), Nothing)

    End Function

    ''' <summary>
    ''' Returns the sections that may be merger/rename candidates by name or browser. A section
    ''' qualifies when its text and the removed entry's text both contain the same value from
    ''' <c> BrowserSecRefs </c> (anywhere in the text, not only in <c> LangSecRef </c>), or when its
    ''' uppercased name contains a word of the old name followed by a space, as a substring.
    ''' We stop reading old name words at the <c> * </c>.
    ''' </summary>
    '''
    ''' <param name="oldNameBroken">
    ''' Space-split word tokens from the removed entry's name, used for substring matching
    ''' </param>
    '''
    ''' <param name="potentialMatchesList">
    ''' Snapshot of added and modified entries to search for candidates
    ''' </param>
    '''
    ''' <param name="snapshotTextMap">
    ''' Pre-computed uppercased string representations of each candidate section, keyed by name.
    ''' A section missing from it is serialized on the spot.
    ''' </param>
    '''
    ''' <param name="oldEntryTextUpper">
    ''' Uppercased string representation of the removed entry, used for browser SecRef matching
    ''' </param>
    '''
    ''' <returns>
    ''' A list of candidate <c> iniSection </c>s whose name or browser SecRef overlaps with the removed entry
    ''' </returns>
    Private Function FindProbableMatches(oldNameBroken As String(),
                                          potentialMatchesList As List(Of iniSection),
                                          snapshotTextMap As Dictionary(Of String, String),
                                          oldEntryTextUpper As String) As List(Of iniSection)

        Dim out = New List(Of iniSection)

        For Each newSection In potentialMatchesList

            Dim upperNewName = newSection.Name.ToUpperInvariant()
            Dim matched = False

            Dim newVerUpper As String = Nothing
            If Not snapshotTextMap.TryGetValue(newSection.Name, newVerUpper) Then
                newVerUpper = newSection.ToString().ToUpperInvariant()
            End If

            For Each browser In BrowserSecRefs

                If Not newVerUpper.Contains(browser) Then Continue For
                If Not oldEntryTextUpper.Contains(browser) Then Continue For

                out.Add(newSection)
                matched = True
                Exit For

            Next

            If matched Then Continue For

            For Each oldNamePiece In oldNameBroken

                If String.Equals(oldNamePiece, "*", StringComparison.InvariantCultureIgnoreCase) Then Exit For

                If upperNewName.IndexOf($"{oldNamePiece.ToUpperInvariant()} ", StringComparison.Ordinal) >= 0 Then out.Add(newSection) : Exit For

            Next

        Next

        Return out

    End Function

    ''' <summary>
    ''' Bundles the three reverse indexes built over potential match FileKey and RegKey values
    ''' for content-aware candidate lookup during removal processing
    ''' </summary>
    Private Class ContentIndexes

        ''' <summary>
        ''' Exact key value → set of section names that contain that value
        ''' </summary>
        Public ReadOnly KeyValueIndex As New Dictionary(Of String, HashSet(Of String))(StringComparer.OrdinalIgnoreCase)

        ''' <summary>
        ''' Path root (first two backslash components) → set of section names,
        ''' catching pattern or flag changes where the root stays the same
        ''' </summary>
        Public ReadOnly PathRootIndex As New Dictionary(Of String, HashSet(Of String))(StringComparer.OrdinalIgnoreCase)

        ''' <summary>
        ''' First path component → list of (prefix before *, section name) pairs
        ''' for wildcard prefix matching
        ''' </summary>
        Public ReadOnly WildcardPrefixes As New Dictionary(Of String, List(Of KeyValuePair(Of String, String)))(StringComparer.OrdinalIgnoreCase)

    End Class

End Class