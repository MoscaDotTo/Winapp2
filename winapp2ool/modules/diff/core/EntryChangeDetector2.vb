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
''' Normalizes deprecated path values to suppress false-positive diffs, delegates rename
''' and merger detection to <c>MergeDetector2</c>, and coordinates key-level analysis
''' via <c>KeyModificationAnalyzer2</c>.
''' </summary>
Public Class EntryChangeDetector2

    Private ReadOnly _state As DiffState
    Private ReadOnly _file1 As iniFile2
    Private ReadOnly _file2 As iniFile2
    Private ReadOnly _mergeDetector As MergeDetector2
    Private ReadOnly _keyAnalyzer As KeyModificationAnalyzer2
    Private ReadOnly _renderer As DiffOutputRenderer2

    ''' <summary>
    ''' Initializes a new instance of <c>EntryChangeDetector2</c>
    ''' </summary>
    ''' 
    ''' <param name="state">
    ''' Shared diff state tracking all entry changes
    ''' </param>
    ''' 
    ''' <param name="file1">
    ''' The old version of winapp2.ini as an <c>iniFile2</c>
    ''' </param>
    ''' 
    ''' <param name="file2">
    ''' The new version of winapp2.ini as an <c>iniFile2</c>
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
    ''' Produces <c>MenuSection</c> output for removed entries with no key matches
    ''' </param>
    Public Sub New(state As DiffState,
                   file1 As iniFile2,
                   file2 As iniFile2,
                   mergeDetector As MergeDetector2,
                   keyAnalyzer As KeyModificationAnalyzer2,
                   renderer As DiffOutputRenderer2)

        _state = state
        _file1 = file1
        _file2 = file2
        _mergeDetector = mergeDetector
        _keyAnalyzer = keyAnalyzer
        _renderer = renderer

    End Sub

    ''' <summary>
    ''' Replaces deprecated path values in all keys of a winapp2.ini <c>iniFile2</c>
    ''' to suppress false-positive diff entries caused by known path renames
    ''' </summary>
    '''
    ''' <param name="winapp">
    ''' The file whose key values will be normalized in place
    ''' </param>
    Public Sub SnuffNoisyChanges(winapp As iniFile2)

        For Each section In winapp

            For Each key In section.Keys : CleanKeyValue(key) : Next

        Next

    End Sub

    ''' <summary>
    ''' Replaces deprecated path values in a single key's value in-place
    ''' </summary>
    ''' 
    ''' <param name="key">
    ''' The key whose value is normalized against <c>PathReplacements</c>
    ''' </param>
    Private Sub CleanKeyValue(key As iniKey2)

        For i = 0 To PathReplacements.Count - 1

            Dim oldPath = PathReplacements.Keys(i)

            If Not key.Value.Contains(oldPath) Then Continue For

            key.Value = key.Value.Replace(oldPath, PathReplacements.Values(i))

        Next

    End Sub

    ''' <summary>
    ''' Records entries present in the new file but not the old file as added
    ''' </summary>
    Public Sub ProcessNewEntries()

        For Each section In _file2

            If _file1.Contains(section.Name) Then Continue For

            _state.ModifiedEntries.AddedEntryNames.Add(section.Name)
            _state.ModifiedEntries.PotentialMatches2.Add(section)

        Next

    End Sub

    ''' <summary>
    ''' Processes the entries in the old file. If an entry is not present in the new file,
    ''' it is recorded as removed. If an entry is present in both files, it is compared 
    ''' against the new version for modifications.
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
            _state.ModifiedEntries.PotentialMatches2.Add(_file2.GetSection(modifiedEntryName))
        Next

    End Sub

    ''' <summary>
    ''' Processes the entries determined to have been "Removed" and categorizes them into 3 bins: <br />
    ''' Entries which have been renamed: all FileKeys / RegKeys match, but there may be minor changes <br />  <br />
    ''' Entries which have been merged: all or most FileKeys / RegKeys / Detects / DetectFiles match,
    ''' but there may be major changes <br /> <br />
    ''' Entries which have actually been removed: key values not found in any new entries <br />
    ''' </summary>
    Public Function ProcessRemovals() As List(Of MenuSection)

        Dim results = New Concurrent.ConcurrentDictionary(Of String, MenuSection)(StringComparer.OrdinalIgnoreCase)
        Dim potentialMatchesSnapshot2 = _state.ModifiedEntries.PotentialMatches2.ToList()

        Dim snapshotTextMap As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        Dim oldEntryTextMap As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        PrePopulateCachesAndTextMaps(potentialMatchesSnapshot2, snapshotTextMap, oldEntryTextMap)

        Dim contentIndexes = BuildContentIndexes(potentialMatchesSnapshot2)
        Dim eligibleNames = BuildEligibleNameSet(potentialMatchesSnapshot2)

        Parallel.ForEach(_state.ModifiedEntries.RemovedEntryNames,
                 Sub(entry)

                     Dim result = ProcessSingleRemoval(entry, potentialMatchesSnapshot2, snapshotTextMap, oldEntryTextMap, contentIndexes, eligibleNames)
                     If result IsNot Nothing Then results(entry) = result

                 End Sub)

        ReconcileRenamesAndMergers()

        Dim out As New List(Of MenuSection)

        For Each key In results.Keys.OrderBy(Function(k) k, StringComparer.OrdinalIgnoreCase)
            out.Add(results(key))
        Next

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
    Private Sub PrePopulateCachesAndTextMaps(potentialMatches As List(Of iniSection2),
                                             snapshotTextMap As Dictionary(Of String, String),
                                             oldEntryTextMap As Dictionary(Of String, String))

        For Each section In potentialMatches
            _state.Caches.CachedNewEntries2(section.Name) = section
            snapshotTextMap(section.Name) = section.ToString().ToUpperInvariant()
        Next

        For Each entryName In _state.ModifiedEntries.RemovedEntryNames
            Dim oldSection2 = _file1.GetSection(entryName)
            _state.Caches.CachedOldEntries2(oldSection2.Name) = oldSection2
            oldEntryTextMap(entryName) = oldSection2.ToString().ToUpperInvariant()
        Next

    End Sub

    ''' <summary>
    ''' Builds reverse indexes over the FileKey and RegKey values in <paramref name="potentialMatches"/>
    ''' to enable fast content-aware candidate lookup during removal processing.
    ''' <list type="bullet">
    '''   <item><term>KeyValueIndex</term>
    '''     <description>Exact key value → set of section names containing that value</description></item>
    '''   <item><term>PathRootIndex</term>
    '''     <description>First two backslash components → set of section names, catching
    '''     wildcard pattern changes where the path root stays the same</description></item>
    '''   <item><term>WildcardPrefixes</term>
    '''     <description>For roots containing <c>*</c>, the prefix before <c>*</c> grouped
    '''     by first path component for efficient lookup</description></item>
    ''' </list>
    ''' </summary>
    '''
    ''' <param name="potentialMatches">
    ''' Snapshot of added and modified entries whose keys are indexed
    ''' </param>
    '''
    ''' <returns>
    ''' A <c>ContentIndexes</c> instance containing all three reverse indexes
    ''' </returns>
    Private Shared Function BuildContentIndexes(potentialMatches As List(Of iniSection2)) As ContentIndexes

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
    ''' eligible for candidacy (i.e. present in <c>AddedEntryNames</c> or <c>ModifiedEntryNames</c>)
    ''' </summary>
    '''
    ''' <param name="potentialMatches">
    ''' Snapshot of added and modified entries to filter
    ''' </param>
    '''
    ''' <returns>
    ''' A <c>HashSet</c> of section names eligible for rename/merger matching
    ''' </returns>
    Private Function BuildEligibleNameSet(potentialMatches As List(Of iniSection2)) As HashSet(Of String)

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
    ''' <c>MergeDetector2.AssessRenamesAndMergers</c>. Returns a <c>MenuSection</c> for entries
    ''' that were truly removed (no rename or merger found), or <c>Nothing</c> if a rename/merger
    ''' was recorded.
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
    ''' Reverse content indexes built by <c>BuildContentIndexes</c>
    ''' </param>
    '''
    ''' <param name="eligibleNames">
    ''' Set of section names eligible for candidacy (added or modified)
    ''' </param>
    '''
    ''' <returns>
    ''' A <c>MenuSection</c> describing the removal if no rename/merger was found;
    ''' <c>Nothing</c> if a rename or merger was recorded in <c>DiffState</c>
    ''' </returns>
    Private Function ProcessSingleRemoval(entryName As String,
                                           potentialMatches As List(Of iniSection2),
                                           snapshotTextMap As Dictionary(Of String, String),
                                           oldEntryTextMap As Dictionary(Of String, String),
                                           indexes As ContentIndexes,
                                           eligibleNames As HashSet(Of String)) As MenuSection

        Dim oldSection2 = _file1.GetSection(entryName)

        If oldSection2.Keys.GetByType("FileKey").Count = 0 AndAlso
           oldSection2.Keys.GetByType("RegKey").Count = 0 Then

            Return _renderer.MakeDiff(oldSection2, 1)

        End If

        Dim allCandidates = GatherCandidateNames(entryName, oldSection2, potentialMatches, snapshotTextMap, oldEntryTextMap, indexes)
        Dim combinedMatches = FilterToEligibleSections(allCandidates, eligibleNames)
        Dim changesRecorded = _mergeDetector.AssessRenamesAndMergers(combinedMatches, oldSection2)

        Return If(changesRecorded, Nothing, _renderer.MakeDiff(oldSection2, 1))

    End Function

    ''' <summary>
    ''' Gathers candidate section names for a removed entry using all three heuristics:
    ''' name/browser-ref matching, exact key value index lookup, path root index lookup,
    ''' and wildcard prefix index lookup
    ''' </summary>
    '''
    ''' <param name="entryName">
    ''' The name of the removed entry
    ''' </param>
    '''
    ''' <param name="oldSection2">
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
                                           oldSection2 As iniSection2,
                                           potentialMatches As List(Of iniSection2),
                                           snapshotTextMap As Dictionary(Of String, String),
                                           oldEntryTextMap As Dictionary(Of String, String),
                                           indexes As ContentIndexes) As HashSet(Of String)

        Dim allCandidates As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        ' 1. Name / browser-ref heuristic
        Dim probableMatches = FindProbableMatches2(entryName.Split(CChar(" ")), potentialMatches, snapshotTextMap, oldEntryTextMap(entryName))
        For Each section In probableMatches : allCandidates.Add(section.Name) : Next

        ' 2. Content-aware: look up each old key value in the reverse indexes
        For Each key In oldSection2.Keys

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
    ''' <paramref name="eligibleNames"/> and resolves each to its <c>iniSection2</c>
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
    ''' A list of <c>iniSection2</c> instances from the new file for each eligible candidate
    ''' </returns>
    Private Function FilterToEligibleSections(candidateNames As HashSet(Of String),
                                               eligibleNames As HashSet(Of String)) As List(Of iniSection2)

        Dim combinedMatches As New List(Of iniSection2)

        For Each candidateName In candidateNames

            If Not eligibleNames.Contains(candidateName) Then Continue For

            Dim s2 = _file2.GetSection(candidateName)
            If s2 IsNot Nothing Then combinedMatches.Add(s2)

        Next

        Return combinedMatches

    End Function

    ''' <summary>
    ''' Converts any remaining rename whose target is also in <c>MergedEntryNames</c>
    ''' into a merger. When the <c>Parallel.ForEach</c> in <c>ProcessRemovals</c> runs,
    ''' a merger's <c>TrackMerger</c> may execute before the competing rename's
    ''' <c>ConfirmRename</c> has registered the rename, causing the cleanup branch
    ''' (lines 385-400 of <c>TrackMerger</c>) to be skipped. This pass catches those
    ''' cases after all parallel work is complete.
    ''' </summary>
    Private Sub ReconcileRenamesAndMergers()

        ' Snapshot the set — we'll mutate the underlying collections during iteration
        Dim renamesToConvert As New List(Of String)

        For Each renamedTarget In _state.MergedEntries.RenamedEntryNames
            If _state.MergedEntries.MergedEntryNames.Contains(renamedTarget) Then renamesToConvert.Add(renamedTarget)
        Next

        For Each target In renamesToConvert

            Dim renameHolder As String = Nothing
            If Not _state.MergedEntries.RenamedEntryPairs.TryGetValue(target, renameHolder) Then Continue For

            ' Add the rename holder into the merge dict alongside any existing merge sources
            If Not _state.MergedEntries.MergeDict.ContainsKey(target) Then _state.MergedEntries.MergeDict.Add(target, New List(Of String))
            If Not _state.MergedEntries.MergeDict(target).Contains(renameHolder) Then _state.MergedEntries.MergeDict(target).Add(renameHolder)

            If Not _state.MergedEntries.OldToNewMergeDict.ContainsKey(renameHolder) Then _state.MergedEntries.OldToNewMergeDict.Add(renameHolder, New List(Of String))
            If Not _state.MergedEntries.OldToNewMergeDict(renameHolder).Contains(target) Then _state.MergedEntries.OldToNewMergeDict(renameHolder).Add(target)

            ' Remove the rename
            _state.MergedEntries.RenamedEntryPairs.Remove(target)
            _state.MergedEntries.RenamedEntryNames.Remove(target)

        Next

    End Sub

    ''' <summary>
    ''' Returns the directory-level path root of a key value for indexing, stripping any
    ''' pipe-delimited flags first. For paths with 3+ backslash components, returns the
    ''' first two (e.g. <c>%AppData%\SomeApp</c>). For paths with exactly 2 components,
    ''' returns the full directory path (e.g. <c>%AppData%\GetRight*</c> from
    ''' <c>%AppData%\GetRight*|GetRight.lst;*.data|RECURSE</c>).
    ''' </summary>
    '''
    ''' <param name="value">
    ''' The raw key value string, optionally containing pipe-delimited flags
    ''' </param>
    '''
    ''' <returns>
    ''' The first one or two backslash-delimited path components, or <c>Nothing</c> if the value has no backslash
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
    ''' variable or drive root), or <c>Nothing</c> if the path has no backslash.
    ''' E.g. <c>%AppData%\GetRight*</c> → <c>%AppData%</c>
    ''' </summary>
    '''
    ''' <param name="pathRoot">
    ''' The path string to extract the first component from
    ''' </param>
    '''
    ''' <returns>
    ''' The substring before the first <c>\</c>, or <c>Nothing</c> if no backslash is present
    ''' </returns>
    Private Shared Function GetFirstComponent(pathRoot As String) As String

        Dim idx = pathRoot.IndexOf("\"c)
        Return If(idx > 0, pathRoot.Substring(0, idx), Nothing)

    End Function

    ''' <summary>
    ''' Produces a list of <c>iniSection2</c>s who may potentially be merger/rename candidates
    ''' based on traits such as section and name similarities
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
    ''' Pre-computed uppercased string representations of each candidate section, keyed by name
    ''' </param>
    '''
    ''' <param name="oldEntryTextUpper">
    ''' Uppercased string representation of the removed entry, used for browser SecRef matching
    ''' </param>
    '''
    ''' <returns>
    ''' A list of candidate <c>iniSection2</c>s whose name or browser SecRef overlaps with the removed entry
    ''' </returns>
    Private Function FindProbableMatches2(oldNameBroken As String(),
                                          potentialMatchesList As List(Of iniSection2),
                                          snapshotTextMap As Dictionary(Of String, String),
                                          oldEntryTextUpper As String) As List(Of iniSection2)

        Dim out = New List(Of iniSection2)

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
        ''' catching wildcard pattern changes where the root stays the same
        ''' </summary>
        Public ReadOnly PathRootIndex As New Dictionary(Of String, HashSet(Of String))(StringComparer.OrdinalIgnoreCase)

        ''' <summary>
        ''' First path component → list of (prefix before *, section name) pairs
        ''' for wildcard prefix matching
        ''' </summary>
        Public ReadOnly WildcardPrefixes As New Dictionary(Of String, List(Of KeyValuePair(Of String, String)))(StringComparer.OrdinalIgnoreCase)

    End Class

End Class