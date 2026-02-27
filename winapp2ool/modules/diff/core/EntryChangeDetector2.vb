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
''' Adapts <c>EntryChangeDetector</c> for use with <c>iniFile2</c>/<c>iniSection2</c>.
''' Categorizes entries as added, removed, or modified between two versions of winapp2.ini.
''' </summary>
Public Class EntryChangeDetector2

    ''' <summary>
    ''' 
    ''' </summary>
    Private ReadOnly _state As DiffState

    ''' <summary>
    ''' 
    ''' </summary>
    Private ReadOnly _file1 As iniFile2

    ''' <summary>
    ''' 
    ''' </summary>
    Private ReadOnly _file2 As iniFile2

    ''' <summary>
    ''' 
    ''' </summary>
    Private ReadOnly _mergeDetector As MergeDetector2

    ''' <summary>
    ''' 
    ''' </summary>
    Private ReadOnly _keyAnalyzer As KeyModificationAnalyzer2

    ''' <summary>
    ''' 
    ''' </summary>
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
    Public Sub SnuffNoisyChanges(winapp As iniFile2)

        For Each section In winapp

            For Each key In section.Keys : CleanKeyValue(key) : Next

        Next

    End Sub

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
    ''' Processes the entries in the old file. If an entry is not present in the new file, it is recorded as removed.
    ''' If an entry is present in both files, it is compared against the new version for modifications.
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
    ''' Entries which have been renamed: all FileKeys / RegKeys match, but there may be minor changes <br />
    ''' Entries which have been merged: all or most FileKeys / RegKeys / Detects / DetectFiles match, but there may be major changes <br />
    ''' Entries which have actually been removed: key values not found in any new entries <br />
    ''' </summary>
    Public Function ProcessRemovals() As List(Of MenuSection)

        Dim out As New List(Of MenuSection)
        Dim results = New Concurrent.ConcurrentDictionary(Of String, MenuSection)(StringComparer.OrdinalIgnoreCase)
        Dim potentialMatchesSnapshot2 = _state.ModifiedEntries.PotentialMatches2.ToList()

        ' Pre-populate new entry cache and pre-compute section text for potential matches
        Dim snapshotTextMap As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        For Each section In potentialMatchesSnapshot2
            _state.Caches.CachedNewEntries2(section.Name) = section
            snapshotTextMap(section.Name) = section.ToString().ToUpperInvariant()
        Next

        ' Pre-populate old entry cache and pre-compute section text for removed entries.
        Dim oldEntryTextMap As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        For Each entryName In _state.ModifiedEntries.RemovedEntryNames
            Dim oldSection2 = _file1.GetSection(entryName)
            _state.Caches.CachedOldEntries2(oldSection2.Name) = oldSection2
            oldEntryTextMap(entryName) = oldSection2.ToString().ToUpperInvariant()
        Next

        ' Build reverse index: key value -> set of section names that contain that value.
        ' Also index path roots (first two components) to catch wildcard pattern changes
        ' where the path root stays the same (e.g. %AppData%\Foo\*.log -> %AppData%\Foo\*).
        Dim keyValueIndex As New Dictionary(Of String, HashSet(Of String))(StringComparer.OrdinalIgnoreCase)
        Dim pathRootIndex As New Dictionary(Of String, HashSet(Of String))(StringComparer.OrdinalIgnoreCase)

        ' Wildcard prefix index: for new keys whose path root contains *, stores the prefix
        ' (everything before *) grouped by first path component for efficient lookup.
        ' E.g. %AppData%\GetRight* -> prefix "%AppData%\GetRight" under group "%AppData%"
        ' This lets us find that %AppData%\GetRightToGo would be captured by the wildcard.
        Dim wildcardPrefixes As New Dictionary(Of String, List(Of KeyValuePair(Of String, String)))(StringComparer.OrdinalIgnoreCase)

        For Each section In potentialMatchesSnapshot2

            For Each key In section.Keys

                If Not key.KeyType.Equals("FileKey", StringComparison.OrdinalIgnoreCase) AndAlso
                   Not key.KeyType.Equals("RegKey", StringComparison.OrdinalIgnoreCase) Then Continue For

                ' Index exact value
                If Not keyValueIndex.ContainsKey(key.Value) Then keyValueIndex(key.Value) = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                keyValueIndex(key.Value).Add(section.Name)

                ' Index path root (first two backslash components) for wildcard resilience
                Dim root = GetPathRoot(key.Value)

                If root Is Nothing Then Continue For

                If Not pathRootIndex.ContainsKey(root) Then pathRootIndex(root) = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                pathRootIndex(root).Add(section.Name)

                ' If root has wildcard, also index the prefix before * grouped by first component
                Dim starIdx = root.IndexOf("*"c)

                If Not starIdx > 0 Then Continue For

                Dim prefix = root.Substring(0, starIdx)
                Dim firstComp = GetFirstComponent(root)

                If firstComp IsNot Nothing Then

                    If Not wildcardPrefixes.ContainsKey(firstComp) Then wildcardPrefixes(firstComp) = New List(Of KeyValuePair(Of String, String))
                    wildcardPrefixes(firstComp).Add(New KeyValuePair(Of String, String)(prefix, section.Name))

                End If

            Next

        Next

        ' Pre-compute the set of names eligible for candidacy (added or modified)
        Dim eligibleNames As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        For Each section In potentialMatchesSnapshot2

            If _state.ModifiedEntries.AddedEntryNames.Contains(section.Name) OrElse
               _state.ModifiedEntries.ModifiedEntryNames.Contains(section.Name) Then eligibleNames.Add(section.Name)

        Next

        Parallel.ForEach(_state.ModifiedEntries.RemovedEntryNames,
                 Sub(entry)

                     Dim oldSection2 = _file1.GetSection(entry)

                     If oldSection2.Keys.GetByType("FileKey").Count = 0 AndAlso
                        oldSection2.Keys.GetByType("RegKey").Count = 0 Then

                         results(entry) = _renderer.MakeDiff(oldSection2, 1)
                         Return

                     End If

                     ' Gather candidates from all three heuristics
                     Dim allCandidates As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

                     ' 1. Name / browser-ref heuristic (existing logic)
                     Dim probableMatches = FindProbableMatches2(entry.Split(CChar(" ")), potentialMatchesSnapshot2, snapshotTextMap, oldEntryTextMap(entry))
                     For Each section In probableMatches : allCandidates.Add(section.Name) : Next

                     ' 2. Content-aware: look up each old key value in the reverse index
                     For Each key In oldSection2.Keys

                         If Not key.KeyType.Equals("FileKey", StringComparison.OrdinalIgnoreCase) AndAlso
                            Not key.KeyType.Equals("RegKey", StringComparison.OrdinalIgnoreCase) Then Continue For

                         Dim exactHits As HashSet(Of String) = Nothing
                         If keyValueIndex.TryGetValue(key.Value, exactHits) Then For Each hit In exactHits : allCandidates.Add(hit) : Next

                         ' Path root lookup catches same-root changes (e.g. flag or pattern changes)
                         Dim root = GetPathRoot(key.Value)

                         If root Is Nothing Then Continue For


                         Dim rootHits As HashSet(Of String) = Nothing
                         If pathRootIndex.TryGetValue(root, rootHits) Then For Each hit In rootHits : allCandidates.Add(hit) : Next

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

                         If Not wildcardPrefixes.TryGetValue(firstComp, prefixList) Then Continue For

                         For Each wp In prefixList

                             If comparableRoot.StartsWith(wp.Key, StringComparison.OrdinalIgnoreCase) Then allCandidates.Add(wp.Value)

                         Next


                     Next

                     ' Filter to only eligible (added/modified) entries
                     Dim combinedMatches As New List(Of iniSection2)

                     For Each candidateName In allCandidates

                         If Not eligibleNames.Contains(candidateName) Then Continue For

                         Dim s2 = _file2.GetSection(candidateName)
                         If s2 IsNot Nothing Then combinedMatches.Add(s2)

                     Next

                     Dim changesRecorded = _mergeDetector.AssessRenamesAndMergers(combinedMatches, oldSection2)

                     If Not changesRecorded Then results(entry) = _renderer.MakeDiff(oldSection2, 1)

                 End Sub)

        ' Post-parallel reconciliation: convert any rename whose target also received
        ' merger content back to a merger. This handles the race where TrackMerger ran
        ' before ConfirmRename registered the rename, so the in-flight cleanup was skipped.
        ReconcileRenamesAndMergers()

        For Each key In results.Keys.OrderBy(Function(k) k, StringComparer.OrdinalIgnoreCase)
            out.Add(results(key))
        Next

        Return out

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
    ''' Returns <c>Nothing</c> if the value has no backslash.
    ''' </summary>
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
    Private Shared Function GetFirstComponent(pathRoot As String) As String

        Dim idx = pathRoot.IndexOf("\"c)
        Return If(idx > 0, pathRoot.Substring(0, idx), Nothing)

    End Function

    ''' <summary>
    ''' Produces a list of <c>iniSection2</c>s who may potentially be merger/rename candidates
    ''' based on traits such as section and name similarities
    ''' </summary>
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

End Class
