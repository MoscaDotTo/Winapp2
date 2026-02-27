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

    Private ReadOnly _state As DiffState
    Private ReadOnly _file1 As iniFile2
    Private ReadOnly _file2 As iniFile2
    Private ReadOnly _mergeDetector As MergeDetector2
    Private ReadOnly _keyAnalyzer As KeyModificationAnalyzer2
    Private ReadOnly _renderer As DiffOutputRenderer2

    ''' <summary>
    ''' Initializes a new instance of <c>EntryChangeDetector2</c>
    ''' </summary>
    ''' <param name="state">Shared diff state tracking all entry changes</param>
    ''' <param name="file1">The old version of winapp2.ini as an <c>iniFile2</c></param>
    ''' <param name="file2">The new version of winapp2.ini as an <c>iniFile2</c></param>
    ''' <param name="mergeDetector">Handles rename and merger detection for removed entries</param>
    ''' <param name="keyAnalyzer">Tracks key-level changes between entry versions</param>
    ''' <param name="renderer">Produces <c>MenuSection</c> output for removed entries with no key matches</param>
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

            For Each key In section.Keys

                CleanKeyValue(key)

            Next

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
            If Not _state.Caches.CachedNewEntries2.ContainsKey(section.Name) Then
                _state.Caches.CachedNewEntries2.Add(section.Name, section)
            End If
            snapshotTextMap(section.Name) = section.ToString().ToUpperInvariant()
        Next

        ' Pre-populate old entry cache and pre-compute section text for removed entries.
        Dim oldEntryTextMap As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        For Each entryName In _state.ModifiedEntries.RemovedEntryNames
            Dim oldSection2 = _file1.GetSection(entryName)
            If Not _state.Caches.CachedOldEntries2.ContainsKey(oldSection2.Name) Then
                _state.Caches.CachedOldEntries2.Add(oldSection2.Name, oldSection2)
            End If
            oldEntryTextMap(entryName) = oldSection2.ToString().ToUpperInvariant()
        Next

        Parallel.ForEach(_state.ModifiedEntries.RemovedEntryNames,
                 Sub(entry)

                     Dim oldSection2 = _file1.GetSection(entry)

                     If oldSection2.Keys.GetByType("FileKey").Count = 0 AndAlso
                        oldSection2.Keys.GetByType("RegKey").Count = 0 Then

                         results(entry) = _renderer.MakeDiff(oldSection2, 1)
                         Return

                     End If

                     Dim probableMatches = FindProbableMatches2(entry.Split(CChar(" ")), potentialMatchesSnapshot2, snapshotTextMap, oldEntryTextMap(entry))
                     Dim allCandidates As New HashSet(Of String)

                     For Each section In probableMatches
                         allCandidates.Add(section.Name)
                     Next

                     For Each section In potentialMatchesSnapshot2
                         allCandidates.Add(section.Name)
                     Next

                     Dim combinedMatches As New List(Of iniSection2)

                     For Each candidateName In allCandidates

                         If _state.ModifiedEntries.AddedEntryNames.Contains(candidateName) OrElse
                            _state.ModifiedEntries.ModifiedEntryNames.Contains(candidateName) Then

                             Dim s2 = _file2.GetSection(candidateName)
                             If s2 IsNot Nothing Then combinedMatches.Add(s2)

                         End If

                     Next

                     Dim changesRecorded = _mergeDetector.AssessRenamesAndMergers(combinedMatches, oldSection2)

                     If Not changesRecorded Then results(entry) = _renderer.MakeDiff(oldSection2, 1)

                 End Sub)

        For Each key In results.Keys.OrderBy(Function(k) k, StringComparer.OrdinalIgnoreCase)
            out.Add(results(key))
        Next

        Return out

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
