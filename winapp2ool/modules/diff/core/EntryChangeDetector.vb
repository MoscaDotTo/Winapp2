'    Copyright (C) 2018-2025 Hazel Ward
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
''' </summary>
Public Class EntryChangeDetector

    Private ReadOnly _state As DiffState
    Private ReadOnly _file1 As iniFile
    Private ReadOnly _file2 As iniFile
    Private ReadOnly _mergeDetector As MergeDetector
    Private ReadOnly _keyAnalyzer As KeyModificationAnalyzer
    Private ReadOnly _renderer As DiffOutputRenderer

    ''' <summary>
    ''' Initializes a new instance of <c>EntryChangeDetector</c>
    ''' </summary>
    '''
    ''' <param name="state">Shared diff state tracking all entry changes</param>
    ''' <param name="file1">The old version of winapp2.ini</param>
    ''' <param name="file2">The new version of winapp2.ini</param>
    ''' <param name="mergeDetector">Handles rename and merger detection for removed entries</param>
    ''' <param name="keyAnalyzer">Tracks key-level changes between entry versions</param>
    ''' <param name="renderer">Produces <c>MenuSection</c> output for removed entries with no key matches</param>
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
    ''' Records entries present in the new file but not the old file as added
    ''' </summary>
    Public Sub ProcessNewEntries()

        For Each section In _file2.Sections.Values

            If _file1.Sections.ContainsKey(section.Name) Then Continue For

            _state.ModifiedEntries.AddedEntryNames.Add(section.Name)
            _state.ModifiedEntries.PotentialMatches.Add(section)

        Next

    End Sub

    ''' <summary>
    ''' Processes the entries in the old file. If an entry is not present in the new file, it is recorded as removed.
    ''' If an entry is present in both files, it is compared against the new version for modifications.
    ''' </summary>
    Public Sub ProcessOldEntries()

        For Each section In _file1.Sections.Values

            If Not _file2.Sections.ContainsKey(section.Name) Then

                _state.ModifiedEntries.RemovedEntryNames.Add(section.Name)
                Continue For

            End If

            _keyAnalyzer.FindModifications(section, _file2.Sections(section.Name))

        Next

        For Each modifiedEntryName In _state.ModifiedEntries.ModifiedEntryNames
            _state.ModifiedEntries.PotentialMatches.Add(_file2.Sections(modifiedEntryName))
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
        Dim potentialMatchesSnapshot = _state.ModifiedEntries.PotentialMatches.ToList()

        ' Pre-populate new entry cache and pre-compute section text for potential matches
        Dim snapshotTextMap As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        For Each section In potentialMatchesSnapshot
            If Not _state.Caches.CachedNewEntries.ContainsKey(section.Name) Then
                _state.Caches.CachedNewEntries.Add(section.Name, New winapp2entry(section))
            End If
            snapshotTextMap(section.Name) = section.ToString().ToUpperInvariant()
        Next

        ' Pre-populate old entry cache and pre-compute section text for removed entries
        Dim oldEntryTextMap As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        For Each entryName In _state.ModifiedEntries.RemovedEntryNames
            Dim oldSection = _file1.getSection(entryName)
            If Not _state.Caches.CachedOldEntries.ContainsKey(oldSection.Name) Then
                _state.Caches.CachedOldEntries.Add(oldSection.Name, New winapp2entry(oldSection))
            End If
            oldEntryTextMap(entryName) = oldSection.ToString().ToUpperInvariant()
        Next

        Parallel.ForEach(_state.ModifiedEntries.RemovedEntryNames,
                 Sub(entry)

                     Dim oldSectionVersion = _file1.getSection(entry)
                     Dim oldWa2Section = _state.Caches.CachedOldEntries(oldSectionVersion.Name)

                     If oldWa2Section.FileKeys.KeyCount = 0 AndAlso oldWa2Section.RegKeys.KeyCount = 0 Then

                         results(entry) = _renderer.MakeDiff(oldSectionVersion, 1)
                         Return

                     End If

                     Dim probableMatches = FindProbableMatches(entry.Split(CChar(" ")), potentialMatchesSnapshot, snapshotTextMap, oldEntryTextMap(entry))
                     Dim allCandidates As New HashSet(Of String)

                     For Each section In probableMatches

                         allCandidates.Add(section.Name)

                     Next

                     For Each section In potentialMatchesSnapshot

                         allCandidates.Add(section.Name)

                     Next

                     Dim combinedMatches As New List(Of iniSection)
                     For Each candidateName In allCandidates

                         If _state.ModifiedEntries.AddedEntryNames.Contains(candidateName) OrElse
                            _state.ModifiedEntries.ModifiedEntryNames.Contains(candidateName) Then

                             combinedMatches.Add(_file2.Sections(candidateName))

                         End If

                     Next

                     Dim changesRecorded = _mergeDetector.AssessRenamesAndMergers(combinedMatches, oldSectionVersion)

                     If Not changesRecorded Then results(entry) = _renderer.MakeDiff(oldSectionVersion, 1)

                 End Sub)

        For Each key In results.Keys.OrderBy(Function(k) k, StringComparer.OrdinalIgnoreCase)
            out.Add(results(key))
        Next
        Return out

    End Function

    ''' <summary>
    ''' Creates parity between certain key values that have been changed as part of the winapp2.ini v23XXXX changes <br />
    ''' Used to ignore broad cases of text conversion in Diff, such as environmental variables changes introduced as part of
    ''' the effort to go Non-CCleaner as default winapp2.ini
    ''' </summary>
    '''
    ''' <param name="winapp">
    ''' A winapp2.ini format <c>iniFile</c>
    ''' </param>
    Public Sub SnuffNoisyChanges(ByRef winapp As iniFile)

        For Each section In winapp.Sections.Values

            For Each key In section.Keys.Keys

                CleanKeyValue(key)

            Next

        Next

    End Sub

    ''' <summary>
    ''' Replaces values in a given <c>iniKey</c> that contain any of the <c>OldPaths</c> with the corresponding <c>NewPaths</c>
    ''' </summary>
    '''
    ''' <param name="key">
    ''' An <c>iniKey</c> to be sanitized
    ''' </param>
    Private Sub CleanKeyValue(ByRef key As iniKey)

        For i = 0 To PathReplacements.Count - 1

            Dim oldPath = PathReplacements.Keys(i)
            Dim newPath = PathReplacements.Values(i)

            If Not key.Value.Contains(oldPath) Then Continue For

            key.Value = key.Value.Replace(oldPath, newPath)

        Next

    End Sub

    ''' <summary>
    ''' Produces a list of iniSections who may potentially be merger/rename candidates based on traits such as section and name similarities
    ''' </summary>
    '''
    ''' <param name="oldNameBroken">
    ''' The old name of the entry broken into pieces about the space character, ie. each "word" in the name
    ''' </param>
    '''
    ''' <param name="potentialMatchesList">
    ''' Snapshot of all candidate sections from the new file that could be rename or merger targets
    ''' </param>
    '''
    ''' <param name="snapshotTextMap">
    ''' Pre-computed map of section name to full section text (uppercased) for all potential match candidates
    ''' </param>
    '''
    ''' <param name="oldEntryTextUpper">
    ''' Pre-computed uppercased full text of the old entry being matched
    ''' </param>
    '''
    ''' <returns>
    ''' A <c>List(Of iniSection)</c> either falling into the same <c>LangSecRef</c> or matching a component of the name of the "old" iniSection
    ''' </returns>
    Private Function FindProbableMatches(oldNameBroken As String(),
                                         potentialMatchesList As List(Of iniSection),
                                         snapshotTextMap As Dictionary(Of String, String),
                                         oldEntryTextUpper As String) As List(Of iniSection)

        Dim out = New List(Of iniSection)

        For Each newName In potentialMatchesList

            Dim upperNewName = newName.Name.ToUpperInvariant()
            Dim matched = False

            Dim newVerUpper As String = Nothing
            If Not snapshotTextMap.TryGetValue(newName.Name, newVerUpper) Then
                newVerUpper = newName.ToString().ToUpperInvariant()
            End If

            For Each browser In BrowserSecRefs

                If Not newVerUpper.Contains(browser) Then Continue For
                If Not oldEntryTextUpper.Contains(browser) Then Continue For

                out.Add(newName)
                matched = True
                Exit For

            Next

            If matched Then Continue For

            For Each oldNamePiece In oldNameBroken

                If String.Equals(oldNamePiece, "*", StringComparison.InvariantCultureIgnoreCase) Then Exit For

                If upperNewName.IndexOf($"{oldNamePiece.ToUpperInvariant()} ", StringComparison.Ordinal) >= 0 Then out.Add(newName) : Exit For

            Next

        Next

        Return out

    End Function

End Class
