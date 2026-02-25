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

''' <summary>
''' Detects and tracks key-level modifications between two versions of an <c>iniSection</c>.
''' </summary>
Public Class KeyModificationAnalyzer

    Private ReadOnly _state As DiffState

    Public Sub New(state As DiffState)

        _state = state

    End Sub

    ''' <summary>
    ''' Determines the changes made to the <c>iniKey</c> values in an <c>iniSection</c> that has
    ''' been updated between versions
    ''' </summary>
    '''
    ''' <param name="oldSection">
    ''' One or more entries from the old version of winapp2.ini as a single unit.
    ''' In the case of tracking multiple mergers, this is a combination of all the keys from all
    ''' the merged entries.
    ''' </param>
    '''
    ''' <param name="newSection">An entry from the new version of winapp2.ini</param>
    Public Sub FindModifications(oldSection As iniSection, newSection As iniSection)
        AnalyzeAndTrackSectionDiff(oldSection, newSection, addToModified:=True, clearExisting:=False)
    End Sub

    ''' <summary>
    ''' Computes modifications between a combined old entry and a new added entry.
    ''' Does NOT add to ModifiedEntryNames since this is an added entry.
    ''' </summary>
    Public Sub FindModificationsForAddedEntry(oldSection As iniSection, newSection As iniSection)
        AnalyzeAndTrackSectionDiff(oldSection, newSection, addToModified:=False, clearExisting:=True)
    End Sub

    ''' <summary>
    ''' Core implementation shared by <c>FindModifications</c> and <c>FindModificationsForAddedEntry</c>.
    ''' Compares two sections, updates all tracking dictionaries, and optionally records the entry
    ''' as modified.
    ''' </summary>
    '''
    ''' <param name="oldSection">The old version of the entry (may be a composite of merged entries)</param>
    ''' <param name="newSection">The new version of the entry</param>
    '''
    ''' <param name="addToModified">
    ''' When <c>True</c>, adds <paramref name="newSection"/> to <c>ModifiedEntryNames</c> and
    ''' tracks name changes. Use <c>False</c> for added entries where the name is new by definition.
    ''' </param>
    '''
    ''' <param name="clearExisting">
    ''' When <c>True</c>, clears any previous tracker data for <paramref name="newSection"/> before
    ''' writing (overwrite semantics). When <c>False</c>, rolls back previously observed statistics
    ''' and merges into existing tracker data.
    ''' </param>
    Private Sub AnalyzeAndTrackSectionDiff(oldSection As iniSection, newSection As iniSection,
                                            addToModified As Boolean, clearExisting As Boolean)

        Dim addedKeys, removedKeys As New keyList

        If oldSection.compareTo(newSection, removedKeys, addedKeys) Then Return

        SyncLock _state.ModifiedEntries.ModifiedEntryNames

            If clearExisting Then

                _state.ModifiedEntries.AddedKeyTracker.Remove(newSection.Name)
                _state.ModifiedEntries.RemovedKeyTracker.Remove(newSection.Name)
                _state.ModifiedEntries.ModifiedKeyTracker.Remove(newSection.Name)

            ElseIf _state.ModifiedEntries.ModifiedEntryNames.Contains(newSection.Name) Then

                RollBackPreviouslyObservedChanges(newSection.Name)

            End If

            Dim updatedKeys = DetermineModifiedKeys(removedKeys, addedKeys)
            If removedKeys.KeyCount + addedKeys.KeyCount + updatedKeys.Count = 0 Then Return

            updateTrackingDictionary(_state.ModifiedEntries.RemovedKeyTracker, removedKeys, oldSection, newSection)
            updateTrackingDictionary(_state.ModifiedEntries.AddedKeyTracker, addedKeys, oldSection, newSection)

            If addToModified Then

                If Not _state.ModifiedEntries.AddedEntryNames.Contains(newSection.Name) Then _state.ModifiedEntries.ModifiedEntryNames.Add(newSection.Name)


                If Not oldSection.Name.Equals(newSection.Name, StringComparison.InvariantCultureIgnoreCase) Then

                    Dim oldName = New iniKey(oldSection.Name) With {.KeyType = "Name"}
                    Dim newName = New iniKey(newSection.Name) With {.KeyType = "Name"}
                    updatedKeys.Add(New KeyValuePair(Of iniKey, iniKey)(newName, oldName))

                End If

            End If

            If updatedKeys.Count = 0 Then Return

            If Not clearExisting AndAlso _state.ModifiedEntries.ModifiedKeyTracker.ContainsKey(newSection.Name) Then

                ' Merge into existing tracker
                For Each kvp In BuildModifications(updatedKeys)

                    If Not _state.ModifiedEntries.ModifiedKeyTracker(newSection.Name).ContainsKey(kvp.Key) Then _state.ModifiedEntries.ModifiedKeyTracker(newSection.Name).Add(kvp.Key, kvp.Value)

                Next

            Else

                _state.ModifiedEntries.ModifiedKeyTracker(newSection.Name) = BuildModifications(updatedKeys)

            End If

        End SyncLock

    End Sub

    ''' <summary>
    ''' In the rare event that an entry has its classification changed from "renamed" to "merged"
    ''' through the Diff process, this function handles rolling back the record keeping associated
    ''' with the first classification
    ''' </summary>
    '''
    ''' <param name="sectionName">
    ''' The name of an entry whose classification is being updated from "renamed" to "merged"
    ''' </param>
    Private Sub RollBackPreviouslyObservedChanges(sectionName As String)

        _state.ModifiedEntries.ModifiedEntryNames.Remove(sectionName)
        _state.ModifiedEntries.AddedKeyTracker.Remove(sectionName)
        _state.ModifiedEntries.RemovedKeyTracker.Remove(sectionName)
        _state.ModifiedEntries.ModifiedKeyTracker.Remove(sectionName)

    End Sub

    ''' <summary>
    ''' Updates a tracking dictionary with key changes between two sections
    ''' </summary>
    Private Sub updateTrackingDictionary(ByRef keyTracker As Dictionary(Of String, keyList),
                                         keys As keyList,
                                         oldSection As iniSection,
                                         newSection As iniSection)

        If keys.KeyCount = 0 Then Return

        If keyTracker.ContainsKey(newSection.Name) Then

            For Each key In keys.Keys

                If Not keyTracker(newSection.Name).Keys.Contains(key) Then keyTracker(newSection.Name).add(key)

            Next

        Else

            keyTracker.Add(newSection.Name, keys)

        End If

    End Sub

    ''' <summary>
    ''' Builds a dictionary which contains as its keys each updated key from the new version of winapp2.ini and as the values a list of the old
    ''' keys which were replaced by the new key
    ''' </summary>
    '''
    ''' <param name="updatedKeys">
    ''' A list of updated keys and the old keys determined to have been superseded by them
    ''' </param>
    '''
    ''' <returns>
    ''' A <c>Dictionary(Of iniKey, keyList)</c> where each <c>iniKey</c> is a key in the new version of the file and
    ''' the <c>keyList</c> contains all of the keys from the old version which have been determined to be captured by the new version
    ''' </returns>
    Private Function BuildModifications(ByRef updatedKeys As List(Of KeyValuePair(Of iniKey, iniKey))) As Dictionary(Of iniKey, keyList)

        Dim modifications As New Dictionary(Of iniKey, keyList)

        For Each kvpair In updatedKeys

            ' Find existing entry by comparing key values, not object references
            Dim existingKey As iniKey = Nothing
            For Each k In modifications.Keys

                If String.Equals(k.Name, kvpair.Key.Name, StringComparison.InvariantCultureIgnoreCase) AndAlso
                   String.Equals(k.Value, kvpair.Key.Value, StringComparison.InvariantCultureIgnoreCase) Then

                    existingKey = k
                    Exit For

                End If

            Next

            If existingKey Is Nothing Then

                modifications.Add(kvpair.Key, New keyList)
                modifications(kvpair.Key).add(kvpair.Value)

            Else

                modifications(existingKey).add(kvpair.Value)

            End If

        Next

        Return modifications

    End Function

    ''' <summary>
    ''' Resolves collisions between added and removed keys so as to identify which added keys are modified versions of removed keys,
    ''' updating the given <c>keyLists</c> accordingly
    ''' </summary>
    '''
    ''' <param name="removedKeys">
    ''' <c>iniKeys</c> determined to have been removed from the newer version of the <c>iniSection</c>
    ''' </param>
    '''
    ''' <param name="addedKeys">
    ''' <c>iniKeys</c> determined to have been added to the newer version of the <c>iniSection</c>
    ''' </param>
    '''
    ''' <returns>
    ''' A list of <c>iniKeys</c> and their matched partners (in [new, old] KeyValuePairs) for the purpose of identifying "modified" keys
    ''' </returns>
    Private Function DetermineModifiedKeys(ByRef removedKeys As keyList,
                                           ByRef addedKeys As keyList) As List(Of KeyValuePair(Of iniKey, iniKey))

        Dim updatedKeys As New List(Of KeyValuePair(Of iniKey, iniKey))
        Dim classifiers = {"LangSecRef", "Section"}
        Dim defunctSingletonKeys = {"Warning", "DetectOS", "SpecialDetect"}
        Dim matchedOldKeyValues As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        For Each key In addedKeys.Keys

            Dim newKeyType = key.KeyType

            For Each sKey In removedKeys.Keys

                If matchedOldKeyValues.Contains(sKey.Value) Then Continue For

                Dim oldKeyType = sKey.KeyType

                Dim shouldExistOnce = classifiers.Contains(newKeyType) AndAlso classifiers.Contains(oldKeyType) OrElse
                                      defunctSingletonKeys.Contains(newKeyType) AndAlso newKeyType = oldKeyType

                Dim newCapturesOld = KeyComparisonStrategyFactory.CompareKeys(key, sKey)
                Dim oldCapturesNew = KeyComparisonStrategyFactory.CompareKeys(sKey, key)

                If Not (shouldExistOnce OrElse newCapturesOld OrElse oldCapturesNew) Then Continue For

                updatedKeys.Add(New KeyValuePair(Of iniKey, iniKey)(key, sKey))
                matchedOldKeyValues.Add(sKey.Value)

            Next

        Next

        For Each pair In updatedKeys

            addedKeys.remove(pair.Key)
            removedKeys.remove(pair.Value)

        Next

        Return updatedKeys

    End Function

End Class
