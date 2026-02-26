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

''' <summary>
''' Adapts <c>KeyModificationAnalyzer</c> for use with <c>iniSection2</c>/<c>iniKey2</c>.
''' Section comparison is implemented directly rather than via <c>iniSection.compareTo</c>.
''' All <c>DiffState</c> writes use <c>iniKey</c>/<c>keyList</c> (converted at the boundary).
''' </summary>
Public Class KeyModificationAnalyzer2

    Private ReadOnly _state As DiffState

    Public Sub New(state As DiffState)
        _state = state
    End Sub

    ''' <summary>
    ''' Determines the changes made to the <c>iniKey2</c> values in an <c>iniSection2</c> that has
    ''' been updated between versions. Callback-compatible with <c>Action(Of iniSection2, iniSection2)</c>.
    ''' </summary>
    Public Sub FindModifications(oldSection As iniSection2, newSection As iniSection2)
        AnalyzeAndTrackSectionDiff2(oldSection, newSection, addToModified:=True, clearExisting:=False)
    End Sub

    ''' <summary>
    ''' Computes modifications between a combined old entry and a new added entry.
    ''' Does NOT add to ModifiedEntryNames since this is an added entry.
    ''' </summary>
    Public Sub FindModificationsForAddedEntry(oldSection As iniSection2, newSection As iniSection2)
        AnalyzeAndTrackSectionDiff2(oldSection, newSection, addToModified:=False, clearExisting:=True)
    End Sub

    ''' <summary>
    ''' Variant of <see cref="FindModifications"/> that accepts a flat key list instead of an
    ''' <c>iniSection2</c> for the old side. Used when combining keys from multiple old entries into
    ''' a synthetic section — avoids <c>iniKeyCollection</c>'s first-write-wins name deduplication
    ''' dropping keys that share a name across source entries (e.g. two FileKey1 values).
    ''' </summary>
    Public Sub FindModificationsFromCombinedKeys(oldKeys As List(Of iniKey2), newSection As iniSection2)
        AnalyzeAndTrackSectionDiff2WithKeyList(oldKeys, newSection, addToModified:=True, clearExisting:=False)
    End Sub

    ''' <summary>
    ''' Variant of <see cref="FindModificationsForAddedEntry"/> that accepts a flat key list.
    ''' </summary>
    Public Sub FindModificationsForAddedEntryFromKeys(oldKeys As List(Of iniKey2), newSection As iniSection2)
        AnalyzeAndTrackSectionDiff2WithKeyList(oldKeys, newSection, addToModified:=False, clearExisting:=True)
    End Sub

    Private Sub AnalyzeAndTrackSectionDiff2WithKeyList(oldKeys As List(Of iniKey2), newSection As iniSection2,
                                                       addToModified As Boolean, clearExisting As Boolean)

        Dim addedKeys, removedKeys As New keyList

        If CompareSections2FromKeyList(oldKeys, newSection, removedKeys, addedKeys) Then Return

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

            updateTrackingDictionary(_state.ModifiedEntries.RemovedKeyTracker, removedKeys, newSection.Name)
            updateTrackingDictionary(_state.ModifiedEntries.AddedKeyTracker, addedKeys, newSection.Name)

            If addToModified AndAlso Not _state.ModifiedEntries.AddedEntryNames.Contains(newSection.Name) Then
                _state.ModifiedEntries.ModifiedEntryNames.Add(newSection.Name)
            End If

            If updatedKeys.Count = 0 Then Return

            If Not clearExisting AndAlso _state.ModifiedEntries.ModifiedKeyTracker.ContainsKey(newSection.Name) Then

                For Each kvp In BuildModifications(updatedKeys)
                    If Not _state.ModifiedEntries.ModifiedKeyTracker(newSection.Name).ContainsKey(kvp.Key) Then _state.ModifiedEntries.ModifiedKeyTracker(newSection.Name).Add(kvp.Key, kvp.Value)
                Next

            Else

                _state.ModifiedEntries.ModifiedKeyTracker(newSection.Name) = BuildModifications(updatedKeys)

            End If

        End SyncLock

    End Sub

    ''' <summary>
    ''' Variant of <see cref="CompareSections2"/> that takes a flat key list for the old side,
    ''' allowing duplicate key names (e.g. two FileKey1 from different source entries).
    ''' </summary>
    Private Function CompareSections2FromKeyList(oldKeys As List(Of iniKey2), new2 As iniSection2,
                                                 ByRef removedKeys As keyList,
                                                 ByRef addedKeys As keyList) As Boolean

        Dim newKeyList = new2.Keys.ToList()
        Dim matched As New HashSet(Of Integer)

        For Each oldKey In oldKeys

            Dim foundMatch = False

            For i = 0 To newKeyList.Count - 1

                If matched.Contains(i) Then Continue For

                If oldKey.KeyType.Equals(newKeyList(i).KeyType, StringComparison.InvariantCultureIgnoreCase) AndAlso
                   oldKey.Value.Equals(newKeyList(i).Value, StringComparison.InvariantCultureIgnoreCase) Then

                    matched.Add(i)
                    foundMatch = True
                    Exit For

                End If

            Next

            If Not foundMatch Then removedKeys.add(DiffFileBridge.ToIniKey(oldKey))

        Next

        For i = 0 To newKeyList.Count - 1
            If Not matched.Contains(i) Then addedKeys.add(DiffFileBridge.ToIniKey(newKeyList(i)))
        Next

        Return removedKeys.KeyCount = 0 AndAlso addedKeys.KeyCount = 0

    End Function

    Private Sub AnalyzeAndTrackSectionDiff2(oldSection As iniSection2, newSection As iniSection2,
                                             addToModified As Boolean, clearExisting As Boolean)

        Dim addedKeys, removedKeys As New keyList

        If CompareSections2(oldSection, newSection, removedKeys, addedKeys) Then Return

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

            updateTrackingDictionary(_state.ModifiedEntries.RemovedKeyTracker, removedKeys, newSection.Name)
            updateTrackingDictionary(_state.ModifiedEntries.AddedKeyTracker, addedKeys, newSection.Name)

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

                For Each kvp In BuildModifications(updatedKeys)
                    If Not _state.ModifiedEntries.ModifiedKeyTracker(newSection.Name).ContainsKey(kvp.Key) Then _state.ModifiedEntries.ModifiedKeyTracker(newSection.Name).Add(kvp.Key, kvp.Value)
                Next

            Else

                _state.ModifiedEntries.ModifiedKeyTracker(newSection.Name) = BuildModifications(updatedKeys)

            End If

        End SyncLock

    End Sub

    ''' <summary>
    ''' Compares two <c>iniSection2</c> objects by <c>KeyType</c> and <c>Value</c>, populating
    ''' <paramref name="removedKeys"/> and <paramref name="addedKeys"/> with the differences (as <c>iniKey</c>).
    ''' Each new key is consumed at most once, so renumbered keys with the same type and value
    ''' (e.g. FileKey1 → FileKey2) are treated as equivalent, matching <c>iniSection.compareTo</c> behavior.
    ''' </summary>
    ''' <returns><c>True</c> if the sections are identical</returns>
    Private Function CompareSections2(old As iniSection2, new2 As iniSection2,
                                      ByRef removedKeys As keyList,
                                      ByRef addedKeys As keyList) As Boolean

        Dim newKeyList = new2.Keys.ToList()
        Dim matched As New HashSet(Of Integer)

        For Each oldKey In old.Keys

            Dim foundMatch = False

            For i = 0 To newKeyList.Count - 1

                If matched.Contains(i) Then Continue For

                If oldKey.KeyType.Equals(newKeyList(i).KeyType, StringComparison.InvariantCultureIgnoreCase) AndAlso
                   oldKey.Value.Equals(newKeyList(i).Value, StringComparison.InvariantCultureIgnoreCase) Then

                    matched.Add(i)
                    foundMatch = True
                    Exit For

                End If

            Next

            If Not foundMatch Then removedKeys.add(DiffFileBridge.ToIniKey(oldKey))

        Next

        For i = 0 To newKeyList.Count - 1
            If Not matched.Contains(i) Then addedKeys.add(DiffFileBridge.ToIniKey(newKeyList(i)))
        Next

        Return removedKeys.KeyCount = 0 AndAlso addedKeys.KeyCount = 0

    End Function

    Private Sub RollBackPreviouslyObservedChanges(sectionName As String)

        _state.ModifiedEntries.ModifiedEntryNames.Remove(sectionName)
        _state.ModifiedEntries.AddedKeyTracker.Remove(sectionName)
        _state.ModifiedEntries.RemovedKeyTracker.Remove(sectionName)
        _state.ModifiedEntries.ModifiedKeyTracker.Remove(sectionName)

    End Sub

    Private Sub updateTrackingDictionary(ByRef keyTracker As Dictionary(Of String, keyList),
                                         keys As keyList,
                                         newSectionName As String)

        If keys.KeyCount = 0 Then Return

        If keyTracker.ContainsKey(newSectionName) Then

            For Each key In keys.Keys
                If Not keyTracker(newSectionName).Keys.Contains(key) Then keyTracker(newSectionName).add(key)
            Next

        Else

            keyTracker.Add(newSectionName, keys)

        End If

    End Sub

    Private Function BuildModifications(ByRef updatedKeys As List(Of KeyValuePair(Of iniKey, iniKey))) As Dictionary(Of iniKey, keyList)

        Dim modifications As New Dictionary(Of iniKey, keyList)

        For Each kvpair In updatedKeys

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
