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
''' Compares two versions of an entry at the key level, identifying added, removed,
''' and updated <c>iniKey2</c> values. Also supports comparison against a flat key list
''' for entries built from multiple merged sources. All results are written to the
''' <c>DiffState</c> key-change trackers.
''' </summary>
Public Class KeyModificationAnalyzer2

    Private ReadOnly _state As DiffState

    ''' <summary>
    ''' Creates a new <c>KeyModificationAnalyzer2</c> bound to the given diff state
    ''' </summary>
    ''' 
    ''' <param name="state">
    ''' Shared diff state whose key-change trackers will be written to
    ''' </param>
    Public Sub New(state As DiffState)

        _state = state

    End Sub

    ''' <summary>
    ''' Determines the changes made to the <c>iniKey2</c> values in an
    ''' <c>iniSection2</c> that has been updated between versions. <br /> <br />
    ''' Callback-compatible with <c>Action(Of iniSection2, iniSection2)</c>.
    ''' </summary>
    '''
    ''' <param name="oldSection">
    ''' The previous version of the entry
    ''' </param>
    ''' 
    ''' <param name="newSection">
    ''' The current version of the entry
    ''' </param>
    Public Sub FindModifications(oldSection As iniSection2,
                                 newSection As iniSection2)

        AnalyzeAndTrackSectionDiff2(oldSection, newSection, addToModified:=True, clearExisting:=False)

    End Sub

    ''' <summary>
    ''' Computes modifications between a combined old entry and a new added entry.
    ''' Does NOT add to ModifiedEntryNames since this is an added entry.
    ''' </summary>
    '''
    ''' <param name="oldSection">
    ''' The old (removed) entry being compared against
    ''' </param>
    ''' 
    ''' <param name="newSection">
    ''' The new (added) entry being analyzed
    ''' </param>
    Public Sub FindModificationsForAddedEntry(oldSection As iniSection2,
                                              newSection As iniSection2)

        AnalyzeAndTrackSectionDiff2(oldSection, newSection, addToModified:=False, clearExisting:=True)

    End Sub

    ''' <summary>
    ''' Variant of <see cref="FindModifications"/> that accepts a flat key list instead of an
    ''' <c>iniSection2</c> for the old side. Used when combining keys from multiple old entries into
    ''' a synthetic section — avoids <c>iniKeyCollection</c>'s first-write-wins name deduplication
    ''' dropping keys that share a name across source entries (e.g. two FileKey1 values).
    ''' </summary>
    '''
    ''' <param name="oldKeys">
    ''' Flat list of keys from one or more combined old entries
    ''' </param>
    ''' 
    ''' <param name="newSection">
    ''' The current (modified) version of the entry
    ''' </param>
    Public Sub FindModificationsFromCombinedKeys(oldKeys As List(Of iniKey2),
                                                 newSection As iniSection2)

        AnalyzeAndTrackSectionDiff2WithKeyList(oldKeys, newSection, addToModified:=True, clearExisting:=False)

    End Sub

    ''' <summary>
    ''' Variant of <see cref="FindModificationsForAddedEntry"/>
    ''' that accepts a flat key list.
    ''' </summary>
    '''
    ''' <param name="oldKeys">
    ''' Flat list of keys from one or more combined old entries
    ''' </param>
    ''' 
    ''' <param name="newSection">
    ''' The new (added) entry being analyzed
    ''' </param>
    Public Sub FindModificationsForAddedEntryFromKeys(oldKeys As List(Of iniKey2),
                                                      newSection As iniSection2)

        AnalyzeAndTrackSectionDiff2WithKeyList(oldKeys, newSection, addToModified:=False, clearExisting:=True)

    End Sub

    ''' <summary>
    ''' Core implementation for flat-key-list comparisons.
    ''' Computes added, removed, and updated keys,
    ''' then writes them to the <c>DiffState</c> trackers 
    ''' under <paramref name="newSection"/>'s name.
    ''' </summary>
    ''' 
    ''' <param name="oldKeys">
    ''' Flat list of keys representing the old side of the comparison
    ''' </param>
    ''' 
    ''' <param name="newSection">
    ''' The new entry whose keys form the new side of the comparison
    ''' </param>
    ''' 
    ''' <param name="addToModified">
    ''' When <c>True</c>, adds <paramref name="newSection"/>'s
    ''' name to <c>ModifiedEntryNames</c>
    ''' (only if not already in <c>AddedEntryNames</c>)
    ''' </param>
    ''' 
    ''' <param name="clearExisting">
    ''' When <c>True</c>, removes any prior tracker entries for 
    ''' <paramref name="newSection"/> before writing;
    ''' when <c>False</c>, rolls back and replaces prior 
    ''' entries if the entry was already tracked as modified
    ''' </param>
    Private Sub AnalyzeAndTrackSectionDiff2WithKeyList(oldKeys As List(Of iniKey2), newSection As iniSection2,
                                                       addToModified As Boolean, clearExisting As Boolean)

        Dim addedKeys, removedKeys As New List(Of iniKey2)

        If CompareKeyLists(oldKeys, newSection.Keys, removedKeys, addedKeys) Then Return

        WriteResultsToTrackers(newSection.Name, Nothing, addedKeys, removedKeys, addToModified, clearExisting)

    End Sub

    ''' <summary>
    ''' Core implementation for section-to-section comparisons. Computes added, removed, and updated keys,
    ''' injects a Name-change sentinel pair when the section names differ, then writes results to the
    ''' <c>DiffState</c> trackers under <paramref name="newSection"/>'s name.
    ''' </summary>
    ''' 
    ''' <param name="oldSection">
    ''' The previous version of the entry</param>
    ''' 
    ''' <param name="newSection">The current version of the entry</param>
    ''' 
    ''' <param name="addToModified">
    ''' When <c>True</c>, adds <paramref name="newSection"/>'s name to <c>ModifiedEntryNames</c>
    ''' (only if not already in <c>AddedEntryNames</c>) and records a Name-change pair if names differ
    ''' </param>
    ''' 
    ''' <param name="clearExisting">
    ''' When <c>True</c>, removes any prior tracker entries for <paramref name="newSection"/> before writing;
    ''' when <c>False</c>, rolls back and replaces prior entries if the entry was already tracked as modified
    ''' </param>
    Private Sub AnalyzeAndTrackSectionDiff2(oldSection As iniSection2,
                                            newSection As iniSection2,
                                            addToModified As Boolean,
                                            clearExisting As Boolean)

        Dim addedKeys, removedKeys As New List(Of iniKey2)

        If CompareKeyLists(oldSection.Keys, newSection.Keys, removedKeys, addedKeys) Then Return

        WriteResultsToTrackers(newSection.Name, oldSection.Name, addedKeys, removedKeys, addToModified, clearExisting)

    End Sub

    ''' <summary>
    ''' Compares two key sequences by <c>KeyType</c> and <c>Value</c>, populating
    ''' <paramref name="removedKeys"/> and <paramref name="addedKeys"/> with the differences.
    ''' Each new key is consumed at most once, so renumbered keys with the same type and value
    ''' (e.g. FileKey1 → FileKey2) are treated as equivalent.
    ''' Accepts any <c>IEnumerable(Of iniKey2)</c> for the old side, supporting both
    ''' <c>iniSection2.Keys</c> and flat key lists with duplicate names.
    ''' </summary>
    ''' 
    ''' <param name="oldKeys">
    ''' The previous version's keys (from a section or a flat merged list)
    ''' </param>
    ''' 
    ''' <param name="newKeys">
    ''' The current version's keys
    ''' </param>
    ''' 
    ''' <param name="removedKeys">
    ''' Populated with old keys not found in <paramref name="newKeys"/>
    ''' </param>
    ''' 
    ''' <param name="addedKeys">
    ''' Populated with new keys not matched by any old key
    ''' </param>
    ''' 
    ''' <returns>
    ''' <c>True</c> if the key lists are identical (no additions or removals)
    ''' </returns>
    Private Shared Function CompareKeyLists(oldKeys As IEnumerable(Of iniKey2),
                                            newKeys As IEnumerable(Of iniKey2),
                                      ByRef removedKeys As List(Of iniKey2),
                                      ByRef addedKeys As List(Of iniKey2)) As Boolean

        Dim newKeyList = newKeys.ToList()
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

            If Not foundMatch Then removedKeys.Add(oldKey)

        Next

        For i = 0 To newKeyList.Count - 1
            If Not matched.Contains(i) Then addedKeys.Add(newKeyList(i))
        Next

        Return removedKeys.Count = 0 AndAlso addedKeys.Count = 0

    End Function

    ''' <summary>
    ''' Writes computed key-level changes to the <c>DiffState</c> trackers under the given section name.
    ''' Handles clearing/rolling back prior tracker entries, recording added/removed/updated keys,
    ''' optionally adding the section to <c>ModifiedEntryNames</c>, and injecting a Name-change
    ''' sentinel pair when <paramref name="oldSectionName"/> differs from <paramref name="newSectionName"/>.
    ''' </summary>
    '''
    ''' <param name="newSectionName">
    ''' The entry name used as the dictionary key in all trackers
    ''' </param>
    '''
    ''' <param name="oldSectionName">
    ''' The old entry name; when non-<c>Nothing</c> and different from <paramref name="newSectionName"/>,
    ''' a Name-change sentinel pair is injected into the updated keys list. <br /> <br />
    ''' Pass <c>Nothing</c> for flat-key-list comparisons where no name change applies.
    ''' </param>
    '''
    ''' <param name="addedKeys">
    ''' Keys present in the new version but not the old; updated in place
    ''' by <c>DetermineModifiedKeys</c> which removes promoted entries
    ''' </param>
    '''
    ''' <param name="removedKeys">
    ''' Keys present in the old version but not the new; updated in place
    ''' by <c>DetermineModifiedKeys</c> which removes promoted entries
    ''' </param>
    '''
    ''' <param name="addToModified">
    ''' When <c>True</c>, adds <paramref name="newSectionName"/> to <c>ModifiedEntryNames</c>
    ''' (only if not already in <c>AddedEntryNames</c>)
    ''' </param>
    '''
    ''' <param name="clearExisting">
    ''' When <c>True</c>, removes any prior tracker entries before writing;
    ''' when <c>False</c>, rolls back and replaces prior entries if already tracked as modified
    ''' </param>
    Private Sub WriteResultsToTrackers(newSectionName As String,
                                       oldSectionName As String,
                                       addedKeys As List(Of iniKey2),
                                       removedKeys As List(Of iniKey2),
                                       addToModified As Boolean,
                                       clearExisting As Boolean)

        SyncLock _state.ModifiedEntries.ModifiedEntryNames

            If clearExisting Then

                _state.ModifiedEntries.AddedKeyTracker2.Remove(newSectionName)
                _state.ModifiedEntries.RemovedKeyTracker2.Remove(newSectionName)
                _state.ModifiedEntries.ModifiedKeyTracker2.Remove(newSectionName)

            ElseIf _state.ModifiedEntries.ModifiedEntryNames.Contains(newSectionName) Then

                RollBackPreviouslyObservedChanges(newSectionName)

            End If

            Dim updatedKeys = DetermineModifiedKeys(removedKeys, addedKeys)
            If removedKeys.Count + addedKeys.Count + updatedKeys.Count = 0 Then Return

            updateTrackingDictionary(_state.ModifiedEntries.RemovedKeyTracker2, removedKeys, newSectionName)
            updateTrackingDictionary(_state.ModifiedEntries.AddedKeyTracker2, addedKeys, newSectionName)

            If addToModified Then

                If Not _state.ModifiedEntries.AddedEntryNames.Contains(newSectionName) Then _state.ModifiedEntries.ModifiedEntryNames.Add(newSectionName)

                If oldSectionName IsNot Nothing AndAlso
                   Not oldSectionName.Equals(newSectionName, StringComparison.InvariantCultureIgnoreCase) Then

                    Dim oldName = New iniKey2($"Name={oldSectionName}")
                    Dim newName = New iniKey2($"Name={newSectionName}")
                    updatedKeys.Add(New KeyValuePair(Of iniKey2, iniKey2)(newName, oldName))

                End If

            End If

            If updatedKeys.Count = 0 Then Return

            MergeModificationsIntoTracker(newSectionName, updatedKeys, clearExisting)

        End SyncLock

    End Sub

    ''' <summary>
    ''' Merges computed updated-key pairs into the <c>ModifiedKeyTracker2</c> for the given entry.
    ''' When <paramref name="clearExisting"/> is <c>False</c> and the tracker already has entries,
    ''' new modifications are added alongside existing ones without overwriting.
    ''' </summary>
    '''
    ''' <param name="sectionName">
    ''' The entry name used as the dictionary key in <c>ModifiedKeyTracker2</c>
    ''' </param>
    '''
    ''' <param name="updatedKeys">
    ''' Pairs of (new key, old key) to write into the tracker
    ''' </param>
    '''
    ''' <param name="clearExisting">
    ''' When <c>False</c> and the tracker already contains <paramref name="sectionName"/>,
    ''' new entries are merged in; otherwise the tracker entry is replaced wholesale
    ''' </param>
    Private Sub MergeModificationsIntoTracker(sectionName As String,
                                              updatedKeys As List(Of KeyValuePair(Of iniKey2, iniKey2)),
                                              clearExisting As Boolean)

        If Not clearExisting AndAlso _state.ModifiedEntries.ModifiedKeyTracker2.ContainsKey(sectionName) Then

            For Each kvp In BuildModifications(updatedKeys)

                If Not _state.ModifiedEntries.ModifiedKeyTracker2(sectionName).ContainsKey(kvp.Key) Then _state.ModifiedEntries.ModifiedKeyTracker2(sectionName).Add(kvp.Key, kvp.Value)

            Next

        Else

            _state.ModifiedEntries.ModifiedKeyTracker2(sectionName) = BuildModifications(updatedKeys)

        End If

    End Sub

    ''' <summary>
    ''' Removes all tracker entries for <paramref name="sectionName"/> so that 
    ''' a subsequent comparison can replace them with a fresh, authoritative result
    ''' </summary>
    ''' 
    ''' <param name="sectionName">
    ''' The entry name whose tracker entries are to be cleared
    ''' </param>
    Private Sub RollBackPreviouslyObservedChanges(sectionName As String)

        _state.ModifiedEntries.ModifiedEntryNames.Remove(sectionName)
        _state.ModifiedEntries.AddedKeyTracker2.Remove(sectionName)
        _state.ModifiedEntries.RemovedKeyTracker2.Remove(sectionName)
        _state.ModifiedEntries.ModifiedKeyTracker2.Remove(sectionName)

    End Sub

    ''' <summary>
    ''' Appends <paramref name="keys"/> into <paramref name="keyTracker"/> under
    ''' <paramref name="newSectionName"/>, inserting a new list entry if one does 
    ''' not yet exist and skipping duplicates otherwise
    ''' </summary>
    ''' 
    ''' <param name="keyTracker">
    ''' The added or removed key tracker dictionary to update
    ''' </param>
    ''' 
    ''' <param name="keys">
    ''' The keys to add; no-ops if empty
    ''' </param>
    ''' 
    ''' <param name="newSectionName">
    ''' The entry name used as the dictionary key
    ''' </param>
    Private Sub updateTrackingDictionary(ByRef keyTracker As Dictionary(Of String, List(Of iniKey2)),
                                         keys As List(Of iniKey2),
                                         newSectionName As String)

        If keys.Count = 0 Then Return

        If keyTracker.ContainsKey(newSectionName) Then

            For Each key In keys
                If Not keyTracker(newSectionName).Contains(key) Then keyTracker(newSectionName).Add(key)
            Next

        Else

            keyTracker.Add(newSectionName, keys)

        End If

    End Sub

    ''' <summary>
    ''' Converts a flat list of (new key, old key) pairs into the tracker dictionary format,
    ''' grouping multiple old keys under the same new key when they share a name and value
    ''' </summary>
    ''' 
    ''' <param name="updatedKeys">
    ''' Pairs of (new <c>iniKey2</c>, old <c>iniKey2</c>) as produced by <c>DetermineModifiedKeys</c>
    ''' </param>
    ''' 
    ''' <returns>
    ''' A dictionary mapping each distinct new key to the list of old keys it replaced
    ''' </returns>
    Private Function BuildModifications(ByRef updatedKeys As List(Of KeyValuePair(Of iniKey2, iniKey2))) As Dictionary(Of iniKey2, List(Of iniKey2))

        Dim modifications As New Dictionary(Of iniKey2, List(Of iniKey2))

        For Each kvpair In updatedKeys

            Dim existingKey As iniKey2 = Nothing

            For Each k In modifications.Keys

                If String.Equals(k.Name, kvpair.Key.Name, StringComparison.InvariantCultureIgnoreCase) AndAlso
                   String.Equals(k.Value, kvpair.Key.Value, StringComparison.InvariantCultureIgnoreCase) Then

                    existingKey = k
                    Exit For

                End If

            Next

            If existingKey Is Nothing Then

                modifications.Add(kvpair.Key, New List(Of iniKey2))
                modifications(kvpair.Key).Add(kvpair.Value)

            Else

                modifications(existingKey).Add(kvpair.Value)

            End If

        Next

        Return modifications

    End Function

    ''' <summary>
    ''' Promotes (new key, old key) pairs from <paramref name="addedKeys"/> and 
    ''' <paramref name="removedKeys"/> into an "updated" list when the keys are equivalent 
    ''' under the comparison strategy, share a singleton key type (e.g. LangSecRef), or represent
    ''' a known defunct singleton replacement. Matched keys are removed from 
    ''' <paramref name="addedKeys"/> and <paramref name="removedKeys"/> in place.
    ''' </summary>
    ''' 
    ''' <param name="removedKeys">
    ''' Keys absent from the new version; matched entries are removed
    ''' </param>
    ''' 
    ''' <param name="addedKeys">
    ''' Keys absent from the old version; matched entries are removed
    ''' </param>
    ''' 
    ''' <returns>
    ''' Pairs of (new <c>iniKey2</c>, old <c>iniKey2</c>) representing 
    ''' keys that were updated rather than purely added/removed
    ''' </returns>
    Private Function DetermineModifiedKeys(ByRef removedKeys As List(Of iniKey2),
                                           ByRef addedKeys As List(Of iniKey2)) As List(Of KeyValuePair(Of iniKey2, iniKey2))

        Dim updatedKeys As New List(Of KeyValuePair(Of iniKey2, iniKey2))
        Dim classifiers = {"LangSecRef", "Section"}
        Dim defunctSingletonKeys = {"Warning", "DetectOS", "SpecialDetect"}
        Dim matchedOldKeyValues As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        For Each key In addedKeys

            Dim newKeyType = key.KeyType

            For Each sKey In removedKeys

                If matchedOldKeyValues.Contains(sKey.Value) Then Continue For

                Dim oldKeyType = sKey.KeyType

                Dim shouldExistOnce = classifiers.Contains(newKeyType) AndAlso classifiers.Contains(oldKeyType) OrElse
                                      defunctSingletonKeys.Contains(newKeyType) AndAlso newKeyType = oldKeyType

                Dim newCapturesOld = KeyComparisonStrategyFactory.CompareKeys(key, sKey)
                Dim oldCapturesNew = KeyComparisonStrategyFactory.CompareKeys(sKey, key)

                If Not (shouldExistOnce OrElse newCapturesOld OrElse oldCapturesNew) Then Continue For

                updatedKeys.Add(New KeyValuePair(Of iniKey2, iniKey2)(key, sKey))
                matchedOldKeyValues.Add(sKey.Value)

            Next

        Next

        For Each pair In updatedKeys

            addedKeys.Remove(pair.Key)
            removedKeys.Remove(pair.Value)

        Next

        Return updatedKeys

    End Function

End Class
