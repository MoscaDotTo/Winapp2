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

        Dim newKeyList = newKeys.Where(Function(k) Not IgnoredKeyTypes.Contains(k.KeyType)).ToList()
        Dim filteredOldKeys = oldKeys.Where(Function(k) Not IgnoredKeyTypes.Contains(k.KeyType))
        Dim matched As New HashSet(Of Integer)

        For Each oldKey In filteredOldKeys

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
    ''' a known defunct singleton replacement. Two further passes then pair RegKeys that share a
    ''' registry path but target different value-names (see <see cref="PairRegKeysBySharedPath"/>)
    ''' and one-for-one Detect/DetectFile detection-criteria swaps (same-type first, then cross-type;
    ''' see <see cref="PairDetectionCriteria"/>). Matched keys are removed from
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
        Dim classifiers = ClassifierKeyTypes
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

        updatedKeys.AddRange(PairRegKeysBySharedPath(removedKeys, addedKeys))
        updatedKeys.AddRange(PairDetectionCriteria(removedKeys, addedKeys))

        Return updatedKeys

    End Function

    ''' <summary>
    ''' Pairs detection-criteria changes as modifications, preferring the least ambiguous reading.
    ''' An entry's detection criteria (its <c>Detect</c> and <c>DetectFile</c> keys) is a single
    ''' conceptual role: how winapp2ool decides the target application is present. <br /> <br />
    ''' Two passes run over the still-unmatched detection keys. First, any detection key type with
    ''' exactly one added and one removed key is paired same-type — e.g. one <c>DetectFile</c>
    ''' replaced by another at a different path. Then, if exactly one added and one removed detection
    ''' key remain (necessarily of different types), they are paired as a cross-type swap — e.g. a
    ''' <c>Detect</c> becoming a <c>DetectFile</c>. Any detection key that cannot be paired one-for-one
    ''' this way is left as a plain add/remove, so a lone removed <c>Detect</c> alongside a matched
    ''' <c>DetectFile</c> swap stays reported as removed.
    ''' </summary>
    '''
    ''' <param name="removedKeys">
    ''' The still-unmatched removed keys; paired detection keys are removed in place
    ''' </param>
    '''
    ''' <param name="addedKeys">
    ''' The still-unmatched added keys; paired detection keys are removed in place
    ''' </param>
    '''
    ''' <returns>
    ''' Pairs of (new <c>iniKey2</c>, old <c>iniKey2</c>) for each one-for-one detection swap found,
    ''' or an empty list if none qualify
    ''' </returns>
    Private Shared Function PairDetectionCriteria(ByRef removedKeys As List(Of iniKey2),
                                                  ByRef addedKeys As List(Of iniKey2)) As List(Of KeyValuePair(Of iniKey2, iniKey2))

        Dim pairs As New List(Of KeyValuePair(Of iniKey2, iniKey2))

        ' Same-type swaps first (least ambiguous), so a DetectFile↔DetectFile change is preferred
        ' over consuming an unrelated Detect that should remain a removal.
        For Each detectionType In DetectionKeyTypes

            Dim dt = detectionType
            PairSingleSwap(removedKeys, addedKeys, pairs, Function(k) k.typeIs(dt, ignoreCase:=True))

        Next

        ' Then a single remaining detection key on each side (necessarily cross-type) is a swap too.
        PairSingleSwap(removedKeys, addedKeys, pairs, Function(k) DetectionKeyTypes.Contains(k.KeyType))

        Return pairs

    End Function

    ''' <summary>
    ''' If exactly one added key and exactly one removed key satisfy <paramref name="predicate"/>,
    ''' records them as a (new, old) pair in <paramref name="pairs"/> and removes both from
    ''' <paramref name="addedKeys"/> and <paramref name="removedKeys"/>. Does nothing otherwise,
    ''' so an ambiguous many-to-one or one-sided change is left untouched.
    ''' </summary>
    '''
    ''' <param name="removedKeys">
    ''' The still-unmatched removed keys; the matched key is removed in place
    ''' </param>
    '''
    ''' <param name="addedKeys">
    ''' The still-unmatched added keys; the matched key is removed in place
    ''' </param>
    '''
    ''' <param name="pairs">
    ''' The accumulator the matched (new, old) pair is appended to
    ''' </param>
    '''
    ''' <param name="predicate">
    ''' Selects which keys participate in this swap (e.g. a specific detection key type)
    ''' </param>
    Private Shared Sub PairSingleSwap(ByRef removedKeys As List(Of iniKey2),
                                      ByRef addedKeys As List(Of iniKey2),
                                      pairs As List(Of KeyValuePair(Of iniKey2, iniKey2)),
                                      predicate As Func(Of iniKey2, Boolean))

        Dim removedMatches = removedKeys.Where(predicate).ToList()
        Dim addedMatches = addedKeys.Where(predicate).ToList()

        If removedMatches.Count <> 1 OrElse addedMatches.Count <> 1 Then Return

        pairs.Add(New KeyValuePair(Of iniKey2, iniKey2)(addedMatches(0), removedMatches(0)))

        addedKeys.Remove(addedMatches(0))
        removedKeys.Remove(removedMatches(0))

    End Sub

    ''' <summary>
    ''' Pairs RegKeys that share an identical registry path (the portion before the first <c> | </c>)
    ''' but target different value-names, promoting each such pair to an updated key. <br /> <br />
    ''' Because the winapp2.ini format permits at most one value-name per RegKey, the registry path
    ''' uniquely groups the value-names cleaned beneath it: a path with exactly one removed and one
    ''' added RegKey is an unambiguous change to that key's cleaning rule rather than an independent
    ''' removal and addition. Paths with more than one removed or added RegKey are left untouched,
    ''' since which removal pairs with which addition cannot be determined.
    ''' </summary>
    '''
    ''' <param name="removedKeys">
    ''' The still-unmatched removed keys; promoted RegKeys are removed in place
    ''' </param>
    '''
    ''' <param name="addedKeys">
    ''' The still-unmatched added keys; promoted RegKeys are removed in place
    ''' </param>
    '''
    ''' <returns>
    ''' Pairs of (new <c>iniKey2</c>, old <c>iniKey2</c>) for every registry path carrying
    ''' exactly one removed and one added RegKey
    ''' </returns>
    Private Shared Function PairRegKeysBySharedPath(ByRef removedKeys As List(Of iniKey2),
                                                    ByRef addedKeys As List(Of iniKey2)) As List(Of KeyValuePair(Of iniKey2, iniKey2))

        Dim pairs As New List(Of KeyValuePair(Of iniKey2, iniKey2))

        Dim removedByPath = GroupRegKeysByPath(removedKeys)
        Dim addedByPath = GroupRegKeysByPath(addedKeys)

        For Each path In removedByPath.Keys

            If removedByPath(path).Count <> 1 Then Continue For

            Dim addedForPath As List(Of iniKey2) = Nothing
            If Not addedByPath.TryGetValue(path, addedForPath) Then Continue For
            If addedForPath.Count <> 1 Then Continue For

            pairs.Add(New KeyValuePair(Of iniKey2, iniKey2)(addedForPath(0), removedByPath(path)(0)))

        Next

        For Each pair In pairs

            addedKeys.Remove(pair.Key)
            removedKeys.Remove(pair.Value)

        Next

        Return pairs

    End Function

    ''' <summary>
    ''' Groups RegKeys by their registry path (the portion before the first <c> | </c>),
    ''' skipping any non-RegKey entries. Path matching is case-insensitive, consistent with
    ''' the key comparison strategies.
    ''' </summary>
    '''
    ''' <param name="keys">
    ''' The keys to group; entries whose type is not RegKey are ignored
    ''' </param>
    '''
    ''' <returns>
    ''' A dictionary mapping each registry path to the RegKeys that target a value beneath it
    ''' </returns>
    Private Shared Function GroupRegKeysByPath(keys As List(Of iniKey2)) As Dictionary(Of String, List(Of iniKey2))

        Dim grouped As New Dictionary(Of String, List(Of iniKey2))(StringComparer.InvariantCultureIgnoreCase)

        For Each key In keys

            If Not key.typeIs("RegKey") Then Continue For

            Dim path = key.PipeSplit(0)

            If Not grouped.ContainsKey(path) Then grouped.Add(path, New List(Of iniKey2))
            grouped(path).Add(key)

        Next

        Return grouped

    End Function

End Class
