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
''' Aggregates the raw key-change trackers in <see cref="DiffState"/> into summary statistics
''' (added, removed, updated, and replaced key counts per entry category), detects keys that
''' moved between entries, and finds browsers added or removed between the two files.
''' </summary>
Public Class DiffStatisticsCalculator

    Private ReadOnly _state As DiffState
    Private ReadOnly _file1 As iniFile
    Private ReadOnly _file2 As iniFile

    ''' <summary>
    ''' Creates a new <c> DiffStatisticsCalculator </c>
    ''' </summary>
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
    Public Sub New(state As DiffState,
                   file1 As iniFile,
                   file2 As iniFile)

        _state = state
        _file1 = file1
        _file2 = file2

    End Sub

    ''' <summary>
    ''' Compares <c> Section </c> key values containing <c> "Web Browser" </c> (case-insensitive)
    ''' between the two files and records any values present in the new file but absent from the old file in
    ''' <c> DiffStatistics.NewBrowserSectionValues </c>, and any values present in the old file
    ''' but absent from the new file in <c> DiffStatistics.RemovedBrowserSectionValues </c>.
    ''' <br /><br />
    ''' Each browser supported by BrowserBuilder receives a unique Section value
    ''' (e.g. <c> "Brave Web Browser" </c>). A novel value therefore represents a newly
    ''' added browser, not merely new entries for an existing one.
    ''' </summary>
    Public Sub DetectNewBrowserSupport()

        Dim oldSectionValues As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        For Each section In _file1

            For Each key In section.Keys

                If Not key.typeIs("Section") Then Continue For
                If key.Value.IndexOf(BrowserSectionSubstring, StringComparison.OrdinalIgnoreCase) < 0 Then Continue For

                oldSectionValues.Add(key.Value)

            Next

        Next

        Dim newSectionValues As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Dim novelValues As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        For Each section In _file2

            For Each key In section.Keys

                If Not key.typeIs("Section") Then Continue For
                If key.Value.IndexOf(BrowserSectionSubstring, StringComparison.OrdinalIgnoreCase) < 0 Then Continue For

                newSectionValues.Add(key.Value)
                If oldSectionValues.Contains(key.Value) Then Continue For
                If novelValues.Add(key.Value) Then _state.Statistics.NewBrowserSectionValues.Add(key.Value)

            Next

        Next

        Dim removedValues As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        For Each oldValue In oldSectionValues

            If newSectionValues.Contains(oldValue) Then Continue For
            If removedValues.Add(oldValue) Then _state.Statistics.RemovedBrowserSectionValues.Add(oldValue)

        Next

        _state.Statistics.NewBrowserSectionValues.Sort(StringComparer.OrdinalIgnoreCase)
        _state.Statistics.RemovedBrowserSectionValues.Sort(StringComparer.OrdinalIgnoreCase)

    End Sub

    ''' <summary>
    ''' Adds up the added, removed and updated key counts of modified entries from the raw
    ''' trackers. We count only names in <c> ModifiedEntryNames </c>, because the trackers also
    ''' hold results for rename and merge targets. Run it after <see cref="DetectCrossEntryMovements"/>,
    ''' so that moved keys are no longer in the trackers.
    ''' </summary>
    Public Sub CalculateInitialStatistics()

        For Each kvp In _state.ModifiedEntries.AddedKeyTracker

            If Not _state.ModifiedEntries.ModifiedEntryNames.Contains(kvp.Key) Then Continue For
            _state.Statistics.ModEntriesAddedKeyTotal += kvp.Value.Count
            _state.Statistics.ModEntriesAddedKeyEntryCount += 1

        Next

        For Each kvp In _state.ModifiedEntries.RemovedKeyTracker

            If Not _state.ModifiedEntries.ModifiedEntryNames.Contains(kvp.Key) Then Continue For
            _state.Statistics.ModEntriesRemovedKeysWithoutReplacementTotal += kvp.Value.Count
            _state.Statistics.ModEntriesRemovedKeyEntryCount += 1

        Next

        ' ModEntriesUpdatedKeyTotal = number of new keys that replaced old keys
        ' ModEntriesReplacedByUpdateTotal = number of old keys that were replaced
        ' ModEntriesUpdatedKeyEntryCount = number of entries with key updates
        For Each kvp In _state.ModifiedEntries.ModifiedKeyTracker

            If Not _state.ModifiedEntries.ModifiedEntryNames.Contains(kvp.Key) Then Continue For

            _state.Statistics.ModEntriesUpdatedKeyEntryCount += 1
            _state.Statistics.ModEntriesUpdatedKeyTotal += kvp.Value.Count

            For Each updateKvp In kvp.Value

                _state.Statistics.ModEntriesReplacedByUpdateTotal += updateKvp.Value.Count

            Next

        Next

    End Sub

    ''' <summary>
    ''' Calculates key-level statistics for renamed entries by reading
    ''' the tracker dictionaries keyed by each rename's new name
    ''' </summary>
    Public Sub CalculateRenameStatistics()

        For Each newName In _state.MergedEntries.RenamedEntryNames

            Dim hasAdded = _state.ModifiedEntries.AddedKeyTracker.ContainsKey(newName) AndAlso
                           _state.ModifiedEntries.AddedKeyTracker(newName).Count > 0

            Dim hasRemoved = _state.ModifiedEntries.RemovedKeyTracker.ContainsKey(newName) AndAlso
                             _state.ModifiedEntries.RemovedKeyTracker(newName).Count > 0

            Dim realUpdateCount = 0
            Dim realReplacedCount = 0

            If _state.ModifiedEntries.ModifiedKeyTracker.ContainsKey(newName) Then

                For Each updateKvp In _state.ModifiedEntries.ModifiedKeyTracker(newName)

                    If updateKvp.Key.typeIs("Name") Then Continue For

                    realUpdateCount += 1
                    realReplacedCount += updateKvp.Value.Count

                Next

            End If

            Dim hasMoves = _state.KeyMovements.MovedOutOf(newName).Count > 0 OrElse
                           _state.KeyMovements.MovedInto(newName).Count > 0

            If Not hasAdded AndAlso Not hasRemoved AndAlso realUpdateCount = 0 AndAlso Not hasMoves Then

                _state.Statistics.RenamedEntriesNameOnlyCount += 1
                Continue For

            End If

            If hasAdded Then

                _state.Statistics.RenamedEntriesAddedKeyTotal += _state.ModifiedEntries.AddedKeyTracker(newName).Count
                _state.Statistics.RenamedEntriesAddedKeyEntryCount += 1

            End If

            If hasRemoved Then

                _state.Statistics.RenamedEntriesRemovedKeyTotal += _state.ModifiedEntries.RemovedKeyTracker(newName).Count
                _state.Statistics.RenamedEntriesRemovedKeyEntryCount += 1

            End If

            If realUpdateCount > 0 Then

                _state.Statistics.RenamedEntriesUpdatedKeyTotal += realUpdateCount
                _state.Statistics.RenamedEntriesReplacedByUpdateTotal += realReplacedCount
                _state.Statistics.RenamedEntriesUpdatedKeyEntryCount += 1

            End If

        Next

    End Sub

    ''' <summary>
    ''' Detects keys that left an entry present in both files and landed in another entry, records
    ''' each in <see cref="KeyMovementTracker.MovedKeys"/>, and takes the moved keys out of the
    ''' added and removed trackers. The landing entry can be modified, renamed, or added; we
    ''' consider every key of an added entry that isn't a rename, since none of them is tracked
    ''' yet. Only <see cref="MovableKeyTypes"/> move. A removed key pairs with a same-type key in
    ''' another entry, preferring one with the same value, then one that captures it or that it
    ''' captures, taking entries in name order (ignoring case). Detection keys move only to the
    ''' same value. Several removed keys can move into one key that covers them all. Run it after
    ''' all parallel processing completes and before <see cref="CalculateInitialStatistics"/>,
    ''' which then counts only the keys that didn't move.
    ''' </summary>
    Public Sub DetectCrossEntryMovements()

        Dim candidates As New List(Of AddedKeyInfo)()
        For Each kvp In _state.ModifiedEntries.AddedKeyTracker

            For Each key In kvp.Value : candidates.Add(New AddedKeyInfo(kvp.Key, key)) : Next

        Next

        For Each entryName In _state.ModifiedEntries.AddedEntryNames

            If _state.MergedEntries.RenamedEntryNames.Contains(entryName) Then Continue For
            If _state.ModifiedEntries.AddedKeyTracker.ContainsKey(entryName) Then Continue For

            For Each key In _file2.GetSection(entryName).Keys : candidates.Add(New AddedKeyInfo(entryName, key)) : Next

        Next

        Dim addedByType As New Dictionary(Of String, List(Of AddedKeyInfo))(StringComparer.OrdinalIgnoreCase)
        For Each info In candidates.OrderBy(Function(c) c.EntryName, StringComparer.OrdinalIgnoreCase)

            If Not MovableKeyTypes.Contains(info.Key.KeyType) Then Continue For
            If Not addedByType.ContainsKey(info.Key.KeyType) Then addedByType(info.Key.KeyType) = New List(Of AddedKeyInfo)
            addedByType(info.Key.KeyType).Add(info)

        Next

        Dim keysToRemoveFromAdded As New Dictionary(Of String, List(Of iniKey))
        Dim keysToRemoveFromRemoved As New Dictionary(Of String, List(Of iniKey))
        For Each kvp In _state.ModifiedEntries.RemovedKeyTracker.OrderBy(Function(e) e.Key, StringComparer.OrdinalIgnoreCase)

            Dim sourceEntry = kvp.Key
            Dim removedKeyList = kvp.Value
            For Each removedKey In removedKeyList.ToList()

                Dim sameTypeAdded As List(Of AddedKeyInfo) = Nothing
                If Not addedByType.TryGetValue(removedKey.KeyType, sameTypeAdded) Then Continue For

                Dim elsewhere = sameTypeAdded.Where(Function(added) Not String.Equals(sourceEntry, added.EntryName, StringComparison.OrdinalIgnoreCase))
                Dim match = elsewhere.FirstOrDefault(Function(added) String.Equals(added.Key.Value, removedKey.Value, StringComparison.OrdinalIgnoreCase))

                ' A parent path Detect captures every child path under it, which says nothing about where a key went
                If match Is Nothing AndAlso Not DetectionKeyTypes.Contains(removedKey.KeyType) Then

                    match = elsewhere.FirstOrDefault(Function(added) KeyComparisonStrategyFactory.CompareKeys(added.Key, removedKey) OrElse
                                                                     KeyComparisonStrategyFactory.CompareKeys(removedKey, added.Key))

                End If

                If match Is Nothing Then Continue For

                Dim targetEntry = match.EntryName
                Dim movementKey = $"{removedKey.Name}{MovementKeySeparator}{removedKey.Value}{MovementKeySeparator}{sourceEntry}"
                _state.KeyMovements.MovedKeys(movementKey) = New KeyMovementInfo(sourceEntry, targetEntry, removedKey, match.Key)
                _state.Statistics.ModEntriesMovedKeysTotal += 1

                If Not keysToRemoveFromRemoved.ContainsKey(sourceEntry) Then keysToRemoveFromRemoved(sourceEntry) = New List(Of iniKey)
                keysToRemoveFromRemoved(sourceEntry).Add(removedKey)

                ' An added entry has no tracker yet; its moved keys are credited when it is itemized
                If Not _state.ModifiedEntries.AddedKeyTracker.ContainsKey(targetEntry) Then Continue For
                If Not keysToRemoveFromAdded.ContainsKey(targetEntry) Then keysToRemoveFromAdded(targetEntry) = New List(Of iniKey)
                keysToRemoveFromAdded(targetEntry).Add(match.Key)

            Next

        Next

        For Each kvp In keysToRemoveFromRemoved

            Dim sourceEntry = kvp.Key
            For Each key In kvp.Value : _state.ModifiedEntries.RemovedKeyTracker(sourceEntry).Remove(key) : Next

            If _state.ModifiedEntries.RemovedKeyTracker(sourceEntry).Count = 0 Then _state.ModifiedEntries.RemovedKeyTracker.Remove(sourceEntry)

        Next

        For Each kvp In keysToRemoveFromAdded

            Dim targetEntry = kvp.Key
            For Each key In kvp.Value : _state.ModifiedEntries.AddedKeyTracker(targetEntry).Remove(key) : Next

            If _state.ModifiedEntries.AddedKeyTracker(targetEntry).Count = 0 Then _state.ModifiedEntries.AddedKeyTracker.Remove(targetEntry)

        Next

        Dim sourceEntries As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Dim intoModified As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Dim intoRenamed As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Dim intoAdded As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        For Each movementInfo In _state.KeyMovements.MovedKeys.Values

            sourceEntries.Add(movementInfo.SourceEntry)

            Dim target = movementInfo.TargetEntry
            If _state.MergedEntries.RenamedEntryNames.Contains(target) Then

                intoRenamed.Add(target)
                _state.Statistics.MovedIntoRenamedKeyTotal += 1

            ElseIf _state.ModifiedEntries.ModifiedEntryNames.Contains(target) Then

                intoModified.Add(target)
                _state.Statistics.MovedIntoModifiedKeyTotal += 1

            Else

                intoAdded.Add(target)
                _state.Statistics.MovedIntoAddedKeyTotal += 1

            End If

        Next

        _state.Statistics.ModEntriesMovedKeysSourceCount = sourceEntries.Count
        _state.Statistics.ModEntriesMovedKeysTargetCount = intoModified.Count + intoRenamed.Count + intoAdded.Count
        _state.Statistics.MovedIntoModifiedEntryCount = intoModified.Count
        _state.Statistics.MovedIntoRenamedEntryCount = intoRenamed.Count
        _state.Statistics.MovedIntoAddedEntryCount = intoAdded.Count

    End Sub

    ''' <summary>
    ''' Calculates the statistics for added entries with merged content. Run it after
    ''' <see cref="DiffOutputRenderer.ItemizeAddedEntriesWithMergers"/>, which writes the
    ''' carried-over keys into the added-key tracker that this reads.
    ''' </summary>
    Public Sub CalculateAddedWithMergersStatistics()

        Dim entriesWithMergers As New List(Of String)
        Dim oldEntryCaptures = BuildOldEntryCaptureTracking(entriesWithMergers)

        ComputeKeyCaptureRates(oldEntryCaptures)
        LogSourceEntryKeyStatus(oldEntryCaptures)

        _state.Statistics.AddedWithMergersEntryCount = entriesWithMergers.Count
        _state.Statistics.AddedWithMergersSourceEntryCount = oldEntryCaptures.Count

        ' Compute per-new-entry stats from the tracked key data.
        ' Note: by the time this runs, AddedKeyTracker entries contain both truly novel keys
        ' and carried-over keys (added there by ItemizeAddedEntriesWithMergers for display).
        ' We separate them by checking against the set of key values from the merged old entries.
        ' Keys that moved in from a surviving entry are neither, and are counted as moved.
        For Each newEntryName In entriesWithMergers

            Dim movedIn = _state.KeyMovements.MovedInto(newEntryName)

            Dim allMergedKeyValues As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            For Each oldEntryName In _state.MergedEntries.MergeDict(newEntryName)


                If Not oldEntryCaptures.ContainsKey(oldEntryName) Then Continue For

                For Each v In oldEntryCaptures(oldEntryName).AllKeyValues : allMergedKeyValues.Add(v) : Next

            Next

            Dim addedKeys = If(_state.ModifiedEntries.AddedKeyTracker.ContainsKey(newEntryName),
                               _state.ModifiedEntries.AddedKeyTracker(newEntryName), New List(Of iniKey))
            Dim removedKeys = If(_state.ModifiedEntries.RemovedKeyTracker.ContainsKey(newEntryName),
                                 _state.ModifiedEntries.RemovedKeyTracker(newEntryName), New List(Of iniKey))
            Dim updatedKeysDict = If(_state.ModifiedEntries.ModifiedKeyTracker.ContainsKey(newEntryName),
                                     _state.ModifiedEntries.ModifiedKeyTracker(newEntryName),
                                     New Dictionary(Of iniKey, List(Of iniKey)))

            Dim novelCount = 0
            Dim carriedOverCount = 0
            For Each k In addedKeys

                If allMergedKeyValues.Contains(k.Value) Then

                    carriedOverCount += 1

                ElseIf Not movedIn.Any(Function(m) String.Equals(m.TargetKey.KeyType, k.KeyType, StringComparison.OrdinalIgnoreCase) AndAlso
                                                   String.Equals(m.TargetKey.Value, k.Value, StringComparison.OrdinalIgnoreCase)) Then

                    novelCount += 1

                End If

            Next

            If novelCount > 0 Then

                _state.Statistics.AddedWithMergersNovelKeysTotal += novelCount
                _state.Statistics.AddedWithMergersNovelKeysEntryCount += 1

            End If

            If carriedOverCount > 0 Then

                _state.Statistics.AddedWithMergersCarriedOverKeysTotal += carriedOverCount
                _state.Statistics.AddedWithMergersCarriedOverKeysEntryCount += 1

            End If

            If removedKeys.Count > 0 Then

                _state.Statistics.AddedWithMergersDroppedKeysTotal += removedKeys.Count
                _state.Statistics.AddedWithMergersDroppedEntryCount += 1

            End If

            If updatedKeysDict.Count = 0 Then Continue For

            _state.Statistics.AddedWithMergersCapturingKeysTotal += updatedKeysDict.Count
            _state.Statistics.AddedWithMergersCapturedKeysTotal += updatedKeysDict.Values.Sum(Function(oldKeys) oldKeys.Count)
            _state.Statistics.AddedWithMergersCapturingEntryCount += 1

        Next

    End Sub

    ''' <summary>
    ''' Initializes capture tracking for the old entries merged into added entries that aren't renames
    ''' </summary>
    '''
    ''' <param name="entriesWithMergers">
    ''' List populated by this method with the names of added entries that contain merged content
    ''' </param>
    '''
    ''' <returns>
    ''' A dictionary mapping each old entry name to its <c> OldEntryKeyTracking </c> instance,
    ''' pre-populated with all key values from that entry
    ''' </returns>
    Private Function BuildOldEntryCaptureTracking(entriesWithMergers As List(Of String)) As Dictionary(Of String, OldEntryKeyTracking)

        Dim oldEntryCaptures As New Dictionary(Of String, OldEntryKeyTracking)(StringComparer.OrdinalIgnoreCase)

        For Each newEntryName In _state.ModifiedEntries.AddedEntryNames

            If _state.MergedEntries.RenamedEntryNames.Contains(newEntryName) Then Continue For
            If Not _state.MergedEntries.MergeDict.ContainsKey(newEntryName) Then Continue For

            entriesWithMergers.Add(newEntryName)

            For Each oldEntryName In _state.MergedEntries.MergeDict(newEntryName)

                If oldEntryCaptures.ContainsKey(oldEntryName) Then Continue For
                If Not _file1.Contains(oldEntryName) Then Continue For

                Dim oldSection = _file1.GetSection(oldEntryName)
                Dim tracking As New OldEntryKeyTracking With {
                    .AllKeyValues = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase),
                    .CapturedKeyValues = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                }

                For Each key In oldSection.Keys : tracking.AllKeyValues.Add(key.Value) : Next

                oldEntryCaptures(oldEntryName) = tracking

            Next

        Next

        Return oldEntryCaptures

    End Function

    ''' <summary>
    ''' For each old entry in <paramref name="oldEntryCaptures"/>, checks which of its keys are present
    ''' in, or captured by a key of, any new entry it was merged into, added or modified, and marks
    ''' them as captured in the tracking record
    ''' </summary>
    '''
    ''' <param name="oldEntryCaptures">
    ''' Tracking dictionary built by <c> BuildOldEntryCaptureTracking </c>; capture sets are updated in place
    ''' </param>
    Private Sub ComputeKeyCaptureRates(oldEntryCaptures As Dictionary(Of String, OldEntryKeyTracking))

        Dim targetsForOld As New Dictionary(Of String, HashSet(Of String))(StringComparer.OrdinalIgnoreCase)
        For Each kvp In _state.MergedEntries.MergeDict

            For Each oldName In kvp.Value

                If Not targetsForOld.ContainsKey(oldName) Then targetsForOld(oldName) = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                targetsForOld(oldName).Add(kvp.Key)

            Next

        Next

        For Each oldEntryKvp In oldEntryCaptures

            Dim tracking = oldEntryKvp.Value
            Dim oldSection = _file1.GetSection(oldEntryKvp.Key)

            Dim relevantTargets As HashSet(Of String) = Nothing
            If Not targetsForOld.TryGetValue(oldEntryKvp.Key, relevantTargets) Then Continue For

            For Each newEntryName In relevantTargets

                Dim newSection = _file2.GetSection(newEntryName)

                Dim newKeyValues = New HashSet(Of String)(newSection.Keys.Select(Function(k) k.Value), StringComparer.OrdinalIgnoreCase)

                For Each oldKey In oldSection.Keys

                    If tracking.CapturedKeyValues.Contains(oldKey.Value) Then Continue For

                    If newKeyValues.Contains(oldKey.Value) Then

                        tracking.CapturedKeyValues.Add(oldKey.Value)
                        Continue For

                    End If

                    For Each newKey In newSection.Keys

                        If Not KeyComparisonStrategyFactory.CompareKeys(newKey, oldKey) Then Continue For

                        tracking.CapturedKeyValues.Add(oldKey.Value)
                        Exit For

                    Next

                Next

            Next

        Next

    End Sub

    ''' <summary>
    ''' Logs per-entry key status showing captured and dropped keys together,
    ''' followed by a summary displaying totals and capture rates by type
    ''' </summary>
    '''
    ''' <param name="oldEntryCaptures">
    ''' Tracking dictionary whose entries are enumerated to produce the per-entry log output
    ''' </param>
    Private Sub LogSourceEntryKeyStatus(oldEntryCaptures As Dictionary(Of String, OldEntryKeyTracking))

        If oldEntryCaptures.Count = 0 Then Return

        Dim categories() = {"[DELETION]", "[DETECTION]", "[CATEGORY]", "[OTHER]"}
        Dim totByCategory As New Dictionary(Of String, Integer)
        Dim capByCategory As New Dictionary(Of String, Integer)

        For Each cat In categories

            totByCategory(cat) = 0
            capByCategory(cat) = 0

        Next

        gLog("SOURCE ENTRY KEY STATUS REPORT:")
        gLog("✓ = Captured, ✗ = Dropped", buffr:=True)

        For Each kvp In oldEntryCaptures

            Dim tracking = kvp.Value
            Dim oldSection = _file1.GetSection(kvp.Key)
            Dim capturedCount = tracking.CapturedKeyValues.Count
            Dim droppedCount = tracking.AllKeyValues.Count - capturedCount

            If capturedCount = 0 AndAlso droppedCount = 0 Then Continue For

            gLog($"[{kvp.Key}] - {capturedCount} captured, {droppedCount} dropped")

            For Each key In oldSection.Keys

                Dim marker = getMarker(key)
                Dim wasCaptured = tracking.CapturedKeyValues.Contains(key.Value)

                gLog($"    {If(wasCaptured, "✓", "✗")} {marker} {key.Name}={key.Value}")

                totByCategory(marker) += 1
                If wasCaptured Then capByCategory(marker) += 1

            Next

            gLog("")

        Next

        Dim labels() = {"DELETION  (FileKey/RegKey):",
                        "DETECTION (Detect/DetectFile):",
                        "CATEGORY  (Section/LangSecRef):",
                        "OTHER     (Warning etc.):"}

        gLog(String.Format("{0,-34}{1,7}{2,10}{3,9}{4,14}", "KEY STATUS SUMMARY:", "Total", "Captured", "Dropped", "Capture Rate"))

        For i = 0 To categories.Length - 1

            Dim tot = totByCategory(categories(i))
            Dim cap = capByCategory(categories(i))
            Dim drp = tot - cap
            Dim rate = If(tot > 0, String.Format("{0:F1}%", cap / CDbl(tot) * 100), "N/A")

            gLog(String.Format("{0,-34}{1,7}{2,10}{3,9}{4,14}", labels(i), tot, cap, drp, rate))

        Next

        Dim grandTot = 0, grandCap = 0
        For Each cat In categories

            grandTot += totByCategory(cat)
            grandCap += capByCategory(cat)

        Next

        Dim grandRate = If(grandTot > 0, String.Format("{0:F1}%", grandCap / CDbl(grandTot) * 100), "N/A")
        gLog(String.Format("{0,-34}{1,7}{2,10}{3,9}{4,14}", "All keys:", grandTot, grandCap, grandTot - grandCap, grandRate))
        gLog("")

    End Sub

    ''' <summary>
    ''' Returns a category marker string for a given key type
    ''' </summary>
    '''
    ''' <param name="key">
    ''' The key whose type is mapped to a category label
    ''' </param>
    '''
    ''' <returns>
    ''' One of <c> [DELETION] </c>, <c> [DETECTION] </c>, <c> [CATEGORY] </c>, or <c> [OTHER] </c>
    ''' </returns>
    Private Function getMarker(key As iniKey) As String

        Dim isDelete = key.KeyType = "FileKey" OrElse key.KeyType = "RegKey"
        Dim isDetect = key.KeyType = "Detect" OrElse key.KeyType = "DetectFile"
        Dim isCategory = key.KeyType = "Section" OrElse key.KeyType = "LangSecRef"
        Dim marker = If(isDelete, "[DELETION]", If(isDetect, "[DETECTION]", If(isCategory, "[CATEGORY]", "[OTHER]")))

        Return marker

    End Function

    ''' <summary>
    ''' An added key paired with the name of the entry to which it was added
    ''' </summary>
    Private Class AddedKeyInfo

        ''' <summary>
        ''' The name of the entry to which the <c> Key </c> was added
        ''' </summary>
        Public Property EntryName As String

        ''' <summary>
        ''' The added key
        ''' </summary>
        Public Property Key As iniKey

        ''' <summary>
        ''' Creates a new <c> AddedKeyInfo </c> instance
        ''' </summary>
        ''' 
        ''' <param name="entry">
        ''' The name of the entry to which the key was added 
        ''' </param>
        ''' 
        ''' <param name="k">
        ''' The added key
        ''' </param>
        Public Sub New(entry As String,
                       k As iniKey)

            EntryName = entry
            Key = k

        End Sub

    End Class

    ''' <summary>
    ''' Tracks which of one old entry's key values its merge targets captured. Every set compares
    ''' values case-insensitively.
    ''' </summary>
    Private Class OldEntryKeyTracking

        ''' <summary>
        ''' All keys present in the old entry
        ''' </summary>
        Public Property AllKeyValues As HashSet(Of String)

        ''' <summary>
        ''' Keys that have been captured by new entries
        ''' </summary>
        Public Property CapturedKeyValues As HashSet(Of String)

    End Class

End Class
