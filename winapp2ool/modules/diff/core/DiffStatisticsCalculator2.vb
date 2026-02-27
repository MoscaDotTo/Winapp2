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
''' Adapts <c>DiffStatisticsCalculator</c> for use with <c>iniFile2</c>/<c>iniSection2</c>.
''' </summary>
Public Class DiffStatisticsCalculator2

    Private ReadOnly _state As DiffState
    Private ReadOnly _file1 As iniFile2
    Private ReadOnly _file2 As iniFile2

    ''' <summary>
    ''' Initializes a new instance of <c>DiffStatisticsCalculator2</c>
    ''' </summary>
    '''
    ''' <param name="state">Shared diff state tracking all entry changes</param>
    ''' <param name="file1">The old version of winapp2.ini as an <c>iniFile2</c></param>
    ''' <param name="file2">The new version of winapp2.ini as an <c>iniFile2</c></param>
    Public Sub New(state As DiffState, file1 As iniFile2, file2 As iniFile2)

        _state = state
        _file1 = file1
        _file2 = file2

    End Sub

    ''' <summary>
    ''' Calculates statistics from raw trackers before movement detection.
    ''' Filters to <c>ModifiedEntryNames</c> only — the tracker dictionaries also contain added-merger
    ''' entries populated by <c>FindModificationsForAddedEntry</c> which must not be counted here.
    ''' </summary>
    Public Sub CalculateInitialStatistics()

        For Each kvp In _state.ModifiedEntries.AddedKeyTracker2

            If Not _state.ModifiedEntries.ModifiedEntryNames.Contains(kvp.Key) Then Continue For
            _state.Statistics.ModEntriesAddedKeyTotal += kvp.Value.Count
            _state.Statistics.ModEntriesAddedKeyEntryCount += 1

        Next

        For Each kvp In _state.ModifiedEntries.RemovedKeyTracker2

            If Not _state.ModifiedEntries.ModifiedEntryNames.Contains(kvp.Key) Then Continue For
            _state.Statistics.ModEntriesRemovedKeysWithoutReplacementTotal += kvp.Value.Count
            _state.Statistics.ModEntriesRemovedKeyEntryCount += 1

        Next

        ' Count updated/replaced keys
        ' ModEntriesUpdatedKeyTotal = number of NEW keys that replaced old keys
        ' ModEntriesReplacedByUpdateTotal = number of OLD keys that were replaced
        ' ModEntriesUpdatedKeyEntryCount = number of entries with key updates
        For Each kvp In _state.ModifiedEntries.ModifiedKeyTracker2

            If Not _state.ModifiedEntries.ModifiedEntryNames.Contains(kvp.Key) Then Continue For

            _state.Statistics.ModEntriesUpdatedKeyEntryCount += 1

            ' kvp.Value is Dictionary(Of iniKey2, List(Of iniKey2))
            ' Each key in this dictionary is a NEW key
            _state.Statistics.ModEntriesUpdatedKeyTotal += kvp.Value.Count

            ' Count how many OLD keys were replaced
            For Each updateKvp In kvp.Value

                ' updateKvp.Value is List(Of iniKey2) of OLD keys replaced by this NEW key
                _state.Statistics.ModEntriesReplacedByUpdateTotal += updateKvp.Value.Count

            Next

        Next

    End Sub

    ''' <summary>
    ''' Detects keys that were removed from one entry and added to another (cross-entry movements).
    ''' Must be called after all parallel processing completes.
    ''' </summary>
    Public Sub DetectCrossEntryMovements()

        Dim addedKeyInfo As New List(Of AddedKeyInfo)()

        ' Build lookup of all added keys
        For Each kvp In _state.ModifiedEntries.AddedKeyTracker2
            Dim entryName = kvp.Key
            Dim keyList2 = kvp.Value

            For Each key In keyList2
                addedKeyInfo.Add(New AddedKeyInfo(entryName, key))
            Next
        Next

        ' Pre-group added keys by KeyType to skip cross-type comparisons in the inner loop
        Dim addedByType As New Dictionary(Of String, List(Of AddedKeyInfo))(StringComparer.OrdinalIgnoreCase)
        For Each info In addedKeyInfo
            If Not addedByType.ContainsKey(info.Key.KeyType) Then addedByType(info.Key.KeyType) = New List(Of AddedKeyInfo)
            addedByType(info.Key.KeyType).Add(info)
        Next

        ' Track keys to remove from trackers after detection
        Dim keysToRemoveFromAdded As New Dictionary(Of String, List(Of iniKey2))
        Dim keysToRemoveFromRemoved As New Dictionary(Of String, List(Of iniKey2))

        ' Check each removed key to see if it was added elsewhere
        For Each kvp In _state.ModifiedEntries.RemovedKeyTracker2
            Dim sourceEntry = kvp.Key
            Dim removedKeyList = kvp.Value

            For Each removedKey In removedKeyList.ToList()

                ' Only compare against added keys of the same type
                Dim sameTypeAdded As List(Of AddedKeyInfo) = Nothing
                If Not addedByType.TryGetValue(removedKey.KeyType, sameTypeAdded) Then Continue For

                ' Look for matching added keys
                For Each addedInfo In sameTypeAdded
                    Dim targetEntry = addedInfo.EntryName
                    Dim addedKey = addedInfo.Key

                    ' Same entry = not a movement
                    If String.Equals(sourceEntry, targetEntry, StringComparison.OrdinalIgnoreCase) Then Continue For

                    ' Use the same comparison logic as DetermineModifiedKeys
                    ' Check both directions: new captures old OR old captures new
                    Dim newCapturesOld = KeyComparisonStrategyFactory.CompareKeys(addedKey, removedKey)
                    Dim oldCapturesNew = KeyComparisonStrategyFactory.CompareKeys(removedKey, addedKey)

                    ' For a movement, we need an exact match (bidirectional equivalence)
                    If Not (newCapturesOld AndAlso oldCapturesNew) Then Continue For

                    ' Track the movement
                    Dim movementKey = $"{removedKey.Name}{MovementKeySeparator}{removedKey.Value}{MovementKeySeparator}{sourceEntry}"
                    _state.KeyMovements.MovedKeys(movementKey) = New KeyMovementInfo(sourceEntry, targetEntry)
                    _state.Statistics.ModEntriesMovedKeysTotal += 1

                    ' Mark keys for removal from trackers
                    If Not keysToRemoveFromRemoved.ContainsKey(sourceEntry) Then keysToRemoveFromRemoved(sourceEntry) = New List(Of iniKey2)

                    keysToRemoveFromRemoved(sourceEntry).Add(removedKey)

                    If Not keysToRemoveFromAdded.ContainsKey(targetEntry) Then keysToRemoveFromAdded(targetEntry) = New List(Of iniKey2)

                    keysToRemoveFromAdded(targetEntry).Add(addedKey)

                    ' Decrement added/removed statistics
                    _state.Statistics.ModEntriesAddedKeyTotal -= 1
                    _state.Statistics.ModEntriesRemovedKeysWithoutReplacementTotal -= 1

                    Exit For ' Only match once per removed key

                Next

            Next

        Next

        ' Actually remove moved keys from trackers
        For Each kvp In keysToRemoveFromRemoved

            Dim sourceEntry = kvp.Key
            For Each key In kvp.Value

                _state.ModifiedEntries.RemovedKeyTracker2(sourceEntry).Remove(key)

            Next

            ' Clean up empty lists
            If _state.ModifiedEntries.RemovedKeyTracker2(sourceEntry).Count = 0 Then _state.ModifiedEntries.RemovedKeyTracker2.Remove(sourceEntry)

        Next

        For Each kvp In keysToRemoveFromAdded

            Dim targetEntry = kvp.Key
            For Each key In kvp.Value

                _state.ModifiedEntries.AddedKeyTracker2(targetEntry).Remove(key)

            Next

            ' Clean up empty lists
            If _state.ModifiedEntries.AddedKeyTracker2(targetEntry).Count = 0 Then _state.ModifiedEntries.AddedKeyTracker2.Remove(targetEntry)

        Next

        ' Count unique source and target entries
        Dim sourceEntries As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Dim targetEntries As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        For Each movementInfo In _state.KeyMovements.MovedKeys.Values
            sourceEntries.Add(movementInfo.SourceEntry)
            targetEntries.Add(movementInfo.TargetEntry)
        Next

        _state.Statistics.ModEntriesMovedKeysSourceCount = sourceEntries.Count
        _state.Statistics.ModEntriesMovedKeysTargetCount = targetEntries.Count

    End Sub

    ''' <summary>
    ''' Calculates accurate statistics for added entries with merged content.
    ''' </summary>
    Public Sub CalculateAddedWithMergersStatistics()

        Dim entriesWithMergers As New List(Of String)
        Dim oldEntryCaptures = BuildOldEntryCaptureTracking(entriesWithMergers)

        ComputeKeyCaptureRates(oldEntryCaptures)

        Dim totals = SummarizeCaptureStatistics(oldEntryCaptures)
        LogSourceEntryKeyStatus(oldEntryCaptures)

        _state.Statistics.AddedWithMergersEntryCount = entriesWithMergers.Count
        _state.Statistics.AddedWithMergersSourceEntryCount = oldEntryCaptures.Count
        _state.Statistics.AddedWithMergersCapturedKeysTotal = totals.CapturedContent
        _state.Statistics.AddedWithMergersDroppedKeysTotal = totals.TotalContent - totals.CapturedContent

        ' Compute per-new-entry stats from the tracked key data.
        ' Note: by the time this runs, AddedKeyTracker entries contain both truly novel keys
        ' and carried-over keys (added there by ItemizeAddedEntriesWithMergers for display).
        ' We separate them by checking against the set of key values from the merged old entries.
        For Each newEntryName In entriesWithMergers

            Dim allMergedKeyValues As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            For Each oldEntryName In _state.MergedEntries.MergeDict(newEntryName)
                If oldEntryCaptures.ContainsKey(oldEntryName) Then
                    For Each v In oldEntryCaptures(oldEntryName).AllKeyValues
                        allMergedKeyValues.Add(v)
                    Next
                End If
            Next

            Dim addedKeys = If(_state.ModifiedEntries.AddedKeyTracker2.ContainsKey(newEntryName),
                               _state.ModifiedEntries.AddedKeyTracker2(newEntryName), New List(Of iniKey2))
            Dim removedKeys = If(_state.ModifiedEntries.RemovedKeyTracker2.ContainsKey(newEntryName),
                                 _state.ModifiedEntries.RemovedKeyTracker2(newEntryName), New List(Of iniKey2))
            Dim updatedKeysDict = If(_state.ModifiedEntries.ModifiedKeyTracker2.ContainsKey(newEntryName),
                                     _state.ModifiedEntries.ModifiedKeyTracker2(newEntryName),
                                     New Dictionary(Of iniKey2, List(Of iniKey2)))

            Dim novelCount = 0
            Dim carriedOverCount = 0
            For Each k In addedKeys
                If allMergedKeyValues.Contains(k.Value) Then carriedOverCount += 1 Else novelCount += 1
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
                _state.Statistics.AddedWithMergersDroppedEntryCount += 1
            End If

            If updatedKeysDict.Count > 0 Then
                _state.Statistics.AddedWithMergersCapturingKeysTotal += updatedKeysDict.Count
                _state.Statistics.AddedWithMergersCapturingEntryCount += 1
            End If

        Next

    End Sub

    ''' <summary>
    ''' Initializes capture tracking for all old entries involved in mergers
    ''' </summary>
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
                    .CapturedKeyValues = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase),
                    .AllContentKeyValues = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase),
                    .CapturedContentKeyValues = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                }

                For Each key In oldSection.Keys

                    tracking.AllKeyValues.Add(key.Value)
                    If key.KeyType = "FileKey" OrElse key.KeyType = "RegKey" Then tracking.AllContentKeyValues.Add(key.Value)

                Next

                oldEntryCaptures(oldEntryName) = tracking

            Next

        Next

        Return oldEntryCaptures

    End Function

    Private Sub ComputeKeyCaptureRates(oldEntryCaptures As Dictionary(Of String, OldEntryKeyTracking))

        ' Build reverse lookup: oldEntryName -> Set(Of newEntryNames) it was merged into
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

            ' Only check the specific new entries this old entry was merged into
            Dim relevantTargets As HashSet(Of String) = Nothing
            If Not targetsForOld.TryGetValue(oldEntryKvp.Key, relevantTargets) Then Continue For

            For Each newEntryName In relevantTargets

                Dim newSection = _file2.GetSection(newEntryName)

                Dim newKeyValues = New HashSet(Of String)(newSection.Keys.Select(Function(k) k.Value), StringComparer.OrdinalIgnoreCase)

                For Each oldKey In oldSection.Keys

                    If tracking.CapturedKeyValues.Contains(oldKey.Value) Then Continue For

                    ' Fast path: exact value match avoids CompareKeys entirely
                    If newKeyValues.Contains(oldKey.Value) Then
                        tracking.CapturedKeyValues.Add(oldKey.Value)
                        If oldKey.KeyType = "FileKey" OrElse oldKey.KeyType = "RegKey" Then tracking.CapturedContentKeyValues.Add(oldKey.Value)
                        Continue For
                    End If

                    ' Slow path: wildcard / regex comparison
                    For Each newKey In newSection.Keys

                        If KeyComparisonStrategyFactory.CompareKeys(newKey, oldKey) Then

                            tracking.CapturedKeyValues.Add(oldKey.Value)
                            If oldKey.KeyType = "FileKey" OrElse oldKey.KeyType = "RegKey" Then tracking.CapturedContentKeyValues.Add(oldKey.Value)
                            Exit For

                        End If

                    Next

                Next

            Next

        Next

    End Sub

    ''' <summary>
    ''' Calculates and logs overall capture rate statistics. Returns key count totals.
    ''' </summary>
    Private Function SummarizeCaptureStatistics(oldEntryCaptures As Dictionary(Of String, OldEntryKeyTracking)) As KeyCaptureTotals

        Dim totals As New KeyCaptureTotals

        For Each tracking In oldEntryCaptures.Values

            totals.Total += tracking.AllKeyValues.Count
            totals.Captured += tracking.CapturedKeyValues.Count
            totals.TotalContent += tracking.AllContentKeyValues.Count
            totals.CapturedContent += tracking.CapturedContentKeyValues.Count

        Next

        Return totals

    End Function

    ''' <summary>
    ''' Logs per-entry key status showing captured and dropped keys together,
    ''' followed by a summary displaying totals and capture rates by type
    ''' </summary>
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
        gLog("✓ = Captured, ✗ = Dropped")
        gLog("")

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

        Dim labels() = {"DELETION  (FileKey/RegKey):", "DETECTION (Detect/DetectFile):", "CATEGORY  (Section/LangSecRef):", "OTHER     (Warning etc.):"}

        gLog(String.Format("{0,-34}{1,7}{2,10}{3,9}{4,8}", "KEY STATUS SUMMARY:", "Total", "Captured", "Dropped", "Capture Rate"))

        For i = 0 To categories.Length - 1

            Dim tot = totByCategory(categories(i))
            Dim cap = capByCategory(categories(i))
            Dim drp = tot - cap
            Dim rate = If(tot > 0, String.Format("{0:F1}%", cap / CDbl(tot) * 100), "N/A")

            gLog(String.Format("{0,-34}{1,7}{2,10}{3,9}{4,8}", labels(i), tot, cap, drp, rate))

        Next

        Dim grandTot = 0, grandCap = 0
        For Each cat In categories

            grandTot += totByCategory(cat)
            grandCap += capByCategory(cat)

        Next

        Dim grandRate = If(grandTot > 0, String.Format("{0:F1}%", grandCap / CDbl(grandTot) * 100), "N/A")
        gLog(String.Format("{0,-34}{1,7}{2,10}{3,9}{4,8}", "All keys:", grandTot, grandCap, grandTot - grandCap, grandRate))
        gLog("")

    End Sub

    ''' <summary>
    ''' Returns a category marker string for a given key type
    ''' </summary>
    Private Function getMarker(key As iniKey2) As String

        Dim isDelete = (key.KeyType = "FileKey" OrElse key.KeyType = "RegKey")
        Dim isDetect = key.KeyType = "Detect" OrElse key.KeyType = "DetectFile"
        Dim isCategory = key.KeyType = "Section" OrElse key.KeyType = "LangSecRef"
        Dim marker = If(isDelete, "[DELETION]", If(isDetect, "[DETECTION]", If(isCategory, "[CATEGORY]", "[OTHER]")))

        Return marker

    End Function

    ''' <summary>
    ''' Helper class for tracking added key information
    ''' </summary>
    Private Class AddedKeyInfo

        ''' <summary>The name of the entry this key was added to</summary>
        Public Property EntryName As String

        ''' <summary>The added key</summary>
        Public Property Key As iniKey2

        Public Sub New(entry As String, k As iniKey2)

            EntryName = entry
            Key = k

        End Sub

    End Class

    ''' <summary>
    ''' Holds aggregate key count totals from capture analysis
    ''' </summary>
    Private Class KeyCaptureTotals

        ''' <summary>Total number of keys across all old entries (all types)</summary>
        Public Property Total As Integer

        ''' <summary>Number of keys captured by any new entry (all types)</summary>
        Public Property Captured As Integer

        ''' <summary>Total number of FileKey and RegKey values across all old entries</summary>
        Public Property TotalContent As Integer

        ''' <summary>Number of FileKey and RegKey values captured by any new entry</summary>
        Public Property CapturedContent As Integer

    End Class

    ''' <summary>
    ''' Helper class for tracking key capture per old entry
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

        ''' <summary>Values of FileKey and RegKey keys in the old entry</summary>
        Public Property AllContentKeyValues As HashSet(Of String)

        ''' <summary>Values of FileKey and RegKey keys from the old entry that were captured by any new entry</summary>
        Public Property CapturedContentKeyValues As HashSet(Of String)

    End Class

End Class
