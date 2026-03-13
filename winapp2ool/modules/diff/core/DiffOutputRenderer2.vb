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
''' Formats and renders diff results as <c>MenuSection</c> output for display and logging.
''' Produces the post-diff summary, entry-level change descriptions (additions, removals,
''' renames, mergers), and key-level itemization of added, removed, and updated keys.
''' </summary>
Public Class DiffOutputRenderer2

    Private ReadOnly _state As DiffState
    Private ReadOnly _file1 As iniFile2
    Private ReadOnly _file2 As iniFile2
    Private ReadOnly _keyAnalyzer As KeyModificationAnalyzer2

    ''' <summary>
    ''' Maps merged target entry name → (key value → source old entry name).
    ''' Built by <c>ItemizeMergers</c> and consumed by <c> ItemizeModifications </c>
    ''' to attribute old keys to their source entries in merger output.
    ''' </summary>
    Private ReadOnly _mergerSourceMaps As New Dictionary(Of String, Dictionary(Of String, String))(StringComparer.OrdinalIgnoreCase)


    ''' <summary>
    ''' Initializes a new instance of <c>DiffOutputRenderer2</c>
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
    ''' <param name="keyAnalyzer">
    ''' Used to compute key-level changes for merger and added-with-merger entries
    ''' </param>
    Public Sub New(state As DiffState,
                   file1 As iniFile2,
                   file2 As iniFile2,
                   keyAnalyzer As KeyModificationAnalyzer2)

        _state = state
        _file1 = file1
        _file2 = file2
        _keyAnalyzer = keyAnalyzer

    End Sub

    ''' <summary>
    ''' Records the summary of the diff results and reports them to the user
    ''' </summary>
    '''
    ''' <returns>
    ''' A <c>MenuSection</c> containing the formatted diff summary
    ''' </returns>
    Public Function LogPostDiff() As MenuSection

        Dim stats = _state.Statistics
        Dim merged = _state.MergedEntries
        Dim modified = _state.ModifiedEntries

        Dim netDiff = _file2.Count - _file1.Count
        Dim oldRemovedNoRepl = modified.RemovedEntryNames.Count - stats.MergedEntryCount - merged.RenamedEntryNames.Count

        Dim modifiedEntriesWithMergers = merged.MergedEntryNames.Where(Function(e) Not modified.AddedEntryNames.Contains(e)).Count()

        Dim mergedIntoModifiedSources As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        For Each entryName In merged.MergedEntryNames

            If modified.AddedEntryNames.Contains(entryName) OrElse Not merged.MergeDict.ContainsKey(entryName) Then Continue For

            For Each oldName In merged.MergeDict(entryName) : mergedIntoModifiedSources.Add(oldName) : Next

        Next

        Dim oldEntriesMergedIntoModified = mergedIntoModifiedSources.Count
        Dim netChange = $"Net entry count change: {netDiff}"

        Dim modifiedSummaryOpener = $"Modified entries: {modified.ModifiedEntryNames.Count}"
        Dim modifiedAdded = $" + {stats.ModEntriesAddedKeyTotal} added keys across {stats.ModEntriesAddedKeyEntryCount} entries "
        Dim modifiedRemoved = $" - {stats.ModEntriesRemovedKeysWithoutReplacementTotal} removed keys without replacement across {stats.ModEntriesRemovedKeyEntryCount} entries "
        Dim modifiedUpdated = $" ~ {stats.ModEntriesUpdatedKeyTotal} updated keys replaced {stats.ModEntriesReplacedByUpdateTotal} old keys across {stats.ModEntriesUpdatedKeyEntryCount} entries"

        Dim removedSummary = $"Removed entries: {modified.RemovedEntryNames.Count}"
        Dim removedMergedIntoModified = $" @ {oldEntriesMergedIntoModified} removed entries have been merged into {modifiedEntriesWithMergers} modified entries "
        Dim removedMergedIntoAdded = $" + {stats.AddedWithMergersSourceEntryCount} removed entries have been merged into {stats.AddedWithMergersEntryCount} added entries"
        Dim removedRenamed = $" & {merged.RenamedEntryNames.Count} removed entries have been renamed"
        Dim removedNoReplacement = $" - {oldRemovedNoRepl} entries have been removed without replacement"
        Dim removedReadded = $" + {stats.RemovedByAdditionCount} removed entries have been merged into {stats.AddedEntryWithMergerCount} added entries"

        Dim hasAddedWithMergers = stats.AddedWithMergersEntryCount > 0

        Dim addedMergersSource = $" @ {stats.AddedWithMergersEntryCount} entries consolidate content from {stats.AddedWithMergersSourceEntryCount} removed entries"
        Dim addedMergersNovel = $"    + {stats.AddedWithMergersNovelKeysEntryCount} entries contain {stats.AddedWithMergersNovelKeysTotal} novel keys (not from merged sources)"
        Dim addedMergersCapturing = $"    ~ {stats.AddedWithMergersCapturingEntryCount} entries contain {stats.AddedWithMergersCapturingKeysTotal} keys capturing {stats.AddedWithMergersCapturedKeysTotal} removed keys"
        Dim addedMergersDropped = $"    - {stats.AddedWithMergersDroppedEntryCount} entries dropped {stats.AddedWithMergersDroppedKeysTotal} keys from merged sources"
        Dim addedMergersCarriedOver = $"    = {stats.AddedWithMergersCarriedOverKeysEntryCount} entries contain {stats.AddedWithMergersCarriedOverKeysTotal} keys carried over unchanged from merged sources"

        Dim plainAddedCount = modified.AddedEntryNames.Count - merged.RenamedEntryNames.Count - stats.AddedWithMergersEntryCount

        Dim added = $"Added entries: {modified.AddedEntryNames.Count}"
        Dim addedPlain = $" + {plainAddedCount} novel entries (without merged content)"
        Dim addedRenamed = $" & {merged.RenamedEntryNames.Count} added entries are renamed versions of removed entries and may contain other minor changes"

        Dim modifiedEntriesHaveAdditions = stats.ModEntriesAddedKeyTotal > 0
        Dim modEntriesHaveRemovals = stats.ModEntriesRemovedKeysWithoutReplacementTotal > 0
        Dim modEntriesHaveUpdates = stats.ModEntriesUpdatedKeyTotal > 0
        Dim hasRenames = merged.RenamedEntryNames.Count > 0
        Dim hasMergedIntoAdded = stats.AddedWithMergersSourceEntryCount > 0
        Dim hasMergedIntoModified = modifiedEntriesWithMergers > 0

        Dim renameStats = stats
        Dim renamedNameOnly = $"    = {renameStats.RenamedEntriesNameOnlyCount} are name-only changes (no key differences)"
        Dim renamedAdded = $"    + {renameStats.RenamedEntriesAddedKeyTotal} added keys across {renameStats.RenamedEntriesAddedKeyEntryCount} entries"
        Dim renamedRemoved = $"    - {renameStats.RenamedEntriesRemovedKeyTotal} removed keys across {renameStats.RenamedEntriesRemovedKeyEntryCount} entries"
        Dim renamedUpdated = $"    ~ {renameStats.RenamedEntriesUpdatedKeyTotal} updated keys replaced {renameStats.RenamedEntriesReplacedByUpdateTotal} old keys across {renameStats.RenamedEntriesUpdatedKeyEntryCount} entries"

        Dim out As New MenuSection

        out.AddTopBorder().AddColoredLine("Diff Summary", ConsoleColor.DarkGreen, centered:=True).AddDivider() _
           .AddColoredLine(netChange, ConsoleColor.White) _
           .AddColoredLine(modifiedSummaryOpener, ConsoleColor.Yellow) _
           .AddColoredLine(modifiedAdded, ConsoleColor.Green, condition:=modifiedEntriesHaveAdditions) _
           .AddColoredLine(modifiedRemoved, ConsoleColor.Red, condition:=modEntriesHaveRemovals) _
           .AddColoredLine(modifiedUpdated, ConsoleColor.Yellow, condition:=modEntriesHaveUpdates) _
           .AddColoredLine(removedSummary, ConsoleColor.Cyan) _
           .AddColoredLine(removedMergedIntoModified, ConsoleColor.Cyan, condition:=hasMergedIntoModified) _
           .AddColoredLine(removedMergedIntoAdded, ConsoleColor.Green, condition:=hasMergedIntoAdded) _
           .AddColoredLine(removedReadded, ConsoleColor.Green, condition:=stats.RemovedByAdditionCount > 0) _
           .AddColoredLine(removedRenamed, ConsoleColor.Magenta, condition:=hasRenames) _
           .AddColoredLine(renamedNameOnly, ConsoleColor.Magenta, condition:=hasRenames AndAlso renameStats.RenamedEntriesNameOnlyCount > 0) _
           .AddColoredLine(renamedAdded, ConsoleColor.Green, condition:=hasRenames AndAlso renameStats.RenamedEntriesAddedKeyTotal > 0) _
           .AddColoredLine(renamedRemoved, ConsoleColor.Red, condition:=hasRenames AndAlso renameStats.RenamedEntriesRemovedKeyTotal > 0) _
           .AddColoredLine(renamedUpdated, ConsoleColor.Yellow, condition:=hasRenames AndAlso renameStats.RenamedEntriesUpdatedKeyTotal > 0) _
           .AddColoredLine(removedNoReplacement, ConsoleColor.Red) _
           .AddColoredLine(added, ConsoleColor.DarkGreen) _
           .AddColoredLine(addedMergersSource, ConsoleColor.DarkCyan, condition:=hasAddedWithMergers) _
           .AddColoredLine(addedMergersNovel, ConsoleColor.Green, condition:=stats.AddedWithMergersNovelKeysTotal > 0) _
           .AddColoredLine(addedMergersCarriedOver, ConsoleColor.Cyan, condition:=stats.AddedWithMergersCarriedOverKeysTotal > 0) _
           .AddColoredLine(addedMergersCapturing, ConsoleColor.Yellow, condition:=stats.AddedWithMergersCapturingKeysTotal > 0) _
           .AddColoredLine(addedMergersDropped, ConsoleColor.Red, condition:=stats.AddedWithMergersDroppedKeysTotal > 0) _
           .AddColoredLine(addedPlain, ConsoleColor.Green, condition:=plainAddedCount > 0) _
           .AddColoredLine(addedRenamed, ConsoleColor.Magenta, condition:=hasRenames) _
           .AddBottomBorder()

        gLog("Diff Summary", ascend:=True, leadr:=True, ascAmt:=2)
        gLog(netChange)
        gLog(modifiedSummaryOpener)
        gLog(modifiedAdded, cond:=modifiedEntriesHaveAdditions)
        gLog(modifiedRemoved, cond:=modEntriesHaveRemovals)
        gLog(modifiedUpdated, cond:=modEntriesHaveUpdates)
        gLog(removedSummary)
        gLog(removedMergedIntoModified, cond:=hasMergedIntoModified)
        gLog(removedMergedIntoAdded, cond:=hasMergedIntoAdded)
        gLog(removedReadded, cond:=stats.RemovedByAdditionCount > 0)
        gLog(removedRenamed, cond:=hasRenames)
        gLog(renamedNameOnly, cond:=hasRenames AndAlso renameStats.RenamedEntriesNameOnlyCount > 0)
        gLog(renamedAdded, cond:=hasRenames AndAlso renameStats.RenamedEntriesAddedKeyTotal > 0)
        gLog(renamedRemoved, cond:=hasRenames AndAlso renameStats.RenamedEntriesRemovedKeyTotal > 0)
        gLog(renamedUpdated, cond:=hasRenames AndAlso renameStats.RenamedEntriesUpdatedKeyTotal > 0)
        gLog(removedNoReplacement)
        gLog(added)
        gLog(addedMergersSource, cond:=hasAddedWithMergers)
        gLog(addedMergersNovel, cond:=stats.AddedWithMergersNovelKeysTotal > 0)
        gLog(addedMergersCarriedOver, cond:=stats.AddedWithMergersCarriedOverKeysTotal > 0)
        gLog(addedMergersCapturing, cond:=stats.AddedWithMergersCapturingKeysTotal > 0)
        gLog(addedMergersDropped, cond:=stats.AddedWithMergersDroppedKeysTotal > 0)
        gLog(addedPlain, cond:=plainAddedCount > 0)
        gLog(addedRenamed, cond:=hasRenames)
        gLog("", descend:=True, descAmt:=2)
        gLog("Diff complete", descend:=True)

        Return out

    End Function

    ''' <summary>
    ''' Records each removed entry from the old version which
    ''' has been merged into an entry in the new version
    ''' </summary>
    '''
    ''' <returns>
    ''' One <c>MenuSection</c> per merged old entry
    ''' </returns>
    Public Function SummarizeMergers() As List(Of MenuSection)

        Dim out As New List(Of MenuSection)

        For Each oldEntry In _state.MergedEntries.OldToNewMergeDict.OrderBy(Function(kvp) kvp.Key, StringComparer.OrdinalIgnoreCase)

            Dim oldName = oldEntry.Key
            Dim newTargets = oldEntry.Value

            _state.Statistics.MergedEntryCount += 1

            Dim result As MenuSection

            result = If(newTargets.Count = 1,
                          MakeDiff(_file1.GetSection(oldName), 4, _file2.GetSection(newTargets(0))),
                          MakeDiffMultiTarget(_file1.GetSection(oldName), newTargets))

            out.Add(result)

        Next

        Return out

    End Function

    ''' <summary>
    ''' Creates a diff section for an entry that was
    ''' split/merged into multiple new entries
    ''' </summary>
    '''
    ''' <param name="oldSection">
    ''' The removed entry
    ''' </param>
    '''
    ''' <param name="newTargets">
    ''' List of new entry names that contain keys from the old entry
    ''' </param>
    '''
    ''' <returns>
    ''' A <c>MenuSection</c> listing the target entry names the old entry was split/merged into
    ''' </returns>
    Public Function MakeDiffMultiTarget(oldSection As iniSection2,
                                        newTargets As List(Of String)) As MenuSection

        Dim result = New MenuSection
        Dim changeStr = $"{oldSection.Name} has been split/merged into {newTargets.Count} entries"

        result.AddColoredLine(changeStr, color:=ConsoleColor.Cyan, centered:=True)
        gLog(changeStr, indent:=True, leadr:=True)

        result.AddColoredLine("Merged into:", color:=ConsoleColor.Yellow, centered:=True)

        For Each target In newTargets

            result.AddColoredLine($"  • {target}", color:=ConsoleColor.Magenta, centered:=True)
            gLog($"  • {target}", indent:=True)

        Next

        If Not ShowFullEntries Then Return result

        result.AddBlank()
        result.AddColoredLine("Old entry:", color:=ConsoleColor.DarkRed, centered:=True)
        gLog("Old entry:", leadr:=True)
        BuildEntrySection(result, oldSection.ToString)

        Return result

    End Function

    ''' <summary>
    ''' Records each removed entry from the old version
    ''' which has been given a new name in the new version.
    ''' Only emits entries that are name-only changes (no key differences);
    ''' entries with key-level changes are handled by <c>ItemizeRenameChanges</c>.
    ''' </summary>
    '''
    ''' <returns>
    ''' One <c>MenuSection</c> per name-only renamed entry
    ''' </returns>
    Public Function SummarizeRenames() As List(Of MenuSection)

        Dim out As New List(Of MenuSection)

        For Each entry In _state.MergedEntries.RenamedEntryNames.OrderBy(Function(s) s, StringComparer.OrdinalIgnoreCase)

            Dim hasChanges = (_state.ModifiedEntries.AddedKeyTracker2.ContainsKey(entry) AndAlso
                              _state.ModifiedEntries.AddedKeyTracker2(entry).Count > 0) OrElse
                             (_state.ModifiedEntries.RemovedKeyTracker2.ContainsKey(entry) AndAlso
                              _state.ModifiedEntries.RemovedKeyTracker2(entry).Count > 0)

            ' Only count non-Name updated keys as real changes
            If Not hasChanges AndAlso _state.ModifiedEntries.ModifiedKeyTracker2.ContainsKey(entry) Then

                For Each kvp In _state.ModifiedEntries.ModifiedKeyTracker2(entry)

                    If Not kvp.Key.typeIs("Name") Then hasChanges = True : Exit For

                Next

            End If

            If hasChanges Then Continue For

            Dim oldName = _state.MergedEntries.RenamedEntryPairs(entry)
            out.Add(MakeDiff(_file1.GetSection(oldName), 3, _file2.GetSection(entry)))

        Next

        Return out

    End Function

    ''' <summary>
    ''' Conducts a Diff of each entry detected as containing merged content.
    ''' Builds a combined <c>iniSection2</c> from all contributing old entries and passes it
    ''' directly to <c>FindModifications</c> without any string serialization roundtrip.
    ''' </summary>
    '''
    ''' <returns>
    ''' <c> MenuSection </c>s itemizing key-level changes
    ''' for each modified entry that received merged content
    ''' </returns>
    Public Function ItemizeMergers() As List(Of MenuSection)

        For Each targetEntry In _state.MergedEntries.MergeDict.Keys _
                                    .OrderBy(Function(k) k, StringComparer.OrdinalIgnoreCase)

            If _state.ModifiedEntries.AddedEntryNames.Contains(targetEntry) Then Continue For
            If _state.MergedEntries.RenamedEntryNames.Contains(targetEntry) Then Continue For

            Dim uniqueKeyValues = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            Dim combinedOldKeys As New List(Of iniKey2)
            Dim sourceEntryMap As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

            For Each oldEntryName In _state.MergedEntries.MergeDict(targetEntry)

                Dim oldEnt = _file1.GetSection(oldEntryName)
                If oldEnt Is Nothing Then Continue For

                For Each key In oldEnt.Keys

                    If uniqueKeyValues.Contains(key.Value) Then Continue For

                    combinedOldKeys.Add(key)
                    uniqueKeyValues.Add(key.Value)
                    sourceEntryMap(key.Value) = oldEntryName

                Next

            Next

            ' Include the target's own old keys if it existed in file1 and wasn't already
            ' listed as a merge source (avoids mutating MergeDict as the old code did)
            If _file1.Contains(targetEntry) AndAlso
               Not _state.MergedEntries.MergeDict(targetEntry).Contains(targetEntry) Then

                For Each key In _file1.GetSection(targetEntry).Keys

                    If uniqueKeyValues.Contains(key.Value) Then Continue For

                    combinedOldKeys.Add(key)
                    uniqueKeyValues.Add(key.Value)
                    sourceEntryMap(key.Value) = targetEntry

                Next

            End If

            _mergerSourceMaps(targetEntry) = sourceEntryMap
            _keyAnalyzer.FindModificationsFromCombinedKeys(combinedOldKeys, _file2.GetSection(targetEntry))

        Next

        Return ItemizeModifications(True)

    End Function

    ''' <summary>
    ''' Outputs each added entry and any entries which have been merged into it
    ''' </summary>
    '''
    ''' <returns>
    ''' One <c>MenuSection</c> per added entry (excluding renames and added-with-merger entries)
    ''' </returns>
    Public Function ItemizeAdditions() As List(Of MenuSection)

        Dim results As New List(Of MenuSection)
        Dim lastEntryHadMergers = False

        For Each entry In _state.ModifiedEntries.AddedEntryNames.OrderBy(Function(s) s, StringComparer.OrdinalIgnoreCase)

            If _state.MergedEntries.RenamedEntryNames.Contains(entry) Then Continue For

            If _state.MergedEntries.MergeDict.ContainsKey(entry) Then Continue For

            Dim section = _file2.GetSection(entry)
            Dim diffSection = MakeDiff(section, 0)

            diffSection.AddLine("", condition:=Not lastEntryHadMergers AndAlso
                                          _state.MergedEntries.MergeDict.ContainsKey(section.Name))

            lastEntryHadMergers = _state.MergedEntries.MergeDict.ContainsKey(section.Name)

            If Not lastEntryHadMergers Then results.Add(diffSection) : Continue For

            _state.Statistics.AddedEntryWithMergerCount += 1
            Dim out = "This entry contains keys merged from the following removed entries:  "

            diffSection.AddColoredLine(out, color:=ConsoleColor.Yellow, centered:=True)
            gLog(out, indent:=True)

            For Each mergedEntry In _state.MergedEntries.MergeDict(section.Name)

                _state.Statistics.RemovedByAdditionCount += 1
                diffSection.AddColoredLine(mergedEntry, color:=ConsoleColor.DarkCyan, centered:=True)
                gLog(mergedEntry, indent:=True, indAmt:=4)

            Next

            diffSection.AddBlank()

            results.Add(diffSection)

        Next

        Return results

    End Function

    ''' <summary>
    ''' Itemizes the ways in which a given entry has been modified and outputs them to the user
    ''' </summary>
    '''
    ''' <param name="isMerger">
    ''' Indicates that the current set of entries which have been 
    ''' modified are the product of merging multiple entries together
    ''' </param>
    '''
    ''' <returns>
    ''' <c>MenuSection</c>s itemizing added, removed, and 
    ''' updated keys for each qualifying modified entry
    ''' </returns>
    Public Function ItemizeModifications(Optional isMerger As Boolean = False) As List(Of MenuSection)

        Dim results = New List(Of MenuSection)

        For Each entry In _state.ModifiedEntries.ModifiedEntryNames.OrderBy(Function(s) s, StringComparer.OrdinalIgnoreCase)

            If Not isMerger AndAlso _state.MergedEntries.MergedEntryNames.Contains(entry) Then Continue For
            If isMerger AndAlso Not _state.MergedEntries.MergedEntryNames.Contains(entry) Then Continue For
            If _state.MergedEntries.RenamedEntryNames.Contains(entry) Then Continue For

            Dim addKeyTypes, remKeyTypes, modKeyTypes As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
            Dim newSectionVer = _file2.GetSection(entry)

            Dim addedKeys = If(_state.ModifiedEntries.AddedKeyTracker2.ContainsKey(entry),
                          _state.ModifiedEntries.AddedKeyTracker2(entry), New List(Of iniKey2))

            Dim removedKeys = If(_state.ModifiedEntries.RemovedKeyTracker2.ContainsKey(entry),
                            _state.ModifiedEntries.RemovedKeyTracker2(entry), New List(Of iniKey2))

            Dim updatedKeysDict = If(_state.ModifiedEntries.ModifiedKeyTracker2.ContainsKey(entry),
                                 _state.ModifiedEntries.ModifiedKeyTracker2(entry), New Dictionary(Of iniKey2, List(Of iniKey2)))

            If removedKeys.Count + addedKeys.Count + updatedKeysDict.Count = 0 Then Continue For

            ' Look up source attribution map for merger entries
            Dim sourceMap As Dictionary(Of String, String) = Nothing
            If isMerger Then _mergerSourceMaps.TryGetValue(entry, sourceMap)

            results.Add(MakeDiff(newSectionVer, 2))
            results.AddRange(ItemizeChangesFromList(addedKeys, True, addKeyTypes, sourceMap))
            results.AddRange(ItemizeChangesFromList(removedKeys, False, remKeyTypes, sourceMap))

            results.AddRange(ItemizeUpdatedKeys(updatedKeysDict, addedKeys, removedKeys, modKeyTypes, sourceMap))

            results.Add(ItemizeMergedEntries(entry, isMerger))

        Next

        Return results

    End Function

    ''' <summary>
    ''' Outputs the changes made to the keys within an entry to the user
    ''' </summary>
    '''
    ''' <param name="updatedKeysDict">
    ''' Map of new key → list of old keys it replaced,
    ''' as recorded by <c>KeyModificationAnalyzer2</c>
    ''' </param>
    '''
    ''' <param name="addedKeys">
    ''' Keys that were purely added (not replacements);
    ''' used to determine log indentation
    ''' </param>
    '''
    ''' <param name="removedKeys">
    ''' Keys that were purely removed (not replacements);
    ''' used to determine log indentation
    ''' </param>
    '''
    ''' <param name="modKeyTypes">
    ''' Accumulator dictionary that tracks the count of updated
    ''' keys per key type for the modification summary
    ''' </param>
    '''
    ''' <param name="sourceEntryMap">
    ''' Optional map of key value → source entry name, used
    ''' to attribute old keys to their origin entries in merger output
    ''' </param>
    '''
    ''' <returns>
    ''' <c>MenuSection</c>s describing each key update
    ''' (one header section plus one detail section per updated key)
    ''' </returns>
    Public Function ItemizeUpdatedKeys(updatedKeysDict As Dictionary(Of iniKey2, List(Of iniKey2)),
                                       addedKeys As List(Of iniKey2),
                                       removedKeys As List(Of iniKey2),
                                       modKeyTypes As Dictionary(Of String, Integer),
                              Optional sourceEntryMap As Dictionary(Of String, String) = Nothing) As List(Of MenuSection)

        Dim result As New List(Of MenuSection)

        If updatedKeysDict.Count = 0 Then Return result

        gLog("", ascend:=True, cond:=addedKeys.Count + removedKeys.Count = 0)

        For Each changeList In updatedKeysDict.Values

            recordModification(modKeyTypes, changeList(0).KeyType)

        Next

        result.Add(summarizeEntryUpdate(modKeyTypes, "Modified"))

        For i = 0 To updatedKeysDict.Count - 1

            Dim output As New MenuSection
            Dim newKey = updatedKeysDict.Keys(i)
            Dim oldKeys = updatedKeysDict.Values(i)
            Dim isRename = newKey.typeIs("Name")
            Dim count = updatedKeysDict.Values(i).Count

            Dim outTxt1 = $"{If(isRename, "Entry Name", newKey.Name)} has been modified{If(Not isRename, $", replacing {count} old key{If(count > 1, "s", "")}", "")}"

            output.AddColoredLine(outTxt1, ConsoleColor.Yellow)
            gLog(outTxt1, indent:=True, indAmt:=1, leadr:=i = 0)

            Dim outTxt2 = $" + New: {If(isRename, newKey.Value, newKey.ToString())}"

            output.AddColoredLine(outTxt2, ConsoleColor.Green)
            gLog(outTxt2, indent:=True, indAmt:=4)

            For Each oldKey In oldKeys

                Dim sourceInfo = ""
                If sourceEntryMap IsNot Nothing AndAlso sourceEntryMap.ContainsKey(oldKey.Value) Then
                    sourceInfo = $" (from [{sourceEntryMap(oldKey.Value)}])"
                End If

                Dim old = $" - Old: {If(isRename, oldKey.Value, oldKey.ToString())}{sourceInfo}"

                output.AddColoredLine(old, ConsoleColor.Red)
                gLog(old, indent:=True, indAmt:=4)

            Next

            output.AddBlank()

            result.Add(output)

        Next

        gLog(descend:=True, cond:=addedKeys.Count + removedKeys.Count = 0)

        Return result

    End Function

    ''' <summary>
    ''' Itemizes the names of any removed entries that were merged into <c><paramref name="entry"/></c>
    ''' </summary>
    '''
    ''' <param name="entry">
    ''' The name of the target entry that received merged content
    ''' </param>
    ''' 
    ''' <param name="isMerger">
    ''' When <c>True</c>, the section label describes changes measured against old entries;
    ''' when <c>False</c>, it names the removed entries whose content was merged in
    ''' </param>
    '''
    ''' <returns>
    ''' A <c>MenuSection</c> listing the source entry names,
    ''' or an empty section if <paramref name="entry"/> has no merge sources
    ''' </returns>
    Public Function ItemizeMergedEntries(entry As String, isMerger As Boolean) As MenuSection

        Dim out As New MenuSection
        If Not _state.MergedEntries.MergeDict.ContainsKey(entry) Then Return out

        Dim outTxt = If(Not isMerger, "This entry contains keys merged from the following removed entries",
                                   "The above changes are measured against the following removed/old entries")

        out.AddColoredLine(outTxt, ConsoleColor.Yellow, centered:=True)
        gLog()
        gLog(outTxt, indent:=True)

        For Each mergedEntry In _state.MergedEntries.MergeDict(entry)

            out.AddColoredLine(mergedEntry, ConsoleColor.DarkCyan, centered:=True)
            gLog(mergedEntry, indent:=True)

        Next

        out.AddBlank()

        Return out

    End Function

    ''' <summary>
    ''' Itemizes key-level changes for each renamed entry, pulling from
    ''' the same trackers populated by the rename's <c>FindModifications</c> callback
    ''' </summary>
    '''
    ''' <returns>
    ''' <c>MenuSection</c>s describing added, removed, and updated keys for each rename
    ''' </returns>
    Public Function ItemizeRenameChanges() As List(Of MenuSection)

        Dim results As New List(Of MenuSection)

        For Each newName In _state.MergedEntries.RenamedEntryNames.OrderBy(Function(s) s, StringComparer.OrdinalIgnoreCase)

            Dim addedKeys = If(_state.ModifiedEntries.AddedKeyTracker2.ContainsKey(newName),
                           _state.ModifiedEntries.AddedKeyTracker2(newName), New List(Of iniKey2))

            Dim removedKeys = If(_state.ModifiedEntries.RemovedKeyTracker2.ContainsKey(newName),
                             _state.ModifiedEntries.RemovedKeyTracker2(newName), New List(Of iniKey2))

            Dim rawUpdatedDict = If(_state.ModifiedEntries.ModifiedKeyTracker2.ContainsKey(newName),
                                 _state.ModifiedEntries.ModifiedKeyTracker2(newName), New Dictionary(Of iniKey2, List(Of iniKey2)))

            ' Strip the Name sentinel — the MakeDiff header already shows the rename
            Dim updatedKeysDict As New Dictionary(Of iniKey2, List(Of iniKey2))
            For Each kvp In rawUpdatedDict

                If Not kvp.Key.typeIs("Name") Then updatedKeysDict.Add(kvp.Key, kvp.Value)

            Next

            If removedKeys.Count + addedKeys.Count + updatedKeysDict.Count = 0 Then Continue For

            Dim oldName = _state.MergedEntries.RenamedEntryPairs(newName)
            Dim addKeyTypes, remKeyTypes, modKeyTypes As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)

            results.Add(MakeDiff(_file1.GetSection(oldName), 3, _file2.GetSection(newName)))
            results.AddRange(ItemizeChangesFromList(addedKeys, True, addKeyTypes, Nothing))
            results.AddRange(ItemizeChangesFromList(removedKeys, False, remKeyTypes, Nothing))
            results.AddRange(ItemizeUpdatedKeys(updatedKeysDict, addedKeys, removedKeys, modKeyTypes))

        Next

        Return results

    End Function

    ''' <summary>
    ''' Outputs the details of a modified entry's changes to the user
    ''' </summary>
    '''
    ''' <param name="section">
    ''' The entry being described (the old version for renames/mergers)
    ''' </param>
    '''
    ''' <param name="changeType">
    ''' Change category: 0 = added, 1 = removed, 2 = modified, 3 = renamed to, 4 = merged into
    ''' </param>
    '''
    ''' <param name="newSection">
    ''' The new version of the entry; required when
    ''' <paramref name="changeType"/> is 3 (rename) or 4 (merge)
    ''' </param>
    '''
    ''' <returns>
    ''' A <c>MenuSection</c> describing the entry's change
    ''' </returns>
    Public Function MakeDiff(section As iniSection2,
                             changeType As Integer,
                    Optional newSection As iniSection2 = Nothing) As MenuSection

        Dim result = New MenuSection
        Dim printColor As ConsoleColor = ConsoleColor.Cyan

        If changeType = 2 OrElse changeType = 3 Then printColor = If(changeType = 2, ConsoleColor.Yellow, ConsoleColor.Magenta)

        Dim renamedOrMergedEntryName = If(newSection IsNot Nothing, newSection.Name, "")
        Dim changeTypeStrs = {"added", "removed", "modified", "renamed to ", "merged into "}
        Dim changeStr = $"{section.Name} has been {changeTypeStrs(changeType)}{renamedOrMergedEntryName}"

        result.AddColoredLine(changeStr, color:=If(changeType >= 2, printColor, If(changeType < 1, ConsoleColor.Green, ConsoleColor.Red)), centered:=True)
        gLog(changeStr, indent:=True, leadr:=True)

        If Not ShowFullEntries Then Return result

        ' Everything below this point prints to the user only if Verbose Mode is enabled
        ' in the Diff settings
        Dim isMergeOrRenamed = changeType >= 3 AndAlso changeType < 5

        If isMergeOrRenamed Then

            result.AddBlank()
            result.AddColoredLine("Old entry:", color:=ConsoleColor.DarkRed, centered:=True)

            gLog()
            gLog("Old entry:", leadr:=True)

        End If

        BuildEntrySection(result, section.ToString)

        If Not isMergeOrRenamed Then Return result

        Dim out = If(changeType = 3, "Renamed entry: ", "Merged entry: ")

        result.AddBlank()
        result.AddColoredLine(out, color:=printColor)

        gLog(out, leadr:=True)
        BuildEntrySection(result, newSection.ToString)

        Return result

    End Function

    ''' <summary>
    ''' Appends each line of an entry string to the given <c>MenuSection</c>
    ''' </summary>
    '''
    ''' <param name="section">
    ''' The <c>MenuSection</c> to append lines to
    ''' </param>
    ''' 
    ''' <param name="entry">
    ''' The string representation of the entry, split on <c>vbCrLf</c>
    ''' </param>
    Public Sub BuildEntrySection(ByRef section As MenuSection,
                                       entry As String)

        Dim splitEntry = entry.Split(CChar(vbCrLf))

        For i = 0 To splitEntry.Length - 1

            Dim line = splitEntry(i).Replace(vbLf, "")
            section.AddLine(line)
            gLog(line, indAmt:=4)

        Next

    End Sub

    ''' <summary>
    ''' Outputs details of keys that moved between entries
    ''' </summary>
    '''
    ''' <returns>
    ''' <c>MenuSection</c>s grouped by source entry, each listing
    ''' the keys that moved and their destination entries;
    ''' empty if no movements were detected
    ''' </returns>
    Public Function ItemizeKeyMovements() As List(Of MenuSection)

        Dim results As New List(Of MenuSection)

        If _state.KeyMovements.MovedKeys.Count = 0 Then Return results

        ' Group movements by source entry
        Dim movementsBySource As New Dictionary(Of String, List(Of KeyMovementDetail))(StringComparer.OrdinalIgnoreCase)

        For Each kvp In _state.KeyMovements.MovedKeys

            Dim parts = kvp.Key.Split(MovementKeySeparator)
            If parts.Length < 3 Then Continue For

            Dim keyName = parts(0)
            Dim keyValue = parts(1)
            Dim movementInfo = kvp.Value
            Dim source = movementInfo.SourceEntry
            Dim target = movementInfo.TargetEntry

            If Not movementsBySource.ContainsKey(source) Then movementsBySource(source) = New List(Of KeyMovementDetail)

            movementsBySource(source).Add(New KeyMovementDetail(keyName, keyValue, target))

        Next

        ' Output movements grouped by source entry
        For Each sourceEntry In movementsBySource.OrderBy(Function(kvp) kvp.Key, StringComparer.OrdinalIgnoreCase)

            Dim out As New MenuSection
            out.AddColoredLine($"{sourceEntry.Key} - Keys moved to other entries:", ConsoleColor.Cyan, centered:=True)
            gLog($"{sourceEntry.Key} - Keys moved to other entries:", indent:=True, leadr:=True)

            For Each movement In sourceEntry.Value

                Dim line = $"  → {movement.KeyName}={movement.KeyValue} moved to [{movement.Target}]"
                out.AddColoredLine(line, ConsoleColor.DarkCyan)
                gLog(line, indent:=True, indAmt:=4)

            Next

            out.AddBlank()
            results.Add(out)

        Next

        Return results

    End Function

    ''' <summary>
    ''' Outputs detailed information for added entries 
    ''' that contain merged content from removed entries.
    ''' Builds combined old entry sections directly from 
    ''' <c>iniKey2</c> objects 
    ''' </summary>
    '''
    ''' <returns>
    ''' <c>MenuSection</c>s describing each added-with-merger entry:
    ''' header, source list, novel/dropped/capturing key breakdowns
    ''' </returns>
    Public Function ItemizeAddedEntriesWithMergers() As List(Of MenuSection)

        Dim results As New List(Of MenuSection)
        Dim processedOldEntries As New HashSet(Of String)

        For Each entry In _state.ModifiedEntries.AddedEntryNames.OrderBy(Function(s) s, StringComparer.OrdinalIgnoreCase)

            ' Skip renames - they're handled separately
            If _state.MergedEntries.RenamedEntryNames.Contains(entry) Then Continue For

            ' Skip entries without merged content - handled by ItemizeAdditions
            If Not _state.MergedEntries.MergeDict.ContainsKey(entry) Then Continue For

            Dim section = _file2.GetSection(entry)
            Dim mergedCount = _state.MergedEntries.MergeDict(entry).Count

            ' Build combined old key list directly — avoids iniKeyCollection name deduplication
            ' dropping keys that share a name across different source entries (e.g. two FileKey1 values)
            Dim uniqueKeyValues = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            Dim combinedOldKeys As New List(Of iniKey2)

            Dim sourceEntryMap As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

            For i = 0 To _state.MergedEntries.MergeDict(entry).Count - 1

                Dim oldEntryName = _state.MergedEntries.MergeDict(entry)(i)
                If Not _file1.Contains(oldEntryName) Then Continue For

                Dim oldEnt = _file1.GetSection(oldEntryName)

                For Each key In oldEnt.Keys

                    If uniqueKeyValues.Contains(key.Value) Then Continue For

                    combinedOldKeys.Add(key)
                    uniqueKeyValues.Add(key.Value)
                    sourceEntryMap(key.Value) = oldEntryName

                Next

                processedOldEntries.Add(oldEntryName)

            Next

            ' Compute the modifications between combined old entry and new entry
            _keyAnalyzer.FindModificationsForAddedEntryFromKeys(combinedOldKeys, _file2.GetSection(entry))

            ' Header: Entry added with merged content
            Dim headerSection As New MenuSection
            Dim headerText = $"{entry} has been added (consolidating {mergedCount} removed entr{If(mergedCount = 1, "y", "ies")})"

            headerSection.AddColoredLine(headerText, ConsoleColor.Green, centered:=True)
            gLog(headerText, indent:=True, leadr:=True)
            results.Add(headerSection)

            ' List merged entries
            Dim mergedListSection As New MenuSection
            mergedListSection.AddColoredLine("Merged from:", ConsoleColor.Yellow, centered:=True)
            gLog("Merged from:", indent:=True)

            For Each mergedEntry In _state.MergedEntries.MergeDict(entry)

                mergedListSection.AddColoredLine($"{mergedEntry}", ConsoleColor.DarkCyan, centered:=True)
                gLog($"  • {mergedEntry}", indent:=True, indAmt:=4)

            Next

            mergedListSection.AddBlank()
            results.Add(mergedListSection)

            ' Show key changes from tracked data
            Dim addedKeys = If(_state.ModifiedEntries.AddedKeyTracker2.ContainsKey(entry),
                          _state.ModifiedEntries.AddedKeyTracker2(entry), New List(Of iniKey2))

            Dim removedKeys = If(_state.ModifiedEntries.RemovedKeyTracker2.ContainsKey(entry),
                            _state.ModifiedEntries.RemovedKeyTracker2(entry), New List(Of iniKey2))

            Dim updatedKeysDict = If(_state.ModifiedEntries.ModifiedKeyTracker2.ContainsKey(entry),
                                 _state.ModifiedEntries.ModifiedKeyTracker2(entry),
                                 New Dictionary(Of iniKey2, List(Of iniKey2)))

            ' Build value+type sets for membership checks
            Dim addedKeyIds As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            For Each k In addedKeys : addedKeyIds.Add($"{k.KeyType}|{k.Value}") : Next


            Dim removedKeyIds As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            For Each k In removedKeys : removedKeyIds.Add($"{k.KeyType}|{k.Value}") : Next

            Dim carriedOverKeys As New List(Of iniKey2)
            For Each key In _file2.GetSection(entry).Keys

                Dim keyId = $"{key.KeyType}|{key.Value}"
                If addedKeyIds.Contains(keyId) OrElse removedKeyIds.Contains(keyId) Then Continue For

                ' Skip if it's in updated keys (modified) — compare by value
                Dim isUpdated = False
                For Each kvp In updatedKeysDict

                    If kvp.Key.Value = key.Value Then isUpdated = True : Exit For

                Next

                If isUpdated Then Continue For

                ' If it's in sourceEntryMap, it was carried over from a merged source
                If sourceEntryMap.ContainsKey(key.Value) Then carriedOverKeys.Add(key)

            Next

            addedKeys.AddRange(carriedOverKeys)

            If addedKeys.Count + removedKeys.Count + updatedKeysDict.Count > 0 Then

                Dim addKeyTypes, remKeyTypes, modKeyTypes As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)

                ' New keys in the added entry (not from merged sources)
                If addedKeys.Count > 0 Then

                    Dim newKeysSection As New MenuSection
                    Dim novelKeysMsg = $"{addedKeys.Count} keys added or carried over from merged sources:"
                    newKeysSection.AddColoredLine(novelKeysMsg, ConsoleColor.Green, centered:=True)
                    gLog()
                    gLog(novelKeysMsg, indent:=True)
                    results.Add(newKeysSection)
                    results.AddRange(ItemizeChangesFromList(addedKeys, True, addKeyTypes, sourceEntryMap))

                End If

                ' Keys from merged entries that were not carried over into this added entry
                If removedKeys.Count > 0 Then

                    Dim droppedSection As New MenuSection

                    Dim KeysNotMergedMsg = $"{removedKeys.Count} keys from merged entries not in this entry:"
                    droppedSection.AddColoredLine(KeysNotMergedMsg, ConsoleColor.DarkYellow, centered:=True)
                    gLog()
                    gLog(KeysNotMergedMsg, indent:=True)
                    results.Add(droppedSection)
                    results.AddRange(ItemizeChangesFromList(removedKeys, False, remKeyTypes, sourceEntryMap))

                End If

                ' Keys that captured/replaced multiple old keys
                If updatedKeysDict.Count > 0 Then

                    Dim capturedSection As New MenuSection
                    Dim capturedMsg = $"{updatedKeysDict.Count} keys capturing content from merged entries"
                    capturedSection.AddColoredLine(capturedMsg, ConsoleColor.Yellow, centered:=True)
                    gLog()
                    gLog(capturedMsg, indent:=True)
                    results.Add(capturedSection)
                    results.AddRange(ItemizeUpdatedKeys(updatedKeysDict, addedKeys, removedKeys, modKeyTypes, sourceEntryMap))

                End If

            End If

            Dim spacer As New MenuSection
            spacer.AddBlank()
            results.Add(spacer)

        Next

        Return results

    End Function

    ''' <summary>
    ''' Increments the change count for <paramref name="keyType"/> in <paramref name="ktDict"/>,
    ''' inserting a zero-initialized entry first if the key is not yet present
    ''' </summary>
    ''' <param name="ktDict">Accumulator dictionary mapping key type to change count</param>
    ''' <param name="keyType">The key type whose count should be incremented</param>
    Private Sub recordModification(ktDict As Dictionary(Of String, Integer), keyType As String)

        If Not ktDict.ContainsKey(keyType) Then ktDict(keyType) = 0
        ktDict(keyType) += 1

    End Sub

    ''' <summary>
    ''' Creates a summary section for a modified entry detailing the number of
    ''' added, removed, or updated keys of each type
    ''' </summary>
    '''
    ''' <param name="ktDict">
    ''' Map of key type → count of changes of that type for the current entry
    ''' </param>
    '''
    ''' <param name="changeType">
    ''' The type of change being summarized (e.g., "Added", "Removed", "Modified")
    ''' </param>
    '''
    ''' <returns>
    ''' A <c>MenuSection</c> containing one line per key type summarizing the count of changes
    ''' </returns>
    Private Function summarizeEntryUpdate(ktDict As Dictionary(Of String, Integer), changeType As String) As MenuSection

        Dim result As New MenuSection
        Dim total = ktDict.Count
        Dim i = 0

        For Each kvp In ktDict

            Dim out = $"{changeType} {kvp.Value} {kvp.Key}{If(kvp.Value > 1, "s", "")}"
            result.AddColoredLine(out, ConsoleColor.Yellow, centered:=True)
            result.AddBlank(i = total - 1)

            gLog(out, indent:=True, leadr:=i = 0)
            i += 1

        Next

        Return result

    End Function

    ''' <summary>
    ''' Prints any added or removed keys from an updated entry to the user
    ''' </summary>
    '''
    ''' <param name="kl">
    ''' List of added or removed keys to itemize
    ''' </param>
    '''
    ''' <param name="wasAdded">
    ''' <c>True</c> if the keys in <paramref name="kl"/> were added; <br />
    ''' <c>False</c> if they were removed
    ''' </param>
    '''
    ''' <param name="ktDict">
    ''' Accumulator dictionary that tracks the count of 
    ''' changed keys per key type for the modification summary
    ''' </param>
    '''
    ''' <param name="sourceEntryMap">
    ''' Optional map of key value → source entry name used to annotate 
    ''' merger origin; novel keys are labeled "(novel)" when present
    ''' </param>
    Private Function ItemizeChangesFromList(kl As List(Of iniKey2),
                                            wasAdded As Boolean,
                                            ktDict As Dictionary(Of String, Integer),
                                   Optional sourceEntryMap As Dictionary(Of String, String) = Nothing) As List(Of MenuSection)

        Dim out As New List(Of MenuSection)

        If kl.Count = 0 Then Return out

        gLog(Nothing, ascend:=True)

        Dim changeTxt = If(wasAdded, "Added", "Removed")

        kl.ForEach(Sub(key) recordModification(ktDict, key.KeyType))

        out.Add(summarizeEntryUpdate(ktDict, changeTxt))

        Dim result As New MenuSection

        For i = 0 To kl.Count - 1

            Dim key = kl(i).ToString()
            Dim color = If(wasAdded, ConsoleColor.Green, ConsoleColor.Red)

            ' Added keys either come from a source (mergers, renames) or are novel
            Dim sourceInfo = ""
            If sourceEntryMap IsNot Nothing AndAlso sourceEntryMap.ContainsKey(kl(i).Value) Then

                sourceInfo = $" (from [{sourceEntryMap(kl(i).Value)}])"

            ElseIf wasAdded AndAlso sourceEntryMap IsNot Nothing Then

                sourceInfo = " (novel)"

            End If

            result.AddColoredLine(key & sourceInfo, color)
            gLog(key & sourceInfo, indent:=True, indAmt:=4)

        Next

        result.AddBlank()
        out.Add(result)

        gLog(Nothing, descend:=True)

        Return out

    End Function

    ''' <summary>
    ''' Helper class for displaying key movement details
    ''' </summary>
    Private Class KeyMovementDetail

        ''' <summary>
        ''' The name portion of the moved key (the part before '=' on disk)
        ''' </summary>
        Public Property KeyName As String

        ''' <summary>
        ''' The value of the moved key (everything after the '=' on disk)
        ''' </summary>
        Public Property KeyValue As String

        ''' <summary>
        ''' The name of the entry this key was moved into
        ''' </summary>
        Public Property Target As String

        ''' <summary>
        ''' Creates a new KeyMovementDetail with the given key name, value, and target entry
        ''' </summary>
        ''' 
        ''' <param name="name">
        ''' The name portion of the moved key (the part before '=' on disk)
        ''' </param>
        ''' 
        ''' <param name="value">
        ''' The value of the moved key (everything after the '=' on disk)
        ''' </param>
        ''' 
        ''' <param name="target">
        ''' The name of the entry this key was moved into
        ''' </param>
        Public Sub New(name As String, value As String, target As String)

            KeyName = name
            KeyValue = value
            Me.Target = target

        End Sub

    End Class

End Class
