Option Strict On
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

Imports System.Linq.Expressions

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
    ''' Label used in place of a key type when itemizing a detection-criteria modification —
    ''' a change to an entry's <c>Detect</c>/<c>DetectFile</c> keys, which are reported as a
    ''' single conceptual "detection criteria" change rather than by individual key type
    ''' </summary>
    Private Const DetectionCriteriaLabel As String = "Detection criteria"

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
        Dim oldRemovedNoRepl = modified.RemovedEntryNames.Count - merged.OldToNewMergeDict.Count - merged.RenamedEntryNames.Count

        Dim modifiedEntriesWithMergers = merged.MergedEntryNames.Where(Function(e) Not modified.AddedEntryNames.Contains(e)).Count()

        Dim mergedIntoModifiedSources As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        For Each entryName In merged.MergedEntryNames

            If modified.AddedEntryNames.Contains(entryName) OrElse Not merged.MergeDict.ContainsKey(entryName) Then Continue For

            For Each oldName In merged.MergeDict(entryName) : mergedIntoModifiedSources.Add(oldName) : Next

        Next

        Dim oldEntriesMergedIntoModified = mergedIntoModifiedSources.Count
        Dim netChange = $"Net entry count change: {_file1.Count} → {_file2.Count} ({If(netDiff >= 0, "+", "")}{netDiff})"

        Dim modifiedSummaryOpener = $"Modified entries: {modified.ModifiedEntryNames.Count}"
        Dim modifiedAdded = $" + {stats.ModEntriesAddedKeyTotal} added keys across {stats.ModEntriesAddedKeyEntryCount} entries "
        Dim modifiedRemoved = $" - {stats.ModEntriesRemovedKeysWithoutReplacementTotal} removed keys without replacement across {stats.ModEntriesRemovedKeyEntryCount} entries "
        Dim modifiedUpdated = $" ~ {stats.ModEntriesUpdatedKeyTotal} updated keys replaced {stats.ModEntriesReplacedByUpdateTotal} old keys across {stats.ModEntriesUpdatedKeyEntryCount} entries"
        Dim movedKeys = $" ~ {stats.ModEntriesMovedKeysTotal} keys moved from {stats.ModEntriesMovedKeysSourceCount} entr{If(stats.ModEntriesMovedKeysSourceCount = 1, "y", "ies")} into {stats.ModEntriesMovedKeysTargetCount} entr{If(stats.ModEntriesMovedKeysTargetCount = 1, "y", "ies")}"
        Dim modifiedMergerNote = $" + {modifiedEntriesWithMergers} entries also received merged content from removed entries (see merged entries below)"

        Dim removedSummary = $"Removed entries: {modified.RemovedEntryNames.Count}"
        Dim removedMergedTotal = merged.OldToNewMergeDict.Count
        Dim removedMergedSummary = $" @ {removedMergedTotal} removed entries have been merged into other entries"
        Dim removedMergedIntoModified = $"    @ {oldEntriesMergedIntoModified} merged into {modifiedEntriesWithMergers} modified entries"
        Dim removedMergedIntoAdded = $"    + {stats.AddedWithMergersSourceEntryCount} merged into {stats.AddedWithMergersEntryCount} added entries"
        Dim removedRenamed = $" & {merged.RenamedEntryNames.Count} removed entries have been renamed"
        Dim removedNoReplacement = $" - {oldRemovedNoRepl} entries have been removed without replacement"
        Dim hasAddedWithMergers = stats.AddedWithMergersEntryCount > 0
        Dim hasMerged = removedMergedTotal > 0

        Dim addedMergersSource = $" @ {stats.AddedWithMergersEntryCount} entries consolidate content from {stats.AddedWithMergersSourceEntryCount} removed entries"
        Dim addedMergersNovel = $"    + {stats.AddedWithMergersNovelKeysEntryCount} entries contain {stats.AddedWithMergersNovelKeysTotal} novel keys (not from merged sources)"
        Dim addedMergersCapturing = $"    ~ {stats.AddedWithMergersCapturingEntryCount} entries contain {stats.AddedWithMergersCapturingKeysTotal} keys capturing {stats.AddedWithMergersCapturedKeysTotal} removed keys"
        Dim addedMergersDropped = $"    - {stats.AddedWithMergersDroppedEntryCount} entries dropped {stats.AddedWithMergersDroppedKeysTotal} keys from merged sources"
        Dim addedMergersCarriedOver = $"    = {stats.AddedWithMergersCarriedOverKeysEntryCount} entries contain {stats.AddedWithMergersCarriedOverKeysTotal} keys carried over unchanged from merged sources"

        Dim plainAddedCount = modified.AddedEntryNames.Where(
            Function(e) Not merged.RenamedEntryNames.Contains(e) AndAlso
                        Not merged.MergeDict.ContainsKey(e)).Count()

        Dim added = $"Added entries: {modified.AddedEntryNames.Count}"
        Dim addedPlain = $" + {plainAddedCount} novel entries (without merged content)"
        Dim renamedInAddedCount = merged.RenamedEntryNames.Where(Function(e) modified.AddedEntryNames.Contains(e)).Count()
        Dim addedRenamed = $" & {renamedInAddedCount} added entries are renamed versions of removed entries and may contain other minor changes"

        Dim hasNewBrowsers = stats.NewBrowserSectionValues.Count > 0
        Dim newBrowserSummary = $" + {stats.NewBrowserSectionValues.Count} new browser{If(stats.NewBrowserSectionValues.Count > 1, "s", "")} added"
        Dim hasRemovedBrowsers = stats.RemovedBrowserSectionValues.Count > 0
        Dim removedBrowserSummary = $" - {stats.RemovedBrowserSectionValues.Count} browser{If(stats.RemovedBrowserSectionValues.Count > 1, "s", "")} removed"

        Dim modifiedEntriesHaveAdditions = stats.ModEntriesAddedKeyTotal > 0
        Dim modEntriesHaveRemovals = stats.ModEntriesRemovedKeysWithoutReplacementTotal > 0
        Dim modEntriesHaveUpdates = stats.ModEntriesUpdatedKeyTotal > 0
        Dim hasMovedKeys = stats.ModEntriesMovedKeysTotal > 0
        Dim hasRenames = merged.RenamedEntryNames.Count > 0
        Dim hasMergedIntoAdded = stats.AddedWithMergersSourceEntryCount > 0
        Dim hasMergedIntoModified = modifiedEntriesWithMergers > 0

        Dim renameStats = stats
        Dim renamedNameOnly = $"    = {renameStats.RenamedEntriesNameOnlyCount} are name-only changes (no key differences)"
        Dim renamedAdded = $"    + {renameStats.RenamedEntriesAddedKeyTotal} added keys across {renameStats.RenamedEntriesAddedKeyEntryCount} entries"
        Dim renamedRemoved = $"    - {renameStats.RenamedEntriesRemovedKeyTotal} removed keys across {renameStats.RenamedEntriesRemovedKeyEntryCount} entries"
        Dim renamedUpdated = $"    ~ {renameStats.RenamedEntriesUpdatedKeyTotal} updated keys replaced {renameStats.RenamedEntriesReplacedByUpdateTotal} old keys across {renameStats.RenamedEntriesUpdatedKeyEntryCount} entries"

        Dim out As New MenuSection
        out.AddTopBorder().AddColoredLine("Diff Summary", ConsoleColor.DarkGreen, centered:=True).AddDivider()

        gLog(Nothing, leadr:=True)

        Using gLogScope("Diff Summary")

            Emit(out, netChange, ConsoleColor.White)
            Emit(out, newBrowserSummary, ConsoleColor.Cyan, hasNewBrowsers)
            Emit(out, removedBrowserSummary, ConsoleColor.Red, hasRemovedBrowsers)
            Emit(out, modifiedSummaryOpener, ConsoleColor.Yellow)
            Emit(out, modifiedAdded, ConsoleColor.Green, modifiedEntriesHaveAdditions)
            Emit(out, modifiedRemoved, ConsoleColor.Red, modEntriesHaveRemovals)
            Emit(out, modifiedUpdated, ConsoleColor.Yellow, modEntriesHaveUpdates)
            Emit(out, movedKeys, ConsoleColor.Cyan, hasMovedKeys)
            Emit(out, modifiedMergerNote, ConsoleColor.DarkCyan, hasMergedIntoModified)
            Emit(out, removedSummary, ConsoleColor.Cyan)
            Emit(out, removedMergedSummary, ConsoleColor.Cyan, hasMerged)
            Emit(out, removedMergedIntoModified, ConsoleColor.Cyan, hasMerged AndAlso hasMergedIntoModified)
            Emit(out, removedMergedIntoAdded, ConsoleColor.Green, hasMerged AndAlso hasMergedIntoAdded)
            Emit(out, removedRenamed, ConsoleColor.Magenta, hasRenames)
            Emit(out, renamedNameOnly, ConsoleColor.Magenta, hasRenames AndAlso renameStats.RenamedEntriesNameOnlyCount > 0)
            Emit(out, renamedAdded, ConsoleColor.Green, hasRenames AndAlso renameStats.RenamedEntriesAddedKeyTotal > 0)
            Emit(out, renamedRemoved, ConsoleColor.Red, hasRenames AndAlso renameStats.RenamedEntriesRemovedKeyTotal > 0)
            Emit(out, renamedUpdated, ConsoleColor.Yellow, hasRenames AndAlso renameStats.RenamedEntriesUpdatedKeyTotal > 0)
            Emit(out, removedNoReplacement, ConsoleColor.Red)
            Emit(out, added, ConsoleColor.DarkGreen)
            Emit(out, addedMergersSource, ConsoleColor.DarkCyan, hasAddedWithMergers)
            Emit(out, addedMergersNovel, ConsoleColor.Green, stats.AddedWithMergersNovelKeysTotal > 0)
            Emit(out, addedMergersCarriedOver, ConsoleColor.Cyan, stats.AddedWithMergersCarriedOverKeysTotal > 0)
            Emit(out, addedMergersCapturing, ConsoleColor.Yellow, stats.AddedWithMergersCapturingKeysTotal > 0)
            Emit(out, addedMergersDropped, ConsoleColor.Red, stats.AddedWithMergersDroppedKeysTotal > 0)
            Emit(out, addedPlain, ConsoleColor.Green, plainAddedCount > 0)
            Emit(out, addedRenamed, ConsoleColor.Magenta, renamedInAddedCount > 0)

            gLog("")

        End Using

        out.AddBottomBorder()

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
        Dim mergeCount = _state.MergedEntries.OldToNewMergeDict.Count

        If mergeCount = 0 Then Return out
        Dim mergeHeader = $"{mergeCount} {If(mergeCount = 1, "entry", "entries")} merged or split into other entries"
        Dim mergeHeaderSection As New MenuSection
        mergeHeaderSection.AddColoredLine(mergeHeader, ConsoleColor.Cyan, centered:=True).AddDivider(solid:=False)
        out.Add(mergeHeaderSection)

        gLog(Nothing, leadr:=True)

        Using gLogScope()

            For Each oldEntry In _state.MergedEntries.OldToNewMergeDict.OrderBy(Function(kvp) kvp.Key, StringComparer.OrdinalIgnoreCase)

                Dim oldName = oldEntry.Key
                Dim newTargets = oldEntry.Value

                Dim result As MenuSection

                result = If(newTargets.Count = 1,
                              MakeDiff(_file1.GetSection(oldName), 4, _file2.GetSection(newTargets(0))),
                              MakeDiffMultiTarget(_file1.GetSection(oldName), newTargets))

                out.Add(result)

            Next

        End Using

        gLog(mergeHeader, leadr:=True)

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
        gLog($"  {changeStr}", leadr:=True)

        result.AddColoredLine("Merged into:", color:=ConsoleColor.Yellow, centered:=True)

        For Each target In newTargets

            result.AddColoredLine($"{target}", color:=ConsoleColor.Magenta, centered:=True)
            gLog($"    • {target}")

        Next

        If Not ShowFullEntries Then Return result

        result.AddBlank()
        result.AddColoredLine("Old entry:", color:=ConsoleColor.DarkRed, centered:=True)
        gLog("Old entry:", leadr:=True)
        BuildEntrySection(result, oldSection.ToString)

        For Each target In newTargets

            Dim targetSection = _file2.GetSection(target)
            If targetSection Is Nothing Then Continue For

            result.AddBlank()
            result.AddColoredLine($"Merged entry: {target}", color:=ConsoleColor.Magenta, centered:=True)
            gLog($"Merged entry: {target}", leadr:=True)
            BuildEntrySection(result, targetSection.ToString)

        Next

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

        Dim qualifying As New List(Of String)

        For Each entry In _state.MergedEntries.RenamedEntryNames.OrderBy(Function(s) s, StringComparer.OrdinalIgnoreCase)

            Dim hasChanges = (_state.ModifiedEntries.AddedKeyTracker2.ContainsKey(entry) AndAlso
                              _state.ModifiedEntries.AddedKeyTracker2(entry).Count > 0) OrElse
                             (_state.ModifiedEntries.RemovedKeyTracker2.ContainsKey(entry) AndAlso
                              _state.ModifiedEntries.RemovedKeyTracker2(entry).Count > 0)

            If Not hasChanges AndAlso _state.ModifiedEntries.ModifiedKeyTracker2.ContainsKey(entry) Then

                For Each kvp In _state.ModifiedEntries.ModifiedKeyTracker2(entry)

                    If Not kvp.Key.typeIs("Name") Then hasChanges = True : Exit For

                Next

            End If

            If Not hasChanges Then qualifying.Add(entry)

        Next

        If qualifying.Count = 0 Then Return out

        Dim header = $"{qualifying.Count} {If(qualifying.Count = 1, "entry", "entries")} renamed (name-only changes)"
        Dim headerSection As New MenuSection
        headerSection.AddColoredLine(header, ConsoleColor.Magenta, centered:=True).AddDivider(solid:=False)
        out.Add(headerSection)

        gLog(Nothing)

        Using gLogScope()

            For Each entry In qualifying

                Dim oldName = _state.MergedEntries.RenamedEntryPairs(entry)
                out.Add(MakeDiff(_file1.GetSection(oldName), 3, _file2.GetSection(entry)))

            Next

        End Using

        gLog(header, leadr:=True)

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

        Dim qualifying = _state.MergedEntries.MergeDict.Keys _
            .OrderBy(Function(k) k, StringComparer.OrdinalIgnoreCase) _
            .Where(Function(targetEntry) Not _state.ModifiedEntries.AddedEntryNames.Contains(targetEntry) AndAlso
                                          Not _state.MergedEntries.RenamedEntryNames.Contains(targetEntry)) _
            .ToList()

        For Each targetEntry In qualifying

            Dim combined = BuildCombinedOldKeys(_state.MergedEntries.MergeDict(targetEntry), targetEntry)
            _mergerSourceMaps(targetEntry) = combined.SourceEntryMap
            _keyAnalyzer.FindModificationsFromCombinedKeys(combined.Keys, _file2.GetSection(targetEntry))

        Next

        If qualifying.Count = 0 Then Return ItemizeModifications(True)

        Dim header = $"{qualifying.Count} modified {If(qualifying.Count = 1, "entry", "entries")} incorporating merged content"
        Dim headerSection As New MenuSection
        headerSection.AddColoredLine(header, ConsoleColor.DarkCyan, centered:=True).AddDivider(solid:=False)

        Dim results As New List(Of MenuSection)
        results.Add(headerSection)

        gLog(Nothing)

        Using gLogScope()

            results.AddRange(ItemizeModifications(True))

        End Using

        gLog(header, leadr:=True)

        Return results

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

        Dim qualifying = _state.ModifiedEntries.AddedEntryNames _
            .OrderBy(Function(s) s, StringComparer.OrdinalIgnoreCase) _
            .Where(Function(entry) Not _state.MergedEntries.RenamedEntryNames.Contains(entry) AndAlso
                                    Not _state.MergedEntries.MergeDict.ContainsKey(entry)) _
            .ToList()

        If qualifying.Count = 0 Then Return results

        Dim addHeader = $"{qualifying.Count} novel {If(qualifying.Count = 1, "entry", "entries")} added"
        Dim addHeaderSection As New MenuSection
        addHeaderSection.AddColoredLine(addHeader, ConsoleColor.DarkGreen, centered:=True).AddDivider(solid:=False)
        results.Add(addHeaderSection)

        gLog(Nothing, leadr:=True)

        Using gLogScope("Added entries:")

            For Each entry In qualifying

                results.Add(MakeDiff(_file2.GetSection(entry), 0))

            Next

        End Using

        gLog(addHeader, leadr:=True)

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

        Dim qualifying As New List(Of String)

        For Each entry In _state.ModifiedEntries.ModifiedEntryNames.OrderBy(Function(s) s, StringComparer.OrdinalIgnoreCase)

            If Not isMerger AndAlso _state.MergedEntries.MergedEntryNames.Contains(entry) Then Continue For
            If isMerger AndAlso Not _state.MergedEntries.MergedEntryNames.Contains(entry) Then Continue For
            If _state.MergedEntries.RenamedEntryNames.Contains(entry) Then Continue For

            Dim changes = GetKeyChanges(entry)
            If changes.RemovedKeys.Count + changes.AddedKeys.Count + changes.UpdatedKeysDict.Count = 0 Then Continue For

            qualifying.Add(entry)

        Next

        If qualifying.Count = 0 Then Return results

        Dim emitEntries = Sub()

                              For Each entry In qualifying

                                  Dim addKeyTypes, remKeyTypes, modKeyTypes As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
                                  Dim newSectionVer = _file2.GetSection(entry)
                                  Dim changes = GetKeyChanges(entry)
                                  Dim sourceMap As Dictionary(Of String, String) = Nothing
                                  If isMerger Then _mergerSourceMaps.TryGetValue(entry, sourceMap)

                                  results.Add(MakeDiff(newSectionVer, 2))
                                  results.AddRange(ItemizeChangesFromList(changes.AddedKeys, True, addKeyTypes, sourceMap))
                                  results.AddRange(ItemizeChangesFromList(changes.RemovedKeys, False, remKeyTypes, sourceMap))
                                  results.AddRange(ItemizeUpdatedKeys(changes.UpdatedKeysDict, changes.AddedKeys, changes.RemovedKeys, modKeyTypes, sourceMap))
                                  results.Add(ItemizeMergedEntries(entry, isMerger))

                              Next

                          End Sub

        If isMerger Then emitEntries() : Return results

        Dim modHeader = $"{qualifying.Count} modified {If(qualifying.Count = 1, "entry", "entries")}"
        Dim modHeaderSection As New MenuSection
        modHeaderSection.AddColoredLine(modHeader, ConsoleColor.Yellow, centered:=True).AddDivider(solid:=False)
        results.Add(modHeaderSection)

        gLog(Nothing, leadr:=True)

        Using gLogScope("Modified entries:")

            emitEntries()

        End Using

        gLog(modHeader)

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

        For Each kvp In updatedKeysDict

            Dim bucket = If(DetectionKeyTypes.Contains(kvp.Key.KeyType), DetectionCriteriaLabel, kvp.Value(0).KeyType)
            recordModification(modKeyTypes, bucket)

        Next

        Using gLogScope()

            result.Add(summarizeEntryUpdate(modKeyTypes, "Modified"))

            For i = 0 To updatedKeysDict.Count - 1

                Dim output As New MenuSection
                Dim newKey = updatedKeysDict.Keys(i)
                Dim oldKeys = updatedKeysDict.Values(i)
                Dim isRename = newKey.typeIs("Name")
                Dim isDetection = Not isRename AndAlso DetectionKeyTypes.Contains(newKey.KeyType)
                Dim count = updatedKeysDict.Values(i).Count

                Dim outTxt1 As String

                If isDetection Then

                    outTxt1 = $"{DetectionCriteriaLabel} modified{If(count > 1, $", replacing {count} old keys", "")}"

                Else

                    Dim outText1EntryName = If(isRename, "Entry Name", newKey.Name)
                    outTxt1 = $"{outText1EntryName} has been modified{If(Not isRename, $", replacing {count} old key{If(count > 1, "s", "")}", "")}"

                End If

                output.AddColoredLine(outTxt1, ConsoleColor.DarkYellow)
                gLog($"  {outTxt1}", leadr:=i = 0)

                Dim outTxt2 = $" + New: {If(isRename, newKey.Value, newKey.ToString())}"

                output.AddColoredLine(outTxt2, ConsoleColor.Green)
                gLog($"        {outTxt2}")

                For Each oldKey In oldKeys

                    Dim sourceInfo = ""
                    Dim hasSourceInfo = sourceEntryMap IsNot Nothing AndAlso sourceEntryMap.ContainsKey(oldKey.Value)
                    If hasSourceInfo Then sourceInfo = $" (from [{sourceEntryMap(oldKey.Value)}])"

                    Dim old = $" - Old: {If(isRename, oldKey.Value, oldKey.ToString())}{sourceInfo}"

                    output.AddColoredLine(old, ConsoleColor.Red)
                    gLog($"        {old}")

                Next

                result.Add(output)

            Next

        End Using

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

        out.AddBlank()
        out.AddColoredLine(outTxt, ConsoleColor.Yellow, centered:=True)
        gLog()
        gLog($"  {outTxt}")

        For Each mergedEntry In _state.MergedEntries.MergeDict(entry)

            out.AddColoredLine(mergedEntry, ConsoleColor.DarkCyan, centered:=True)
            gLog($"  {mergedEntry}")

        Next
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

        Dim qualifying As New List(Of Tuple(Of String, List(Of iniKey2), List(Of iniKey2), Dictionary(Of iniKey2, List(Of iniKey2))))

        For Each newName In _state.MergedEntries.RenamedEntryNames.OrderBy(Function(s) s, StringComparer.OrdinalIgnoreCase)

            Dim changes = GetKeyChanges(newName)
            Dim addedKeys = changes.AddedKeys
            Dim removedKeys = changes.RemovedKeys
            Dim rawUpdatedDict = changes.UpdatedKeysDict

            Dim updatedKeysDict As New Dictionary(Of iniKey2, List(Of iniKey2))
            For Each kvp In rawUpdatedDict

                If Not kvp.Key.typeIs("Name") Then updatedKeysDict.Add(kvp.Key, kvp.Value)

            Next

            If removedKeys.Count + addedKeys.Count + updatedKeysDict.Count = 0 Then Continue For

            qualifying.Add(Tuple.Create(newName, addedKeys, removedKeys, updatedKeysDict))

        Next

        If qualifying.Count = 0 Then Return results

        Dim msg = $"Minor changes to {qualifying.Count} renamed {If(qualifying.Count = 1, "entry", "entries")}"
        Dim headerSection As New MenuSection
        headerSection.AddColoredLine(msg, ConsoleColor.Magenta, centered:=True).AddDivider(solid:=False)
        results.Add(headerSection)

        gLog(Nothing, leadr:=True)

        Using gLogScope("Minor changes to renamed entries:")

            For Each item In qualifying

                Dim newName = item.Item1
                Dim addedKeys = item.Item2
                Dim removedKeys = item.Item3
                Dim updatedKeysDict = item.Item4
                Dim oldName = _state.MergedEntries.RenamedEntryPairs(newName)
                Dim addKeyTypes, remKeyTypes, modKeyTypes As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)

                results.Add(MakeDiff(_file1.GetSection(oldName), 3, _file2.GetSection(newName)))
                results.AddRange(ItemizeChangesFromList(addedKeys, True, addKeyTypes, Nothing))
                results.AddRange(ItemizeChangesFromList(removedKeys, False, remKeyTypes, Nothing))
                results.AddRange(ItemizeUpdatedKeys(updatedKeysDict, addedKeys, removedKeys, modKeyTypes))

            Next

        End Using

        gLog(msg, leadr:=True)

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

        If changeType = 2 OrElse changeType = 3 Then printColor = If(changeType = 2, ConsoleColor.DarkYellow, ConsoleColor.Magenta)

        Dim renamedOrMergedEntryName = If((changeType = 3 OrElse changeType = 4) AndAlso newSection IsNot Nothing, newSection.Name, "")
        Dim changeTypeStrs = {"added", "removed", "modified", "renamed to ", "merged into "}
        Dim changeStr = $"{section.Name} has been {changeTypeStrs(changeType)}{renamedOrMergedEntryName}"

        result.AddBlank(changeType = 2)
        result.AddColoredLine(changeStr, color:=If(changeType >= 2, printColor, GetRedGreen(changeType = 1)), centered:=True)
        gLog($"  {changeStr}", leadr:=True)

        If Not ShowFullEntries Then Return result

        ' Everything below this point prints to the user only if Verbose Mode is enabled
        Dim isMergeOrRenamed = changeType >= 3 AndAlso changeType < 5

        If isMergeOrRenamed Then

            result.AddBlank()
            result.AddColoredLine("Old entry:", color:=ConsoleColor.DarkRed, centered:=True)

            gLog()
            gLog("    Old entry:", leadr:=True)

        End If

        BuildEntrySection(result, section.ToString)

        If Not isMergeOrRenamed Then Return result

        Dim out = If(changeType = 3, "Renamed entry: ", "Merged entry: ")

        result.AddBlank()
        result.AddColoredLine(out, color:=printColor, centered:=True)

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
            gLog($"        {line}")

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

        Dim sourceCount = movementsBySource.Count
        Dim movHeader = $"keys moved between entries ({sourceCount} source {If(sourceCount = 1, "entry", "entries")})"
        Dim movHeaderSection As New MenuSection
        movHeaderSection.AddColoredLine(movHeader, ConsoleColor.Cyan, centered:=True).AddDivider(solid:=False)
        results.Add(movHeaderSection)
        Dim totKeys = 0
        gLog(Nothing, leadr:=True)

        Using gLogScope("Cross-Entry key movements: ")

            For Each sourceEntry In movementsBySource.OrderBy(Function(kvp) kvp.Key, StringComparer.OrdinalIgnoreCase)

                Dim out As New MenuSection
                out.AddColoredLine($"{sourceEntry.Key}", ConsoleColor.Cyan, centered:=True)
                gLog($"{sourceEntry.Key}", leadr:=True)

                For Each movement In sourceEntry.Value

                    Dim line = $"  -> {movement.KeyName}={movement.KeyValue} moved to [{movement.Target}]"
                    out.AddColoredLine(line, ConsoleColor.DarkCyan)
                    gLog($"  {line}")
                    totKeys += 1

                Next

                out.AddBlank()
                results.Add(out)

            Next

        End Using

        gLog($"{totKeys} {movHeader}", leadr:=True)

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

        Dim qualifying = _state.ModifiedEntries.AddedEntryNames _
            .OrderBy(Function(s) s, StringComparer.OrdinalIgnoreCase) _
            .Where(Function(entry) Not _state.MergedEntries.RenamedEntryNames.Contains(entry) AndAlso
                                    _state.MergedEntries.MergeDict.ContainsKey(entry)) _
            .ToList()

        If qualifying.Count = 0 Then Return results

        Dim awmHeader = $"{qualifying.Count} added {If(qualifying.Count = 1, "entry", "entries")} consolidating removed content"
        Dim awmHeaderSection As New MenuSection
        awmHeaderSection.AddColoredLine(awmHeader, ConsoleColor.Green, centered:=True).AddDivider(solid:=False)
        results.Add(awmHeaderSection)

        gLog(Nothing, leadr:=True)

        Using gLogScope("Added entries containing merged content:")

            For Each entry In qualifying

                Dim section = _file2.GetSection(entry)
                Dim mergedCount = _state.MergedEntries.MergeDict(entry).Count

                Dim combined = BuildCombinedOldKeys(_state.MergedEntries.MergeDict(entry), "")
                Dim sourceEntryMap = combined.SourceEntryMap

                _keyAnalyzer.FindModificationsForAddedEntryFromKeys(combined.Keys, _file2.GetSection(entry))

                Dim headerSection As New MenuSection
                Dim headerText = $"{entry} has been added (consolidating {mergedCount} removed entr{If(mergedCount = 1, "y", "ies")})"

                headerSection.AddColoredLine(headerText, ConsoleColor.Green, centered:=True)
                gLog(headerText, leadr:=True)
                results.Add(headerSection)

                Dim mergedListSection As New MenuSection
                mergedListSection.AddColoredLine("Merged from:", ConsoleColor.Yellow, centered:=True)
                gLog("Merged from:")

                For Each mergedEntry In _state.MergedEntries.MergeDict(entry)

                    mergedListSection.AddColoredLine($"{mergedEntry}", ConsoleColor.DarkCyan, centered:=True)
                    gLog($"  • {mergedEntry}")

                Next

                mergedListSection.AddBlank()
                results.Add(mergedListSection)

                Dim entryChanges = GetKeyChanges(entry)
                Dim addedKeys = entryChanges.AddedKeys
                Dim removedKeys = entryChanges.RemovedKeys
                Dim updatedKeysDict = entryChanges.UpdatedKeysDict

                Dim addedKeyIds As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                For Each k In addedKeys : addedKeyIds.Add($"{k.KeyType}|{k.Value}") : Next


                Dim removedKeyIds As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                For Each k In removedKeys : removedKeyIds.Add($"{k.KeyType}|{k.Value}") : Next

                Dim carriedOverKeys As New List(Of iniKey2)
                For Each key In _file2.GetSection(entry).Keys

                    Dim keyId = $"{key.KeyType}|{key.Value}"
                    If addedKeyIds.Contains(keyId) OrElse removedKeyIds.Contains(keyId) Then Continue For

                    Dim isUpdated = False
                    For Each kvp In updatedKeysDict

                        If kvp.Key.Value = key.Value Then isUpdated = True : Exit For

                    Next

                    If isUpdated Then Continue For

                    If sourceEntryMap.ContainsKey(key.Value) Then carriedOverKeys.Add(key)

                Next

                addedKeys.AddRange(carriedOverKeys)

                If addedKeys.Count + removedKeys.Count + updatedKeysDict.Count > 0 Then

                    Dim addKeyTypes, remKeyTypes, modKeyTypes As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)

                    If addedKeys.Count > 0 Then

                        Dim newKeysSection As New MenuSection
                        Dim novelKeysMsg = $"{addedKeys.Count} keys added or carried over from merged sources:"
                        newKeysSection.AddColoredLine(novelKeysMsg, ConsoleColor.Green, centered:=True)
                        gLog()
                        gLog(novelKeysMsg)
                        results.Add(newKeysSection)
                        results.AddRange(ItemizeChangesFromList(addedKeys, True, addKeyTypes, sourceEntryMap))

                    End If

                    If removedKeys.Count > 0 Then

                        Dim droppedSection As New MenuSection

                        Dim KeysNotMergedMsg = $"{removedKeys.Count} keys from merged entries not in this entry:"
                        droppedSection.AddColoredLine(KeysNotMergedMsg, ConsoleColor.DarkYellow, centered:=True)
                        gLog()
                        gLog(KeysNotMergedMsg)
                        results.Add(droppedSection)
                        results.AddRange(ItemizeChangesFromList(removedKeys, False, remKeyTypes, sourceEntryMap))

                    End If

                    If updatedKeysDict.Count > 0 Then

                        Dim capturedSection As New MenuSection
                        Dim capturedMsg = $"{updatedKeysDict.Count} keys capturing content from merged entries"
                        capturedSection.AddColoredLine(capturedMsg, ConsoleColor.Yellow, centered:=True)
                        gLog()
                        gLog(capturedMsg)
                        results.Add(capturedSection)
                        results.AddRange(ItemizeUpdatedKeys(updatedKeysDict, addedKeys, removedKeys, modKeyTypes, sourceEntryMap))

                    End If

                End If

                results.Add(New MenuSection().AddBlank)

            Next

        End Using

        gLog(awmHeader, leadr:=True)

        Return results

    End Function

    ''' <summary>
    ''' Outputs a summary section listing Section values that represent newly added browser
    ''' support i.e. <c> Section </c> key values containing <c> "Web Browser" </c> that
    ''' appear in the new file but not the old file.
    ''' Returns an empty list when no new browser support was detected.
    ''' </summary>
    '''
    ''' <returns>
    ''' Zero or one <c> MenuSection </c> summarising new browser support
    ''' </returns>
    Public Function ItemizeNewBrowsers() As List(Of MenuSection)

        Return ItemizeBrowserChanges(_state.Statistics.NewBrowserSectionValues, True)

    End Function

    ''' <summary>
    ''' Outputs a summary section listing Section values for Web Browsers which are either 
    ''' newly added or newly removed
    ''' </summary>
    ''' 
    ''' <param name="ObservedValues">
    ''' Section key values relating to web browsers
    ''' </param>
    ''' 
    ''' <param name="wasAdded">
    ''' Indicates that the current assessment is for newly added browsers
    ''' </param>
    ''' <returns></returns>
    Public Shared Function ItemizeBrowserChanges(ByRef ObservedValues As List(Of String),
                                                wasAdded As Boolean) As List(Of MenuSection)

        Dim results As New List(Of MenuSection)

        If ObservedValues.Count = 0 Then Return results

        Dim plural = $"web browser{If(ObservedValues.Count = 1, "", "s")}"
        Dim header = $"{ObservedValues.Count} {plural} {If(wasAdded, "added", "removed")}"

        Dim out As New MenuSection
        Dim pColor = GetRedGreen(Not wasAdded)
        out.AddColoredLine(header, pColor, True).AddDivider(solid:=False)

        gLog(Nothing, leadr:=True)

        Using gLogScope($"Web Browser {If(wasAdded, "additions", "removals")}")

            gLog("")

            For Each value In ObservedValues

                out.AddColoredLine(value, pColor, True)
                gLog(value, buffr:=True)

            Next


        End Using

        gLog(header)

        results.Add(out)
        Return results

    End Function

    ''' <summary>
    ''' Produces one <c>MenuSection</c> listing all web browsers whose support was
    ''' removed in the new file, or an empty list if none were removed.
    ''' </summary>
    '''
    ''' <returns>
    ''' A list containing a single <c>MenuSection</c> with the removed browser header and
    ''' one line per removed browser value, or an empty list if there are no removed browsers
    ''' </returns>
    Public Function ItemizeRemovedBrowsers() As List(Of MenuSection)

        Return ItemizeBrowserChanges(_state.Statistics.RemovedBrowserSectionValues, False)

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

            Dim suffix = If(kvp.Value > 1 AndAlso kvp.Key <> DetectionCriteriaLabel, "s", "")
            Dim out = $"{changeType} {kvp.Value} {kvp.Key}{suffix}"
            result.AddColoredLine(out, ConsoleColor.Yellow, centered:=True)
            result.AddBlank(i = total - 1)

            gLog($"  {out}", leadr:=i = 0)
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

        Dim changeTxt = If(wasAdded, "Added", "Removed")

        kl.ForEach(Sub(key) recordModification(ktDict, key.KeyType))

        Using gLogScope()

            out.Add(summarizeEntryUpdate(ktDict, changeTxt))

            Dim result As New MenuSection

            For i = 0 To kl.Count - 1

                Dim key = kl(i).ToString()
                Dim color = If(wasAdded, ConsoleColor.Green, ConsoleColor.Red)

                Dim sourceInfo = ""
                If sourceEntryMap IsNot Nothing AndAlso sourceEntryMap.ContainsKey(kl(i).Value) Then

                    sourceInfo = $" (from [{sourceEntryMap(kl(i).Value)}])"

                ElseIf wasAdded AndAlso sourceEntryMap IsNot Nothing Then

                    sourceInfo = " (novel)"

                End If

                result.AddColoredLine(key & sourceInfo, color)
                gLog($"        {key & sourceInfo}")

            Next

            result.AddBlank()
            out.Add(result)

        End Using

        Return out

    End Function

    Private Sub Emit(section As MenuSection, text As String, color As ConsoleColor, Optional condition As Boolean = True)
        section.AddColoredLine(text, color, condition:=condition)
        gLog(text, cond:=condition)
    End Sub

    Private Class EntryKeyChanges

        Public ReadOnly AddedKeys As List(Of iniKey2)
        Public ReadOnly RemovedKeys As List(Of iniKey2)
        Public ReadOnly UpdatedKeysDict As Dictionary(Of iniKey2, List(Of iniKey2))

        Public Sub New(added As List(Of iniKey2),
                       removed As List(Of iniKey2),
                       updated As Dictionary(Of iniKey2, List(Of iniKey2)))

            AddedKeys = added
            RemovedKeys = removed
            UpdatedKeysDict = updated

        End Sub

    End Class

    Private Function GetKeyChanges(entry As String) As EntryKeyChanges

        Dim added = If(_state.ModifiedEntries.AddedKeyTracker2.ContainsKey(entry),
                       _state.ModifiedEntries.AddedKeyTracker2(entry), New List(Of iniKey2))

        Dim removed = If(_state.ModifiedEntries.RemovedKeyTracker2.ContainsKey(entry),
                         _state.ModifiedEntries.RemovedKeyTracker2(entry), New List(Of iniKey2))

        Dim updated = If(_state.ModifiedEntries.ModifiedKeyTracker2.ContainsKey(entry),
                         _state.ModifiedEntries.ModifiedKeyTracker2(entry),
                         New Dictionary(Of iniKey2, List(Of iniKey2)))

        Return New EntryKeyChanges(added, removed, updated)

    End Function

    Private Class CombinedOldKeyResult

        Public ReadOnly Keys As List(Of iniKey2)
        Public ReadOnly SourceEntryMap As Dictionary(Of String, String)

        Public Sub New(keys As List(Of iniKey2), sourceMap As Dictionary(Of String, String))

            Me.Keys = keys
            SourceEntryMap = sourceMap

        End Sub

    End Class

    ''' <summary>
    ''' Builds a deduplicated list of <c>iniKey2</c> objects from all old entries
    ''' named in <paramref name="mergeSourceNames"/>, plus optionally from
    ''' <paramref name="targetEntry"/> itself if it existed in file1 and is not
    ''' already a named merge source.
    ''' </summary>
    '''
    ''' <param name="mergeSourceNames">
    ''' Names of the old entries whose keys are to be combined
    ''' </param>
    '''
    ''' <param name="targetEntry">
    ''' The entry receiving the merged content; when non-empty,
    ''' its own file1 keys are appended if it existed in file1
    ''' and is not listed as a source. Pass an empty string to skip this step.
    ''' </param>
    '''
    ''' <returns>
    ''' A <c>CombinedOldKeyResult</c> with the deduplicated key list
    ''' and the value → source-entry attribution map
    ''' </returns>
    Private Function BuildCombinedOldKeys(mergeSourceNames As IEnumerable(Of String),
                                          targetEntry As String) As CombinedOldKeyResult

        Dim uniqueKeyValues As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Dim combinedKeys As New List(Of iniKey2)
        Dim sourceMap As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

        For Each oldEntryName In mergeSourceNames

            Dim oldEnt = _file1.GetSection(oldEntryName)
            If oldEnt Is Nothing Then Continue For

            For Each key In oldEnt.Keys

                If uniqueKeyValues.Contains(key.Value) Then Continue For

                combinedKeys.Add(key)
                uniqueKeyValues.Add(key.Value)
                sourceMap(key.Value) = oldEntryName

            Next

        Next

        If targetEntry <> "" AndAlso
           _file1.Contains(targetEntry) AndAlso
           Not mergeSourceNames.Contains(targetEntry) Then

            For Each key In _file1.GetSection(targetEntry).Keys

                If uniqueKeyValues.Contains(key.Value) Then Continue For

                combinedKeys.Add(key)
                uniqueKeyValues.Add(key.Value)
                sourceMap(key.Value) = targetEntry

            Next

        End If

        Return New CombinedOldKeyResult(combinedKeys, sourceMap)

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
