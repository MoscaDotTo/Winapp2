'    Copyright (C) 2018-2022 Hazel Ward
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
Imports System.Globalization
''' <summary> Performs a "Diff" on two winapp2.ini files </summary>
''' Docs last updated: 2020-09-01 
Module Diff

    ''' <summary> The old or local version of winapp2.ini to be diffed </summary>
    ''' Docs last updated: 2020-09-02 | Code last updated: 2020-09-02
    Public Property DiffFile1 As iniFile = New iniFile(Environment.CurrentDirectory, "winapp2.ini", mExist:=True)

    ''' <summary> The new or remote version of winapp2.ini against which DiffFile1 will be compared </summary>
    ''' Docs last updated: 2020-09-02 | Code last updated: 2020-09-02
    Public Property DiffFile2 As iniFile = New iniFile(Environment.CurrentDirectory, "", mExist:=True)

    ''' <summary> The path to which the log will optionally be saved </summary>
    ''' Docs last updated: 2020-09-02 | Code last updated: 2020-09-02
    Public Property DiffFile3 As iniFile = New iniFile(Environment.CurrentDirectory, "diff.txt")

    ''' <summary> Indicates that a remote winapp2.ini should be downloaded to use as <c> DiffFile2 </c> </summary>
    ''' Docs last updated: 2020-09-02 | Code last updated: 2020-09-02
    Public Property DownloadDiffFile As Boolean = Not isOffline

    ''' <summary> Indicates that the diff output should be saved to disk </summary>
    ''' Docs last updated: 2020-09-02 | Code last updated: 2020-09-02
    Public Property SaveDiffLog As Boolean = False

    ''' <summary> Indicates that the module settings have been modified from their defaults </summary>
    ''' Docs last updated: 2020-09-02 | Code last updated: 2020-09-02
    Public Property DiffModuleSettingsChanged As Boolean = False

    ''' <summary> Indicates that the remote file should be trimmed for the local system before diffing </summary>
    ''' Docs last updated: 2020-09-02 | Code last updated: 2020-09-02
    Public Property TrimRemoteFile As Boolean = Not isOffline

    ''' <summary> The number of entries Diff determines to have been added between versions </summary>
    ''' Docs last updated: 2020-09-02 | Code last updated: 2020-09-02
    Public Property AddedEntryCount As Integer = 0

    ''' <summary> The number of entries Diff determines to have been modified between versions </summary>
    ''' Docs last updated: 2020-09-02 | Code last updated: 2020-09-02
    Public Property ModifiedEntryCount As Integer = 0

    ''' <summary> The number of entries Diff determines to have been removed between versions </summary>
    ''' Docs last updated: 2020-09-02 | Code last updated: 2020-09-02
    Public Property RemovedEntryCount As Integer = 0

    ''' <summary> The number of entries Diff determines to have been merged between versions  </summary>
    ''' Docs last upated 2022-06-20 | Code last updated 2022-06-20 
    Public Property MergedEntryCount As Integer = 0

    ''' <summary> The number of entries Diff determines to have been renamed between versions  </summary>
    ''' Docs last upated 2022-06-20 | Code last updated 2022-06-20 
    Public Property RenamedEntryCount As Integer = 0


    ''' <summary> The total number of keys that were added to modified entries </summary>
    ''' Docs last updated: 2020-09-02 | Code last updated: 2020-09-02
    Private Property ModEntriesAddedKeyTotal As Integer = 0

    ''' <summary> The total number of keys that were removed from modified entries </summary>
    ''' Docs last updated: 2020-09-02 | Code last updated: 2020-09-02
    Private Property ModEntriesRemovedKeyTotal As Integer = 0

    ''' <summary> The total number of keys that were updated in modified entries </summary>
    ''' Docs last updated: 2020-09-02 | Code last updated: 2020-09-02
    Private Property ModEntrtiesUpdatedKeyTotal As Integer = 0

    ''' <summary> Indicates that full entries should be printed in the Diff output. <br/> <br/> Called "verbose mode" in the menu </summary>
    ''' Docs last updated: 2020-09-02 | Code last updated: 2020-09-02
    Public Property ShowFullEntries As Boolean = False

    ''' <summary> Holds the log from the most recent run of the Differ to display back to the user </summary>
    ''' Docs last updated: 2020-09-02 | Code last updated: 2020-09-02
    Public MostRecentDiffLog As String = ""

    ''' <summary> Handles the commandline args for Diff </summary>
    ''' Docs last updated: 2020-09-02 | Code last updated: 2020-09-02
    '''  Diff args:
    ''' -d          : download the latest winapp2.ini
    ''' -ncc        : download the latest non-ccleaner winapp2.ini (implies -d)
    ''' -savelog    : save the diff.txt log
    Public Sub handleCmdLine()
        initDefaultDiffSettings()
        handleDownloadBools(DownloadDiffFile)
        ' Make sure we have a name set for the new file if we're downloading or else the diff will not run
        If DownloadDiffFile Then DiffFile2.Name = If(RemoteWinappIsNonCC, "Online non-ccleaner winapp2.ini", "Online winapp2.ini")
        invertSettingAndRemoveArg(SaveDiffLog, "-savelog")
        getFileAndDirParams(DiffFile1, DiffFile2, DiffFile3)
        If DiffFile2.Name.Length <> 0 Then initDiff()
    End Sub

    ''' <summary> Runs the Differ from outside the module </summary>
    ''' <param name="firstFile"> The local winapp2.ini file to diff against the master GitHub copy </param>
    ''' Docs last updated: 2020-09-02 | Code last updated: 2020-09-02
    Public Sub remoteDiff(firstFile As iniFile)
        DiffFile1 = firstFile
        DownloadDiffFile = True
        initDiff()
    End Sub

    ''' <summary> Carries out the main set of Diffing operations </summary>
    ''' Docs last updated: 2020-09-02 | Code last updated: 2020-09-02
    Public Sub initDiff()
        If Not enforceFileHasContent(DiffFile1) Then Return
        If DownloadDiffFile Then
            Dim downloadedIniFile = getRemoteIniFile(winapp2link)
            DiffFile2.Sections = downloadedIniFile.Sections
            ' We don't actually do anything with the comments but in case we ever do, we may as well retain them 
            DiffFile2.Comments = downloadedIniFile.Comments
        Else
            If Not enforceFileHasContent(DiffFile2) Then Return
        End If
        If TrimRemoteFile AndAlso DownloadDiffFile Then
            Dim tmp As New winapp2file(DiffFile2)
            Trim.trimFile(tmp)
            DiffFile2.Sections = tmp.toIni.Sections
        End If
        logInitDiff()
        compareTo()
        logPostDiff()
        print(3, pressEnterStr)
        crl()
        MostRecentDiffLog = If(SaveDiffLog, getLogSliceFromGlobal("Beginning diff", "Diff complete"), "hasBeenRun")
        DiffFile3.overwriteToFile(MostRecentDiffLog, SaveDiffLog)
        setHeaderText(If(SaveDiffLog, DiffFile3.Name & " saved", "Diff complete"))
    End Sub

    '''<summary> Logs the initial portion of the diff output for the user </summary>
    ''' Docs last updated: 2020-09-02 | Code last updated: 2020-09-02
    Private Sub logInitDiff()
        print(3, "Diffing, please wait. This may take a moment.")
        clrConsole()
        Dim oldVersionNum = getVer(DiffFile1)
        Dim newVersionNum = getVer(DiffFile2)
        gLog($"Beginning diff between{oldVersionNum} and{newVersionNum}", ascend:=True)
        print(6, $"Changes between{oldVersionNum} and{newVersionNum}", enStrCond:=False)
    End Sub

    ''' <summary> Gets the version from winapp2.ini </summary>
    ''' <param name="someFile"> A winapp2.ini format <c> iniFile </c> </param>
    ''' Docs last updated: 2020-09-02 | Code last updated: 2020-09-02
    Private Function getVer(someFile As iniFile) As String
        Dim ver = If(someFile.Comments.Count > 0, someFile.Comments(0).Comment.ToString(CultureInfo.InvariantCulture).ToUpperInvariant, "000000")
        Return If(ver.Contains("VERSION"), ver.TrimStart(CChar(";")).Replace("VERSION:", "version"), " version not given")
    End Function

    ''' <summary> Logs and prints the summary of the Diff </summary>
    ''' Docs last updated: 2020-09-02 | Code last updated: 2022-06-20
    Private Sub logPostDiff()
        gLog($"Added entries: {AddedEntryCount}", indent:=True)
        gLog($"Modified entries: {ModifiedEntryCount}", indent:=True)
        gLog($"Removed entries: {RemovedEntryCount}", indent:=True)
        gLog("Diff complete", descend:=True)
        print(0, Nothing, closeMenu:=True)
        print(4, "Summary", conjoin:=True)
        print(0, $"Added entries: {AddedEntryCount}", colorLine:=True, enStrCond:=True)
        print(0, $"Removed entries: {RemovedEntryCount}", colorLine:=True, enStrCond:=False)
        print(0, $"Modified entries: {ModifiedEntryCount}", colorLine:=True, useArbitraryColor:=True, arbitraryColor:=ConsoleColor.Yellow)
        print(0, $" * {ModEntriesAddedKeyTotal} added keys (total between all entries)", colorLine:=True, enStrCond:=True, cond:=ModEntriesAddedKeyTotal > 0)
        print(0, $" * {ModEntriesRemovedKeyTotal} removed keys (total between all entries)", colorLine:=True, enStrCond:=False, cond:=ModEntriesRemovedKeyTotal > 0)
        print(0, $" * {ModEntrtiesUpdatedKeyTotal} updated keys (total between all entries)", colorLine:=True, useArbitraryColor:=True, arbitraryColor:=ConsoleColor.Yellow, cond:=ModEntrtiesUpdatedKeyTotal > 0)
        print(0, $"Renamed entries: {RenamedEntryCount}", colorLine:=True, useArbitraryColor:=True, arbitraryColor:=ConsoleColor.Magenta)
        print(0, $"Merged entries: {MergedEntryCount}", colorLine:=True, useArbitraryColor:=True, arbitraryColor:=ConsoleColor.Cyan, closeMenu:=True)
    End Sub

    ''' <summary> Compares two winapp2.ini format <c> iniFiles </c> and quantifies the differences to the user </summary>
    ''' Docs last updated: 2020-09-02 | Code last updated: 2022-06-28
    Private Sub compareTo()
        AddedEntryCount = 0
        RemovedEntryCount = 0
        ModifiedEntryCount = 0
        ModEntriesAddedKeyTotal = 0
        ModEntriesRemovedKeyTotal = 0
        ModEntrtiesUpdatedKeyTotal = 0
        RenamedEntryCount = 0
        MergedEntryCount = 0
        Dim renamedEntryNameList As New strList
        Dim mergedEntryNameList As New strList
        Dim modifiedEntryNameList As New strList
        Dim addedEntryNameList As New strList
        Dim removedEntryNameList As New strList
        Dim mergeDict As New Dictionary(Of String, List(Of String))

        ' Determine the names of the entries who appear only in the "new" file 
        For Each section In DiffFile2.Sections.Values
            addedEntryNameList.add(section.Name, Not DiffFile1.Sections.ContainsKey(section.Name))
        Next

        ' Determine the names of the entries who appear only in the "old" file 
        For Each section In DiffFile1.Sections.Values
            removedEntryNameList.add(section.Name, Not DiffFile2.Sections.ContainsKey(section.Name))
        Next

        ' Determine which entries who appear in both files have been modified 
        For Each section In DiffFile1.Sections.Values
            Dim modCounts, addCounts, remCounts As New List(Of Integer)
            Dim modKeyTypes, addKeyTypes, remKeyTypes As New List(Of String)

            ' If any entry exists in both DiffFile1 and DiffFile2, and does not pass the compareTo, then it has been modified 
            If DiffFile2.Sections.Keys.Contains(section.Name) Then
                Dim sSection = DiffFile2.Sections(section.Name)
                Dim addedKeys, removedKeys As New keyList
                Dim updatedKeys As New List(Of KeyValuePair(Of iniKey, iniKey))

                If Not section.compareTo(sSection, removedKeys, addedKeys) Then
                    modifiedEntryNameList.add(sSection.Name)
                    ' Determine the types of modifications made to the entry 
                    chkLsts(removedKeys, addedKeys, updatedKeys)
                    ' Silently ignore any entries with only alphabetization changes
                    If removedKeys.KeyCount + addedKeys.KeyCount + updatedKeys.Count = 0 Then Continue For
                    getDiff(sSection, 2, ModifiedEntryCount)
                    getChangesFromList(addedKeys, True, addKeyTypes, addCounts)
                    getChangesFromList(removedKeys, False, remKeyTypes, remCounts)
                    If updatedKeys.Count > 0 Then
                        gLog("Modifed Keys:", ascend:=True, ascAmt:=2)
                        updatedKeys.ForEach(Sub(pair) recordModification(modKeyTypes, modCounts, pair.Value.KeyType))
                        summarizeEntryUpdate(modKeyTypes, modCounts, "Modified")
                        For i = 0 To updatedKeys.Count - 1
                            Dim pair = updatedKeys(i)
                            recordModification(modKeyTypes, modCounts, pair.Value.KeyType)
                            gLog(pair.Key.Name, indent:=True, indAmt:=2)
                            gLog("Old: " & pair.Value.toString, indent:=True, indAmt:=3)
                            gLog("New: " & pair.Key.toString, indent:=True, indAmt:=3)
                            print(0, "Old: " & pair.Value.toString, colorLine:=True)
                            print(0, "New: " & pair.Key.toString, colorLine:=True, enStrCond:=True)
                            print(0, Nothing, cond:=i = updatedKeys.Count - 1, conjoin:=i = updatedKeys.Count - 1, fillBorder:=False)
                        Next
                        gLog(descend:=True, descAmt:=2)
                    End If
                    addCounts.ForEach(Sub(count) ModEntriesAddedKeyTotal += count)
                    remCounts.ForEach(Sub(count) ModEntriesRemovedKeyTotal += count)
                    modCounts.ForEach(Sub(count) ModEntrtiesUpdatedKeyTotal += count)
                End If
            End If
        Next

        ' Entries that appear in the old file but not in the new file had one of three things occur: 
        ' They had their name changed but none of the keys changed, we'll consider this a "rename" 
        ' They were merged into another key, or renamed directly and then that renamed key was modified some additional way, we'll consider this a "merge" 
        ' They were actually removed from the file without their contents being merged into another entry or with their contents merged into another entry
        ' in a way in which we cannot track (ie. through use of wildcards), in which case we will just consider the entry removed 
        For Each entry In removedEntryNameList.Items
            Dim regKeyCountsMatch = False
            Dim fileKeyCountsMatch = False
            Dim newMergedOrRenamedName = ""
            Dim highestMatchCount = 0
            Dim entryWasRenamedOrMerged = False
            Dim oldSectionVersion = DiffFile1.getSection(entry)

            For Each sSection In DiffFile2.Sections.Values
                ' Don't consider the contents of entries that are neither new nor modified in the new version 
                If Not addedEntryNameList.contains(sSection.Name) And Not modifiedEntryNameList.contains(sSection.Name) Then Continue For
                Dim allFileKeysMatched = False
                Dim someFileKeysMatched = False
                Dim allRegKeysMatched = False
                Dim someRegKeysMatched = False
                fileKeyCountsMatch = False
                regKeyCountsMatch = False
                Dim wa2sSection As New winapp2entry(sSection)
                Dim curwa2Section As New winapp2entry(oldSectionVersion)
                Dim totalMatches = 0
                assessKeyMatches(curwa2Section.FileKeys, wa2sSection.FileKeys, fileKeyCountsMatch, someFileKeysMatched, allFileKeysMatched, totalMatches)
                assessKeyMatches(curwa2Section.RegKeys, wa2sSection.RegKeys, regKeyCountsMatch, someRegKeysMatched, allRegKeysMatched, totalMatches)
                ' If we hit matches along the way, we'll remember the name of the entry with the largest number of matching keys to assess at the end 
                ' but only if it exists in the set of entries who have been added or modified 
                If (curwa2Section.FileKeys.KeyCount > 0 And someFileKeysMatched) Or (someRegKeysMatched And curwa2Section.RegKeys.KeyCount > 0) And
                     totalMatches > highestMatchCount And strList.IsInAny({addedEntryNameList, modifiedEntryNameList}, sSection.Name) Then
                    newMergedOrRenamedName = sSection.Name
                End If
                ' If all the filekeys and regkeys and their respective counts match between two entries then it 
                ' stands to reason that the old version of the key was renamed into the new version 
                If allFileKeysMatched And allRegKeysMatched And fileKeyCountsMatch And regKeyCountsMatch Then
                    getDiff(oldSectionVersion, 3, RenamedEntryCount, newMergedOrRenamedName, sSection)
                    entryWasRenamedOrMerged = True
                    renamedEntryNameList.add(newMergedOrRenamedName)
                    Continue For
                ElseIf allFileKeysMatched And allRegKeysMatched Then
                    ' Likewise, if all the keys matched but the counts don't match, then the entry was probably merged 
                    getDiff(oldSectionVersion, 4, MergedEntryCount, newMergedOrRenamedName, sSection)
                    entryWasRenamedOrMerged = True
                    mergedEntryNameList.add(newMergedOrRenamedName)
                    If Not mergeDict.ContainsKey(newMergedOrRenamedName) Then mergeDict.Add(newMergedOrRenamedName, New List(Of String))
                    mergeDict(newMergedOrRenamedName).Add(entry)
                    Continue For
                End If
            Next
            ' At this stage we'll consider ourselves unable to definitively state that an entry has been renamed and classify the entry 
            ' as having been merged if it 
            If Not entryWasRenamedOrMerged Then
                If strList.IsInAny({modifiedEntryNameList, renamedEntryNameList, addedEntryNameList, mergedEntryNameList}, newMergedOrRenamedName) Then
                    mergedEntryNameList.add(newMergedOrRenamedName)
                    getDiff(oldSectionVersion, 4, MergedEntryCount, newMergedOrRenamedName, DiffFile2.Sections(newMergedOrRenamedName))
                    If Not mergeDict.ContainsKey(newMergedOrRenamedName) Then mergeDict.Add(newMergedOrRenamedName, New List(Of String))
                    mergeDict(newMergedOrRenamedName).Add(entry)
                    Continue For
                End If
                ' If at this stage we haven't determined the entry to exist as part of another entry, it was probably removed for realsies 
                getDiff(oldSectionVersion, 1, RemovedEntryCount)
            End If
        Next
        ' Any sections from the new file which are not found in the old file have been added 
        For Each section In DiffFile2.Sections.Values
            If Not DiffFile1.Sections.Keys.Contains(section.Name) And Not renamedEntryNameList.contains(section.Name) Then
                getDiff(section, 0, AddedEntryCount)
                If mergeDict.ContainsKey(section.Name) Then
                    print(0, "This entry contains keys merged from the following removed entries:", isCentered:=True, colorLine:=True, useArbitraryColor:=True, arbitraryColor:=ConsoleColor.Yellow)
                    For Each mergedEntry In mergeDict(section.Name)
                        print(0, mergedEntry, isCentered:=True, colorLine:=True, useArbitraryColor:=True, arbitraryColor:=ConsoleColor.DarkCyan)
                    Next
                    print(0, "", conjoin:=True, fillBorder:=False)
                End If
            End If
        Next
    End Sub

    ''' <summary> Counts the number of matching key contents between two ini keyLists </summary>
    ''' <param name="currentKeyList"> The list of keys from the "old" version of the entry </param>
    ''' <param name="newKeyList"> The list of keys from the "new" version of the entry </param>
    ''' <param name="countTracker"> Tracks the number of matches observed between the two given keyLists </param>
    ''' <param name="someKeysMatchedTracker"> Tracks whether or not any keys have matched </param>
    ''' <param name="allKeysMatchedTracker"> Tracks whether or not all the keys have matched </param>
    ''' <param name="MatchCount"> The number of matches observed </param>
    Public Sub assessKeyMatches(currentKeyList As keyList, newKeyList As keyList, ByRef countTracker As Boolean, ByRef someKeysMatchedTracker As Boolean, ByRef allKeysMatchedTracker As Boolean, ByRef MatchCount As Integer)
        If currentKeyList.KeyCount = newKeyList.KeyCount Then countTracker = True
        ' If there's nothing to match, consider the keys matched 
        If currentKeyList.KeyCount = 0 Then
            someKeysMatchedTracker = True
            allKeysMatchedTracker = True
        Else
            ' Otherwise, determine whether or not some or all the FileKeys have matched 
            For Each key In currentKeyList.Keys
                For Each newFileKey In newKeyList.Keys
                    If String.Equals(newFileKey.Value, key.Value, StringComparison.InvariantCultureIgnoreCase) Then
                        someKeysMatchedTracker = True
                        MatchCount += 1
                        Exit For
                    End If
                Next
            Next
            If MatchCount = currentKeyList.KeyCount Then allKeysMatchedTracker = True
        End If
    End Sub

    ''' <summary> Records the number of changes made in a modified entry </summary>
    ''' <param name="ktList">The KeyTypes for the type of change being observed </param>
    ''' <param name="countsList">The counts of the changed KeyTypes </param>
    ''' <param name="keyType"> A KeyType from a key that has been changed and whose change will be recorded </param>
    ''' Docs last updated: 2020-09-02 | Code last updated: 2020-09-02
    Private Sub recordModification(ByRef ktList As List(Of String), ByRef countsList As List(Of Integer), ByRef keyType As String)
        If Not ktList.Contains(keyType) Then
            ktList.Add(keyType)
            countsList.Add(1)
        Else
            countsList(ktList.IndexOf(keyType)) += 1
        End If
    End Sub

    ''' <summary> Outputs the entry update summary </summary>
    ''' <param name="keyTypeList"> The KeyTypes that have been updated </param>
    ''' <param name="countList"> The quantity of keys by KeyType who have been modified </param>
    ''' <param name="changeType">The type of change being summarized </param>
    ''' Docs last updated: 2020-09-02 | Code last updated: 2020-09-02
    Private Sub summarizeEntryUpdate(keyTypeList As List(Of String), countList As List(Of Integer), changeType As String)
        For i = 0 To keyTypeList.Count - 1
            gLog($"{changeType} {countList(i)} {keyTypeList(i)}s")
            print(0, $"{changeType} {countList(i)} {keyTypeList(i)}s", colorLine:=True, useArbitraryColor:=True, arbitraryColor:=ConsoleColor.Yellow, isCentered:=True, trailingBlank:=i = keyTypeList.Count - 1)
        Next
    End Sub

    ''' <summary> Prints any added or removed keys from an updated entry to the user </summary>
    ''' <param name="kl"> The iniKeys that have been added/removed from an entry </param>
    ''' <param name="wasAdded"> <c> True </c> if keys in <c> <paramref name="kl"/> </c> were added, <c> False </c> otherwise </param>
    ''' <param name="ktList"> The KeyTypes of modified keys </param>
    ''' <param name="countList"> The counts of the KeyTypes for modified keys </param>
    ''' Docs last updated: 2020-09-04 | Code last updated: 2020-09-04
    Private Sub getChangesFromList(kl As keyList, wasAdded As Boolean, ByRef ktList As List(Of String), ByRef countList As List(Of Integer))
        If kl.KeyCount = 0 Then Return
        Dim changeTxt = If(wasAdded, "Added", "Removed")
        gLog(changeTxt, indent:=True, ascend:=True)
        Dim tmpKtList = ktList
        Dim tmpCountList = countList
        kl.Keys.ForEach(Sub(key) recordModification(tmpKtList, tmpCountList, key.KeyType))
        ktList = tmpKtList
        countList = tmpCountList
        summarizeEntryUpdate(ktList, countList, changeTxt)
        For i = 0 To kl.KeyCount - 1
            Dim key = kl.Keys(i)
            print(0, key.toString, colorLine:=True, enStrCond:=wasAdded)
            print(0, Nothing, cond:=i = kl.KeyCount - 1, conjoin:=True, fillBorder:=False)
            gLog($"{key.toString}", indent:=True, indAmt:=4)
        Next
        gLog(descend:=True)
    End Sub

    ''' <summary> Determines the category of change associated with keys found by Diff </summary>
    ''' <param name="removedKeys"> <c> iniKeys </c> determined to have been removed from the newer version of the <c> iniSection </c> </param>
    ''' <param name="addedKeys"> <c> iniKeys </c> determined to have been added to the newer version of the <c> iniSection </c> </param>
    ''' <param name="updatedKeys"> <c> iniKeys </c> determined to have been modified in the newer version of the <c> iniSection </c> </param>
    ''' Docs last updated: 2020-09-02 | Code last updated: 2020-09-02
    Private Sub chkLsts(ByRef removedKeys As keyList, ByRef addedKeys As keyList, ByRef updatedKeys As List(Of KeyValuePair(Of iniKey, iniKey)))
        Dim akAlpha As New keyList
        Dim rkAlpha As New keyList
        For i = 0 To addedKeys.KeyCount - 1
            Dim key = addedKeys.Keys(i)
            For j = 0 To removedKeys.KeyCount - 1
                Dim skey = removedKeys.Keys(j)
                If key.compareNames(skey) Then
                    Select Case key.KeyType
                        Case "FileKey", "ExcludeKey", "RegKey"
                            Dim oldKey As New winapp2KeyParameters(key)
                            Dim newKey As New winapp2KeyParameters(skey)
                            ' If the path has changed, the key has been updated
                            If oldKey.PathString <> newKey.PathString Then updateKeys(updatedKeys, key, skey) : Exit For
                            Dim oldArgsUpper As New List(Of String)
                            Dim newArgsUpper As New List(Of String)
                            oldKey.ArgsList.ForEach(Sub(arg) oldArgsUpper.Add(arg.ToUpperInvariant))
                            newKey.ArgsList.ForEach(Sub(arg) newArgsUpper.Add(arg.ToUpperInvariant))
                            Dim oldArgs = oldArgsUpper.ToArray
                            Dim newArgs = newArgsUpper.ToArray
                            ' If the arguments aren't identical, the key has been updated 
                            If oldArgs.Except(newArgs).Any Or newArgs.Except(oldArgs).Any Then updateKeys(updatedKeys, key, skey) : Exit For
                            ' If we get this far, it's just an alphabetization change and can be ignored silently
                            akAlpha.add(skey)
                            rkAlpha.add(key)
                        Case Else
                            ' Other keys don't require such complex legwork, thankfully. If their values don't match, they've been updated
                            updateKeys(updatedKeys, key, skey, Not key.compareValues(skey))
                    End Select
                End If
            Next
        Next
        ' Update the keyLists
        For Each pair In updatedKeys
            addedKeys.remove(pair.Key)
            removedKeys.remove(pair.Value)
        Next
        addedKeys.remove(akAlpha.Keys)
        removedKeys.remove(rkAlpha.Keys)
    End Sub

    ''' <summary> Performs change tracking for chkLst </summary>
    ''' <param name="updLst"> The list of updated keys </param>
    ''' <param name="key"> An added key </param>
    ''' <param name="skey"> A removed key </param>
    ''' <param name="cond"> Indicates that the keys should be updated <br> Optional, Default: <c> True </c> </br> </param>
    ''' Docs last updated: 2020-09-02 | Code last updated: 2020-09-02
    Private Sub updateKeys(ByRef updLst As List(Of KeyValuePair(Of iniKey, iniKey)), key As iniKey, skey As iniKey, Optional cond As Boolean = True)
        If cond Then updLst.Add(New KeyValuePair(Of iniKey, iniKey)(key, skey))
    End Sub

    ''' <summary> Logs the changes that have been made to a modified entry </summary>
    ''' <param name="section"> The modified entry whose changes are being observed </param>
    ''' <param name="changeType"> The type of change to observe (as it will be described to the user) <br />
    ''' <list type="bullet">
    ''' <item> <description> <c> 0 </c>: "added"    </description> </item>
    ''' <item> <description> <c> 1 </c>: "removed"  </description> </item>
    ''' <item> <description> <c> 2 </c>: "modified" </description> </item>
    ''' <item> <description> <c> 3 </c>: "renamed to " </description> </item>
    ''' <item> <description> <c> 4 </c>: "merged into " </description> </item>
    ''' </list> </param>
    ''' <param name="changeCounter"> A pointer to the counter for the type of change being tracked </param>
    ''' <param name="renamedOrMergedEntryName"> In the event that an entry was renamed or merged, this is the name of the entry containing its keys </param>
    ''' Docs last updated: 2022-06-20 | Code last updated: 2022-06-20
    Private Sub getDiff(section As iniSection, changeType As Integer, ByRef changeCounter As Integer, Optional renamedOrMergedEntryName As String = "", Optional newSection As iniSection = Nothing, Optional mergeDict As Dictionary(Of String, List(Of String)) = Nothing)
        Dim changeTypeStrs = {"added", "removed", "modified", "renamed to ", "merged into "}
        changeCounter += 1
        Dim printColor As ConsoleColor = ConsoleColor.Cyan
        Select Case changeType
            Case 2
                printColor = ConsoleColor.Yellow
            Case 3
                printColor = ConsoleColor.Magenta
        End Select
        Dim changeStr = $"{section.Name} has been {changeTypeStrs(changeType)}{renamedOrMergedEntryName}"
        gLog(changeStr, indent:=True, leadr:=True)
        print(0, Nothing)
        print(0, changeStr, isCentered:=True, fillBorder:=False, colorLine:=True, useArbitraryColor:=changeType >= 2, enStrCond:=changeType < 1, arbitraryColor:=printColor)
        print(0, Nothing, conjoin:=True, fillBorder:=False)
        If ShowFullEntries Then
            print(0, "")
            print(0, "Old entry: ", cond:=changeType >= 3)
            For Each line In section.ToString.Split(CChar(vbCrLf))
                gLog(line.Replace(vbLf, ""), indent:=True, indAmt:=4)
                print(0, line.Replace(vbLf, ""))
            Next
            print(0, "")
            If changeType >= 3 Then
                print(0, "")
                print(0, If(changeType = 3, "Renamed entry: ", "Merged entry: "), cond:=changeType >= 3)
                gLog(If(changeType = 3, "Renamed Entry: ", "Merged entry: "))
                For Each line In newSection.ToString.Split(CChar(vbCrLf))
                    gLog(line.Replace(vbLf, ""), indent:=True, indAmt:=4)
                    print(0, line.Replace(vbLf, ""))
                Next
                print(0, "")
            End If
        End If
    End Sub

End Module