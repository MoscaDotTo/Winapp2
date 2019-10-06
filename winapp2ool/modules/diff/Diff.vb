'    Copyright (C) 2018-2019 Robbie Ward
' 
'    This file is a part of Winapp2ool
' 
'    Winapp2ool is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    Winap2ool is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with Winapp2ool.  If not, see <http://www.gnu.org/licenses/>.
Option Strict On
''' <summary> Performs a diff operation on two winapp2.ini files </summary>
Module Diff
    ''' <summary> The old or local version of winapp2.ini to be diffed </summary>
    Public Property DiffFile1 As iniFile = New iniFile(Environment.CurrentDirectory, "winapp2.ini", mExist:=True)
    ''' <summary> The new or remote version of winapp2.ini to be diffed against </summary>
    Public Property DiffFile2 As iniFile = New iniFile(Environment.CurrentDirectory, "", mExist:=True)
    ''' <summary> The path to which the log will optionally be saved </summary>
    Public Property DiffFile3 As iniFile = New iniFile(Environment.CurrentDirectory, "diff.txt")
    ''' <summary> Indicates that a remote winapp2.ini should be downloaded to use as <c> DiffFile2 </c> </summary>
    Public Property DownloadDiffFile As Boolean = Not isOffline
    ''' <summary> Indicates that the diff output should be saved to disk </summary>
    Public Property SaveDiffLog As Boolean = False
    ''' <summary> Indicates that the module settings have been modified from their defaults </summary>
    Public Property ModuleSettingsChanged As Boolean = False
    ''' <summary> Indicates that the remote file should be trimmed for the local system before diffing </summary>
    Public Property TrimRemoteFile As Boolean = Not isOffline

    ''' <summary> The number of entries Diff determines to have been added between versions </summary>
    Public Property AddedEntryCount As Integer = 0
    ''' <summary> The number of entries Diff determines to have been modified between versions </summary>
    Public Property ModifiedEntryCount As Integer = 0
    ''' <summary> The number of entries Diff determines to have been removed between versions </summary>
    Public Property RemovedEntryCount As Integer = 0

    ''' <summary> Indicates that full entries should be printed in the Diff output. <br/> <br/> Called "verbose mode" in the menu </summary>
    Public Property ShowFullEntries As Boolean = False
    ''' <summary> Holds the log from the most recent run of the Differ to display back to the user </summary>
    Private Property MostRecentDiffLog As String = ""

    ''' <summary> Handles the commandline args for Diff </summary>
    '''  Diff args:
    ''' -d          : download the latest winapp2.ini
    ''' -ncc        : download the latest non-ccleaner winapp2.ini (implies -d)
    ''' -savelog    : save the diff.txt log
    Public Sub handleCmdLine()
        initDefaultSettings()
        handleDownloadBools(DownloadDiffFile)
        ' Make sure we have a name set for the new file if we're downloading or else the diff will not run
        If DownloadDiffFile Then DiffFile2.Name = If(RemoteWinappIsNonCC, "Online non-ccleaner winapp2.ini", "Online winapp2.ini")
        invertSettingAndRemoveArg(SaveDiffLog, "-savelog")
        getFileAndDirParams(DiffFile1, DiffFile2, DiffFile3)
        If Not DiffFile2.Name = "" Then initDiff()
    End Sub

    ''' <summary> Restores the default state of the module's parameters </summary>
    Private Sub initDefaultSettings()
        DownloadDiffFile = Not isOffline
        TrimRemoteFile = Not isOffline
        ShowFullEntries = False
        DiffFile3.resetParams()
        DiffFile2.resetParams()
        DiffFile1.resetParams()
        SaveDiffLog = False
        ModuleSettingsChanged = False
    End Sub

    ''' <summary> Runs the Differ from outside the module </summary>
    ''' <param name="firstFile"> The local winapp2.ini file to diff against the master GitHub copy </param>
    Public Sub remoteDiff(firstFile As iniFile)
        DiffFile1 = firstFile
        DownloadDiffFile = True
        initDiff()
    End Sub

    ''' <summary> Prints the main menu to the user </summary>
    Public Sub printMenu()
        Console.WindowHeight = If(ModuleSettingsChanged, 34, 32)
        printMenuTop({"Observe the differences between two ini files"})
        print(1, "Run (default)", "Run the diff tool", enStrCond:=Not (DiffFile2.Name = "" And Not DownloadDiffFile), colorLine:=True)
        print(0, "Select Older/Local File:", leadingBlank:=True)
        print(1, "File Chooser", "Choose a new name or location for your older ini file")
        print(0, "Select Newer/Remote File:", leadingBlank:=True)
        print(5, GetNameFromDL(True), "diffing against the latest winapp2.ini version on GitHub", cond:=Not isOffline, enStrCond:=DownloadDiffFile, leadingBlank:=True)
        print(5, "Remote file trimming", "trimming the remote winapp2.ini before diffing", cond:=DownloadDiffFile = True, enStrCond:=TrimRemoteFile, trailingBlank:=True)
        print(1, "File Chooser", "Choose a new name or location for your newer ini file", Not DownloadDiffFile, isOffline, True)
        print(0, "Log Settings:")
        print(5, "Toggle Log Saving", "automatic saving of the Diff output", leadingBlank:=True, trailingBlank:=Not SaveDiffLog, enStrCond:=SaveDiffLog)
        print(1, "File Chooser (log)", "Change where Diff saves its log", SaveDiffLog, trailingBlank:=True)
        print(5, "Verbose Mode", "printing full entries in the diff output", enStrCond:=ShowFullEntries, trailingBlank:=True)
        print(0, $"Older file: {replDir(DiffFile1.Path)}")
        print(0, $"Newer file: {If(DiffFile2.Name = "" And Not DownloadDiffFile, "Not yet selected", If(DownloadDiffFile, GetNameFromDL(True), replDir(DiffFile2.Path)))}",
                                  closeMenu:=Not SaveDiffLog And Not ModuleSettingsChanged And MostRecentDiffLog = "")
        print(0, $"Log   file: {replDir(DiffFile3.Path)}", cond:=SaveDiffLog, closeMenu:=(Not ModuleSettingsChanged) And MostRecentDiffLog = "")
        print(2, "Diff", cond:=ModuleSettingsChanged, closeMenu:=MostRecentDiffLog = "")
        print(1, "Log Viewer", "Show the most recent Diff log", cond:=Not MostRecentDiffLog = "", closeMenu:=True, leadingBlank:=True)
    End Sub

    ''' <summary> Handles the user input from the main menu </summary>
    ''' <param name="input"> The user's input </param>
    Public Sub handleUserInput(input As String)
        Select Case True
            Case input = "0"
                exitModule()
            Case input = "1" Or input = ""
                If Not denyActionWithHeader(DiffFile2.Name = "" And Not DownloadDiffFile, "Please select a file against which to diff") Then initDiff()
            Case input = "2"
                changeFileParams(DiffFile1, ModuleSettingsChanged)
            Case input = "3" And Not isOffline
                toggleDownload(DownloadDiffFile, ModuleSettingsChanged)
                DiffFile2.Name = GetNameFromDL(DownloadDiffFile)
            Case input = "4" And DownloadDiffFile
                toggleSettingParam(TrimRemoteFile, "Trimming", ModuleSettingsChanged)
            Case (input = "4" And Not (DownloadDiffFile Or isOffline)) Or (input = "3" And isOffline)
                changeFileParams(DiffFile2, ModuleSettingsChanged)
            Case (input = "5" And Not isOffline) Or (input = "4" And isOffline)
                toggleSettingParam(SaveDiffLog, "Log Saving", ModuleSettingsChanged)
            Case SaveDiffLog And ((input = "6" And Not isOffline) Or (input = "5" And isOffline))
                changeFileParams(DiffFile3, ModuleSettingsChanged)
            Case input = "6" And Not SaveDiffLog Or input = "7" And SaveDiffLog
                toggleSettingParam(ShowFullEntries, "Verbose Mode", ModuleSettingsChanged)
            Case ModuleSettingsChanged And ( 'Online Case below
                                        (Not isOffline And ((Not SaveDiffLog And input = "7") Or
                                        (SaveDiffLog And input = "8"))) Or
                                        (isOffline And ((input = "5") Or (input = "6" And SaveDiffLog)))) ' Offline case
                resetModuleSettings("Diff", AddressOf initDefaultSettings)
            Case Not MostRecentDiffLog = "" And ((input = "7" And Not ModuleSettingsChanged) Or (input = "8" And ModuleSettingsChanged))
                MostRecentDiffLog = getLogSliceFromGlobal("Beginning diff", "Diff complete")
                printSlice(MostRecentDiffLog)
            Case Else
                setHeaderText(invInpStr, True)
        End Select
    End Sub

    ''' <summary> Carries out the main set of Diffing operations </summary>
    Private Sub initDiff()
        If Not enforceFileHasContent(DiffFile1) Then Exit Sub
        If DownloadDiffFile Then DiffFile2 = getRemoteIniFile(winapp2link, DiffFile2) Else If Not enforceFileHasContent(DiffFile2) Then Exit Sub
        If TrimRemoteFile And DownloadDiffFile Then
            Dim tmp As New winapp2file(DiffFile2)
            Trim.trim(tmp)
            DiffFile2.Sections = tmp.toIni.Sections
        End If
        logInitDiff()
        compareTo()
        logPostDiff()
        print(3, anyKeyStr)
        crk()
        MostRecentDiffLog = If(SaveDiffLog, getLogSliceFromGlobal("Beginning diff", "Diff complete"), "hasBeenRun")
        DiffFile3.overwriteToFile(MostRecentDiffLog, SaveDiffLog)
        setHeaderText(If(SaveDiffLog, DiffFile3.Name & " saved", "Diff complete"))
    End Sub

    '''<summary>Logs the initial portion of the diff output for the user</summary>
    Private Sub logInitDiff()
        print(3, "Diffing, please wait. This may take a moment.")
        clrConsole()
        Dim oldVersionNum = getVer(DiffFile1)
        Dim newVersionNum = getVer(DiffFile2)
        gLog($"Beginning diff between{oldVersionNum} and{newVersionNum}", ascend:=True)
        print(3, $"Changes between{oldVersionNum} and{newVersionNum}")
    End Sub

    ''' <summary>Gets the version from winapp2.ini</summary>
    ''' <param name="someFile">winapp2.ini format iniFile object</param>
    Private Function getVer(someFile As iniFile) As String
        Dim ver = If(someFile.Comments.Count > 0, someFile.Comments(0).Comment.ToString.ToLower, "000000")
        Return If(ver.Contains("version"), ver.TrimStart(CChar(";")).Replace("version:", "version"), " version not given")
    End Function

    ''' <summary> Logs and prints the summary of the Diff </summary>
    Private Sub logPostDiff()
        gLog($"Added entries: {AddedEntryCount}", indent:=True)
        gLog($"Modified entries: {ModifiedEntryCount}", indent:=True)
        gLog($"Removed entries: {RemovedEntryCount}", indent:=True)
        gLog("Diff complete", descend:=True)
        print(4, "Summary", conjoin:=True)
        print(0, $"Added entries: {AddedEntryCount}")
        print(0, $"Modified entries: {ModifiedEntryCount}")
        print(0, $"Removed entries: {RemovedEntryCount}", closeMenu:=True, trailr:=True)
    End Sub

    ''' <summary>Compares two winapp2.ini format iniFiles and builds the output for the user containing the differences</summary>
    Private Sub compareTo()
        AddedEntryCount = 0
        RemovedEntryCount = 0
        ModifiedEntryCount = 0
        Dim comparedList As New strList
        For Each section In DiffFile1.Sections.Values
            ' If we're looking at an entry in the old file and the new file contains it, and we haven't yet processed this entry
            If DiffFile2.Sections.Keys.Contains(section.Name) And Not comparedList.contains(section.Name) Then
                Dim sSection = DiffFile2.Sections(section.Name)
                ' And if that entry in the new file does not compareTo the entry in the old file, we have a modified entry
                Dim addedKeys, removedKeys As New keyList
                Dim updatedKeys As New List(Of KeyValuePair(Of iniKey, iniKey))
                If Not section.compareTo(sSection, removedKeys, addedKeys) Then
                    chkLsts(removedKeys, addedKeys, updatedKeys)
                    ' Silently ignore any entries with only alphabetization changes
                    If removedKeys.KeyCount + addedKeys.KeyCount + updatedKeys.Count = 0 Then Continue For
                    getDiff(sSection, 2, ModifiedEntryCount)
                    getChangesFromList(addedKeys, True)
                    getChangesFromList(removedKeys, False)
                    If updatedKeys.Count > 0 Then
                        print(0, "Modified:", isCentered:=True)
                        gLog("Modifed Keys:", ascend:=True, ascAmt:=2)
                        For Each pair In updatedKeys
                            gLog(pair.Key.Name, indent:=True, indAmt:=2)
                            gLog("Old: " & pair.Value.toString, indent:=True, indAmt:=3)
                            gLog("New: " & pair.Key.toString, indent:=True, indAmt:=3)
                            print(0, "Old: " & pair.Value.toString, colorLine:=True)
                            print(0, "New: " & pair.Key.toString, colorLine:=True, enStrCond:=True)
                        Next
                        gLog("", descend:=True, descAmt:=2)
                    End If
                End If
            ElseIf Not DiffFile2.Sections.Keys.Contains(section.Name) And Not comparedList.contains(section.Name) Then
                ' If we do not have the entry in the new file, it has been removed between versions 
                getDiff(section, 1, RemovedEntryCount)
            End If
            comparedList.add(section.Name)
        Next
        ' Any sections from the new file which are not found in the old file have been added
        For Each section In DiffFile2.Sections.Values
            If Not DiffFile1.Sections.Keys.Contains(section.Name) Then getDiff(section, 0, AddedEntryCount)
        Next
    End Sub

    ''' <summary> Handles the Added and Removed cases for changes </summary>
    ''' <param name="kl"> A list of iniKeys that have been added/removed</param>
    ''' <param name="wasAdded"> True if the change type is Added, False if Removed </param>
    Private Sub getChangesFromList(kl As keyList, wasAdded As Boolean)
        If kl.KeyCount = 0 Then Exit Sub
        Dim changeTxt = If(wasAdded, "Added:", "Removed:")
        print(0, changeTxt, isCentered:=True)
        gLog(changeTxt, indent:=True, ascend:=True)
        For Each key In kl.Keys
            print(0, key.toString, colorLine:=True, enStrCond:=wasAdded)
            gLog($"{key.toString}", indent:=True, indAmt:=4)
        Next
        gLog("", descend:=True)
    End Sub

    ''' <summary> Determines the category of change associated with key found by Diff </summary>
    ''' <param name="removedKeys"> <c> iniKeys </c> determined to have been removed from the newer version of the <c> iniSection </c> </param>
    ''' <param name="addedKeys"> <c> iniKeys </c> determined to have been added to the newer version of the <c> iniSection </c> </param>
    ''' <param name="updatedKeys"> <c> iniKeys </c> determined to have been modified in the newer version of the <c> iniSection </c> </param>
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
                            If Not oldKey.PathString = newKey.PathString Then
                                updateKeys(updatedKeys, key, skey)
                                Exit For
                            End If
                            oldKey.ArgsList.Sort()
                            newKey.ArgsList.Sort()
                            ' Check the number of arguments provided to the key
                            If oldKey.ArgsList.Count = newKey.ArgsList.Count Then
                                For k = 0 To oldKey.ArgsList.Count - 1
                                    ' If the args count matches but the sorted state of the args doesn't, the key has been updated
                                    If Not oldKey.ArgsList(k).Equals(newKey.ArgsList(k), StringComparison.InvariantCultureIgnoreCase) Then
                                        updateKeys(updatedKeys, key, skey)
                                        Exit For
                                    End If
                                Next
                                ' If we get this far, it's just an alphabetization change and can be ignored silently
                                akAlpha.add(skey)
                                rkAlpha.add(key)
                            Else
                                ' If the count doesn't match, something has definitely changed
                                updateKeys(updatedKeys, key, skey)
                                Exit For
                            End If
                        Case Else
                            ' Other keys don't require such complex legwork, thankfully. If their values don't match, they've been updated
                            If Not key.compareValues(skey) Then updateKeys(updatedKeys, key, skey)
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

    ''' <summary>Performs change tracking for chkLst </summary>
    ''' <param name="updLst">The list of updated keys</param>
    ''' <param name="key">An added key</param>
    ''' <param name="skey">A removed key</param>
    Private Sub updateKeys(ByRef updLst As List(Of KeyValuePair(Of iniKey, iniKey)), key As iniKey, skey As iniKey)
        updLst.Add(New KeyValuePair(Of iniKey, iniKey)(key, skey))
    End Sub

    ''' <summary> Logs the changes that have been made to a modified entry </summary>
    ''' <param name="section"> The modified entry whose changes are being observed </param>
    ''' <param name="changeType"> The type of change to observe (as it will be described to the user) <br />
    ''' <list type="bullet">
    ''' <item> <description> <c> 0 </c>: "added"    </description> </item>
    ''' <item> <description> <c> 1 </c>: "removed"  </description> </item>
    ''' <item> <description> <c> 2 </c>: "modified" </description> </item>
    ''' </list></param>
    ''' <param name="changeCounter"> A pointer to the counter for the type of change being tracked </param>
    Private Sub getDiff(section As iniSection, changeType As Integer, ByRef changeCounter As Integer)
        Dim changeTypeStrs = {"added", "removed", "modified"}
        changeCounter += 1
        gLog($"{section.Name} has been {changeTypeStrs(changeType)}", indent:=True, leadr:=True)
        print(0, $"{section.Name} has been {changeTypeStrs(changeType)}", isCentered:=True, colorLine:=changeType <= 1, enStrCond:=If(changeType = 1, False, True))
        If ShowFullEntries Then
            print(0, "")
            For Each line In section.ToString.Split(CChar(vbCrLf))
                gLog(line.Replace(vbLf, ""), indent:=True, indAmt:=4)
                print(0, line.Replace(vbLf, ""))
            Next
            print(0, "")
        End If
    End Sub
End Module