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

''' <summary>
''' A module whose purpose is to allow a user to perform a diff on two winapp2.ini files
''' </summary>
Module Diff

    ''' <summary>The old or local version of winapp2.ini to be diffed</summary>
    Public Property DiffFile1 As iniFile = New iniFile(Environment.CurrentDirectory, "winapp2.ini", mExist:=True)
    '''<summary>The new or remote version of winapp2.ini to be diffed</summary>
    Public Property DiffFile2 As iniFile = New iniFile(Environment.CurrentDirectory, "", mExist:=True)
    '''<summary>The path to which the log will optionally be saved</summary>
    Public Property DiffFile3 As iniFile = New iniFile(Environment.CurrentDirectory, "diff.txt")
    '''<summary>Indicates whether or not we are downloading a remote winapp2.ini</summary>
    Public Property DownloadDiffFile As Boolean = Not isOffline
    '''<summary>Indicates whether or not the diff output should be saved to disk</summary>
    Public Property SaveDiffLog As Boolean = False
    '''<summary>Indicates that the module settings have been changed</summary>
    Public Property ModuleSettingsChanged As Boolean = False
    ''' <summary>Holds the output that will be shown to the user and optionally saved to disk</summary>
    Private Property outputToFile As String = ""
    '''<summary>Indicates that the remote file should be trimmed for the local system before diffing</summary>
    Public Property TrimRemoteFile As Boolean = Not isOffline

    '''<summary>The number of entries Diff determines to have been added in a new version</summary>
    Public Property AddedEntryCount As Integer = 0
    '''<summary>The number of entries Diff determines to have been modified in a new version</summary>
    Public Property ModifiedEntryCount As Integer = 0
    '''<summary>The number of entries Diff determines to have been removed in a new version</summary>
    Public Property RemovedEntryCount As Integer = 0

    ''' <summary>Handles the commandline args for Diff</summary>
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

    ''' <summary>Restores the default state of the module's parameters</summary>
    Private Sub initDefaultSettings()
        DownloadDiffFile = False
        DiffFile3.resetParams()
        DiffFile2.resetParams()
        DiffFile1.resetParams()
        SaveDiffLog = False
        ModuleSettingsChanged = False
    End Sub

    ''' <summary>Runs the Differ from outside the module </summary>
    ''' <param name="firstFile">The old winapp2.ini file</param>
    Public Sub remoteDiff(firstFile As iniFile, Optional dl As Boolean = True)
        DownloadDiffFile = dl
        DiffFile1 = firstFile
        initDiff()
    End Sub

    ''' <summary>Prints the main menu to the user</summary>
    Public Sub printMenu()
        Console.WindowHeight = If(ModuleSettingsChanged, 32, 30)
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
        print(0, $"Older file: {replDir(DiffFile1.Path)}")
        print(0, $"Newer file: {If(DiffFile2.Name = "" And Not DownloadDiffFile, "Not yet selected", If(DownloadDiffFile, GetNameFromDL(True), replDir(DiffFile2.Path)))}", closeMenu:=Not SaveDiffLog And Not ModuleSettingsChanged)
        print(0, $"Log   file: {replDir(DiffFile3.Path)}", cond:=SaveDiffLog, closeMenu:=Not ModuleSettingsChanged)
        print(2, "Diff", cond:=ModuleSettingsChanged, closeMenu:=True)
    End Sub

    ''' <summary>Handles the user input from the main menu</summary>
    ''' <param name="input">The String containing the user's input from the menu</param>
    Public Sub handleUserInput(input As String)
        Select Case True
            Case input = "0"
                exitModule()
            Case input = "1" Or input = ""
                If Not denyActionWithTopper(DiffFile2.Name = "" And Not DownloadDiffFile, "Please select a file against which to diff") Then initDiff()
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
            Case ModuleSettingsChanged And ( 'Online Case below
                                        (Not isOffline And ((Not SaveDiffLog And input = "6") Or
                                        (SaveDiffLog And input = "7"))) Or
                                        (isOffline And ((input = "5") Or (input = "6" And SaveDiffLog)))) ' Offline case
                resetModuleSettings("Diff", AddressOf initDefaultSettings)
            Case Else
                setHeaderText(invInpStr, True)
        End Select
    End Sub

    ''' <summary>Carries out the main set of Diffing operations</summary>
    Private Sub initDiff()
        outputToFile = ""
        DiffFile1.validate()
        If DownloadDiffFile Then DiffFile2 = getRemoteIniFile(winapp2link)
        DiffFile2.validate()
        If Not (enforceFileHasContent(DiffFile1) And enforceFileHasContent(DiffFile2)) Then Exit Sub
        If TrimRemoteFile Then
            Dim tmp As New winapp2file(DiffFile2)
            Trim.trim(tmp)
            DiffFile2.Sections = tmp.toIni.Sections
        End If
        logInitDiff()
        differ()
        logPostDiff()
        Console.WriteLine()
        printMenuLine(bmenu(anyKeyStr))
        If Not SuppressOutput Then Console.ReadLine()
        DiffFile3.overwriteToFile(outputToFile, SaveDiffLog)
        setHeaderText("Diff Complete")
    End Sub

    '''<summary>Logs the initial portion of the diff output for the user</summary>
    Private Sub logInitDiff()
        print(3, "Diffing, please wait. This may take a moment.")
        clrConsole()
        Dim oldVersionNum = getVer(DiffFile1)
        Dim newVersionNum = getVer(DiffFile2)
        log(tmenu($"Changes made between{oldVersionNum} and{newVersionNum}"))
        log(menu(menuStr02))
        log(menu(menuStr00))
    End Sub

    ''' <summary>Gets the version from winapp2.ini</summary>
    ''' <param name="someFile">winapp2.ini format iniFile object</param>
    Private Function getVer(someFile As iniFile) As String
        Dim ver = If(someFile.Comments.Count > 0, someFile.Comments(0).Comment.ToString.ToLower, "000000")
        Return If(ver.Contains("version"), ver.TrimStart(CChar(";")).Replace("version:", "version"), " version not given")
    End Function

    ''' <summary>Performs the diff and outputs the info to the user</summary>
    Private Sub differ()
        ' Compare the files and then enumerate their changes
        Dim outList = compareTo()
        AddedEntryCount = 0
        ModifiedEntryCount = 0
        RemovedEntryCount = 0
        For Each change In outList.Items
            Select Case True
                Case change.Contains("has been added")
                    AddedEntryCount += 1
                Case change.Contains(" has been removed")
                    RemovedEntryCount += 1
                Case Else
                    ModifiedEntryCount += 1
            End Select
            log(change)
        Next
    End Sub

    ''' <summary>Logs the summary of the Diff for output to the user </summary>
    Private Sub logPostDiff()
        log(menu("Diff complete.", True))
        log(menu(menuStr03))
        log(menu("Summary", True))
        log(menu(menuStr01))
        log(menu($"Added entries: {AddedEntryCount}"))
        log(menu($"Modified entries: {ModifiedEntryCount}"))
        log(menu($"Removed entries: {RemovedEntryCount}"))
        log(menu(menuStr02))
    End Sub

    ''' <summary>Compares two winapp2.ini format iniFiles and builds the output for the user containing the differences</summary>
    Private Function compareTo() As strList
        Dim outList, comparedList As New strList
        For Each section In DiffFile1.Sections.Values
            ' If we're looking at an entry in the old file and the new file contains it, and we haven't yet processed this entry
            If DiffFile2.Sections.Keys.Contains(section.Name) And Not comparedList.contains(section.Name) Then
                Dim sSection As iniSection = DiffFile2.Sections(section.Name)
                ' And if that entry in the new file does not compareTo the entry in the old file, we have a modified entry
                Dim addedKeys, removedKeys As New keyList
                Dim updatedKeys As New List(Of KeyValuePair(Of iniKey, iniKey))
                If Not section.compareTo(sSection, removedKeys, addedKeys) Then
                    chkLsts(removedKeys, addedKeys, updatedKeys)
                    ' Silently ignore any entries with only alphabetization changes
                    If removedKeys.KeyCount + addedKeys.KeyCount + updatedKeys.Count = 0 Then Continue For
                    Dim tmp = getDiff(sSection, "modified")
                    tmp = getChangesFromList(addedKeys, tmp, $"{pNL("Added:")}")
                    tmp = getChangesFromList(removedKeys, tmp, $"{pNL("Removed:", addedKeys.KeyCount > 0)}")
                    If updatedKeys.Count > 0 Then
                        tmp += aNL($"{pNL("Modified:", removedKeys.KeyCount > 0 Or addedKeys.KeyCount > 0)}")
                        updatedKeys.ForEach(Sub(pair) appendStrs({aNL(pNL(pair.Key.Name)), $"Old:   {aNL(pair.Key.toString)}", $"New:   {aNL(pair.Value.toString)}"}, tmp))
                    End If
                    tmp += pNL(menuStr00)
                    outList.add(tmp)
                End If
            ElseIf Not DiffFile2.Sections.Keys.Contains(section.Name) And Not comparedList.contains(section.Name) Then
                ' If we do not have the entry in the new file, it has been removed between versions 
                outList.add(getDiff(section, "removed"))
            End If
            comparedList.add(section.Name)
        Next
        ' Any sections from the new file which are not found in the old file have been added
        For Each section In DiffFile2.Sections.Values
            If Not DiffFile1.Sections.Keys.Contains(section.Name) Then outList.add(getDiff(section, "added"))
        Next
        Return outList
    End Function

    ''' <summary>Handles the Added and Removed cases for changes </summary>
    ''' <param name="keyList">A list of iniKeys that have been added/removed</param>
    ''' <param name="out">The output text to be appended to</param>
    ''' <param name="changeTxt">The text to appear in the output</param>
    Private Function getChangesFromList(keyList As keyList, out As String, changeTxt As String) As String
        If keyList.KeyCount = 0 Then Return out
        out += aNL(changeTxt)
        keyList.Keys.ForEach(Sub(key) out += key.toString & Environment.NewLine)
        Return out
    End Function

    ''' <summary>Observes lists of added and removed keys from a section for diffing, adds any changes to the updated key </summary>
    ''' <param name="removedKeys">The list of iniKeys that were removed from the newer version of the file</param>
    ''' <param name="addedKeys">The list of iniKeys that were added to the newer version of the file</param>
    ''' <param name="updatedKeys">The list containing iniKeys rationalized by this function as having been updated rather than added or removed</param>
    Private Sub chkLsts(ByRef removedKeys As keyList, ByRef addedKeys As keyList, ByRef updatedKeys As List(Of KeyValuePair(Of iniKey, iniKey)))
        Dim akAlpha As New keyList
        Dim rkAlpha As New keyList
        For i As Integer = 0 To addedKeys.KeyCount - 1
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

    ''' <summary>Returns a string containing a menu box listing the change type and entry, followed by the entry's toString</summary>
    ''' <param name="section">an iniSection object to be diffed</param>
    ''' <param name="changeType">The type of change to observe</param>
    Private Function getDiff(section As iniSection, changeType As String) As String
        Dim out = ""
        appendStrs({aNL(mkMenuLine($"{section.Name} has been {changeType}.", "c")), aNL(aNL(mkMenuLine(menuStr02, ""))), aNL(section.ToString)}, out)
        If Not changeType = "modified" Then out += menuStr00
        Return out
    End Function

    ''' <summary>Saves a String to the log file</summary>
    ''' <param name="toLog">The string to be appended to the log</param>
    Private Sub log(toLog As String)
        cwl(toLog)
        outputToFile += aNL(toLog)
    End Sub
End Module