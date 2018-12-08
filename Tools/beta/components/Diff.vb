'    Copyright (C) 2018 Robbie Ward
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
Imports System.IO

''' <summary>
''' A module whose purpose is to allow a user to perform a diff on two winapp2.ini files
''' </summary>
Module Diff
    'File handlers
    Dim oFile As iniFile = New iniFile(Environment.CurrentDirectory, "winapp2.ini")
    Dim nFile As iniFile = New iniFile(Environment.CurrentDirectory, "")
    Dim logFile As iniFile = New iniFile(Environment.CurrentDirectory, "diff.txt")
    Dim outputToFile As String
    'Module parameters
    Dim download As Boolean = False
    Dim downloadNCC As Boolean = False
    Dim saveLog As Boolean = False
    Dim settingsChanged As Boolean = False

    ''' <summary>
    ''' Initializes the default module settings and returns references to them to the calling function
    ''' </summary>
    ''' <param name="firstFile">The old winapp2.ini file</param>
    ''' <param name="secondFile">The new winapp2.ini file</param>
    ''' <param name="thirdFile">The log file</param>
    ''' <param name="d">The boolean representing whether or not a file should be downloaded</param>
    ''' <param name="dncc">The boolean representing whether or not the non-ccleaner file should be downloaded</param>
    ''' <param name="sl">The boolean representing whether or not we should save our log</param>
    Public Sub initDiffParams(ByRef firstFile As iniFile, ByRef secondFile As iniFile, ByRef thirdFile As iniFile, ByRef d As Boolean, ByRef dncc As Boolean, ByRef sl As Boolean)
        initDefaultSettings()
        firstFile = oFile
        secondFile = nFile
        thirdFile = logFile
        d = download
        dncc = downloadNCC
        sl = saveLog
    End Sub

    ''' <summary>
    ''' Runs the Differ from outside the module
    ''' </summary>
    ''' <param name="firstFile">The old winapp2.ini file</param>
    ''' <param name="secondFile">The new winapp2.ini file</param>
    ''' <param name="thirdFile">The log file</param>
    ''' <param name="d">The boolean representing whether or not a file should be downloaded</param>
    ''' <param name="dncc">The boolean representing whether or not the non-ccleaner file should be downloaded</param>
    ''' <param name="sl">The boolean representing whether or not we should save our log</param>
    Public Sub remoteDiff(ByRef firstFile As iniFile, secondFile As iniFile, thirdFile As iniFile, d As Boolean, dncc As Boolean, sl As Boolean)
        oFile = firstFile
        nFile = secondFile
        logFile = thirdFile
        download = d
        downloadNCC = dncc
        saveLog = sl
        initDiff()
    End Sub

    ''' <summary>
    ''' Restores the default state of the module's parameters
    ''' </summary>
    Private Sub initDefaultSettings()
        oFile.resetParams()
        nFile.resetParams()
        logFile.resetParams()
        download = False
        downloadNCC = False
        saveLog = False
        settingsChanged = False
    End Sub

    ''' <summary>
    ''' Prints the main menu to the user
    ''' </summary>
    Public Sub printMenu()
        Console.WindowHeight = If(settingsChanged, 32, 30)
        printMenuTop({"Observe the differences between two ini files"})
        print(1, "Run (default)", "Run the diff tool")
        print(0, "Select Older/Local File:", leadingBlank:=True)
        print(1, "winapp2.ini", "Use the default name")
        print(1, "File Chooser", "Choose a new name or location for your older ini file")
        print(0, "Select Newer/Remote File:", leadingBlank:=True)
        print(1, "winapp2.ini (online)", $"{enStr(download)} diffing against the latest winapp2.ini version on GitHub", cond:=Not isOffline, leadingBlank:=True)
        print(1, "winapp2.ini (non-ccleaner)", $"{enStr(downloadNCC)} diffing against the latest non-ccleaner winapp2.ini version on GitHub", download, trailingBlank:=True)
        print(1, "File Chooser", "Choose a new name or location for your newer ini file", Not download, isOffline, True)
        print(0, "Log Settings:")
        print(1, "Toggle Log Saving", $"{enStr(saveLog)} automatic saving of the Diff output", leadingBlank:=True, trailingBlank:=Not saveLog)
        print(1, "File Chooser (log)", "Change where Diff saves its log", saveLog, trailingBlank:=True)
        print(0, $"Older file: {replDir(oFile.path)}")
        print(0, $"Newer file: {If(nFile.name = "" And Not download, "Not yet selected", If(download, GetNameFromDL(download, downloadNCC), replDir(nFile.path)))}", closeMenu:=Not saveLog And Not settingsChanged)
        print(0, $"Log   file: {replDir(logFile.path)}", cond:=saveLog, closeMenu:=Not settingsChanged)
        print(2, "Diff", cond:=settingsChanged, closeMenu:=True)
    End Sub

    ''' <summary>
    ''' Handles the user input from the main menu
    ''' </summary>
    ''' <param name="input">The String containing the user's input from the menu</param>
    Public Sub handleUserInput(input As String)
        Select Case True
            Case input = "0"
                exitModule("Diff")
            Case input = "1" Or input = ""
                If Not denyActionWithTopper(nFile.name = "" And Not download, "Please Select a file against which To diff") Then initDiff()
            Case input = "2"
                oFile.name = "winapp2.ini"
                menuHeaderText = "Old/Local filename Set"
            Case input = "3"
                changeFileParams(oFile, settingsChanged)
            Case input = "4" And Not isOffline
                toggleDownload(download, settingsChanged)
                nFile.name = If(download, GetNameFromDL(download, downloadNCC), "")
                If downloadNCC And Not download Then downloadNCC = False
            Case input = "5" And download
                toggleDownload(downloadNCC, settingsChanged)
                nFile.name = If(download, GetNameFromDL(download, downloadNCC), "")
            Case (input = "5" And Not (download Or isOffline)) Or (input = "4" And isOffline)
                changeFileParams(nFile, settingsChanged)
            Case (input = "6" And Not isOffline) Or (input = "5" And isOffline)
                toggleSettingParam(saveLog, "Log Saving ", settingsChanged)
            Case saveLog And ((input = "7" And Not isOffline) Or (input = "6" And isOffline))
                changeFileParams(logFile, settingsChanged)
            Case settingsChanged And (Not isOffline And ((input = "7" And Not saveLog) Or (input = "8" And saveLog)) Or (isOffline And (input = "6") Or (input = "7" And saveLog)))
                resetModuleSettings("Diff", AddressOf initDefaultSettings)
            Case Else
                menuHeaderText = invInpStr
        End Select
    End Sub

    ''' <summary>
    ''' Carries out the main set of Diffing operations
    ''' </summary>
    Private Sub initDiff()
        outputToFile = ""
        oFile.validate()
        If download Then nFile = getRemoteWinapp(downloadNCC)
        nFile.validate()
        If pendingExit() Then Exit Sub
        differ()
        If saveLog Then logFile.overwriteToFile(outputToFile)
        menuHeaderText = "Diff Complete"
    End Sub

    ''' <summary>
    ''' Gets the version from winapp2.ini
    ''' </summary>
    ''' <param name="someFile">winapp2.ini format iniFile object</param>
    ''' <returns></returns>
    Private Function getVer(someFile As iniFile) As String
        Dim ver As String = If(someFile.comments.Count > 0, someFile.comments(0).comment.ToString.ToLower, "000000")
        Return If(ver.Contains("version"), ver.TrimStart(CChar(";")).Replace("version:", "version"), " version not given")
    End Function

    ''' <summary>
    ''' Performs the diff and outputs the info to the user
    ''' </summary>
    Private Sub differ()
        print(3, "Diffing, please wait. This may take a moment.")
        Console.Clear()
        'Collect & verify version #s and print them out for the menu
        Dim fver As String = getVer(oFile)
        Dim sver As String = getVer(nFile)
        log(tmenu($"Changes made between{fver} and{sver}"))
        log(menu(menuStr02))
        log(menu(menuStr00))
        'Compare the files and then ennumerate their changes
        Dim outList As List(Of String) = compareTo()
        Dim remCt As Integer = 0
        Dim modCt As Integer = 0
        Dim addCt As Integer = 0
        For Each change In outList
            If change.Contains("has been added") Then
                addCt += 1
            ElseIf change.Contains("has been removed") Then
                remCt += 1
            Else
                modCt += 1
            End If
            log(change)
        Next
        'Print the summary to the user
        log(menu("Diff complete.", True))
        log(menu(menuStr03))
        log(menu("Summary", True))
        log(menu(menuStr01))
        log(menu($"Added entries: {addCt}"))
        log(menu($"Modified entries: {modCt}"))
        log(menu($"Removed entries: {remCt}"))
        log(menu(menuStr02))
        Console.WriteLine()
        printMenuLine(bmenu(anyKeyStr))
        Console.ReadKey()
    End Sub

    ''' <summary>
    ''' Compares two winapp2.ini format iniFiles and builds the output for the user containing the differences
    ''' </summary>
    ''' <returns></returns>
    Private Function compareTo() As List(Of String)
        Dim outList, comparedList As New List(Of String)
        For Each section In oFile.sections.Values
            'If we're looking at an entry in the old file and the new file contains it, and we haven't yet processed this entry
            If nFile.sections.Keys.Contains(section.name) And Not comparedList.Contains(section.name) Then
                Dim sSection As iniSection = nFile.sections(section.name)
                'And if that entry in the new file does not compareTo the entry in the old file, we have a modified entry
                Dim addedKeys, removedKeys As New List(Of iniKey)
                Dim updatedKeys As New List(Of KeyValuePair(Of iniKey, iniKey))
                If Not section.compareTo(sSection, removedKeys, addedKeys) Then
                    chkLsts(removedKeys, addedKeys, updatedKeys)
                    'Silently ignore any entries with only alphabetization changes
                    If removedKeys.Count + addedKeys.Count + updatedKeys.Count = 0 Then Continue For
                    Dim tmp As String = getDiff(sSection, "modified")
                    getChangesFromList(addedKeys, tmp, $"{prependNewLines()}Added:")
                    getChangesFromList(removedKeys, tmp, $"{prependNewLines(addedKeys.Count > 0)}Removed:")
                    If updatedKeys.Count > 0 Then
                        tmp += appendNewLine($"{prependNewLines(removedKeys.Count > 0 Or addedKeys.Count > 0)}Modified:")
                        For Each pair In updatedKeys
                            appendStrs({appendNewLine(prependNewLines() & pair.Key.name), $"Old:   {appendNewLine(pair.Key.toString)}", $"New:   {appendNewLine(pair.Value.toString)}"}, tmp)
                        Next
                    End If
                    tmp += prependNewLines(False) & menuStr00
                    outList.Add(tmp)
                End If
            ElseIf Not nFile.sections.Keys.Contains(section.name) And Not comparedList.Contains(section.name) Then
                'If we do not have the entry in the new file, it has been removed between versions 
                outList.Add(getDiff(section, "removed"))
            End If
            comparedList.Add(section.name)
        Next
        For Each section In nFile.sections.Values
            'Any sections from the new file which are not found in the old file have been added
            If Not oFile.sections.Keys.Contains(section.name) Then outList.Add(getDiff(section, "added"))
        Next
        Return outList
    End Function

    ''' <summary>
    ''' Handles the Added and Removed cases for changes 
    ''' </summary>
    ''' <param name="keyList">A list of iniKeys that have been added/removed</param>
    ''' <param name="out">The output text to be appended to</param>
    ''' <param name="changeTxt">The text to appear in the output</param>
    Private Sub getChangesFromList(keyList As List(Of iniKey), ByRef out As String, changeTxt As String)
        If keyList.Count = 0 Then Exit Sub
        out += appendNewLine(changeTxt)
        For Each key In keyList
            out += appendNewLine(key.toString)
        Next
    End Sub

    ''' <summary>
    ''' Observes lists of added and removed keys from a section for diffing, adds any changes to the updated key 
    ''' </summary>
    ''' <param name="removedKeys">The list of iniKeys that were removed</param>
    ''' <param name="addedKeys">The list of iniKeys that were added</param>
    ''' <param name="updatedKeys">The list containing iniKeys rationalized by this function as having been updated rather than added or removed</param>
    Private Sub chkLsts(ByRef removedKeys As List(Of iniKey), ByRef addedKeys As List(Of iniKey), ByRef updatedKeys As List(Of KeyValuePair(Of iniKey, iniKey)))
        Dim rkTemp, akTemp As New List(Of iniKey)
        rkTemp = removedKeys.ToList
        akTemp = addedKeys.ToList
        For Each key In removedKeys
            For Each skey In addedKeys
                If key.name.ToLower = skey.name.ToLower Then
                    Dim oldKey As New winapp2KeyParameters(key)
                    Dim newKey As New winapp2KeyParameters(skey)
                    'Make sure the given path hasn't changed
                    If Not oldKey.paramString = newKey.paramString Then
                        updateKeys(updatedKeys, akTemp, rkTemp, key, skey)
                        Continue For
                    End If
                    oldKey.argsList.Sort()
                    newKey.argsList.Sort()
                    If oldKey.argsList.Count = newKey.argsList.Count Then
                        For i As Integer = 0 To oldKey.argsList.Count - 1
                            If Not oldKey.argsList(i).ToLower = newKey.argsList(i).ToLower Then
                                updateKeys(updatedKeys, akTemp, rkTemp, key, skey)
                                Exit For
                            End If
                        Next
                        'If we get this far, it's probably just an alphabetization change and can be ignored silenty
                        akTemp.Remove(skey)
                        rkTemp.Remove(key)
                    Else
                        'If the count doesn't match, something has definitely changed
                        updateKeys(updatedKeys, akTemp, rkTemp, key, skey)
                    End If
                End If
            Next
        Next
        'Update the lists
        addedKeys = akTemp
        removedKeys = rkTemp
    End Sub

    ''' <summary>
    ''' Performs change tracking for chkLst 
    ''' </summary>
    ''' <param name="updLst">The list of updated keys</param>
    ''' <param name="aKeys">The list of added keys</param>
    ''' <param name="rKeys">The list of removed keys</param>
    ''' <param name="key">A removed inikey</param>
    ''' <param name="skey">An added iniKey</param>
    Private Sub updateKeys(ByRef updLst As List(Of KeyValuePair(Of iniKey, iniKey)), ByRef aKeys As List(Of iniKey), ByRef rKeys As List(Of iniKey), key As iniKey, skey As iniKey)
        updLst.Add(New KeyValuePair(Of iniKey, iniKey)(key, skey))
        rKeys.Remove(key)
        aKeys.Remove(skey)
    End Sub

    ''' <summary>
    ''' Returns a string containing a menu box listing the change type and entry, followed by the entry's toString
    ''' </summary>
    ''' <param name="section">an iniSection object to be diffed</param>
    ''' <param name="changeType">The type of change to observe</param>
    ''' <returns></returns>
    Private Function getDiff(section As iniSection, changeType As String) As String
        Dim out As String = ""
        appendStrs({appendNewLine(mkMenuLine($"{section.name} has been {changeType}.", "c")), appendNewLine(appendNewLine(mkMenuLine(menuStr02, ""))), appendNewLine(section.ToString)}, out)
        If Not changeType = "modified" Then out += menuStr00
        Return out
    End Function

    ''' <summary>
    ''' Saves a String to the log file
    ''' </summary>
    ''' <param name="toLog">The string to be appended to the log</param>
    Private Sub log(toLog As String)
        cwl(toLog)
        outputToFile += appendNewLine(toLog)
    End Sub
End Module