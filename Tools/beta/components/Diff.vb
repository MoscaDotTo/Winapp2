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
    Dim nFile As iniFile = New iniFile(Environment.CurrentDirectory, "", "winapp2.ini")
    Dim logFile As iniFile = New iniFile(Environment.CurrentDirectory, "diff.txt")
    Dim outputToFile As String

    'Menu settings
    Dim settingsChanged As Boolean = False

    'Boolean module parameters
    Dim downloadFile As Boolean = False
    Dim downloadNCC As Boolean = False
    Dim saveLog As Boolean = False

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
        d = downloadFile
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
        downloadFile = d
        downloadNCC = dncc
        saveLog = sl
        initDiff()
    End Sub

    'Restore all the module settings to their default state
    ''' <summary>
    ''' Restores the default state of the module's parameters
    ''' </summary>
    Private Sub initDefaultSettings()
        oFile.resetParams()
        nFile.resetParams()
        logFile.resetParams()
        downloadFile = False
        downloadNCC = False
        saveLog = False
        settingsChanged = False
    End Sub

    ''' <summary>
    ''' Resets the settings to their default state
    ''' </summary>
    Private Sub resetSettings()
        initDefaultSettings()
        menuTopper = "Diff settings have been reset to their defaults"
    End Sub

    ''' <summary>
    ''' Prints the main menu to the user
    ''' </summary>
    Private Sub printMenu()
        printMenuTop({"Observe the differences between two ini files"}, True)
        printMenuOpt("Run (default)", "Run the diff tool")

        printBlankMenuLine()
        printMenuLine("Select Older/Local File:", "l")
        printMenuOpt("winapp2.ini", "Use the default name")
        printMenuOpt("File Chooser", "Choose a new name or location for your older ini file")

        printBlankMenuLine()
        printMenuLine("Select Newer/Remote File:", "l")
        printMenuOpt("winapp2.ini (online)", "Diff against the latest winapp2.ini version on GitHub")
        printMenuOpt("winapp2.ini (non-ccleaner)", "Diff against the latest non-ccleaner winapp2.ini version on GitHub")
        printMenuOpt("File Chooser", "Choose a new name of location for your newer ini file")

        printBlankMenuLine()
        printMenuLine("Log Settings:", "l")
        printMenuOpt("Toggle Log Saving", enStr(saveLog) & " automatic saving of the Diff output")
        printIf(saveLog, "opt", "File Chooser (log)", "Change where Diff saves its log")

        printBlankMenuLine()
        printMenuLine("Older file: " & replDir(oFile.path), "l")
        printMenuLine("Newer file: " & If(nFile.name = "", "Not yet selected", replDir(nFile.path)), "l")

        printIf(settingsChanged, "reset", "Diff", "")
        printMenuLine(menuStr02)
    End Sub

    ''' <summary>
    ''' The main event loop for the Differ 
    ''' </summary>
    Public Sub main()
        initMenu("Diff", 35)
        Console.WindowHeight = 40
        outputToFile = ""

        Do Until exitCode
            Console.Clear()
            printMenu()
            Console.WriteLine()
            Console.Write(promptStr)

            Dim input As String = Console.ReadLine()
            handleUserInput(input)
        Loop
        revertMenu()
    End Sub

    Private Sub handleUserInput(input As String)
        Select Case True
            Case input = "0"
                Console.WriteLine("Exiting diff...")
                exitCode = True
            Case input = "1" Or input = ""
                If Not denyActionWithTopper(nFile.name = "", "Please select a file against which to diff") Then initDiff()
            Case input = "2"
                oFile.name = "winapp2.ini"
            Case input = "3"
                changeFileParams(oFile, settingsChanged)
            Case input = "4"
                If Not denySettingOffline() Then
                    toggleSettingParam(downloadFile, "Download ", settingsChanged)
                    nFile.name = "Online"
                End If
            Case input = "5"
                If Not denySettingOffline() Then
                    Select Case True
                        Case (Not downloadFile And Not downloadNCC) Or (downloadFile And downloadNCC)
                            toggleSettingParam(downloadFile, "Download ", settingsChanged)
                            toggleSettingParam(downloadNCC, "Download Non-CCleaner ", settingsChanged)
                        Case downloadFile And Not downloadNCC
                            toggleSettingParam(downloadNCC, "Download Non-CCleaner ", settingsChanged)
                    End Select
                    nFile.name = If(downloadNCC, "Online (non-ccleaner)", "")
                End If
            Case input = "6"
                changeFileParams(nFile, settingsChanged)
            Case input = "7"
                toggleSettingParam(saveLog, "Log Saving ", settingsChanged)
            Case input = "8" And saveLog
                changeFileParams(logFile, settingsChanged)
            Case (input = "9" And (settingsChanged And saveLog)) Or (input = "8" And Not saveLog And settingsChanged)
                resetSettings()
            Case Else
                menuTopper = invInpStr
        End Select
    End Sub

    ''' <summary>
    ''' Carries out the main set of Diffing operations
    ''' </summary>
    Private Sub initDiff()
        oFile.validate()
        If downloadFile Then
            'Make sure we capture the name before overriding it on the next line so the menu works properly
            Dim tmpName As String = nFile.name
            nFile = getRemoteWinapp(downloadNCC)
            nFile.name = tmpName
        Else
            nFile.validate()
        End If
        differ()
        If saveLog Then logFile.overwriteToFile(outputToFile)
        Console.Clear()
        menuTopper = "Diff Complete"
    End Sub

    ''' <summary>
    ''' Performs the diff and outputs the info to the user
    ''' </summary>
    Private Sub differ()
        If exitCode Then Exit Sub
        Console.Clear()

        'collect & verify version #s and print them out for the menu
        Dim fver As String = If(oFile.comments.Count > 0, oFile.comments(0).comment.ToString, "000000")
        fver = If(fver.ToLower.Contains("version"), fver.TrimStart(CChar(";")).Replace("Version:", "version"), " version not given")

        Dim sver As String = If(nFile.comments.Count > 0, nFile.comments(0).comment.ToString, "000000")
        sver = If(sver.ToLower.Contains("version"), sver.TrimStart(CChar(";")).Replace("Version:", "version"), " version not given")

        outputToFile = appendNewLine(tmenu("Changes made between" & fver & " and" & sver))
        outputToFile += appendNewLine(menu(menuStr02))
        outputToFile += appendNewLine(menu(menuStr00))

        'compare the files and then ennumerate their changes
        Dim outList As List(Of String) = compareTo()
        Dim remCt As Integer = 0
        Dim modCt As Integer = 0
        Dim addCt As Integer = 0
        For Each change In outList
            If change.Contains("has been added.") Then
                addCt += 1
            ElseIf change.Contains("has been removed") Then
                remCt += 1
            Else
                modCt += 1
            End If
            outputToFile += appendNewLine(change)
        Next

        'Print the summary to the user
        outputToFile += appendNewLine(menu("Diff complete.", "c"))
        outputToFile += appendNewLine(menu(menuStr03))
        outputToFile += appendNewLine(menu("Summary", "c"))
        outputToFile += appendNewLine(menu(menuStr01))
        outputToFile += appendNewLine(menu("Added entries: " & addCt, "l"))
        outputToFile += appendNewLine(menu("Modified entries: " & modCt, "l"))
        outputToFile += appendNewLine(menu("Removed entries: " & remCt, "l"))
        outputToFile += appendNewLine(menu(menuStr02))
        If Not suppressOutput Then Console.Write(outputToFile)

        Console.WriteLine()
        printMenuLine(bmenu("Press any key to return to the winapp2ool menu.", "l"))
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

                'and if that entry in the new file does not compareTo the entry in the old file, we have a modified entry
                Dim addedKeys, removedKeys As New List(Of iniKey)
                Dim updatedKeys As New List(Of KeyValuePair(Of iniKey, iniKey))
                If Not section.compareTo(sSection, removedKeys, addedKeys) Then
                    chkLsts(removedKeys, addedKeys, updatedKeys)
                    'Silently ignore any entries with only alphabetization changes
                    If removedKeys.Count + addedKeys.Count + updatedKeys.Count = 0 Then Continue For
                    Dim tmp As String = getDiff(sSection, "modified.")
                    getChangesFromList(addedKeys, tmp, prependNewLines(True) & "Added:")
                    getChangesFromList(removedKeys, tmp, prependNewLines(addedKeys.Count > 0) & "Removed:")
                    If updatedKeys.Count > 0 Then
                        tmp += appendNewLine(prependNewLines(removedKeys.Count > 0 Or addedKeys.Count > 0) & "Modified:")
                        For Each pair In updatedKeys
                            tmp += appendNewLine(prependNewLines(True) & pair.Key.name)
                            tmp += "old:   " & appendNewLine(pair.Key.toString)
                            tmp += "new:   " & appendNewLine(pair.Value.toString)
                        Next
                    End If
                    tmp += prependNewLines(False) & menuStr00
                    outList.Add(tmp)
                End If

            ElseIf Not nFile.sections.Keys.Contains(section.name) And Not comparedList.Contains(section.name) Then
                'If we do not have the entry in the new file, it has been removed between versions 
                outList.Add(getDiff(section, "removed."))
            End If
            comparedList.Add(section.name)
        Next

        For Each section In nFile.sections.Values
            'Any sections from the new file which are not found in the old file have been added
            If Not oFile.sections.Keys.Contains(section.name) Then outList.Add(getDiff(section, "added."))
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
        If keyList.Count > 0 Then
            out += changeTxt
            For Each key In keyList
                out += prependNewLines(True) & key.toString
            Next
        End If
    End Sub

    ''' <summary>
    ''' Observes lists of added and removed keys from a section for diffing, adds any changes to the updated key 
    ''' </summary>
    ''' <param name="removedKeys"></param>
    ''' <param name="addedKeys"></param>
    ''' <param name="updatedKeys"></param>
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
        out += mkMenuLine(section.name & " has been " & changeType, "c") & Environment.NewLine
        out += mkMenuLine(menuStr02, "") & Environment.NewLine & Environment.NewLine
        out += section.ToString & Environment.NewLine
        If Not changeType = "modified." Then out += menuStr00
        Return out
    End Function
End Module