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
Imports System.IO

''' <summary>
''' A module whose purpose is to perform some housekeeping on ccleaner.ini to help clean up after winapp2.ini
''' </summary>
Module CCiniDebug
    ' File Handlers
    Dim winappFile As iniFile = New iniFile(Environment.CurrentDirectory, "winapp2.ini")
    Dim ccFile As iniFile = New iniFile(Environment.CurrentDirectory, "ccleaner.ini")
    Dim outputFile As iniFile = New iniFile(Environment.CurrentDirectory, "ccleaner.ini", "ccleaner-debugged.ini")
    ' Module parameters
    Dim pruneFile As Boolean = True
    Dim saveFile As Boolean = True
    Dim sortFile As Boolean = True
    Dim settingsChanged As Boolean = False

    ''' <summary>Restores the default state of the module's parameters</summary>
    Private Sub initDefaultSettings()
        ccFile.resetParams()
        outputFile.resetParams()
        winappFile.resetParams()
        pruneFile = True
        saveFile = True
        sortFile = True
        settingsChanged = False
    End Sub

    ''' <summary>Handles the commandline args for CCiniDebug</summary>
    '''  CCiniDebug args:
    ''' -noprune    : disable pruning of stale winapp2.ini entries
    ''' -nosort     : disable sorting ccleaner.ini alphabetically
    ''' -nosave     : disable saving the modified ccleaner.ini back to file
    Public Sub handleCmdlineArgs()
        initDefaultSettings()
        invertSettingAndRemoveArg(pruneFile, "-noprune")
        invertSettingAndRemoveArg(sortFile, "-nosort")
        invertSettingAndRemoveArg(saveFile, "-nosave")
        getFileAndDirParams(winappFile, ccFile, outputFile)
        initDebug()
    End Sub

    ''' <summary>Prints the CCiniDebug menu to the user</summary>
    Public Sub printMenu()
        printMenuTop({"Sort alphabetically the contents of ccleaner.ini and prune stale winapp2.ini settings"})
        print(1, "Run (default)", "Debug ccleaner.ini", trailingBlank:=True)
        print(5, "Toggle Pruning", "removal of dead winapp2.ini settings", enStrCond:=pruneFile)
        print(5, "Toggle Saving", "automatic saving of changes made by CCiniDebug", enStrCond:=saveFile)
        print(5, "Toggle Sorting", "alphabetical sorting of ccleaner.ini", enStrCond:=sortFile, trailingBlank:=True)
        print(1, "File Chooser (ccleaner.ini)", "Choose a new ccleaner.ini name or location")
        print(1, "File Chooser (winapp2.ini)", "Choose a new winapp2.ini name or location", pruneFile, trailingBlank:=Not saveFile)
        print(1, "File Chooser (save)", "Change where CCiniDebug saves its changes", saveFile, trailingBlank:=True)
        print(0, $"Current ccleaner.ini:  {replDir(ccFile.path)}")
        print(0, $"Current winapp2.ini:   {replDir(winappFile.path)}", cond:=pruneFile)
        print(0, $"Current save location: {replDir(outputFile.path)}", cond:=saveFile, closeMenu:=Not settingsChanged)
        print(2, "CCiniDebug", cond:=settingsChanged, closeMenu:=True)
    End Sub

    ''' <summary>Handles the user's input from the menu</summary>
    ''' <param name="input">The string containing the user's input</param>
    Public Sub handleUserInput(input As String)
        Select Case True
            Case input = "0"
                exitModule()
            Case (input = "1" Or input = "") And (pruneFile Or saveFile Or sortFile)
                initDebug()
            Case input = "2"
                toggleSettingParam(pruneFile, "Pruning", settingsChanged)
            Case input = "3"
                toggleSettingParam(saveFile, "Autosaving", settingsChanged)
            Case input = "4"
                toggleSettingParam(sortFile, "Sorting", settingsChanged)
            Case input = "5"
                changeFileParams(ccFile, settingsChanged)
            Case input = "6" And pruneFile
                changeFileParams(winappFile, settingsChanged)
            Case saveFile And ((input = "6" And Not pruneFile) Or (input = "7" And pruneFile))
                changeFileParams(outputFile, settingsChanged)
            Case settingsChanged And ((input = "6" And Not (pruneFile Or saveFile)) Or (input = "7" And (pruneFile Xor saveFile)) Or (input = "8" And pruneFile And saveFile))
                resetModuleSettings("CCiniDebug", AddressOf initDefaultSettings)
            Case Not (pruneFile Or saveFile Or sortFile)
                setHeaderText("Please enable at least one option", True)
            Case Else
                setHeaderText(invInpStr, True)
        End Select
    End Sub

    ''' <summary>Performs the debugging process</summary>
    Private Sub initDebug()
        ccFile.validate()
        If pruneFile Then winappFile.validate()
        If pendingExit() Then Exit Sub
        clrConsole()
        printMenuLine(tmenu("CCiniDebug Results"))
        printMenuLine(menuStr03)
        If pruneFile Then prune(ccFile.Sections("Options"))
        If sortFile Then sortCC()
        If saveFile Then outputFile.overwriteToFile(ccFile.toString)
        print(0, $"{If(saveFile, $"{outputFile.Name} saved", "Analysis complete")}. {anyKeyStr}", isCentered:=True, closeMenu:=True)
        If Not SuppressOutput Then Console.ReadKey()
    End Sub

    ''' <summary>Scans for and removes stale winapp2.ini entry settings from the Options section of a ccleaner.ini file</summary>
    ''' <param name="optionsSec">The iniSection object containing the Options from ccleaner.ini</param>
    Private Sub prune(ByRef optionsSec As iniSection)
        print(0, $"Scanning {ccFile.Name} for settings left over from removed winapp2.ini entries", leadingBlank:=True, trailingBlank:=True)
        Dim tbTrimmed As New List(Of Integer)
        For i As Integer = 0 To optionsSec.Keys.keyCount - 1
            Dim optionStr As String = optionsSec.Keys.Keys(i).toString
            ' Only operate on (app) keys belonging to winapp2.ini
            If optionStr.StartsWith("(App)") And optionStr.Contains("*") Then
                Dim toRemove As New List(Of String) From {"(App)", "=True", "=False"}
                toRemove.ForEach(Sub(param) optionStr = optionStr.Replace(param, ""))
                If Not winappFile.Sections.ContainsKey(optionStr) Then
                    tbTrimmed.Add(i)
                End If
            End If
        Next
        print(0, $"{tbTrimmed.Count} orphaned settings detected", leadingBlank:=True, trailingBlank:=True)
        optionsSec.removeKeys(tbTrimmed)
    End Sub

    ''' <summary>Sorts the keys in the Options (only) section of ccleaner.ini</summary>
    Private Sub sortCC()
        Dim lineList As List(Of String) = ccFile.Sections("Options").getKeysAsList
        lineList.Sort()
        lineList.Insert(0, "[Options]")
        ccFile.Sections("Options") = New iniSection(lineList)
    End Sub
End Module