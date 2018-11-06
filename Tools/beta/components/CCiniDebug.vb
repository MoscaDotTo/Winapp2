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
''' A module whose purpose is to perform some housekeeping on ccleaner.ini to help clean up after winapp2.ini
''' </summary>
Module CCiniDebug

    'File Handlers
    Dim winappFile As iniFile = New iniFile(Environment.CurrentDirectory, "winapp2.ini")
    Dim ccFile As iniFile = New iniFile(Environment.CurrentDirectory, "ccleaner.ini")
    Dim outputFile As iniFile = New iniFile(Environment.CurrentDirectory, "ccleaner.ini", "ccleaner-debugged.ini")

    'Menu vars
    Dim settingsChanged As Boolean = False

    'Boolean module parameters
    Dim pruneFile As Boolean = True
    Dim saveFile As Boolean = True
    Dim sortFile As Boolean = True

    ''' <summary>
    ''' Restores the default state of the module's parameters
    ''' </summary>
    Private Sub initDefaultSettings()
        ccFile.resetParams()
        outputFile.resetParams()
        winappFile.resetParams()
        pruneFile = True
        saveFile = True
        sortFile = True
        settingsChanged = False
    End Sub

    ''' <summary>
    ''' Initializes the default module settings and returns references to them to the calling function
    ''' </summary>
    ''' <param name="firstFile">The iniFile object to represent winapp2.ini</param>
    ''' <param name="secondFile">The iniFile object to represent ccleaner.ini</param>
    ''' <param name="thirdFile">The iniFile object containing the save path information</param>
    ''' <param name="pf">The boolean for pruning</param>
    ''' <param name="sa">The boolean for saving</param>
    ''' <param name="so">The boolean for sorting</param>
    Public Sub initCCDebugParams(ByRef firstFile As iniFile, ByRef secondFile As iniFile, ByRef thirdFile As iniFile, ByRef pf As Boolean, ByRef sa As Boolean, ByRef so As Boolean)
        initDefaultSettings()
        firstFile = winappFile
        secondFile = ccFile
        thirdFile = outputFile
        pf = pruneFile
        sa = saveFile
        so = sortFile
    End Sub

    ''' <summary>
    ''' Runs the debugger when called from outside the module
    ''' </summary>
    ''' <param name="firstfile">The winapp2.ini iniFile</param>
    ''' <param name="secondfile">The ccleaner.ini iniFile</param>
    ''' <param name="thirdfile">The iniFile with the save location</param>
    ''' <param name="pf">Boolean for pruning</param>
    ''' <param name="sa">Boolean for saving</param>
    ''' <param name="so">Boolean for sorting</param>
    Public Sub remoteCC(firstfile As iniFile, secondfile As iniFile, thirdfile As iniFile, pf As Boolean, sa As Boolean, so As Boolean)
        winappFile = firstfile
        ccFile = secondfile
        outputFile = thirdfile
        pruneFile = pf
        saveFile = sa
        sortFile = so
        initDebug()
    End Sub

    ''' <summary>
    ''' Prints the CCiniDebug menu to the user
    ''' </summary>
    Private Sub printMenu()
        printMenuTop({"Sort alphabetically the contents of ccleaner.ini and prune stale winapp2.ini settings"}, True)
        printMenuOpt("Run (default)", "Debug ccleaner.ini")

        printBlankMenuLine()
        printMenuOpt("Toggle Pruning", enStr(pruneFile) & " removal of dead winapp2.ini settings")
        printMenuOpt("Toggle Saving", enStr(saveFile) & " automatic saving of changes made by CCiniDebug")
        printMenuOpt("Toggle Sorting", enStr(sortFile) & " alphabetical sorting of ccleaner.ini")

        printBlankMenuLine()
        printMenuOpt("File Chooser (ccleaner.ini)", "Choose a new ccleaner.ini name or location")
        printIf(pruneFile, "opt", "File Chooser (winapp2.ini)", "Choose a new winapp2.ini name or location")
        printIf(saveFile, "opt", "File Chooser (save)", "Change where CCiniDebug saves its changes")

        printBlankMenuLine()
        printMenuLine("Current ccleaner.ini:  " & replDir(ccFile.path), "l")
        printIf(pruneFile, "line", "Current winapp2.ini:   " & replDir(winappFile.path), "l")
        printIf(saveFile, "line", "Current save location: " & replDir(outputFile.path), "l")
        printIf(settingsChanged, "reset", "CCiniDebug", "")
        printMenuLine(menuStr02)
    End Sub

    ''' <summary>
    ''' Resets the module settings to their defaults
    ''' </summary>
    Private Sub resetSettings()
        initDefaultSettings()
        menuTopper = "CCiniDebug settings have been reset to their defaults"
    End Sub

    ''' <summary>
    ''' The main event loop for the program
    ''' </summary>
    Public Sub main()
        initMenu("CCiniDebug", 35)
        Do Until exitCode
            Console.Clear()
            printMenu()
            Console.WriteLine()
            Console.Write(promptStr)
            handleUserInput(Console.ReadLine)
        Loop
        revertMenu()
    End Sub

    ''' <summary>
    ''' Handles the user's input from the menu
    ''' </summary>
    ''' <param name="input">The string containing the user's input</param>
    Private Sub handleUserInput(input As String)
        Select Case True
            Case input = "0"
                Console.WriteLine("Returning to winapp2ool menu...")
                exitCode = True
            Case (input = "1" Or input = "") And (pruneFile Or saveFile Or sortFile)
                initDebug()
            Case input = "2"
                toggleSettingParam(pruneFile, "Pruning ", settingsChanged)
            Case input = "3"
                toggleSettingParam(saveFile, "Autosaving ", settingsChanged)
            Case input = "4"
                toggleSettingParam(sortFile, "Sorting ", settingsChanged)
            Case input = "5"
                changeFileParams(ccFile, settingsChanged)
            Case input = "6" And pruneFile
                changeFileParams(winappFile, settingsChanged)
            Case (input = "6" And Not pruneFile And saveFile) Or (input = "7" And pruneFile And saveFile)
                changeFileParams(outputFile, settingsChanged)
            Case (input = "6" And Not pruneFile And Not saveFile And settingsChanged) Or (input = "7" And ((pruneFile Or saveFile) And Not (pruneFile And saveFile) And settingsChanged)) Or input = "8" And pruneFile And saveFile And settingsChanged
                resetSettings()
            Case Not (pruneFile Or saveFile Or sortFile)
                menuTopper = "Please enable at least one option"
            Case Else
                menuTopper = invInpStr
        End Select
    End Sub

    ''' <summary>
    ''' Performs the debugging process
    ''' </summary>
    Private Sub initDebug()
        Console.Clear()
        ccFile.validate()
        If exitCode Then Exit Sub
        Console.Clear()

        'Load winapp2.ini and use it to prune ccleaner.ini of stale entries
        If pruneFile Then
            winappFile.validate()
            If exitCode Then Exit Sub
            Console.Clear()
            printMenuLine(tmenu("CCiniDebug Results"))
            printMenuLine(menuStr03)
            printMenuLine("Pruning...", "c")
            printBlankMenuLine()
            prune(ccFile.sections("Options"))
            printBlankMenuLine()
        Else
            printMenuLine(tmenu("CCiniDebug Results"))
            printMenuLine(menuStr03)
        End If
        'Sort the keys
        If sortFile Then
            printMenuLine("Sorting..", "c")
            sortCC()
            printMenuLine("Sorting complete.", "l")
            printBlankMenuLine()
        End If
        'Write ccleaner.ini back to file
        If saveFile Then
            outputFile.overwriteToFile(ccFile.toString)
            printMenuLine(outputFile.name & " saved. " & anyKeyStr, "c")
        Else
            printMenuLine("Analysis complete. " & anyKeyStr, "c")
        End If
        printMenuLine(menuStr02)
        If Not suppressOutput Then Console.ReadKey()
        'Flip the exitCode boolean
        revertMenu()
    End Sub

    ''' <summary>
    ''' Scans for and removes stale winapp2.ini entry settings from the Options section of a ccleaner.ini file
    ''' </summary>
    ''' <param name="optionsSec">The iniSection object containing the Options from ccleaner.ini</param>
    Private Sub prune(ByRef optionsSec As iniSection)
        printMenuLine("Scanning " & ccFile.name & " for settings left over from removed winapp2.ini entries", "l")
        printBlankMenuLine()

        'collect the keys we must remove
        Dim tbTrimmed As New List(Of Integer)

        For i As Integer = 0 To optionsSec.keys.Count - 1
            Dim optionStr As String = optionsSec.keys.Values(i).toString
            'only operate on (app) keys belonging to winapp2.ini
            If optionStr.StartsWith("(App)") And optionStr.Contains("*") Then
                optionStr = optionStr.Replace("(App)", "")
                optionStr = optionStr.Replace("=True", "")
                optionStr = optionStr.Replace("=False", "")
                If Not winappFile.sections.ContainsKey(optionStr) Then
                    printMenuLine(optionsSec.keys.Values(i).lineString, "l")
                    tbTrimmed.Add(optionsSec.keys.Keys(i))
                End If
            End If
        Next
        printBlankMenuLine()
        printMenuLine(tbTrimmed.Count & " orphaned settings detected", "l")
        'Remove the keys
        For Each key In tbTrimmed
            optionsSec.keys.Remove(key)
        Next
    End Sub

    ''' <summary>
    ''' Sorts the keys in the Options (only) section of ccleaner.ini
    ''' </summary>
    Private Sub sortCC()
        Dim lineList As List(Of String) = ccFile.sections("Options").getKeysAsList
        lineList.Insert(0, "[Options]")
        ccFile.sections("Options") = New iniSection(lineList)
    End Sub
End Module