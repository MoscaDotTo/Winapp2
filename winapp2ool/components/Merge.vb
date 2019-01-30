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
''' A module that facilitates the merging of two user defined iniFile objects
''' </summary>
Module Merge
    ' File handlers
    Dim winappFile As iniFile = New iniFile(Environment.CurrentDirectory, "winapp2.ini")
    Dim mergeFile As iniFile = New iniFile(Environment.CurrentDirectory, "")
    Dim outputFile As iniFile = New iniFile(Environment.CurrentDirectory, "winapp2.ini", "winapp2-merged.ini")
    ' Module parameters
    Dim mergeMode As Boolean = True
    Dim settingsChanged As Boolean = False

    ''' <summary>
    ''' Handles the commandline args for Merge
    ''' </summary>
    '''  Merge args:
    ''' -mm         : toggle mergemode from add&amp;replace to add&amp;remove
    ''' Preset merge file choices
    ''' -r          : removed entries.ini 
    ''' -c          : custom.ini 
    ''' -w          : winapp3.ini
    ''' -a          : archived entries.ini 
    Public Sub handleCmdLine()
        initDefaultSettings()
        invertSettingAndRemoveArg(mergeMode, "-mm")
        invertSettingAndRemoveArg(False, "-r", mergeFile.name, "Removed Entries.ini")
        invertSettingAndRemoveArg(False, "-c", mergeFile.name, "Custom.ini")
        invertSettingAndRemoveArg(False, "-w", mergeFile.name, "winapp3.ini")
        invertSettingAndRemoveArg(False, "-a", mergeFile.name, "Archived Entries.ini")
        getFileAndDirParams(winappFile, mergeFile, outputFile)
        If Not mergeFile.name = "" Then initMerge()
    End Sub

    ''' <summary>
    ''' Restores the default state of the module's parameters
    ''' </summary>
    Private Sub initDefaultSettings()
        winappFile.resetParams()
        mergeFile.resetParams()
        outputFile.resetParams()
        settingsChanged = False
    End Sub

    ''' <summary>
    ''' Prints the main menu to the user
    ''' </summary>
    Public Sub printMenu()
        printMenuTop({"Merge the contents of two ini files, while either replacing (default) or removing sections with the same name."})
        print(1, "Run (default)", "Merge the two ini files")
        print(0, "Preset Merge File Choices:", leadingBlank:=True, trailingBlank:=True)
        print(1, "Removed Entries", "Select 'Removed Entries.ini'")
        print(1, "Custom", "Select 'Custom.ini'")
        print(1, "winapp3.ini", "Select 'winapp3.ini'", trailingBlank:=True)
        print(1, "File Chooser (winapp2.ini)", "Choose a new name or location for winapp2.ini")
        print(1, "File Chooser (merge)", "Choose a name or location for merging")
        print(1, "File Chooser (save)", "Choose a new save location for the merged file", trailingBlank:=True)
        print(0, $"Current winapp2.ini: {replDir(winappFile.path)}")
        print(0, $"Current merge file : {If(mergeFile.name = "", "Not yet selected", replDir(mergeFile.path))}")
        print(0, $"Current save target: {replDir(outputFile.path)}", trailingBlank:=True)
        print(1, "Toggle Merge Mode", "Switch between merge modes.")
        print(0, $"Current mode: {If(mergeMode, "Add & Replace", "Add & Remove")}", closeMenu:=Not settingsChanged)
        print(2, "Merge", cond:=settingsChanged, closeMenu:=True)
        Console.WindowHeight = If(settingsChanged, 32, 30)
    End Sub

    ''' <summary>
    ''' Processes the user's input and acts accordingly based on the state of the program
    ''' </summary>
    ''' <param name="input">The String containing the user's input</param>
    Public Sub handleUserInput(input As String)
        Select Case True
            Case input = "0"
                exitModule()
            Case input = "1" Or input = ""
                If Not denyActionWithTopper(mergeFile.name = "", "You must select a file to merge") Then initMerge()
            Case input = "2"
                changeMergeName("Removed entries.ini")
            Case input = "3"
                changeMergeName("Custom.ini")
            Case input = "4"
                changeMergeName("winapp3.ini")
            Case input = "5"
                changeFileParams(winappFile, settingsChanged)
            Case input = "6"
                changeFileParams(mergeFile, settingsChanged)
            Case input = "7"
                changeFileParams(outputFile, settingsChanged)
            Case input = "8"
                toggleSettingParam(mergeMode, "Merge Mode ", settingsChanged)
            Case input = "9" And settingsChanged
                resetModuleSettings("Merge", AddressOf initDefaultSettings)
            Case Else
                menuHeaderText = invInpStr
        End Select
    End Sub

    ''' <summary>
    ''' Changes the merge file's name
    ''' </summary>
    ''' <param name="newName">the new name for the merge file</param>
    Private Sub changeMergeName(newName As String)
        mergeFile.name = newName
        settingsChanged = True
        menuHeaderText = "Merge filename set"
    End Sub

    ''' <summary>
    ''' Validates iniFiles and begins the merging process
    ''' </summary>
    Public Sub initMerge()
        Console.Clear()
        winappFile.validate()
        mergeFile.validate()
        If pendingExit() Then Exit Sub
        merge()
    End Sub

    ''' <summary>
    ''' Conducts the merger of our two iniFiles
    ''' </summary>
    Private Sub merge()
        processMergeMode(winappFile, mergeFile)
        Dim tmp As New winapp2file(winappFile)
        Dim tmp2 As New winapp2file(mergeFile)
        print(0, bmenu($"Merging {winappFile.name} with {mergeFile.name}"))
        ' Add the entries from the second file to their respective sections in the first file
        For i As Integer = 0 To tmp.winapp2entries.Count - 1
            tmp.winapp2entries(i).AddRange(tmp2.winapp2entries(i))
        Next
        tmp.rebuildToIniFiles()
        For Each section In tmp.entrySections
            section.sortSections(replaceAndSort(section.getSectionNamesAsList, "-", "  "))
        Next
        Dim out As String = tmp.winapp2string
        outputFile.overwriteToFile(out)
        print(0, bmenu($"Finished merging files. {anyKeyStr}"))
        If Not suppressOutput Then Console.ReadKey()
    End Sub

    ''' <summary>
    ''' Performs conflict resolution for the merge process
    ''' </summary>
    ''' <param name="first">The base iniFile that will be modified by Merge</param>
    ''' <param name="second">The iniFile to be merged into the base</param>
    Private Sub processMergeMode(ByRef first As iniFile, ByRef second As iniFile)
        Dim removeList As New List(Of String)
        For Each section In second.sections.Keys
            If first.sections.Keys.Contains(section) Then
                ' If mergemode is true, replace the match. otherwise, remove the match
                If mergeMode Then first.sections.Item(section) = second.sections.Item(section) Else first.sections.Remove(section)
                removeList.Add(section)
            End If
        Next
        ' Remove any processed sections from the second file so that only entries to add remain 
        For Each section In removeList
            second.sections.Remove(section)
        Next
    End Sub
End Module