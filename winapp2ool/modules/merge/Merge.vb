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
''' A module that facilitates the merging of two user defined iniFile objects
''' </summary>
Module Merge
    '''<summary>True if replacing collisions, False if removing them</summary>
    Public Property mergeMode As Boolean = True
    '''<summary>Indicates whether or not the module's settings have changed</summary>
    Public Property ModuleSettingsChanged As Boolean = False

    '''<summary>The master file to be merged into</summary>
    Public Property MergeFile1 As New iniFile(Environment.CurrentDirectory, "winapp2.ini")

    '''<summary> The file whose contents will be merged into MergeFile1</summary>
    Public Property MergeFile2 As New iniFile(Environment.CurrentDirectory, "")

    ''' <summary>Stores the path for where the merged file should be saved to disk (overwrites MergeFile1 by default)</summary>
    Public Property MergeFile3 As New iniFile(Environment.CurrentDirectory, "winapp2.ini", "winapp2-merged.ini")

    ''' <summary>Handles the commandline args for Merge</summary>
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
        invertSettingAndRemoveArg(False, "-r", MergeFile2.Name, "Removed Entries.ini")
        invertSettingAndRemoveArg(False, "-c", MergeFile2.Name, "Custom.ini")
        invertSettingAndRemoveArg(False, "-w", MergeFile2.Name, "winapp3.ini")
        invertSettingAndRemoveArg(False, "-a", MergeFile2.Name, "Archived Entries.ini")
        getFileAndDirParams(MergeFile1, MergeFile2, MergeFile3)
        If Not MergeFile2.Name = "" Then initMerge()
    End Sub

    ''' <summary>Restores the default state of the module's parameters</summary>
    Private Sub initDefaultSettings()
        MergeFile1.resetParams()
        MergeFile2.resetParams()
        MergeFile3.resetParams()
        mergeMode = True
        ModuleSettingsChanged = False
    End Sub

    ''' <summary>Prints the main menu to the user</summary>
    Public Sub printMenu()
        printMenuTop({"Merge the contents of two ini files, while either replacing (default) or removing sections with the same name."})
        print(1, "Run (default)", "Merge the two ini files", enStrCond:=Not (MergeFile2.Name = ""), colorLine:=True)
        print(0, "Preset Merge File Choices:", leadingBlank:=True, trailingBlank:=True)
        print(1, "Removed Entries", "Select 'Removed Entries.ini'")
        print(1, "Custom", "Select 'Custom.ini'")
        print(1, "winapp3.ini", "Select 'winapp3.ini'", trailingBlank:=True)
        print(1, "File Chooser (winapp2.ini)", "Choose a new name or location for winapp2.ini")
        print(1, "File Chooser (merge)", "Choose a name or location for merging")
        print(1, "File Chooser (save)", "Choose a new save location for the merged file", trailingBlank:=True)
        print(0, $"Current winapp2.ini: {replDir(MergeFile1.Path)}")
        print(0, $"Current merge file : {If(MergeFile2.Name = "", "Not yet selected", replDir(MergeFile2.Path))}", enStrCond:=Not MergeFile2.Name = "", colorLine:=True)
        print(0, $"Current save target: {replDir(MergeFile3.Path)}", trailingBlank:=True)
        print(1, "Toggle Merge Mode", "Switch between merge modes.")
        print(0, $"Current mode: {If(mergeMode, "Add & Replace", "Add & Remove")}", closeMenu:=Not ModuleSettingsChanged)
        print(2, "Merge", cond:=ModuleSettingsChanged, closeMenu:=True)
        Console.WindowHeight = If(ModuleSettingsChanged, 32, 30)
    End Sub

    ''' <summary>Processes the user's input and acts accordingly based on the state of the program</summary>
    ''' <param name="input">The String containing the user's input</param>
    Public Sub handleUserInput(input As String)
        Select Case True
            Case input = "0"
                exitModule()
            Case input = "1" Or input = ""
                If Not denyActionWithTopper(MergeFile2.Name = "", "You must select a file to merge") Then initMerge()
            Case input = "2"
                changeMergeName("Removed entries.ini")
            Case input = "3"
                changeMergeName("Custom.ini")
            Case input = "4"
                changeMergeName("winapp3.ini")
            Case input = "5"
                changeFileParams(MergeFile1, ModuleSettingsChanged)
            Case input = "6"
                changeFileParams(MergeFile2, ModuleSettingsChanged)
            Case input = "7"
                changeFileParams(MergeFile3, ModuleSettingsChanged)
            Case input = "8"
                toggleSettingParam(mergeMode, "Merge Mode ", ModuleSettingsChanged)
            Case input = "9" And ModuleSettingsChanged
                resetModuleSettings("Merge", AddressOf initDefaultSettings)
            Case Else
                setHeaderText(invInpStr, True)
        End Select
    End Sub

    ''' <summary>Changes the merge file's name</summary>
    ''' <param name="newName">the new name for the merge file</param>
    Private Sub changeMergeName(newName As String)
        MergeFile2.Name = newName
        ModuleSettingsChanged = True
        setHeaderText("Merge filename set")
    End Sub

    ''' <summary>Validates iniFiles and begins the merging process</summary>
    Public Sub initMerge()
        clrConsole()
        MergeFile1.validate()
        MergeFile2.validate()
        If Not (enforceFileHasContent(MergeFile1) And enforceFileHasContent(MergeFile2)) Then Exit Sub
        print(0, tmenu($"Merging {MergeFile1.Name} with {MergeFile2.Name}"))
        merge()
        print(0, "", closeMenu:=True)
        print(3, $"Finished merging files. {anyKeyStr}")
        crk()
    End Sub

    ''' <summary>Conducts the merger of our two iniFiles</summary>
    Private Sub merge()
        processMergeMode(MergeFile1, MergeFile2)
        Dim tmp As New winapp2file(MergeFile1)
        Dim tmp2 As New winapp2file(MergeFile2)
        ' Add the entries from the second file to their respective sections in the first file
        Console.ForegroundColor = ConsoleColor.Green
        For i As Integer = 0 To tmp.Winapp2entries.Count - 1
            tmp.Winapp2entries(i).AddRange(tmp2.Winapp2entries(i))
            tmp2.Winapp2entries(i).ForEach(Sub(ent As winapp2entry) print(0, $"Adding {ent.Name}"))
        Next
        Console.ResetColor()
        tmp.rebuildToIniFiles()
        For Each section In tmp.EntrySections
            section.sortSections(replaceAndSort(section.namesToStrList, "-", "  "))
        Next
        Dim out As String = tmp.winapp2string
        MergeFile3.overwriteToFile(out)
    End Sub

    '''<summary>Allows other modules to call Merge on iniFile objects</summary>
    Public Sub merge(ByRef mergeFile As iniFile, ByRef sourceFile As iniFile)
        MergeFile1 = mergeFile
        MergeFile2 = sourceFile
    End Sub

    ''' <summary>Performs conflict resolution for the merge process</summary>
    ''' <param name="first">The base iniFile that will be modified by Merge</param>
    ''' <param name="second">The iniFile to be merged into the base</param>
    Private Sub processMergeMode(ByRef first As iniFile, ByRef second As iniFile)
        Dim removeList As New List(Of String)
        Console.ForegroundColor = If(mergeMode, ConsoleColor.Green, ConsoleColor.Red)
        For Each section In second.Sections.Keys
            If first.Sections.Keys.Contains(section) Then
                print(0, $"{If(mergeMode, "Replacing", "Removing")} {first.Sections.Item(section).Name}")
                ' If mergemode is true, replace the match. otherwise, remove the match
                If mergeMode Then first.Sections.Item(section) = second.Sections.Item(section) Else first.Sections.Remove(section)
                removeList.Add(section)
            End If
        Next
        ' Remove any processed sections from the second file so that only entries to add remain 
        For Each section In removeList
            second.Sections.Remove(section)
        Next
    End Sub
End Module