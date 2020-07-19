Option Strict On
'    Copyright (C) 2018-2020 Robbie Ward
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
Imports System.Globalization

''' <summary> Facilitates the merger of one <c> iniFile </c> into another </summary>
Module Merge
    ''' <summary> <c> True </c> if replacing collisions, <c> False </c> if removing them </summary>
    Public Property mergeMode As Boolean = True
    ''' <summary> Indicates that module's settings have been modified from their defaults </summary>
    Public Property ModuleSettingsChanged As Boolean = False

    ''' <summary> The file with whose contents <c> MergeFile2 </c> will be merged </summary>
    Public Property MergeFile1 As New iniFile(Environment.CurrentDirectory, "winapp2.ini")

    ''' <summary> The file whose contents will be merged into <c> MergeFile1 </c> </summary>
    Public Property MergeFile2 As New iniFile(Environment.CurrentDirectory, "")

    ''' <summary>Stores the path to which the merged file should be written back to disk (overwrites <c> MergeFile1 </c> by default) </summary>
    Public Property MergeFile3 As New iniFile(Environment.CurrentDirectory, "winapp2.ini", "winapp2-merged.ini")

    ''' <summary> Handles the commandline args for Merge </summary>
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
        If Not MergeFile2.Name.Length = 0 Then initMerge()
    End Sub

    ''' <summary> Restores the default state of the module's properties </summary>
    Private Sub initDefaultSettings()
        MergeFile1.resetParams()
        MergeFile2.resetParams()
        MergeFile3.resetParams()
        mergeMode = True
        ModuleSettingsChanged = False
        restoreDefaultSettings(NameOf(Merge), AddressOf createMergeSettingsSection)
    End Sub

    ''' <summary> Loads values from disk into memory for the Merge module settings </summary>
    Public Sub getSeralizedMergeSettings()
        For Each kvp In settingsDict(NameOf(Merge))
            Select Case kvp.Key
                Case NameOf(MergeFile1) & "_Name"
                    MergeFile1.Name = kvp.Value
                Case NameOf(MergeFile1) & "_Dir"
                    MergeFile1.Dir = kvp.Value
                Case NameOf(MergeFile2) & "_Name"
                    MergeFile2.Name = kvp.Value
                Case NameOf(MergeFile2) & "_Dir"
                    MergeFile2.Dir = kvp.Value
                Case NameOf(MergeFile3) & "_Name"
                    MergeFile3.Name = kvp.Value
                Case NameOf(MergeFile3) & "_Dir"
                    MergeFile1.Dir = kvp.Value
            End Select
        Next
    End Sub

    '''<summary> Adds the current (typically default) state of the module's settings into the disk-writable settings representation </summary>
    Public Sub createMergeSettingsSection()
        Dim mergeSettingsTuple As New List(Of String) From {NameOf(mergeMode), tsInvariant(mergeMode), NameOf(ModuleSettingsChanged), tsInvariant(ModuleSettingsChanged),
            NameOf(MergeFile1), MergeFile1.Name, MergeFile1.Dir, NameOf(MergeFile2), MergeFile2.Name, MergeFile2.Dir, NameOf(MergeFile3), MergeFile3.Name, MergeFile3.Dir}
        createModuleSettingsSection(NameOf(Merge), mergeSettingsTuple, 2)
    End Sub

    ''' <summary> Prints the <c> Merge </c> menu to the user, includes some predefined merge files choices for ease of access </summary>
    Public Sub printMenu()
        printMenuTop({"Merge the contents of two ini files, while either replacing (default) or removing sections with the same name."})
        print(1, "Run (default)", "Merge the two ini files", enStrCond:=Not MergeFile2.Name.Length = 0, colorLine:=True)
        print(0, "Preset Merge File Choices:", leadingBlank:=True, trailingBlank:=True)
        print(1, "Removed Entries", "Select 'Removed Entries.ini'")
        print(1, "Custom", "Select 'Custom.ini'")
        print(1, "winapp3.ini", "Select 'winapp3.ini'", trailingBlank:=True)
        print(1, "File Chooser (winapp2.ini)", "Choose a new name or location for winapp2.ini")
        print(1, "File Chooser (merge)", "Choose a name or location for merging")
        print(1, "File Chooser (save)", "Choose a new save location for the merged file", trailingBlank:=True)
        print(0, $"Current winapp2.ini: {replDir(MergeFile1.Path)}")
        print(0, $"Current merge file : {If(MergeFile2.Name.Length = 0, "Not yet selected", replDir(MergeFile2.Path))}", enStrCond:=Not MergeFile2.Name.Length = 0, colorLine:=True)
        print(0, $"Current save target: {replDir(MergeFile3.Path)}", trailingBlank:=True)
        print(1, "Toggle Merge Mode", "Switch between merge modes.")
        print(0, $"Current mode: {If(mergeMode, "Add & Replace", "Add & Remove")}", closeMenu:=Not ModuleSettingsChanged)
        print(2, NameOf(Merge), cond:=ModuleSettingsChanged, closeMenu:=True)
        Console.WindowHeight = If(ModuleSettingsChanged, 32, 30)
    End Sub

    ''' <summary> Handles the user's input from the main menu </summary>
    ''' <param name="input"> The user's input </param>
    Public Sub handleUserInput(input As String)
        Select Case True
            Case input = "0"
                exitModule()
            Case input = "1" Or input.Length = 0
                If Not denyActionWithHeader(MergeFile2.Name.Length = 0, "You must select a file to merge") Then initMerge()
            Case input = "2"
                changeMergeName("Removed entries.ini")
            Case input = "3"
                changeMergeName("Custom.ini")
            Case input = "4"
                changeMergeName("winapp3.ini")
            Case input = "5"
                changeFileParams(MergeFile1, ModuleSettingsChanged, NameOf(Merge), NameOf(MergeFile1), NameOf(ModuleSettingsChanged))
            Case input = "6"
                changeFileParams(MergeFile2, ModuleSettingsChanged, NameOf(Merge), NameOf(MergeFile2), NameOf(ModuleSettingsChanged))
            Case input = "7"
                changeFileParams(MergeFile3, ModuleSettingsChanged, NameOf(Merge), NameOf(MergeFile3), NameOf(ModuleSettingsChanged))
            Case input = "8"
                toggleSettingParam(mergeMode, "Merge Mode ", ModuleSettingsChanged, NameOf(Merge), NameOf(mergeMode), NameOf(ModuleSettingsChanged))
            Case input = "9" And ModuleSettingsChanged
                resetModuleSettings("Merge", AddressOf initDefaultSettings)
            Case Else
                setHeaderText(invInpStr, True)
        End Select
    End Sub

    ''' <summary> Changes the merge file's name </summary>
    ''' <param name="newName"> The new <c> Name </c> for <c> MergeFile2 </c> </param>
    Private Sub changeMergeName(newName As String)
        MergeFile2.Name = newName
        ModuleSettingsChanged = True
        setHeaderText("Merge filename set")
    End Sub

    ''' <summary> Validates the <c> iniFiles </c> and kicks off the merging process </summary>
    Public Sub initMerge()
        clrConsole()
        If Not (enforceFileHasContent(MergeFile1) And enforceFileHasContent(MergeFile2)) Then Return
        print(4, $"Merging {MergeFile1.Name} with {MergeFile2.Name}")
        mergeFiles(True)
        print(0, "", closeMenu:=True)
        print(3, $"Finished merging files. {anyKeyStr}")
        crk()
    End Sub

    ''' <summary> Conducts the merger of our two iniFiles </summary>
    ''' <param name="isWinapp2"> Indicates that the <c> iniFiles </c> being operated on are of winapp2.ini syntax </param>
    Private Sub mergeFiles(isWinapp2 As Boolean)
        resolveConflicts(MergeFile1, MergeFile2)
        ' After conflicts are resolved, remaining entries need only be added 
        For Each section In MergeFile2.Sections.Values
            MergeFile1.Sections.Add(section.Name, section)
            print(0, $"Adding {section.Name}", colorLine:=True, enStrCond:=True)
        Next
        If isWinapp2 Then
            Dim tmp As New winapp2file(MergeFile1)
            tmp.sortInneriniFiles()
            MergeFile1.Sections = tmp.toIni.Sections
            MergeFile3.overwriteToFile(tmp.winapp2string)
        Else
            MergeFile1.sortSections(MergeFile1.namesToStrList)
            MergeFile3.overwriteToFile(MergeFile1.toString)
        End If
    End Sub

    ''' <summary> Facilitates merging <c> iniFiles </c> from outside the module </summary>
    ''' <param name="mergeFile"> An <c> iniFile </c> to whom content will be added </param>
    ''' <param name="sourceFile">An <c> iniFile </c> whose content will be added to <c> <paramref name="mergeFile"/> </c> </param>
    ''' <param name="isWinapp"> Indicates that the <c> iniFiles </c> being worked with contain winapp2.ini syntax </param>
    ''' <param name="mm"> Indicates the <c> MergeMode </c> for the merger <br/> 
    ''' <c> True </c> to replace conflicts, <c> False </c> to remove them </param>
    Public Sub RemoteMerge(ByRef mergeFile As iniFile, ByRef sourceFile As iniFile, isWinapp As Boolean, mm As Boolean)
        MergeFile1 = mergeFile
        MergeFile2 = sourceFile
        mergeMode = mm
        mergeFiles(isWinapp)
    End Sub

    ''' <summary> Performs conflict resolution for the merge process, handling the case where <c> iniSections </c> 
    ''' in both files have the same <c> Name </c> </summary>
    ''' <param name="first"> The master <c> iniFile </c> </param>
    ''' <param name="second"> The <c> iniFile </c> whose contents will be merged into <c> <paramref name="first"/> </c> </param>
    Private Sub resolveConflicts(ByRef first As iniFile, ByRef second As iniFile)
        Dim removeList As New List(Of String)
        For Each section In second.Sections.Keys
            If first.Sections.Keys.Contains(section) Then
                print(0, $"{If(mergeMode, "Replacing", "Removing")} {first.Sections.Item(section).Name}", colorLine:=True, enStrCond:=mergeMode)
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