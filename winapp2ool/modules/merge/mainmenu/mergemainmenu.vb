'    Copyright (C) 2018-2025 Hazel Ward
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
Module mergemainmenu

    ''' <summary> Prints the <c> Merge </c> menu to the user, includes some predefined merge files choices for ease of access </summary>
    Public Sub printMergeMainMenu()

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
        print(0, $"Current mode: {If(mergeMode, "Add & Replace", "Add & Remove")}", closeMenu:=Not MergeModuleSettingsChanged)
        print(2, NameOf(Merge), cond:=MergeModuleSettingsChanged, closeMenu:=True)
        Console.WindowHeight = If(MergeModuleSettingsChanged, 32, 30)

    End Sub

    ''' <summary> Handles the user's input from the main menu </summary>
    ''' <param name="input"> The user's input </param>
    Public Sub handleMergeMainMenuUserInput(input As String)
        Select Case True
            Case input = "0"
                exitModule()
            Case input = "1" OrElse input.Length = 0
                If Not denyActionWithHeader(MergeFile2.Name.Length = 0, "You must select a file to merge") Then initMerge()
            Case input = "2"
                changeMergeName("Removed entries.ini")
            Case input = "3"
                changeMergeName("Custom.ini")
            Case input = "4"
                changeMergeName("winapp3.ini")
            Case input = "5"
                changeFileParams(MergeFile1, MergeModuleSettingsChanged, NameOf(Merge), NameOf(MergeFile1), NameOf(MergeModuleSettingsChanged))
            Case input = "6"
                changeFileParams(MergeFile2, MergeModuleSettingsChanged, NameOf(Merge), NameOf(MergeFile2), NameOf(MergeModuleSettingsChanged))
            Case input = "7"
                changeFileParams(MergeFile3, MergeModuleSettingsChanged, NameOf(Merge), NameOf(MergeFile3), NameOf(MergeModuleSettingsChanged))
            Case input = "8"
                toggleSettingParam(mergeMode, "Merge Mode ", MergeModuleSettingsChanged, NameOf(Merge), NameOf(mergeMode), NameOf(MergeModuleSettingsChanged))
            Case input = "9" And MergeModuleSettingsChanged
                resetModuleSettings("Merge", AddressOf initDefaultMergeSettings)
            Case Else
                setHeaderText(invInpStr, True)
        End Select
    End Sub

End Module
