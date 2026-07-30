'    Copyright (C) 2018-2026 Hazel Ward
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

''' <summary>
''' Displays the EntryBuilder main menu and handles user input, allowing configuration
''' of module settings from the UI
''' </summary>
Public Module entryBuilderMainMenu

    ''' <summary>
    ''' Builds the EntryBuilder main menu with all options and their dispatch handlers
    ''' registered inline
    ''' </summary>
    '''
    ''' <returns>
    ''' A fully configured <c> MenuSection </c> ready to print or dispatch
    ''' </returns>
    Private Function buildEntryBuilderMenu() As MenuSection

        Dim menuDesc = {"Generate winapp2.ini entries from a shorthand DSL (pass-through + WebView scaffold expansion)",
                        "Consult the winapp2ool ReadMe before using this module!"}

        Return MenuSection.CreateCompleteMenu(NameOf(EntryBuilder), menuDesc, ConsoleColor.DarkCyan) _
            .AddDispatchedOption("Run (default)", "Generate winapp2.ini entries from the shorthand source files",
                Sub() initEntryBuilder()) _
            .AddBlank() _
            .AddDispatchedOption("Choose source directory", "Select the directory containing per-letter shorthand source files",
                Sub() changeFile2Params(EntryBuilderFile1, EntryBuilderModuleSettingsChanged, NameOf(EntryBuilder), NameOf(EntryBuilderFile1), NameOf(EntryBuilderModuleSettingsChanged), "Source directory")) _
            .AddDispatchedOption("Choose save target", "Select where to save the generated entries",
                Sub() changeFile2Params(EntryBuilderFile2, EntryBuilderModuleSettingsChanged, NameOf(EntryBuilder), NameOf(EntryBuilderFile2), NameOf(EntryBuilderModuleSettingsChanged), "Save target")) _
            .AddDispatchedOption("Choose scaffolds directory", "Select the shared scaffold catalog directory",
                Sub() changeFile2Params(EntryBuilderFile3, EntryBuilderModuleSettingsChanged, NameOf(EntryBuilder), NameOf(EntryBuilderFile3), NameOf(EntryBuilderModuleSettingsChanged), "Scaffolds directory")) _
            .AddBlank() _
            .AddDispatchedToggle("output splitting", "writing per-letter files into the save target's directory", EntryBuilderSplitOutput,
                Sub() toggleModuleSetting("Split output", NameOf(EntryBuilder), GetType(entryBuilderSettings), NameOf(EntryBuilderSplitOutput), NameOf(EntryBuilderModuleSettingsChanged))) _
            .AddBlank() _
            .AddColoredFileInfo("Current source directory:            ", EntryBuilderFile1.Dir, ConsoleColor.DarkYellow) _
            .AddColoredFileInfo("Current save target:                 ", EntryBuilderFile2.Path(), ConsoleColor.Yellow) _
            .AddColoredFileInfo("Current scaffolds directory:         ", EntryBuilderFile3.Dir, ConsoleColor.DarkYellow) _
            .AddBlank(EntryBuilderModuleSettingsChanged) _
            .AddDispatchedResetOpt(NameOf(EntryBuilder), EntryBuilderModuleSettingsChanged,
                Sub() resetModuleSettings(NameOf(EntryBuilder), AddressOf InitDefaultEntryBuilderSettings))

    End Function

    ''' <summary>
    ''' Prints the EntryBuilder main menu to the user
    ''' </summary>
    Public Sub printEntryBuilderMenu()

        buildEntryBuilderMenu().Print()

    End Sub

    ''' <summary>
    ''' Handles the user input from the EntryBuilder menu
    ''' </summary>
    '''
    ''' <param name="input">
    ''' The user's input
    ''' </param>
    Public Sub handleEntryBuilderInput(input As String)

        Dim intInput As Integer

        If Not Integer.TryParse(input, intInput) Then

            If input.Length = 0 Then initEntryBuilder() : Return

            setNextMenuHeaderText(invInpStr, printColor:=ConsoleColor.Red)
            Return

        End If

        If intInput = 0 Then exitModule() : Return

        If Not buildEntryBuilderMenu().Dispatch(intInput) Then setNextMenuHeaderText(invInpStr, printColor:=ConsoleColor.Red)

    End Sub

End Module
