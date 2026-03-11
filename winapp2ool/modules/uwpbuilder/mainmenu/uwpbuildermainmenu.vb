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
''' Displays the UWPBuilder main menu and handles user input, allowing configuration
''' of module settings from the UI
''' </summary>
Public Module uwpbuildermainmenu

    ''' <summary>
    ''' Builds the UWPBuilder main menu with all options and their dispatch handlers registered inline
    ''' </summary>
    '''
    ''' <returns>
    ''' A fully configured <c> MenuSection </c> ready to print or dispatch
    ''' </returns>
    Private Function buildUWPBuilderMenu() As MenuSection

        Dim menuDesc = {"Generate winapp2.ini entries for Universal Windows Platform applications",
                        "Consult the winapp2ool ReadMe before using this module!"}

        Return MenuSection.CreateCompleteMenu(NameOf(UWPBuilder), menuDesc, ConsoleColor.DarkCyan) _
            .AddDispatchedOption("Run (default)", "Generate UWP app winapp2.ini entries",
                Sub() initUWPBuilder()) _
            .AddBlank() _
            .AddDispatchedOption("Choose source directory", "Select the directory containing UWP.ini and AppInfo\",
                Sub() changeFile2Params(UWPFile1, UWPBuilderModuleSettingsChanged, NameOf(UWPBuilder), NameOf(UWPFile1), NameOf(UWPBuilderModuleSettingsChanged), "Source directory")) _
            .AddDispatchedOption("Choose save target", "Select where to save the generated entries",
                Sub() changeFile2Params(UWPFile2, UWPBuilderModuleSettingsChanged, NameOf(UWPBuilder), NameOf(UWPFile2), NameOf(UWPBuilderModuleSettingsChanged), "Save target")) _
            .AddBlank() _
            .AddColoredFileInfo("Current source directory: ", UWPFile1.Dir, ConsoleColor.DarkYellow) _
            .AddColoredFileInfo("Current save target:      ", UWPFile2.Path(), ConsoleColor.Yellow) _
            .AddBlank(UWPBuilderModuleSettingsChanged) _
            .AddDispatchedResetOpt(NameOf(UWPBuilder), UWPBuilderModuleSettingsChanged,
                Sub() resetModuleSettings(NameOf(UWPBuilder), AddressOf InitDefaultUWPBuilderSettings))

    End Function

    ''' <summary>
    ''' Prints the UWPBuilder main menu to the user
    ''' </summary>
    Public Sub printUWPBuilderMenu()

        buildUWPBuilderMenu().Print()

    End Sub

    ''' <summary>
    ''' Handles the user input from the UWPBuilder menu
    ''' </summary>
    '''
    ''' <param name="input">
    ''' The user's input
    ''' </param>
    Public Sub handleUWPBuilderInput(input As String)

        Dim intInput As Integer

        If Not Integer.TryParse(input, intInput) Then

            If input.Length = 0 Then initUWPBuilder() : Return

            setNextMenuHeaderText(invInpStr, printColor:=ConsoleColor.Red)
            Return

        End If

        If intInput = 0 Then exitModule() : Return

        If Not buildUWPBuilderMenu().Dispatch(intInput) Then setNextMenuHeaderText(invInpStr, printColor:=ConsoleColor.Red)

    End Sub

End Module
