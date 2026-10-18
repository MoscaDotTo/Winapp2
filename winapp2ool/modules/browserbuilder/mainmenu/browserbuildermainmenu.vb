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
''' Displays the Browser Builder main menu and handles user input, allowing user configuration
''' of module settings from the UI
''' </summary>
Public Module browserbuildermainmenu

    ''' <summary>
    ''' Builds the BrowserBuilder main menu with all options and their dispatch handlers registered inline
    ''' </summary>
    '''
    ''' <returns>
    ''' A fully configured <c> MenuSection </c> ready to print or dispatch
    ''' </returns>
    Friend Function buildBrowserBuilderMenu() As MenuSection

        Dim menuDesc = {"Generate winapp2.ini entries for web browsers and web views",
                        "Supports both Chromium-based and Gecko-based entry generation",
                        "Consult the winapp2ool ReadMe before using this module!"}

        Return MenuSection.CreateCompleteMenu(NameOf(BrowserBuilder), menuDesc, ConsoleColor.DarkMagenta) _
            .AddDispatchedOption("Run (default)", "Generate web browser winapp2.ini entries",
                Sub() initBrowserBuilder(), isDefault:=True) _
            .AddBlank() _
            .AddDispatchedOption("Choose source directory", "Select the directory containing chromium.ini, gecko.ini, and flavor correction files",
                Sub() changeFileParams(BuilderFile1, BrowserBuilderModuleSettingsChanged, NameOf(BrowserBuilder), NameOf(BuilderFile1), NameOf(BrowserBuilderModuleSettingsChanged), "Source directory")) _
            .AddDispatchedOption("Choose save target", "Select a new location on disk to which generated entries should be saved",
                Sub() changeFileParams(BuilderFile2, BrowserBuilderModuleSettingsChanged, NameOf(BrowserBuilder), NameOf(BuilderFile2), NameOf(BrowserBuilderModuleSettingsChanged), "Save target")) _
            .AddBlank() _
            .AddColoredFileInfo("Current source directory: ", BuilderFile1.Dir, ConsoleColor.DarkYellow) _
            .AddColoredFileInfo("Current save target:      ", BuilderFile2.Path(), ConsoleColor.Yellow) _
            .AddBlank(BrowserBuilderModuleSettingsChanged) _
            .AddDispatchedResetOpt(NameOf(BrowserBuilder), BrowserBuilderModuleSettingsChanged,
                Sub() resetModuleSettings(NameOf(BrowserBuilder), AddressOf InitDefaultBrowserBuilderSettings))

    End Function
End Module
