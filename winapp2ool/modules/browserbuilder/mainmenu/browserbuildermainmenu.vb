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

''' <summary>
''' Displays the Browser Builder main menu and handles user input, allowing user configuration
''' of module settings from the UI
''' </summary>
''' 
''' Docs last updated: 2025-07-30 | Code last updated: 2025-07-30
Public Module browserbuildermainmenu

    ''' <summary>
    ''' Prints the BrowserBuilder main menu to the user
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-07-02 | Code last updated: 2025-07-02
    Public Sub printBrowserBuilderMenu()

        printMenuTop({"Generate winapp2.ini entries for web browsers and web views",
                     "Supports both Chromium-based and Gecko-based entry generation",
                     "Consult the winapp2ool ReadMe before using this module!"})

        print(1, "Run (Default)", "Run the browser builder")
        print(1, "File Chooser (chromium.ini)", "Configure the path to the chromium browser rules", leadingBlank:=True, trailingBlank:=True)
        print(1, "File Chooser (gecko.ini)", "Configure the path to the geck browser rules", trailingBlank:=True)
        print(1, "File Chooser (save)", "Configure the path for the output file", trailingBlank:=True)
        print(1, "File Chooser (additions)", "Configure the path to the file containing any additional entry code", trailingBlank:=True)
        print(0, $"Current chromium.ini path: {replDir(BuilderFile1.Path)}")
        print(0, $"Current gecko.ini path: {replDir(builderfile2.Path)}")
        print(0, $"Current save path: {replDir(builderfile3.Path)}", closeMenu:=Not BrowserBuilderModuleSettingsChanged)
        print(2, NameOf(BrowserBuilder), cond:=BrowserBuilderModuleSettingsChanged, closeMenu:=True)

    End Sub

    ''' <summary>
    ''' Handles the user input from the BrowserBuilder menu
    ''' </summary>
    ''' 
    ''' <param name="input">The String containing the user's input</param>
    ''' 
    ''' Docs last updated: 2025-07-02 | Code last updated: 2025-07-02
    Public Sub handleBrowserBuilderInput(input As String)

        If input Is Nothing Then argIsNull(NameOf(input)) : Return

        Select Case True

            ' Option Name:                                 Exit 
            ' Option States:
            ' Default                                      -> 0 (Default) 
            Case input = "0"

                exitModule()

            ' Option Name:                                 Run (default) 
            ' Option States:
            ' Default                                      -> 1 (Default)
            Case (input = "1" OrElse input.Length = 0)

                initBrowserBuilder()

            ' Option Name:                                 File Chooser (chromium.ini) 
            ' Option States:
            ' Default                                      -> 2 (Default)
            Case input = "2"

                changeFileParams(BuilderFile1, BrowserBuilderModuleSettingsChanged, NameOf(BrowserBuilder), NameOf(BuilderFile1), NameOf(BrowserBuilderModuleSettingsChanged))

            ' Option Name:                                 File Chooser (gecko.ini) 
            ' Option States:
            ' Default                                      -> 3 (Default)
            Case input = "3"

                changeFileParams(builderfile2, BrowserBuilderModuleSettingsChanged, NameOf(BrowserBuilder), NameOf(builderfile2), NameOf(BrowserBuilderModuleSettingsChanged))

            ' Option Name:                                 File Chooser (browsers.ini) 
            ' Option States:
            ' Default                                      -> 4 (Default)
            Case input = "4"

                changeFileParams(builderfile3, BrowserBuilderModuleSettingsChanged, NameOf(BrowserBuilder), NameOf(builderfile3), NameOf(BrowserBuilderModuleSettingsChanged))

            ' Option Name:                                 Reset Settings       
            ' Option States:
            ' Module Settings not changed                  -> Unavailable (not displayed)
            ' Module Settings changed                      -> 5 (Default)
            Case BrowserBuilderModuleSettingsChanged AndAlso input = "5"

                resetModuleSettings(NameOf(BrowserBuilder), AddressOf initDefaultBrowserBuilderSettings)

            Case Else

                setHeaderText(invInpStr, True)

        End Select

    End Sub

End Module
