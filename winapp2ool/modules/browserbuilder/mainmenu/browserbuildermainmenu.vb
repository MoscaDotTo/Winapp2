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
''' Docs last updated: 2025-07-30 
Public Module browserbuildermainmenu

    ''' <summary>
    ''' Prints the BrowserBuilder main menu to the user
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-07-02 | Code last updated: 2025-08-22
    Public Sub printBrowserBuilderMenu()

        Console.WindowHeight = 40
        Console.WindowWidth = 130

        Dim menuDesc = {"Generate winapp2.ini entries for web browsers and web views",
                        "Supports both Chromium-based and Gecko-based entry generation",
                        "Consult the winapp2ool ReadMe before using this module!"}

        Dim menu = MenuSection.CreateCompleteMenu(NameOf(BrowserBuilder), menuDesc, ConsoleColor.DarkMagenta)

        menu.AddOption("Run (default)", "Generate web browser winapp2.ini entries").AddBlank() _
            .AddOption("Choose chromium ruleset", "Select a new generative ruleset for chromium browsers") _
            .AddOption("Choose gecko ruleset", "Select a new generative ruleset for gecko browsers").AddBlank _
            .AddOption("Choose save target", "Select a new location on disk to which generated entries should be saved").AddBlank _
            .AddOption("Choose section removals", "Select a new flavor file for section removals") _
            .AddOption("Choose key name removals", "Select a new flavor file for key removals by name") _
            .AddOption("Choose key value removals", "Select a new flavor file for key removals by value") _
            .AddOption("Choose section replacements", "Select a new flavor file for section replacements") _
            .AddOption("Choose key replacements", "Select a new flavor file for key value replacements") _
            .AddOption("Choose additions", "Select a new flavor file for section and key additions").AddBlank() _
            .AddColoredFileInfo($"Current chromium.ini:       ", BuilderFile1.Path, ConsoleColor.DarkYellow) _
            .AddColoredFileInfo($"Current gecko.ini:          ", BuilderFile2.Path, ConsoleColor.DarkRed) _
            .AddColoredFileInfo($"Current save target:        ", BuilderFile3.Path, ConsoleColor.Yellow).AddBlank() _
            .AddFileInfo($"Section removals rules:     ", BuilderFile5.Path) _
            .AddFileInfo($"Key name removals rules:    ", BuilderFile6.Path) _
            .AddFileInfo($"Key value removals rules:   ", BuilderFile7.Path) _
            .AddFileInfo($"Section replacements rules: ", BuilderFile8.Path) _
            .AddFileInfo($"Key replacement rules:      ", BuilderFile9.Path) _
            .AddFileInfo($"Content addition rules      ", BuilderFile4.Path) _
            .AddBlank(BrowserBuilderModuleSettingsChanged) _
            .AddResetOpt(NameOf(BrowserBuilder), BrowserBuilderModuleSettingsChanged)

        menu.Print()

    End Sub

    ''' <summary>
    ''' Handles the user input from the BrowserBuilder menu
    ''' </summary>
    ''' 
    ''' <param name="input">
    ''' The user's input 
    ''' </param>
    ''' 
    ''' Docs last updated: 2025-08-22 | Code last updated: 2025-08-22
    Public Sub handleBrowserBuilderInput(input As String)

        Dim fileOpts = getFileOpts()
        Dim fileNums = getMenuNumbering(fileOpts, 2)

        Dim resetNum = CType(fileOpts.Count + 2, String)

        Select Case True

            ' Exit
            ' Notes: Always "0"
            Case input = "0"

                exitModule()

            ' Run (default) 
            ' Notes: Always "1", also triggered by no input if run conditions are otherwise satisfied
            Case (input = "1" OrElse input.Length = 0)

                initBrowserBuilder()

            ' File Selectors
            ' chromium.ini
            ' gecko.ini
            ' save target
            ' flavor files
            Case fileNums.Contains(input)

                Dim i = CType(input, Integer) - 2

                Dim fileName = fileOpts.Keys(i)
                Dim fileObj = fileOpts(fileName)

                changeFileParams(fileObj, BrowserBuilderModuleSettingsChanged, NameOf(BrowserBuilder), fileName, NameOf(BrowserBuilderModuleSettingsChanged))

            ' Reset Settings       
            ' Notes: Always the last option, only appears if settings have been changed
            Case BrowserBuilderModuleSettingsChanged AndAlso input = resetNum

                resetModuleSettings(NameOf(BrowserBuilder), AddressOf initDefaultBrowserBuilderSettings)

            Case Else

                setHeaderText(invInpStr, True)

        End Select

    End Sub

    ''' <summary>
    ''' Determines the current set of file selectors displayed on the menu and returns a Dictionary 
    ''' of those options and their respective files <br />
    ''' <br />
    ''' The set of possible files includes:
    ''' <list type="bullet">
    '''     
    '''     <item>
    '''     chromium.ini
    '''     </item>
    '''     
    '''     <item>
    '''     gecko.ini
    '''     </item>
    '''     
    '''     <item>
    '''     save target
    '''     </item>
    '''     
    '''     <item>
    '''     Section removals
    '''     </item>
    '''     
    '''     <item>
    '''     Key name removals
    '''     </item>
    '''     
    '''     <item>
    '''     Key value removals
    '''     </item>
    '''     
    '''     <item>
    '''     Section replacements
    '''     </item>
    '''     
    '''     <item>
    '''     Key replacements
    '''     </item>
    '''     
    '''     <item>
    '''     Additions
    '''     </item>
    '''     
    ''' </list>
    ''' 
    ''' </summary>
    ''' 
    ''' <returns> 
    ''' The set of <c> iniFile </c> properties for an object currently displayed on the menu
    ''' </returns>
    ''' 
    ''' Docs last updated: 2025-08-22 | Code last updated: 2025-08-22
    Private Function getFileOpts() As Dictionary(Of String, iniFile)

        Dim selectors As New Dictionary(Of String, iniFile)

        selectors.Add(NameOf(BuilderFile1), BuilderFile1)
        selectors.Add(NameOf(BuilderFile2), BuilderFile2)
        selectors.Add(NameOf(BuilderFile3), BuilderFile3)

        selectors.Add(NameOf(BuilderFile5), BuilderFile5)
        selectors.Add(NameOf(BuilderFile6), BuilderFile6)
        selectors.Add(NameOf(BuilderFile9), BuilderFile9)
        selectors.Add(NameOf(BuilderFile7), BuilderFile7)
        selectors.Add(NameOf(BuilderFile8), BuilderFile8)
        selectors.Add(NameOf(BuilderFile4), BuilderFile4)

        Return selectors

    End Function

End Module
