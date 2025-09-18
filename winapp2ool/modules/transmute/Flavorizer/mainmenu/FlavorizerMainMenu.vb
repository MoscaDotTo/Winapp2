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
''' Displays the Flavorizer main menu to the user and handles their input accordingly
''' </summary>
''' 
''' Docs last updated: 2025-08-01
Public Module FlavorizerMainMenu

    ''' <summary>
    ''' Prints the Flavorizer menu to the user, showing all correction file options
    ''' and current settings
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-08-01 | Code last updated: 2025-08-06
    Public Sub printFlavorizerMainMenu()

        Console.WindowHeight = 45

        Dim menuDescriptionLines = {"Apply a 'flavor' to an ini file using a collection of correction files",
                                    "Flavorization is applied in this order:",
                                    "Section Removal → Key Name Removal → Key Value Removal → Section Replacement → Key Replacement → Additions",
                                    "All flavor files are optional - specify only the ones you need"}

        Dim menu = MenuSection.CreateCompleteMenu(NameOf(Flavorizer), menuDescriptionLines, ConsoleColor.Yellow)

        menu.AddColoredOption("Run (default)", "Apply the flavor to the base file", GetRedGreen(FlavorizerFile1.Name.Length = 0)).AddBlank() _
            .AddColoredOption("Auto detect Flavor", "Attempt to detect the flavor files by name", ConsoleColor.Yellow) _
            .AddColoredOption("winapp2.ini formatting", $"{enStr(FlavorizeAsWinapp)} formatting output as winapp2.ini", If(FlavorizeAsWinapp, ConsoleColor.Green, ConsoleColor.Red)).AddBlank() _
            .AddOption("Change base file", "Select the base file to be flavorized") _
            .AddOption("Change save target", "Select where to save the flavorized result").AddBlank() _
            .AddOption("Change target directory", "Change the directory within which to scan for Flavor files").AddBlank() _
            .AddOption("Change section removal file", "Select the file containing sections to remove entirely") _
            .AddOption("Change key name removal file", "Select the file containing keys to remove by name") _
            .AddOption("Change key value removal file", "Select the file containing keys to remove by value") _
            .AddOption("Change section replacement file", "Select the file containing sections to replace entirely") _
            .AddOption("Change key replacement file", "Select the file containing keys to replace by name") _
            .AddOption("Change additions file", "Select the file containing sections/keys to add").AddBlank() _
            .AddColoredLine("Base and save files:", ConsoleColor.Magenta) _
            .AddLine($"Base file: {replDir(FlavorizerFile1.Path)}") _
            .AddLine($"Save target: {replDir(FlavorizerFile2.Path)}").AddBlank() _
            .AddColoredLine($"Auto detect target directory: {FlavorizerFile9.Dir}", ConsoleColor.Yellow).AddBlank() _
            .AddColoredLine("Flavor files (applied in order):", ConsoleColor.Magenta) _
            .AddColoredLine($"Section removal: {getFileMenuName(FlavorizerFile3)}", getFileMenuColor(FlavorizerFile3)) _
            .AddColoredLine($"Key name removal: {getFileMenuName(FlavorizerFile4)}", getFileMenuColor(FlavorizerFile4)) _
            .AddColoredLine($"Key value removal: {getFileMenuName(FlavorizerFile5)}", getFileMenuColor(FlavorizerFile5)) _
            .AddColoredLine($"Section replacement: {getFileMenuName(FlavorizerFile6)}", getFileMenuColor(FlavorizerFile6)) _
            .AddColoredLine($"Key replacement: {getFileMenuName(FlavorizerFile7)}", getFileMenuColor(FlavorizerFile7)) _
            .AddColoredLine($"Additions: {getFileMenuName(FlavorizerFile8)}", getFileMenuColor(FlavorizerFile8)) _
            .AddBlank(FlavorizerModuleSettingsChanged) _
            .AddResetOpt(NameOf(Flavorizer), FlavorizerModuleSettingsChanged)

        menu.Print()

    End Sub

    ''' <summary>
    ''' Handles the user's input from the main menu
    ''' </summary>
    ''' 
    ''' <param name="input">
    ''' The user's input
    ''' </param>
    ''' 
    ''' Docs last updated: 2025-08-01 | Code last updated: 2025-08-23
    Public Sub handleFlavorizerMainMenuUserInput(input As String)

        Dim fileOpts = getFileOpts()
        Dim fileNums = getMenuNumbering(fileOpts, 4)

        Dim resetNum = CType(4 + fileOpts.Count + 1, String)

        Select Case True

            ' Exit 
            ' Notes: Always "0"
            Case input = "0"

                exitModule()

            ' Run (default)
            ' Notes: Always "1", also triggered by no input if run conditions are otherwise satisfied 
            Case input = "1" OrElse input.Length = 0

                If Not denyActionWithHeader(FlavorizerFile1.Name.Length = 0, "You must select a base file") Then initFlavorizer()

            ' Detect Flavor Files
            ' Notes: Always "2"
            Case input = "2"

                DetectFlavorFiles(FlavorizerFile9.Dir)

            ' Toggle winapp2.ini formatting 
            ' Notes: Always "3"
            Case input = "3"

                toggleSettingParam(FlavorizeAsWinapp, "Winapp2.ini formatting", FlavorizerModuleSettingsChanged,
                                   NameOf(Flavorizer), NameOf(FlavorizeAsWinapp), NameOf(FlavorizerModuleSettingsChanged))

            ' File selectors
            ' Base file 
            ' Save target
            ' Target directory
            ' Section removals
            ' Key name removals
            ' Key value removals
            ' Section replacements
            ' Key replacements
            ' Additions
            Case fileNums.Contains(input)

                Dim i = CType(input, Integer) - 4

                Dim fileName = fileOpts.Keys(i)
                Dim fileObj = fileOpts(fileName)

                changeFileParams(fileObj, FlavorizerModuleSettingsChanged, NameOf(Flavorize),
                                 fileName, NameOf(FlavorizerModuleSettingsChanged))


                ' This file is used only to hold the Target directory and never holds a name 
                FlavorizerFile9.Name = ""
                updateSettings(NameOf(Flavorizer), $"{NameOf(FlavorizerFile9)}_Name", "")

            ' Reset settings
            ' Notes: Only available after a setting has been changed, always comes last in the option list
            Case FlavorizerModuleSettingsChanged AndAlso input = resetNum

                resetModuleSettings(NameOf(Flavorizer), AddressOf initDefaultFlavorizerSettings)

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
    ''' <item>
    ''' Base file
    ''' </item>
    ''' 
    ''' <item>
    ''' Save target
    ''' </item>
    ''' 
    ''' <item>
    ''' Target directory
    ''' </item>
    '''     
    ''' <item>
    ''' Section removals
    ''' </item>
    ''' 
    ''' <item>
    ''' Key name removals
    ''' </item>
    ''' 
    ''' <item>
    ''' Key value removals
    ''' </item>
    ''' 
    ''' <item>
    ''' Section replacements
    ''' </item>
    ''' 
    ''' <item>
    ''' Key replacements
    ''' </item>
    ''' 
    ''' <item>
    ''' Additions
    ''' </item>
    ''' 
    ''' </list>
    ''' 
    ''' </summary>
    ''' 
    ''' <returns> 
    ''' The set of <c> iniFile </c> properties currently displayed on the menu
    ''' </returns>
    ''' 
    ''' Docs last updated: 2025-08-23 | Code last updated: 2025-08-23
    Private Function getFileOpts() As Dictionary(Of String, iniFile)

        Dim selectors As New Dictionary(Of String, iniFile) From {
            {NameOf(FlavorizerFile1), FlavorizerFile1},
            {NameOf(FlavorizerFile2), FlavorizerFile2},
            {NameOf(FlavorizerFile9), FlavorizerFile9},
            {NameOf(FlavorizerFile3), FlavorizerFile3},
            {NameOf(FlavorizerFile4), FlavorizerFile4},
            {NameOf(FlavorizerFile5), FlavorizerFile5},
            {NameOf(FlavorizerFile6), FlavorizerFile6},
            {NameOf(FlavorizerFile7), FlavorizerFile7},
            {NameOf(FlavorizerFile8), FlavorizerFile8}
        }

        Return selectors

    End Function

End Module
