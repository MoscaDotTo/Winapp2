Option Strict On
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

''' <summary>
''' Displays the Flavorizer main menu to the user and handles their input accordingly
''' </summary>
''' 
''' Docs last updated: 2025-08-01 | Code last updated: 2025-08-06
Public Module FlavorizerMainMenu

    ''' <summary>
    ''' Prints the Flavorizer menu to the user, showing all correction file options
    ''' and current settings
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-08-01 | Code last updated: 2025-08-06
    Public Sub printFlavorizerMainMenu()

        Console.WindowHeight = Console.LargestWindowHeight

        Dim menuDescriptionLines = {"Apply a 'flavor' to an ini file using a collection of correction files",
                                    "Flavorization is applied in this order:",
                                    "Section Removal → Key Name Removal → Key Value Removal → Section Replacement → Key Replacement → Additions",
                                    "All flavor files are optional - specify only the ones you need"}

        Dim menu = MenuSection.CreateCompleteMenu(NameOf(Flavorizer), menuDescriptionLines, ConsoleColor.Yellow)

        menu.AddToggle("Run (default)", "Apply the flavor to the base file", FlavorizerFile1.Name.Length > 0).AddBlank() _
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
            .AddResetOpt(NameOf(Flavorizer), FlavorizerModuleSettingsChanged) _
            .Print(False)

    End Sub

    ''' <summary>
    ''' Handles the user's input from the main menu
    ''' </summary>
    ''' 
    ''' <param name="input">
    ''' The user's input
    ''' </param>
    ''' 
    ''' Docs last updated: 2025-08-01 | Code last updated: 2025-08-06
    Public Sub handleFlavorizerMainMenuUserInput(input As String)

        Dim flavorFiles As New Dictionary(Of String, iniFile) From {
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

        Dim fileOpts = {"4", "5", "6", "7", "8", "9", "10", "11", "12"}

        Select Case True

            ' Option Name:                                 Exit 
            ' Option States:
            ' Default                                      -> 0 (Default)
            Case input = "0"

                exitModule()

            ' Option Name:                                 Run 
            ' Option States:
            ' Default                                      -> 1 (Default)
            Case input = "1" OrElse input.Length = 0

                If Not denyActionWithHeader(FlavorizerFile1.Name.Length = 0, "You must select a base file") Then initFlavorizer()


            ' Option Name:                                 Detect Flavor Files
            ' Option States:
            ' Default                                      -> 11 (Default)
            Case input = "2"

                DetectFlavorFiles(FlavorizerFile9.Dir)

            ' Option Name:                                 Toggle winapp2.ini formatting 
            ' Option States:
            ' Default                                      -> 10 (Default)
            Case input = "3"

                toggleSettingParam(FlavorizeAsWinapp, "Winapp2.ini formatting", FlavorizerModuleSettingsChanged,
                                   NameOf(Flavorizer), NameOf(FlavorizeAsWinapp), NameOf(FlavorizerModuleSettingsChanged))

            ' Option Name:                                 File Choosers
            ' Option States:
            ' Change base file                             -> 4 
            ' Change save target                           -> 5
            ' Change target directory                      -> 6
            ' Section removal                              -> 7
            ' Key name removal                             -> 8
            ' Key value removal                            -> 9
            ' Section replacement                          -> 10
            ' Key replacement                              -> 11
            ' Additions                                    -> 12
            Case fileOpts.Contains(input)

                Dim i = CType(input, Integer) - 4
                changeFileParams(flavorFiles(flavorFiles.Keys(i)), FlavorizerModuleSettingsChanged,
                                 NameOf(Flavorizer), flavorFiles.Keys(i), NameOf(FlavorizerModuleSettingsChanged))

                ' This file is used only to hold the Target directory and never holds a name 
                If input = "6" Then

                    FlavorizerFile9.Name = ""
                    updateSettings(NameOf(Flavorizer), $"{NameOf(FlavorizerFile9)}_Name", "")

                End If

            ' Option Name:                                 Reset Flavorizer Settings
            ' Option States:
            ' FlavorizerModuleSettingsChanged = False     -> Unavailable (not displayed)
            ' FlavorizerModuleSettingsChanged = True      -> 13 (Default)
            Case input = "13" AndAlso FlavorizerModuleSettingsChanged

                resetModuleSettings(NameOf(Flavorizer), AddressOf initDefaultFlavorizerSettings)

            Case Else

                setHeaderText(invInpStr, True)

        End Select

    End Sub

End Module
