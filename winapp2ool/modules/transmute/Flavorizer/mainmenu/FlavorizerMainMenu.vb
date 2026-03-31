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
''' Displays the Flavorizer main menu to the user and handles their input accordingly
''' </summary>
Public Module FlavorizerMainMenu

    ''' <summary>
    ''' Builds and returns the Flavorizer main menu. <br />
    ''' Both <c> printFlavorizerMainMenu </c> and <c> handleFlavorizerMainMenuUserInput </c>
    ''' call this function to ensure menu numbering seen by the user stays in sync with dispatch.
    ''' </summary>
    Private Function buildFlavorizerMenu() As MenuSection

        Console.WindowHeight = 45

        Dim menuDescriptionLines = {"Apply a 'flavor' to an ini file using a collection of correction files",
                                    "Flavorization is applied in this order:",
                                    "Section Removal → Key Name Removal → Key Value Removal → Section Replacement → Key Replacement → Additions",
                                    "All flavor files are optional - specify only the ones you need"}

        Return MenuSection.CreateCompleteMenu(NameOf(Flavorizer), menuDescriptionLines, ConsoleColor.Yellow) _
            .AddDispatchedColoredOption("Run (default)", "Apply the flavor to the base file", GetRedGreen(FlavorizerFile1.Name.Length = 0),
                Sub() If Not denyActionWithHeader(FlavorizerFile1.Name.Length = 0, "You must select a base file") Then initFlavorizer()) _
            .AddBlank() _
            .AddDispatchedColoredOption("Auto detect Flavor", "Attempt to detect the flavor files by name", ConsoleColor.Yellow,
                Sub() DetectFlavorFiles(FlavorizerFile9.Dir)) _
            .AddDispatchedToggle("winapp2.ini formatting", $"{enStr(FlavorizeAsWinapp)} formatting output as winapp2.ini", FlavorizeAsWinapp,
                Sub() toggleModuleSetting("winapp2.ini formatting", NameOf(Flavorizer), GetType(FlavorizerSettings),
                                          NameOf(FlavorizeAsWinapp), NameOf(FlavorizerModuleSettingsChanged))) _
            .AddBlank() _
            .AddDispatchedOption("Change base file", "Select the base file to be flavorized",
                Sub()
                    changeFile2Params(FlavorizerFile1, FlavorizerModuleSettingsChanged, NameOf(Flavorizer),
                                      NameOf(FlavorizerFile1), NameOf(FlavorizerModuleSettingsChanged))
                End Sub) _
            .AddDispatchedOption("Change save target", "Select where to save the flavorized result",
                Sub()
                    changeFile2Params(FlavorizerFile2, FlavorizerModuleSettingsChanged, NameOf(Flavorizer),
                                      NameOf(FlavorizerFile2), NameOf(FlavorizerModuleSettingsChanged))
                End Sub) _
            .AddBlank() _
            .AddDispatchedOption("Change target directory", "Change the directory within which to scan for Flavor files",
                Sub()
                    changeFile2Params(FlavorizerFile9, FlavorizerModuleSettingsChanged, NameOf(Flavorizer),
                                      NameOf(FlavorizerFile9), NameOf(FlavorizerModuleSettingsChanged))
                    FlavorizerFile9.Name = ""
                    SaveModule2(NameOf(Flavorizer), GetType(FlavorizerSettings))
                End Sub) _
            .AddBlank() _
            .AddDispatchedOption("Change section removal file", "Select the file containing sections to remove entirely",
                Sub()
                    changeFile2Params(FlavorizerFile3, FlavorizerModuleSettingsChanged, NameOf(Flavorizer),
                                      NameOf(FlavorizerFile3), NameOf(FlavorizerModuleSettingsChanged))
                End Sub) _
            .AddDispatchedOption("Change key name removal file", "Select the file containing keys to remove by name",
                Sub()
                    changeFile2Params(FlavorizerFile4, FlavorizerModuleSettingsChanged, NameOf(Flavorizer),
                                      NameOf(FlavorizerFile4), NameOf(FlavorizerModuleSettingsChanged))
                End Sub) _
            .AddDispatchedOption("Change key value removal file", "Select the file containing keys to remove by value",
                Sub()
                    changeFile2Params(FlavorizerFile5, FlavorizerModuleSettingsChanged, NameOf(Flavorizer),
                                      NameOf(FlavorizerFile5), NameOf(FlavorizerModuleSettingsChanged))
                End Sub) _
            .AddDispatchedOption("Change section replacement file", "Select the file containing sections to replace entirely",
                Sub()
                    changeFile2Params(FlavorizerFile6, FlavorizerModuleSettingsChanged, NameOf(Flavorizer),
                                      NameOf(FlavorizerFile6), NameOf(FlavorizerModuleSettingsChanged))
                End Sub) _
            .AddDispatchedOption("Change key replacement file", "Select the file containing keys to replace by name",
                Sub()
                    changeFile2Params(FlavorizerFile7, FlavorizerModuleSettingsChanged, NameOf(Flavorizer),
                                      NameOf(FlavorizerFile7), NameOf(FlavorizerModuleSettingsChanged))
                End Sub) _
            .AddDispatchedOption("Change additions file", "Select the file containing sections/keys to add",
                Sub()
                    changeFile2Params(FlavorizerFile8, FlavorizerModuleSettingsChanged, NameOf(Flavorizer),
                                      NameOf(FlavorizerFile8), NameOf(FlavorizerModuleSettingsChanged))
                End Sub) _
            .AddBlank() _
            .AddColoredLine("Base and save files:", ConsoleColor.Magenta) _
            .AddLine($"Base file: {replDir(FlavorizerFile1.Path())}") _
            .AddLine($"Save target: {replDir(FlavorizerFile2.Path())}") _
            .AddBlank() _
            .AddColoredLine($"Auto detect target directory: {FlavorizerFile9.Dir}", ConsoleColor.Yellow) _
            .AddBlank() _
            .AddColoredLine("Flavor files (applied in order):", ConsoleColor.Magenta) _
            .AddColoredLine($"Section removal: {getFileMenuName(FlavorizerFile3)}", getFileMenuColor(FlavorizerFile3)) _
            .AddColoredLine($"Key name removal: {getFileMenuName(FlavorizerFile4)}", getFileMenuColor(FlavorizerFile4)) _
            .AddColoredLine($"Key value removal: {getFileMenuName(FlavorizerFile5)}", getFileMenuColor(FlavorizerFile5)) _
            .AddColoredLine($"Section replacement: {getFileMenuName(FlavorizerFile6)}", getFileMenuColor(FlavorizerFile6)) _
            .AddColoredLine($"Key replacement: {getFileMenuName(FlavorizerFile7)}", getFileMenuColor(FlavorizerFile7)) _
            .AddColoredLine($"Additions: {getFileMenuName(FlavorizerFile8)}", getFileMenuColor(FlavorizerFile8)) _
            .AddBlank(FlavorizerModuleSettingsChanged) _
            .AddDispatchedResetOpt(NameOf(Flavorizer), FlavorizerModuleSettingsChanged,
                Sub() resetModuleSettings(NameOf(Flavorizer), AddressOf initDefaultFlavorizerSettings))

    End Function

    ''' <summary>
    ''' Prints the Flavorizer menu to the user, showing all correction file options and current settings
    ''' </summary>
    Public Sub printFlavorizerMainMenu()

        buildFlavorizerMenu().Print()

    End Sub

    ''' <summary>
    ''' Handles the user's input from the main menu
    ''' </summary>
    '''
    ''' <param name="input">
    ''' The user's input
    ''' </param>
    Public Sub handleFlavorizerMainMenuUserInput(input As String)

        Dim intInput As Integer

        If Not Integer.TryParse(input, intInput) Then

            If input.Length = 0 Then
                If Not denyActionWithHeader(FlavorizerFile1.Name.Length = 0, "You must select a base file") Then initFlavorizer()
                Return
            End If

            setNextMenuHeaderText(invInpStr, printColor:=ConsoleColor.Red)
            Return

        End If

        If intInput = 0 Then exitModule() : Return

        If Not buildFlavorizerMenu().Dispatch(intInput) Then setNextMenuHeaderText(invInpStr, printColor:=ConsoleColor.Red)

    End Sub

    ''' <summary>
    ''' Returns a display name for a Flavorizer file chooser entry in the menu
    ''' </summary>
    '''
    ''' <param name="f">
    ''' The <c> iniFileChooser </c> whose display name will be returned
    ''' </param>
    '''
    ''' <returns>
    ''' The file name if set, or "not set" otherwise
    ''' </returns>
    Private Function getFileMenuName(f As iniFileChooser) As String

        Return If(f.Name.Length = 0, "not set", f.Name)

    End Function

    ''' <summary>
    ''' Returns the display color for a Flavorizer file chooser entry in the menu
    ''' </summary>
    '''
    ''' <param name="f">
    ''' The <c> iniFileChooser </c> whose display color will be returned
    ''' </param>
    '''
    ''' <returns>
    ''' <c> ConsoleColor.Green </c> if the file is set and exists on disk, <c> ConsoleColor.Red </c> otherwise
    ''' </returns>
    Private Function getFileMenuColor(f As iniFileChooser) As ConsoleColor

        Return GetRedGreen(Not (f.Name.Length > 0 AndAlso f.Exists()))

    End Function

End Module
