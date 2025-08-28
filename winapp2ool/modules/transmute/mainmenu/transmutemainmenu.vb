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
''' Displays the Transmute main menu to the user and handles their input accordingly
''' </summary>
''' 
''' Docs last updated: 2025-07-15 | Code last updated: 2025-07-15
Public Module transmuteMainMenu

    ''' <summary> 
    ''' Prints the <c> Transmute </c> menu to the user, includes predefined source files choices and sub mode options
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-07-15 | Code last updated: 2025-07-15
    Public Sub printTransmuteMainMenu()

        adjustTransmuteConsoleHeight()

        Dim modeColor = If(Transmutator = TransmuteMode.Add, ConsoleColor.Green, If(Transmutator = TransmuteMode.Replace, ConsoleColor.Yellow, ConsoleColor.Red))
        Dim subModeColor = If(TransmuteModeIsReplace, ConsoleColor.Magenta, If(TransmuteRemoveMode = RemoveMode.BySection, ConsoleColor.Magenta, ConsoleColor.Cyan))
        Dim replaceModeDesc = $"Currently replacing {If(ReplaceModeIsByKey, "individual key values within particular", "entire")} sections in the base file"
        Dim removeModeDesc = $"Currently removing {If(RemoveModeIsByKey, "individual keys within particular", "entire")} sections from the base file"
        Dim keyModeColor = If(TransmuteRemoveKeyMode = RemoveKeyMode.ByName, ConsoleColor.Magenta, ConsoleColor.DarkYellow)
        Dim keyRemovalModeDesc = $"Currently removing keys by {If(RemoveKeyModeIsByName, "name matching only", "name (numberless) and value pair matching")}"

        Dim menuDescriptionLines = {"Transmute an ini file with sections and keys provided by a second ini file",
                      "Features by mode:",
                      "Add entire sections or individual keys to specific sections",
                      "Replace entire sections or individual keys within specific sections by Value",
                      "Remove entire sections or individual keys within specific sections by Name or Value"}

        Dim menu = MenuSection.CreateCompleteMenu("Transmute", menuDescriptionLines, ConsoleColor.Magenta)

        menu.AddColoredOption("Run (default", "Perform the transmutation", GetRedGreen(TransmuteFile1.Name.Length = 0)) _
            .AddColoredOption("Open Flavorizer", "Apply a complex series of modifications to an ini file", ConsoleColor.Yellow).AddBlank() _
            .AddLine("Preset source files") _
            .AddOption("Removed Entries", "Select 'Removed Entries.ini'") _
            .AddOption("Custom", "Select 'Custom.ini'") _
            .AddOption("winapp3.ini", "Select 'winapp3.ini'") _
            .AddOption("browsers.ini", "Select 'browsers.ini'").AddBlank() _
            .AddOption("Change base file", "Select a new base file to be modified") _
            .AddOption("Change source file", "Select the source file providing modifications for the base file") _
            .AddOption("Change save target", "Select a save target for the output").AddBlank() _
            .AddOption("Change transmute mode", "Cycle through primary transmute modes (Add/Replace/Remove)") _
            .AddColoredLine($"Transmute mode: {Transmutator}", modeColor) _
            .AddBlank(TransmuteModeIsReplace) _
            .AddOption("Change Replace mode", "Toggle between replacing by section or by key", TransmuteModeIsReplace) _
            .AddColoredLine(replaceModeDesc, subModeColor, condition:=TransmuteModeIsReplace) _
            .AddBlank(TransmuteModeIsRemove) _
            .AddOption("Change Remove mode", "Toggle between removing by section or by key", TransmuteModeIsRemove) _
            .AddColoredLine(removeModeDesc, subModeColor, condition:=TransmuteModeIsRemove) _
            .AddOption("Change Key Removal mode", "Toggle between removing keys by name or by value", TransmuteModeIsRemove AndAlso RemoveModeIsByKey) _
            .AddColoredLine(keyRemovalModeDesc, keyModeColor, condition:=TransmuteModeIsRemove AndAlso RemoveModeIsByKey) _
            .AddBlank(TransmuteModuleSettingsChanged) _
            .AddResetOpt(NameOf(Transmute), TransmuteModuleSettingsChanged)

        menu.Print()

    End Sub

    ''' <summary>
    ''' Once per run of the module, we'll set the console height to be large enough so as to 
    ''' be able to display the largest number of options that we display 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-07-15 | Code last updated: 2025-07-15
    Private Sub adjustTransmuteConsoleHeight()

        If Console.WindowHeight > 42 Then Return

        Console.WindowHeight = 43

    End Sub

    ''' <summary> 
    ''' Handles the user's input from the main menu 
    ''' </summary>
    ''' 
    ''' <param name="input"> 
    ''' The user's input 
    ''' </param>
    ''' 
    ''' Docs last updated: 2025-07-15 | Code last updated: 2025-07-15
    Public Sub handleTransmuteUserInput(input As String)

        Dim presetFileNames As New Dictionary(Of String, String) From {
            {"3", "Removed Entries.ini"},
            {"4", "Custom.ini"},
            {"5", "winapp3.ini"},
            {"6", "browsers.ini"}
        }

        Dim transmuteFiles As New Dictionary(Of String, iniFile) From {
            {NameOf(TransmuteFile1), TransmuteFile1},
            {NameOf(TransmuteFile2), TransmuteFile2},
            {NameOf(TransmuteFile3), TransmuteFile3}
        }

        Dim fileOpts = {"7", "8", "9"}

        Select Case True

            ' Option Name:                                 Exit 
            ' Option States:
            ' Default                                      -> 0 (Default)
            Case input = "0"

                exitModule()

            ' Option Name:                                 Run (default)
            ' Option States:
            ' Default                                      -> 1 (default)
            Case input = "1" OrElse input.Length = 0

                If Not denyActionWithHeader(TransmuteFile2.Name.Length = 0, "You must select a source file") Then initTransmute()

            ' Option Name:                                 Open Flavorizer
            ' Option States:
            ' Default                                      -> 2 (default)
            Case input = "2"

                initModule("Flavorizer", AddressOf printFlavorizerMainMenu, AddressOf handleFlavorizerMainMenuUserInput)

            ' Option Name:                                 Preset File Names
            ' Option States:
            ' Removed Entries.ini                          -> 3 (default)
            ' Custom.ini                                   -> 4 (default)
            ' winapp3.ini                                  -> 5 (default)
            ' browsers.ini                                 -> 6 (default)

            Case presetFileNames.ContainsKey(input)

                changeBaseFileName(presetFileNames(input))

          ' Option Name:                                   File Selectors
          ' Option States:
          ' Change base file                              -> 7 (default)
          ' Change source file                            -> 8 (default)
          ' Change save target                            -> 9 (default)
            Case fileOpts.Contains(input)

                Dim i = CType(input, Integer) - 7
                changeFileParams(transmuteFiles(transmuteFiles.Keys(i)), TransmuteModuleSettingsChanged, NameOf(Transmute),
                                 transmuteFiles.Keys(i), NameOf(TransmuteModuleSettingsChanged))

            ' Option Name:                                 Change Transmute Mode 
            ' Option States:
            ' Default                                      -> 10 (default)
            Case input = "10"

                cycleEnumSetting(Transmutator, GetType(TransmuteMode), "Transmutator",
                                 TransmuteModuleSettingsChanged, NameOf(Transmute), NameOf(TransmuteModuleSettingsChanged))

                updateAllTransmuteEnumSettings()

            ' Option Name:                                 Change Replace Mode
            ' Option States:
            ' TransmuteModeIsReplace = False               -> Unavailable (not displayed)
            ' Default                                      -> 11 (default)
            Case input = "11" AndAlso TransmuteModeIsReplace

                cycleEnumSetting(TransmuteReplaceMode, GetType(ReplaceMode), "Replace Mode",
                                 TransmuteModuleSettingsChanged, NameOf(Transmute), NameOf(TransmuteModuleSettingsChanged))

                updateAllTransmuteEnumSettings()

            ' Option Name:                                 Change Remove Mode
            ' Option States:
            ' TransmuteModeIsRemove = False                -> Unavailable (not displayed)
            ' Default                                      -> 12 (default)
            Case input = "11" AndAlso TransmuteModeIsRemove

                cycleEnumSetting(TransmuteRemoveMode, GetType(RemoveMode), "Remove Mode",
                                 TransmuteModuleSettingsChanged, NameOf(Transmute), NameOf(TransmuteModuleSettingsChanged))

                updateAllTransmuteEnumSettings()

            ' Option Name:                                 Change Remove Key Mode
            ' Option States:
            ' RemoveModeIsByKey = False                    -> Unavailable (not displayed)
            ' Default                                      -> 11 (default)
            Case input = "12" AndAlso TransmuteModeIsRemove AndAlso RemoveModeIsByKey

                cycleEnumSetting(TransmuteRemoveKeyMode, GetType(RemoveKeyMode), "Remove Key Mode",
                                 TransmuteModuleSettingsChanged, NameOf(Transmute), NameOf(TransmuteModuleSettingsChanged))

                updateAllTransmuteEnumSettings()

            ' Option Name:                                  Reset Settings 
            ' Option States: 
            ' ModuleSettingsChanged = False                 -> Unavailable (not displayed) 
            ' TransmuteModeIsAdd                            -> 11 (default)
            ' TransmuteModeIsReplace (+1)                   -> 12  
            ' TransmuteModeIsRemove (+1) and BySection (+0) -> 12
            ' TransmuteModeIsRemove (+1) and ByKey (+1)     -> 13
            Case TransmuteModuleSettingsChanged AndAlso
                 input = computeMenuNumber(11, {TransmuteModeIsAdd, TransmuteModeIsReplace,
                                               TransmuteModeIsRemove AndAlso RemoveModeIsBySection,
                                               TransmuteModeIsRemove AndAlso RemoveModeIsByKey}, {0, 1, 1, 2})

                resetModuleSettings(NameOf(Transmute), AddressOf initDefaultTransmuteSettings)

            Case Else

                setHeaderText(invInpStr, True, printColor:=ConsoleColor.Red)

        End Select

    End Sub

End Module