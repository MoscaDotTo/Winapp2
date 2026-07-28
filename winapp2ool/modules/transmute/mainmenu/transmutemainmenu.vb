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
''' Displays the Transmute main menu to the user and handles their input accordingly
''' </summary>
Public Module transmuteMainMenu

    ''' <summary>
    ''' Builds and returns the Transmute main menu. <br />
    ''' Both <c> printTransmuteMainMenu </c> and <c> handleTransmuteUserInput </c>
    ''' call this function to ensure menu numbering seen by the user stays in sync with dispatch.
    ''' </summary>
    Private Function buildTransmuteMenu() As MenuSection

        adjustTransmuteConsoleHeight()

        Dim TransmuteModeIsReplace = Transmutator = TransmuteMode.Replace
        Dim ReplaceModeIsByKey = TransmuteReplaceMode = ReplaceMode.ByKey
        Dim TransmuteModeIsRemove = Transmutator = TransmuteMode.Remove
        Dim RemoveModeIsByKey = TransmuteRemoveMode = RemoveMode.ByKey

        Dim modeColor = If(Transmutator = TransmuteMode.Add, ConsoleColor.Green, If(Transmutator = TransmuteMode.Replace, ConsoleColor.Yellow, ConsoleColor.Red))
        Dim subModeColor = If(TransmuteModeIsReplace, ConsoleColor.Magenta, If(TransmuteRemoveMode = RemoveMode.BySection, ConsoleColor.Magenta, ConsoleColor.Cyan))
        Dim replaceModeDesc = $"Currently replacing {If(ReplaceModeIsByKey, "individual key values within particular", "entire")} sections in the base file"
        Dim removeModeDesc = $"Currently removing {If(RemoveModeIsByKey, "individual keys within particular", "entire")} sections from the base file"
        Dim keyModeColor = If(TransmuteRemoveKeyMode = RemoveKeyMode.ByName, ConsoleColor.Magenta, ConsoleColor.DarkYellow)
        Dim keyRemovalModeDesc = $"Currently removing keys by {If(TransmuteRemoveKeyMode = RemoveKeyMode.ByName, "name matching only", "name (numberless) and value pair matching")}"

        Dim menuDescriptionLines = {"Transmute an ini file with sections and keys provided by a second ini file",
                                    "Add entire sections or individual keys to specific sections",
                                    "Replace entire sections or individual keys Values within specific sections by Name",
                                    "Remove entire sections or individual keys within specific sections by Name or Value"}

        Return MenuSection.CreateCompleteMenu("Transmute", menuDescriptionLines, ConsoleColor.Magenta) _
            .AddDispatchedColoredOption("Run (default)", "Perform the transmutation", GetRedGreen(TransmuteFile1.Name.Length = 0),
                Sub() If Not denyActionWithHeader(TransmuteFile2.Name.Length = 0, "You must select a valid source file") Then initTransmute()) _
            .AddDispatchedColoredOption("Open Flavorizer", "Apply a complex series of modifications to an ini file", ConsoleColor.Yellow,
                Sub() initModule("Flavorizer", AddressOf printFlavorizerMainMenu, AddressOf handleFlavorizerMainMenuUserInput)) _
            .AddBlank() _
            .AddLine("Preset source file names:") _
            .AddDispatchedOption("Removed Entries", "Select 'Removed Entries.ini'", Sub() setSourceFileName("Removed Entries.ini")) _
            .AddDispatchedOption("custom", "Select 'custom.ini'", Sub() setSourceFileName("custom.ini")) _
            .AddDispatchedOption("winapp3.ini", "Select 'winapp3.ini'", Sub() setSourceFileName("winapp3.ini")) _
            .AddDispatchedOption("browsers.ini", "Select 'browsers.ini'", Sub() setSourceFileName("browsers.ini")) _
            .AddDispatchedOption("uwp.ini", "Select 'uwp.ini'", Sub() setSourceFileName("uwp.ini")) _
            .AddBlank() _
            .AddDispatchedOption("Change base file", "Select a new base file to be modified",
                Sub() changeFile2Params(TransmuteFile1, TransmuteModuleSettingsChanged, NameOf(Transmute), NameOf(TransmuteFile1), NameOf(TransmuteModuleSettingsChanged))) _
            .AddDispatchedOption("Change source file", "Select the source file providing modifications for the base file",
                Sub() changeFile2Params(TransmuteFile2, TransmuteModuleSettingsChanged, NameOf(Transmute), NameOf(TransmuteFile2), NameOf(TransmuteModuleSettingsChanged))) _
            .AddDispatchedOption("Change save target", "Select a save target for the output",
                Sub() changeFile2Params(TransmuteFile3, TransmuteModuleSettingsChanged, NameOf(Transmute), NameOf(TransmuteFile3), NameOf(TransmuteModuleSettingsChanged))) _
            .AddBlank() _
            .AddColoredFileInfo($"Base file:   ", TransmuteFile1.Path(), ConsoleColor.Magenta) _
            .AddColoredFileInfo($"Source file: ", TransmuteFile2.Path(), ConsoleColor.Cyan) _
            .AddColoredFileInfo($"Save target: ", TransmuteFile3.Path(), ConsoleColor.Yellow) _
            .AddBlank() _
            .AddDispatchedToggle("Syntax", "saving the output with winapp2.ini syntax", UseWinapp2Syntax,
                Sub() toggleModuleSetting("Winapp2.ini Syntax", NameOf(Transmute), GetType(transmuteSettings), NameOf(UseWinapp2Syntax), NameOf(TransmuteModuleSettingsChanged))) _
            .AddDispatchedToggle("Global sections", "treating [*] and [*Map:] source sections as global operations", RecognizeGlobalSections,
                Sub() toggleModuleSetting("Global Sections", NameOf(Transmute), GetType(transmuteSettings), NameOf(RecognizeGlobalSections), NameOf(TransmuteModuleSettingsChanged))) _
            .AddBlank() _
            .AddDispatchedOption("Change transmute mode", "Cycle through primary transmute modes (Add/Replace/Remove)",
                Sub() CycleEnumProperty(NameOf(Transmutator), "Transmute Mode", GetType(transmuteSettings), NameOf(Transmute),
                                        TransmuteModuleSettingsChanged, NameOf(TransmuteModuleSettingsChanged), ConsoleColor.Magenta)) _
            .AddColoredLine($"Transmute mode: {Transmutator}", modeColor) _
            .AddBlank(TransmuteModeIsReplace) _
            .AddDispatchedOption("Change Replace mode", "Toggle between replacing by section or by key", condition:=TransmuteModeIsReplace,
                handler:=Sub() CycleEnumProperty(NameOf(TransmuteReplaceMode), "Replace Mode", GetType(transmuteSettings), NameOf(Transmute),
                                                 TransmuteModuleSettingsChanged, NameOf(TransmuteModuleSettingsChanged), ConsoleColor.Magenta)) _
            .AddColoredLine(replaceModeDesc, subModeColor, condition:=TransmuteModeIsReplace) _
            .AddBlank(TransmuteModeIsRemove) _
            .AddDispatchedOption("Change Remove mode", "Toggle between removing by section or by key", condition:=TransmuteModeIsRemove,
                handler:=Sub() CycleEnumProperty(NameOf(TransmuteRemoveMode), "Remove Mode", GetType(transmuteSettings), NameOf(Transmute),
                                                 TransmuteModuleSettingsChanged, NameOf(TransmuteModuleSettingsChanged), ConsoleColor.Magenta)) _
            .AddColoredLine(removeModeDesc, subModeColor, condition:=TransmuteModeIsRemove) _
            .AddDispatchedOption("Change Key Removal mode", "Toggle between removing keys by name or by value", condition:=TransmuteModeIsRemove AndAlso RemoveModeIsByKey,
                handler:=Sub() CycleEnumProperty(NameOf(TransmuteRemoveKeyMode), "Key Removal Mode", GetType(transmuteSettings), NameOf(Transmute),
                                                 TransmuteModuleSettingsChanged, NameOf(TransmuteModuleSettingsChanged), ConsoleColor.Magenta)) _
            .AddColoredLine(keyRemovalModeDesc, keyModeColor, condition:=TransmuteModeIsRemove AndAlso RemoveModeIsByKey) _
            .AddBlank(TransmuteModuleSettingsChanged) _
            .AddDispatchedResetOpt(NameOf(Transmute), TransmuteModuleSettingsChanged, Sub() resetModuleSettings(NameOf(Transmute), AddressOf initDefaultTransmuteSettings))

    End Function

    ''' <summary>
    ''' Sets <c> TransmuteFile2 </c>'s name to <paramref name="fileName"/>, persists the change,
    ''' and updates the menu header to confirm the selection
    ''' </summary>
    '''
    ''' <param name="fileName">The preset source file name to assign</param>
    Private Sub setSourceFileName(fileName As String)

        TransmuteFile2.Name = fileName
        TransmuteModuleSettingsChanged = True
        SaveModule2(NameOf(Transmute), GetType(transmuteSettings))
        setNextMenuHeaderText($"Source file name set to {TransmuteFile2.Name}", printColor:=ConsoleColor.Yellow)

    End Sub

    ''' <summary>
    ''' Prints the <c> Transmute </c> menu to the user, includes predefined source file choices and sub mode options
    ''' </summary>
    Public Sub printTransmuteMainMenu()

        buildTransmuteMenu().Print()

    End Sub

    ''' <summary>
    ''' Once per run of the module, we'll set the console height to be large enough so as to
    ''' be able to display the largest number of options that we display
    ''' </summary>
    Private Sub adjustTransmuteConsoleHeight()

        If Console.WindowHeight <= 43 Then Console.WindowHeight = 44

    End Sub

    ''' <summary>
    ''' Handles the user's input from the main menu
    ''' </summary>
    '''
    ''' <param name="input">
    ''' The user's input
    ''' </param>
    Public Sub handleTransmuteUserInput(input As String)

        Dim intInput As Integer

        If Not Integer.TryParse(input, intInput) Then

            If input.Length = 0 Then

                If Not denyActionWithHeader(TransmuteFile2.Name.Length = 0, "You must select a valid source file") Then initTransmute()
                Return

            End If

            setNextMenuHeaderText(invInpStr, printColor:=ConsoleColor.Red)
            Return

        End If

        If intInput = 0 Then exitModule() : Return

        If Not buildTransmuteMenu().Dispatch(intInput) Then setNextMenuHeaderText(invInpStr, printColor:=ConsoleColor.Red)

    End Sub

End Module
