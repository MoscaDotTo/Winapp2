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

Imports System.Reflection

''' <summary>
''' Displays the Transmute main menu to the user and handles their input accordingly
''' </summary>
Public Module transmuteMainMenu

    ''' <summary> 
    ''' Prints the <c> Transmute </c> menu to the user, includes predefined source files choices and sub mode options
    ''' </summary>
    Public Sub printTransmuteMainMenu()

        adjustTransmuteConsoleHeight()
        updateAllTransmuteEnumSettings()

        Dim TransmuteModeIsReplace = Transmutator = TransmuteMode.Replace
        Dim ReplaceModeIsByKey = TransmuteReplaceMode = ReplaceMode.ByKey
        Dim TransmuteModeIsRemove = Transmutator = TransmuteMode.Remove
        Dim RemoveModeIsByKey = TransmuteRemoveMode = RemoveMode.ByKey
        Dim RemoveKeyModeIsByName = TransmuteRemoveKeyMode = RemoveKeyMode.ByName

        Dim modeColor = If(Transmutator = TransmuteMode.Add, ConsoleColor.Green, If(Transmutator = TransmuteMode.Replace, ConsoleColor.Yellow, ConsoleColor.Red))
        Dim subModeColor = If(TransmuteModeIsReplace, ConsoleColor.Magenta, If(TransmuteRemoveMode = RemoveMode.BySection, ConsoleColor.Magenta, ConsoleColor.Cyan))
        Dim replaceModeDesc = $"Currently replacing {If(ReplaceModeIsByKey, "individual key values within particular", "entire")} sections in the base file"
        Dim removeModeDesc = $"Currently removing {If(RemoveModeIsByKey, "individual keys within particular", "entire")} sections from the base file"
        Dim keyModeColor = If(TransmuteRemoveKeyMode = RemoveKeyMode.ByName, ConsoleColor.Magenta, ConsoleColor.DarkYellow)
        Dim keyRemovalModeDesc = $"Currently removing keys by {If(RemoveKeyModeIsByName, "name matching only", "name (numberless) and value pair matching")}"

        Dim menuDescriptionLines = {"Transmute an ini file with sections and keys provided by a second ini file",
                                    "Add entire sections or individual keys to specific sections",
                                    "Replace entire sections or individual keys Values within specific sections by Name",
                                    "Remove entire sections or individual keys within specific sections by Name or Value"}

        Dim menu = MenuSection.CreateCompleteMenu("Transmute", menuDescriptionLines, ConsoleColor.Magenta)

        menu.AddColoredOption("Run (default)", "Perform the transmutation", GetRedGreen(TransmuteFile1.Name.Length = 0)) _
            .AddColoredOption("Open Flavorizer", "Apply a complex series of modifications to an ini file", ConsoleColor.Yellow).AddBlank() _
            .AddLine("Preset source file names:") _
            .AddOption("Removed Entries", "Select 'Removed Entries.ini'") _
            .AddOption("Custom", "Select 'Custom.ini'") _
            .AddOption("winapp3.ini", "Select 'winapp3.ini'") _
            .AddOption("browsers.ini", "Select 'browsers.ini'").AddBlank() _
            .AddOption("Change base file", "Select a new base file to be modified") _
            .AddOption("Change source file", "Select the source file providing modifications for the base file") _
            .AddOption("Change save target", "Select a save target for the output").AddBlank() _
            .AddColoredFileInfo($"Base file:   ", TransmuteFile1.Path, ConsoleColor.Magenta) _
            .AddColoredFileInfo($"Source file: ", TransmuteFile2.Path, ConsoleColor.Cyan) _
            .AddColoredFileInfo($"Save target: ", TransmuteFile3.Path, ConsoleColor.Yellow).AddBlank _
            .AddToggle("Syntax", "saving the output with winapp2.ini syntax", UseWinapp2Syntax).AddBlank() _
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
    Private Sub adjustTransmuteConsoleHeight()

        If Console.WindowHeight <= 42 Then Console.WindowHeight = 43

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

            setNextMenuHeaderText(invInpStr, printColor:=ConsoleColor.Red)
            Return

        End If

        Dim presetFileNames As New List(Of String) From {
            "Removed Entries.ini",
            "Custom.ini",
            "winapp3.ini",
            "browsers.ini"}

        Dim transmuteFiles As New Dictionary(Of String, iniFile) From {
            {NameOf(TransmuteFile1), TransmuteFile1},
            {NameOf(TransmuteFile2), TransmuteFile2},
            {NameOf(TransmuteFile3), TransmuteFile3}}

        Dim fileDescs = {
                         "base file",
                         "source file",
                         "save target"}

        Dim enumOpts = getEnumOpts()
        Dim staticOpts = 3
        Dim minPresetNum = staticOpts
        Dim maxPresetNum = minPresetNum + presetFileNames.Count - 1
        Dim minFileSelectNum = maxPresetNum + 1
        Dim maxFileSelectNum = minFileSelectNum + transmuteFiles.Count - 1
        Dim toggleOpt = maxFileSelectNum + 1
        Dim minEnumNum = toggleOpt + 1
        Dim maxEnumNum = minEnumNum + enumOpts.Count - 1
        Dim resetNum = maxEnumNum + 1

        Select Case True

            ' Exit 
            Case intInput = 0
            exitModule()

            ' Run (default)
            Case intInput = 1 OrElse input.Length = 0

                If Not denyActionWithHeader(TransmuteFile2.Name.Length = 0, "You must select a valid source file") Then initTransmute()

            ' Open Flavorizer
            Case intInput = 2

                initModule("Flavorizer", AddressOf printFlavorizerMainMenu, AddressOf handleFlavorizerMainMenuUserInput)

            ' Preset Source File Names
            Case intInput >= minPresetNum AndAlso intInput <= maxPresetNum

                Dim i = intInput - staticOpts
                TransmuteFile2.Name = presetFileNames(i)
                updateSettings(NameOf(Transmute), NameOf(TransmuteFile2) & "_Name", presetFileNames(i))
                updateSettings(NameOf(Transmute), NameOf(TransmuteModuleSettingsChanged), True.ToString)

                setNextMenuHeaderText($"Source file name set to {TransmuteFile2.Name}", printColor:=ConsoleColor.Yellow)

          ' File Selectors
            Case intInput >= minFileSelectNum AndAlso intInput <= maxFileSelectNum

                Dim i = intInput - maxPresetNum - 1
                changeFileParams(transmuteFiles(transmuteFiles.Keys(i)), TransmuteModuleSettingsChanged, NameOf(Transmute),
                                 transmuteFiles.Keys(i), NameOf(TransmuteModuleSettingsChanged), fileDescs(i))

            ' Toggles
            Case intInput = toggleOpt

                toggleModuleSetting("Winapp2.ini Syntax", NameOf(Transmute), GetType(transmuteSettings),
                                    NameOf(UseWinapp2Syntax), NameOf(TransmuteModuleSettingsChanged))

            ' Transmute Mode and sub modes
            Case intInput >= minEnumNum AndAlso intInput <= maxEnumNum

                Dim i = intInput - minEnumNum
                Dim propName = enumOpts.Keys(i)
                CycleEnumProperty(propName, enumOpts(propName), GetType(transmuteSettings), NameOf(Transmute),
                                  TransmuteModuleSettingsChanged, NameOf(TransmuteModuleSettingsChanged), ConsoleColor.Magenta)

                updateAllTransmuteEnumSettings()

            ' Reset Settings 
            Case TransmuteModuleSettingsChanged AndAlso intInput = resetNum

                resetModuleSettings(NameOf(Transmute), AddressOf initDefaultTransmuteSettings)

            Case Else

                setNextMenuHeaderText(invInpStr, printColor:=ConsoleColor.Red)

        End Select

    End Sub

    ''' <summary>
    ''' Determines the current set of enum options displayed on the menu and returns a Dictionary 
    ''' with the property names as keys and their descriptions as values
    ''' 
    ''' The set of possible enums includes:
    ''' <list type="bullet">
    '''     
    '''     <item>
    '''     Transmute mode
    '''     </item>
    '''     
    '''     <item>
    '''     Replace Mode (only available when Transmute mode is Replace)
    '''     </item>
    '''     
    '''     <item>
    '''     Remove Mode (only available when Transmute mode is Remove)
    '''     </item>
    '''     
    '''     <item>
    '''     Key Removal Mode (only available when Transmute mode is Remove and Remove mode is By Key)
    '''     </item>
    '''     
    ''' </list>
    ''' 
    ''' </summary>
    Private Function getEnumOpts() As Dictionary(Of String, String)

        Dim enumOpts As New Dictionary(Of String, String) From {{NameOf(Transmutator), "Transmute Mode"}}

        If Transmutator = TransmuteMode.Replace Then enumOpts.Add(NameOf(TransmuteReplaceMode), "Replace Mode")

        If Transmutator = TransmuteMode.Remove Then

            enumOpts.Add(NameOf(TransmuteRemoveMode), "Remove Mode")
            If TransmuteRemoveMode = RemoveMode.ByKey Then enumOpts.Add(NameOf(TransmuteRemoveKeyMode), "Key Removal Mode")

        End If

        Return enumOpts

    End Function

End Module