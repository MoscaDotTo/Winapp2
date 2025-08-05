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


        printMenuTop({"Transmute an ini file with sections and keys provided by a second ini file",
                      "Features by mode:",
                      "Add entire sections or individual keys to specific sections",
                      "Replace entire sections or individual keys within specific sections by Value",
                      "Remove entire sections or individual keys within specific sections by Name or Value"})

        print(1, "Run (default)", "Perform the transmutation", enStrCond:=Not TransmuteFile2.Name.Length = 0, colorLine:=True)

        print(1, "Open Flavorizer", "Apply a complex series of modifications to an ini file", leadingBlank:=True)

        print(0, "Preset source File Choices:", leadingBlank:=True, trailingBlank:=True)
        print(1, "Removed Entries", "Select 'Removed Entries.ini'")
        print(1, "Custom", "Select 'Custom.ini'")
        print(1, "winapp3.ini", "Select 'winapp3.ini'", trailingBlank:=True)

        print(1, "File Chooser (base)", "Choose a new name or location for the base file")
        print(1, "File Chooser (source)", "Choose a name or location for the source file")
        print(1, "File Chooser (save)", "Choose a new save location for the output of the transmutation", trailingBlank:=True)

        print(0, $"Current base file: {replDir(TransmuteFile1.Path)}")
        print(0, $"Current source file : {If(TransmuteFile2.Name.Length = 0, "Not yet selected", replDir(TransmuteFile2.Path))}", enStrCond:=Not TransmuteFile2.Name.Length = 0, colorLine:=True)
        print(0, $"Current save target: {replDir(TransmuteFile3.Path)}", trailingBlank:=True)

        Dim modeColor = If(Transmutator = TransmuteMode.Add, ConsoleColor.Green, If(Transmutator = TransmuteMode.Replace, ConsoleColor.Yellow, ConsoleColor.Red))
        print(1, "Change transmute mode", "Cycle through primary transmute modes (Add/Replace/Remove)")
        print(0, $"Current mode: {Transmutator}", leadingBlank:=True, colorLine:=True, useArbitraryColor:=True, arbitraryColor:=modeColor, closeMenu:=TransmuteModeIsAdd AndAlso Not TransmuteModuleSettingsChanged)

        Dim subModeColor = If(TransmuteModeIsReplace, ConsoleColor.Magenta, If(TransmuteRemoveMode = RemoveMode.BySection, ConsoleColor.Magenta, ConsoleColor.Cyan))
        Dim replaceModeDesc = If(ReplaceModeIsByKey, "Replace individual key values in within particular sections in the base file ", "Replace entire sections in the base file")
        print(1, "Change Replace mode", "Toggle between replacing by section or by key", TransmuteModeIsReplace)
        print(0, $"Current replace mode: {TransmuteReplaceMode}: {replaceModeDesc}", cond:=TransmuteModeIsReplace, leadingBlank:=True, colorLine:=True, useArbitraryColor:=True, arbitraryColor:=subModeColor)

        Dim removeModeDesc = If(RemoveModeIsByKey, "Remove individual keys in within particular sections in the base file ", "Remove entire sections in the base file")
        print(1, "Change Remove mode", "Toggle between removing by section or by key", TransmuteModeIsRemove)
        print(0, $"Current remove mode: {TransmuteRemoveMode}: {removeModeDesc}", cond:=TransmuteModeIsRemove, leadingBlank:=True, colorLine:=True, useArbitraryColor:=True, arbitraryColor:=subModeColor)

        Dim keyModeColor = If(TransmuteRemoveKeyMode = RemoveKeyMode.ByName, ConsoleColor.Magenta, ConsoleColor.DarkYellow)
        print(1, "Change Key Removal mode", "Toggle between removing keys by name or by value", TransmuteModeIsRemove AndAlso RemoveModeIsByKey)
        print(0, $"Current remove key mode: {TransmuteRemoveKeyMode}", cond:=TransmuteModeIsRemove AndAlso RemoveModeIsByKey, leadingBlank:=True, colorLine:=True, useArbitraryColor:=True, arbitraryColor:=keyModeColor)

        print(2, NameOf(Transmute), cond:=TransmuteModuleSettingsChanged, closeMenu:=True)

    End Sub


    ''' <summary>
    ''' Once per run of the module, we'll set the console height to be large enough so as to 
    ''' be able to display the largest number of options that we display 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-07-15 | Code last updated: 2025-07-15
    Private Sub adjustTransmuteConsoleHeight()

        Dim baseHeight = Console.WindowHeight

        If baseHeight > 42 Then Return

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
    Public Sub handleTransmuteMainMenuUserInput(input As String)

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

            ' Option Name:                                 Preset File Name: Removed Entries 
            ' Option States:
            ' Default                                      -> 3 (default)
            Case input = "3"

                changeBaseFileName("Removed entries.ini")

            ' Option Name:                                 Preset File Name: Custom
            ' Option States:
            ' Default                                      -> 4 (default)
            Case input = "4"

                changeBaseFileName("Custom.ini")

            ' Option Name:                                 Preset File Name: winapp3.ini
            ' Option States:
            ' Default                                      -> 5 (default)
            Case input = "5"

                changeBaseFileName("winapp3.ini")

            ' Option Name:                                 File Chooser (base)
            ' Option States:
            ' Default                                      -> 6 (default)

            Case input = "6"

                changeFileParams(TransmuteFile1, TransmuteModuleSettingsChanged,
                                 NameOf(Transmute), NameOf(TransmuteFile1), NameOf(TransmuteModuleSettingsChanged))

            ' Option Name:                                 File Chooser (source)
            ' Option States:
            ' Default                                      -> 7 (default)
            Case input = "7"

                changeFileParams(TransmuteFile2, TransmuteModuleSettingsChanged,
                                 NameOf(Transmute), NameOf(TransmuteFile2), NameOf(TransmuteModuleSettingsChanged))

            ' Option Name:                                 File Chooser (save)
            ' Option States:
            ' Default                                      -> 8 (default)
            Case input = "8"

                changeFileParams(TransmuteFile3, TransmuteModuleSettingsChanged,
                                 NameOf(Transmute), NameOf(TransmuteFile3), NameOf(TransmuteModuleSettingsChanged))

            ' Option Name:                                 Change Transmute Mode 
            ' Option States:
            ' Default                                      -> 9 (default)
            Case input = "9"

                cycleEnumSetting(Transmutator, GetType(TransmuteMode), "Transmutator",
                                 TransmuteModuleSettingsChanged, NameOf(Transmute), NameOf(TransmuteModuleSettingsChanged))

                updateAllTransmuteEnumSettings()

            ' Option Name:                                 Change Replace Mode
            ' Option States:
            ' TransmuteModeIsReplace = False               -> Unavailable (not displayed)
            ' Default                                      -> 10 (default)
            Case input = "10" And Transmutator = TransmuteMode.Replace

                cycleEnumSetting(TransmuteReplaceMode, GetType(ReplaceMode), "Replace Mode",
                                 TransmuteModuleSettingsChanged, NameOf(Transmute), NameOf(TransmuteModuleSettingsChanged))

                updateAllTransmuteEnumSettings()

            ' Option Name:                                 Change Remove Mode
            ' Option States:
            ' TransmuteModeIsRemove = False                -> Unavailable (not displayed)
            ' Default                                      -> 10 (default)
            Case input = "10" And Transmutator = TransmuteMode.Remove

                cycleEnumSetting(TransmuteRemoveMode, GetType(RemoveMode), "Remove Mode",
                                 TransmuteModuleSettingsChanged, NameOf(Transmute), NameOf(TransmuteModuleSettingsChanged))

                updateAllTransmuteEnumSettings()

            ' Option Name:                                 Change Remove Key Mode
            ' Option States:
            ' RemoveModeIsByKey = False                    -> Unavailable (not displayed)
            ' Default                                      -> 11 (default)
            Case input = "11" And Transmutator = TransmuteMode.Remove And TransmuteRemoveMode = RemoveMode.ByKey

                cycleEnumSetting(TransmuteRemoveKeyMode, GetType(RemoveKeyMode), "Remove Key Mode",
                                 TransmuteModuleSettingsChanged, NameOf(Transmute), NameOf(TransmuteModuleSettingsChanged))

                updateAllTransmuteEnumSettings()

            ' Option Name:                                  Reset Settings 
            ' Option States: 
            ' ModuleSettingsChanged = False                 -> Unavailable (not displayed) 
            ' TransmuteModeIsAdd                            -> 10 (default)
            ' TransmuteModeIsReplace (+1)                   -> 11  
            ' TransmuteModeIsRemove (+1) and BySection (+0) -> 11
            ' TransmuteModeIsRemove (+1) and ByKey (+1)     -> 12
            Case TransmuteModuleSettingsChanged AndAlso
                 input = computeMenuNumber(10, {TransmuteModeIsAdd, TransmuteModeIsReplace,
                                               TransmuteModeIsRemove AndAlso RemoveModeIsBySection,
                                               TransmuteModeIsRemove AndAlso RemoveModeIsByKey}, {0, 1, 1, 2})

                resetModuleSettings(NameOf(Transmute), AddressOf initDefaultTransmuteSettings)

            Case Else

                setHeaderText(invInpStr, True)

        End Select

    End Sub

End Module