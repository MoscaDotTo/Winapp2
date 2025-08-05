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
''' Docs last updated: 2025-08-01 | Code last updated: 2025-08-01
Public Module FlavorizerMainMenu

    ''' <summary>
    ''' Prints the Flavorizer menu to the user, showing all correction file options
    ''' and current settings
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-08-01 | Code last updated: 2025-08-05
    Public Sub printFlavorizerMainMenu()

        Console.WindowHeight = Console.LargestWindowHeight

        printMenuTop({"Apply a 'flavor' to an ini file using correction files",
                      "Flavorization applies corrections in this order:",
                      "Section Removal → Key Name Removal → Key Value Removal → Section Replacement → Key Replacement → Additions",
                      "All correction files are optional - specify only the ones you need"})

        print(1, "Run (default)", "Apply the flavor to the base file", enStrCond:=FlavorizerFile1.Name.Length > 0, colorLine:=True)

        print(0, "Base and save files:", leadingBlank:=True, trailingBlank:=True)
        print(1, "File Chooser (base)", "Choose the base file to be flavorized")
        print(1, "File Chooser (save)", "Choose where to save the flavorized result")

        print(0, "Correction files (applied in order):", leadingBlank:=True)
        print(1, "Section Removal", "Choose file containing sections to remove entirely")
        print(1, "Key Name Removal", "Choose file containing keys to remove by name")
        print(1, "Key Value Removal", "Choose file containing keys to remove by value")
        print(1, "Section Replacement", "Choose file containing sections to replace entirely")
        print(1, "Key Replacement", "Choose file containing keys to replace by name")
        print(1, "Additions", "Choose file containing sections/keys to add", trailingBlank:=True)

        print(0, "Current Files:")
        print(0, $"Base file: {replDir(FlavorizerFile1.Path)}")
        print(0, $"Save file: {replDir(FlavorizerFile2.Path)}")
        print(0, $"Section removal: {If(FlavorizerFile3.Name.Length = 0, "Not specified", replDir(FlavorizerFile3.Path))}", enStrCond:=FlavorizerFile3.Name.Length > 0)
        print(0, $"Key name removal: {If(FlavorizerFile4.Name.Length = 0, "Not specified", replDir(FlavorizerFile4.Path))}", enStrCond:=FlavorizerFile4.Name.Length > 0)
        print(0, $"Key value removal: {If(FlavorizerFile5.Name.Length = 0, "Not specified", replDir(FlavorizerFile5.Path))}", enStrCond:=FlavorizerFile5.Name.Length > 0)
        print(0, $"Section replacement: {If(FlavorizerFile6.Name.Length = 0, "Not specified", replDir(FlavorizerFile6.Path))}", enStrCond:=FlavorizerFile6.Name.Length > 0)
        print(0, $"Key replacement: {If(FlavorizerFile7.Name.Length = 0, "Not specified", replDir(FlavorizerFile7.Path))}", enStrCond:=FlavorizerFile7.Name.Length > 0)
        print(0, $"Additions: {If(FlavorizerFile8.Name.Length = 0, "Not specified", replDir(FlavorizerFile8.Path))}", enStrCond:=FlavorizerFile8.Name.Length > 0, trailingBlank:=True)

        print(1, "winapp2.ini formatting", "Toggle formatting output as winapp2.ini")
        print(0, $"winapp format: {FlavorizeAsWinapp}", leadingBlank:=True, colorLine:=True, enStrCond:=FlavorizeAsWinapp)

        print(1, "Auto detect Flavor", "Attempt to detect the flavor files by name", leadingBlank:=True)
        print(1, "Change target directory", "Change the directory within which to scan for Flavor files")
        print(0, $"Current target directory: {FlavorizerFile9.Dir}", closeMenu:=Not FlavorizerModuleSettingsChanged)

        print(2, NameOf(Flavorizer), cond:=FlavorizerModuleSettingsChanged, closeMenu:=True)

    End Sub

    ''' <summary>
    ''' Handles the user's input from the main menu
    ''' </summary>
    ''' 
    ''' <param name="input">
    ''' The user's input
    ''' </param>
    ''' 
    ''' Docs last updated: 2025-08-01 | Code last updated: 2025-08-01
    Public Sub handleFlavorizerMainMenuUserInput(input As String)

        Dim flavorFiles As New Dictionary(Of String, iniFile) From {
            {NameOf(FlavorizerFile1), FlavorizerFile1},
            {NameOf(FlavorizerFile2), FlavorizerFile2},
            {NameOf(FlavorizerFile3), FlavorizerFile3},
            {NameOf(FlavorizerFile4), FlavorizerFile4},
            {NameOf(FlavorizerFile5), FlavorizerFile5},
            {NameOf(FlavorizerFile6), FlavorizerFile6},
            {NameOf(FlavorizerFile7), FlavorizerFile7},
            {NameOf(FlavorizerFile8), FlavorizerFile8}
        }

        Dim fileOpts = {"2", "3", "4", "5", "6", "7", "8", "9"}

        Select Case True

            Case input = "0"

                exitModule()

            Case input = "1" OrElse input.Length = 0

                If Not denyActionWithHeader(FlavorizerFile1.Name.Length = 0, "You must select a base file") Then initFlavorizer()

            Case fileOpts.Contains(input)

                Dim i = CType(input, Integer) - 2
                changeFileParams(flavorFiles(flavorFiles.Keys(i)), FlavorizerModuleSettingsChanged, NameOf(Flavorizer),
                                 flavorFiles.Keys(i), NameOf(FlavorizerModuleSettingsChanged))

            Case input = "10"

                toggleSettingParam(FlavorizeAsWinapp, "Winapp2.ini formatting", FlavorizerModuleSettingsChanged,
                                   NameOf(Flavorizer), NameOf(FlavorizeAsWinapp), NameOf(FlavorizerModuleSettingsChanged))

            Case input = "11"

                DetectFlavorFiles(FlavorizerFile9.Dir)

            Case input = "12"

                changeFileParams(FlavorizerFile9, FlavorizerModuleSettingsChanged, NameOf(Flavorizer),
                                  NameOf(FlavorizerFile9), NameOf(FlavorizerModuleSettingsChanged))

                ' This file is used only to hold the Target directory and never holds a name 
                FlavorizerFile9.Name = ""
                updateSettings(NameOf(Flavorizer), $"{NameOf(FlavorizerFile9)}_Name", "")

            Case input = "13" AndAlso FlavorizerModuleSettingsChanged

                resetModuleSettings(NameOf(Flavorizer), AddressOf initDefaultFlavorizerSettings)

            Case Else

                setHeaderText(invInpStr, True)

        End Select

    End Sub

End Module