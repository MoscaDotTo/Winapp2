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
''' Prints the main menu for the WinappDebug module to the user and handles user input
''' </summary>
''' 
''' Docs last updated: 2023-07-20 | Code last updated: 2023-07-20
Module lintmainmenu

    ''' <summary>
    ''' Displays the main <c> WinappDebug </c> menu to the user 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2023-07-20 | Code last updated: 2023-07-20
    Public Sub printLintMainMenu()

        printMenuTop({"Scan winapp2.ini for style and syntax errors, and attempt to repair them where possible."})
        print(1, "Run (Default)", "Run the debugger")
        print(1, "File Chooser (winapp2.ini)", "Choose a different file name or path for winapp2.ini", leadingBlank:=True, trailingBlank:=True)
        print(5, "Toggle Saving", "saving the file after correcting errors", enStrCond:=SaveChanges)
        print(1, "File Chooser (save)", "Save a copy of changes made to a new file instead of overwriting winapp2.ini", SaveChanges, trailingBlank:=True)
        print(1, "Toggle Scan Settings", "Enable or disable individual scan and correction routines", leadingBlank:=Not SaveChanges, trailingBlank:=True)
        print(5, "Toggle Default Value Audit", "enforcing a specific value for Default keys", enStrCond:=overrideDefaultVal, trailingBlank:=Not overrideDefaultVal)
        print(1, "Toggle Expected Default", $"Currently enforcing that Default keys have a value of: {expectedDefaultValue}", cond:=overrideDefaultVal, trailingBlank:=True)
        print(0, $"Current winapp2.ini:  {replDir(winappDebugFile1.Path)}", closeMenu:=Not SaveChanges And Not LintModuleSettingsChanged And MostRecentLintLog.Length = 0)
        print(0, $"Current save target:  {replDir(winappDebugFile3.Path)}", cond:=SaveChanges, closeMenu:=Not LintModuleSettingsChanged And MostRecentLintLog.Length = 0)
        print(2, NameOf(WinappDebug), cond:=LintModuleSettingsChanged, closeMenu:=MostRecentLintLog.Length = 0)
        print(1, "Log Viewer", "Show the most recent lint results", cond:=Not MostRecentLintLog.Length = 0, closeMenu:=True, leadingBlank:=True)

    End Sub

    ''' <summary> 
    ''' Handles the user input for the <c> WinappDebug </c> main menu
    ''' </summary>
    ''' 
    ''' <param name="input">
    ''' The user's input 
    ''' </param>
    ''' 
    ''' Docs last updated: 2023-07-20 | Code last updated: 2022-12-04
    Public Sub handleLintUserInput(input As String)

        If input Is Nothing Then argIsNull(NameOf(input)) : Return

        Dim saveXorOverride = SaveChanges Xor overrideDefaultVal
        Dim saveAndOverride = SaveChanges AndAlso overrideDefaultVal

        Select Case True

            ' Option Name:                                 Exit 
            ' Option states:
            ' Default                                      -> 0 (default)
            Case input = "0"

                exitModule()

            ' Option Name:                                 Run (default)
            ' Option states:
            ' Default                                      -> 1 (default)
            Case input = "1" Or input.Length = 0

                initDebug()

            ' Option Name:                                 File Chooser (winapp2.ini)
            ' Option states:
            ' Default                                      -> 2 (default)
            Case input = "2"

                changeFileParams(winappDebugFile1, LintModuleSettingsChanged, NameOf(WinappDebug), NameOf(winappDebugFile1), NameOf(LintModuleSettingsChanged))

            ' Option Name:                                Toggle Saving 
            ' Option states:
            ' Default                                      -> 3 (default)
            Case input = "3"

                toggleSettingParam(SaveChanges, "Saving", LintModuleSettingsChanged, NameOf(WinappDebug), NameOf(SaveChanges), NameOf(LintModuleSettingsChanged))

            ' Option Name:                                 File Chooser (save)
            ' Option states:
            ' Not saving changes                           -> Unavailable (not displayed)
            ' Saving changes                               -> 4 (default)
            Case input = "4" AndAlso SaveChanges

                changeFileParams(winappDebugFile3, LintModuleSettingsChanged, NameOf(WinappDebug), NameOf(winappDebugFile3), NameOf(LintModuleSettingsChanged))

            ' Option Name:                                 Toggle Scan Settings
            ' Option states:
            ' Not saving                                   -> 4 (default)
            ' Saving (+1)                                  -> 5
            Case input = computeMenuNumber(4, {SaveChanges}, {1})

                initModule("Scan Settings", AddressOf advSettings.printMenu, AddressOf advSettings.handleUserInput)
                Console.WindowHeight = 30

            ' Option Name:                                  Toggle Default Value Audit 
            ' Option states:
            ' Not saving                                   -> 5 (default)
            ' Saving (+1)                                  -> 6
            Case input = computeMenuNumber(5, {SaveChanges}, {1})


                toggleSettingParam(overrideDefaultVal, "Default Value Overriding", LintModuleSettingsChanged, NameOf(WinappDebug), NameOf(overrideDefaultVal), NameOf(LintModuleSettingsChanged))

            ' Option Name:                           Reset Settings       
            ' Option states:
            ' Module Settings not changed                  -> Unavailable (not displayed)
            ' Not Saving, Not auditing defaults            -> 6 (default)
            ' Saving (+1), not auditing                    -> 7
            ' Not saving, auditing (+1)                    -> 7
            ' Saving and auditing                          -> 8 
            Case LintModuleSettingsChanged And input = computeMenuNumber(6, {SaveChanges, overrideDefaultVal}, {1, 1})

                resetModuleSettings("WinappDebug", AddressOf InitDefaultLintSettings)

            ' Option Name:                                 Log Viewer
            ' Option states:
            ' No log exists                                -> Unavailable (not displayed) 
            ' No settings changes, not saving, no audit    -> 6 (default) 
            ' Settings changed(+1), not saving, no audit   -> 7
            ' Settings changed(+1), saving(+1), no audit   -> 8
            ' Settings changed(+1), not saving, audit(+1)  -> 8 
            ' Settings changed(+1), saving(+1), audit(+1)  -> 9
            Case Not MostRecentLintLog.Length = 0 AndAlso input = computeMenuNumber(6, {LintModuleSettingsChanged, SaveChanges, overrideDefaultVal}, {1, 1, 1})
                printSlice(MostRecentLintLog)

            ' Option Name:                                 Toggle Expected Default
            ' Option states:
            ' Not auditing                                 -> Unavailable (not displayed)
            ' Not saving                                   -> 6
            ' Saving (+1)                                  -> 7
            Case overrideDefaultVal And input = computeMenuNumber(6, {SaveChanges}, {1})

                toggleSettingParam(expectedDefaultValue, "Expected Default Value", LintModuleSettingsChanged, NameOf(WinappDebug), NameOf(expectedDefaultValue), NameOf(LintModuleSettingsChanged))

            Case Else

                setHeaderText(invInpStr, True)

        End Select

    End Sub

End Module