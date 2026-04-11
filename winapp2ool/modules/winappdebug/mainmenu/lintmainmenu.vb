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
''' Prints the main menu for the WinappDebug module to the user and handles user input
''' </summary>
Module lintmainmenu

    ''' <summary>
    ''' Builds the WinappDebug main menu with all options and their dispatch handlers registered inline.
    ''' Called by both <c> printLintMainMenu </c> (to render) and <c> handleLintUserInput </c>
    ''' (to dispatch), so the displayed option numbers and the dispatch table are always in sync.
    ''' </summary>
    Private Function buildLintMenu() As MenuSection

        Return MenuSection.CreateCompleteMenu("WinappDebug",
            {"Scan winapp2.ini for style and syntax errors, and attempt to repair them where possible."},
            ConsoleColor.DarkCyan) _
            .AddDispatchedOption("Run (Default)", "Run the debugger", Sub() InitDebug()) _
            .AddBlank() _
            .AddDispatchedOption("File Chooser (winapp2.ini)", "Choose a different file name or path for winapp2.ini",
                Sub() changeFile2Params(winappDebugFile1, LintModuleSettingsChanged, NameOf(WinappDebug), NameOf(winappDebugFile1), NameOf(LintModuleSettingsChanged))) _
            .AddBlank() _
            .AddDispatchedToggle("Saving", "saving the file after correcting errors", SaveChanges,
                Sub() toggleSettingParam(SaveChanges, "Saving", LintModuleSettingsChanged, NameOf(WinappDebug), NameOf(SaveChanges), NameOf(LintModuleSettingsChanged))) _
            .AddDispatchedOption("File Chooser (save)", "Save a copy of changes made to a new file instead of overwriting winapp2.ini", condition:=SaveChanges,
                handler:=Sub() changeFile2Params(winappDebugFile3, LintModuleSettingsChanged, NameOf(WinappDebug), NameOf(winappDebugFile3), NameOf(LintModuleSettingsChanged))) _
            .AddBlank() _
            .AddDispatchedOption("Scan Settings", "Enable or disable individual scan and correction routines",
                Sub()
                    initModule("Scan Settings", AddressOf advSettings.printMenu, AddressOf advSettings.handleUserInput)
                    Console.WindowHeight = 30
                End Sub) _
            .AddBlank() _
            .AddDispatchedToggle("Default Value Audit", "enforcing a specific value for Default keys", overrideDefaultVal,
                Sub() toggleSettingParam(overrideDefaultVal, "Default Value Overriding", LintModuleSettingsChanged, NameOf(WinappDebug), NameOf(overrideDefaultVal), NameOf(LintModuleSettingsChanged))) _
            .AddDispatchedOption("Toggle Expected Default", $"Currently enforcing that Default keys have a value of: {expectedDefaultValue}",
                Sub() toggleSettingParam(expectedDefaultValue, "Expected Default Value", LintModuleSettingsChanged, NameOf(WinappDebug), NameOf(expectedDefaultValue), NameOf(LintModuleSettingsChanged)),
                overrideDefaultVal) _
            .AddBlank() _
            .AddFileInfo("Current winapp2.ini:  ", winappDebugFile1.Path) _
            .AddFileInfo("Current save target:  ", winappDebugFile3.Path, condition:=SaveChanges) _
            .AddBlank(LintModuleSettingsChanged) _
            .AddDispatchedResetOpt("WinappDebug", LintModuleSettingsChanged,
                Sub() resetModuleSettings("WinappDebug", AddressOf InitDefaultLintSettings)) _
            .AddBlank(MostRecentLintLog.Length > 0) _
            .AddDispatchedColoredOption("Log Viewer", "Show the most recent lint results", ConsoleColor.Yellow,
                Sub() printSlice(MostRecentLintLog.ToString()),
                MostRecentLintLog.Length > 0)

    End Function

    ''' <summary>
    ''' Displays the main <c> WinappDebug </c> menu to the user
    ''' </summary>
    Public Sub printLintMainMenu()

        buildLintMenu().Print()

    End Sub

    ''' <summary>
    ''' Handles the user input for the <c> WinappDebug </c> main menu
    ''' </summary>
    '''
    ''' <param name="input">
    ''' The user's input
    ''' </param>
    Public Sub handleLintUserInput(input As String)

        If input Is Nothing Then argIsNull(NameOf(input)) : Return

        Dim intInput As Integer

        If Not Integer.TryParse(input, intInput) Then
            If input.Length = 0 Then initDebug() : Return
            setNextMenuHeaderText(invInpStr, printColor:=ConsoleColor.Red)
            Return
        End If

        If intInput = 0 Then exitModule() : Return

        If Not buildLintMenu().Dispatch(intInput) Then setNextMenuHeaderText(invInpStr, printColor:=ConsoleColor.Red)

    End Sub

End Module
