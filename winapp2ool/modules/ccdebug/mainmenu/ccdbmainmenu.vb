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
''' Displays the CCiniDebug main menu and handles user input accordingly
''' </summary>
Module ccdbmainmenu

    ''' <summary>
    ''' Builds and returns the CCiniDebug main menu. <br />
    ''' Both <c> printCCDBMainMenu </c> and <c> handleCCDBMUserInput </c> call this function
    ''' to ensure the menu numbering seen by the user is always in sync with dispatch.
    ''' </summary>
    Private Function buildCCDBMenu() As MenuSection

        Dim menuDesc As String() = {"Sort alphabetically the contents of the CCleaner Classic (v1-6) ccleaner.ini and prune stale winapp2.ini settings",
                                    "This module is not compatible with CCleaner 7's ccleaner.ini"}
        Dim noOptionsSelected = Not (PruneStaleEntries OrElse SaveDebuggedFile OrElse SortFileForOutput)

        Return MenuSection.CreateCompleteMenu(NameOf(CCiniDebug), menuDesc, ConsoleColor.Red) _
            .AddDispatchedColoredOption("Run (default)", "Debug ccleaner.ini", GetRedGreen(noOptionsSelected), AddressOf CheckOptsAndDebug) _
            .AddBlank() _
            .AddDispatchedToggle("Toggle pruning", "removal of orphaned winapp2.ini settings", PruneStaleEntries,
                Sub() toggleModuleSetting("Toggle pruning", NameOf(CCiniDebug), GetType(ccdebugsettings), NameOf(PruneStaleEntries), NameOf(CCDBSettingsChanged))) _
            .AddDispatchedToggle("Toggle saving", "automatic saving of changes made by CCiniDebug", SaveDebuggedFile,
                Sub() toggleModuleSetting("Toggle saving", NameOf(CCiniDebug), GetType(ccdebugsettings), NameOf(SaveDebuggedFile), NameOf(CCDBSettingsChanged))) _
            .AddDispatchedToggle("Toggle sorting", "alphabetical sorting of the contents of ccleaner.ini", SortFileForOutput,
                Sub() toggleModuleSetting("Toggle sorting", NameOf(CCiniDebug), GetType(ccdebugsettings), NameOf(SortFileForOutput), NameOf(CCDBSettingsChanged))) _
            .AddBlank() _
            .AddDispatchedOption("Choose winapp2.ini", "Select a new supplemental winapp2.ini file", condition:=PruneStaleEntries,
                handler:=Sub() changeFile2Params(CCDebugFile1, CCDBSettingsChanged, NameOf(CCiniDebug), NameOf(CCDebugFile1), NameOf(CCDBSettingsChanged))) _
            .AddDispatchedOption("Choose ccleaner.ini", "Select a new ccleaner.ini file for debugging",
                Sub() changeFile2Params(CCDebugFile2, CCDBSettingsChanged, NameOf(CCiniDebug), NameOf(CCDebugFile2), NameOf(CCDBSettingsChanged))) _
            .AddDispatchedOption("Choose save target", "Select a new save target for the debugged ccleaner.ini", condition:=SaveDebuggedFile,
                handler:=Sub() changeFile2Params(CCDebugFile3, CCDBSettingsChanged, NameOf(CCiniDebug), NameOf(CCDebugFile3), NameOf(CCDBSettingsChanged))) _
            .AddBlank() _
            .AddLine($"Current winapp2.ini:   {replDir(CCDebugFile1.Path())}", condition:=PruneStaleEntries) _
            .AddLine($"Current ccleaner.ini:  {replDir(CCDebugFile2.Path())}") _
            .AddLine($"Current save target:   {replDir(CCDebugFile3.Path())}", condition:=SaveDebuggedFile) _
            .AddBlank(CCDBSettingsChanged) _
            .AddDispatchedResetOpt(NameOf(CCiniDebug), CCDBSettingsChanged, Sub() resetModuleSettings(NameOf(CCiniDebug), AddressOf initDefaultCCDBSettings))

    End Function

    ''' <summary>
    ''' Prints the CCiniDebug menu to the user
    ''' </summary>
    Public Sub printCCDBMainMenu()

        buildCCDBMenu().Print()

    End Sub

    ''' <summary>
    ''' Handles the user's input from the CCiniDebug main menu
    ''' </summary>
    '''
    ''' <param name="input">
    ''' The user's input
    ''' </param>
    Public Sub handleCCDBMUserInput(input As String)

        Dim intInput As Integer

        If Not Integer.TryParse(input, intInput) Then

            If input.Length = 0 Then

                CheckOptsAndDebug()
                Return

            End If

            setNextMenuHeaderText(invInpStr, printColor:=ConsoleColor.Red)
            Return

        End If

        If intInput = 0 Then exitModule() : Return

        If Not buildCCDBMenu().Dispatch(intInput) Then setNextMenuHeaderText(invInpStr, printColor:=ConsoleColor.Red)

    End Sub

    ''' <summary>
    ''' Ensures that at least one option has been selected and kicks off the debugger
    ''' </summary>
    Private Sub CheckOptsAndDebug()

        Dim noOpts = Not (PruneStaleEntries OrElse SaveDebuggedFile OrElse SortFileForOutput)
        If Not denyActionWithHeader(noOpts, "Please enable at least one option") Then initCCDebug()

    End Sub

End Module
