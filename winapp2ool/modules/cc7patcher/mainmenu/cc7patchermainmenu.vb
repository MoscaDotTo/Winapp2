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
''' Displays the CC7Patcher module main menu and handles user input.
''' Called by both <c> printCC7PatcherMenu </c> (to render) and <c> handleCC7PatcherInput </c>
''' (to dispatch), so the displayed option numbers and the dispatch table are always in sync.
''' </summary>
'''
''' Docs last updated: 2026-03-12 | Code last updated: 2026-03-12
Public Module cc7patchermainmenu

    ''' <summary>
    ''' Builds the CC7Patcher main menu with all options and their dispatch handlers registered inline
    ''' </summary>
    '''
    ''' Docs last updated: 2026-03-12 | Code last updated: 2026-03-12
    Private Function buildCC7PatcherMenu() As MenuSection

        Dim menuDescLines = {"Patch ccleaner.ini with winapp2.ini entries compatible with CCleaner 7"}

        Return MenuSection.CreateCompleteMenu(NameOf(CC7Patcher), menuDescLines, ConsoleColor.Yellow) _
            .AddDispatchedColoredOption("Run (default)", "Install winapp2.ini for CCleaner 7", ConsoleColor.Yellow,
                Sub() initCC7Patcher()) _
            .AddBlank() _
            .AddDispatchedToggle("Trim", "trimming winapp2.ini before installation", TrimBeforePatching,
                Sub() toggleModuleSetting("Trim", NameOf(CC7Patcher), GetType(cc7patchersettings),
                                          NameOf(TrimBeforePatching), NameOf(CC7PatcherModuleSettingsChanged))) _
            .AddDispatchedToggle("Download", "downloading the latest winapp2.ini from GitHub", DownloadWinapp2,
                Sub() toggleModuleSetting("Download", NameOf(CC7Patcher), GetType(cc7patchersettings),
                                          NameOf(DownloadWinapp2), NameOf(CC7PatcherModuleSettingsChanged)),
                Not isOffline) _
            .AddBlank() _
            .AddDispatchedOption("Change winapp2.ini", "Select the winapp2.ini file to install",
                Sub() changeFile2Params(CC7PatcherFile1, CC7PatcherModuleSettingsChanged, NameOf(CC7Patcher),
                                        NameOf(CC7PatcherFile1), NameOf(CC7PatcherModuleSettingsChanged)),
                Not DownloadWinapp2) _
            .AddDispatchedOption("Change ccleaner.ini", "Select the ccleaner.ini file to be patched",
                Sub() changeFile2Params(CC7PatcherFile2, CC7PatcherModuleSettingsChanged, NameOf(CC7Patcher),
                                        NameOf(CC7PatcherFile2), NameOf(CC7PatcherModuleSettingsChanged))) _
            .AddDispatchedOption("Change output file", "Select where to save the patched ccleaner.ini",
                Sub() changeFile2Params(CC7PatcherFile3, CC7PatcherModuleSettingsChanged, NameOf(CC7Patcher),
                                        NameOf(CC7PatcherFile3), NameOf(CC7PatcherModuleSettingsChanged))) _
            .AddBlank() _
            .AddColoredFileInfo("Current winapp2.ini:  ", If(DownloadWinapp2, "Online", CC7PatcherFile1.Path()), ConsoleColor.Green) _
            .AddColoredFileInfo("Current ccleaner.ini: ", CC7PatcherFile2.Path(), ConsoleColor.Red) _
            .AddColoredFileInfo("Output file:          ", CC7PatcherFile3.Path(), ConsoleColor.Cyan) _
            .AddBlank(CC7PatcherModuleSettingsChanged) _
            .AddDispatchedResetOpt(NameOf(CC7Patcher), CC7PatcherModuleSettingsChanged,
                Sub() resetModuleSettings(NameOf(CC7Patcher), AddressOf InitDefaultCC7PatcherSettings))

    End Function

    ''' <summary>
    ''' Prints the CC7Patcher menu to the user
    ''' </summary>
    '''
    ''' Docs last updated: 2026-03-12 | Code last updated: 2026-03-12
    Public Sub printCC7PatcherMenu()

        buildCC7PatcherMenu().Print()

    End Sub

    ''' <summary>
    ''' Handles user input for the CC7Patcher menu
    ''' </summary>
    '''
    ''' <param name="input">
    ''' The user's input
    ''' </param>
    '''
    ''' Docs last updated: 2026-03-12 | Code last updated: 2026-03-12
    Public Sub handleCC7PatcherInput(input As String)

        Dim intInput As Integer

        If Not Integer.TryParse(input, intInput) Then

            setNextMenuHeaderText(invInpStr, printColor:=ConsoleColor.Red)
            Return

        End If

        If intInput = 0 Then exitModule() : Return

        If Not buildCC7PatcherMenu().Dispatch(intInput) Then setNextMenuHeaderText(invInpStr, printColor:=ConsoleColor.Red)

    End Sub

End Module
