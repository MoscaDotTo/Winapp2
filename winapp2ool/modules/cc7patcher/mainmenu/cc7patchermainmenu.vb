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
Public Module cc7patchermainmenu

    ''' <summary>
    ''' Builds the CC7Patcher main menu with all options and their dispatch handlers registered inline
    ''' </summary>
    Private Function buildCC7PatcherMenu() As MenuSection

        Dim menuDescLines = {"Patch ccleaner.ini with winapp2.ini entries compatible with CCleaner 7"}
        Dim settingsChangeName = NameOf(CC7PatcherModuleSettingsChanged)
        Dim moduleName = NameOf(CC7Patcher)

        Return MenuSection.CreateCompleteMenu(moduleName, menuDescLines, ConsoleColor.Yellow) _
            .AddDispatchedColoredOption("Run (default)", "Install winapp2.ini for CCleaner 7", ConsoleColor.Yellow,
                Sub() initCC7Patcher()) _
            .AddBlank() _
            .AddDispatchedToggle("Trim", "trimming winapp2.ini before installation", TrimBeforePatching,
                Sub() toggleModuleSetting("Trim", moduleName, GetType(cc7patchersettings), NameOf(TrimBeforePatching), settingsChangeName)) _
            .AddDispatchedToggle("Download", "downloading the latest winapp2.ini from GitHub", DownloadWinapp2, condition:=Not isOffline,
                handler:=Sub() toggleModuleSetting("Download", moduleName, GetType(cc7patchersettings), NameOf(DownloadWinapp2), settingsChangeName)) _
            .AddBlank() _
            .AddDispatchedOption("Change winapp2.ini", "Select the winapp2.ini file to install", condition:=Not DownloadWinapp2,
                handler:=Sub() changeFile2Params(CC7PatcherFile1, CC7PatcherModuleSettingsChanged, moduleName, NameOf(CC7PatcherFile1), settingsChangeName)) _
            .AddDispatchedOption("Change ccleaner.ini", "Select the ccleaner.ini file to be patched",
                Sub() changeFile2Params(CC7PatcherFile2, CC7PatcherModuleSettingsChanged, moduleName, NameOf(CC7PatcherFile2), settingsChangeName)) _
            .AddDispatchedOption("Change output file", "Select where to save the patched ccleaner.ini",
                Sub() changeFile2Params(CC7PatcherFile3, CC7PatcherModuleSettingsChanged, moduleName, NameOf(CC7PatcherFile3), settingsChangeName)) _
            .AddBlank() _
            .AddColoredFileInfo("Current winapp2.ini:  ", If(DownloadWinapp2, "Online", CC7PatcherFile1.Path()), ConsoleColor.Green) _
            .AddColoredFileInfo("Current ccleaner.ini: ", CC7PatcherFile2.Path(), ConsoleColor.Red) _
            .AddColoredFileInfo("Output file:          ", CC7PatcherFile3.Path(), ConsoleColor.Cyan) _
            .AddBlank(CC7PatcherModuleSettingsChanged) _
            .AddDispatchedResetOpt(moduleName, CC7PatcherModuleSettingsChanged,
                Sub() resetModuleSettings(moduleName, AddressOf InitDefaultCC7PatcherSettings))

    End Function

    ''' <summary>
    ''' Prints the CC7Patcher menu to the user
    ''' </summary>
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
    Public Sub handleCC7PatcherInput(input As String)

        Dim intInput As Integer

        If Not Integer.TryParse(input, intInput) Then

            If input.Length = 0 Then

                If Not denyActionWithHeader(Not CC7PatcherFile2.Exists(), "CCleaner.ini not found") Then initCC7Patcher()
                Return

            End If

            setNextMenuHeaderText(invInpStr, printColor:=ConsoleColor.Red)
            Return

        End If

        If intInput = 0 Then exitModule() : Return

        If Not buildCC7PatcherMenu().Dispatch(intInput) Then setNextMenuHeaderText(invInpStr, printColor:=ConsoleColor.Red)

    End Sub

End Module
