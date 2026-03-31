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
''' Displays the global settings menu to the user and handles input from that menu
''' </summary>
Module globalsettingsmenu

    ''' <summary>
    ''' Initializes the default state of the Winapp2ool module settings
    ''' </summary>
    Private Sub initDefaultSettings()

        InitDefaultToolSettings()

    End Sub

    ''' <summary>
    ''' Builds the Winapp2ool global settings menu
    ''' </summary>
    Private Function buildMainToolSettingsMenu() As MenuSection

        Dim menuDesc = {"Change high level settings for winapp2ool",
                       "Enable reading and writing settings from disk to persist any changes made here"}

        Dim menu = MenuSection.CreateCompleteMenu("Winapp2ool Global Settings", menuDesc, ConsoleColor.DarkGreen)

        menu.AddBlank() _
            .AddDispatchedToggle("Saving Settings", "saving a copy of winapp2ool's settings to the disk",
                saveSettingsToDisk,
                Sub() toggleModuleSetting(NameOf(saveSettingsToDisk), NameOf(Winapp2ool), GetType(maintoolsettings), NameOf(saveSettingsToDisk), NameOf(toolSettingsHaveChanged))) _
            .AddDispatchedToggle("Reading settings", "overriding winapp2ool's default settings at launch using winapp2ool.ini", readSettingsFromDisk,
                Sub() toggleModuleSetting(NameOf(readSettingsFromDisk), NameOf(Winapp2ool), GetType(maintoolsettings), NameOf(readSettingsFromDisk), NameOf(toolSettingsHaveChanged))) _
            .AddDispatchedToggle("Beta Participation", "participating in the 'beta' builds of winapp2ool (requires a restart)", isBeta,
                Sub()
                    If denyActionWithHeader(DotNetFrameworkOutOfDate, "Winapp2ool beta requires .NET 4.6 or higher") Then Return
                    toggleModuleSetting(NameOf(isBeta), NameOf(Winapp2ool), GetType(maintoolsettings), NameOf(isBeta), NameOf(toolSettingsHaveChanged))
                    autoUpdate()
                End Sub) _
            .AddDispatchedToggle("Offline Mode", "forcing winapp2ool into offline mode", isOffline,
                Sub() toggleModuleSetting(NameOf(isOffline), NameOf(Winapp2ool), GetType(maintoolsettings), NameOf(isOffline), NameOf(toolSettingsHaveChanged))).AddBlank() _
            .AddDispatchedColoredOption("Change Flavor", "Cycle the current flavor of winapp2.ini to the next", ConsoleColor.DarkMagenta,
                Sub() CycleEnumProperty(NameOf(CurrentWinappFlavor), "Flavor", GetType(maintoolsettings), NameOf(Winapp2ool), toolSettingsHaveChanged, NameOf(toolSettingsHaveChanged), ConsoleColor.DarkMagenta)) _
            .AddColoredLine($"Current Flavor: {CurrentWinappFlavor.ToString}", ConsoleColor.Magenta).AddBlank() _
            .AddDispatchedColoredOption("View Log", "Print winapp2ool's current internal log", ConsoleColor.DarkYellow,
                Sub() printLog()) _
            .AddDispatchedColoredOption("Save log", "Save winapp2ool's current internal log to disk", ConsoleColor.DarkYellow,
                Sub() saveGlobalLog()) _
            .AddDispatchedColoredOption("Change Save Target", "Select a new filename or path to which the winapp2ool log should be saved", ConsoleColor.DarkYellow,
                Sub() changeFile2Params(GlobalLogFile, toolSettingsHaveChanged, NameOf(Winapp2ool), NameOf(GlobalLogFile), NameOf(toolSettingsHaveChanged))).AddBlank() _
            .AddColoredFileInfo("Current save target: ", GlobalLogFile.Path(), ConsoleColor.DarkYellow).AddBlank() _
            .AddDispatchedOption("Visit GitHub", "Open the Winapp2 GitHub page in your default web browser",
                Sub() Process.Start(gitLink)) _
            .AddBlank() _
            .AddDispatchedResetOpt(NameOf(Winapp2ool), toolSettingsHaveChanged, Sub() initDefaultSettings())

        Return menu

    End Function

    ''' <summary>
    ''' Prints the Winapp2ool global settings menu to the user
    ''' </summary>
    Public Sub printMainToolSettingsMenu()

        buildMainToolSettingsMenu().Print()

    End Sub

    ''' <summary>
    ''' Handles user input for the Winapp2ool global settings menu
    ''' </summary>
    '''
    ''' <param name="input">
    ''' The user's input
    ''' </param>
    Public Sub handleMainToolSettingsInput(input As String)

        Dim intInput As Integer
        If Not Integer.TryParse(input, intInput) Then
            setNextMenuHeaderText(invInpStr, printColor:=ConsoleColor.Red)
            Return
        End If

        If intInput = 0 Then
            exitModule()
            Return
        End If

        If Not buildMainToolSettingsMenu().Dispatch(intInput) Then
            setNextMenuHeaderText(invInpStr, printColor:=ConsoleColor.Red)
        End If

    End Sub

End Module
