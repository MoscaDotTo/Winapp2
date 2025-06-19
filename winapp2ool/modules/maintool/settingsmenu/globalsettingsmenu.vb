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
''' 
''' Docs last updated: 2023-07-19 | Code last updated: 2023-07-19
Module globalsettingsmenu

    ''' <summary> 
    ''' Initalizes the default state of the winapp2ool module settings 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2020-07-14 | Code last updated: 2020-07-14
    Private Sub initDefaultSettings()

        GlobalLogFile.resetParams()
        RemoteWinappIsNonCC = False
        saveSettingsToDisk = False
        readSettingsFromDisk = False
        toolSettingsHaveChanged = False

        restoreDefaultSettings(NameOf(Winapp2ool), AddressOf createToolSettingsSection)

    End Sub

    ''' <summary> 
    ''' Prints the winapp2ool settings menu to the user 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2023-07-19 | Code last updated: 2023-07-19
    Public Sub printMainToolSettingsMenu()

        Dim toolDesc = "Change some high level settings, including saving & reading settings from disk"
        Dim toggleSavingSettingsName = "Toggle Saving Settings"
        Dim toggleSavingSettingsDesc = "saving a copy of winapp2ool's settings to the disk"
        Dim toggleReadingSettingsName = "Toggle Reading Settings"
        Dim toggleReadingSettingsDesc = "overriding winapp2ool's default settings with those found in winapp2ool.ini"
        Dim toggleNonCCleanerName = "Toggle Non-CCleaner Mode"
        Dim toggleNonCCleanerDesc = "using the Non-CCleaner version of winapp2.ini by default"
        Dim viewLogName = "View Log"
        Dim viewLogDesc = "Print winapp2ool's internal log"
        Dim fileChooserName = "File Chooser (log)"
        Dim fileChooserDesc = "Change the filename or path to which the winapp2ool log should be saved"
        Dim saveLogName = "Save Log"
        Dim saveLogDesc = "Save winapp2ool's internal log to the disk"
        Dim currentLogTarget = $"Current log file target: {replDir(GlobalLogFile.Path)}"
        Dim VisitGitHubName = "Visit GitHub"
        Dim VisitGitHubDesc = "Open the Winapp2 GitHub page in your default web browser"
        Dim ToggleBetaName = "Toggle Beta Participation"
        Dim ToggleBetaDesc = "participating in the 'beta' builds of winapp2ool (requires a restart)"
        Dim ToggleOfflineName = "Toggle Offline Mode"
        Dim ToggleOfflineDesc = "Force winapp2ool into offline mode"

        printMenuTop({toolDesc})
        print(5, toggleSavingSettingsName, toggleSavingSettingsDesc, enStrCond:=saveSettingsToDisk, leadingBlank:=True)
        print(5, toggleReadingSettingsName, toggleReadingSettingsDesc, enStrCond:=readSettingsFromDisk, trailingBlank:=True)
        print(5, toggleNonCCleanerName, toggleNonCCleanerDesc, enStrCond:=RemoteWinappIsNonCC, trailingBlank:=True)
        print(1, viewLogName, viewLogDesc)
        print(1, fileChooserName, fileChooserDesc)
        print(1, saveLogName, saveLogDesc)
        print(0, currentLogTarget, leadingBlank:=True, trailingBlank:=True)
        print(1, VisitGitHubName, VisitGitHubDesc, trailingBlank:=True)
        print(5, ToggleBetaName, ToggleBetaDesc, enStrCond:=isBeta)
        print(5, ToggleOfflineName, ToggleOfflineDesc, enStrCond:=isOffline, closeMenu:=Not toolSettingsHaveChanged)
        print(2, NameOf(Winapp2ool), cond:=toolSettingsHaveChanged, closeMenu:=True)

    End Sub

    ''' <summary> 
    ''' Handles the user input for the winapp2ool settings menu 
    ''' </summary>
    ''' 
    ''' <param name="input">
    ''' The user's input
    ''' </param>
    ''' 
    ''' Docs last updated: 2020-07-14 | Code last updated: 2020-07-14
    Public Sub handleMainToolSettingsInput(input As String)

        Select Case True

            ' Option Name:                                 Exit
            ' Option States:
            ' Default                                      -> 0 (default)
            Case input = "0"

                exitModule()

            ' Option Name:                                 Toggle Saving Settings
            ' Option States:
            ' Default                                      -> 1 (default)
            Case input = "1"

                toggleSettingParam(saveSettingsToDisk, "Saving settings to disk", toolSettingsHaveChanged, NameOf(Winapp2ool), NameOf(saveSettingsToDisk), NameOf(toolSettingsHaveChanged))

            ' Option Name:                                 Toggle Reading Settings
            ' Option States:
            ' Default                                      -> 2 (default)
            Case input = "2"

                toggleSettingParam(readSettingsFromDisk, "Reading settings from disk", toolSettingsHaveChanged, NameOf(Winapp2ool), NameOf(readSettingsFromDisk), NameOf(toolSettingsHaveChanged))

            ' Option Name:                                 Toggle Non-CCleaner Mode
            ' Option States:
            ' Default                                      -> 3 (default)
            Case input = "3"

                toggleSettingParam(RemoteWinappIsNonCC, "Non-CCleaner mode", toolSettingsHaveChanged, NameOf(Winapp2ool), NameOf(RemoteWinappIsNonCC), NameOf(toolSettingsHaveChanged))

            ' Option Name:                                 View Log
            ' Option States:
            ' Default                                      -> 4 (default)
            Case input = "4"

                printLog()

            ' Option Name:                                 File Chooser (log)
            ' Option States:
            ' Default                                      -> 5 (default)
            Case input = "5"

                changeFileParams(GlobalLogFile, toolSettingsHaveChanged, NameOf(Winapp2ool), NameOf(GlobalLogFile), NameOf(toolSettingsHaveChanged))

            ' Option Name:                                 Save Log
            ' Option States:
            ' Default                                      -> 6 (default)
            Case input = "6"

                GlobalLogFile.overwriteToFile(logger.toString)

            ' Option Name:                                 Visit GitHub
            ' Option States:
            ' Default                                      -> 7 (default)
            Case input = "7"

                Process.Start(gitLink)

            ' Option Name:                                 Toggle Beta Participation
            ' Option States:
            ' Default                                      -> 8 (default)
            Case input = "8"

                If Not denyActionWithHeader(DotNetFrameworkOutOfDate, "Winapp2ool beta requires .NET 4.6 or higher") Then

                    toggleSettingParam(isBeta, "Beta Participation", toolSettingsHaveChanged, NameOf(Winapp2ool), NameOf(isBeta), NameOf(toolSettingsHaveChanged))
                    autoUpdate()

                End If

            ' Option Name:                                 Toggle Offline Mode
            ' Option States:
            ' Default                                      -> 9 (default)
            Case input = "9"

                isOffline = True

            ' Option Name:                                 Reset Settings
            ' Option States:
            ' Default                                      -> 10 (default)
            Case input = "10" And toolSettingsHaveChanged

                initDefaultSettings()

            Case Else

                setHeaderText(invInpStr, True)

        End Select

    End Sub

End Module