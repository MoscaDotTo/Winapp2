'    Copyright (C) 2018-2020 Robbie Ward
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
''' <summary> Displays the global settings menu to the user and handles input from that menu  </summary>
Module globalsettingsmenu

    ''' <summary> Initalizes the default state of the winapp2ool module settings </summary>
    ''' Docs last updated: 2020-07-14 | Code last updated: 2020-07-14
    Private Sub initDefaultSettings()
        GlobalLogFile.resetParams()
        RemoteWinappIsNonCC = False
        saveSettingsToDisk = False
        readSettingsFromDisk = False
        toolSettingsHaveChanged = False
        restoreDefaultSettings(NameOf(Winapp2ool), AddressOf createToolSettingsSection)
    End Sub

    ''' <summary> Prints the winapp2ool settings menu to the user </summary>
    ''' Docs last updated: 2020-07-14 | Code last updated: 2020-07-14
    Public Sub printMainToolSettingsMenu()
        printMenuTop({"Change some high level settings, including saving & reading settings from disk"})
        print(5, "Toggle Saving Settings", $"saving a copy of winapp2ool's settings to the disk", enStrCond:=saveSettingsToDisk, leadingBlank:=True)
        print(5, "Toggle Reading Settings", $"overriding winapp2ool's default settings with those found in winapp2ool.ini", enStrCond:=readSettingsFromDisk, trailingBlank:=True)
        print(5, "Toggle Non-CCleaner Mode", $"using the Non-CCleaner version of winapp2.ini by default", enStrCond:=RemoteWinappIsNonCC, trailingBlank:=True)
        print(1, "View Log", "Print winapp2ool's internal log")
        print(1, "File Chooser (log)", "Change the filename or path to which the winapp2ool log should be saved")
        print(1, "Save Log", "Save winapp2ool's internal log to the disk")
        print(0, $"Current log file target: {replDir(GlobalLogFile.Path)}", leadingBlank:=True, trailingBlank:=True)
        print(1, "Visit GitHub", "Open the winapp2.ini/winapp2ool GitHub in your default web browser", trailingBlank:=True)
        print(5, "Toggle Beta Participation", $"participating in the 'beta' builds of winapp2ool (requires a restart)", enStrCond:=isBeta, closeMenu:=Not toolSettingsHaveChanged)
        print(2, NameOf(Winapp2ool), cond:=toolSettingsHaveChanged, closeMenu:=True)
    End Sub

    ''' <summary> Handles the user input for the winapp2ool settings menu </summary>
    ''' <param name="input"> The user's input </param>
    ''' Docs last updated: 2020-07-14 | Code last updated: 2020-07-14
    Public Sub handleMainToolSettingsInput(input As String)
        Select Case True
            Case input = "0"
                exitModule()
            Case input = "1"
                toggleSettingParam(saveSettingsToDisk, "Saving settings to disk", toolSettingsHaveChanged, NameOf(Winapp2ool), NameOf(saveSettingsToDisk), NameOf(toolSettingsHaveChanged))
            Case input = "2"
                toggleSettingParam(readSettingsFromDisk, "Reading settings from disk", toolSettingsHaveChanged, NameOf(Winapp2ool), NameOf(readSettingsFromDisk), NameOf(toolSettingsHaveChanged))
            Case input = "3"
                toggleSettingParam(RemoteWinappIsNonCC, "Non-CCleaner mode", toolSettingsHaveChanged, NameOf(Winapp2ool), NameOf(RemoteWinappIsNonCC), NameOf(toolSettingsHaveChanged))
            Case input = "4"
                printLog()
            Case input = "5"
                changeFileParams(GlobalLogFile, toolSettingsHaveChanged, NameOf(Winapp2ool), NameOf(GlobalLogFile), NameOf(toolSettingsHaveChanged))
            Case input = "6"
                GlobalLogFile.overwriteToFile(logger.toString)
            Case input = "7"
                Process.Start(gitLink)
            Case input = "8"
                If Not denyActionWithHeader(DotNetFrameworkOutOfDate, "Winapp2ool beta requires .NET 4.6 or higher") Then
                    toggleSettingParam(isBeta, "Beta Participation", toolSettingsHaveChanged, NameOf(Winapp2ool), NameOf(isBeta), NameOf(toolSettingsHaveChanged))
                    autoUpdate()
                End If
            Case input = "9" And toolSettingsHaveChanged
                initDefaultSettings()
            Case Else
                setHeaderText(invInpStr, True)
        End Select
    End Sub

End Module
