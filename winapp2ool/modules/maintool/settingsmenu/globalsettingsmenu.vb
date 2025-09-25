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
''' </summary
Module globalsettingsmenu

    ''' <summary> 
    ''' Initalizes the default state of the winapp2ool module settings 
    ''' </summary>
    Private Sub initDefaultSettings()

        GlobalLogFile.resetParams()
        RemoteWinappIsNonCC = False
        saveSettingsToDisk = False
        readSettingsFromDisk = False
        toolSettingsHaveChanged = False
        CurrentWinappFlavor = WinappFlavor.CCleaner

        restoreDefaultSettings(NameOf(Winapp2ool), AddressOf createToolSettingsSection)

    End Sub

    ''' <summary> 
    ''' Prints the winapp2ool settings menu to the user 
    ''' </summary>
    Public Sub printMainToolSettingsMenu()

        Dim currentLogTarget = $"Current log file target: {replDir(GlobalLogFile.Path)}"
        Dim VisitGitHubName = "Visit GitHub"
        Dim VisitGitHubDesc = "Open the Winapp2 GitHub page in your default web browser"
        Dim ToggleBetaName = "Toggle Beta Participation"
        Dim ToggleBetaDesc = "participating in the 'beta' builds of winapp2ool (requires a restart)"
        Dim ToggleOfflineName = "Toggle Offline Mode"
        Dim ToggleOfflineDesc = "Force winapp2ool into offline mode"

        Dim menuDesc = {"Change high level settings for winapp2ool",
                       "Enable reading and writing settings from disk to persist any changes made here"}

        Dim menu = MenuSection.CreateCompleteMenu("Winapp2ool Global Settings", menuDesc, ConsoleColor.DarkGreen)
        menu.AddBlank() _
            .AddToggle("Saving Settings", "saving a copy of winapp2ool's settings to the disk", saveSettingsToDisk) _
            .AddToggle("Reading settings", "overriding winapp2ool's default settings at launch using winapp2ool.ini", readSettingsFromDisk) _
            .AddToggle("Beta Participation", "participating in the 'beta' builds of winapp2ool (requires a restart)", isBeta) _
            .AddToggle("Offline Mode", "forcing winapp2ool into ofline mode", isOffline).AddBlank() _
            .AddColoredOption("Change Flavor", "Cycle the current flavor of winapp2.ini to the next", ConsoleColor.DarkMagenta) _
            .AddColoredLine($"Current Flavor: {CurrentWinappFlavor.ToString}", ConsoleColor.Magenta).AddBlank _
            .AddColoredOption("View Log", "Print winapp2ool's current internal log", ConsoleColor.DarkYellow) _
            .AddColoredOption("Save log", "Save winapp2ool's current internal log to disk", ConsoleColor.DarkYellow) _
            .AddColoredOption("Change Save Target", "Select a new filename of path to which the winapp2ool log should be saved", ConsoleColor.DarkYellow).AddBlank _
            .AddColoredFileInfo("Current save target: ", GlobalLogFile.Path, ConsoleColor.DarkYellow).AddBlank _
            .AddOption("Visit GitHub", "Open the Winapp2 GitHub page in your default web browser").AddBlank() _
            .AddResetOpt(NameOf(Winapp2ool), toolSettingsHaveChanged)

        menu.Print()

    End Sub

    ''' <summary>
    ''' Returns the set of toggles in the winapp2ool settings menu
    ''' </summary>
    ''' <returns></returns>
    Private Function getToggleOpts() As Dictionary(Of String, String)

        Dim toggles As New Dictionary(Of String, String)

        toggles.Add("Saving Settings", NameOf(saveSettingsToDisk))
        toggles.Add("Reading settings", NameOf(readSettingsFromDisk))
        toggles.Add("Beta Mode", NameOf(isBeta))
        toggles.Add("Offline Mode", NameOf(isOffline))

        Return toggles

    End Function

    ''' <summary>
    ''' Returns the set of enums in the winapp2ool settings menu
    ''' </summary>
    ''' <returns></returns>
    Private Function getEnumOpts() As Dictionary(Of String, String)

        Dim enums As New Dictionary(Of String, String)

        enums.Add("Flavor", NameOf(CurrentWinappFlavor))

        Return enums

    End Function
    ''' <summary> 
    ''' Handles the user input for the winapp2ool settings menu 
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

        Dim toggles = getToggleOpts()
        Dim enums = getEnumOpts()

        Dim staticOpts = 1
        Dim minToggleNum = staticOpts
        Dim maxToggleNum = minToggleNum + toggles.Count - 1
        Dim minCycleNum = maxToggleNum + 1
        Dim maxCycleNum = minCycleNum + enums.Count - 1
        Dim minLogNum = maxCycleNum + 1
        Dim maxLogNum = minLogNum + 2
        Dim minGHNum = maxLogNum + 1
        Dim maxGHNum = minGHNum
        Dim resetNum = maxGHNum + 1

        Select Case True

            ' Exit
            Case intInput = 0

                exitModule()

            ' Toggles
            Case intInput >= minToggleNum AndAlso intInput <= maxToggleNum

                Dim i = intInput - staticOpts

                If toggles.Keys(i) = "Beta Mode" Then

                    If denyActionWithHeader(DotNetFrameworkOutOfDate, "Winapp2ool beta requires .NET 4.6 or higher") Then Return

                    toggleModuleSetting(toggles.Keys(i), NameOf(Winapp2ool), GetType(maintoolsettings), toggles.Values(i), NameOf(toolSettingsHaveChanged))
                    autoUpdate()

                End If

                toggleModuleSetting(toggles.Keys(i), NameOf(Winapp2ool), GetType(maintoolsettings), toggles.Values(i), NameOf(toolSettingsHaveChanged))

            ' Enums 
            Case intInput >= minCycleNum AndAlso intInput <= maxCycleNum

                Dim i = intInput - maxToggleNum - 1
                CycleEnumProperty(enums.Values(i), "Flavor", GetType(maintoolsettings), NameOf(Winapp2ool), toolSettingsHaveChanged, NameOf(toolSettingsHaveChanged), ConsoleColor.DarkMagenta)


            ' View Log
            Case intInput = minLogNum

                printLog()

            ' Save Log
            Case intInput = minLogNum + 1

                GlobalLogFile.overwriteToFile(logger.toString)

            ' File Selector
            Case intInput = maxLogNum

                changeFileParams(GlobalLogFile, toolSettingsHaveChanged, NameOf(Winapp2ool), NameOf(GlobalLogFile), NameOf(toolSettingsHaveChanged))


            ' Visit GitHub
            Case intInput = minGHNum

                Process.Start(gitLink)

            ' Reset Settings
            Case intInput = resetNum AndAlso toolSettingsHaveChanged

                initDefaultSettings()

            Case Else

                setNextMenuHeaderText(invInpStr, printColor:=ConsoleColor.Red)

        End Select

    End Sub

End Module
