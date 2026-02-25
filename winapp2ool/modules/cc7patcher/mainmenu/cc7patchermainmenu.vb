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
''' Displays the CC7Patcher module main menu and handles user input
''' </summary>
Public Module cc7patchermainmenu

    ''' <summary>
    ''' Prints the CC7Patcher menu to the user
    ''' </summary>
    Public Sub printCC7PatcherMenu()

        Dim menuDescLines = {"Patch ccleaner.ini with winapp2.ini entries compatible with CCleaner 7"}

        Dim menu = MenuSection.CreateCompleteMenu(NameOf(CC7Patcher), menuDescLines, ConsoleColor.Yellow)

        menu.AddBlank() _
            .AddColoredOption("Run (default)", "Install winapp2.ini for CCleaner 7", ConsoleColor.Yellow).AddBlank _
            .AddToggle("Trim", "trimming winapp2.ini before installation", TrimBeforePatching) _
            .AddToggle("Download", "downloading the latest winapp2.ini from GitHub", DownloadWinapp2, Not isOffline).AddBlank() _
            .AddColoredOption("Change winapp2.ini", "Select the winapp2.ini file to install", ConsoleColor.Green, Not DownloadWinapp2) _
            .AddColoredOption("Change ccleaner.ini", "Select the ccleaner.ini file to be patched", ConsoleColor.Red) _
            .AddColoredOption("Change output file", "Select where to save the patched ccleaner.ini", ConsoleColor.Cyan).AddBlank() _
            .AddColoredFileInfo("Current winapp2.ini:  ", If(DownloadWinapp2, "Online", CC7PatcherFile1.Path), ConsoleColor.Green) _
            .AddColoredFileInfo("Current ccleaner.ini: ", CC7PatcherFile2.Path, ConsoleColor.Red) _
            .AddColoredFileInfo("Output file:          ", CC7PatcherFile3.Path, ConsoleColor.Cyan).AddBlank() _
            .AddResetOpt(NameOf(CC7Patcher), CC7PatcherModuleSettingsChanged)

        menu.Print()

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

            setNextMenuHeaderText(invInpStr, printColor:=ConsoleColor.Red)
            Return

        End If

        Dim fileOpts = getFileOpts()
        Dim toggleOpts = getToggleOpts()

        Dim exitNum = 0
        Dim runNum = exitNum + 1
        Dim toggleStartNum = runNum + 1
        Dim toggleEndNum = toggleStartNum + toggleOpts.Count - 1
        Dim fileStartNum = toggleEndNum + 1
        Dim fileEndNum = fileStartNum + fileOpts.Count - 1
        Dim resetNum = fileEndNum + 1

        Select Case True

            Case intInput = exitNum

                exitModule()

            Case intInput = runNum

                ' Run 

            Case intInput >= toggleStartNum AndAlso intInput <= toggleEndNum

                Dim i = intInput - toggleStartNum

                Dim toggleMenuText = toggleOpts.Keys(i)
                Dim toggleName = toggleOpts(toggleMenuText)

                toggleModuleSetting(toggleMenuText, NameOf(CC7Patcher), GetType(cc7patchersettings),
                                    toggleName, NameOf(CC7PatcherModuleSettingsChanged))

            Case intInput >= fileStartNum AndAlso intInput <= fileEndNum

                Dim i = intInput - fileStartNum

                Dim fileName = fileOpts.Keys(i)
                Dim fileObj = fileOpts(fileName)

                changeFileParams(fileObj, CC7PatcherModuleSettingsChanged, NameOf(CC7Patcher),
                                 fileName, NameOf(CC7PatcherModuleSettingsChanged))

            Case CC7PatcherModuleSettingsChanged AndAlso intInput = resetNum

                resetModuleSettings(NameOf(CC7Patcher), AddressOf initDefaultCC7PatcherSettings)

            Case Else

                setNextMenuHeaderText(invInpStr, printColor:=ConsoleColor.Red)

        End Select

    End Sub

    Private Function getToggleOpts() As Dictionary(Of String, String)

        Dim toggles As New Dictionary(Of String, String)

        toggles.Add("Trimming", NameOf(TrimBeforePatching))
        If Not isOffline Then toggles.Add("Downloading", NameOf(DownloadWinapp2))

        Return toggles

    End Function

    Private Function getFileOpts() As Dictionary(Of String, iniFile)

        Dim selectors As New Dictionary(Of String, iniFile)

        If Not DownloadWinapp2 Then selectors.Add(NameOf(CC7PatcherFile1), CC7PatcherFile1)
        selectors.Add(NameOf(CC7PatcherFile2), CC7PatcherFile2)
        selectors.Add(NameOf(CC7PatcherFile3), CC7PatcherFile3)

        Return selectors

    End Function

End Module