Option Strict On
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

Imports System.ComponentModel.Design
Imports Microsoft.VisualBasic.Logging

''' <summary>
''' Provides the user interface for the Combine module
''' </summary>
''' 
''' Docs last updated: 2025-08-20 | Code last updated: 2025-08-20
Module combinemainmenu

    ''' <summary>
    ''' Prints the Combine module menu to the user
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-08-20 | Code last updated: 2025-08-20
    Public Sub printCombineMainMenu()

        Dim menuDesc = {"Combine all ini files from a target directory and its children into a single file"}
        Dim noLog = MostRecentCombineLog.Length = 0

        Dim menu = MenuSection.CreateCompleteMenu(NameOf(Combine), menuDesc, ConsoleColor.DarkCyan)
        menu.AddOption("Run (default)", "Combine ini files").AddBlank() _
            .AddOption("Change the target directory", "File Chooser (target)") _
            .AddOption("Change the save file location", "File Chooser (save)").AddBlank() _
            .AddColoredOption("Log Viewer", "View the detailed Combine log", ConsoleColor.Yellow, Not noLog) _
            .AddBlank(Not noLog) _
            .AddLine($"Current target directory: {replDir(CombineFile1.Dir)}") _
            .AddLine($"Current save location:    {replDir(CombineFile3.Path)}", CombineModuleSettingsChanged) _
            .AddResetOpt(NameOf(Combine), CombineModuleSettingsChanged)

        menu.Print()

    End Sub

    ''' <summary>
    ''' Handles the user input from the Combine menu
    ''' </summary>
    ''' 
    ''' <param name="input">
    ''' The user's input string
    ''' </param>
    ''' 
    ''' Docs last updated: 2025-08-20 | Code last updated: 2025-08-20
    Public Sub handleCombineUserInput(input As String)

        Dim noLog = MostRecentCombineLog.Length = 0

        Dim fileOpts = getFileOpts()
        Dim fileNums = getMenuNumbering(fileOpts, 2)

        Dim logViewerNum = CType(fileOpts.Count + 2, String)
        Dim resetNum = CType(fileOpts.Count + 2 + If(noLog, 0, 1), String)

        Select Case True

            ' Exit 
            ' Notes: Always "0"
            Case input = "0"

                exitModule()

            ' Run (default)
            ' Notes: Always "1", also triggered by no input if run conditions are otherwise satisfied 
            Case (input = "1" Or input = "")

                initCombine(CombineFile1.Dir, CombineFile3)

            ' File Selectors
            ' Target dir (dir only)
            Case fileNums.Contains(input)

                Dim i = CType(input, Integer) - 2

                Dim fileName = fileOpts.Keys(i)
                Dim fileObj = fileOpts(fileName)

                changeFileParams(fileObj, CombineModuleSettingsChanged, NameOf(Combine),
                                 fileName, NameOf(CombineModuleSettingsChanged))

                ' This file is used only to hold the Target directory and never holds a name 
                CombineFile1.Name = ""
                updateSettings(NameOf(Combine), $"{NameOf(CombineFile1)}_Name", "")

            ' Log Viewer
            ' Notes: Only available after Combine has been run at least once during the current session
            Case Not noLog AndAlso input = logViewerNum

                printSlice(MostRecentCombineLog)

            ' Reset settings
            ' Notes: Only available after a setting has been changed, always comes last in the option list
            Case CombineModuleSettingsChanged And input = resetNum

                resetModuleSettings(NameOf(Combine), AddressOf initDefaultCombineSettings)

            Case Else

                setHeaderText(invInpStr, True)

        End Select

    End Sub

    ''' <summary>
    ''' Determines the current set of file selectors displayed on the menu and returns a Dictionary 
    ''' of those options and their respective files <br />
    ''' <br />
    ''' The set of possible files includes:
    ''' <list type="bullet">
    '''     
    '''     <item>
    '''     Source files target directory 
    '''     </item>
    '''     
    '''     <item>
    '''     Save target
    '''     </item>
    '''     
    ''' </list>
    ''' 
    ''' </summary>
    ''' 
    ''' <returns> 
    ''' The set of <c> iniFile </c> properties currently displayed on the menu
    ''' </returns>
    ''' 
    ''' Docs last updated: 2025-08-23 | Code last updated: 2025-08-23
    Private Function getFileOpts() As Dictionary(Of String, iniFile)

        Dim selectors As New Dictionary(Of String, iniFile)

        selectors.Add(NameOf(CombineFile1), CombineFile1)
        selectors.Add(NameOf(CombineFile3), CombineFile3)

        Return selectors

    End Function

End Module
