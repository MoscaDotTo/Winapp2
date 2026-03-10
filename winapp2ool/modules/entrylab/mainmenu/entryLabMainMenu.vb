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
''' Displays the Entry Lab menu and routes to generative sub-modules.
''' Entry Lab is a navigation hub grouping all modules that <em>generate</em>
''' winapp2.ini entries from templates, keeping the main menu stable as
''' new generators are added.
''' </summary>
Module entryLabMainMenu

    ''' <summary>
    ''' Prints the Entry Lab menu to the user
    ''' </summary>
    Public Sub printEntryLabMenu()

        Dim menuDesc = {"Entry Lab is a collection of tools that generate winapp2.ini entries from templates",
                        "Select a tool below to get started"}

        Dim menu = MenuSection.CreateCompleteMenu("Entry Lab", menuDesc, ConsoleColor.Cyan).AddBlank

        menu.AddColoredOption(NameOf(BrowserBuilder), "Generate winapp2.ini entries for web browsers", ConsoleColor.DarkYellow) _
            .AddColoredOption(NameOf(UWPBuilder), "Generate winapp2.ini entries for Universal Windows Platform apps", ConsoleColor.Blue)

        menu.Print()

    End Sub

    ''' <summary>
    ''' Handles user input for the Entry Lab menu
    ''' </summary>
    '''
    ''' <param name="input">
    ''' The user's input
    ''' </param>
    Public Sub handleEntryLabInput(input As String)

        Dim modules = New Dictionary(Of String, KeyValuePair(Of Action, Action(Of String))) From {
            {NameOf(BrowserBuilder), New KeyValuePair(Of Action, Action(Of String))(AddressOf printBrowserBuilderMenu, AddressOf handleBrowserBuilderInput)},
            {NameOf(UWPBuilder), New KeyValuePair(Of Action, Action(Of String))(AddressOf printUWPBuilderMenu, AddressOf handleUWPBuilderInput)}
        }

        Dim moduleOpts As String() = {"1", "2"}

        Select Case True

            Case input = "0"

                exitModule()

            Case moduleOpts.Contains(input) OrElse input.Length = 0

                If input.Length = 0 Then input = "1"
                Dim i = CType(input, Integer) - 1
                initModule(modules.Keys(i), modules.Values(i).Key, modules.Values(i).Value)

            Case Else

                setNextMenuHeaderText(invInpStr, printColor:=ConsoleColor.Red)

        End Select

    End Sub

End Module
