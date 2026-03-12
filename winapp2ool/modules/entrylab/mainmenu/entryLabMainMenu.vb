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
    ''' Builds the Entry Lab menu with all options and their dispatch handlers registered inline.
    ''' Called by both <c> printEntryLabMenu </c> (to render) and <c> handleEntryLabInput </c>
    ''' (to dispatch), so the displayed option numbers and the dispatch table are always in sync.
    ''' </summary>
    Private Function buildEntryLabMenu() As MenuSection

        Dim menuDesc = {"Entry Lab is a collection of tools that generate winapp2.ini entries from templates",
                        "Select a tool below to get started"}

        Return MenuSection.CreateCompleteMenu("Entry Lab", menuDesc, ConsoleColor.Cyan).AddBlank() _
            .AddDispatchedColoredOption(NameOf(BrowserBuilder), "Generate winapp2.ini entries for web browsers", ConsoleColor.DarkYellow,
                Sub() initModule(NameOf(BrowserBuilder), AddressOf printBrowserBuilderMenu, AddressOf handleBrowserBuilderInput)) _
            .AddDispatchedColoredOption(NameOf(UWPBuilder), "Generate winapp2.ini entries for Universal Windows Platform apps", ConsoleColor.Blue,
                Sub() initModule(NameOf(UWPBuilder), AddressOf printUWPBuilderMenu, AddressOf handleUWPBuilderInput))

    End Function

    ''' <summary>
    ''' Prints the Entry Lab menu to the user
    ''' </summary>
    Public Sub printEntryLabMenu()

        buildEntryLabMenu().Print()

    End Sub

    ''' <summary>
    ''' Handles user input for the Entry Lab menu
    ''' </summary>
    '''
    ''' <param name="input">
    ''' The user's input
    ''' </param>
    Public Sub handleEntryLabInput(input As String)

        Dim intInput As Integer

        If Not Integer.TryParse(input, intInput) Then

            ' Allow an empty input to default to BrowserBuilder (option 1)
            If input.Length = 0 Then initModule(NameOf(BrowserBuilder), AddressOf printBrowserBuilderMenu, AddressOf handleBrowserBuilderInput) : Return

            setNextMenuHeaderText(invInpStr, printColor:=ConsoleColor.Red)
            Return

        End If

        If intInput = 0 Then exitModule() : Return

        If Not buildEntryLabMenu().Dispatch(intInput) Then setNextMenuHeaderText(invInpStr, printColor:=ConsoleColor.Red)

    End Sub

End Module
