'    Copyright (C) 2018-2021 Hazel Ward
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
''' This module serves as a simple interface for testing new ideas that don't fit neatly into an existing module
''' The ideas here should be considered alpha and the code here should be considered spaghetti 
''' </summary>
Module Minefield
    ''' <summary>Prints the main Minefield menu</summary>
    Public Sub printMenu()
        printMenuTop({"A testing ground for new ideas/features, watch your step!", "Some options may not do anything."})
        print(1, "Java Entry Maker", "Clean up after the messy JRE installer")
        print(1, "Babel", "Generate winapp2.ini entries for lang files")
        print(1, "Outlook", "Generate winapp2.ini entries with custom outlook profile support")
        print(1, "Search", "Generate lists of sections", closeMenu:=True)
    End Sub

    '''<summary>Prints the main Babel menu</summary>
    Public Sub printBabelMenu()
        printMenuTop({"This tool will attempt to generate winapp2.ini entries to remove language files from your system. Use EXTREME caution!"})
        print(1, "Run (Default)", "Run the tool")
    End Sub

    '''<summary>Handles the user input for Babel</summary>
    '''<param name="input">The String containing the user's input</param>
    Public Sub handleBabelInput(input As String)
        Select Case input
            Case "0"
                exitModule()
            Case "1", ""

        End Select
    End Sub

    ''' <summary>Handles user input for the main Minefield menu</summary>
    ''' <param name="input">The String containing the user's input</param>
    Public Sub handleUserInput(input As String)
        Select Case input
            Case "0"
                exitModule()
            Case "1"
                initModule("Java Entry Maker", AddressOf printJMMenu, AddressOf handleJMInput)
            Case "2"
                initModule("Babel", AddressOf printBabelMenu, AddressOf handleBabelInput)
            Case "3"
                initModule("Outooker", AddressOf printOutlookMenu, AddressOf handleOutlookInput)
            Case "4"
                initModule("Search", AddressOf printSearchMenu, AddressOf handleSearchInput)
        End Select
    End Sub

    Public Sub printSearchMenu()
        printMenuTop({"Outputs all sections that fit the given pattern in a file"})
        print(1, "winapp2.ini", "Search winapp2.ini by section criteria", closeMenu:=True)
    End Sub

    Public Sub handleSearchInput(input As String)
        Select Case True
            Case input = "0"
                exitModule()
            Case input = "1" Or input.Length = 0
                clrConsole()
                Console.Write("Enter the Section or LangSecRef you'd like from the file:")
                Dim sect = Console.ReadLine
                Dim iwinapp2 = New iniFile(Environment.CurrentDirectory, "winapp2.ini")
                iwinapp2.validate()
                Dim winapp2 = New winapp2file(iwinapp2)
                Dim out = ""
                For Each lst In winapp2.Winapp2entries
                    For Each entry In lst
                        checkSection(entry, entry.SectionKey, sect, out)
                        checkSection(entry, entry.LangSecRef, sect, out)
                    Next
                Next
                cwl(out)
                crk()
        End Select
    End Sub

    Private Sub checkSection(entry As winapp2entry, lst As keyList, sect As String, ByRef out As String)
        If lst.KeyCount = 0 Then Exit Sub

        If lst.Keys(0).Value = sect Then
            Dim tmp = entry.dumpToListOfStrings
            For Each line In tmp
                out += line & Environment.NewLine
            Next
            out += Environment.NewLine
        End If


    End Sub

    ''' <summary>Prints the GameMaker menu</summary>
    Private Sub printGMMenu()
        printMenuTop({"This tool will allow for a more meta approach to creating entries for games, particularly steam."})
        print(1, "Run (Disabled)", "Attempt to generate entries", closeMenu:=True)
    End Sub

    ''' <summary>Handles input for GameMaker</summary>
    ''' <param name="input">The String containing the user's input</param>
    Private Sub handleGMInput(input As String)
        Select Case input
            Case "0"
                exitModule()
        End Select
    End Sub
End Module