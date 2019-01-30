'    Copyright (C) 2018-2019 Robbie Ward
' 
'    This file is a part of Winapp2ool
' 
'    Winapp2ool is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    Winap2ool is distributed in the hope that it will be useful,
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

    ''' <summary>
    ''' Prints the main Minefield menu
    ''' </summary>
    Public Sub printMenu()
        printMenuTop({"A testing ground for new ideas/features, watch your step!"})
        print(1, "Java Entry Maker", "Clean up after the messy JRE installer")
        print(1, "Babel", "Generate winapp2.ini entries for lang files", closeMenu:=True)
    End Sub

    Public Sub printBabelMenu()
        printMenuTop({"This tool will attempt to generate winapp2.ini entries to remove language files from your system. Use EXTREME caution!"})
        print(1, "Run (Default)", "Run the tool")
    End Sub

    Public Sub handleBabelInput(input As String)
        Select Case input
            Case "0"
                exitCode = True
            Case "1", ""

        End Select
    End Sub

    ''' <summary>
    ''' Handles user input for the main Minefield menu
    ''' </summary>
    ''' <param name="input">The String containing the user's input</param>
    Public Sub handleUserInput(input As String)
        Select Case input
            Case "0"
                exitCode = True
            Case "1"
                initModule("Java Entry Maker", AddressOf printJMMenu, AddressOf handleJMInput)
            Case "2"
                initModule("Babel", AddressOf printBabelMenu, AddressOf handleBabelInput)
        End Select
    End Sub

    ''' <summary>
    ''' Prints the GameMaker menu
    ''' </summary>
    Private Sub printGMMenu()
        printMenuTop({"This tool will allow for a more meta approach to creating entries for games, particularly steam."})
        print(1, "Run (Disabled)", "Attempt to generate entries", closeMenu:=True)
    End Sub

    ''' <summary>
    ''' Handles input for GameMaker
    ''' </summary>
    ''' <param name="input"></param>
    Private Sub handleGMInput(input As String)
        Select Case input
            Case "0"
                exitCode = True
        End Select
    End Sub
End Module