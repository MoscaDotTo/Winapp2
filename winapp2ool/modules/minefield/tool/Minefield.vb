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
''' Minefield is a winapp2ool module which serves as a testing ground for new ideas and features
''' that don't fit into an existing module. It is not intended for general use, and the features 
''' here may be incomplete, experimental, or broken.
''' </summary>
Module Minefield

    ''' <summary>
    ''' Prints the main Minefield menu
    ''' </summary>
    Public Sub printMenu()

        Dim menuDescriptionLines = {"A testing ground for new ideas/features, watch your step!",
                                    "Items in this menu may be incomplete, experimental, or broken entirely",
                                    "Watch your step!"}

        Dim menu = MenuSection.CreateCompleteMenu("Minefield", menuDescriptionLines, ConsoleColor.DarkBlue)

        menu.Print()

    End Sub

    ''' <summary>
    ''' Handles user input for the main Minefield menu
    ''' </summary>
    ''' 
    ''' <param name="input">
    ''' The user's input
    ''' </param>
    Public Sub handleUserInput(input As String)

        Select Case True

            Case input = "0"

                exitModule()

            Case Else

                setNextMenuHeaderText(invInpStr, printColor:=ConsoleColor.DarkRed)

        End Select

    End Sub

End Module