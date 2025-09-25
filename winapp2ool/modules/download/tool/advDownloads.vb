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
''' A sub menu of the Downloader module for advanced users who want more than just the CCleaner flavor of winapp2.ini from the repo 
''' </summary>
Public Module advDownloads

    ''' <summary> 
    ''' Prints the advanced downloads menu 
    ''' </summary>
    Public Sub printAdvMenu()

        Dim menuDesc = {"Warning!", "Files in this menu are not recommended for use by beginners!"}
        Dim menu = MenuSection.CreateCompleteMenu("Advanced Downloads", menuDesc, ConsoleColor.DarkCyan)
        menu.AddBlank _
            .AddOption("Winapp3.ini", "Extended and/or potentially unsafe entries") _
            .AddOption("Archived entries.ini", "Entries for old or discontinued software")

        menu.Print()

    End Sub

    ''' <summary>
    ''' Handles the user input for the advanced download menu 
    ''' </summary>
    ''' 
    ''' <param name="input">
    ''' The user's input 
    ''' </param>
    Public Sub handleAdvInput(input As String)

        Select Case input

            Case "0"

                exitModule()

            Case "1"

                downloadFile.Name = "winapp3.ini"
                download(downloadFile, wa3link)

            Case "2"

                downloadFile.Name = "Archived entries.ini"
                download(downloadFile, archivedLink)

            Case Else

                setNextMenuHeaderText(invInpStr, printColor:=ConsoleColor.Red)

        End Select

    End Sub

End Module
