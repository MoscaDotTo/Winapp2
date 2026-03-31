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
''' A sub menu of the Downloader module for advanced users who want more than just the CCleaner flavor of winapp2.ini from the repo
''' </summary>
Public Module advDownloads

    ''' <summary>
    ''' Builds the advanced downloads menu with all options and their dispatch handlers registered inline
    ''' </summary>
    Private Function buildAdvMenu() As MenuSection

        Dim menuDesc = {"Warning!", "Files in this menu are not recommended for use by beginners!"}

        Return MenuSection.CreateCompleteMenu("Advanced Downloads", menuDesc, ConsoleColor.DarkCyan) _
            .AddDispatchedOption("Winapp3.ini", "Extended and/or potentially unsafe entries",
                Sub()
                    downloadFile.Name = "winapp3.ini"
                    download(downloadFile, wa3link)
                End Sub) _
            .AddDispatchedOption("Archived entries.ini", "Entries for old or discontinued software",
                Sub()
                    downloadFile.Name = "Archived entries.ini"
                    download(downloadFile, archivedLink)
                End Sub)

    End Function

    ''' <summary>
    ''' Prints the advanced downloads menu to the user
    ''' </summary>
    Public Sub printAdvMenu()

        buildAdvMenu().Print()

    End Sub

    ''' <summary>
    ''' Handles user input for the advanced downloads menu
    ''' </summary>
    '''
    ''' <param name="input">
    ''' The user's input
    ''' </param>
    Public Sub handleAdvInput(input As String)

        Dim intInput As Integer

        If Not Integer.TryParse(input, intInput) Then

            setNextMenuHeaderText(invInpStr, printColor:=ConsoleColor.Red)
            Return

        End If

        If intInput = 0 Then exitModule() : Return

        If Not buildAdvMenu().Dispatch(intInput) Then setNextMenuHeaderText(invInpStr, printColor:=ConsoleColor.Red)

    End Sub

End Module
