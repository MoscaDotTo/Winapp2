'    Copyright (C) 2018-2024 Hazel Ward
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
''' <summary> A sub menu of the Downloader module for advanced users who want more than just the CCleaner flavor of winapp2.ini from the repo </summary>
Public Module advDownloads

    ''' <summary> Prints the advanced downloads menu </summary>
    ''' Docs last updated: 2020-09-14 | Code last updated: 2020-09-14
    Public Sub printAdvMenu()
        printMenuTop({"Warning!", "Files in this menu are not recommended for use by beginners."})
        print(1, "Winapp3.ini", "Extended and/or potentially unsafe entries")
        print(1, "Archived entries.ini", "Entries for old or discontinued software")
        print(1, "Java.ini", "Used to generate a winapp2.ini entry that cleans up after the Java installer", closeMenu:=True)
    End Sub

    ''' <summary> Handles the user input for the advanced download menu </summary>
    ''' <param name="input">The user's input </param>
    ''' Docs last updated: 2020-09-14 | Code last updated: 2020-09-14
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
            Case "3"
                downloadFile.Name = "java.ini"
                download(downloadFile, javaLink)
            Case Else
                setHeaderText(invInpStr, True)
        End Select
    End Sub

End Module
