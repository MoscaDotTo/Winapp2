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
''' Allows users to download files from GitHub through a simple menu 
''' </summary>
''' 
''' Docs last updated: 2020-09-14 | Code last updated: 2020-09-14
Module Downloader

    ''' <summary> 
    ''' The web address of the winapp2.ini GitHub 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2020-09-14 | Code last updated: 2020-09-14
    Public ReadOnly Property gitLink As String = "https://github.com/MoscaDotTo/Winapp2/"

    ''' <summary>
    ''' The web address of the CCleaner version of winapp2.ini 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2020-09-14 | Code last updated: 2020-09-14
    Public ReadOnly Property wa2Link As String = "https://raw.githubusercontent.com/MoscaDotTo/Winapp2/master/Winapp2.ini"

    ''' <summary>
    ''' The web address of the Non-CCleaner version of winapp2.ini 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2020-09-14 | Code last updated: 2020-09-14
    Public ReadOnly Property nonccLink As String = "https://raw.githubusercontent.com/MoscaDotTo/Winapp2/master/Non-CCleaner/Winapp2.ini"

    ''' <summary> 
    ''' The web address of winapp2ool.exe 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2020-09-14 | Code last updated: 2020-09-14
    Public ReadOnly Property toolLink As String = "https://github.com/MoscaDotTo/Winapp2/raw/master/winapp2ool/bin/Release/winapp2ool.exe"

    ''' <summary> 
    ''' The web address of the beta build of winapp2ool.exe 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2020-09-14 | Code last updated: 2020-09-14
    Public ReadOnly Property betaToolLink As String = "https://github.com/MoscaDotTo/Winapp2/raw/Branch1/winapp2ool/bin/Release/winapp2ool.exe"

    ''' <summary> 
    ''' The web address of version.txt (winapp2ool's public version identifer) 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2020-09-14 | Code last updated: 2020-09-14
    Public ReadOnly Property toolVerLink As String = "https://raw.githubusercontent.com/MoscaDotTo/Winapp2/master/winapp2ool/version.txt"

    ''' <summary> 
    ''' The web address of version.txt on the beta branch 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2020-09-14 | Code last updated: 2020-09-14
    Public ReadOnly Property betaToolVerLink As String = "https://raw.githubusercontent.com/MoscaDotTo/Winapp2/Branch1/winapp2ool/version.txt"

    ''' <summary> 
    ''' The web address of removed entries.ini
    ''' </summary>
    ''' 
    ''' Docs last updated: 2020-09-14 | Code last updated: 2020-09-14
    Public ReadOnly Property removedLink As String = "https://raw.githubusercontent.com/MoscaDotTo/Winapp2/master/Non-CCleaner/Removed%20entries.ini"

    ''' <summary> 
    ''' The web address of winapp3.ini 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2020-09-14 | Code last updated: 2020-09-14
    Public ReadOnly Property wa3link As String = "https://raw.githubusercontent.com/MoscaDotTo/Winapp2/master/Winapp3/Winapp3.ini"

    ''' <summary>
    ''' The web address of archived entries.ini 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2020-09-14 | Code last updated: 2020-09-14
    Public ReadOnly Property archivedLink As String = "https://raw.githubusercontent.com/MoscaDotTo/Winapp2/master/Winapp3/Archived%20entries.ini"

    ''' <summary> 
    ''' The web address of java.ini
    ''' </summary>
    ''' 
    ''' Docs last updated: 2020-09-14 | Code last updated: 2020-09-14
    Public ReadOnly Property javaLink As String = "https://raw.githubusercontent.com/MoscaDotTo/Winapp2/master/Winapp3/java.ini"

    ''' <summary> 
    ''' The web address of the winapp2ool ReadMe file 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2020-09-14 | Code last updated: 2020-09-14
    Public ReadOnly Property readMeLink As String = "https://raw.githubusercontent.com/MoscaDotTo/Winapp2/master/winapp2ool/Readme.md"

    ''' <summary> 
    ''' Handles the commandline args for the Downloader 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2020-09-14 | Code last updated: 2020-09-14
    Public Sub handleCmdLine()

        Dim fileLink = ""

        If cmdargs.Count > 0 Then

            Select Case cmdargs(0).ToUpperInvariant

                Case "1", "2", "WINAPP2"

                    fileLink = If(cmdargs(0) <> "2", wa2Link, nonccLink)
                    downloadFile.Name = "winapp2.ini"

                Case "3", "WINAPP2OOL"

                    fileLink = toolLink
                    downloadFile.Name = "winapp2ool.exe"

                Case "4", "REMOVED"

                    fileLink = removedLink
                    downloadFile.Name = "Removed Entries.ini"

                Case "5", "WINAPP3"

                    fileLink = wa3link
                    downloadFile.Name = "winapp3.ini"

                Case "6", "ARCHIVED"

                    fileLink = archivedLink
                    downloadFile.Name = "Archived Entries.ini"

                Case "7", "JAVA"

                    fileLink = javaLink
                    downloadFile.Name = "java.ini"

                Case "8", "README"

                    fileLink = readMeLink
                    downloadFile.Name = "readme.txt"

            End Select

            cmdargs.RemoveAt(0)

        End If

        getFileAndDirParams(downloadFile, New iniFile, New iniFile)

        If downloadFile.Dir = Environment.CurrentDirectory And downloadFile.Name = "winapp2ool.exe" Then autoUpdate()

        download(downloadFile, fileLink)

    End Sub

    ''' <summary> 
    ''' Returns the link to winapp2.ini of the apprpriate flavor for the current tool configuration 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2020-09-14 | Code last updated: 2020-09-14
    Public Function winapp2link() As String

        Return If(RemoteWinappIsNonCC, nonccLink, wa2Link)

    End Function

    ''' <summary>
    ''' Returns the link to winapp2ool.exe on the appropriate branch for the current tool configuration 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2020-09-14 | Code last updated: 2020-09-14
    Public Function toolExeLink() As String

        Return If(isBeta, betaToolLink, toolLink)

    End Function

    ''' <summary> 
    ''' Returns the online download status (name) of winapp2.ini as a <c> String </c>, empty string if not downloading 
    ''' </summary>
    ''' 
    ''' <param name="shouldDownload"> 
    ''' Indicates that a module is configured to download a remote winapp2.ini 
    ''' </param>
    ''' 
    ''' Docs last updated: 2020-09-14 | Code last updated: 2020-09-14
    Public Function GetNameFromDL(shouldDownload As Boolean) As String

        Return If(shouldDownload, If(RemoteWinappIsNonCC, "Online (Non-CCleaner)", "Online"), "")

    End Function

End Module