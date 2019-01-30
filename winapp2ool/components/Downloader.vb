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
Imports System.IO
Imports System.Net

''' <summary>
''' This module contains functions that allow the user to reach online resources. 
''' Its primary user-facing functionality is to present the list of downloads from the GitHub to the user
''' </summary>
Module Downloader
    ' Links to GitHub resources
    Public Const wa2Link As String = "https://raw.githubusercontent.com/MoscaDotTo/Winapp2/master/Winapp2.ini"
    Public Const nonccLink As String = "https://raw.githubusercontent.com/MoscaDotTo/Winapp2/master/Non-CCleaner/Winapp2.ini"
    Public Const toolLink As String = "https://github.com/MoscaDotTo/Winapp2/raw/master/winapp2ool/bin/Release/winapp2ool.exe"
    Public Const toolVerLink As String = "https://raw.githubusercontent.com/MoscaDotTo/Winapp2/master/winapp2ool/version.txt"
    Public Const removedLink As String = "https://raw.githubusercontent.com/MoscaDotTo/Winapp2/master/Non-CCleaner/Removed%20entries.ini"
    Public Const wa3link As String = "https://raw.githubusercontent.com/MoscaDotTo/Winapp2/master/Winapp3/Winapp3.ini"
    Public Const archivedLink As String = "https://raw.githubusercontent.com/MoscaDotTo/Winapp2/master/Winapp3/Archived%20entries.ini"
    Public Const javaLink As String = "https://raw.githubusercontent.com/MoscaDotTo/Winapp2/master/Winapp3/java.ini"
    Public Const readMeLink As String = "https://raw.githubusercontent.com/MoscaDotTo/Winapp2/master/winapp2ool/Readme.md"
    ' File handler
    Dim downloadFile As iniFile = New iniFile(Environment.CurrentDirectory, "")

    ''' <summary>
    ''' Handles the commandline args for the Downloader 
    ''' </summary>
    Public Sub handleCmdLine()
        Dim fileLink As String = ""
        If cmdargs.Count > 0 Then
            Select Case cmdargs(0).ToLower
                Case "1", "2", "winapp2"
                    fileLink = If(Not cmdargs(0) = "2", wa2Link, nonccLink)
                    downloadFile.name = "winapp2.ini"
                Case "3", "winapp2ool"
                    fileLink = toolLink
                    downloadFile.name = "winapp2ool.exe"
                Case "4", "removed"
                    fileLink = removedLink
                    downloadFile.name = "Removed Entries.ini"
                Case "5", "winapp3"
                    fileLink = wa3link
                    downloadFile.name = "winapp3.ini"
                Case "6", "archived"
                    fileLink = archivedLink
                    downloadFile.name = "Archived Entries.ini"
                Case "7", "java"
                    fileLink = javaLink
                    downloadFile.name = "java.ini"
                Case "8", "readme"
                    fileLink = readMeLink
                    downloadFile.name = "readme.txt"
            End Select
            cmdargs.RemoveAt(0)
        End If
        getFileAndDirParams(downloadFile, New iniFile, New iniFile)
        If downloadFile.dir = Environment.CurrentDirectory And downloadFile.name = "winapp2ool.exe" Then autoUpdate()
        download(fileLink)
    End Sub

    ''' <summary>
    ''' Prints the main menu to the user
    ''' </summary>
    Public Sub printMenu()
        printMenuTop({"Download files from the winapp2 GitHub"})
        print(1, "Winapp2.ini", "Download the latest winapp2.ini")
        print(1, "Non-CCleaner", "Download the latest non-ccleaner winapp2.ini")
        print(1, "Winapp2ool", "Download the latest winapp2ool.exe")
        print(1, "Removed Entries.ini", "Download only entries used to create the non-ccleaner winapp2.ini", leadingBlank:=True)
        print(1, "Directory", "Change the save directory", trailingBlank:=True)
        print(1, "Advanced", "Settings for power users")
        print(1, "ReadMe", "The winapp2ool ReadMe")
        print(0, $"Save directory: {replDir(downloadFile.dir)}", leadingBlank:=True, closeMenu:=True)
    End Sub

    ''' <summary>
    ''' Prints the advanced downloads menu
    ''' </summary>
    Private Sub printAdvMenu()
        printMenuTop({"Warning!", "Files in this menu are not recommended for use by beginners."})
        print(1, "Winapp3.ini", "Extended and/or potentially unsafe entries")
        print(1, "Archived entries.ini", "Entries for old or discontinued software")
        print(1, "Java.ini", "Used to generate a winapp2.ini entry that cleans up after the Java installer", closeMenu:=True)
    End Sub

    ''' <summary>
    ''' Handles the user input for the advanced download menu
    ''' </summary>
    ''' <param name="input">The string containing the user's input</param>
    Private Sub handleAdvInput(input As String)
        Select Case input
            Case "0"
                exitModule()
            Case "1"
                downloadFile.name = "winapp3.ini"
                download(wa3link)
            Case "2"
                downloadFile.name = "Archived entries.ini"
                download(archivedLink)
            Case "3"
                downloadFile.name = "java.ini"
                download(javaLink)
            Case Else
                menuHeaderText = invInpStr
        End Select
    End Sub

    ''' <summary>
    ''' Handles the user input for the menu
    ''' </summary>
    ''' <param name="input">The String containing the user's input</param>
    Public Sub handleUserInput(input As String)
        Select Case input
            Case "0"
                Console.WriteLine("Returning to winapp2ool menu...")
                exitCode = True
            Case "1", "2"
                downloadFile.name = "winapp2.ini"
                Dim link As String = If(input = "1", wa2Link, nonccLink)
                download(link)
                If downloadFile.dir = Environment.CurrentDirectory Then checkedForUpdates = False
            Case "3"
                ' Feature gate downloading the executable behind .NET 4.6+
                If Not denyActionWithTopper(dnfOOD, "This option requires a newer version of the .NET Framework") Then
                    If downloadFile.dir = Environment.CurrentDirectory Then
                        autoUpdate()
                    Else
                        downloadFile.name = "winapp2ool.exe"
                        download(toolLink)
                    End If
                End If
            Case "4"
                downloadFile.name = "Removed entries.ini"
                download(removedLink)
            Case "5"
                dirChooser(downloadFile.dir)
                undoAnyPendingExits()
                menuHeaderText = "Save directory changed"
            Case "6"
                initModule("Advanced Downloads", AddressOf printAdvMenu, AddressOf handleAdvInput)
            Case "7"
                ' It's actually a .md but the user doesn't need to know that  
                downloadFile.name = "Readme.txt"
                download(readMeLink)
            Case Else
                menuHeaderText = invInpStr
        End Select
    End Sub

    ''' <summary>
    ''' Feteches a remote winapp2.ini file from GitHub
    ''' </summary>
    ''' <param name="ncc">The Boolean specifying whether the non-CCleaner version should be downloaded</param>
    ''' <returns></returns>
    Public Function getRemoteWinapp(ncc As Boolean) As iniFile
        Return getRemoteIniFile(If(ncc, nonccLink, wa2Link))
    End Function

    ''' <summary>
    ''' Reads a file until a specified line number 
    ''' </summary>
    ''' <param name="path">The path (or online address) of the file</param>
    ''' <param name="lineNum">The line number to read to</param>
    ''' <param name="remote">The boolean specifying whether the resource is remote (online)</param>
    ''' <returns></returns>
    Public Function getFileDataAtLineNum(path As String, Optional lineNum As Integer = 1, Optional remote As Boolean = True) As String
        Dim reader As StreamReader = Nothing
        Dim out As String = ""
        Try
            reader = If(remote, New StreamReader(New WebClient().OpenRead(path)), New StreamReader(path))
            out = getTargetLine(reader, lineNum)
        Catch ex As Exception
            exc(ex)
            Return ""
        Finally
            reader.Close()
        End Try
        Return If(out = Nothing, "", out)
    End Function

    ''' <summary>
    ''' Returns true if we are able to connect to the internet, otherwise, returns false.
    ''' </summary>
    Public Function checkOnline() As Boolean
        Dim reader As StreamReader
        Try
            reader = New StreamReader(New WebClient().OpenRead("http://www.github.com"))
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Returns an iniFile object created using an online resource
    ''' </summary>
    ''' ie. GitHub
    ''' <param name="address">A URL pointing to an online .ini file</param>
    ''' <returns></returns>
    Public Function getRemoteIniFile(address As String) As iniFile
        Dim reader As StreamReader
        Try
            Dim client As New WebClient
            reader = New StreamReader(client.OpenRead(address))
            Dim wholeFile As String = reader.ReadToEnd
            wholeFile += Environment.NewLine
            Dim splitFile As String() = wholeFile.Split(CChar(Environment.NewLine))
            ' Workaround for java.ini until the underlying reason for mismatches between line endings can be discovered
            If address = javaLink Then splitFile = wholeFile.Split(CChar(vbLf))
            For i As Integer = 0 To splitFile.Count - 1
                splitFile(i) = splitFile(i).Replace(vbCr, "").Replace(vbLf, "")
            Next
            Return New iniFile(splitFile)
        Catch ex As Exception
            exc(ex)
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Reads a file only until reaching a specific line and then returns that line as a String
    ''' </summary>
    ''' <param name="reader">An open file stream</param>
    ''' <param name="lineNum">The target line number</param>
    ''' <returns></returns>
    Private Function getTargetLine(reader As StreamReader, lineNum As Integer) As String
        Dim out As String = ""
        Dim curLine As Integer = 1
        While curLine <= lineNum
            out = reader.ReadLine()
            curLine += 1
        End While
        Return If(out = Nothing, "", out)
    End Function

    ''' <summary>
    ''' Handles a request to download a file from outside the module
    ''' </summary>
    ''' <param name="dir">The directory to which the file should be downloaded</param>
    ''' <param name="name">The name with which to save the file</param>
    ''' <param name="link">The URL to download from</param>
    ''' <param name="prompt">Boolean specifying whether or not the user should be asked to overwrite the file should it exist</param>
    Public Sub remoteDownload(dir As String, name As String, link As String, prompt As Boolean)
        downloadFile.dir = dir
        downloadFile.name = name
        download(link, prompt)
    End Sub

    ''' <summary>
    ''' Downloads the latest version of winapp2ool.exe and replaces the currently running executable with it before launching that new executable and closing the program.
    ''' </summary>
    Public Sub autoUpdate()
        downloadFile.dir = Environment.CurrentDirectory
        downloadFile.name = "winapp2ool updated.exe"
        Dim backupName As String = $"winapp2ool v{currentVersion}.exe.bak"
        Try
            ' Remove any existing backups of this version
            If File.Exists($"{Environment.CurrentDirectory}\{backupName}") Then File.Delete($"{Environment.CurrentDirectory}\{backupName}")
            ' Remove any old update files that didn't get renamed for whatever reason
            If File.Exists(downloadFile.path) Then File.Delete(downloadFile.path)
            download(toolLink, False)
            ' Rename the executables and launch the new one
            File.Move("winapp2ool.exe", backupName)
            File.Move("winapp2ool updated.exe", "winapp2ool.exe")
            System.Diagnostics.Process.Start($"{Environment.CurrentDirectory}\winapp2ool.exe")
            Environment.Exit(0)
        Catch ex As Exception
            exc(ex)
        End Try
    End Sub

    ''' <summary>
    ''' Prompts the user to rename or overwrite a file if necessary before downloading.
    ''' </summary>
    ''' <param name="link">The URL to be downloaded from</param>
    ''' <param name="prompt">The boolean indicating whether or not the user should be prompted to rename the file should it exist already.</param>
    Private Sub download(link As String, Optional prompt As Boolean = True)
        Dim givenName As String = downloadFile.name
        ' Don't try to download to a directory that doesn't exist
        If Not Directory.Exists(downloadFile.dir) Then Directory.CreateDirectory(downloadFile.dir)
        ' If the file exists and we're prompting or overwrite, do that.
        If prompt And File.Exists(downloadFile.path) And Not suppressOutput Then
            cwl($"{downloadFile.name} already exists in the target directory.")
            Console.Write("Enter a new file name, or leave blank to overwrite the existing file: ")
            Dim nfilename As String = Console.ReadLine()
            If nfilename.Trim <> "" Then downloadFile.name = nfilename
        End If
        cwl($"Downloading {givenName}...")
        Dim success As Boolean = dlFile(link, downloadFile.path)
        cwl($"Download {If(success, "Complete.", "Failed.")}")
        cwl(If(success, "Downloaded ", $"Unable to download {downloadFile.name} to {downloadFile.dir}"))
        menuHeaderText = $"Download {If(success, "", "in")}complete: {downloadFile.name}"
        If Not success Then Console.ReadLine()
    End Sub

    ''' <summary>
    ''' Performs the download, returns a boolean indicating the success of the download.
    ''' </summary>
    ''' <param name="link">The URL to be downloaded from</param>
    ''' <param name="path">The file path to be saved to</param>
    ''' <returns></returns>
    Private Function dlFile(link As String, path As String) As Boolean
        Try
            Dim dl As New WebClient
            dl.DownloadFile(New Uri(link), path)
            Return True
        Catch ex As Exception
            exc(ex)
            Return False
        End Try
    End Function
End Module