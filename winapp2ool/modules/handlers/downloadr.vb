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
''' <summary>Provides functions to download files to the rest of the tool</summary>
Module downloadr

    ''' <summary>Performs the download, returns a boolean indicating the success of the download.</summary>
    ''' <param name="link">The URL to be downloaded from</param>
    ''' <param name="path">The file path to be saved to</param>
    Public Function dlFile(link As String, path As String) As Boolean
        Try
            Dim dl As New WebClient
            dl.DownloadFile(New Uri(link), path)
            dl.Dispose()
            Return True
        Catch ex As Exception
            exc(ex)
            Return False
        End Try
    End Function

    ''' <summary>Prompts the user to rename or overwrite a file if necessary before downloading.</summary>
    ''' <param name="link">The URL to be downloaded from</param>
    ''' <param name="prompt">The boolean indicating whether or not the user should be prompted to rename the file should it exist already.</param>
    ''' <param name="quietly">Optional boolean indicating all output should be suppressed (default: false)</param>
    Public Sub download(pathHolder As iniFile, link As String, Optional prompt As Boolean = True, Optional quietly As Boolean = False)
        Dim givenName = pathHolder.Name
        ' Don't try to download to a directory that doesn't exist
        If Not Directory.Exists(pathHolder.Dir) Then Directory.CreateDirectory(pathHolder.Dir)
        ' If the file exists and we're prompting or overwrite, do that.
        If prompt And File.Exists(pathHolder.Path) And Not SuppressOutput And Not quietly Then
            cwl($"{pathHolder.Name} already exists in the target directory.")
            Console.Write("Enter a new file name, or leave blank to overwrite the existing file: ")
            Dim nfilename = Console.ReadLine()
            If nfilename.Trim <> "" Then pathHolder.Name = nfilename
        End If
        If Not prompt And File.Exists(pathHolder.Path) Then File.Delete(pathHolder.Path)
        cwl($"Downloading {givenName}...", Not quietly)
        Dim success = dlFile(link, pathHolder.Path)
        cwl($"Download {If(success, "Complete.", "Failed.")}", Not quietly)
        cwl(If(success, "Downloaded ", $"Unable to download {pathHolder.Name} to {pathHolder.Dir}"), Not quietly)
        setHeaderText($"Download {If(success, "", "in")}complete: {pathHolder.Name}", Not success, Not quietly)
        If Not success Then Console.ReadLine()
    End Sub

    ''' <summary>Reads a file until a specified line number0</summary>
    ''' <param name="path">The path (or online address) of the file</param>
    ''' <param name="lineNum">The line number to read to</param>
    ''' <param name="remote">The boolean specifying whether the resource is remote (online)</param>
    Public Function getFileDataAtLineNum(path As String, Optional lineNum As Integer = 1, Optional remote As Boolean = True) As String
        Dim out As String
        Try
            Dim reader = If(remote, New StreamReader(New WebClient().OpenRead(path)), New StreamReader(path))
            out = getTargetLine(reader, lineNum)
            reader.Close()
        Catch ex As Exception
            exc(ex)
            Return ""
        End Try
        Return If(Not out = Nothing, out, "")
    End Function

    ''' <summary>Returns true if we are able to connect to the internet, otherwise, returns false.</summary>
    Public Function checkOnline() As Boolean
        Dim reader As StreamReader
        Try
            reader = New StreamReader(New WebClient().OpenRead("http://www.github.com"))
            gLog("Established connection to GitHub")
            reader.Close()
            Return True
        Catch ex As Exception
            exc(ex)
            Return False
        End Try
    End Function

    ''' <summary>Reads a file only until reaching a specific line and then returns that line as a String</summary>
    ''' <param name="reader">An open file stream</param>
    ''' <param name="lineNum">The target line number</param>
    Private Function getTargetLine(reader As StreamReader, lineNum As Integer) As String
        Dim out = ""
        Dim curLine = 1
        While curLine <= lineNum
            out = reader.ReadLine()
            curLine += 1
        End While
        Return out
    End Function


    ''' <summary>Returns an iniFile object created using an online resource ie. GitHub</summary>
    ''' <param name="address">A URL pointing to an online .ini file</param>
    Public Function getRemoteIniFile(address As String, Optional ByRef someFile As iniFile = Nothing) As iniFile
        Try
            Dim client As New WebClient
            Dim reader = New StreamReader(client.OpenRead(address))
            Dim wholeFile = reader.ReadToEnd
            wholeFile += Environment.NewLine
            Dim splitFile = wholeFile.Split(CChar(Environment.NewLine))
            ' Workaround for java.ini until the underlying reason for mismatches between line endings can be discovered
            If address = javaLink Then splitFile = wholeFile.Split(CChar(vbLf))
            For i = 0 To splitFile.Count - 1
                splitFile(i) = splitFile(i).Replace(vbCr, "").Replace(vbLf, "")
            Next
            reader.Close()
            Dim someFileExists = someFile IsNot Nothing
            Dim out = New iniFile(splitFile) With {.Dir = If(someFileExists, someFile.Dir, Environment.CurrentDirectory),
                                                   .InitDir = If(someFileExists, someFile.InitDir, Environment.CurrentDirectory),
                                                   .mustExist = If(someFileExists, someFile.mustExist, False),
                                                   .Name = If(someFileExists, someFile.Name, ""),
                                                   .InitName = If(someFileExists, someFile.InitName, ""),
                                                   .SecondName = If(someFileExists, someFile.SecondName, "")}
            reader.Close()
            client.Dispose()
            Return out
        Catch ex As Exception
            exc(ex)
            Return Nothing
        End Try
    End Function
End Module
