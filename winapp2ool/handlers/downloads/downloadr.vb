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

Imports System.IO
Imports System.Net

''' <summary> 
''' Downloadr is a winapp2ool module which provides the ability to download files from the 
''' internet to the rest of the application
''' </summary>
Module downloadr

    ''' <summary>
    ''' Attempts to download a file from the internet, overwriting any file already at
    ''' <paramref name="path"/>. A network failure is only logged. Any other exception goes
    ''' through <see cref="exc"/>, which marks the run failed and waits for Enter unless
    ''' output is suppressed.
    ''' </summary>
    ''' 
    ''' <param name="link">
    ''' A URL pointing to a file to be downloaded 
    ''' </param>
    ''' 
    ''' <param name="path">
    ''' The path on disk to which the downloaded file should be saved 
    ''' </param>
    ''' 
    ''' <returns> <c> True </c> if the download is successful, <br />
    ''' <c> False </c> otherwise 
    ''' </returns>
    Public Function dlFile(link As String,
                           path As String) As Boolean

        Try

            Dim dl As New WebClient
            dl.DownloadFile(New Uri(link), path)
            dl.Dispose()
            Return True

        Catch ex As WebException

            handleWebException(ex)
            Return False

        Catch e As Exception

            exc(e)
            Return False

        End Try

    End Function

    ''' <summary>
    ''' Downloads a file from the internet to the path described by an <c> iniFileChooser </c>,
    ''' creating its directory if needed. If the file already exists, we either offer to
    ''' rename the download or overwrite it. A failed download marks the run failed and waits
    ''' for Enter unless output is suppressed, even when <paramref name="quietly"/> is set.
    ''' </summary>
    '''
    ''' <param name="pathHolder">
    ''' The <c> iniFileChooser </c> describing the save location. A name the user types at the
    ''' rename prompt replaces its <c> Name </c>.
    ''' </param>
    '''
    ''' <param name="link">
    ''' A URL pointing to a file to be downloaded
    ''' </param>
    '''
    ''' <param name="prompt">
    ''' Indicates whether to offer a new file name when the target already exists. The prompt
    ''' only appears when neither output is suppressed nor <paramref name="quietly"/> is set,
    ''' and otherwise the download overwrites the file. When <c> False </c>, we delete the
    ''' existing file before downloading. <br /><br />
    ''' Optional, Default: <c> True </c>
    ''' </param>
    '''
    ''' <param name="quietly">
    ''' Indicates whether to skip the progress messages and the rename prompt <br /><br />
    ''' Optional, Default: <c> False </c>
    ''' </param>
    Public Sub download(pathHolder As iniFileChooser,
                        link As String,
               Optional prompt As Boolean = True,
               Optional quietly As Boolean = False)

        If Not Directory.Exists(pathHolder.Dir) Then Directory.CreateDirectory(pathHolder.Dir)

        If prompt AndAlso File.Exists(pathHolder.Path()) AndAlso Not SuppressOutput AndAlso Not quietly Then
            cwl($"{pathHolder.Name} already exists in the target directory.")
            Console.Write("Enter a new file name, or leave blank to overwrite the existing file: ")
            Dim nfilename = Console.ReadLine()
            If Not nfilename.Trim.Length = 0 Then pathHolder.Name = nfilename
        End If

        If Not prompt Then fDelete(pathHolder.Path())

        cwl($"Downloading {pathHolder.Name}...", Not quietly)

        Dim success = dlFile(link, pathHolder.Path())

        cwl($"Download {If(success, "Complete.", "Failed.")}", Not quietly)
        cwl($"{If(success, "Downloaded", "Unable to download")} {pathHolder.Name} to {pathHolder.Dir}", Not quietly)

        setNextMenuHeaderText($"Download {If(success, "", "in")}complete: {pathHolder.Name}", Not success AndAlso Not quietly, ConsoleColor.Red)

        If Not success Then

            markRunFailed()
            crl()

        End If

    End Sub

    ''' <summary>
    ''' Reads a file until a specified line number, returns the contents of that line.
    ''' A missing or unreadable file throws to the caller.
    ''' </summary>
    '''
    ''' <param name="path">The path of the file to read</param>
    '''
    ''' <param name="lineNum">
    ''' The 1-based line number to return from the file <br /><br />
    ''' Optional, Default: <c> 1 </c>
    ''' </param>
    '''
    ''' <returns>
    ''' The line's text, or <c> "" </c> if the file has fewer lines. We log that case without
    ''' marking the run failed.
    ''' </returns>
    Public Function getFileDataAtLineNum(path As String,
                                Optional lineNum As Integer = 1) As String

        Dim out As String = ""

        Try

            Dim reader = New StreamReader(path)
            out = getTargetLine(reader, lineNum)
            reader.Close()

            If out Is Nothing Then Throw New ArgumentException(paramName:=NameOf(lineNum), message:=$"{lineNum} didn't return any data. It may be greater than the number of lines in the file.")

        Catch ex As ArgumentException

            handleInvalidArgException(ex, False, True)
            Return ""

        End Try

        Return out

    End Function

    ''' <summary>
    ''' Attempts to open <c> https://github.com </c>. A failure is logged through
    ''' <see cref="handleWebException"/>, and only a <c> WebException </c> is caught.
    ''' </summary>
    ''' 
    ''' <returns> 
    ''' <c> True </c> If the connection is successful, <br />
    ''' <c> False </c> otherwise
    ''' </returns>
    Public Function checkOnline() As Boolean

        Try

            Dim wc As New WebClient

            gLog("Attempting to connect to GitHub")
            wc.OpenRead("https://github.com").Close()
            gLog("Established connection to GitHub")
            wc.Dispose()

            Return True

        Catch ex As WebException

            handleWebException(ex)
            Return False

        End Try

    End Function

    ''' <summary> 
    ''' Reads a file only until reaching a specific line and then returns that line as a String 
    ''' </summary>
    ''' 
    ''' <param name="reader">
    ''' An open file stream 
    ''' </param>
    ''' 
    ''' <param name="lineNum">
    ''' The target line number 
    ''' </param>
    ''' 
    ''' <returns>
    ''' The String on the line given by <paramref name="lineNum"/>, <c> Nothing </c> if the
    ''' file ends before it, or <c> "" </c> if <paramref name="lineNum"/> is less than 1
    ''' </returns>
    Private Function getTargetLine(reader As StreamReader,
                                   lineNum As Integer) As String

        Dim out = ""
        Dim curLine = 1

        While curLine <= lineNum

            out = reader.ReadLine()
            curLine += 1

        End While

        Return out

    End Function

    ''' <summary>
    ''' Returns the first line of a remote text file, reading no further than that line
    ''' </summary>
    '''
    ''' <param name="link">
    ''' A URL pointing to a text file
    ''' </param>
    '''
    ''' <returns>
    ''' The file's first line, or an empty string if the file is empty, <br />
    ''' <c> Nothing </c> if the download fails
    ''' </returns>
    Public Function getRemoteFirstLine(link As String) As String

        Try

            Using client As New WebClient

                Using reader As New StreamReader(client.OpenRead(link))

                    Return If(reader.ReadLine(), "")

                End Using

            End Using

        Catch ex As WebException

            handleWebException(ex)
            Return Nothing

        Catch ex As IOException

            handleIOException(ex)
            Return Nothing

        End Try

    End Function

    ''' <summary>
    ''' Attempts to create an <c> iniFile </c> using the data provided by <paramref name="address"/>.
    ''' The download is parsed straight from the network, without staging a copy on disk.
    ''' The result's <c> Dir </c> is the <c> %temp% </c> folder and its <c> Name </c> is the
    ''' last segment of the URL, though nothing is written there.
    ''' </summary>
    '''
    ''' <param name="address">
    ''' A URL pointing to a .ini file to be downloaded
    ''' </param>
    '''
    ''' <returns>
    ''' An <c> iniFile </c> parsed from the remote data, whatever it contains, <br />
    ''' <c> Nothing </c> if the download fails with a network or IO error
    ''' </returns>
    Public Function getRemoteIniFile(address As String) As iniFile

        Try

            Using client As New WebClient

                Using reader As New StreamReader(client.OpenRead(address))

                    Return iniFile.FromStream(reader, Environment.GetEnvironmentVariable("temp"), address.Split("/"c).Last)

                End Using

            End Using

        Catch ex As WebException

            handleWebException(ex)
            Return Nothing

        Catch ex As IOException

            handleIOException(ex)
            Return Nothing

        End Try

    End Function

End Module
