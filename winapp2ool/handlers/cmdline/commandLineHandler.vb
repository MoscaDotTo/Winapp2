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

Imports System.IO

''' <summary>
''' This module helps handle any commandline arguments passed to winapp2ool
''' </summary>
Public Module commandLineHandler

    ''' <summary>
    ''' The current list of the command line args (mutable)
    ''' </summary>
    ''' 
    ''' Docs last updated: 2023-07-20 | Code last updated: 2023-07-20
    Public Property cmdargs As List(Of String)

    ''' <summary>
    ''' Flips a boolean setting and removes its associated argument from the args list
    ''' </summary>
    ''' 
    ''' <param name="setting">
    ''' A boolean module setting whose state will be inverted 
    ''' </param>
    ''' 
    ''' <param name="arg">
    ''' A commandline argument targeting <c> <paramref name="setting"/> </c>
    ''' </param>
    ''' 
    ''' <param name="name">
    ''' A File name to be modified if the arg is found
    ''' <br /> Optional, default: <c> "" </c>
    ''' </param>
    ''' 
    ''' <param name="newname">
    ''' The nenw name with which <c> <paramref name="name"/> </c> will be replaced if the arg is found 
    ''' <br /> Optional, default: <c> "" </c>
    ''' </param>
    ''' 
    ''' Docs last updated: 2023-07-20 | Code last updated: 2023-07-20
    Public Sub invertSettingAndRemoveArg(ByRef setting As Boolean, arg As String, Optional ByRef name As String = "", Optional ByRef newname As String = "")

        If Not cmdargs.Contains(arg) Then Return

        gLog($"Found argument: {arg}")
        setting = Not setting
        cmdargs.Remove(arg)
        name = newname

    End Sub

    ''' <summary>
    ''' Processes whether to download and which file to download
    ''' </summary>
    ''' 
    ''' <param name="download">
    ''' Indicates that winapp2.ini should be downloaded from GitHub 
    ''' </param>
    ''' 
    ''' Docs last updated: 2023-07-20 | Code last updated: 2023-07-20
    Public Sub handleDownloadBools(ByRef download As Boolean)

        invertSettingAndRemoveArg(download, "-d")

        Dim networkErr = "Winapp2ool is currently in offline mode, but you have issued commands that require a network connection. Please try again with a network connection."
        If download And isOffline Then printErrExit(networkErr)

    End Sub

    ''' <summary>
    ''' Renames or modifies the filename of an <c> iniFile </c> via the commandline 
    ''' </summary>
    ''' 
    ''' <param name="arg">
    ''' Commandline arg pointing to some particular file in a module 
    ''' <br /> eg. -2f or -1d 
    ''' </param>
    ''' 
    ''' <param name="givenFile">
    ''' An <c> iniFile </c> whose path will be modified 
    ''' </param>
    ''' 
    ''' <remarks> 
    ''' Supports appending child folders to the current directory 
    ''' <br /> eg. -2f "\folder1\folder2\file.ini"
    ''' </remarks>
    ''' 
    ''' Docs last updated: 2023-07-20 | Code last updated: 2023-07-20
    Private Sub getFileName(arg As String, ByRef givenFile As iniFile)

        If cmdargs.Count < 2 Then Return

        Dim ind = cmdargs.IndexOf(arg)
        Dim curArg = cmdargs(ind + 1)

        givenFile.Name = curArg

        If curArg.StartsWith("\", StringComparison.InvariantCulture) AndAlso Not curArg.LastIndexOf("\", StringComparison.InvariantCulture) = 0 Then

            Dim split = curArg.Split(CChar("\"))

            For i As Integer = 1 To split.Length - 2

                givenFile.Dir += $"\{split(i)}"

            Next

            givenFile.Name = split.Last

        End If

        cmdargs.RemoveAt(ind)
        cmdargs.RemoveAt(ind)

    End Sub

    ''' <summary>
    ''' Applies a new directory and name to an iniFile object
    ''' </summary>
    ''' 
    ''' <param name="flag">
    ''' The flag preceeding the file/path parameter in the arg list
    ''' </param>
    ''' 
    ''' <param name="file">
    ''' The iniFile object to be modified
    ''' </param>
    ''' 
    ''' Docs last updated: 2023-07-20 | Code last updated: 2023-07-20
    Private Sub getFileNameAndDir(flag As String, ByRef file As iniFile)

        If cmdargs.Count < 2 Then Return

        Dim ind As Integer = cmdargs.IndexOf(flag)
        getFileParams(cmdargs(ind + 1), file)
        cmdargs.RemoveAt(ind)
        cmdargs.RemoveAt(ind)

    End Sub

    ''' <summary>Takes in a full form filepath with directory and assigns the directory and filename components to the given iniFile object</summary>
    ''' <param name="arg">The filepath argument</param>
    ''' <param name="file">The iniFile object to be modified</param>
    Private Sub getFileParams(ByRef arg As String, ByRef file As iniFile)

        file.Dir = If(arg.StartsWith("\", StringComparison.InvariantCulture), Environment.CurrentDirectory & "\", "")
        Dim splitArg As String() = arg.Split(CChar("\"))

        file.Name = splitArg.Last

        If splitArg.Length < 2 Then Return

        For i As Integer = 0 To splitArg.Length - 2

            file.Dir += splitArg(i) & "\"

        Next

    End Sub

    ''' <summary>Initializes the processing of the commandline args and hands the remaining arguments off to the respective module's handler</summary>
    Public Sub processCommandLineArgs()
        cmdargs = Environment.GetCommandLineArgs.ToList
        Dim argStr = String.Join(",", cmdargs)
        gLog($"Found commandline args: {argStr}")
        ' 0th index holds the executable name, we don't need it. 
        cmdargs.RemoveAt(0)
        checkUpdates(cmdargs.Contains("-autoupdate"))
        If updateIsAvail Then autoUpdate()
        ' The s is for silent, if we have this flag, don't give any output or ask for input along the happy path
        invertSettingAndRemoveArg(SuppressOutput, "-s")
        ' Toggle the tool to use the non-ccleaner version of winapp2.ini
        invertSettingAndRemoveArg(RemoteWinappIsNonCC, "-ncc")
        If cmdargs.Count > 0 Then
            Select Case cmdargs(0)
                Case "1", "-1", "debug", "-debug"
                    cmdargs.RemoveAt(0)
                    HandleLintCmdLine()
                Case "2", "-2", "trim", "-trim"
                    cmdargs.RemoveAt(0)
                    Trim.handleCmdLine()
                Case "3", "-3", "merge", "-merge"
                    cmdargs.RemoveAt(0)
                    Merge.handleCmdLine()
                Case "4", "-4", "diff", "-diff"
                    cmdargs.RemoveAt(0)
                    Diff.HandleCmdLine()
                Case "5", "-5", "ccdebug", "-ccdebug"
                    cmdargs.RemoveAt(0)
                    CCiniDebug.handleCmdlineArgs()
                Case "6", "-6", "download", "-download"
                    cmdargs.RemoveAt(0)
                    Downloader.handleCmdLine()
            End Select
        End If
    End Sub

    ''' <summary>Gets the directory and name info for each file given (if any)</summary>
    ''' <param name="ff">The "first file" from a module</param>
    ''' <param name="sf">The "second file" from a module</param>
    ''' <param name="tf">The "third file" from a module</param>
    Public Sub getFileAndDirParams(ByRef ff As iniFile, ByRef sf As iniFile, ByRef tf As iniFile)
        If ff Is Nothing Then argIsNull(NameOf(ff)) : Return
        If sf Is Nothing Then argIsNull(NameOf(sf)) : Return
        If tf Is Nothing Then argIsNull(NameOf(tf)) : Return
        validateArgs()
        getParams(1, ff)
        getParams(2, sf)
        getParams(3, tf)
    End Sub

    ''' <summary>Processes numerically ordered directory (d) and file (f) commandline args on a per-file basis</summary>
    ''' <param name="someNumber">The number (1-indexed) of our current internal iniFile</param>
    ''' <param name="someFile">A reference to the iniFile object being operated on</param>
    Private Sub getParams(someNumber As Integer, ByRef someFile As iniFile)
        Dim argStr As String = $"-{someNumber}"
        If cmdargs.Contains($"{argStr}d") Then getFileNameAndDir($"{argStr}d", someFile)
        If cmdargs.Contains($"{argStr}f") Then getFileName($"{argStr}f", someFile)
        ' Make sure there's there's no double slashes or leading/trailing slashes in the file parameters
        someFile.Dir = someFile.Dir.Replace(CChar("\\"), CChar("\"))
        If someFile.Name.StartsWith("\", StringComparison.InvariantCulture) Then someFile.Name = someFile.Name.TrimStart(CChar("\"))
        If someFile.Dir.EndsWith("\", StringComparison.InvariantCulture) Then someFile.Dir = someFile.Dir.TrimEnd(CChar("\"))
    End Sub

    ''' <summary>Enforces that commandline args are properly formatted in {"-flag","data"} format</summary>
    Private Sub validateArgs()
        Dim vArgs As String() = {"-1d", "-1f", "-2d", "-2f", "-3d", "-3f"}
        If cmdargs.Count > 1 Then
            Try
                For i As Integer = 0 To cmdargs.Count - 2
                    If Not vArgs.Contains(cmdargs(i)) Or
                        (Not Directory.Exists(cmdargs(i + 1)) And Not Directory.Exists($"{Environment.CurrentDirectory}\{cmdargs(i + 1)}") And
                                Not File.Exists(cmdargs(i + 1)) And File.Exists($"{Environment.CurrentDirectory}\{cmdargs(i + 1)}")) Then
                        Throw New ArgumentException("Invalid commandline arguments given.")
                    End If
                    ' Increment i here by 1 to skip the next item (i+1) since we've already processed it above
                    i += 1
                Next
            Catch ex As ArgumentException
                handleInvalidArgException(ex)
                printErrExit(ex.Message)
            End Try
        End If
    End Sub

    ''' <summary>Prints an error to the user and exits the application after they have pressed a key</summary>
    ''' <param name="errTxt">The text to be printed to the user</param>
    Private Sub printErrExit(errTxt As String)
        Console.WriteLine($"{errTxt} Press any key to exit.")
        Console.ReadKey()
        Environment.Exit(0)
    End Sub
End Module