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

''' <summary>
''' This module handles the commandline args presented to winapp2ool and attempts to pass them off to their respective modules
''' </summary>
Module commandLineHandler
    ' File Handlers
    Dim firstFile As iniFile
    Dim secondFile As iniFile
    Dim thirdFile As iniFile
    ' The current state of the command line args at any point
    Dim args As New List(Of String)

    ''' <summary>
    ''' Flips a boolean setting and removes its associated argument from the args list
    ''' </summary>
    ''' <param name="setting">The boolean setting to be flipped</param>
    ''' <param name="arg">The string containing the argument that flips the boolean</param>
    ''' <param name="name">Optional reference to a file name to be replaced</param>
    ''' <param name="newname">Optional replacement file name</param>
    Private Sub invertSettingAndRemoveArg(ByRef setting As Boolean, ByRef arg As String, Optional ByRef name As String = "", Optional ByRef newname As String = "")
        If args.Contains(arg) Then
            setting = Not setting
            args.Remove(arg)
        End If
        name = newname
    End Sub

    ''' <summary>
    ''' Processes whether to download and which file to download
    ''' </summary>
    ''' <param name="download">The boolean indicating winapp2.ini should be downloaded</param>
    ''' <param name="ncc">The boolean indicating that the non-ccleaner variant of winapp2.ini should be used</param>
    Private Sub handleDownloadBools(ByRef download As Boolean, ByRef ncc As Boolean)
        ' Download a winapp2 to trim?
        invertSettingAndRemoveArg(download, "-d")
        invertSettingAndRemoveArg(ncc, "-ncc")
        ' -ncc implies -d 
        If ncc And Not download Then download = True
        If download And isOffline Then printErrExit("Winapp2ool is currently in offline mode, but you have issued commands that require a network connection. Please try again with a network connection.")
    End Sub

    ''' <summary>
    ''' Handles the commandline args for WinappDebug
    ''' </summary>
    ''' WinappDebug specific command line args
    ''' -c      : enable autocorrect
    Private Sub autoDebug()
        Dim correctErrors As Boolean
        initDebugParams(firstFile, secondFile, correctErrors)
        'Toggle on autocorrect (off by default)
        invertSettingAndRemoveArg(correctErrors, "-c")
        getFileAndDirParams()
        remoteDebug(firstFile, secondFile, correctErrors)
    End Sub

    ''' <summary>
    ''' Handles the commandline args for Trim
    ''' </summary>
    ''' Trim args:
    ''' -d     : download the latest winapp2.ini
    ''' -ncc   : download the latest non-ccleaner winapp2.ini (implies -d)
    Private Sub autoTrim()
        Dim download As Boolean
        Dim ncc As Boolean
        initTrimParams(firstFile, secondFile, download, ncc)
        handleDownloadBools(download, ncc)
        getFileAndDirParams()
        remoteTrim(firstFile, secondFile, download, ncc)
    End Sub

    ''' <summary>
    ''' Handles the commandline args for Merge
    ''' </summary>
    '''  Merge args:
    ''' -mm      : toggle mergemode from add&amp;replace to add&amp;remove
    ''' Preset merge file choices
    ''' -r       : removed entries.ini 
    ''' -c       : custom.ini 
    ''' -w       : winapp3.ini
    ''' -a       : archived entries.ini 
    Private Sub autoMerge()
        Dim mergeMode As Boolean
        initMergeParams(firstFile, secondFile, thirdFile, mergeMode)
        invertSettingAndRemoveArg(mergeMode, "-mm")
        invertSettingAndRemoveArg(False, "-r", secondFile.name, "Removed Entries.ini")
        invertSettingAndRemoveArg(False, "-c", secondFile.name, "Custom.ini")
        invertSettingAndRemoveArg(False, "-w", secondFile.name, "winapp3.ini")
        invertSettingAndRemoveArg(False, "-a", secondFile.name, "Archived Entries.ini")
        getFileAndDirParams()
        If Not secondFile.name = "" Then remoteMerge(firstFile, secondFile, thirdFile, mergeMode)
    End Sub

    ''' <summary>
    ''' Handles the commandline args for Diff
    ''' </summary>
    '''  Diff args:
    ''' -d          : download the latest winapp2.ini
    ''' -ncc        : download the latest non-ccleaner winapp2.ini (implies -d)
    ''' -savelog    : save the diff.txt log
    Private Sub autoDiff()
        Dim ncc As Boolean
        Dim download As Boolean
        Dim save As Boolean
        initDiffParams(firstFile, secondFile, thirdFile, download, ncc, save)
        handleDownloadBools(download, ncc)
        If download Then secondFile.name = If(ncc, "Online non-ccleaner winapp2.ini", "Online winapp2.ini")
        invertSettingAndRemoveArg(save, "-savelog")
        getFileAndDirParams()
        If Not secondFile.name = "" Then remoteDiff(firstFile, secondFile, thirdFile, download, ncc, save)
    End Sub

    ''' <summary>
    ''' Handles the commandline args for CCiniDebug
    ''' </summary>
    '''  CCiniDebug args:
    ''' -noprune : disable pruning of stale winapp2.ini entries
    ''' -nosort  : disable sorting ccleaner.ini alphabetically
    ''' -nosave  : disable saving the modified ccleaner.ini back to file
    Private Sub autoccini()
        Dim prune As Boolean
        Dim save As Boolean
        Dim sort As Boolean
        initCCDebugParams(firstFile, secondFile, thirdFile, prune, save, sort)
        invertSettingAndRemoveArg(prune, "-noprune")
        invertSettingAndRemoveArg(sort, "-nosort")
        invertSettingAndRemoveArg(save, "-nosave")
        getFileAndDirParams()
        remoteCC(firstFile, secondFile, thirdFile, prune, save, sort)
    End Sub

    ''' <summary>
    ''' Processes the commandline args for Downloader
    ''' </summary>
    Private Sub autodownload()
        Dim fileLink As String = ""
        Dim downloadDir As String = Environment.CurrentDirectory
        Dim downloadName As String = ""
        If args.Count > 0 Then
            Select Case args(0).ToLower
                Case "1", "2", "winapp2"
                    fileLink = If(Not args.Equals("2"), wa2Link, nonccLink)
                    downloadName = "winapp2.ini"
                Case "3", "winapp2ool"
                    fileLink = toolLink
                    downloadName = "winapp2ool.exe"
                Case "4", "removed"
                    fileLink = removedLink
                    downloadName = "Removed Entries.ini"
                Case "5", "winapp3"
                    fileLink = wa3link
                    downloadName = "winapp3.ini"
                Case "6", "archived"
                    fileLink = archivedLink
                    downloadName = "Archived Entries.ini"
                Case "7", "java"
                    fileLink = javaLink
                    downloadName = "java.ini"
            End Select
            args.RemoveAt(0)
        End If
        getFileAndDirParams()
        If Not fileLink = "" Then remoteDownload(downloadDir, downloadName, fileLink, False)
    End Sub

    ''' <summary>
    ''' Renames an iniFile object if provided a commandline arg to do so
    ''' </summary>
    ''' <param name="flag">The flag that precedes the name specification in the args list</param>
    ''' <param name="name">The reference to the name parameter of an iniFile object</param>
    Private Sub getFileName(flag As String, ByRef name As String)
        If args.Count >= 2 Then
            Dim ind As Integer = args.IndexOf(flag)
            name = $"\{args(ind + 1)}"
            args.RemoveAt(ind)
            args.RemoveAt(ind)
        End If
    End Sub

    ''' <summary>
    ''' Applies a new directory and name to an iniFile object 
    ''' </summary>
    ''' <param name="flag">The flag preceeding the file/path parameter in the arg list</param>
    ''' <param name="file">The iniFile object to be modified</param>
    Private Sub getFileNameAndDir(flag As String, ByRef file As iniFile)
        If args.Count >= 2 Then
            Dim ind As Integer = args.IndexOf(flag)
            getFileParams(args(ind + 1), file)
            args.RemoveAt(ind)
            args.RemoveAt(ind)
        End If
    End Sub

    ''' <summary>
    ''' Takes in a full form filepath with directory and assigns the directory and filename components to the given iniFile object
    ''' </summary>
    ''' <param name="arg">The filepath argument</param>
    ''' <param name="file">The iniFile object to be modified</param>
    Private Sub getFileParams(ByRef arg As String, ByRef file As iniFile)
        ' Start either a blank path or, support appending children folders to the current path
        file.dir = If(arg.StartsWith("\"), Environment.CurrentDirectory & "\", "")
        Dim splitArg As String() = arg.Split(CChar("\"))
        If splitArg.Count >= 2 Then
            For i As Integer = 0 To splitArg.Count - 2
                file.dir += splitArg(i) & "\"
            Next
        End If
        file.name = splitArg.Last
    End Sub

    ''' <summary>
    ''' Initializes the processing of the commandline args and hands the remaining arguments off to the respective module's handler
    ''' </summary>
    Public Sub processCommandLineArgs()
        args.AddRange(Environment.GetCommandLineArgs)
        args.RemoveAt(0)
        ' The s is for silent, if we havxe this flag, don't give any output to the user under normal circumstances 
        invertSettingAndRemoveArg(suppressOutput, "-s")
        If args.Count > 0 Then
            Select Case args(0)
                Case "1", "-1"
                    args.RemoveAt(0)
                    autoDebug()
                Case "2", "-2"
                    args.RemoveAt(0)
                    autoTrim()
                Case "3", "-3"
                    args.RemoveAt(0)
                    autoMerge()
                Case "4", "-4"
                    args.RemoveAt(0)
                    autoDiff()
                Case "5", "-5"
                    args.RemoveAt(0)
                    autoccini()
                Case "6", "-6"
                    args.RemoveAt(0)
                    autodownload()
            End Select
        End If
    End Sub

    ''' <summary>
    ''' Gets the directory and name info for each file given (if any)
    ''' </summary>
    Private Sub getFileAndDirParams()
        validateArgs()
        getParams(1, firstFile)
        getParams(2, secondFile)
        getParams(3, thirdFile)
    End Sub

    ''' <summary>
    ''' Processes numerically ordered directory (d) and file (f) commandline args on a per-file basis
    ''' </summary>
    ''' <param name="someNumber">The number (1-indexed) of our current internal iniFile</param>
    ''' <param name="someFile">A reference to the iniFile object being operated on</param>
    Private Sub getParams(someNumber As Integer, someFile As iniFile)
        Dim argStr As String = $"-{someNumber}"
        If args.Contains($"{argStr}d") Then getFileNameAndDir($"{argStr}d", someFile)
        If args.Contains($"{argStr}f") Then getFileName($"{argStr}f", someFile.name)
    End Sub

    ''' <summary>
    ''' Enforces that commandline args are properly formatted in {"-flag","data"} format
    ''' </summary>
    Private Sub validateArgs()
        Dim vArgs As String() = {"-1d", "-1f", "-2d", "-2f", "-3d", "-3f"}
        If args.Count > 1 Then
            Try
                For i As Integer = 0 To args.Count - 2
                    If Not vArgs.Contains(args(i)) Or
                        (Not Directory.Exists(args(i + 1)) And Not Directory.Exists($"{Environment.CurrentDirectory}\{args(i + 1)}") And
                                Not File.Exists(args(i + 1)) And File.Exists($"{Environment.CurrentDirectory}\{args(i + 1)}")) Then
                        printErrExit("Invalid command line arguements given.")
                    End If
                    i += 1
                Next
            Catch ex As Exception
                printErrExit("Invalid command line arguements given.")
            End Try
        End If
    End Sub

    ''' <summary>
    ''' Prints an error to the user and exits the application after they have pressed a key
    ''' </summary>
    ''' <param name="errTxt">The String containing the erorr text to be printed to the user</param>
    Private Sub printErrExit(errTxt As String)
        Console.WriteLine($"{errTxt} Press any key to exit.")
        Console.ReadKey()
        Environment.Exit(0)
    End Sub
End Module