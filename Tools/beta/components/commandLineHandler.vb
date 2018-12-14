'    Copyright (C) 2018 Robbie Ward
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

    'File Handlers
    Dim firstFile As iniFile
    Dim secondFile As iniFile
    Dim thirdFile As iniFile

    'The current state of the command line args at any point
    Dim args As New List(Of String)

    ''' <summary>
    ''' Flips a boolean setting and removes its associated argument from the args list
    ''' </summary>
    ''' <param name="setting">The boolean setting to be flipped</param>
    ''' <param name="arg">The string containing the argument that flips the boolean</param>
    Private Sub invertSettingAndRemoveArg(ByRef setting As Boolean, ByRef arg As String)
        If args.Contains(arg) Then
            setting = Not setting
            args.Remove(arg)
        End If
    End Sub

    ''' <summary>
    ''' Calls invertSettingAndRemoveArg and assigns a new name to the given reference variable
    ''' </summary>
    ''' <param name="setting">The boolean setting</param>
    ''' <param name="arg">The arg controlling the setting</param>
    ''' <param name="name">A reference to a String to be renamed</param>
    ''' <param name="newname">The new value for the String to be renamed</param>
    Private Sub invertAndRemoveAndRename(ByRef setting As Boolean, ByRef arg As String, ByRef name As String, newname As String)
        invertSettingAndRemoveArg(setting, arg)
        name = newname
    End Sub

    ''' <summary>
    ''' Processes whether to download and which file to download
    ''' </summary>
    ''' <param name="download">The boolean indicating winapp2.ini should be downloaded</param>
    ''' <param name="ncc">The boolean indicating that the non-ccleaner variant of winapp2.ini should be used</param>
    Private Sub handleDownloadBools(ByRef download As Boolean, ByRef ncc As Boolean)
        'Download a winapp2 to trim?
        invertSettingAndRemoveArg(download, "-d")
        'Download the non ccleaner ini? 
        invertSettingAndRemoveArg(ncc, "-ncc")
        '-ncc implies -d 
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
        'Get default params
        initDebugParams(firstFile, secondFile, correctErrors)
        'Toggle on autocorrect (off by default)
        invertSettingAndRemoveArg(correctErrors, "-c")
        'Get any file parameters from flags
        getFileAndDirParams()
        'Initialize the debug
        remoteDebug(firstFile, secondFile, correctErrors)
    End Sub

    ''' <summary>
    ''' Handles the commandline args for Trim
    ''' </summary>
    ''' Trim commandline args
    ''' -d     : download the latest winapp2.ini
    ''' -ncc   : download the latest non-ccleaner winapp2.ini (implies -d)
    Private Sub autoTrim()
        Dim download As Boolean
        Dim ncc As Boolean
        initTrimParams(firstFile, secondFile, download, ncc)
        'Are we downloading winapp2.ini?
        handleDownloadBools(download, ncc)
        'Get any file parameters from flags
        getFileAndDirParams()
        'Initalize the trim
        remoteTrim(firstFile, secondFile, download, ncc)
    End Sub

    ''' <summary>
    ''' Handles the commandline args for Merge
    ''' </summary>
    '''  Merge specific command line args
    ''' -mm      : toggle mergemode from replace and add to replace and remove
    ''' -r       : use removed entries.ini as the merge file's name
    ''' -c       : use custom.ini as the merge file's name
    Private Sub autoMerge()
        Dim mergeMode As Boolean
        initMergeParams(firstFile, secondFile, thirdFile, mergeMode)
        invertSettingAndRemoveArg(mergeMode, "-mm")
        'Detect any preset filename calls
        invertAndRemoveAndRename(False, "-r", secondFile.name, "Removed Entries.ini")
        invertAndRemoveAndRename(False, "-c", secondFile.name, "Custom.ini")
        'Get any file parameters from flags
        getFileAndDirParams()
        'If we have a secondfile, initiate the merge
        If Not secondFile.name = "" Then remoteMerge(firstFile, secondFile, thirdFile, mergeMode)
    End Sub

    ''' <summary>
    ''' Handles the commandline args for Diff
    ''' </summary>
    '''  Diff commandline args
    ''' -d          : download the latest winapp2.ini
    ''' -ncc        : download the latest non-ccleaner winapp2.ini (implies -d)
    ''' -savelog    : save the diff.txt log
    Private Sub autoDiff()
        Dim ncc As Boolean
        Dim download As Boolean
        Dim save As Boolean
        initDiffParams(firstFile, secondFile, thirdFile, download, ncc, save)
        'Downloading winapp2.ini?
        handleDownloadBools(download, ncc)
        'If downloading, set secondFile's name correctly
        If download Then secondFile.name = If(ncc, "Online non-ccleaner winapp2.ini", "Online winapp2.ini")
        'Save diff.txt?
        invertSettingAndRemoveArg(save, "-savelog")
        'Get any file parameters from flags
        getFileAndDirParams()
        'Only run if we have a second file (because we assume we're running on a winapp2.ini file by default)
        If Not secondFile.name = "" Then remoteDiff(firstFile, secondFile, thirdFile, download, ncc, save)
    End Sub

    ''' <summary>
    ''' Handles the commandline args for CCiniDebug
    ''' </summary>
    '''  CCiniDebug commandline args:
    ''' -noprune : disable pruning of stale winapp2.ini entries
    ''' -nosort  : disable sorting ccleaner.ini alphabetically
    ''' -nosave  : disable saving the modified ccleaner.ini back to file
    Private Sub autoccini()
        Dim prune As Boolean
        Dim save As Boolean
        Dim sort As Boolean
        initCCDebugParams(firstFile, secondFile, thirdFile, prune, save, sort)
        'Check any provided booleans
        invertSettingAndRemoveArg(prune, "-noprune")
        invertSettingAndRemoveArg(sort, "-nosort")
        invertSettingAndRemoveArg(save, "-nosave")
        'Get any file parameters from flags
        getFileAndDirParams()
        'run ccinidebug
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
            Select Case args(0)
                Case "1", "2"
                    fileLink = If(args(0) = "1", wa2Link, nonccLink)
                    downloadName = "winapp2.ini"
                Case "3"
                    fileLink = toolLink
                    downloadName = "winapp2ool.exe"
                Case "4"
                    fileLink = removedLink
                    downloadName = "Removed Entries.ini"
            End Select
            args.RemoveAt(0)
        End If
        'Get any file parameters from flags
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
            name = "\" & args(ind + 1)
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
        'Start either a blank path or, support appending children folders to the current path
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
        'Remove the first arg which is simply the name of the executable
        args.RemoveAt(0)
        'The s is for silent, if we havxe this flag, don't give any output to the user under normal circumstances 
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
        'Make sure the args are formatted correctly so the path/name can be extracted 
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
        Dim argStr As String = "-" & someNumber
        If args.Contains(argStr & "d") Then getFileNameAndDir(argStr & "d", someFile)
        If args.Contains(argStr & "f") Then getFileName(argStr & "f", someFile.name)
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