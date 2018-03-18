Option Strict On
Module commandLineHandler

    'Handle commandline args for WinappDebug
    Private Sub autoDebug(ByRef args As List(Of String))

        'The default run parameters
        Dim filePath As String = Environment.CurrentDirectory
        Dim fileName As String = "\winapp2.ini"
        Dim correctErrors As Boolean = False

        'Toggle on autocorrect (off by default)
        If args.Contains("-c") Then
            correctErrors = True
            args.Remove("-c")
        End If

        'At this stage, any properly formatted command line args should contain only an ordered pair 
        'of a flag and either a full file path or just a file name 

        'Validate that the remainder of the args (if any) are formatted correctly
        validateArgs(args, "-d -f")

        '-d is a flag for providing a file with path
        If args(0) = "-d" Then
            args.RemoveAt(0)
            getFileParams(args(0), filePath, fileName)
        End If

        '-f is a flag for simply providing a file in the current directory
        If args(0) = "-f" Then
            args.RemoveAt(0)
            fileName = "\" & args(0)
        End If

        'Initialize the debug
        remoteDebug(filePath, fileName, correctErrors)
    End Sub

    'Handle commandline args for Trim
    Private Sub autoTrim(ByRef args As List(Of String))
        Dim winappDir As String = Environment.CurrentDirectory
        Dim winappName As String = "\winapp2.ini"
        Dim saveDir As String = Environment.CurrentDirectory
        Dim saveName As String = "\winapp2.ini"
        Dim download As Boolean = False
        Dim ncc As Boolean = False

        'Download a winapp2 to trim?
        If args.Contains("-d") Then
            download = True
            args.Remove("-d")
        End If

        If args.Contains("-ncc") Then
            ncc = True
            download = True
            args.Remove("-ncc")
        End If

        'Validate the contents of args
        validateArgs(args, "-wf -wd -sd -sf")

        'Parse the winapp2 file parameters if given
        If args.Contains("-wd") Then getFileNameAndDir(args, "-wd", winappDir, winappName)
        If args.Contains("-wf") Then getFileName(args, "-wf", winappName)

        'parse the save file parameters if given, if not given, match the winapp parameters
        If args.Contains("-sd") Then
            getFileNameAndDir(args, "-sd", saveDir, saveName)
        Else
            saveDir = winappDir
        End If

        If args.Contains("-sf") Then
            getFileName(args, "-sf", saveName)
        Else
            saveName = winappName
        End If

        'Initalize the trim
        remoteTrim(winappDir, winappName, saveDir, saveName, download, ncc)

    End Sub

    'Handle commandline args for Merge
    Private Sub autoMerge(ByRef args As List(Of String))
        Dim mergeMode As Boolean = True
        Dim winappDir As String = Environment.CurrentDirectory
        Dim winappName As String = "\winapp2.ini"
        Dim sDir As String = Environment.CurrentDirectory
        Dim sName As String = ""

        If args.Contains("-mm") Then
            mergeMode = False
            args.Remove("-mm")
        End If

        If args.Contains("-r") Then
            sName = "\Removed Entries.ini"
            args.Remove("-r")
        End If

        If args.Contains("-c") Then
            sName = "\Custom.ini"
            args.Remove("-c")
        End If

        'validate args
        validateArgs(args, "-fd -ff -sd -sf")

        'Collect any winapp2 info
        If args.Contains("-fd") Then getFileNameAndDir(args, "-fd", winappDir, winappName)
        If args.Contains("-ff") Then getFileName(args, "-ff", winappName)

        'Collect the merge file's info
        If args.Contains("-sd") Then getFileNameAndDir(args, "-sd", sDir, sName)
        If args.Contains("-sf") Then getFileName(args, "-sf", sName)

        'If we got a merge file, initiate the merge
        If Not sName = "" Then remoteMerge(winappDir, winappName, sDir, sName, mergeMode)
    End Sub

    'Handle commandline args for Diff
    Private Sub autoDiff(ByRef args As List(Of String))
        Dim ncc As Boolean = False
        Dim download As Boolean = False
        Dim firstDir As String = Environment.CurrentDirectory
        Dim firstName As String = "\winapp2.ini"
        Dim secondDir As String = Environment.CurrentDirectory
        Dim secondName As String = ""
        Dim save As Boolean = False

        'Download & Diff?
        If args.Contains("-d") Then
            secondName = "null"
            download = True
            args.Remove("-d")
        End If

        'Downloading non-ccleaner ini?
        If args.Contains("-ncc") Then
            ncc = True
            args.Remove("-ncc")
        End If

        'Save diff.txt?
        If args.Contains("-l") Then
            save = True
            args.Remove("-l")
        End If

        'validate args
        validateArgs(args, "-fd -ff -sd -sf")

        'Get the info for the first file
        If args.Contains("-fd") Then getFileNameAndDir(args, "-fd", firstDir, firstName)
        If args.Contains("-ff") Then getFileName(args, "-ff", firstName)

        'Get the info for the second file (if necessary)
        If args.Contains("-sd") Then getFileNameAndDir(args, "-sd", secondDir, secondName)
        If args.Contains("-sf") Then getFileName(args, "-sf", secondName)

        'Only run if we have a second file (because we assume we're running on a winapp2.ini file by default)
        If Not secondName = "" Then remoteDiff(firstDir, firstName, secondDir, secondName, download, ncc, save)
    End Sub

    'Handle commandline args for CCiniDebug
    Private Sub autoccini(ByRef args As List(Of String))
        Dim winappDir As String = Environment.CurrentDirectory
        Dim winappName As String = "\winapp2.ini"
        Dim ccDir As String = Environment.CurrentDirectory
        Dim ccName As String = "\ccleaner.ini"
        Dim prune As Boolean = False

        'Prune?
        If args.Contains("-p") Then
            prune = True
        End If

        'validate args
        validateArgs(args, "-wd -wf -cd -cf")

        'get the winapp2 info
        If args.Contains("-wd") Then getFileNameAndDir(args, "-wd", winappDir, winappName)
        If args.Contains("-wf") Then getFileName(args, "-wf", winappName)

        'get the ccleaner.ini info
        If args.Contains("-cd") Then getFileNameAndDir(args, "-cd", ccDir, ccName)
        If args.Contains("-cf") Then getFileName(args, "-wc", ccName)

        remoteCC(winappDir, winappName, ccDir, ccName, prune)
    End Sub

    Private Sub autodownload(ByRef args As List(Of String))
        Dim downloadDir As String = Environment.CurrentDirectory & "\winapp2ool downloads"
        Dim downloadName As String = ""
        Dim downloadFile As String

        If args.Contains("-p") Then
            downloadDir = Environment.CurrentDirectory
            args.Remove("-p")
        End If

        If args.Count > 0 Then
            Select Case args(0)
                Case "1"
                    downloadFile = wa2Link
                    downloadName = "\winapp2.ini"
                Case "2"
                    downloadFile = nonccLink
                    downloadName = "\winapp2.ini"
                Case "3"
                    downloadFile = toolLink
                    downloadName = "\winapp2ool.exe"
                Case "4"
                    downloadFile = removedLink
                    downloadName = "\Removed Entries.ini"
                Case Else
                    downloadFile = ""
            End Select
            args.RemoveAt(0)
        Else
            downloadFile = ""
        End If

        validateArgs(args, "-dd -df")

        If args.Contains("-dd") Then getFileNameAndDir(args, "-dd", downloadDir, downloadName)
        If args.Contains("-df") Then getFileName(args, "-df", downloadName)

        'If we're downloading winapp2ool, make sure we don't try to overwrite the currently running exe
        If downloadDir = Environment.CurrentDirectory And downloadFile = toolLink Then downloadDir += "\winapp2ool downloads"

        If Not downloadFile = "" Then remoteDownload(downloadDir, downloadName, downloadFile)

    End Sub

    Private Sub getFileName(ByRef args As List(Of String), flag As String, ByRef name As String)
        'Extract (what we assume to be) the file parameter from a file specification flag
        If args.Count >= 2 Then
            Dim ind As Integer = args.IndexOf(flag)
            name = "\" & args(ind + 1)
            args.RemoveAt(ind)
            args.RemoveAt(ind)
        End If
    End Sub

    Private Sub getFileNameAndDir(ByRef args As List(Of String), flag As String, ByRef path As String, ByRef name As String)
        'Take in a file with directory listing and split it into its constituent parts, saving them
        If args.Count >= 2 Then
            Dim ind As Integer = args.IndexOf(flag)
            getFileParams(args(ind + 1), path, name)
            args.RemoveAt(ind)
            args.RemoveAt(ind)
        End If
    End Sub

    Private Sub getFileParams(ByRef arg As String, ByRef path As String, ByRef name As String)
        'Start either a blank path or, support appending children folders to the current path
        path = If(arg.StartsWith("\"), Environment.CurrentDirectory & "\", "")

        'This function should receive a file path in full form from the command line and return it in pieces
        Dim splitArg As String() = arg.Split(CChar("\"))
        If splitArg.Count >= 2 Then
            For i As Integer = 0 To splitArg.Count - 2
                path += splitArg(i) & "\"
            Next
        End If
        name = splitArg.Last

    End Sub

    Public Sub processCommandLineArgs()

        'Build the arguments as a list of strings
        Dim args As New List(Of String)
        args.AddRange(Environment.GetCommandLineArgs)

        'Remove the first arg which is simply the name of the executable
        args.RemoveAt(0)

        'the s is for silent, if we have this flag, don't give any output to the user 
        If args.Contains("-s") Then
            suppressOutput = True
            args.Remove("-s")
        End If

        If args.Count > 0 Then

            Select Case args(0)
                Case "1"
                    args.RemoveAt(0)
                    autoDebug(args)
                Case "2"
                    args.RemoveAt(0)
                    autoTrim(args)
                Case "3"
                    args.RemoveAt(0)
                    autoMerge(args)
                Case "4"
                    args.RemoveAt(0)
                    autoDiff(args)
                Case "5"
                    args.RemoveAt(0)
                    autoccini(args)
                Case "6"
                    args.RemoveAt(0)
                    autodownload(args)
            End Select

        End If
    End Sub

    Private Sub validateArgs(args As List(Of String), validArgs As String)
        'This sub ensures that command line args are properly formatted in a ("-flag","data"...) format
        If args.Count > 0 Then
            Dim vArgs As New List(Of String)
            vArgs.AddRange(validArgs.Split(CChar(" ")))

            For i As Integer = 0 To args.Count - 1
                If Not vArgs.Contains(args(i)) Then
                    Console.WriteLine("Invalid command line arguements given. Press any key to exit.")
                    Console.ReadKey()
                    Environment.Exit(0)
                End If
                i += 1
            Next
        End If
    End Sub
End Module