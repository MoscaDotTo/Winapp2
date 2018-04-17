Option Strict On
Imports System.IO

Module commandLineHandler

    'File Handlers
    Dim firstFile As IFileHandlr
    Dim secondFile As IFileHandlr
    Dim thirdFile As IFileHandlr


    'Handle commandline args for WinappDebug
    Private Sub autoDebug(ByRef args As List(Of String))

        Dim correctErrors As Boolean

        'Get default params
        initDebugParams(firstFile, secondFile, correctErrors)

        'Toggle on autocorrect (off by default)
        If args.Contains("-c") Then
            correctErrors = True
            args.Remove("-c")
        End If

        'Get any file parameters from flags
        validateAndParse(args)

        'Initialize the debug
        remoteDebug(firstFile, secondFile, correctErrors)
    End Sub

    'Handle commandline args for Trim
    Private Sub autoTrim(ByRef args As List(Of String))

        Dim download As Boolean
        Dim ncc As Boolean

        initTrimParams(firstFile, secondFile, download, ncc)

        'Download a winapp2 to trim?
        If args.Contains("-download") Then
            download = True
            args.Remove("-download")
        End If

        'Download the non ccleaner ini? Implies -d
        If args.Contains("-ncc") Then
            ncc = True
            download = True
            args.Remove("-ncc")
        End If

        'Get any file parameters from flags
        validateAndParse(args)

        'Initalize the trim
        remoteTrim(firstFile, secondFile, download, ncc)

    End Sub

    'Handle commandline args for Merge
    Private Sub autoMerge(ByRef args As List(Of String))

        Dim mergeMode As Boolean

        initMergeParams(firstFile, secondFile, thirdFile, mergeMode)

        If args.Contains("-mm") Then
            mergeMode = False
            args.Remove("-mm")
        End If

        If args.Contains("-r") Then
            secondFile.name = "Removed Entries.ini"
            args.Remove("-r")
        End If

        If args.Contains("-c") Then
            secondFile.name = "Custom.ini"
            args.Remove("-c")
        End If

        'Get any file parameters from flags
        validateAndParse(args)

        'If we have a secondfile, initiate the merge
        If Not secondFile.name = "" Then remoteMerge(firstFile, secondFile, thirdFile, mergeMode)
    End Sub

    'Handle commandline args for Diff
    Private Sub autoDiff(ByRef args As List(Of String))
        Dim ncc As Boolean
        Dim download As Boolean
        Dim save As Boolean = False

        initDiffParams(firstFile, secondFile, thirdFile, download, ncc, save)

        'Download & Diff?
        If args.Contains("-download") Then
            secondFile.name = "Online winapp2.ini"
            download = True
            args.Remove("-download")
        End If

        'Downloading non-ccleaner ini?
        If args.Contains("-ncc") Then
            ncc = True
            download = True
            secondFile.name = "Online non-ccleaner winapp2.ini"
            args.Remove("-ncc")
        End If

        'Save diff.txt?
        If args.Contains("-l") Then
            save = True
            args.Remove("-l")
        End If

        'Only run if we have a second file (because we assume we're running on a winapp2.ini file by default)
        If Not secondFile.name = "" Then remoteDiff(firstFile, secondFile, thirdFile, download, ncc, save)
    End Sub

    'Handle commandline args for CCiniDebug
    Private Sub autoccini(ByRef args As List(Of String))
        Dim prune As Boolean
        Dim save As Boolean
        Dim sort As Boolean

        initCCDebugParams(firstFile, secondFile, thirdFile, prune, save, sort)

        'Prune?
        If args.Contains("-noprune") Then
            prune = False
            args.Remove("-noprune")
        End If

        If args.Contains("-nosave") Then
            save = False
            args.Remove("-nosave")
        End If

        If args.Contains("-nosort") Then
            sort = False
            args.Remove("-nosort")
        End If

        'Get any file parameters from flags
        validateAndParse(args)

        'run ccinidebug
        remoteCC(firstFile, secondFile, thirdFile, prune, save, sort)
    End Sub

    Private Sub autodownload(ByRef args As List(Of String))

        Dim downloadFile As String
        Dim downloadDir As String = Environment.CurrentDirectory & "\winapp2ool downloads"
        Dim downloadName As String

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
                    downloadName = ""
            End Select
            args.RemoveAt(0)
        Else
            downloadFile = ""
            downloadName = ""
        End If

        'Get any file parameters from flags
        validateAndParse(args)

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

    Private Sub getFileNameAndDir(ByRef args As List(Of String), flag As String, ByRef file As IFileHandlr)
        'Take in a file with directory listing and split it into its constituent parts, saving them
        If args.Count >= 2 Then
            Dim ind As Integer = args.IndexOf(flag)
            getFileParams(args(ind + 1), file)
            args.RemoveAt(ind)
            args.RemoveAt(ind)
        End If
    End Sub

    Private Sub getFileParams(ByRef arg As String, ByRef file As IFileHandlr)
        'Start either a blank path or, support appending children folders to the current path
        file.dir = If(arg.StartsWith("\"), Environment.CurrentDirectory & "\", "")

        'This function should receive a file path in full form from the command line and return it in pieces
        Dim splitArg As String() = arg.Split(CChar("\"))
        If splitArg.Count >= 2 Then
            For i As Integer = 0 To splitArg.Count - 2
                file.dir += splitArg(i) & "\"
            Next
        End If
        file.name = splitArg.Last

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

    Private Sub getFileAndDirParams(ByRef args As List(Of String))
        'Get the info for the first file
        If args.Contains("-1d") Then getFileNameAndDir(args, "-1d", firstFile)
        If args.Contains("-1f") Then getFileName(args, "-1f", firstFile.name)

        'Get the info for the second file (if necessary)
        If args.Contains("-2d") Then getFileNameAndDir(args, "-2d", secondFile)
        If args.Contains("-2f") Then getFileName(args, "-2f", secondFile.name)

        If args.Contains("-3d") Then getFileNameAndDir(args, "-3d", thirdFile)
        If args.Contains("-3f") Then getFileName(args, "-3f", thirdFile.name)

    End Sub

    Private Sub validateAndParse(args As List(Of String))
        'validate args
        validateArgs(args)

        'Get file params
        getFileAndDirParams(args)
    End Sub

    Private Sub validateArgs(args As List(Of String))
        'This sub ensures that command line args are properly formatted in a ("-flag","data"...) format

        Dim vArgs As New List(Of String) From {"-1d", "-1f", "-2d", "-2f", "-3d", "-3f"}
        If args.Count > 0 Then
            Try
                For i As Integer = 0 To args.Count - 1
                    If Not vArgs.Contains(args(i)) Then
                        Console.WriteLine("Invalid command line arguements given. Press any key to exit.")
                        Console.ReadKey()
                        Environment.Exit(0)
                    Else
                        If Not Directory.Exists(args(i + 1)) And Not Directory.Exists(Environment.CurrentDirectory & "\" & args(i + 1)) And
                                Not File.Exists(args(i + 1)) And File.Exists(Environment.CurrentDirectory & "\" & args(i + 1)) Then
                            Console.WriteLine("Invalid command line arguements given. Press any key to exit.")
                            Console.ReadKey()
                            Environment.Exit(0)
                        End If
                    End If
                    i += 1
                Next
            Catch ex As Exception
                Console.WriteLine("Invalid command line arguements given. Press any key to exit.")
                Console.ReadKey()
                Environment.Exit(0)
            End Try
        End If
    End Sub
End Module