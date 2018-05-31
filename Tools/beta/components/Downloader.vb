Option Strict On
Imports System.IO
Imports System.Net
Module Downloader

    'Links to GitHub resources
    Public wa2Link As String = "https://raw.githubusercontent.com/MoscaDotTo/Winapp2/master/Winapp2.ini"
    Public nonccLink As String = "https://raw.githubusercontent.com/MoscaDotTo/Winapp2/master/Non-CCleaner/Winapp2.ini"
    Public toolLink As String = "https://github.com/MoscaDotTo/Winapp2/raw/master/Tools/beta/winapp2ool.exe"
    Public toolVerLink As String = "https://raw.githubusercontent.com/MoscaDotTo/Winapp2/master/Tools/beta/version.txt"
    Public removedLink As String = "https://raw.githubusercontent.com/MoscaDotTo/Winapp2/master/Non-CCleaner/Removed%20entries.ini"

    'File handler
    Dim downloadFile As iniFile = New iniFile(Environment.CurrentDirectory, "")

    Private Sub printMenu()
        printMenuTop({"Download files from the winapp2 GitHub"}, True)
        printMenuOpt("Winapp2.ini", "Download the latest winapp2.ini")
        printMenuOpt("Non-CCleaner", "Download the latest non-ccleaner winapp2.ini")
        printMenuOpt("Winapp2ool", "Download the latest winapp2ool.exe")
        printMenuOpt("Removed Entries.ini", "Download only entries used to create the non-ccleaner winapp2.ini")
        printBlankMenuLine()
        printMenuOpt("Directory", "Change the save directory")
        printBlankMenuLine()
        printMenuLine("Save directory: " & replDir(downloadFile.dir), "l")
        printMenuLine(menuStr02)
    End Sub

    Public Sub main()
        initMenu("Download", 35)

        Do Until exitCode
            Console.Clear()
            printMenu()
            Console.WriteLine()
            Console.Write("Enter a number: ")
            Dim input As String = Console.ReadLine
            Select Case input
                Case "0"
                    Console.WriteLine("Returning to winapp2ool menu...")
                    exitCode = True
                Case "1", "2"
                    downloadFile.name = "winapp2.ini"
                    Dim link As String = If(input = "1", wa2Link, nonccLink)
                    download(wa2Link, True)
                Case "3"
                    If downloadFile.dir = Environment.CurrentDirectory Then
                        autoUpdate()
                    Else
                        downloadFile.name = "winapp2ool.exe"
                        download(toolLink, True)
                    End If
                Case "4"
                    downloadFile.name = "Removed entries.ini"
                    download(removedLink, True)
                Case "5"
                    dirChooser(downloadFile.dir)
                    undoAnyPendingExits()
                    menuTopper = "Save directory changed"
                Case Else
                    menuTopper = invInpStr
            End Select
        Loop
        revertMenu()
    End Sub

    'fetch winapp2.ini from github (ncc or otherwise)
    Public Function getRemoteWinapp(ncc As Boolean) As iniFile
        Return If(ncc, getRemoteIniFile(nonccLink, "\winapp2.ini"), getRemoteIniFile(wa2Link, "\winapp2.ini"))
    End Function

    'Read a file until a specified line and return that line 
    Public Function getFileDataAtLineNum(address As String, lineNum As Integer) As String
        Dim reader As StreamReader = Nothing
        Try
            reader = New StreamReader(address)
            Return getTargetLine(reader, lineNum)
        Catch ex As Exception
            exc(ex)
            Return Nothing
        Finally
            reader.Close()
        End Try
    End Function

    'Load a remote ini file
    Public Function getRemoteIniFile(address As String, name As String) As iniFile
        Dim reader As StreamReader = Nothing
        Try
            Dim client As New WebClient
            reader = New StreamReader(client.OpenRead(address))
            Dim wholeFile As String = reader.ReadToEnd
            wholeFile += Environment.NewLine
            Dim splitFile As String() = wholeFile.Split(CType(Environment.NewLine, Char()))
            Return New iniFile(splitFile, name)
        Catch ex As Exception
            exc(ex)
            Return Nothing
        End Try
    End Function

    'Read a file only until a specific line and then return that line
    Private Function getTargetLine(reader As StreamReader, lineNum As Integer) As String
        Dim out As String = ""
        Dim curLine As Integer = 1

        While curLine <= lineNum
            out = reader.ReadLine()
            curLine += 1
        End While
        Return out

    End Function

    'Load a remote file and toss it into getTargetLine, return the targeted line as a string
    Public Function getRemoteFileDataAtLineNum(address As String, lineNum As Integer) As String
        Dim reader As StreamReader = Nothing
        Try
            Dim client As New WebClient
            reader = New StreamReader(client.OpenRead(address))
            Return getTargetLine(reader, lineNum)
            reader.Close()

        Catch ex As Exception

            If ex.Message.StartsWith("The remote name could not be resolved") Then
                Console.WriteLine("Error: Could not establish connection to " & address)
                Console.WriteLine("Please check your internet connection settings and try again. If you feel this is a bug, please report it on GitHub.")
                Console.WriteLine()
            Else
                exc(ex)
            End If
            Return ""
        End Try
    End Function

    'Handle a call to download a file from an external module
    Public Sub remoteDownload(dir As String, name As String, link As String, prompt As Boolean)
        downloadFile.dir = dir
        downloadFile.name = name
        download(link, prompt)
    End Sub

    'Download the latest version of winapp2ool.exe and replace the currently running version with it, then launch that new version.
    Public Sub autoUpdate()
        downloadFile.dir = Environment.CurrentDirectory
        downloadFile.name = "winapp2ool updated.exe"
        Try
            'Remove any existing backups of this version
            If File.Exists(Environment.CurrentDirectory & "\winapp2ool v" & currentVersion & ".exe.bak") Then File.Delete(Environment.CurrentDirectory & "\winapp2ool v" & currentVersion & ".exe.bak")

            'Remove any old update files that didn't get renamed for whatever reason
            If File.Exists(downloadFile.path) Then File.Delete(downloadFile.path)
            download(toolLink, False)

            'Rename the old executable
            File.Move("winapp2ool.exe", "winapp2ool v" & currentVersion & ".exe.bak")

            'Rename the new executable
            File.Move("winapp2ool updated.exe", "winapp2ool.exe")

            'Start the new executable and exit the current one
            System.Diagnostics.Process.Start(Environment.CurrentDirectory & "\winapp2ool.exe")
            Environment.Exit(0)

        Catch ex As Exception
            exc(ex)
        End Try
    End Sub

    Private Sub download(link As String, prompt As Boolean)

        'Remember the initial given name
        Dim givenName As String = downloadFile.name

        ' Don't try to download to a directory that doesn't exist
        If Not Directory.Exists(downloadFile.dir) Then Directory.CreateDirectory(downloadFile.dir)

        If prompt Then
            If File.Exists(downloadFile.path) And Not suppressOutput Then
                Console.WriteLine(downloadFile.name & " already exists in the target directory.")
                Console.Write("Enter a new file name, or leave blank to overwrite the existing file: ")
                Dim nfilename As String = Console.ReadLine()
                If nfilename.Trim <> "" Then downloadFile.name = nfilename
            End If
        End If

        cwl("Downloading " & givenName & "...")

        'Preform the actual download
        Try
            Dim dl As New WebClient
            dl.DownloadFile(New Uri(link), downloadFile.path)
        Catch ex As Exception
            exc(ex)
            Console.ReadKey()
        End Try

        cwl("Download complete.")
        cwl("Downloaded " & downloadFile.name & " to " & downloadFile.dir)
        menuTopper = "Download complete: " & downloadFile.name
    End Sub
End Module