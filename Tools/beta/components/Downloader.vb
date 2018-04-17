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
    Dim downloadFile As IFileHandlr = New IFileHandlr(Environment.CurrentDirectory & "\winapp2ool downloads", "")

    'Menu settings
    Dim exitCode As Boolean = False
    Dim menuTopper As String = ""

    Private Sub printMenu()
        printMenuLine(tmenu(menuTopper))
        printMenuLine(menuStr03)
        printMenuLine("Menu: Enter a number to select", "c")
        printMenuLine(menuStr01)
        printMenuLine("0. Exit                         - Return to the winapp2ool menu", "l")
        printMenuLine("1. Winapp2.ini                  - Download the latest winapp2.ini", "l")
        printMenuLine("2. Non-CCleaner                 - Download the latest non-ccleaner winapp2.ini", "l")
        printMenuLine("3. Winapp2ool                   - Download the latest winapp2ool.exe", "l")
        printMenuLine("4. Removed Entries.ini          - Download only entries used to create the non-ccleaner winapp2.ini", "l")
        printMenuLine(menuStr01)
        printMenuLine("5. Directory                    - Change the save directory", "l")
        printMenuLine(menuStr01)
        printMenuLine("Save directory: " & replDir(downloadFile.dir), "l")
        printMenuLine(menuStr02)
    End Sub

    Public Sub main()
        exitCode = False
        menuTopper = "Download"

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
                Case "1"
                    download("winapp2.ini", wa2Link, downloadFile.dir)
                Case "2"
                    download("winapp2.ini", nonccLink, downloadFile.dir)
                Case "3"
                    download("winapp2ool.exe", toolLink, downloadFile.dir)
                Case "4"
                    download("Removed entries.ini", removedLink, downloadFile.dir)
                Case "5"
                    dChooser(downloadFile.dir, exitCode)
                    menuTopper = "Changed save directory"
                Case Else
                    menuTopper = invInpStr
            End Select
        Loop
    End Sub

    'fetch winapp2.ini from github (ncc or otherwise)
    Public Function getRemoteWinapp(ncc As Boolean) As iniFile
        Return If(ncc, getRemoteIniFile(nonccLink, "\winapp2.ini"), getRemoteIniFile(wa2Link, "\winapp2.ini"))
    End Function

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

    Public Sub remoteDownload(dir As String, name As String, link As String)
        download(name, link, dir)
    End Sub

    Private Sub download(fileName As String, fileLink As String, downloadDir As String)

        Dim givenName As String = fileName

        ' Don't try to download to a place that doesn't exist
        If Not Directory.Exists(downloadDir) Then Directory.CreateDirectory(downloadDir)

        'Prompt to overwrite files if they exist already
        If File.Exists(downloadDir & "\" & fileName) And Not suppressOutput Then
            Console.WriteLine(fileName & " already exists in the target directory.")
            Console.Write("Enter a new file name, or leave blank to overwrite the existing file: ")
            Dim nfilename As String = Console.ReadLine()
            If nfilename.Trim <> "" Then fileName = nfilename
        End If

        cwl("Downloading " & givenName & "...")

        'Preform the actual download
        Try
            Dim dl As New WebClient
            dl.DownloadFile(New Uri(fileLink), downloadDir & "\" & fileName)
        Catch ex As Exception
            exc(ex)
            Console.ReadKey()
        End Try
        cwl("Download complete.")
        cwl("Downloaded " & fileName & " to " & downloadDir)
        menuTopper = "Download complete: " & fileName

    End Sub
End Module