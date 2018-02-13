Option Strict On
Imports System.IO
Imports System.Net

Module Downloader

    Public wa2Link As String = "https://raw.githubusercontent.com/MoscaDotTo/Winapp2/master/Winapp2.ini"
    Public nonccLink As String = "https://raw.githubusercontent.com/MoscaDotTo/Winapp2/master/Non-CCleaner/Winapp2.ini"
    Public toolLink As String = "https://github.com/MoscaDotTo/Winapp2/raw/master/Tools/beta/winapp2ool.exe"
    Public toolVerLink As String = "https://raw.githubusercontent.com/MoscaDotTo/Winapp2/master/Tools/beta/version.txt"
    Dim downloadDir As String = Environment.CurrentDirectory & "\winapp2ool downloads"

    Dim exitCode As Boolean = False
    Dim hasMenuTopper As Boolean = False

    Private Sub printMenu()
        If Not hasMenuTopper Then
            hasMenuTopper = True
            tmenu("Download")
        End If
        menu(menuStr03, "")
        menu("Menu: Enter a number to select", "c")
        menu(menuStr01, "")
        menu("0. Exit                         - Return to the winapp2ool menu", "l")
        menu("1. winapp2.ini                  - Download the latest winapp2.ini", "l")
        menu("2. Non-CCleaner                 - Download the latest non-ccleaner winapp2.ini", "l")
        menu("3. winapp2ool                   - Download the latest winapp2ool.exe", "l")
        menu("4. directory                    - Change the download directory", "l")
        menu(menuStr02, "")

    End Sub
    Public Sub main()
        exitCode = False
        hasMenuTopper = False
        Console.Clear()

        Do Until exitCode
            printMenu()
            Console.WriteLine()
            Console.Write("Enter a number: ")
            Dim input As String = Console.ReadLine
            Select Case input
                Case "0"
                    Console.WriteLine("Returning to winapp2ool menu...")
                    exitCode = True
                Case "1"
                    download("winapp2.ini", wa2Link, downloadDir)
                    Console.Clear()
                    tmenu("Download complete:")
                    menu("winapp2.ini", "c")
                Case "2"
                    download("winapp2.ini", nonccLink, downloadDir)
                    Console.Clear()
                    tmenu("Download complete:")
                    menu("winapp2.ini", "c")
                Case "3"
                    download("winapp2ool.exe", toolLink, downloadDir)
                    Console.Clear()
                    tmenu("Download complete:")
                    menu("winapp2ool.exe", "c")
                Case "4"
                    dChooser(downloadDir, exitCode)
                    Console.Clear()
                    tmenu("Current download directory:")
                    menu(downloadDir, "l")
                Case Else
                    Console.Clear()
                    tmenu("Invalid input. Please try again.")
            End Select
        Loop
    End Sub

    Public Function getFileDataAtLineNum(address As String, lineNum As Integer) As String
        Dim reader As StreamReader = Nothing
        Try
            reader = New StreamReader(address)
            Return getTargetLine(reader, lineNum)
        Catch ex As Exception
            Console.WriteLine("Error: " & ex.ToString)
            Console.WriteLine("Please report this error on GitHub")
            Console.WriteLine()
            Return Nothing
        Finally
            reader.Close()
        End Try
    End Function

    Private Function getTargetLine(reader As StreamReader, lineNum As Integer) As String
        Dim out As String = ""
        Dim curLine As Integer = 0

        While curLine < lineNum
            out = reader.ReadLine()
            curLine += 1
        End While
        Return out

    End Function

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
                Console.WriteLine("Error: " & ex.ToString)
                Console.WriteLine("Please report this error on GitHub")
                Console.WriteLine()
            End If
            Return ""
        End Try
    End Function

    Private Sub download(fileName As String, fileLink As String, downloadDir As String)

        Dim givenName As String = fileName

        If Not Directory.Exists(downloadDir) Then
            Directory.CreateDirectory(downloadDir)
        End If

        If File.Exists(downloadDir & "\" & fileName) Then
            Console.WriteLine(fileName & " already exists in the target directory.")
            Console.Write("Enter a new file name, or leave blank to overwrite the existing file: ")
            Dim nfilename As String = Console.ReadLine()
            If nfilename.Trim <> "" Then
                fileName = nfilename
            End If
        End If

        Console.WriteLine("Downloading " & givenName & "...")

        Try
            Dim dl As New WebClient
            dl.DownloadFileAsync(New Uri(fileLink), (downloadDir & "\" & fileName))
        Catch ex As Exception
            Console.WriteLine("Error: " & ex.ToString)
            Console.WriteLine("Please report this error on GitHub")
            Console.WriteLine()
        End Try
        Console.WriteLine("Download complete.")
        Console.WriteLine("Downloaded " & fileName & " to " & downloadDir)
    End Sub
End Module