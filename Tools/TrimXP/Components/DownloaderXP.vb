Imports System.IO
Imports System.Net

Module DownloaderXP
    'Links
    Public wa2Link As String = "https://raw.githubusercontent.com/MoscaDotTo/Winapp2/master/Winapp2.ini"
    Public nonccLink As String = "https://raw.githubusercontent.com/MoscaDotTo/Winapp2/master/Non-CCleaner/Winapp2.ini"

    'File handler
    Dim downloadFile As iniFile = New iniFile(Environment.CurrentDirectory, "")

    'fetch winapp2.ini from github (ncc or otherwise)
    Public Function getRemoteWinapp(ncc As Boolean) As iniFile
        Return If(ncc, getRemoteIniFile(nonccLink, "\winapp2.ini"), getRemoteIniFile(wa2Link, "\winapp2.ini"))
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

    Private Sub download(link As String, prompt As Boolean)

        'Remember the initial given name
        Dim givenName As String = downloadFile.name

        ' Don't try to download to a directory that doesn't exist
        If Not Directory.Exists(downloadFile.dir) Then Directory.CreateDirectory(downloadFile.dir)

        If prompt Then
            If File.Exists(downloadFile.path) Then
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
