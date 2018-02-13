Option Strict On
Imports System.IO
Module Winapp2ool

    Dim currentVersion As Double = 0.6
    Dim latestVersion As String

    Dim checkedForUpdates As Boolean = False
    Dim updateIsAvail As Boolean = False

    Dim latestWa2Ver As String = ""
    Dim localWa2Ver As String = ""

    Dim waUpdateIsAvail As Boolean = False
    Dim menuHasTopper As Boolean = False

    Private Sub printMenu()
        If menuHasTopper Then
            mMenu("Winapp2ool - A multitool for winapp2.ini and related files")
        Else
            tmenu("Winapp2ool - A multitool for winapp2.ini and related files")
            menu(menuStr03)
            menuHasTopper = True
        End If
        menu("Menu: Enter a number to select", "c")
        If updateIsAvail Then
            menu(menuStr01, "")
            menu("A new version of Winapp2ool is available!", "c")
            menu("Current:   v" & currentVersion, "c")
            menu("Available: v" & latestVersion, "c")
        End If
        If waUpdateIsAvail And Not localWa2Ver = "0" Then
            menu(menuStr01, "")
            menu("A new version of winapp2.ini is available!", "c")
            menu("Current:   v" & localWa2Ver, "c")
            menu("Available: v" & latestWa2Ver, "c")
        End If
        menu(menuStr01, "")
        menu("0. Exit             - Exit the application", "l")
        menu("1. WinappDebug      - Load the WinappDebug tool to check for errors in winapp2.ini", "l")
        menu("2. CCiniDebug       - Load the CCiniDebug tool to sort and trim ccleaner.ini", "l")
        menu("3. Diff             - Load the Diff tool to observe the changes between two winapp2.ini files", "l")
        menu("4. Trim             - Load the Trim tool to debloat winapp2.ini for your system", "l")
        menu("5. Downloader       - Load the Downloader tool to download files from the Winapp2 GitHub", "l")
        menu(menuStr02, "")
    End Sub

    Public Sub main()
        Console.Title = "Winapp2ool v" & currentVersion & " beta"
        Console.WindowWidth = 120

        checkUpdates()
        Dim exitCode As Boolean = False
        Do Until exitCode = True
            printMenu()
            Console.WriteLine()
            Console.Write("Enter a number: ")
            Dim input As String = Console.ReadLine

            Select Case input
                Case "0"
                    exitCode = True
                    Console.WriteLine("Exiting...")
                    Environment.Exit(1)
                Case "1"
                    WinappDebug.main()
                    Console.Clear()
                    tmenu("Finished running WinappDebug")
                Case "2"
                    CCiniDebug.Main()
                    Console.Clear()
                    tmenu("Finished running CCiniDebug")
                Case "3"
                    Diff.main()
                    Console.Clear()
                    tmenu("Finished running Diff")
                Case "4"
                    Trim.main()
                    Console.Clear()
                    tmenu("Finished running Trim")
                Case "5"
                    Downloader.main()
                    Console.Clear()
                    tmenu("Finished running Downloader")
                Case Else
                    Console.Clear()
                    tmenu("Invalid input. Please try again.")
            End Select
        Loop
    End Sub

    Private Sub checkUpdates()
        Try
            latestVersion = getRemoteFileDataAtLineNum(toolVerLink, 1)

            If CDbl(latestVersion) > currentVersion Then
                updateIsAvail = True
            End If
            latestWa2Ver = getRemoteFileDataAtLineNum(wa2Link, 1).Split(CChar(" "))(2)
            If Not File.Exists(Environment.CurrentDirectory & "\winapp2.ini") Then
                localWa2Ver = "0"
                tmenu("/!\ Update check failed. /!\")
                menuHasTopper = True
            Else
                localWa2Ver = getFileDataAtLineNum(Environment.CurrentDirectory & "\winapp2.ini", 1).Split(CChar(" "))(2)
            End If
            If CDbl(latestWa2Ver) > CDbl(localWa2Ver) Then
                waUpdateIsAvail = True
            End If

        Catch ex As Exception
            tmenu("/!\ Update check failed. /!\")
            menuHasTopper = True
        End Try
    End Sub

End Module