Option Strict On
Imports System.IO
Module Winapp2ool

    Dim currentVersion As Double = 0.7
    Dim latestVersion As String

    Dim checkedForUpdates As Boolean = False
    Dim updateIsAvail As Boolean = False

    Dim latestWa2Ver As String = ""
    Dim localWa2Ver As String = ""

    Dim waUpdateIsAvail As Boolean = False
    Dim menuHasTopper As Boolean = False

    Private Sub printMenu()
        If menuHasTopper Then
            printMenuLine(mMenu("Winapp2ool - A multitool for winapp2.ini and related files"))
        Else
            printMenuLine(tmenu("Winapp2ool - A multitool for winapp2.ini and related files"))
            printMenuLine(menuStr03)
            menuHasTopper = True
        End If
        printMenuLine("Menu: Enter a number to select", "c")
        If updateIsAvail Then
            printMenuLine(menuStr01)
            printMenuLine("A new version of Winapp2ool is available!", "c")
            printMenuLine("Current:   v" & currentVersion, "c")
            printMenuLine("Available: v" & latestVersion, "c")
        End If
        If waUpdateIsAvail And Not localWa2Ver = "0" Then
            printMenuLine(menuStr01)
            printMenuLine("A new version of winapp2.ini is available!", "c")
            printMenuLine("Current:   v" & localWa2Ver, "c")
            printMenuLine("Available: v" & latestWa2Ver, "c")
        End If
        printMenuLine(menuStr01)
        printMenuLine("0. Exit             - Exit the application", "l")
        printMenuLine("1. WinappDebug      - Check for and correct errors in winapp2.ini", "l")
        printMenuLine("2. Trim             - Debloat winapp2.ini for your system", "l")
        printMenuLine("3. Merge            - Merge the contents of an ini file into winapp2.ini", "l")
        printMenuLine("4. Diff             - Observe the changes between two winapp2.ini files", "l")
        printMenuLine("5. CCiniDebug       - Sort and trim ccleaner.ini", "l")
        printMenuLine(menuStr01)
        printMenuLine("6. Downloader       - Download files from the Winapp2 GitHub", "l")
        printMenuLine(menuStr02)
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
                    printMenuLine(tmenu("Finished running WinappDebug"))
                Case "2"
                    Trim.main()
                    Console.Clear()
                    printMenuLine(tmenu("Finished running Trim"))
                Case "3"
                    Merge.main()
                    Console.Clear()
                    printMenuLine(tmenu("Finished running Merge"))
                Case "4"
                    Diff.main()
                    Console.Clear()
                    printMenuLine(tmenu("Finished running Diff"))
                Case "5"
                    CCiniDebug.Main()
                    Console.Clear()
                    printMenuLine(tmenu("Finished running CCiniDebug"))
                Case "6"
                    Downloader.main()
                    Console.Clear()
                    printMenuLine(tmenu("Finished running Downloader"))
                Case Else
                    Console.Clear()
                    printMenuLine(tmenu("Invalid input. Please try again."))
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
                printMenuLine(tmenu("/!\ Update check failed. /!\"))
                menuHasTopper = True
            Else
                localWa2Ver = getFileDataAtLineNum(Environment.CurrentDirectory & "\winapp2.ini", 1).Split(CChar(" "))(2)
            End If
            If CDbl(latestWa2Ver) > CDbl(localWa2Ver) Then
                waUpdateIsAvail = True
            End If

        Catch ex As Exception
            printMenuLine(tmenu("/!\ Update check failed. /!\"))
            menuHasTopper = True
        End Try
    End Sub

    Public Sub exc(ByRef ex As Exception)
        Console.WriteLine("Error: " & ex.ToString)
        Console.WriteLine("Please report this error on GitHub")
        Console.WriteLine()
    End Sub

End Module