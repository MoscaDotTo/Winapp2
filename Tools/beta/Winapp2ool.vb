Option Strict On
Imports System.IO
Module Winapp2ool

    Dim currentVersion As Double = 0.85
    Dim latestVersion As String

    Dim checkedForUpdates As Boolean = False
    Dim updateIsAvail As Boolean = False

    Dim latestWa2Ver As String = ""
    Dim localWa2Ver As String = ""

    'This boolean will prevent us from printing or asking for input under most circumstances, triggered by the -s command line argument 
    Public suppressOutput As Boolean = False

    Dim waUpdateIsAvail As Boolean = False
    Dim menuTopper As String = ""

    Private Sub printMenu()
        printMenuLine(tmenu(menuTopper))
        printMenuLine(menuStr03)
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
        menuTopper = "Winapp2ool - A multitool for winapp2.ini and related files"
        Console.Title = "Winapp2ool v" & currentVersion & " beta"
        Console.WindowWidth = 120
        processCommandLineArgs()
        If suppressOutput Then Environment.Exit(1)
        checkUpdates()
        Dim exitCode As Boolean = False
        Do Until exitCode = True
            Console.Clear()
            printMenu()
            cwl()
            Console.Write("Enter a number: ")
            Dim input As String = Console.ReadLine

            Select Case input
                Case "0"
                    exitCode = True
                    cwl("Exiting...")
                    Environment.Exit(1)
                Case "1"
                    WinappDebug.main()
                    menuTopper = "Finished running WinappDebug"
                Case "2"
                    Trim.main()
                    menuTopper = "Finished running Trim"
                Case "3"
                    Merge.main()
                    menuTopper = "Finished running Merge"
                Case "4"
                    Diff.main()
                    menuTopper = "Finished running Diff"
                Case "5"
                    CCiniDebug.Main()
                    menuTopper = "Finished running CCiniDebug"
                Case "6"
                    Downloader.main()
                    menuTopper = "Finished running Downloader"
                Case Else
                    menuTopper = invInpStr
            End Select
        Loop
    End Sub

    Private Sub checkUpdates()
        Try
            'Check for winapp2ool.exe updates
            latestVersion = getRemoteFileDataAtLineNum(toolVerLink, 1)
            updateIsAvail = CDbl(latestVersion) > currentVersion

            'Check for winapp2.ini updates
            latestWa2Ver = getRemoteFileDataAtLineNum(wa2Link, 1).Split(CChar(" "))(2)
            If Not File.Exists(Environment.CurrentDirectory & "\winapp2.ini") Then
                localWa2Ver = "0"
                menuTopper = "/!\ Winapp2.ini update check failed. /!\"
            Else
                localWa2Ver = getFileDataAtLineNum(Environment.CurrentDirectory & "\winapp2.ini", 1).Split(CChar(" "))(2)
            End If
            waUpdateIsAvail = CDbl(latestWa2Ver) > CDbl(localWa2Ver)

        Catch ex As Exception
            menuTopper = "/!\ Update check failed. /!\"
        End Try
    End Sub

    Public Sub exc(ByRef ex As Exception)
        cwl("Error: " & ex.ToString)
        cwl("Please report this error on GitHub")
        cwl()
    End Sub

    Public Sub cwl()
        If Not suppressOutput Then Console.WriteLine()
    End Sub

    Public Sub cwl(msg As String)
        If Not suppressOutput Then Console.WriteLine(msg)
    End Sub

    'Prompt the user to change a file's parameters, flag it as changed, and mark settings as having changed
    Public Sub changeFileParams(ByRef someFile As IFileHandlr, menuTopper As String, ByRef settingsChangedSetting As Boolean, ByRef ec As Boolean)
        fChooser(someFile.dir, someFile.name, ec, someFile.initName, someFile.secondName)
        settingsChangedSetting = True
        menuTopper = If(someFile.secondName <> "", someFile.initName, "save file") & " parameters updated"
    End Sub

    'Toggle a parameter on or off and mark settings as having changed
    Public Sub toggleSettingParam(ByRef setting As Boolean, paramText As String, ByRef topper As String, ByRef settingsChangedSetting As Boolean)
        setting = Not setting
        topper = paramText & If(setting, "enabled", "disabled")
        settingsChangedSetting = True
    End Sub

    Public Function getMenuNumber(valList As List(Of Boolean), lowestStartingNum As Integer) As Integer
        For Each setting In valList
            If setting Then lowestStartingNum += 1
        Next
        Return lowestStartingNum
    End Function

End Module