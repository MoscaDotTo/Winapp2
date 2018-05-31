Option Strict On
Imports System.IO
Module Winapp2ool

    'Update and compatibility booleans
    Public currentVersion As Double = 0.85
    Dim latestVersion As String = ""
    Dim latestWa2Ver As String = ""
    Dim localWa2Ver As String = ""
    Dim checkedForUpdates As Boolean = False
    Dim updateIsAvail As Boolean = False
    Dim dnfOOD As Boolean = False
    Dim waUpdateIsAvail As Boolean = False

    'This boolean will prevent us from printing or asking for input under most circumstances, triggered by the -s command line argument 
    Public suppressOutput As Boolean = False

    'Dummy iniFile obj for when we don't need actual data or when it'll be filled in later
    Public eini As iniFile = New iniFile("", "")
    Public lwinapp2File As iniFile = New iniFile(Environment.CurrentDirectory, "winapp2.ini")

    'Inform the user when an update is available 
    Private Sub printUpdNotif(cond As Boolean, updName As String, oldVer As String, newVer As String)
        If cond Then
            printMenuLine("A new version of " & updName & " is available!", "c")
            printMenuLine("Current:   v" & oldVer, "c")
            printMenuLine("Available: v" & newVer, "c")
            printBlankMenuLine()
        End If
    End Sub

    Private Sub printMenu()
        printMenuTop({}, False)

        printIf(dnfOOD, "line", "Your .NET Framework is out of date. Please update to version 4.6+.", "c")
        printUpdNotif(waUpdateIsAvail And Not localWa2Ver = "0", "winapp2.ini", localWa2Ver, latestWa2Ver)
        printUpdNotif(updateIsAvail, "Winapp2ool", currentVersion.ToString, latestVersion)

        printMenuOpt("Exit", "Exit the application")
        printMenuOpt("WinappDebug", "Check for and correct errors in winapp2.ini")
        printMenuOpt("Trim", "Debloat winapp2.ini for your system")
        printMenuOpt("Merge", "Merge the contents of an ini file into winapp2.ini")
        printMenuOpt("Diff", "Observe the changes between two winapp2.ini files")
        printMenuOpt("CCiniDebug", "Sort and trim ccleaner.ini")

        printBlankMenuLine()
        printMenuOpt("Downloader", "Download files from the Winapp2 GitHub")

        If waUpdateIsAvail Then
            printBlankMenuLine()
            printMenuOpt("Update", "Update your local copy of winapp2.ini")
            printMenuOpt("Update & Trim", "Download and trim the latest winapp2.ini")
            printMenuOpt("Show update diff", "See the difference between your local file and the latest")
        End If

        If updateIsAvail Then
            printBlankMenuLine()
            printMenuOpt("Update", "Get the latest winapp2ool.exe")
        End If

        If waUpdateIsAvail And updateIsAvail Then Console.WindowHeight += 2
        printMenuLine(menuStr02)
    End Sub

    Public Sub main()

        Console.Title = "Winapp2ool v" & currentVersion & " beta"
        Console.WindowWidth = 120
        initMenu("Winapp2ool - A multitool for winapp2.ini", 35)
        processCommandLineArgs()

        If suppressOutput Then Environment.Exit(1)

        Do Until exitCode
            checkUpdates()
            Console.Clear()
            printMenu()
            cwl()
            Console.Write("Enter a number: ")
            Dim input As String = Console.ReadLine

            Select Case True
                Case input = "0"
                    exitCode = True
                    cwl("Exiting...")
                    Environment.Exit(1)
                Case input = "1"
                    WinappDebug.main()
                    menuTopper = "Finished running WinappDebug"
                Case input = "2"
                    Trim.main()
                    menuTopper = "Finished running Trim"
                Case input = "3"
                    Merge.main()
                    menuTopper = "Finished running Merge"
                Case input = "4"
                    Diff.main()
                    menuTopper = "Finished running Diff"
                Case input = "5"
                    CCiniDebug.main()
                    menuTopper = "Finished running CCiniDebug"
                Case input = "6"
                    Downloader.main()
                    menuTopper = "Finished running Downloader"
                Case input = "7" And waUpdateIsAvail
                    Console.Clear()
                    Console.Write("Downloading, this may take a moment...")
                    remoteDownload(Environment.CurrentDirectory, "winapp2.ini", wa2Link, False)
                    undoAnyPendingExits()
                Case input = "8" And waUpdateIsAvail
                    Console.Clear()
                    Console.Write("Downloading & triming, this may take a moment...")
                    remoteTrim(eini, lwinapp2File, True, False)
                    undoAnyPendingExits()
                Case input = "9" And waUpdateIsAvail
                    Console.Clear()
                    Console.Write("Downloading & diffing, this may take a moment...")
                    remoteDiff(lwinapp2File, eini, eini, True, False, False)
                    undoAnyPendingExits()
                    menuTopper = "Diff Complete"
                Case (input = "10" And (updateIsAvail And waUpdateIsAvail)) Or (input = "7" And (Not waUpdateIsAvail And updateIsAvail))
                    Console.WriteLine("Downloading and updating winapp2ool.exe, this may take a moment...")
                    autoUpdate()
                Case Else
                    menuTopper = invInpStr
            End Select
        Loop
    End Sub

    'Check the versions of winapp2ool, .NET, and winapp2.ini and record if any are outdated
    Private Sub checkUpdates()
        Try
            'Query the latest winapp2ool.exe and winapp2.ini versions if we haven't already 
            If latestVersion = "" Then latestVersion = getRemoteFileDataAtLineNum(toolVerLink, 1)
            If latestWa2Ver = "" Then latestWa2Ver = getRemoteFileDataAtLineNum(wa2Link, 1).Split(CChar(" "))(2)
            'Check the local winapp2.ini version
            If Not File.Exists(Environment.CurrentDirectory & "\winapp2.ini") Then
                localWa2Ver = "0"
                menuTopper = "/!\ Winapp2.ini update check failed. /!\"
            Else
                Dim verStr As String = getFileDataAtLineNum(Environment.CurrentDirectory & "\winapp2.ini", 1).Split(CChar(" "))(2)
                localWa2Ver = verStr
            End If
        Catch ex As Exception
            menuTopper = "/!\ Update check failed. /!\"
            localWa2Ver = "0"
        End Try

        'Observe whether or not updates are available
        waUpdateIsAvail = CDbl(latestWa2Ver) > CDbl(localWa2Ver)
        updateIsAvail = CDbl(latestVersion) > currentVersion

        'winapp2ool requires .NET 4.6 or higher, all versions of which report the following version
        If Not Environment.Version.ToString = "4.0.30319.42000" Then dnfOOD = True

    End Sub

    'Print out exceptions and any other information related to them that a user might need
    Public Sub exc(ByRef ex As Exception)
        If ex.Message.ToString.Contains("SSL/TLS") Then
            cwl("Error: download could not be completed.")
            cwl("This issue is caused by an out of date .NET Framework.")
            cwl("Please update .NET Framework to version 4.6 or higher and try again.")
            cwl("If the issue persists after updating .NET Framework, please report this error on GitHub.")
        Else
            cwl("Error: " & ex.ToString)
            cwl("Please report this error on GitHub")
            cwl()
        End If
    End Sub

    'print an empty line as long as we're not supressing output
    Public Sub cwl()
        If Not suppressOutput Then Console.WriteLine()
    End Sub

    'print a line as long as we're not supressing output
    Public Sub cwl(msg As String)
        If Not suppressOutput Then Console.WriteLine(msg)
    End Sub

    'Prompt the user to change a file's parameters, flag it as changed, and mark settings as having changed
    Public Sub changeFileParams(ByRef someFile As iniFile, ByRef settingsChangedSetting As Boolean)
        fileChooser(someFile)
        settingsChangedSetting = True
        menuTopper = If(someFile.secondName = "", someFile.initName, "save file") & " " & If(exitCode, "parameter update aborted", "parameters updated")
        undoAnyPendingExits()
    End Sub

    'Toggle a parameter on or off and mark settings as having changed
    Public Sub toggleSettingParam(ByRef setting As Boolean, paramText As String, ByRef settingsChangedSetting As Boolean)
        setting = Not setting
        menuTopper = paramText & If(setting, "enabled", "disabled")
        settingsChangedSetting = True
    End Sub

End Module