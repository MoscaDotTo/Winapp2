'    Copyright (C) 2018-2019 Robbie Ward
' 
'    This file is a part of Winapp2ool
' 
'    Winapp2ool is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    Winap2ool is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with Winapp2ool.  If not, see <http://www.gnu.org/licenses/>.
Option Strict On
Imports Microsoft.Win32

Module Winapp2ool
    ''' <summary>Indicates whether or not winapp2ool is in "Non-CCleaner" mode and should collect the appropriate ini </summary>
    Public Property RemoteWinappIsNonCC As Boolean = False
    '''<summary>Indicates whether or not the .NET Framework installed on the current machine is below the targeted version</summary>
    Public Property DotNetFrameworkOutOfDate As Boolean = False
    ''' <summary>Indicates whether or not winapp2ool currently has access to the internet</summary>
    Public Property isOffline As Boolean = False
    '''<summary>Indicates that this build is beta and should check the beta branch link for updates</summary>
    Public Property isBeta As Boolean = True

    ''' <summary>Denies the ability to access online-only functions if offline</summary>
    Public Function denySettingOffline() As Boolean
        gLog("Action was unable to complete because winapp2ool is offline", isOffline)
        setHeaderText("This option is unavailable while in offline mode", True, isOffline)
        Return isOffline
    End Function

    ''' <summary>Prints the main menu to the user</summary>
    Private Sub printMenu()
        checkUpdates(Not isOffline)
        printMenuTop({}, False)
        print(4, "Winapp2ool is currently in offline mode", cond:=isOffline)
        print(4, "Your .NET Framework is out of date", cond:=DotNetFrameworkOutOfDate)
        printUpdNotif(waUpdateIsAvail, "winapp2.ini", localWa2Ver, latestWa2Ver)
        printUpdNotif(updateIsAvail, "Winapp2ool", currentVersion, latestVersion)
        print(1, "Exit", "Exit the application")
        print(1, "WinappDebug", "Check for and correct errors in winapp2.ini")
        print(1, "Trim", "Debloat winapp2.ini for your system")
        print(1, "Merge", "Merge the contents of an ini file into winapp2.ini")
        print(1, "Diff", "Observe the changes between two winapp2.ini files")
        print(1, "CCiniDebug", "Sort and trim ccleaner.ini", trailingBlank:=True)
        print(1, "Downloader", "Download files from the Winapp2 GitHub", closeMenu:=Not isOffline And Not waUpdateIsAvail And Not updateIsAvail)
        If waUpdateIsAvail And Not isOffline Then
            print(1, "Update", "Update your local copy of winapp2.ini", leadingBlank:=True)
            print(1, "Update & Trim", "Download and trim the latest winapp2.ini")
            print(1, "Show update diff", "See the difference between your local file and the latest", closeMenu:=Not updateIsAvail)
        End If
        print(1, "Update", "Get the latest winapp2ool.exe", updateIsAvail And Not DotNetFrameworkOutOfDate, True, closeMenu:=True)
        print(1, "Go online", "Retry your internet connection", isOffline, True, closeMenu:=True)
        Console.WindowHeight = If(waUpdateIsAvail And updateIsAvail, 32, 30)
    End Sub

    ''' <summary>Presents the winapp2ool menu to the user, initiates the main event loop for the application</summary>
    Public Sub main()
        gLog($"Starting application")
        Console.Title = $"Winapp2ool v{currentVersion}"
        Console.WindowWidth = 126
        ' winapp2ool requires .NET 4.6 or higher for full functionality, all versions of which report the following version
        If Not Environment.Version.ToString = "4.0.30319.42000" Then DotNetFrameworkOutOfDate = True
        gLog($".NET Framework is out of date. Found {Environment.Version.ToString}", DotNetFrameworkOutOfDate)
        ' winapp2ool requires internet access for some functions
        chkOfflineMode()
        processCommandLineArgs()
        If SuppressOutput Then Environment.Exit(1)
        initModule($"Winapp2ool v{currentVersion} - A multitool for winapp2.ini", AddressOf printMenu, AddressOf handleUserInput)
    End Sub

    ''' <summary>Handles the user input for the menu</summary>
    ''' <param name="input">The String containing the user input</param>
    Private Sub handleUserInput(input As String)
        Select Case True
            Case input = "0"
                ExitCode = True
                cwl("Exiting...")
                Environment.Exit(0)
            Case input = "1"
                initModule("WinappDebug", AddressOf WinappDebug.printMenu, AddressOf WinappDebug.handleUserInput)
            Case input = "2"
                initModule("Trim", AddressOf Trim.printMenu, AddressOf Trim.handleUserInput)
            Case input = "3"
                initModule("Merge", AddressOf Merge.printMenu, AddressOf Merge.handleUserInput)
            Case input = "4"
                initModule("Diff", AddressOf Diff.printMenu, AddressOf Diff.handleUserInput)
            Case input = "5"
                initModule("CCiniDebug", AddressOf CCiniDebug.printMenu, AddressOf CCiniDebug.handleUserInput)
            Case input = "6"
                If Not denySettingOffline() Then initModule("Downloader", AddressOf Downloader.printMenu, AddressOf Downloader.handleUserInput)
            Case input = "7" And isOffline
                chkOfflineMode()
                setHeaderText("Winapp2ool was unable to establish a network connection. You are still in offline mode.", True, isOffline)
            Case input = "7" And waUpdateIsAvail
                clrConsole()
                cwl("Downloading, this may take a moment...")
                download(New iniFile(Environment.CurrentDirectory, "winapp2.ini"), winapp2link, False)
                checkedForUpdates = False
            Case input = "8" And waUpdateIsAvail
                clrConsole()
                cwl("Downloading & trimming, this may take a moment...")
                remoteTrim(New iniFile("", ""), New iniFile(Environment.CurrentDirectory, "winapp2.ini"), True)
                checkedForUpdates = False
            Case input = "9" And waUpdateIsAvail
                clrConsole()
                cwl("Downloading & diffing, this may take a moment...")
                remoteDiff(New iniFile(Environment.CurrentDirectory, "winapp2.ini"))
                setHeaderText("Diff Complete")
            Case (input = "10" And (updateIsAvail And waUpdateIsAvail)) Or (input = "7" And (Not waUpdateIsAvail And updateIsAvail)) And Not DotNetFrameworkOutOfDate
                cwl("Downloading and updating winapp2ool.exe, this may take a moment...")
                autoUpdate()
            Case input = "m"
                initModule("Minefield", AddressOf Minefield.printMenu, AddressOf Minefield.handleUserInput)
            Case input = "savelog"
                GlobalLogFile.overwriteToFile(logger.toString)
            Case input = "printlog"
                printLog()
            Case Else
                setHeaderText(invInpStr, True)
        End Select
    End Sub

    ''' <summary>Attempts to return the Windows version number, return 0.0 if it cannot</summary>
    Public Function getWinVer() As Double
        gLog("Checking Windows version")
        ' We can return very quickly on Windows 10 using this registry key. Unknown if it exists on earlier versions
        If Registry.GetValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion", "CurrentMajorVersionNumber", Nothing) IsNot Nothing Then
            Dim tmp = Registry.GetValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion", "CurrentMajorVersionNumber", Nothing).ToString
            gLog($"Found Windows {tmp}")
            Return CDbl(tmp)
        End If
        Dim osVersion = System.Environment.OSVersion.ToString().Replace("Microsoft Windows NT ", "")
        Dim ver = osVersion.Split(CChar("."))
        Dim out = Val($"{ver(0)}.{ver(1)}")
        gLog($"Found Windows {out}")
        ' This might not act completely correctly on Windows 8.1 but usage of that seems small enough that it wont be an issue
        If Not {5.1, 6.0, 6.1, 6.2, 6.3}.Contains(out) Then
            Console.WriteLine("Unable to determine which version of Windows you are running.")
            Console.WriteLine()
            Console.WriteLine("If you see this message, please report your Windows version on GitHub along with the following information:")
            Console.WriteLine($"out: {out}")
            Console.WriteLine()
            Console.WriteLine("Press any key to continue")
            Console.ReadKey()
            out = 0.0
        End If
        Return out
    End Function

    ''' <summary>Appends a series of values onto a String</summary>
    ''' <param name="toAppend">The values to append</param>
    ''' <param name="out">The given string to be extended</param>
    Public Sub appendStrs(toAppend As String(), ByRef out As String, Optional delim As Boolean = False, Optional delimchar As Char = CChar(","))
        If Not delim Then
            For Each param In toAppend
                out += param
            Next
        Else
            For i = 0 To toAppend.Count - 2
                out += toAppend(i) & $"{delimchar} "
            Next
            out += toAppend.Last
        End If
    End Sub

    ''' <summary>Returns the first portion of a registry or filepath parameterization</summary>
    ''' <param name="val">The directory listing to be split</param>
    Public Function getFirstDir(val As String) As String
        Return val.Split(CChar("\"))(0)
    End Function

    ''' <summary>Initializes a module's menu, prints it, and handles the user input. Effectively the main event loop for winapp2ool and its components</summary>
    ''' <param name="name">The name of the module</param>
    ''' <param name="showMenu">The function that prints the module's menu</param>
    ''' <param name="handleInput">The function that handles the module's input</param>
    Public Sub initModule(name As String, showMenu As Action, handleInput As Action(Of String))
        gLog("", ascend:=True)
        gLog($"Loading module {name}")
        initMenu(name)
        Try
            Do Until ExitCode
                clrConsole()
                showMenu()
                Console.Write(Environment.NewLine & promptStr)
                handleInput(Console.ReadLine)
            Loop
            revertMenu()
            setHeaderText($"{name} closed")
            gLog($"Exiting {name}", descend:=True, leadr:=True)
        Catch ex As Exception
            exc(ex)
        End Try
    End Sub

    ''' <summary>Checks a String for casing errors against a provided array of cased strings, returns the input string if no error is detected</summary>
    ''' <param name="caseArray">The parent array of cased Strings</param>
    ''' <param name="inputText">The String to be checked for casing errors</param>
    Public Function getCasedString(caseArray As String(), inputText As String) As String
        For Each casedText In caseArray
            If inputText.Equals(casedText, StringComparison.InvariantCultureIgnoreCase) Then Return casedText
        Next
        Return inputText
    End Function

    ''' <summary>Ensures that an iniFile has content and informs the user if it does not. Returns false if there are no section.</summary>
    ''' <param name="iFile">An iniFile to be checked for content</param>
    Public Function enforceFileHasContent(iFile As iniFile) As Boolean
        If iFile.Sections.Count = 0 Then
            setHeaderText($"{iFile.Name} was empty or not found", True)
            gLog($"{iFile.Name} was empty or not found", indent:=True)
            Return False
        End If
        Return True
    End Function

    ''' <summary>
    ''' Waits for the user to press a key if we're not supressing output
    ''' </summary>
    Public Sub crk()
        If Not SuppressOutput Then Console.ReadKey()
    End Sub
End Module