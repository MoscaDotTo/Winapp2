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
Imports System.IO
Imports Microsoft.Win32

Module Winapp2ool
    ''' <summary>Indicates the latest available verson of winapp2ool from GitHub</summary>
    Private Property latestVersion As String = ""
    ''' <summary>Indicates the latest available version of winapp2.ini from GitHub</summary>
    Private Property latestWa2Ver As String = ""
    ''' <summary>Indicates the local version of winapp2.ini (if available)</summary>
    Private Property localWa2Ver As String = "000000"
    ''' <summary>Indicates that a winapp2ool update is available from GitHub</summary>
    Private Property updateIsAvail As Boolean = False
    ''' <summary>Indicates that a winapp2.ini update is available from GitHub</summary>
    Private Property waUpdateIsAvail As Boolean = False
    ''' <summary>The current version of the executable as used for version checking against GitHub</summary>
    Public ReadOnly Property currentVersion As String = System.Reflection.Assembly.GetExecutingAssembly.FullName.Split(CChar(","))(1).Substring(9)
    ''' <summary>Indicates whether or not winapp2ool is in "Non-CCleaner" mode and should collect the appropriate ini </summary>
    Public Property RemoteWinappIsNonCC As Boolean = False
    '''<summary>Indicates whether or not the .NET Framework installed on the current machine is below the targeted version</summary>
    Public Property DotNetFrameworkOutOfDate As Boolean = False
    ''' <summary>Indicates whether or not winapp2ool currently has access to the internet</summary>
    Public Property isOffline As Boolean = False
    ''' <summary>Indicates whether or not winapp2ool has already checked for updates</summary>
    Public Property checkedForUpdates As Boolean = False
    '''<summary>Indicates that this build is beta and should check the beta branch link for updates</summary>
    Public Property isBeta As Boolean = True

    ''' <summary>Returns the link to the appropriate winapp2.ini file based on the mode the tool is in</summary>
    Public Function winapp2link() As String
        Return If(RemoteWinappIsNonCC, nonccLink, wa2Link)
    End Function

    Public Function toolExeLink() As String
        Return If(isBeta, betaToolLink, toolLink)
    End Function

    ''' <summary>Informs the user when an update is available</summary>
    ''' <param name="cond">The update condition</param>
    ''' <param name="updName">The item (winapp2.ini or winapp2ool) for which there is a pending update</param>
    ''' <param name="oldVer">The old (currently in use) version</param>
    ''' <param name="newVer">The updated version pending download</param>
    Private Sub printUpdNotif(cond As Boolean, updName As String, oldVer As String, newVer As String)
        If cond Then
            gLog($"Update available for {updName} from {oldVer} to {newVer}")
            Console.ForegroundColor = ConsoleColor.Green
            print(0, $"A new version of {updName} is available!", isCentered:=True)
            print(0, $"Current  : v{oldVer}", isCentered:=True)
            print(0, $"Available: v{newVer}", trailingBlank:=True, isCentered:=True)
            Console.ResetColor()
        End If
    End Sub

    ''' <summary>Denies the ability to access online-only functions if offline</summary>
    Public Function denySettingOffline() As Boolean
        gLog("Action was unable to complete because winapp2ool is offline", isOffline)
        If isOffline Then setHeaderText("This option is unavailable while in offline mode", True)
        Return isOffline
    End Function

    ''' <summary>Prints the main menu to the user</summary>
    Private Sub printMenu()
        If Not isOffline Then checkUpdates()
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
                If isOffline Then setHeaderText("Winapp2ool was unable to establish a network connection. You are still in offline mode.", True)
            Case input = "7" And waUpdateIsAvail
                clrConsole()
                cwl("Downloading, this may take a moment...")
                remoteDownload(Environment.CurrentDirectory, "winapp2.ini", winapp2link, False)
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
                logger.printLog()
            Case Else
                setHeaderText(invInpStr, True)
        End Select
    End Sub

    ''' <summary>Attempts to return the version number from the first line of winapp2.ini, returns "000000" if it can't</summary>
    Private Sub getLocalWinapp2Version()
        If Not File.Exists(Environment.CurrentDirectory & "\winapp2.ini") Then localWa2Ver = "000000 (File not found)" : Exit Sub
        Dim localStr = getFileDataAtLineNum(Environment.CurrentDirectory & "\winapp2.ini", remote:=False).ToLower
        localWa2Ver = If(Not localStr.Contains("version"), "000000", localStr.Split(CChar(" "))(2))
    End Sub

    ''' <summary>Checks the versions of winapp2ool, .NET, and winapp2.ini and records if any are outdated.</summary>
    Private Sub checkUpdates()
        If checkedForUpdates Then Exit Sub
        Try
            gLog("Checking for updates")
            ' Query the latest winapp2ool.exe and winapp2.ini versions 
            Dim tmp = Environment.GetEnvironmentVariable("temp")
            remoteDownload($"{Environment.GetEnvironmentVariable("temp")}\", "winapp2ool.exe", toolExeLink, False)
            latestVersion = System.Reflection.Assembly.LoadFile($"{Environment.GetEnvironmentVariable("temp")}\winapp2ool.exe").FullName.Split(CChar(","))(1).Substring(9)
            latestWa2Ver = getFileDataAtLineNum(winapp2link).Split(CChar(" "))(2)
            ' This should only be true if a user somehow has internet but cannot otherwise connect to the GitHub resources used to check for updates
            ' In this instance we should consider the update check to have failed and put the application into offline mode
            If latestVersion = "" Or latestWa2Ver = "" Then updateCheckFailed("online", True) : Exit Try
            ' Observe whether or not updates are available, using val to avoid conversion mistakes
            updateIsAvail = Val(latestVersion.Replace(".", "")) > Val(currentVersion.Replace(".", ""))
            getLocalWinapp2Version()
            waUpdateIsAvail = Val(latestWa2Ver) > Val(localWa2Ver)
            checkedForUpdates = True
            gLog("Update check complete:")
            gLog($"Winapp2ool:")
            gLog("Local: " & currentVersion, indent:=True)
            gLog("Remote: " & latestVersion, indent:=True)
            gLog("Winapp2.ini:")
            gLog("Local:" & localWa2Ver, indent:=True)
            gLog("Remote: " & latestWa2Ver, indent:=True)
        Catch ex As Exception
            exc(ex)
            updateCheckFailed("winapp2ool or winapp2.ini")
        End Try
    End Sub

    ''' <summary>Handles the case where the update check has failed</summary>
    ''' <param name="name">The name of the component whose update check failed</param>
    ''' <param name="chkOnline">A flag specifying that the internet connection should be retested</param>
    Private Sub updateCheckFailed(name As String, Optional chkOnline As Boolean = False)
        setHeaderText($"/!\ {name} update check failed. /!\", True)
        localWa2Ver = "000000"
        If chkOnline Then chkOfflineMode()
    End Sub

    ''' <summary>Updates the offline status of winapp2ool</summary>
    Private Sub chkOfflineMode()
        gLog("Checking online status")
        isOffline = Not checkOnline()
    End Sub

    ''' <summary>Prompts the user to change a file's parameters, marks both settings and the file as having been changed </summary>
    ''' <param name="someFile">A file whose parameters will be changed</param>
    ''' <param name="settingsChangedSetting">The boolean indicating that a setting has been changed</param>
    Public Sub changeFileParams(ByRef someFile As iniFile, ByRef settingsChangedSetting As Boolean)
        Dim curName = someFile.Name
        Dim curDir = someFile.Dir
        initModule("File Chooser", AddressOf someFile.printFileChooserMenu, AddressOf someFile.handleFileChooserInput)
        settingsChangedSetting = True
        Dim fileChanged = Not someFile.Name = curName Or Not someFile.Dir = curDir
        setHeaderText($"{If(someFile.SecondName = "", someFile.InitName, "save file")} parameters update{If(Not fileChanged, " aborted", "d")}", Not fileChanged)
    End Sub

    ''' <summary>Toggles a setting's boolean state and marks its tracker true</summary>
    ''' <param name="setting">A boolean to be toggled</param>
    ''' <param name="paramText">The string explaining the setting being toggled</param>
    ''' <param name="settingsChangedSetting">The boolean indicating that the setting has been modified</param>
    Public Sub toggleSettingParam(ByRef setting As Boolean, paramText As String, ByRef settingsChangedSetting As Boolean)
        gLog($"Toggling {paramText}", indent:=True)
        setHeaderText($"{paramText} {enStr(setting)}d")
        setting = Not setting
        settingsChangedSetting = True
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

    ''' <summary>Handles toggling downloading of winapp2.ini from menus</summary>
    ''' <param name="download">The download Boolean</param>
    ''' <param name="settingsChanged">The Boolean indicating that settings have changed</param>
    Public Sub toggleDownload(ByRef download As Boolean, ByRef settingsChanged As Boolean)
        If Not denySettingOffline() Then toggleSettingParam(download, "Downloading", settingsChanged)
    End Sub

    ''' <summary>Returns the online download status (name) of winapp2.ini as a String, empty string if not downloading</summary>
    ''' <param name="shouldDownload">The boolean indicating whether or not a module will be downloading </param>
    Public Function GetNameFromDL(shouldDownload As Boolean) As String
        Return If(shouldDownload, If(RemoteWinappIsNonCC, "Online (Non-CCleaner)", "Online"), "")
    End Function

    ''' <summary>Resets a module's settings to the defaults</summary>
    ''' <param name="name">The name of the module</param>
    ''' <param name="setDefaultParams">The function that resets the module's settings</param>
    Public Sub resetModuleSettings(name As String, setDefaultParams As Action)
        gLog($"Restoring {name}'s module settings to default", indent:=True)
        setDefaultParams()
        setHeaderText($"{name} settings have been reset to their defaults.")
    End Sub

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