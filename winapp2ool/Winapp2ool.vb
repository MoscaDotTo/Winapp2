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
    ' Update and compatibility settings
    Public currentVersion As String = "1.02"
    Dim latestVersion As String = ""
    Dim latestWa2Ver As String = ""
    Dim localWa2Ver As String = ""
    Public checkedForUpdates As Boolean = False
    Dim updateIsAvail As Boolean = False
    Dim waUpdateIsAvail As Boolean = False
    Public isOffline As Boolean
    Public dnfOOD As Boolean = False
    ' Dummy iniFile obj for when we don't need actual data or when it'll be filled in later
    Public eini As iniFile = New iniFile("", "")
    Public lwinapp2File As iniFile = New iniFile(Environment.CurrentDirectory, "winapp2.ini")

    ''' <summary>
    ''' Informs the user when an update is available
    ''' </summary>
    ''' <param name="cond">The update condition</param>
    ''' <param name="updName">The item (winapp2.ini or winapp2ool) for which there is a pending update</param>
    ''' <param name="oldVer">The old (currently in use) version</param>
    ''' <param name="newVer">The updated version pending download</param>
    Private Sub printUpdNotif(cond As Boolean, updName As String, oldVer As String, newVer As String)
        If cond Then
            print(0, $"A new version of {updName} is available!", isCentered:=True)
            print(0, $"Current  : v{oldVer}", isCentered:=True)
            print(0, $"Available: v{newVer}", trailingBlank:=True, isCentered:=True)
        End If
    End Sub

    ''' <summary>
    ''' Denies the ability to access online-only functions if offline
    ''' </summary>
    ''' <returns></returns>
    Public Function denySettingOffline() As Boolean
        If isOffline Then menuHeaderText = "This option is unavailable while in offline mode"
        Return isOffline
    End Function

    ''' <summary>
    ''' Prints the main menu to the user
    ''' </summary>
    Private Sub printMenu()
        printMenuTop({}, False)
        print(4, "Winapp2ool is currently in offline mode", cond:=isOffline)
        print(4, "Your .NET Framework is out of date", cond:=dnfOOD)
        printUpdNotif(waUpdateIsAvail And Not localWa2Ver = "0", "winapp2.ini", localWa2Ver, latestWa2Ver)
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
        print(1, "Update", "Get the latest winapp2ool.exe", updateIsAvail And Not dnfOOD, True, closeMenu:=True)
        print(1, "Go online", "Retry your internet connection", isOffline, True, closeMenu:=True)
        Console.WindowHeight = If(waUpdateIsAvail And updateIsAvail, 32, 30)
    End Sub

    ''' <summary>
    ''' Presents the winapp2ool menu to the user, initiates the main event loop for the application
    ''' </summary>
    Public Sub main()
        Console.Title = $"Winapp2ool v{currentVersion}"
        Console.WindowWidth = 120
        ' winapp2ool requires .NET 4.6 or higher for full functionality, all versions of which report the following version
        If Not Environment.Version.ToString = "4.0.30319.42000" Then dnfOOD = True
        ' winapp2ool requires internet access for some functions
        chkOfflineMode()
        processCommandLineArgs()
        If suppressOutput Then Environment.Exit(1)
        initModule($"Winapp2ool v{currentVersion} - A multitool for winapp2.ini", AddressOf printMenu, AddressOf handleUserInput, Not isOffline)
    End Sub

    ''' <summary>
    ''' Handles the user input for the menu
    ''' </summary>
    ''' <param name="input">The String containing the user input</param>
    Private Sub handleUserInput(input As String)
        Select Case True
            Case input = "0"
                exitCode = True
                cwl("Exiting...")
                Environment.Exit(0)
            Case input = "1"
                initModule("WinappDebug", AddressOf WinappDebug.printMenu, AddressOf WinappDebug.handleUserInput)
            Case input = "2"
                initModule("Trim", AddressOf Trim.printmenu, AddressOf Trim.handleUserInput)
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
            Case input = "7" And waUpdateIsAvail
                Console.Clear()
                Console.Write("Downloading, this may take a moment...")
                remoteDownload(Environment.CurrentDirectory, "winapp2.ini", wa2Link, False)
                checkedForUpdates = False
                undoAnyPendingExits()
            Case input = "8" And waUpdateIsAvail
                Console.Clear()
                Console.Write("Downloading & triming, this may take a moment...")
                remoteTrim(eini, lwinapp2File, True, False)
                checkedForUpdates = False
                undoAnyPendingExits()
            Case input = "9" And waUpdateIsAvail
                Console.Clear()
                Console.Write("Downloading & diffing, this may take a moment...")
                remoteDiff(lwinapp2File)
                undoAnyPendingExits()
                menuHeaderText = "Diff Complete"
            Case (input = "10" And (updateIsAvail And waUpdateIsAvail)) Or (input = "7" And (Not waUpdateIsAvail And updateIsAvail)) And Not dnfOOD
                Console.WriteLine("Downloading and updating winapp2ool.exe, this may take a moment...")
                autoUpdate()
            Case input.ToLower = "m"
                initModule("Minefield", AddressOf Minefield.printMenu, AddressOf Minefield.handleUserInput)
            Case Else
                menuHeaderText = invInpStr
        End Select
    End Sub

    ''' <summary>
    ''' Attempts to return the version number from the first line of winapp2.ini, returns "000000" if it can't
    ''' </summary>
    Private Sub getLocalWinapp2Version()
        If Not File.Exists(Environment.CurrentDirectory & "\winapp2.ini") Then updateCheckFailed("winapp2.ini") : Exit Sub
        Dim localStr As String = getFileDataAtLineNum(Environment.CurrentDirectory & "\winapp2.ini", remote:=False).ToLower
        localWa2Ver = If(Not localStr.Contains("version"), "000000", localStr.Split(CChar(" "))(2))
    End Sub

    ''' <summary>
    ''' Checks the versions of winapp2ool, .NET, and winapp2.ini and records if any are outdated.
    ''' </summary>
    Private Sub checkUpdates()
        If checkedForUpdates Then Exit Sub
        Try
            ' Query the latest winapp2ool.exe and winapp2.ini versions 
            latestVersion = getFileDataAtLineNum(toolVerLink)
            latestWa2Ver = getFileDataAtLineNum(wa2Link).Split(CChar(" "))(2)
            ' This should only be true if a user somehow has internet but cannot otherwise connect to the GitHub resources used to check for updates
            ' In this instance we should consider the update check to have failed and put the application into offline mode
            If latestVersion = "" Or latestWa2Ver = "" Then updateCheckFailed("online", True) : Exit Try
            ' Observe whether or not updates are available, using val to avoid conversion mistakes
            updateIsAvail = Val(latestVersion) > Val(currentVersion)
            getLocalWinapp2Version()
            waUpdateIsAvail = Val(latestWa2Ver) > Val(localWa2Ver)
            checkedForUpdates = True
        Catch ex As Exception
            exc(ex)
            updateCheckFailed("winapp2ool or winapp2.ini")
        End Try
    End Sub

    ''' <summary>
    ''' Handles the case where the update check has failed
    ''' </summary>
    ''' <param name="name">The name of the component whose update check failed</param>
    ''' <param name="chkOnline">A flag specifying that that the internet connection should be retested</param>
    Private Sub updateCheckFailed(name As String, Optional chkOnline As Boolean = False)
        menuHeaderText = $"/!\ {name} update check failed. /!\"
        localWa2Ver = "000000"
        If chkOnline Then chkOfflineMode()
    End Sub

    ''' <summary>
    ''' Updates the offline status of winapp2ool
    ''' </summary>
    Private Sub chkOfflineMode()
        isOffline = Not checkOnline()
    End Sub

    ''' <summary>
    ''' Prompts the user to change a file's parameters, marks both settings and the file as having been changed 
    ''' </summary>
    ''' <param name="someFile">A file whose parameters will be changed</param>
    ''' <param name="settingsChangedSetting">The boolean indicating that a setting has been changed</param>
    Public Sub changeFileParams(ByRef someFile As iniFile, ByRef settingsChangedSetting As Boolean)
        fileChooser(someFile)
        settingsChangedSetting = True
        menuHeaderText = $"{If(someFile.secondName = "", someFile.initName, "save file")} parameters update{If(exitCode, " aborted", "d")}"
        undoAnyPendingExits()
    End Sub

    ''' <summary>
    ''' Toggles a setting's boolean state and marks its tracker true
    ''' </summary>
    ''' <param name="setting">A boolean to be toggled</param>
    ''' <param name="paramText">The string explaining the setting being toggled</param>
    ''' <param name="settingsChangedSetting">The boolean indicating that the setting has been modified</param>
    Public Sub toggleSettingParam(ByRef setting As Boolean, paramText As String, ByRef settingsChangedSetting As Boolean)
        setting = Not setting
        menuHeaderText = $"{paramText} {If(setting, "enabled", "disabled")}"
        settingsChangedSetting = True
    End Sub

    ''' <summary>
    ''' Returns 1 or 2 newline characters conditionally
    ''' </summary>
    ''' <param name="cond">The parameter under which to return two newlines</param>
    ''' <returns></returns>
    Public Function prependNewLines(Optional cond As Boolean = False) As String
        Return If(cond, Environment.NewLine & Environment.NewLine, Environment.NewLine)
    End Function

    ''' <summary>
    ''' Attempts to return the Windows version number, return 0.0 if it cannot
    ''' </summary>
    ''' <returns></returns>
    Public Function getWinVer() As Double
        ' We can return very quickly on Windows 10 using this registry key. Unknown if it exists on earlier versions
        If Registry.GetValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion", "CurrentMajorVersionNumber", Nothing) IsNot Nothing Then
            Dim tmp As String = Registry.GetValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion", "CurrentMajorVersionNumber", Nothing).ToString
            Return CDbl(tmp)
        End If
        Dim osVersion As String = System.Environment.OSVersion.ToString().Replace("Microsoft Windows NT ", "")
        Dim ver As String() = osVersion.Split(CChar("."))
        Dim out As Double = Val($"{ver(0)}.{ver(1)}")
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

    ''' <summary>
    ''' Handles toggling downloading of winapp2.ini from menus
    ''' </summary>
    ''' <param name="download">The download Boolean</param>
    ''' <param name="settingsChanged">The Boolean indicating that settings have changed</param>
    Public Sub toggleDownload(ByRef download As Boolean, ByRef settingsChanged As Boolean)
        If Not denySettingOffline() Then toggleSettingParam(download, "Downloading ", settingsChanged)
    End Sub

    ''' <summary>
    ''' Returns the online download status of winapp2.ini as a String, empty string if not downloading
    ''' </summary>
    ''' <param name="d">The normal download Boolean</param>
    ''' <param name="dncc">The non-CCleaner download Boolean</param>
    ''' <returns></returns>
    Public Function GetNameFromDL(d As Boolean, dncc As Boolean) As String
        Return If(d, If(dncc, "Online (non-ccleaner)", "Online"), "")
    End Function

    ''' <summary>
    ''' Resets a module's settings to the defaults
    ''' </summary>
    ''' <param name="name">The name of the module</param>
    ''' <param name="setDefaultParams">The function that resets the module's settings</param>
    Public Sub resetModuleSettings(name As String, setDefaultParams As Action)
        setDefaultParams()
        menuHeaderText = $"{name} settings have been reset to their defaults."
    End Sub

    ''' <summary>
    ''' Appends a series of values onto a String
    ''' </summary>
    ''' <param name="toAppend">The values to append</param>
    ''' <param name="out">The given string to be extended</param>
    Public Sub appendStrs(toAppend As String(), ByRef out As String)
        For Each param In toAppend
            out += param
        Next
    End Sub

    ''' <summary>
    ''' Initializes a module's menu, prints it, and handles the user input. Effectively the main event loop for winapp2ool and its components
    ''' </summary>
    ''' <param name="name">The name of the module</param>
    ''' <param name="callMenu">The function that prints the module's menu</param>
    ''' <param name="handleInput">The function that handle's the module's input</param>
    ''' <param name="chkUpd">The Boolean indicating that winapp2ool should check for updates (only called by the main menu)</param>
    Public Sub initModule(name As String, callMenu As Action, handleInput As Action(Of String), Optional chkUpd As Boolean = False)
        initMenu(name)
        Do Until exitCode
            If chkUpd Then checkUpdates()
            Console.Clear()
            callMenu()
            Console.Write(Environment.NewLine & promptStr)
            handleInput(Console.ReadLine)
        Loop
        revertMenu()
        menuHeaderText = $"{name} closed"
    End Sub
End Module