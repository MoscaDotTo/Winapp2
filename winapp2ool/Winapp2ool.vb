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
    ''' <summary>Indicates that winapp2ool is in "Non-CCleaner" mode and should collect the appropriate ini from GitHub</summary>
    Public Property RemoteWinappIsNonCC As Boolean = False
    '''<summary>Indicates that the .NET Framework installed on the current machine is below the targeted version (.NET Framework 4.5)</summary>
    Public Property DotNetFrameworkOutOfDate As Boolean = False
    ''' <summary>Indicates that winapp2ool currently has access to the internet</summary>
    Public Property isOffline As Boolean = False
    '''<summary>Indicates that this build is beta and should check the beta branch link for updates</summary>
    Public Property isBeta As Boolean = True

    ''' <summary>Prints the main winapp2ool menu to the user</summary>
    Private Sub printMenu()
        checkUpdates(Not isOffline)
        printMenuTop({}, False)
        print(0, "Winapp2ool is currently in offline mode", cond:=isOffline, colorLine:=True, enStrCond:=(False), isCentered:=True, trailingBlank:=True)
        print(0, "Your .NET Framework is out of date", cond:=DotNetFrameworkOutOfDate, colorLine:=True, enStrCond:=(False), isCentered:=True, trailingBlank:=True)
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

    ''' <summary>Processes the commandline args and then initalizes the main winapp2ool module</summary>
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
    ''' <param name="input">The user's input</param>
    Private Sub handleUserInput(input As String)
        Select Case True
            Case input = "0"
                exitModule()
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
                waUpdateIsAvail = False
            Case input = "8" And waUpdateIsAvail
                clrConsole()
                cwl("Downloading & trimming, this may take a moment...")
                remoteTrim(New iniFile("", ""), New iniFile(Environment.CurrentDirectory, "winapp2.ini"), True)
                waUpdateIsAvail = False
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

    ''' <summary>Checks the version of Windows on the current system and returns it as a Double</summary>
    ''' <returns>The Windows version running on the machine, <c>0.0</c> if the windows version cannot be determined</returns>
    Public Function getWinVer() As Double
        gLog("Checking Windows version")
        Dim osVersion = System.Environment.OSVersion.ToString().Replace("Microsoft Windows NT ", "")
        Dim ver = osVersion.Split(CChar("."))
        Dim out = Val($"{ver(0)}.{ver(1)}")
        gLog($"Found Windows {out}")
        Return out
    End Function

    ''' <summary>Returns the first portion of a registry or filepath parameterization</summary>
    ''' <param name="val">A Windows filesystem or registry path from which the root should be returned</param>
    ''' <returns>The root directory given by <paramref name="val"/></returns>
    Public Function getFirstDir(val As String) As String
        Return val.Split(CChar("\"))(0)
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

    ''' <summary>Waits for the user to press a key if <c>SuppressOutput</c> is <c>False</c></summary>
    Public Sub crk()
        If Not SuppressOutput Then Console.ReadKey()
    End Sub
End Module