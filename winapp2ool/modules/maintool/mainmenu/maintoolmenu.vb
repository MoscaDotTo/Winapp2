'    Copyright (C) 2018-2020 Robbie Ward
' 
'    This file is a part of Winapp2ool
' 
'    Winapp2ool is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    Winapp2ool is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with Winapp2ool.  If not, see <http://www.gnu.org/licenses/>.
Option Strict On
''' <summary> Displays the main winapp2ool menu to the user and handles input from that menu </summary>
Module maintoolmenu

    ''' <summary> Prints the main winapp2ool menu to the user </summary>
    ''' Docs last updated: 2020-07-14 | Code last updated: 2020-07-14
    Public Sub printToolMainMenu()
        checkUpdates(Not isOffline And Not checkedForUpdates)
        printMenuTop(Array.Empty(Of String)(), False)
        print(0, "Winapp2ool is currently in offline mode", cond:=isOffline, colorLine:=True, enStrCond:=(False), isCentered:=True, trailingBlank:=True)
        print(0, "Your .NET Framework is out of date", cond:=DotNetFrameworkOutOfDate, colorLine:=True, enStrCond:=(False), isCentered:=True, trailingBlank:=True)
        print(0, "Winapp2ool is currently running from the temporary folder, some functions may be impacted", cond:=cantDownloadExecutable, colorLine:=True, enStrCond:=(False), isCentered:=True, trailingBlank:=True)
        printUpdNotif(waUpdateIsAvail, "winapp2.ini", localWa2Ver, latestWa2Ver)
        printUpdNotif(updateIsAvail, "Winapp2ool", currentVersion, latestVersion)
        print(1, "Exit", "Exit the application")
        print(1, "WinappDebug", "Check for and correct errors in winapp2.ini")
        print(1, "Trim", "Debloat winapp2.ini for your system")
        print(1, "Merge", "Merge the contents of an ini file into winapp2.ini")
        print(1, "Diff", "Observe the changes between two winapp2.ini files")
        print(1, "CCiniDebug", "Sort and trim ccleaner.ini", trailingBlank:=True)
        print(1, "Downloader", "Download files from the Winapp2 GitHub")
        print(1, "Settings", "Manage Winapp2ool's settings", closeMenu:=Not (isOffline Or waUpdateIsAvail Or updateIsAvail), arbitraryColor:=ConsoleColor.Yellow, colorLine:=True, useArbitraryColor:=True)
        If waUpdateIsAvail And Not isOffline Then
            print(1, "Update", "Update your local copy of winapp2.ini", leadingBlank:=True)
            print(1, "Update & Trim", "Download and trim the latest winapp2.ini")
            print(1, "Show update diff", "See the difference between your local file and the latest", closeMenu:=Not updateIsAvail)
        End If
        print(1, "Update", "Get the latest Winapp2ool.exe", updateIsAvail And Not DotNetFrameworkOutOfDate, True, closeMenu:=True)
        print(1, "Go online", "Retry your internet connection", isOffline, True, closeMenu:=True)
        Console.WindowHeight = If(waUpdateIsAvail And updateIsAvail, 34, 32)
    End Sub

    ''' <summary> Handles the user input for the menu </summary>
    ''' <param name="input"> The <c> String </c> containing the user's input </param>
    ''' Docs last updated: 2020-07-14 | Code last updated: 2020-07-14
    Public Sub handleToolMainUserInput(input As String)
        Select Case True
            Case input = "0"
                exitModule()
                cwl("Exiting...")
                Environment.Exit(0)
            Case input = "1"
                initModule(NameOf(WinappDebug), AddressOf WinappDebug.printMenu, AddressOf WinappDebug.handleUserInput)
            Case input = "2"
                initModule(NameOf(Trim), AddressOf Trim.printMenu, AddressOf Trim.handleUserInput)
            Case input = "3"
                initModule(NameOf(Merge), AddressOf Merge.printMenu, AddressOf Merge.handleUserInput)
            Case input = "4"
                initModule(NameOf(Diff), AddressOf printDiffMainMenu, AddressOf handleDiffMainMenuUserInput)
            Case input = "5"
                initModule(NameOf(CCiniDebug), AddressOf printCCDBMainMenu, AddressOf handleCCDBMainMenuUserInput)
            Case input = "6"
                If Not denySettingOffline() Then initModule("Downloader", AddressOf printDownloadMainMenu, AddressOf handleDownloadUserInput)
            Case input = "7"
                initModule("Winapp2ool Settings", AddressOf printMainToolSettingsMenu, AddressOf handleMainToolSettingsInput)
            Case input = "8" And isOffline
                chkOfflineMode()
                setHeaderText("Winapp2ool was unable to establish a network connection. You are still in offline mode.", True, isOffline)
            Case input = "8" And waUpdateIsAvail
                clrConsole()
                cwl("Downloading, this may take a moment...")
                download(New iniFile(Environment.CurrentDirectory, "winapp2.ini"), winapp2link, False)
                waUpdateIsAvail = False
            Case input = "9" And waUpdateIsAvail
                clrConsole()
                cwl("Downloading & trimming, this may take a moment...")
                remoteTrim(New iniFile(), New iniFile(Environment.CurrentDirectory, "winapp2.ini"), True)
                waUpdateIsAvail = False
            Case input = "10" And waUpdateIsAvail
                clrConsole()
                cwl("Downloading & diffing, this may take a moment...")
                remoteDiff(New iniFile(Environment.CurrentDirectory, "winapp2.ini"))
                setHeaderText("Diff Complete")
            Case (input = "11" And (updateIsAvail And waUpdateIsAvail)) Or (input = "8" And (Not waUpdateIsAvail And updateIsAvail)) And Not (DotNetFrameworkOutOfDate Or cantDownloadExecutable)
                cwl("Downloading and updating Winapp2ool.exe, this may take a moment...")
                autoUpdate()
            Case input = "m"
                initModule("Minefield", AddressOf Minefield.printMenu, AddressOf Minefield.handleUserInput)
            Case input = "savelog"
                GlobalLogFile.overwriteToFile(logger.toString)
            Case input = "printlog"
                printLog()
            Case input = "forceupdate"
                autoUpdate()
            Case Else
                setHeaderText(invInpStr, True)
        End Select
    End Sub

End Module
