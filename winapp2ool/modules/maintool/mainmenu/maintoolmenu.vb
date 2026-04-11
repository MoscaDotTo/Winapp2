'    Copyright (C) 2018-2025 Hazel Ward
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

''' <summary>
''' Displays the main winapp2ool menu to the user and handles input from that menu accordingly
''' </summary>
Module maintoolmenu

    ''' <summary>
    ''' Builds the main winapp2ool menu
    ''' </summary>
    Private Function buildToolMainMenu() As MenuSection

        If Console.WindowWidth < 130 Then Console.WindowWidth = 130

        checkUpdates(Not isOffline AndAlso Not checkedForUpdates)
        Dim UpdatesAvailable = Not isOffline AndAlso (waUpdateIsAvail OrElse updateIsAvail)

        Dim menuDesc = {"Welcome to Winapp2ool! Check out the ReadMe on GitHub for help!"}
        Dim headerColor = If(UpdatesAvailable, ConsoleColor.Green, ConsoleColor.Cyan)

        Dim menu = MenuSection.CreateCompleteMenu($"{NameOf(Winapp2ool)} v{currentVersion}", menuDesc, headerColor, True)

        menu.AddBlank() _
            .AddDispatchedOption("WinappDebug", "Scan for and correct style and syntax errors in winapp2.ini",
                Sub() initModule(NameOf(WinappDebug), AddressOf printLintMainMenu, AddressOf handleLintUserInput)) _
            .AddDispatchedOption("Trim", "Optimize winapp2.ini for your system",
                Sub() initModule(NameOf(Trim), AddressOf printTrimMenu, AddressOf handleTrimUserInput)) _
            .AddDispatchedOption("Transmute", "Add, replace, or remove entire sections or individual keys from winapp2.ini",
                Sub() initModule(NameOf(Transmute), AddressOf printTransmuteMainMenu, AddressOf handleTransmuteUserInput)) _
            .AddDispatchedOption("Diff", "Generate a context-aware changelog between two winapp2.ini files",
                Sub() initModule(NameOf(Diff), AddressOf printDiffMainMenu, AddressOf handleDiffUserInput)) _
            .AddDispatchedOption("CCiniDebug", "Remove stale winapp2.ini configurations from ccleaner.ini",
                Sub() initModule(NameOf(CCiniDebug), AddressOf printCCDBMainMenu, AddressOf handleCCDBMUserInput)) _
            .AddDispatchedOption("Entry Lab", "Generate winapp2.ini entries from templates",
                Sub() initModule("Entry Lab", AddressOf printEntryLabMenu, AddressOf handleEntryLabInput)) _
            .AddDispatchedOption("Combine", "Join together a collection of ini files into one",
                Sub() initModule(NameOf(Combine), AddressOf printCombineMainMenu, AddressOf handleCombineUserInput)) _
            .AddDispatchedOption("CC7Patcher", "Install winapp2.ini for CCleaner 7",
                Sub() initModule(NameOf(CC7Patcher), AddressOf printCC7PatcherMenu, AddressOf handleCC7PatcherInput)).AddBlank() _
            .AddDispatchedOption("Downloader", "Download files from the Winapp2 GitHub",
                Sub() If Not denySettingOffline() Then initModule(NameOf(Downloader), AddressOf printDownloadMainMenu, AddressOf handleDownloadUserInput)) _
            .AddDispatchedColoredOption("Settings", "Manage Winapp2ool's settings", ConsoleColor.Yellow,
                Sub() initModule("Winapp2ool Settings", AddressOf printMainToolSettingsMenu, AddressOf handleMainToolSettingsInput))

        ' Update notifications
        getUpdateNotification(waUpdateIsAvail, "winapp2.ini", localWa2Ver, latestWa2Ver, menu)
        getUpdateNotification(updateIsAvail, "Winapp2ool", currentVersion, latestVersion, menu)

        ' User warnings for limited feature availability
        If isOffline OrElse DotNetFrameworkOutOfDate OrElse cantDownloadExecutable Then

            menu.AddBlank() _
                .AddColoredLine("Winapp2ool is currently in offline mode", ConsoleColor.Red, centered:=True, condition:=isOffline) _
                .AddColoredLine("Your .NET Framework is out of date", ConsoleColor.Red, centered:=True, condition:=DotNetFrameworkOutOfDate) _
                .AddColoredLine("Winapp2ool is unable to automatically update", ConsoleColor.Red, centered:=True, condition:=cantDownloadExecutable)

        End If

        menu.AddBlank(waUpdateIsAvail) _
            .AddDispatchedOption("Update Winapp2.ini", "Update your local copy of winapp2.ini",
                Sub()
                    clrConsole()
                    cwl("Downloading, this may take a moment...")
                    download(New iniFile(Environment.CurrentDirectory, "winapp2.ini"), getWinappLink, False)
                    waUpdateIsAvail = False
                End Sub, condition:=waUpdateIsAvail) _
            .AddDispatchedOption("Update & Trim", "Download and trim the latest winapp2.ini",
                Sub()
                    clrConsole()
                    cwl("Downloading & trimming, this may take a moment...")
                    remoteTrim(New iniFileChooser(Environment.CurrentDirectory, ""),
                               New iniFileChooser(Environment.CurrentDirectory, "winapp2.ini", "winapp2.ini"),
                               True)
                    waUpdateIsAvail = False
                End Sub, condition:=waUpdateIsAvail) _
            .AddDispatchedOption("Show winapp2.ini changelog", "See the difference between your local file and the latest",
                Sub()
                    clrConsole()
                    cwl("Downloading & diffing, this may take a moment...")
                    DiffRemoteFile(New iniFileChooser(Environment.CurrentDirectory, "winapp2.ini"))
                    setNextMenuHeaderText("Diff Complete", printColor:=ConsoleColor.Green)
                End Sub, condition:=waUpdateIsAvail) _
            .AddDispatchedOption("Show trimmed changelog", "See the difference between your trimmed local file and the latest",
                                 Sub()
                                     clrConsole()
                                     cwl("Downloading, trimming, & diffing, this may take a moment...")
                                     DiffRemoteFile(New iniFileChooser(Environment.CurrentDirectory, "winapp2.ini"), True)
                                     setNextMenuHeaderText("Diff Complete", printColor:=ConsoleColor.Green)
                                 End Sub, condition:=waUpdateIsAvail) _
            .AddBlank(waUpdateIsAvail) _
            .AddDispatchedOption("Update Winapp2ool", "Get the latest Winapp2ool.exe",
                Sub()
                    cwl("Downloading and updating Winapp2ool.exe, this may take a moment...")
                    autoUpdate()
                End Sub, condition:=updateIsAvail AndAlso Not DotNetFrameworkOutOfDate) _
            .AddDispatchedOption("Go online", "Retry your internet connection",
                Sub()
                    chkOfflineMode()
                    setNextMenuHeaderText("Winapp2ool was unable to establish a network connection. You are still in offline mode.", printColor:=ConsoleColor.Red)
                End Sub, condition:=isOffline)

        Return menu

    End Function

    ''' <summary>
    ''' Prints the main winapp2ool menu to the user
    ''' </summary>
    Public Sub printToolMainMenu()

        buildToolMainMenu().Print()

    End Sub

    ''' <summary>
    ''' Handles the user input for the menu
    ''' </summary>
    '''
    ''' <param name="input">
    ''' The user's input
    ''' </param>
    Public Sub handleToolMainUserInput(input As String)

        ' Hidden options not listed on the menu
        Select Case input
            Case "m"
                initModule("Minefield", AddressOf Minefield.printMenu, AddressOf Minefield.handleUserInput)
                Return
            Case "savelog"
                IO.File.WriteAllText(GlobalLogFile.Path(), logger.toString)
                Return
            Case "printlog"
                printLog()
                Return
            Case "forceupdate"
                autoUpdate()
                Return
        End Select

        Dim intInput As Integer
        If Not Integer.TryParse(input, intInput) Then
            setNextMenuHeaderText(invInpStr, printColor:=ConsoleColor.Red)
            Return
        End If

        If intInput = 0 Then
            exitModule()
            Return
        End If

        If Not buildToolMainMenu().Dispatch(intInput) Then
            setNextMenuHeaderText(invInpStr, printColor:=ConsoleColor.Red)
        End If

    End Sub

    ''' <summary>
    ''' Adds information to the menu indicating to the user that an
    ''' update is available for <see cref="winapp2ool"/> or winapp2.ini
    ''' </summary>
    '''
    ''' <param name="cond">
    ''' Indicates whether or not an update for <see cref="winapp2ool"/>
    ''' or winapp2.ini is available <br />
    ''' <c> True </c> if an update is available <br />
    ''' <c> False </c> if no update is available
    ''' </param>
    '''
    ''' <param name="updName">
    ''' The name of the file (winapp2.ini or <see cref="winapp2ool"/>)
    ''' for which there is a pending update
    ''' </param>
    '''
    ''' <param name="oldVer">
    ''' The old (currently in use) version
    ''' </param>
    '''
    ''' <param name="newVer">
    ''' The updated version pending download
    ''' </param>
    Private Sub getUpdateNotification(cond As Boolean,
                                      updName As String,
                                      oldVer As String,
                                      newVer As String,
                                      menu As MenuSection)

        If Not cond Then Return

        gLog($"Update available for {updName} from {oldVer} to {newVer}")

        menu.AddBlank() _
            .AddColoredLine($"A new version of {updName} is available!", ConsoleColor.Green, True) _
            .AddColoredLine($"Current: v{oldVer}", ConsoleColor.Green, True) _
            .AddColoredLine($"Available: v{newVer}", ConsoleColor.Green, True)

        Console.WindowHeight += 2

    End Sub

End Module
