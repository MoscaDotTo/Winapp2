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
    ''' Prints the main winapp2ool menu to the user 
    ''' </summary>
    Public Sub printToolMainMenu()

        If Console.WindowWidth < 130 Then Console.WindowWidth = 130

        checkUpdates(Not isOffline AndAlso Not checkedForUpdates)
        Dim UpdatesAvailable = Not isOffline AndAlso (waUpdateIsAvail OrElse updateIsAvail)

        ' User warnings for limited feature availability
        Dim menuDesc = {"Welcome to Winapp2ool! Check out the ReadMe on GitHub for help!"}

        Dim headerColor = If(UpdatesAvailable, ConsoleColor.Green, ConsoleColor.Cyan)

        Dim menu = MenuSection.CreateCompleteMenu($"{NameOf(Winapp2ool)} v{currentVersion}", menuDesc, headerColor, True)

        menu.AddBlank() _
            .AddOption("WinappDebug", "Scan for and correct style and syntax errors in winapp2.ini") _
            .AddOption("Trim", "Optimize winapp2.ini for your system") _
            .AddOption("Transmute", "Add, replace, or remove entire sections or individual keys from winapp2.ini") _
            .AddOption("Diff", "Generate a context-aware changelog between two winapp2.ini files") _
            .AddOption("CCiniDebug", "Remove stale winapp2.ini configurations from ccleaner.ini") _
            .AddOption("Browser Builder", "Generate winapp2.ini entries for web browsers") _
            .AddOption("Combine", "Join together a collection of ini files into one").AddBlank() _
            .AddOption("Downloader", "Download files from the Winapp2 GitHub") _
            .AddColoredOption("Settings", "Manage Winapp2ool's settings", ConsoleColor.Yellow)

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
            .AddOption("Update Winapp2.ini", "Update your local copy of winapp2.ini", condition:=waUpdateIsAvail) _
            .AddOption("Update & Trim", "Download and trim the latest winapp2.ini", condition:=waUpdateIsAvail) _
            .AddOption("Show winapp2.ini changelog", "See the difference between your local file and the latest", condition:=waUpdateIsAvail) _
            .AddBlank(waUpdateIsAvail) _
            .AddOption("Update Winapp2ool", "Get the latest Winapp2ool.exe", condition:=updateIsAvail AndAlso Not DotNetFrameworkOutOfDate) _
            .AddOption("Go online", "Retry your internet connection", condition:=isOffline)

        menu.Print()

    End Sub

    ''' <summary> 
    ''' Handles the user input for the menu 
    ''' </summary>
    ''' 
    ''' <param name="input">
    ''' The user's input
    ''' </param>
    Public Sub handleToolMainUserInput(input As String)

        Dim modules = New Dictionary(Of String, KeyValuePair(Of Action, Action(Of String))) From {
        {NameOf(WinappDebug), New KeyValuePair(Of Action, Action(Of String))(AddressOf printLintMainMenu, AddressOf handleLintUserInput)},
        {NameOf(Trim), New KeyValuePair(Of Action, Action(Of String))(AddressOf printTrimMenu, AddressOf handleTrimUserInput)},
        {NameOf(Transmute), New KeyValuePair(Of Action, Action(Of String))(AddressOf printTransmuteMainMenu, AddressOf handleTransmuteUserInput)},
        {NameOf(Diff), New KeyValuePair(Of Action, Action(Of String))(AddressOf printDiffMainMenu, AddressOf handleDiffUserInput)},
        {NameOf(CCiniDebug), New KeyValuePair(Of Action, Action(Of String))(AddressOf printCCDBMainMenu, AddressOf handleCCDBMUserInput)},
        {NameOf(BrowserBuilder), New KeyValuePair(Of Action, Action(Of String))(AddressOf printBrowserBuilderMenu, AddressOf handleBrowserBuilderInput)},
        {NameOf(Combine), New KeyValuePair(Of Action, Action(Of String))(AddressOf printCombineMainMenu, AddressOf handleCombineUserInput)},
        {NameOf(Downloader), New KeyValuePair(Of Action, Action(Of String))(AddressOf printDownloadMainMenu, AddressOf handleDownloadUserInput)},
        {"Winapp2ool Settings", New KeyValuePair(Of Action, Action(Of String))(AddressOf printMainToolSettingsMenu, AddressOf handleMainToolSettingsInput)},
        {"Minefield", New KeyValuePair(Of Action, Action(Of String))(AddressOf Minefield.printMenu, AddressOf Minefield.handleUserInput)}
        }

        Dim moduleOpts = {"1", "2", "3", "4", "5", "6", "7", "8", "9", "m"}

        Select Case True

            ' Option Name:                                 Exit
            ' Option States:
            ' Default                                      -> 0 (default)
            Case input = "0"

                exitModule()
                cwl("Exiting...")
                Environment.Exit(0)

            ' Option Section:                             Open Module
            ' Option States:
            ' WinappDebug                                  -> 1 
            ' Trim                                         -> 2 
            ' Transmute                                    -> 3 
            ' Diff                                         -> 4 
            ' CCiniDebug                                   -> 5 
            ' Browser Builder                              -> 6 
            ' Combine                                      -> 7
            ' Downloader                                   -> 8 
            ' Winapp2ool Settings                          -> 9 
            ' Minefield                                    -> m
            ' Notes: Minefield is a hidden option not listed on the menu
            Case moduleOpts.Contains(input)

                ' secret minefield menu
                If input = "m" Then input = "10"

                Dim i = CType(input, Integer) - 1

                ' Downloader requires an internet connection to launch 
                Dim isDownloader = modules.Keys(i) = NameOf(Downloader)
                Dim canLaunch = (isDownloader AndAlso Not denySettingOffline()) OrElse Not isDownloader

                If canLaunch Then initModule(modules.Keys(i), modules.Values(i).Key, modules.Values(i).Value)

            ' Go online
            ' Notes: Only available if offline 
            ' Always appears after the standard suite of options 
            Case input = "10" AndAlso isOffline

                chkOfflineMode()
                setHeaderText("Winapp2ool was unable to establish a network connection. You are still in offline mode.", True, isOffline)

            ' Update (winapp2.ini)
            ' Notes: Only available if an update to winapp2.ini is available
            ' Always appears after the standard suite of options but before the tool updater
            Case input = "10" AndAlso waUpdateIsAvail

                clrConsole()
                cwl("Downloading, this may take a moment...")
                download(New iniFile(Environment.CurrentDirectory, "winapp2.ini"), getWinappLink, False)
                waUpdateIsAvail = False

            ' Update & Trim winapp2.ini
            ' Notes: Only available if an update to winapp2.ini is available
            Case input = "11" AndAlso waUpdateIsAvail

                clrConsole()
                cwl("Downloading & trimming, this may take a moment...")
                remoteTrim(New iniFile(), New iniFile(Environment.CurrentDirectory, "winapp2.ini"), True)
                waUpdateIsAvail = False

            ' Show Update Diff
            ' Notes: Only available if an update to winapp2.ini is available
            Case input = "12" AndAlso waUpdateIsAvail

                clrConsole()
                cwl("Downloading & diffing, this may take a moment...")
                DiffRemoteFile(New iniFile(Environment.CurrentDirectory, "winapp2.ini"))
                setHeaderText("Diff Complete")

            ' Update (winapp2ool.exe)
            ' Notes: Only available if an update is available and the executable can be updated
            ' Always appears as the last option in the menu when available
            Case updateIsAvail AndAlso Not cantDownloadExecutable AndAlso input = computeMenuNumber(10, {waUpdateIsAvail}, {3})

                cwl("Downloading and updating Winapp2ool.exe, this may take a moment...")
                autoUpdate()

            ' Save winapp2ool log 
            ' Notes: hidden option not listed on menu
            Case input = "savelog"

                GlobalLogFile.overwriteToFile(logger.toString)

            ' Print winapp2ool log 
            ' Notes: hidden option not listed on menu
            Case input = "printlog"

                printLog()

            ' Force winapp2ool update 
            ' Notes: hidden option not listed on menu
            Case input = "forceupdate"

                autoUpdate()

            Case Else

                setHeaderText(invInpStr, True)

        End Select

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
