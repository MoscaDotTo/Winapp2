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
''' Displays the main winapp2ool menu to the user and handles input from that menu 
''' </summary>
''' 
''' Docs last updated: 2023-07-19 | Code last updated: 2023-07-19
Module maintoolmenu

    ''' <summary> 
    ''' Prints the main winapp2ool menu to the user 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2023-07-19 | Code last updated: 2023-07-19
    Public Sub printToolMainMenu()

        checkUpdates(Not isOffline AndAlso Not checkedForUpdates)
        Dim UpdatesAvailable = Not isOffline AndAlso (waUpdateIsAvail OrElse updateIsAvail)

        printMenuTop(Array.Empty(Of String)(), False)

        ' User warnings for limited feature availability
        If isOffline OrElse DotNetFrameworkOutOfDate OrElse cantDownloadExecutable Then

            Dim statusSection As New MenuSection("")

            statusSection.AddLine("Winapp2ool is currently in offline mode", centered:=True, condition:=isOffline) _
                         .AddLine("Your .NET Framework is out of date", centered:=True, condition:=DotNetFrameworkOutOfDate) _
                         .AddLine("Winapp2ool is currently running from the temporary folder, some functions may be impacted", centered:=True, condition:=cantDownloadExecutable) _
                         .Print(withDivider:=False)

            PrintBlank()

        End If

        ' Update notifications
        printUpdNotif(waUpdateIsAvail, "winapp2.ini", localWa2Ver, latestWa2Ver)
        printUpdNotif(updateIsAvail, "Winapp2ool", currentVersion, latestVersion)

        ' Core tools section
        Dim coreToolsSection As New MenuSection("")

        coreToolsSection.AddOption("Exit", "Exit the application") _
                        .AddOption("WinappDebug", "Check for and correct errors in winapp2.ini") _
                        .AddOption("Trim", "Debloat winapp2.ini for your system") _
                        .AddOption("Transmute", "Add, replace, or remove entire sections or individual keys from winapp2.ini") _
                        .AddOption("Diff", "Observe the changes between two winapp2.ini files") _
                        .AddOption("CCiniDebug", "Sort and trim ccleaner.ini") _
                        .AddOption("Browser Builder", "Generate winapp2.ini entries for web browsers") _
                        .AddBlank() _
                        .AddOption("Downloader", "Download files from the Winapp2 GitHub") _
                        .AddColoredOption("Settings", "Manage Winapp2ool's settings", ConsoleColor.Yellow) _
                        .Print(withDivider:=False)

        If waUpdateIsAvail AndAlso Not isOffline Then
            Dim updateSection As New MenuSection("")
            updateSection.AddBlank() _
                         .AddOption("Update", "Update your local copy of winapp2.ini") _
                         .AddOption("Update & Trim", "Download and trim the latest winapp2.ini") _
                         .AddOption("Show update diff", "See the difference between your local file and the latest") _
                         .Print(withDivider:=False)
        End If

        PrintOption("Update", "Get the latest Winapp2ool.exe", condition:=updateIsAvail AndAlso Not DotNetFrameworkOutOfDate)
        PrintOption("Go online", "Retry your internet connection", condition:=isOffline)

        EndMenu()

    End Sub

    ''' <summary> 
    ''' Handles the user input for the menu 
    ''' </summary>
    ''' 
    ''' <param name="input">
    ''' The <c> String </c> containing the user's input 
    ''' </param>
    ''' 
    ''' Docs last updated: 2020-07-14 | Code last updated: 2020-07-14
    Public Sub handleToolMainUserInput(input As String)

        Dim modules = New Dictionary(Of String, KeyValuePair(Of Action, Action(Of String))) From {
        {NameOf(WinappDebug), New KeyValuePair(Of Action, Action(Of String))(AddressOf printLintMainMenu, AddressOf handleLintUserInput)},
        {NameOf(Trim), New KeyValuePair(Of Action, Action(Of String))(AddressOf printTrimMenu, AddressOf handleTrimUserInput)},
        {NameOf(Transmute), New KeyValuePair(Of Action, Action(Of String))(AddressOf printTransmuteMainMenu, AddressOf handleTransmuteUserInput)},
        {NameOf(Diff), New KeyValuePair(Of Action, Action(Of String))(AddressOf printDiffMainMenu, AddressOf handleDiffUserInput)},
        {NameOf(CCiniDebug), New KeyValuePair(Of Action, Action(Of String))(AddressOf printCCDBMainMenu, AddressOf handleCCDBMUserInput)},
        {NameOf(BrowserBuilder), New KeyValuePair(Of Action, Action(Of String))(AddressOf printBrowserBuilderMenu, AddressOf handleBrowserBuilderInput)},
        {NameOf(Downloader), New KeyValuePair(Of Action, Action(Of String))(AddressOf printDownloadMainMenu, AddressOf handleDownloadUserInput)},
        {"Winapp2ool Settings", New KeyValuePair(Of Action, Action(Of String))(AddressOf printMainToolSettingsMenu, AddressOf handleMainToolSettingsInput)},
        {"Minefield", New KeyValuePair(Of Action, Action(Of String))(AddressOf Minefield.printMenu, AddressOf Minefield.handleUserInput)}
        }

        Dim moduleOpts = {"1", "2", "3", "4", "5", "6", "7", "8", "m"}

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
            ' Downloader                                   -> 7 
            ' Winapp2ool Settings                          -> 8 
            ' Minefield                                    -> m 
            Case moduleOpts.Contains(input)

                ' secret minefield menu
                If input = "m" Then input = "9"

                Dim i = CType(input, Integer) - 1

                ' Downloader requires an internet connection to launch 
                Dim isDownloader = modules.Keys(i) = NameOf(Downloader)
                Dim canLaunch = isDownloader AndAlso Not denySettingOffline() OrElse Not isDownloader

                If canLaunch Then initModule(modules.Keys(i), modules.Values(i).Key, modules.Values(i).Value)

            ' Option Name:                                 Go online
            ' Option States:
            ' Default                                      -> Unavailable (default)
            ' Offline:                                    -> 9 
            Case input = "9" AndAlso isOffline

                chkOfflineMode()
                setHeaderText("Winapp2ool was unable to establish a network connection. You are still in offline mode.", True, isOffline)

            ' Option Name:                                 Update
            ' Option States:
            ' Default                                      -> Unavailable (default)
            ' Winapp2.ini Update available:                -> 9
            Case input = "9" AndAlso waUpdateIsAvail

                clrConsole()
                cwl("Downloading, this may take a moment...")
                download(New iniFile(Environment.CurrentDirectory, "winapp2.ini"), winapp2link, False)
                waUpdateIsAvail = False

            ' Option Name:                                 Update & Trim
            ' Option States:
            ' Default                                      -> Unavailable (default)
            ' Winapp2.ini Update available:                -> 10
            Case input = "10" AndAlso waUpdateIsAvail

                clrConsole()
                cwl("Downloading & trimming, this may take a moment...")
                remoteTrim(New iniFile(), New iniFile(Environment.CurrentDirectory, "winapp2.ini"), True)
                waUpdateIsAvail = False

            ' Option Name:                                 Show Update Diff
            ' Option States:
            ' Default                                      -> Unavailable (default)
            ' Winapp2.ini Update available:                -> 11
            Case input = "11" AndAlso waUpdateIsAvail

                clrConsole()
                cwl("Downloading & diffing, this may take a moment...")
                DiffRemoteFile(New iniFile(Environment.CurrentDirectory, "winapp2.ini"))
                setHeaderText("Diff Complete")

            ' Option Name:                                 Update (winapp2ool.exe)
            ' Option States:
            ' Default                                      -> Unavailable (default)
            ' Winapp2ool and Winapp2.ini Update available: -> 12
            ' Winapp2ool Update available:                 -> 9
            Case updateIsAvail AndAlso Not (DotNetFrameworkOutOfDate OrElse cantDownloadExecutable) AndAlso (input = "12" AndAlso waUpdateIsAvail) OrElse (input = "9" AndAlso (Not waUpdateIsAvail))

                cwl("Downloading and updating Winapp2ool.exe, this may take a moment...")
                autoUpdate()

            ' Option Name:                                 Minefield
            ' Option States:
            ' Default                                      -> m (default)
            Case input = "m"

                initModule("Minefield", AddressOf Minefield.printMenu, AddressOf Minefield.handleUserInput)

            ' Option Name:                                 Save winapp2ool log 
            ' Option States:
            ' Default                                      -> savelog (default)
            Case input = "savelog"

                GlobalLogFile.overwriteToFile(logger.toString)

            ' Option Name:                                 Print winapp2ool log 
            ' Option States:
            ' Default                                      -> printlog (default)
            Case input = "printlog"

                printLog()

            ' Option Name:                                 Force winapp2ool update 
            ' Option States:
            ' Default                                      -> forceupdate
            Case input = "forceupdate"

                autoUpdate()

            Case Else

                setHeaderText(invInpStr, True)

        End Select

    End Sub

End Module