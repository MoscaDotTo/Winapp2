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
Module Winapp2ool
    ''' <summary> Indicates that winapp2ool is in "Non-CCleaner" mode and should collect the appropriate ini from GitHub </summary>
    ''' Docs last updated: Before 2020-04-18 | Code last updated: Before 2020-04-18
    Public Property RemoteWinappIsNonCC As Boolean = False
    ''' <summary> Indicates that the .NET Framework installed on the current machine is below the targeted version (.NET Framework 4.5) </summary>
    ''' Docs last updated: Before 2020-04-18 | Code last updated: Before 2020-04-18
    Public Property DotNetFrameworkOutOfDate As Boolean = False
    ''' <summary> Indicates that winapp2ool currently has access to the internet </summary>
    ''' Docs last updated: Before 2020-04-18 | Code last updated: Before 2020-04-18
    Public Property isOffline As Boolean = False
    ''' <summary> Indicates that this build is beta and should check the beta branch link for updates </summary>
    ''' Docs last updated: Before 2020-04-18 | Code last updated: Before 2020-04-18
    Public Property isBeta As Boolean = False
    ''' <summary> Inidcates that we're unable to download the executable </summary>
    ''' Docs last updated: Before 2020-04-18 | Code last updated: Before 2020-04-18
    Public Property cantDownloadExecutable As Boolean = False
    ''' <summary> Indicates that winapp2ool.exe has already been downloaded during this session and prevents us from redownloading it </summary>
    ''' Docs last updated: Before 2020-04-18 | Code last updated: Before 2020-04-18
    Public Property alreadyDownloadedExecutable As Boolean = False
    ''' <summary> Indicates that the module's settings have been changed </summary>
    ''' Docs last updated: Before 2020-04-18 | Code last updated: Before 2020-04-18
    Public Property toolSettingsHaveChanged As Boolean = False

    ''' <summary> Initalizes the default state of the winapp2ool module settings </summary>
    ''' Docs last updated: Before 2020-04-18 | Code last updated: Before 2020-04-18
    Private Sub initDefaultSettings()
        GlobalLogFile.resetParams()
        RemoteWinappIsNonCC = False
        saveSettingsToDisk = False
        readSettingsFromDisk = False
        toolSettingsHaveChanged = False
        restoreDefaultSettings(NameOf(Winapp2ool), AddressOf createToolSettingsSection)
    End Sub

    ''' <summary> Loads values from disk into memory for the winapp2ool module settings </summary>
    ''' Docs last updated: Before 2020-04-18 | Code last updated: Before 2020-04-18
    Public Sub getSeralizedToolSettings()
        For Each kvp In settingsDict(NameOf(Winapp2ool))
            Select Case kvp.Key
                Case NameOf(isBeta)
                    isBeta = CBool(kvp.Value)
                Case NameOf(readSettingsFromDisk)
                    readSettingsFromDisk = CBool(kvp.Value)
                Case NameOf(saveSettingsToDisk)
                    saveSettingsToDisk = CBool(kvp.Value)
                Case NameOf(RemoteWinappIsNonCC)
                    RemoteWinappIsNonCC = CBool(kvp.Value)
                Case NameOf(toolSettingsHaveChanged)
                    toolSettingsHaveChanged = CBool(kvp.Value)
                Case NameOf(GlobalLogFile) & "_Dir"
                    GlobalLogFile.Dir = kvp.Value
                Case NameOf(GlobalLogFile) & "_Name"
                    GlobalLogFile.Name = kvp.Value
            End Select
        Next
    End Sub

    '''<summary> Adds the current (typically default) state of the module's settings into the disk-writable settings representation </summary>
    ''' Docs last updated: Before 2020-04-18 | Code last updated: Before 2020-04-18
    Public Sub createToolSettingsSection()
        Dim compCult = System.Globalization.CultureInfo.InvariantCulture
        createModuleSettingsSection(NameOf(Winapp2ool), {
                                        getSettingIniKey(NameOf(Winapp2ool), NameOf(isBeta), isBeta.ToString(compCult)),
                                        getSettingIniKey(NameOf(Winapp2ool), NameOf(saveSettingsToDisk), saveSettingsToDisk.ToString(compCult)),
                                        getSettingIniKey(NameOf(Winapp2ool), NameOf(readSettingsFromDisk), readSettingsFromDisk.ToString(compCult)),
                                        getSettingIniKey(NameOf(Winapp2ool), NameOf(RemoteWinappIsNonCC), RemoteWinappIsNonCC.ToString(compCult)),
                                        getSettingIniKey(NameOf(Winapp2ool), NameOf(toolSettingsHaveChanged), toolSettingsHaveChanged.ToString(compCult)),
                                        getSettingIniKey(NameOf(Winapp2ool), NameOf(GlobalLogFile) & "_Dir", GlobalLogFile.Dir),
                                        getSettingIniKey(NameOf(Winapp2ool), NameOf(GlobalLogFile) & "_Name", GlobalLogFile.Name)
                                    })
    End Sub

    ''' <summary> Prints the main winapp2ool menu to the user </summary>
    ''' Docs last updated: Before 2020-04-18 | Code last updated: Before 2020-04-18
    Private Sub printMenu()
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

    ''' <summary> Initilizes winapp2ool by checking for internet connectivity, serializing any settings from disk,
    ''' and processing any commandline arguments  before loading the tool's main menu </summary>
    ''' Docs last updated: 2020-04-18 | Code last updated: Before 2020-04-18
    Public Sub main()
        gLog($"Starting application")
        ' winapp2ool requires internet access for some functions
        chkOfflineMode()
        currentVersion = FileVersionInfo.GetVersionInfo(Environment.GetCommandLineArgs(0)).FileVersion
        ' Set the console stage 
        Console.Title = $"Winapp2ool v{currentVersion}"
        Console.WindowWidth = 126
        ' winapp2ool requires .NET 4.6 or higher for full functionality, all versions of which report the following version
        If Not Environment.Version.ToString = "4.0.30319.42000" Then DotNetFrameworkOutOfDate = True
        gLog($".NET Framework is out of date. Found {Environment.Version}", DotNetFrameworkOutOfDate)
        ' Make sure we're not operating in the temporary directory 
        cantDownloadExecutable = Environment.CurrentDirectory.Equals(Environment.GetEnvironmentVariable("temp"), StringComparison.InvariantCultureIgnoreCase)
        loadSettings()
        processCommandLineArgs()
        If SuppressOutput Then Environment.Exit(1)
        initModule($"Winapp2ool v{currentVersion} - A multitool for winapp2.ini", AddressOf printMenu, AddressOf handleUserInput)
    End Sub

    ''' <summary> Handles the user input for the menu </summary>
    ''' <param name="input"> The <c> String </c> containing the user's input </param>
    ''' Docs last updated: 2020-04-18 | Code last updated: 2020-04-18
    Private Sub handleUserInput(input As String)
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
                initModule(NameOf(Diff), AddressOf Diff.printMenu, AddressOf Diff.handleUserInput)
            Case input = "5"
                initModule(NameOf(CCiniDebug), AddressOf CCiniDebug.printMenu, AddressOf CCiniDebug.handleUserInput)
            Case input = "6"
                If Not denySettingOffline() Then initModule("Downloader", AddressOf Downloader.printMenu, AddressOf Downloader.handleUserInput)
            Case input = "7"
                initModule("Winapp2ool Settings", AddressOf printToolSettingsMenu, AddressOf handleToolSettingsInput)
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
            Case (input = "11" And (updateIsAvail And waUpdateIsAvail)) Or (input = "7" And (Not waUpdateIsAvail And updateIsAvail)) And Not (DotNetFrameworkOutOfDate Or cantDownloadExecutable)
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

    ''' <summary> Checks the version of Windows on the current system and returns it as a Double </summary>
    ''' <returns> The Windows version running on the machine, <c> 0.0 </c> if the windows version cannot be determined </returns>
    ''' Docs last updated: 2020-04-18 | Code last updated: Before 2020-04-18
    Public Function getWinVer() As Double
        gLog("Checking Windows version")
        Dim osVersion = System.Environment.OSVersion.ToString().Replace("Microsoft Windows NT ", "")
        Dim ver = osVersion.Split(CChar("."))
        Dim out = Val($"{ver(0)}.{ver(1)}")
        gLog($"Found Windows {out}")
        Return out
    End Function

    ''' <summary> Returns the first portion of a registry or filepath parameterization </summary>
    ''' <param name="val"> A Windows filesystem or registry path from which the root should be returned </param>
    ''' <returns> The root directory given by <paramref name="val"/> </returns>
    ''' Docs last updated: 2020-04-18 | Code last updated: Before 2020-04-18
    Public Function getFirstDir(val As String) As String
        Return val.Split(CChar("\"))(0)
    End Function

    ''' <summary> Ensures that an <c> iniFile </c> has content and informs the user if it does not. Returns <c> False </c> if there are no sections </summary>
    ''' <param name="iFile"> An <c> iniFile </c> to be checked for content </param>
    ''' Docs last updated: 2020-04-18 | Code last updated: Before 2020-04-18
    Public Function enforceFileHasContent(iFile As iniFile) As Boolean
        iFile.validate()
        If iFile.Sections.Count = 0 Then
            setHeaderText($"{iFile.Name} was empty or not found", True)
            gLog($"{iFile.Name} was empty or not found", indent:=True)
            Return False
        End If
        Return True
    End Function

    ''' <summary> Prints the winapp2ool settings menu to the user </summary>
    ''' Docs last updated: 2020-04-18 | Code last updated: 2020-04-18
    Private Sub printToolSettingsMenu()
        printMenuTop({"Change some high level settings, including saving & reading settings from disk"})
        print(5, "Toggle Saving Settings", $"saving a copy of winapp2ool's settings to the disk", enStrCond:=saveSettingsToDisk, leadingBlank:=True)
        print(5, "Toggle Reading Settings", $"overriding winapp2ool's default settings with those found in winapp2ool.ini", enStrCond:=readSettingsFromDisk, trailingBlank:=True)
        print(5, "Toggle Non-CCleaner Mode", $"using the Non-CCleaner version of winapp2.ini by default", enStrCond:=RemoteWinappIsNonCC, trailingBlank:=True)
        print(1, "View Log", "Print winapp2ool's internal log")
        print(1, "File Chooser (log)", "Change the filename or path to which the winapp2ool log should be saved")
        print(1, "Save Log", "Save winapp2ool's internal log to the disk")
        print(0, $"Current log file target: {replDir(GlobalLogFile.Path)}", leadingBlank:=True, trailingBlank:=True)
        print(1, "Visit GitHub", "Open the winapp2.ini/winapp2ool GitHub in your default web browser", trailingBlank:=True)
        print(5, "Toggle Beta Participation", $"participating in the 'beta' builds of winapp2ool (requires a restart)", enStrCond:=isBeta, closeMenu:=Not toolSettingsHaveChanged)
        print(2, NameOf(Winapp2ool), cond:=toolSettingsHaveChanged, closeMenu:=True)
    End Sub

    ''' <summary> Handles the user input for the winapp2ool settings menu </summary>
    ''' <param name="input"> The user's input </param>
    ''' Docs last updated: 2020-04-18 | Code last updated: 2020-04-18
    Private Sub handleToolSettingsInput(input As String)
        Select Case True
            Case input = "0"
                exitModule()
            Case input = "1"
                toggleSettingParam(saveSettingsToDisk, "Saving settings to disk", toolSettingsHaveChanged, NameOf(Winapp2ool), NameOf(saveSettingsToDisk), NameOf(toolSettingsHaveChanged))
            Case input = "2"
                toggleSettingParam(readSettingsFromDisk, "Reading settings from disk", toolSettingsHaveChanged, NameOf(Winapp2ool), NameOf(readSettingsFromDisk), NameOf(toolSettingsHaveChanged))
            Case input = "3"
                toggleSettingParam(RemoteWinappIsNonCC, "Non-CCleaner mode", toolSettingsHaveChanged, NameOf(Winapp2ool), NameOf(RemoteWinappIsNonCC), NameOf(toolSettingsHaveChanged))
            Case input = "4"
                printLog()
            Case input = "5"
                changeFileParams(GlobalLogFile, toolSettingsHaveChanged, NameOf(Winapp2ool), NameOf(GlobalLogFile), NameOf(toolSettingsHaveChanged))
            Case input = "6"
                GlobalLogFile.overwriteToFile(logger.toString)
            Case input = "7"
                Process.Start(gitLink)
            Case input = "8"
                If Not denyActionWithHeader(DotNetFrameworkOutOfDate, "Winapp2ool beta requires .NET 4.6 or higher") Then
                    toggleSettingParam(isBeta, "Beta Participation", toolSettingsHaveChanged, NameOf(Winapp2ool), NameOf(isBeta), NameOf(toolSettingsHaveChanged))
                    autoUpdate()
                End If
            Case input = "9" And toolSettingsHaveChanged
                initDefaultSettings()
            Case Else
                setHeaderText(invInpStr, True)
        End Select
    End Sub
End Module