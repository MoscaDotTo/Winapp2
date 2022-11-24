'    Copyright (C) 2018-2022 Hazel Ward
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
''' <summary> Displays the Diff main menu to the user and handles their input </summary>
''' Docs last updated: 2020-08-30
Module diffmainmenu

    ''' <summary> Prints the Diff main menu to the user </summary>
    ''' Docs last updated: 2020-07-23 | Code last updated: 2020-07-19
    Public Sub printDiffMainMenu()
        ' Force disable the Download if it's enabled in offline mode 
        If isOffline AndAlso DownloadDiffFile Then DownloadDiffFile = False
        Console.WindowHeight = If(DiffModuleSettingsChanged, 34, 32)
        printMenuTop({"Observe the differences between two ini files"})
        print(1, "Run (default)", "Run the diff tool", enStrCond:=Not (DiffFile2.Name.Length = 0 AndAlso Not DownloadDiffFile), colorLine:=True)
        print(0, "Select Older/Local File:", leadingBlank:=True)
        print(1, "File Chooser", "Choose a new name or location for your older ini file")
        print(0, "Select Newer/Remote File:", leadingBlank:=True)
        print(5, GetNameFromDL(True), "diffing against the latest winapp2.ini version on GitHub", cond:=Not isOffline, enStrCond:=DownloadDiffFile, leadingBlank:=True)
        print(5, "Remote file trimming", "trimming the remote winapp2.ini before diffing", cond:=Not isOffline AndAlso DownloadDiffFile = True, enStrCond:=TrimRemoteFile, trailingBlank:=True)
        print(1, "File Chooser", "Choose a new name or location for your newer ini file", isOffline OrElse Not DownloadDiffFile, isOffline, True)
        print(0, "Log Settings:")
        print(5, "Toggle Log Saving", "automatic saving of the Diff output", leadingBlank:=True, trailingBlank:=Not SaveDiffLog, enStrCond:=SaveDiffLog)
        print(1, "File Chooser (log)", "Change where Diff saves its log", SaveDiffLog, trailingBlank:=True)
        print(5, "Verbose Mode", "printing full entries in the diff output", enStrCond:=ShowFullEntries, trailingBlank:=True)
        print(0, $"Older file: {replDir(DiffFile1.Path)}")
        print(0, $"Newer file: {If(DiffFile2.Name.Length = 0 AndAlso Not DownloadDiffFile, "Not yet selected", If(DownloadDiffFile, GetNameFromDL(True), replDir(DiffFile2.Path)))}",
                                  closeMenu:=Not SaveDiffLog AndAlso Not DiffModuleSettingsChanged AndAlso MostRecentDiffLog.Length = 0)
        print(0, $"Log   file: {replDir(DiffFile3.Path)}", cond:=SaveDiffLog, closeMenu:=(Not DiffModuleSettingsChanged) AndAlso MostRecentDiffLog.Length = 0)
        print(2, "Diff", cond:=DiffModuleSettingsChanged, closeMenu:=MostRecentDiffLog.Length = 0)
        print(1, "Log Viewer", "Show the most recent Diff log", cond:=Not MostRecentDiffLog.Length = 0, closeMenu:=True, leadingBlank:=True)
    End Sub

    ''' <summary> Handles the user input from the Diff main menu </summary>
    ''' <param name="input"> The user's input </param>
    ''' Docs last updated: 2022-11-21 | Code last updated: 2022-11-21
    Public Sub handleDiffMainMenuUserInput(input As String)

        Select Case True

            ' Option Name:                                 Exit
            ' Option States:
            ' Default                                      -> 0 (default)
            Case input = "0"

                exitModule()

            ' Option Name:                                 Run (default)
            ' Option States:
            ' Default                                      -> 1 (default)
            Case input = "1" OrElse input.Length = 0

                If Not denyActionWithHeader(DiffFile2.Name.Length = 0 AndAlso Not DownloadDiffFile, "Please select a file against which to diff") Then initDiff()

            ' Option Name:                                 File Chooser (Older File)
            ' Option States:
            ' Default                                      -> 2 (default)
            Case input = "2"

                changeFileParams(DiffFile1, DiffModuleSettingsChanged, NameOf(Diff), NameOf(DiffFile1), NameOf(DiffModuleSettingsChanged))

            ' Option Name:                                 Online
            ' Option States:
            ' Offline                                      -> Unavailable (not displayed)
            ' Online                                       -> 3 (default) 
            Case input = "3" AndAlso Not isOffline

                If Not denySettingOffline() Then
                    toggleSettingParam(DownloadDiffFile, "Downloading", DiffModuleSettingsChanged, NameOf(Diff), NameOf(DownloadDiffFile), NameOf(DiffModuleSettingsChanged))
                End If
                DiffFile2.Name = GetNameFromDL(DownloadDiffFile)

            ' Option Name:                                 Remote file trimming
            ' Option States:
            ' Offline                                      -> Unavailable (not displayed) 
            ' Online, not downloading                      -> Unavailable (not displayed)
            ' Online                                       -> 4 (default) 
            Case input = "4" AndAlso DownloadDiffFile

                toggleSettingParam(TrimRemoteFile, "Trimming", DiffModuleSettingsChanged, NameOf(Diff), NameOf(TrimRemoteFile), NameOf(DiffModuleSettingsChanged))

            ' Option Name:                                 File Chooser (Newer File)
            ' Option States:
            ' Downloading                                  -> Unavailable (not displayed) 
            ' Offline (-1)                                 -> 3
            ' Online, not downloading                      -> 4 (default)
            Case input = computeMenuNumber(4, {isOffline}, {-1})

                changeFileParams(DiffFile2, DiffModuleSettingsChanged, NameOf(Diff), NameOf(DiffFile2), NameOf(DiffModuleSettingsChanged))

            ' Option Name:                                 Toggle Log Saving
            ' Option States:
            ' Offline (-1)                                 -> 4 
            ' Online                                       -> 5 (default)
            Case input = computeMenuNumber(5, {isOffline}, {-1})

                toggleSettingParam(SaveDiffLog, "Log Saving", DiffModuleSettingsChanged, NameOf(Diff), NameOf(SaveDiffLog), NameOf(DiffModuleSettingsChanged))

            ' Option Name:                                 File Chooser (log)
            ' Option States:
            ' Not Saving Log                               -> Unavailable (not displayed) 
            ' Offline (-1)                                 -> 5 
            ' Online                                       -> 6 (default)
            Case SaveDiffLog AndAlso input = computeMenuNumber(6, {isOffline}, {-1})

                changeFileParams(DiffFile3, DiffModuleSettingsChanged, NameOf(Diff), NameOf(DiffFile3), NameOf(DiffModuleSettingsChanged))

            ' Option Name:                                 Verbose Mode
            ' Option States:
            ' Offline (-1), not saving log                 -> 5
            ' Online, not saving log                       -> 6 (default)
            ' Offline (-1), saving log (+1)                -> 6
            Case input = computeMenuNumber(6, {isOffline, SaveDiffLog}, {-1, 1})

                toggleSettingParam(ShowFullEntries, "Verbose Mode", DiffModuleSettingsChanged, NameOf(Diff), NameOf(ShowFullEntries), NameOf(DiffModuleSettingsChanged))

            ' Option Name:                                 Reset Settings
            ' Option States:
            ' DiffModuleSettingsChanged = False            -> Unavailable (not displayed) 
            ' Offline (-1), not saving log                 -> 6
            ' Online, Not saving log                       -> 7 (default)
            ' Offline (-1), saving log (+1)                -> 7
            ' Online, saving log (+1)                      -> 8 

            Case DiffModuleSettingsChanged AndAlso input = computeMenuNumber(7, {isOffline, SaveDiffLog}, {-1, 1})
                resetModuleSettings(NameOf(Diff), AddressOf initDefaultDiffSettings)

            ' Option Name:                                 Log Viewer
            ' Option States
            ' Offline (-1), not saving log                 -> 7
            ' Online, not saving log                       -> 8 (default)
            ' Offline (-1), saving log (+1)                -> 8
            ' Online, saving log (+1)                      -> 9
            Case Not MostRecentDiffLog.Length = 0 AndAlso input = computeMenuNumber(8, {isOffline, SaveDiffLog}, {-1, 1})

                MostRecentDiffLog = getLogSliceFromGlobal("Beginning diff", "Diff complete")
                printSlice(MostRecentDiffLog)

            Case Else

                setHeaderText(invInpStr, True)

        End Select

    End Sub

End Module