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
''' Displays the Diff main menu to the user and handles their input accordingly
''' </summary>
Module diffmainmenu

    ''' <summary>
    ''' Builds the Diff main menu with all options and their dispatch handlers registered inline.
    ''' Called by both <c> printDiffMainMenu </c> (to render) and <c> handleDiffUserInput </c>
    ''' (to dispatch), so the displayed option numbers and the dispatch table are always in sync.
    ''' </summary>
    Private Function buildDiffMenu() As MenuSection

        Dim newFileHasName = Not (DiffFile2.Name.Length = 0 AndAlso Not DownloadDiffFile)
        Dim newerFileText = If(Not newFileHasName, "Not yet selected", If(DownloadDiffFile, GetNameFromDL(True), replDir(DiffFile2.Path())))
        Dim olderFileText = If(DownloadDiffFile, "Local", "Older")
        Dim menuDesc = {"Observes the differences between two ini files"}

        If isOffline Then DownloadDiffFile = False

        Console.WindowHeight = 36

        Return MenuSection.CreateCompleteMenu(NameOf(Diff), menuDesc, ConsoleColor.DarkCyan) _
            .AddDispatchedColoredOption("Run (default)", "Perform a Diff operation", GetRedGreen(Not newFileHasName),
                Sub() If Not denyActionWithHeader(DiffFile2.Name.Length = 0 AndAlso Not DownloadDiffFile, "Please select a file against which to diff") Then ConductDiff()) _
            .AddBlank() _
            .AddDispatchedToggle("diffing against GitHub", "performing the Diff operation against the newest file on GitHub", DownloadDiffFile,
                Sub() toggleModuleSetting("Remote Diffing", NameOf(Diff), GetType(diffsettings), NameOf(DownloadDiffFile), NameOf(DiffModuleSettingsChanged)),
                Not isOffline) _
            .AddDispatchedToggle("remote file trim", "trimming the remote file before diffing for a more bespoke diff", TrimRemoteFile,
                Sub() toggleModuleSetting("Remote file trimming", NameOf(Diff), GetType(diffsettings), NameOf(TrimRemoteFile), NameOf(DiffModuleSettingsChanged)),
                DownloadDiffFile) _
            .AddDispatchedToggle("log saving", "saving of the diff output to disk", SaveDiffLog,
                Sub() toggleModuleSetting("Log saving", NameOf(Diff), GetType(diffsettings), NameOf(SaveDiffLog), NameOf(DiffModuleSettingsChanged))) _
            .AddDispatchedToggle("verbose mode", "printing full entries in the diff output", ShowFullEntries,
                Sub() toggleModuleSetting("Verbose mode", NameOf(Diff), GetType(diffsettings), NameOf(ShowFullEntries), NameOf(DiffModuleSettingsChanged))) _
            .AddBlank() _
            .AddDispatchedOption("Choose older/local file", "Select the older version of the file against which to diff",
                Sub() changeFile2Params(DiffFile1, DiffModuleSettingsChanged, NameOf(Diff), NameOf(DiffFile1), NameOf(DiffModuleSettingsChanged))) _
            .AddDispatchedOption("Choose newer file", "Select the newer version of the file to see what has changed",
                Sub() changeFile2Params(DiffFile2, DiffModuleSettingsChanged, NameOf(Diff), NameOf(DiffFile2), NameOf(DiffModuleSettingsChanged)),
                Not DownloadDiffFile) _
            .AddDispatchedOption("Choose save target", "Select where to save the diff output",
                Sub() changeFile2Params(DiffFile3, DiffModuleSettingsChanged, NameOf(Diff), NameOf(DiffFile3), NameOf(DiffModuleSettingsChanged)),
                SaveDiffLog) _
            .AddBlank() _
            .AddBlank(Not MostRecentDiffLog = "") _
            .AddDispatchedColoredOption("Log Viewer", "View the most recent Diff log", ConsoleColor.Yellow,
                Sub()
                    MostRecentDiffLog = getLogSliceFromGlobal(DiffLogStartPhrase, DiffLogEndPhrase)
                    printSlice(MostRecentDiffLog)
                End Sub,
                Not MostRecentDiffLog = "") _
            .AddBlank() _
            .AddColoredFileInfo($"{olderFileText} file:", DiffFile1.Path(), ConsoleColor.DarkYellow) _
            .AddColoredLine($"Newer file: {newerFileText}", ConsoleColor.Magenta) _
            .AddColoredFileInfo("Save target: ", DiffFile3.Path(), ConsoleColor.Cyan, condition:=SaveDiffLog) _
            .AddBlank(DiffModuleSettingsChanged) _
            .AddDispatchedResetOpt(NameOf(Diff), DiffModuleSettingsChanged,
                Sub() resetModuleSettings(NameOf(Diff), AddressOf InitDefaultDiffSettings))

    End Function

    ''' <summary>
    ''' Prints the Diff main menu to the user
    ''' </summary>
    Public Sub printDiffMainMenu()

        buildDiffMenu().Print()

    End Sub

    ''' <summary>
    ''' Handles the user input from the Diff main menu
    ''' </summary>
    '''
    ''' <param name="input">
    ''' The user's input
    ''' </param>
    Public Sub handleDiffUserInput(input As String)

        Dim intInput As Integer

        If Not Integer.TryParse(input, intInput) Then

            ' Allow an empty input to trigger a run if the run conditions are otherwise satisfied
            If input.Length = 0 AndAlso Not denyActionWithHeader(DiffFile2.Name.Length = 0 AndAlso Not DownloadDiffFile, "Please select a file against which to diff") Then ConductDiff() : Return

            setNextMenuHeaderText(invInpStr, printColor:=ConsoleColor.Red)
            Return

        End If

        If intInput = 0 Then exitModule() : Return

        If Not buildDiffMenu().Dispatch(intInput) Then setNextMenuHeaderText(invInpStr, printColor:=ConsoleColor.Red)

    End Sub

End Module
