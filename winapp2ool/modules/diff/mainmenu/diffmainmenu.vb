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
''' 
''' Docs last updated: 2025-08-21
Module diffmainmenu

    ''' <summary> 
    ''' Prints the Diff main menu to the user 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2023-07-19 | Code last updated: 2025-08-19
    Public Sub printDiffMainMenu()

        Dim newFileHasName = Not (DiffFile2.Name.Length = 0 AndAlso Not DownloadDiffFile)
        Dim newerFileText = If(Not newFileHasName, "Not yet selected", If(DownloadDiffFile, GetNameFromDL(True), replDir(DiffFile2.Path)))
        Dim olderFileText = If(DownloadDiffFile, "Local", "Older")
        Dim menuDesc = {"Observes the differences between two ini files"}

        Dim noLog = MostRecentDiffLog.Length = 0
        Dim SettingsUnchangedAndNoLog = Not DiffModuleSettingsChanged AndAlso noLog

        If isOffline Then DownloadDiffFile = False

        Console.WindowHeight = 36

        Dim menu As New MenuSection
        menu = MenuSection.CreateCompleteMenu(NameOf(Diff), menuDesc, ConsoleColor.DarkCyan)
        menu.AddColoredOption("Run (default)", "Perform a Diff operation", GetRedGreen(Not newFileHasName)).AddBlank() _
            .AddToggle("diffing against GitHub", "performing the Diff operation against the newest file on GitHub", DownloadDiffFile, Not isOffline) _
            .AddToggle("remote file trim", "trimming the remote file before diffing for a more bespoke diff", TrimRemoteFile, DownloadDiffFile) _
            .AddToggle("log saving", "saving of the diff output to disk", SaveDiffLog) _
            .AddToggle("verbose mode", "printing full entries in the diff output", ShowFullEntries) _
            .AddBlank() _
            .AddOption("Choose older/local file", "Select the older version of the file against which to diff") _
            .AddOption("Choose newer file", "Select the newer version of the file to see what has changed", Not DownloadDiffFile) _
            .AddOption("Choose save target", "Select where to save the diff output", SaveDiffLog).AddBlank() _
            .AddBlank(Not MostRecentDiffLog = "") _
            .AddColoredOption("Log Viewer", "View the most recent Diff log", ConsoleColor.Yellow, Not MostRecentDiffLog = "").AddBlank() _
            .AddColoredFileInfo($"{olderFileText} file:", DiffFile1.Path, ConsoleColor.DarkYellow) _
            .AddColoredLine($"Newer file: {newerFileText}", ConsoleColor.Magenta) _
            .AddColoredFileInfo($"Save target: ", DiffFile3.Path, ConsoleColor.Cyan, condition:=SaveDiffLog) _
            .AddBlank(DiffModuleSettingsChanged) _
            .AddResetOpt(NameOf(Diff), DiffModuleSettingsChanged)

        menu.Print()

    End Sub

    ''' <summary> 
    ''' Handles the user input from the Diff main menu 
    ''' </summary>
    ''' 
    ''' <param name="input"> 
    ''' The user's input 
    ''' </param>
    ''' 
    ''' Docs last updated: 2022-11-21 | Code last updated: 2025-08-21
    Public Sub handleDiffUserInput(input As String)

        Dim newFileEmpty = DiffFile2.Name.Length = 0 AndAlso Not DownloadDiffFile
        Dim emptyFileErr = "Please select a file against which to diff"

        Dim intInput As Integer
        If Not Integer.TryParse(input, intInput) Then

            ' Allow an empty input to trigger a run if the run conditions are otherwise satisfied
            If input.Length = 0 AndAlso Not denyActionWithHeader(newFileEmpty, emptyFileErr) Then ConductDiff() : Return

            setNextMenuHeaderText(invInpStr, printColor:=ConsoleColor.Red)

            Return

        End If

        Dim toggleOpts = getToggleOpts()
        Dim fileOpts = getFileOpts()

        Dim staticOpts = 2
        Dim minToggleNum = staticOpts
        Dim maxToggleNum = minToggleNum + toggleOpts.Count - 1
        Dim minFileNum = maxToggleNum + 1
        Dim maxFileNum = minFileNum + fileOpts.Count - 1
        Dim logNum = If(MostRecentDiffLog.Length > 0, maxFileNum + 1, maxFileNum)
        Dim resetNum = logNum + 1

        Select Case True

            ' Exit 
            Case intInput = 0

                exitModule()

            ' Run (default)
            Case intInput = 1

                If Not denyActionWithHeader(newFileEmpty, emptyFileErr) Then ConductDiff()

            ' Toggles
            Case intInput >= minToggleNum AndAlso intInput <= maxToggleNum

                Dim i = intInput - staticOpts

                Dim toggleMenuText = toggleOpts.Keys(i)
                Dim toggleName = toggleOpts(toggleMenuText)

                toggleModuleSetting(toggleMenuText, NameOf(Diff), GetType(diffsettings),
                                    toggleName, NameOf(DiffModuleSettingsChanged))

            ' File Selectors
            Case intInput >= minFileNum AndAlso intInput <= maxFileNum

                Dim i = intInput - staticOpts - toggleOpts.Count

                Dim fileName = fileOpts.Keys(i)
                Dim fileObj = fileOpts(fileName)

                changeFileParams(fileObj, DiffModuleSettingsChanged, NameOf(Diff), fileName, NameOf(DiffModuleSettingsChanged))

            ' Log Viewer
            Case MostRecentDiffLog.Length > 0 AndAlso intInput = logNum

                MostRecentDiffLog = getLogSliceFromGlobal(DiffLogStartPhrase, DiffLogEndPhrase)
                printSlice(MostRecentDiffLog)

            ' Reset settings
            Case DiffModuleSettingsChanged AndAlso intInput = resetNum

                resetModuleSettings(NameOf(Diff), AddressOf InitDefaultDiffSettings)

            Case Else

                setNextMenuHeaderText(invInpStr, printColor:=ConsoleColor.Red)

        End Select

    End Sub

    ''' <summary>
    ''' Determines the current set of toggles displayed on the menu and returns a Dictionary 
    ''' of those options and their respective toggle names <br />
    ''' <br />
    ''' The set of possible toggles includes:
    ''' <list type="bullet">
    '''     
    '''     <item>
    '''     Remote diffinng (unavailable when offline)
    '''     </item>
    '''     
    '''     <item>
    '''     Remote file trimming (unavailable when not downloading)
    '''     </item>
    '''     
    '''     <item>
    '''     Excludes
    '''     </item>
    '''     
    ''' </list>
    '''  
    ''' </summary>
    ''' 
    ''' <returns>
    ''' The set of available toggles for the Trim module, with their names on the 
    ''' menu as keys and the respective property names as values
    ''' </returns>
    ''' 
    ''' Docs last updated: 2025-08-12 | Code last updated: 2025-08-12
    Private Function getToggleOpts() As Dictionary(Of String, String)

        Dim baseToggles As New Dictionary(Of String, String)

        If Not isOffline Then baseToggles.Add("Remote Diffing", NameOf(DownloadDiffFile))
        If DownloadDiffFile Then baseToggles.Add("Remote file trimming", NameOf(TrimRemoteFile))

        baseToggles.Add("Log saving", NameOf(SaveDiffLog))
        baseToggles.Add("Verbose mode", NameOf(ShowFullEntries))

        Return baseToggles

    End Function

    ''' <summary>
    ''' Determines the current set of file selectors displayed on the menu and returns a Dictionary 
    ''' of those options and their respective files <br />
    ''' <br />
    ''' The set of possible files includes:
    ''' <list type="bullet">
    '''     
    '''     <item>
    '''     Older/Local file
    '''     </item>
    '''     
    '''     <item>
    '''     Newer file (not available when downloading)
    '''     </item>
    '''     
    '''     <item>
    '''     Save target (not available when not saving log)
    '''     </item>
    '''     
    ''' </list>
    ''' 
    ''' </summary>
    ''' 
    ''' <returns> 
    ''' The set of <c> iniFile </c> properties currently displayed on the menu
    ''' </returns>
    ''' 
    ''' Docs last updated: 2025-08-12 | Code last updated: 2025-08-12
    Private Function getFileOpts() As Dictionary(Of String, iniFile)

        Dim selectors As New Dictionary(Of String, iniFile)

        selectors.Add(NameOf(DiffFile1), DiffFile1)

        If Not DownloadDiffFile Then selectors.Add(NameOf(DiffFile2), DiffFile2)

        If SaveDiffLog Then selectors.Add(NameOf(DiffFile3), DiffFile3)

        Return selectors

    End Function

End Module
