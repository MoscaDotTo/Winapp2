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
    ''' Builds the Diff main menu. When offline, building it also turns <c> DownloadDiffFile </c> off.
    ''' </summary>
    Friend Function buildDiffMenu() As MenuSection

        Dim newFileHasName = Not (DiffFile2.Name.Length = 0 AndAlso Not DownloadDiffFile)
        Dim newerFileText = If(Not newFileHasName, "Not yet selected", If(DownloadDiffFile, GetNameFromDL(True), replDir(DiffFile2.Path())))
        Dim olderFileText = If(DownloadDiffFile, "Local", "Older")
        Dim menuDesc = {"Observes the differences between two ini files"}

        If isOffline Then DownloadDiffFile = False

        Return MenuSection.CreateCompleteMenu(NameOf(Diff), menuDesc, ConsoleColor.DarkCyan) _
            .AddDispatchedColoredOption("Run (default)", "Perform a Diff operation", GetRedGreen(Not newFileHasName),
                Sub() If Not denyActionWithHeader(DiffFile2.Name.Length = 0 AndAlso Not DownloadDiffFile, "Please select a file against which to diff") Then ConductDiff(), isDefault:=True) _
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
                Sub() changeFileParams(DiffFile1, DiffModuleSettingsChanged, NameOf(Diff), NameOf(DiffFile1), NameOf(DiffModuleSettingsChanged))) _
            .AddDispatchedOption("Choose newer file", "Select the newer version of the file to see what has changed",
                Sub() changeFileParams(DiffFile2, DiffModuleSettingsChanged, NameOf(Diff), NameOf(DiffFile2), NameOf(DiffModuleSettingsChanged)),
                Not DownloadDiffFile) _
            .AddDispatchedOption("Choose save target", "Select where to save the diff output",
                Sub() changeFileParams(DiffFile3, DiffModuleSettingsChanged, NameOf(Diff), NameOf(DiffFile3), NameOf(DiffModuleSettingsChanged)),
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
End Module
