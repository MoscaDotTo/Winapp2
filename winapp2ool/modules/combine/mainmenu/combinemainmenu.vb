'    Copyright (C) 2018-2026 Hazel Ward
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
''' Displays the Combine module main menu and handles user input
''' </summary>
Public Module combinemainmenu

    ''' <summary>
    ''' Builds the Combine main menu
    ''' </summary>
    Friend Function buildCombineMenu() As MenuSection

        Dim menuDesc = {"Combine all ini files from a target directory and its children into a single file"}
        Dim noLog = MostRecentCombineLog.Length = 0

        Return MenuSection.CreateCompleteMenu(NameOf(Combine), menuDesc, ConsoleColor.DarkCyan) _
            .AddDispatchedOption("Run (default)", "Combine ini files", Sub() initCombine(CombineFile1.Dir, CombineFile3), isDefault:=True) _
            .AddBlank() _
            .AddDispatchedOption("Change the target directory", "File Chooser (target)",
                Sub() changeFileParams(CombineFile1, CombineModuleSettingsChanged, NameOf(Combine),
                                        NameOf(CombineFile1), NameOf(CombineModuleSettingsChanged))) _
            .AddDispatchedOption("Change the save file location", "File Chooser (save)",
                Sub() changeFileParams(CombineFile3, CombineModuleSettingsChanged, NameOf(Combine),
                                        NameOf(CombineFile3), NameOf(CombineModuleSettingsChanged))) _
            .AddBlank() _
            .AddDispatchedToggle("Strict Mode", "failing the build instead of merging when processing duplicate section names", CombineStrictNames,
                Sub() toggleModuleSetting("Strict name checking", NameOf(Combine), GetType(combinesettings), NameOf(CombineStrictNames), NameOf(CombineModuleSettingsChanged))) _
            .AddBlank() _
            .AddDispatchedColoredOption("Log Viewer", "View the detailed Combine log", ConsoleColor.Yellow,
                Sub() printSlice(MostRecentCombineLog), Not noLog) _
            .AddBlank(Not noLog) _
            .AddColoredFileInfo("Current target directory: ", CombineFile1.Dir, ConsoleColor.Cyan) _
            .AddColoredFileInfo("Current save location:    ", CombineFile3.Path(), ConsoleColor.Cyan) _
            .AddBlank(CombineModuleSettingsChanged) _
            .AddDispatchedResetOpt(NameOf(Combine), CombineModuleSettingsChanged,
                Sub() resetModuleSettings(NameOf(Combine), AddressOf InitDefaultCombineSettings))

    End Function
End Module
