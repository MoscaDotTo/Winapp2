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
''' Displays the CC7Patcher module main menu and handles user input
''' </summary>
'''
Public Module cc7patchermainmenu

    ''' <summary>
    ''' Builds the CC7Patcher main menu
    ''' </summary>
    Friend Function buildCC7PatcherMenu() As MenuSection

        Dim menuDescLines = {"Patch ccleaner.ini with winapp2.ini entries compatible with CCleaner 7"}
        Dim settingsChangeName = NameOf(CC7PatcherModuleSettingsChanged)
        Dim moduleName = NameOf(CC7Patcher)

        Return MenuSection.CreateCompleteMenu(moduleName, menuDescLines, ConsoleColor.Yellow) _
            .AddDispatchedColoredOption("Run (default)", "Install winapp2.ini for CCleaner 7", ConsoleColor.Yellow,
                Sub() initCC7Patcher(), isDefault:=True) _
            .AddBlank() _
            .AddDispatchedToggle("Trim", "trimming winapp2.ini before installation", TrimBeforePatching,
                Sub() toggleModuleSetting("Trim", moduleName, GetType(cc7patchersettings), NameOf(TrimBeforePatching), settingsChangeName)) _
            .AddDispatchedToggle("include list", "never trimming certain entries", UseCC7PatcherIncludes,
                Sub() toggleModuleSetting("Include list", moduleName, GetType(cc7patchersettings), NameOf(UseCC7PatcherIncludes), settingsChangeName),
                TrimBeforePatching) _
            .AddDispatchedToggle("exclude list", "always trimming certain entries", UseCC7PatcherExcludes,
                Sub() toggleModuleSetting("Exclude list", moduleName, GetType(cc7patchersettings), NameOf(UseCC7PatcherExcludes), settingsChangeName),
                TrimBeforePatching) _
            .AddDispatchedToggle("Download", "downloading the latest winapp2.ini from GitHub", DownloadWinapp2, condition:=Not isOffline,
                handler:=Sub() toggleModuleSetting("Download", moduleName, GetType(cc7patchersettings), NameOf(DownloadWinapp2), settingsChangeName)) _
            .AddBlank() _
            .AddDispatchedOption("Change winapp2.ini", "Select the winapp2.ini file to install", condition:=Not DownloadWinapp2,
                handler:=Sub() changeFileParams(CC7PatcherFile1, CC7PatcherModuleSettingsChanged, moduleName, NameOf(CC7PatcherFile1), settingsChangeName)) _
            .AddDispatchedOption("Change ccleaner.ini", "Select the ccleaner.ini file to be patched",
                Sub() changeFileParams(CC7PatcherFile2, CC7PatcherModuleSettingsChanged, moduleName, NameOf(CC7PatcherFile2), settingsChangeName)) _
            .AddDispatchedOption("Change output file", "Select where to save the patched ccleaner.ini",
                Sub() changeFileParams(CC7PatcherFile3, CC7PatcherModuleSettingsChanged, moduleName, NameOf(CC7PatcherFile3), settingsChangeName)) _
            .AddDispatchedOption("Choose includes file", "Select a file containing entry names which should never be trimmed",
                Sub() changeFileParams(CC7PatcherFile4, CC7PatcherModuleSettingsChanged, moduleName, NameOf(CC7PatcherFile4), settingsChangeName),
                TrimBeforePatching AndAlso UseCC7PatcherIncludes) _
            .AddDispatchedOption("Choose excludes file", "Select a file containing entry names which should always be trimmed",
                Sub() changeFileParams(CC7PatcherFile5, CC7PatcherModuleSettingsChanged, moduleName, NameOf(CC7PatcherFile5), settingsChangeName),
                TrimBeforePatching AndAlso UseCC7PatcherExcludes) _
            .AddBlank() _
            .AddColoredFileInfo("Current winapp2.ini:  ", If(DownloadWinapp2, "Online", CC7PatcherFile1.Path()), ConsoleColor.Green) _
            .AddColoredFileInfo("Current ccleaner.ini: ", CC7PatcherFile2.Path(), ConsoleColor.Red) _
            .AddColoredFileInfo("Output file:          ", CC7PatcherFile3.Path(), ConsoleColor.Cyan) _
            .AddFileInfo("Includes file:        ", CC7PatcherFile4.Path(), condition:=TrimBeforePatching AndAlso UseCC7PatcherIncludes) _
            .AddFileInfo("Excludes file:        ", CC7PatcherFile5.Path(), condition:=TrimBeforePatching AndAlso UseCC7PatcherExcludes) _
            .AddBlank(CC7PatcherModuleSettingsChanged) _
            .AddDispatchedResetOpt(moduleName, CC7PatcherModuleSettingsChanged,
                Sub() resetModuleSettings(moduleName, AddressOf InitDefaultCC7PatcherSettings))

    End Function
End Module
