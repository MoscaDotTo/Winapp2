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
''' Manages the settings of the CC7Patcher module for the purpose of syncing to disk
''' </summary>
''' 
''' Docs last updated: 2025-11-19 | Code last updated: 2025-11-19
Public Module cc7patchersettingshandler

    ''' <summary>
    ''' Restores the default state of the CC7Patcher module's properties
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-11-19 | Code last updated: 2025-11-19
    Public Sub initDefaultCC7PatcherSettings()

        CC7PatcherFile1.resetParams()
        CC7PatcherFile2.resetParams()
        CC7PatcherFile3.resetParams()
        DownloadWinapp2 = True
        TrimBeforePatching = False
        CC7PatcherModuleSettingsChanged = False
        restoreDefaultSettings(NameOf(CC7Patcher), AddressOf createCC7PatcherSettingsSection)

    End Sub

    ''' <summary>
    ''' Assigns the module settings to CC7Patcher based on the current disk-writable settings representation
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-11-19 | Code last updated: 2025-11-19
    Public Sub getSerializedCC7PatcherSettings()

        LoadModuleSettingsFromDict(NameOf(CC7Patcher), GetType(cc7patchersettings))

    End Sub

    ''' <summary>
    ''' Adds the current state of the module's settings into the disk-writable settings representation
    ''' </summary>
    ''' 
    ''' <remarks>
    ''' Most often, this is the default state of these settings
    ''' </remarks>
    ''' 
    ''' Docs last updated: 2025-11-19 | Code last updated: 2025-11-19
    Public Sub createCC7PatcherSettingsSection()

        Dim settingsModule = GetType(cc7patchersettings)
        Dim moduleName = NameOf(CC7Patcher)

        createModuleSettingsSection(moduleName, settingsModule)

    End Sub

End Module