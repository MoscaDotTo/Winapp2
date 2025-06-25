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
''' Manages the settings of the Diff module for the purpose of syncing to disk 
''' </summary>
''' 
''' Docs last updated: 2024-05-08
Module diffsettingshandler

    ''' <summary> 
    ''' Restores the default state of the Diff module's properties 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2020-06-10 | Code last updated: 2020-06-10
    Public Sub InitDefaultDiffSettings()

        DownloadDiffFile = Not isOffline
        TrimRemoteFile = Not isOffline
        ShowFullEntries = False
        SaveDiffLog = False
        DiffModuleSettingsChanged = False
        DiffFile3.resetParams()
        DiffFile2.resetParams()
        DiffFile1.resetParams()
        restoreDefaultSettings(NameOf(Diff), AddressOf CreateDiffSettingsSection)

    End Sub

    ''' <summary> 
    ''' Assigns the module settings to Diff based on the current disk-writable settings representation
    ''' </summary>
    ''' 
    ''' Docs last updated: 2024-05-08 | Code last updated: 2024-05-08
    Public Sub GetSerializedDiffSettings()

        LoadModuleSettingsFromDict(NameOf(Diff), GetType(diffsettings))

    End Sub

    ''' <summary>
    ''' Adds the current state of the module's settings into the disk-writable settings representation 
    ''' </summary>
    ''' 
    ''' <remarks>
    ''' Most often, this is the default state of these settings 
    ''' </remarks>
    ''' 
    ''' Docs last updated: 2023-06-12 | Code last updated: 2024-05-08
    Public Sub CreateDiffSettingsSection()

        Dim diffSettingsTuple = GetSettingsTupleWithReflection(GetType(diffsettings))

        createModuleSettingsSection(NameOf(Diff), diffSettingsTuple, getNumBools(GetType(diffsettings)), getNumFiles(GetType(diffsettings)))

    End Sub

End Module
