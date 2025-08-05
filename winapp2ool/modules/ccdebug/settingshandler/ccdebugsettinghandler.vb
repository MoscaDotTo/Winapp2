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
''' Manages the settings of the CCiniDebug module for the purpose of syncing to disk 
''' </summary>
''' 
''' Docs last updated: 2024-05-08 | Code last updated: 2025-06-25
Module ccdebugsettinghandler

    ''' <summary> 
    ''' Restores the default state the CCiniDebug module's properties 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2024-05-08 | Code last updated: 2020-07-18
    Public Sub initDefaultCCDBSettings()

        CCDebugFile1.resetParams()
        CCDebugFile2.resetParams()
        CCDebugFile3.resetParams()
        PruneStaleEntries = True
        SaveDebuggedFile = True
        SortFileForOutput = True
        CCDBSettingsChanged = False
        restoreDefaultSettings(NameOf(CCiniDebug), AddressOf createCCDBSettingsSection)

    End Sub

    ''' <summary>
    ''' Assigns the module settings to CCiniDebug based on the current disk-writable settings representation
    ''' </summary>
    ''' 
    ''' Docs last updated: 2024-05-08 | Code last updated: 2025-06-25
    Public Sub getSerializedDebugSettings()

        LoadModuleSettingsFromDict(NameOf(CCiniDebug), GetType(ccdebugsettings))

    End Sub

    ''' <summary>
    ''' Adds the current state of the module's settings into the disk-writable settings representation 
    ''' </summary>
    ''' 
    ''' <remarks>
    ''' Most often, this is the default state of these settings 
    ''' </remarks>
    ''' 
    ''' Docs last updated: 2024-05-08 | Code last updated: 2025-06-25
    Public Sub createCCDBSettingsSection()

        Dim ccdebugSettingsTuple = GetSettingsTupleWithReflection(GetType(ccdebugsettings))

        createModuleSettingsSection(NameOf(CCiniDebug), ccdebugSettingsTuple, getNumBools(GetType(ccdebugsettings)), getNumFiles(GetType(ccdebugsettings)))

    End Sub

End Module