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
Module mergesettinghandler

    ''' <summary> 
    ''' Restores the default state of the module's properties 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2020-09-14 | Code last updated: 2020-09-14
    Public Sub initDefaultMergeSettings()

        MergeFile1.resetParams()
        MergeFile2.resetParams()
        MergeFile3.resetParams()
        mergeMode = True
        MergeModuleSettingsChanged = False
        restoreDefaultSettings(NameOf(Merge), AddressOf createMergeSettingsSection)

    End Sub

    ''' <summary> 
    ''' Loads values from disk into memory for the Merge module settings 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2020-09-14 | Code last updated: 2020-09-14
    Public Sub getSeralizedMergeSettings()

        LoadModuleSettingsFromDict(NameOf(Merge), GetType(mergesettings))

    End Sub

    '''<summary>
    '''Adds the current (typically default) state of the module's settings into the disk-writable settings representation 
    '''</summary>
    '''
    ''' Docs last updated: 2020-09-14 | Code last updated: 2025-06-25
    Public Sub createMergeSettingsSection()

        Dim mergeSettingsTuple = GetSettingsTupleWithReflection(GetType(mergesettings))

        createModuleSettingsSection(NameOf(Merge), mergeSettingsTuple, getNumBools(GetType(mergesettings)), getNumFiles(GetType(mergesettings)))

    End Sub

    ''' <summary> 
    ''' Changes the merge file's <c> Name </c>
    ''' </summary>
    ''' <param name="newName"> The new <c> Name </c> for <c> MergeFile2 </c> </param>
    ''' 
    ''' Docs last updated: 2020-09-14 | Code last updated: 2020-09-14
    Public Sub changeMergeName(newName As String)

        MergeFile2.Name = newName
        MergeModuleSettingsChanged = True
        setHeaderText("Merge filename set")

    End Sub

End Module
