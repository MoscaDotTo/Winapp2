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
''' <summary> Syncs settings for the Merge module to and from the disk </summary>
''' Docs last updated: 2020-09-14 | Code last updated: 2020-09-14
Module mergesettinghandler

    ''' <summary> Restores the default state of the module's properties </summary>
    ''' Docs last updated: 2020-09-14 | Code last updated: 2020-09-14
    Public Sub initDefaultMergeSettings()
        MergeFile1.resetParams()
        MergeFile2.resetParams()
        MergeFile3.resetParams()
        mergeMode = True
        MergeModuleSettingsChanged = False
        restoreDefaultSettings(NameOf(Merge), AddressOf createMergeSettingsSection)
    End Sub

    ''' <summary> Loads values from disk into memory for the Merge module settings </summary>
    ''' Docs last updated: 2020-09-14 | Code last updated: 2020-09-14
    Public Sub getSeralizedMergeSettings()
        For Each kvp In settingsDict(NameOf(Merge))
            Select Case kvp.Key
                Case NameOf(MergeFile1) & "_Name"
                    MergeFile1.Name = kvp.Value
                Case NameOf(MergeFile1) & "_Dir"
                    MergeFile1.Dir = kvp.Value
                Case NameOf(MergeFile2) & "_Name"
                    MergeFile2.Name = kvp.Value
                Case NameOf(MergeFile2) & "_Dir"
                    MergeFile2.Dir = kvp.Value
                Case NameOf(MergeFile3) & "_Name"
                    MergeFile3.Name = kvp.Value
                Case NameOf(MergeFile3) & "_Dir"
                    MergeFile1.Dir = kvp.Value
            End Select
        Next
    End Sub

    '''<summary> Adds the current (typically default) state of the module's settings into the disk-writable settings representation </summary>
    ''' Docs last updated: 2020-09-14 | Code last updated: 2020-09-14
    Public Sub createMergeSettingsSection()
        Dim mergeSettingsTuple As New List(Of String) From {NameOf(mergeMode), tsInvariant(mergeMode), NameOf(MergeModuleSettingsChanged), tsInvariant(MergeModuleSettingsChanged),
            NameOf(MergeFile1), MergeFile1.Name, MergeFile1.Dir, NameOf(MergeFile2), MergeFile2.Name, MergeFile2.Dir, NameOf(MergeFile3), MergeFile3.Name, MergeFile3.Dir}
        createModuleSettingsSection(NameOf(Merge), mergeSettingsTuple, 2)
    End Sub

    ''' <summary> Changes the merge file's <c> Name </c> </summary>
    ''' <param name="newName"> The new <c> Name </c> for <c> MergeFile2 </c> </param>
    ''' Docs last updated: 2020-09-14 | Code last updated: 2020-09-14
    Public Sub changeMergeName(newName As String)
        MergeFile2.Name = newName
        MergeModuleSettingsChanged = True
        setHeaderText("Merge filename set")
    End Sub
End Module
