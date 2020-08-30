'    Copyright (C) 2018-2020 Robbie Ward
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
''' <summary> Syncs the Diff module settings to and from disk </summary>
''' Docs last updated: 2020-08-30
Module diffsettingshandler

    ''' <summary> Restores the default state of the Diff module's properties </summary>
    ''' Docs last updated: 2020-07-23 | Code last updated: 2020-07-19
    Public Sub initDefaultDiffSettings()
        DownloadDiffFile = Not isOffline
        TrimRemoteFile = Not isOffline
        ShowFullEntries = False
        DiffFile3.resetParams()
        DiffFile2.resetParams()
        DiffFile1.resetParams()
        SaveDiffLog = False
        DiffModuleSettingsChanged = False
        restoreDefaultSettings(NameOf(Diff), AddressOf createDiffSettingsSection)
    End Sub

    ''' <summary> Assigns Diff module property values based on  </summary>
    ''' Docs last updated: 2020-07-23 | Code last updated: 2020-07-19
    Public Sub getSerializedDiffSettings()
        For Each kvp In settingsDict(NameOf(Diff))
            Select Case kvp.Key
                Case NameOf(DiffFile1) & "_Name"
                    DiffFile1.Name = kvp.Value
                Case NameOf(DiffFile1) & "_Dir"
                    DiffFile1.Dir = kvp.Value
                Case NameOf(DiffFile2) & "_Name"
                    DiffFile2.Name = kvp.Value
                Case NameOf(DiffFile2) & "_Dir"
                    DiffFile2.Dir = kvp.Value
                Case NameOf(DiffFile3) & "_Name"
                    DiffFile3.Name = kvp.Value
                Case NameOf(DiffFile3) & "_Dir"
                    DiffFile3.Dir = kvp.Value
                Case NameOf(DownloadDiffFile)
                    DownloadDiffFile = CBool(kvp.Value)
                Case NameOf(TrimRemoteFile)
                    TrimRemoteFile = CBool(kvp.Value)
                Case NameOf(ShowFullEntries)
                    ShowFullEntries = CBool(kvp.Value)
                Case NameOf(SaveDiffLog)
                    SaveDiffLog = CBool(kvp.Value)
                Case NameOf(DiffModuleSettingsChanged)
                    DiffModuleSettingsChanged = CBool(kvp.Value)
            End Select
        Next
    End Sub

    ''' <summary> Adds the current (typically default) state of the module's settings into the disk-writable settings representation </summary>
    ''' Docs last updated: 2020-08-30 | Code last updated: 2020-07-19
    Public Sub createDiffSettingsSection()
        Dim diffSettingsTuples As New List(Of String) From {NameOf(downloadFile), tsInvariant(DownloadDiffFile), NameOf(TrimRemoteFile), tsInvariant(TrimRemoteFile),
            NameOf(ShowFullEntries), tsInvariant(ShowFullEntries), NameOf(SaveDiffLog), tsInvariant(SaveDiffLog), NameOf(DiffModuleSettingsChanged), tsInvariant(DiffModuleSettingsChanged),
            NameOf(DiffFile1), DiffFile1.Name, DiffFile1.Dir, NameOf(DiffFile2), DiffFile2.Name, DiffFile2.Dir, NameOf(DiffFile3), DiffFile3.Name, DiffFile3.Dir}
        createModuleSettingsSection(NameOf(Diff), diffSettingsTuples, 5)
    End Sub

End Module
