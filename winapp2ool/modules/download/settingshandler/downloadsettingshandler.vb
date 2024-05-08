'    Copyright (C) 2018-2024 Hazel Ward
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
''' <summary> Syncs settings for the Download module to and from the disk </summary>
''' Docs last updated: 2020-09-14 | Code last updated: 2020-09-14
Module downloadsettingshandler

    ''' <summary> Loads values from disk into memory for the Downloader module settings </summary>
    ''' Docs last updated: 2020-09-14 | Code last updated: 2020-09-14
    Public Sub getSerializedDownloaderSettings()
        For Each kvp In settingsDict(NameOf(Downloader))
            Select Case kvp.Key
                Case NameOf(downloadFile) & "_Dir"
                    downloadFile.Dir = kvp.Value
                Case NameOf(DownloadModuleSettingsChanged)
                    DownloadModuleSettingsChanged = CBool(kvp.Value)
            End Select
        Next
    End Sub

    ''' <summary> Adds the current (typically default) state of the module's settings into the disk-writable settings representation </summary>
    ''' Docs last updated: 2020-09-14 | Code last updated: 2020-09-14
    Public Sub createDownloadSettingsSection()
        Dim downloadSettingTuple As New List(Of String) From {NameOf(DownloadModuleSettingsChanged), tsInvariant(DownloadModuleSettingsChanged), NameOf(downloadFile), "", downloadFile.Dir}
        createModuleSettingsSection(NameOf(Downloader), downloadSettingTuple, 1, 1)
    End Sub

End Module
