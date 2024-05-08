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
''' <summary> Syncs settings for the Trim module to and from the disk </summary>
''' Docs last updated: 2022-11-24 | Code last updated: 2022-11-24
Public Module trimsettingshandler
    ''' <summary> Restores the default state of the module's parameters </summary>
    ''' Docs last updated: 2022-11-21 | Code last updated: 2022-11-21
    Public Sub initDefaultTrimSettings()

        TrimFile1.resetParams()
        TrimFile2.resetParams()
        TrimFile3.resetParams()
        TrimFile4.resetParams()
        DownloadFileToTrim = False
        ModuleSettingsChanged = False
        UseTrimIncludes = False
        UseTrimExcludes = False
        restoreDefaultSettings(NameOf(Trim), AddressOf createTrimSettingsSection)

    End Sub

    ''' <summary> Loads values from disk into memory for the Trim module settings </summary>
    ''' Docs last updated: 2022-11-21 | Code last updated: 2022-11-21
    Public Sub getSerializedTrimSettings()

        For Each kvp In settingsDict(NameOf(Trim))

            Select Case kvp.Key

                Case NameOf(TrimFile1) & "_Name"

                    TrimFile1.Name = kvp.Value

                Case NameOf(TrimFile1) & "_Dir"

                    TrimFile1.Dir = kvp.Value

                Case NameOf(TrimFile2) & "_Name"

                    TrimFile2.Name = kvp.Value

                Case NameOf(TrimFile2) & "_Dir"

                    TrimFile2.Dir = kvp.Value

                Case NameOf(TrimFile3) & "_Name"

                    TrimFile3.Name = kvp.Value

                Case NameOf(TrimFile3) & "_Dir"

                    TrimFile3.Dir = kvp.Value

                Case NameOf(TrimFile4) & "_Name"

                    TrimFile4.Name = kvp.Value

                Case NameOf(TrimFile4) & "_Dir"

                    TrimFile4.Dir = kvp.Value

                Case NameOf(DownloadFileToTrim)

                    DownloadFileToTrim = CBool(kvp.Value)

                Case NameOf(ModuleSettingsChanged)

                    ModuleSettingsChanged = CBool(kvp.Value)

                Case NameOf(UseTrimIncludes)

                    UseTrimIncludes = CBool(kvp.Value)

                Case NameOf(UseTrimExcludes)

                    UseTrimExcludes = CBool(kvp.Value)

            End Select

        Next

    End Sub

    ''' <summary> Adds the current (typically default) state of the module's settings into the disk-writable settings representation </summary>
    ''' Docs last updated: 2022-11-21 | Code last updated: 2022-11-21
    Public Sub createTrimSettingsSection()

        Dim trimSettingsTuple As New List(Of String) From {NameOf(DownloadFileToTrim), tsInvariant(DownloadFileToTrim), NameOf(UseTrimIncludes), tsInvariant(UseTrimIncludes),
        NameOf(UseTrimExcludes), tsInvariant(UseTrimExcludes), NameOf(ModuleSettingsChanged), tsInvariant(ModuleSettingsChanged), NameOf(TrimFile1), TrimFile1.Name, TrimFile1.Dir,
        NameOf(TrimFile2), TrimFile2.Name, TrimFile2.Dir, NameOf(TrimFile3), TrimFile3.Name, TrimFile3.Dir, NameOf(TrimFile4), TrimFile4.Name, TrimFile4.Dir}
        createModuleSettingsSection(NameOf(Trim), trimSettingsTuple, 4, 4)

    End Sub

End Module