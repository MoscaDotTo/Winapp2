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
''' <summary> Syncs the winapp2ool module settings to and from disk  </summary>
Module mainToolSettingsHandler

    ''' <summary> Loads values from disk into memory for the winapp2ool module settings </summary>
    ''' Docs last updated: 2020-07-14 | Code last updated: 2020-07-14
    Public Sub getSeralizedToolSettings()
        For Each kvp In settingsDict(NameOf(Winapp2ool))
            Select Case kvp.Key
                Case NameOf(isBeta)
                    isBeta = CBool(kvp.Value)
                Case NameOf(readSettingsFromDisk)
                    readSettingsFromDisk = CBool(kvp.Value)
                Case NameOf(saveSettingsToDisk)
                    saveSettingsToDisk = CBool(kvp.Value)
                Case NameOf(RemoteWinappIsNonCC)
                    RemoteWinappIsNonCC = CBool(kvp.Value)
                Case NameOf(toolSettingsHaveChanged)
                    toolSettingsHaveChanged = CBool(kvp.Value)
                Case NameOf(GlobalLogFile) & "_Dir"
                    GlobalLogFile.Dir = kvp.Value
                Case NameOf(GlobalLogFile) & "_Name"
                    GlobalLogFile.Name = kvp.Value
            End Select
        Next
    End Sub

    '''<summary> Adds the current (typically default) state of the module's settings into the disk-writable settings representation </summary>
    ''' Docs last updated: 2020-07-14 | Code last updated: 2020-07-14
    Public Sub createToolSettingsSection()
        Dim compCult = System.Globalization.CultureInfo.InvariantCulture
        createModuleSettingsSection(NameOf(Winapp2ool), {
                                        getSettingIniKey(NameOf(Winapp2ool), NameOf(isBeta), isBeta.ToString(compCult)),
                                        getSettingIniKey(NameOf(Winapp2ool), NameOf(saveSettingsToDisk), saveSettingsToDisk.ToString(compCult)),
                                        getSettingIniKey(NameOf(Winapp2ool), NameOf(readSettingsFromDisk), readSettingsFromDisk.ToString(compCult)),
                                        getSettingIniKey(NameOf(Winapp2ool), NameOf(RemoteWinappIsNonCC), RemoteWinappIsNonCC.ToString(compCult)),
                                        getSettingIniKey(NameOf(Winapp2ool), NameOf(toolSettingsHaveChanged), toolSettingsHaveChanged.ToString(compCult)),
                                        getSettingIniKey(NameOf(Winapp2ool), NameOf(GlobalLogFile) & "_Dir", GlobalLogFile.Dir),
                                        getSettingIniKey(NameOf(Winapp2ool), NameOf(GlobalLogFile) & "_Name", GlobalLogFile.Name)
                                    })
    End Sub


End Module