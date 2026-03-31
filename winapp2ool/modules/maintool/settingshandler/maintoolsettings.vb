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
''' Holds the persisted settings for the main <c> Winapp2ool </c> module.
''' </summary>
Public Module maintoolsettings

    ''' <summary>
    ''' Holds the filesystem location to which the log file will optionally be saved.
    ''' </summary>
    Public Property GlobalLogFile As iniFileChooser = New iniFileChooser(Environment.CurrentDirectory, "winapp2ool.log", mustExist:=False)

    ''' <summary>
    ''' Indicates that the module's settings have been changed
    ''' </summary>
    Public Property toolSettingsHaveChanged As Boolean = False

    ''' <summary>
    ''' Indicates that changes to the application's settings should be serialized back to the disk
    ''' </summary>
    Public Property saveSettingsToDisk As Boolean = False

    ''' <summary>
    ''' Indicates that settings who are read from the disk should override the corresponding default module settings
    ''' </summary>
    Public Property readSettingsFromDisk As Boolean = False

    ''' <summary>
    ''' Indicates that this build is beta and should check the beta branch link for updates
    ''' </summary>
    Public Property isBeta As Boolean = False

    ''' <summary>
    ''' The currently selected winapp.ini flavor
    ''' </summary>
    Public Property CurrentWinappFlavor As Winapp2ool.WinappFlavor = Winapp2ool.WinappFlavor.CCleaner

    ''' <summary>
    ''' Restores all <c> Winapp2ool </c> settings to their defaults and persists the reset to disk
    ''' </summary>
    Public Sub InitDefaultToolSettings()

        GlobalLogFile.ResetParams()
        toolSettingsHaveChanged = False
        saveSettingsToDisk = False
        readSettingsFromDisk = False
        isBeta = False
        CurrentWinappFlavor = WinappFlavor.CCleaner
        SaveModule2(NameOf(Winapp2ool), GetType(maintoolsettings))

    End Sub

End Module
