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
    ''' Indicates whether the module's settings have been changed from their defaults
    ''' </summary>
    Public Property toolSettingsHaveChanged As Boolean = False

    ''' <summary>
    ''' Indicates whether changes to the application's settings are written back to
    ''' <c> winapp2ool.ini </c>. <see cref="SaveSettings"/> refuses to write while it is
    ''' <c> False </c>. Turning it off is itself never written, so the file keeps the old
    ''' <c> True </c> and saving comes back on at the next launch.
    ''' </summary>
    Public Property saveSettingsToDisk As Boolean = False

    ''' <summary>
    ''' Indicates whether the settings in <c> winapp2ool.ini </c> override the other modules'
    ''' defaults at launch. When <c> False </c>, only this module's own settings are read.
    ''' </summary>
    Public Property readSettingsFromDisk As Boolean = False

    ''' <summary>
    ''' Indicates whether to follow the beta builds. When <c> True </c>, the update check,
    ''' the self-updater and the Downloader use the beta copies of winapp2ool.exe, its
    ''' signature and its version file.
    ''' </summary>
    Public Property isBeta As Boolean = False

    ''' <summary>
    ''' The currently selected winapp2.ini flavor
    ''' </summary>
    Public Property CurrentWinappFlavor As Winapp2ool.WinappFlavor = Winapp2ool.WinappFlavor.CCleaner

    ''' <summary>
    ''' Restores all <c> Winapp2ool </c> settings to their defaults and records them in
    ''' <see cref="SettingsFile"/> through <see cref="SaveModule"/>. That only changes memory:
    ''' the reset reaches disk when a later flush gets past the save gate, and since the reset
    ''' turns <see cref="saveSettingsToDisk"/> off, that happens only if saving is turned back on
    ''' this session.
    ''' </summary>
    Public Sub InitDefaultToolSettings()

        GlobalLogFile.ResetParams()
        toolSettingsHaveChanged = False
        saveSettingsToDisk = False
        readSettingsFromDisk = False
        isBeta = False
        CurrentWinappFlavor = WinappFlavor.CCleaner
        SaveModule(NameOf(Winapp2ool), GetType(maintoolsettings))

    End Sub

End Module
