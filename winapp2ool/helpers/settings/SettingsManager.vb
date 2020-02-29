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
'''<summary>Holds functions that modify and manage settings for winapp2ool and its sub modules</summary>
Module SettingsManager
    ''' <summary>Prompts the user to change a file's parameters, marks both settings and the file as having been changed </summary>
    ''' <param name="someFile">A pointer to an iniFile whose parameters will be changed</param>
    ''' <param name="settingsChangedSetting">A pointer to the boolean indicating that a module's settings been modified from their default state </param>
    Public Sub changeFileParams(ByRef someFile As iniFile, ByRef settingsChangedSetting As Boolean, Optional callingModule As String = "", Optional settingName As String = "")
        Dim curName = someFile.Name
        Dim curDir = someFile.Dir
        initModule("File Chooser", AddressOf someFile.printFileChooserMenu, AddressOf someFile.handleFileChooserInput)
        settingsChangedSetting = True
        Dim fileChanged = Not someFile.Name = curName Or Not someFile.Dir = curDir
        updateSettings(callingModule, $"{settingName}_Dir", someFile.Dir)
        updateSettings(callingModule, $"{settingName}_Name", someFile.Name)
        setHeaderText($"{If(someFile.SecondName.Length = 0, someFile.InitName, "save file")} parameters update{If(Not fileChanged, " aborted", "d")}", Not fileChanged)
    End Sub

    ''' <summary>Toggles a setting's boolean state and marks its tracker true</summary>
    ''' <param name="setting">A pointer to the boolean representing a module setting to be toggled</param>
    ''' <param name="paramText">A string explaining the setting being toggled</param>
    ''' <param name="mSettingsChanged">A pointer to the boolean indicating that a module's settings been modified from their default state</param>
    Public Sub toggleSettingParam(ByRef setting As Boolean, paramText As String, ByRef mSettingsChanged As Boolean, Optional callingModule As String = "",
                                  Optional settingName As String = "", Optional settingChangedName As String = "")
        gLog($"Toggling {paramText}", indent:=True)
        setHeaderText($"{paramText} {enStr(setting)}d", True, True, If(Not setting, ConsoleColor.Green, ConsoleColor.Red))
        setting = Not setting
        mSettingsChanged = True
        updateSettings(callingModule, settingName, setting.ToString)
        updateSettings(callingModule, settingChangedName, mSettingsChanged.ToString)
        settingsFile.overwriteToFile(settingsFile.toString, settingName = NameOf(saveSettingsToDisk) Or settingName = NameOf(readSettingsFromDisk))
    End Sub

    ''' <summary>Handles toggling downloading of winapp2.ini from menus</summary>
    ''' <param name="download">A pointer to a boolean that indicates whether or not a file should be downloaded</param>
    ''' <param name="settingsChanged">A pointer to the boolean indicating that a module's settings been modified from their default state</param>
    Public Sub toggleDownload(ByRef download As Boolean, ByRef settingsChanged As Boolean)
        If Not denySettingOffline() Then toggleSettingParam(download, "Downloading", settingsChanged)
    End Sub

    ''' <summary>Resets a module's settings to the defaults</summary>
    ''' <param name="name">The name of the module whose settings will be reset</param>
    ''' <param name="setDefaultParams">The function that resets the module's settings to their default state</param>
    Public Sub resetModuleSettings(name As String, setDefaultParams As Action)
        gLog($"Restoring {name}'s module settings to their default states", indent:=True)
        setDefaultParams()
        setHeaderText($"{name} settings have been reset to their defaults.")
    End Sub

    ''' <summary>Denies the ability to access online-only functions if offline</summary>
    Public Function denySettingOffline() As Boolean
        gLog("Action was unable to complete because winapp2ool is offline", isOffline)
        setHeaderText("This option is unavailable while in offline mode", True, isOffline)
        Return isOffline
    End Function
End Module