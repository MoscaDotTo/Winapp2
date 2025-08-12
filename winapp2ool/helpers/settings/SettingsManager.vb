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
Imports System.Globalization

''' <summary>
''' Provides functions to manage winapp2ool module settings, including modifying file parameters, 
''' individual module settings (Boolean and Enum), resetting the state of a module's settings to
''' their defaults, and gating functions behind internet access
''' </summary>
''' 
''' Docs last updated: 2025-07-22 | Code last updated: 2025-07-22
Module SettingsManager

    ''' <summary>
    ''' Prompts the user to change a file's parameters, marks both settings and the file as having been changed 
    ''' </summary>
    ''' 
    ''' <param name="someFile">
    ''' A pointer to an iniFile whose parameters will be changed
    ''' </param>
    ''' 
    ''' <param name="settingsChangedSetting">
    ''' A pointer to the boolean indicating that a module's settings been modified from their default state 
    ''' </param>
    ''' 
    ''' Docs last updated: 2025-07-22 | Code last updated: 2025-08-05
    Public Sub changeFileParams(ByRef someFile As iniFile,
                                ByRef settingsChangedSetting As Boolean,
                                      callingModule As String,
                                      settingName As String,
                                      settingChangedName As String)

        Dim curName = someFile.Name
        Dim curDir = someFile.Dir

        initModule("File Chooser", AddressOf someFile.printFileChooserMenu, AddressOf someFile.handleFileChooserInput)

        Dim fileChanged = Not someFile.Name = curName OrElse Not someFile.Dir = curDir
        If Not settingsChangedSetting Then settingsChangedSetting = fileChanged

        setHeaderText($"{If(someFile.SecondName.Length = 0, someFile.InitName, "save file")} parameters update{If(Not fileChanged, " aborted", "d")}", Not fileChanged)

        If Not fileChanged Then Return

        updateSettings(callingModule, $"{settingName}_Dir", someFile.Dir)
        updateSettings(callingModule, $"{settingName}_Name", someFile.Name)
        updateSettings(callingModule, settingChangedName, settingsChangedSetting.ToString(CultureInfo.InvariantCulture))

        settingsFile.overwriteToFile(settingsFile.toString, Not IsCommandLineMode AndAlso saveSettingsToDisk)

    End Sub

    ''' <summary>
    ''' Toggles a setting's boolean state and marks its tracker true
    ''' </summary>
    ''' 
    ''' <param name="setting">
    ''' A pointer to the boolean representing a module setting to be toggled
    ''' </param>
    ''' 
    ''' <param name="paramText">
    ''' A string explaining the setting being toggled
    ''' </param>
    ''' 
    ''' <param name="mSettingsChanged">
    ''' A pointer to the boolean indicating that a module's settings been modified from their default state
    ''' </param>
    ''' 
    ''' <param name="callingModule">
    ''' The name of the module calling this function
    ''' </param>
    ''' 
    ''' <param name="settingName">
    ''' The name of the setting being toggled
    ''' </param>
    ''' 
    ''' <param name="settingChangedName">
    ''' The name of the settings changed flag
    ''' </param>  
    ''' 
    ''' Docs last updated: 2025-07-22 | Code last updated: 2025-07-22
    Public Sub toggleSettingParam(ByRef setting As Boolean,
                                        paramText As String,
                                  ByRef mSettingsChanged As Boolean,
                                        callingModule As String,
                                        settingName As String,
                                        settingChangedName As String)

        gLog($"Toggling {paramText} from {setting} to {Not setting}", indent:=True)
        'setHeaderText($"{paramText} {enStr(setting)}d", True, True, If(Not setting, ConsoleColor.Green, ConsoleColor.Red))
        setNextMenuHeaderText($"{paramText} {enStr(setting)}d", printColor:=If(Not setting, ConsoleColor.Green, ConsoleColor.Red))
        setting = Not setting
        mSettingsChanged = True

        updateSettings(callingModule, settingName, setting.ToString(CultureInfo.InvariantCulture))
        updateSettings(callingModule, settingChangedName, mSettingsChanged.ToString(CultureInfo.InvariantCulture))

        Dim isSaveReadSetting = settingName = NameOf(saveSettingsToDisk) OrElse settingName = NameOf(readSettingsFromDisk)
        settingsFile.overwriteToFile(settingsFile.toString, Not IsCommandLineMode AndAlso isSaveReadSetting)

    End Sub

    ''' <summary>
    ''' Resets a module's settings to the defaults
    ''' </summary>
    ''' 
    ''' <param name="name">
    ''' The name of the module whose settings will be reset
    ''' </param>
    ''' 
    ''' <param name="setDefaultParams">
    ''' The function that resets the module's settings to their default state
    ''' </param>
    ''' 
    ''' Docs last updated: 2025-07-22 | Code last updated: 2025-07-22
    Public Sub resetModuleSettings(name As String,
                                   setDefaultParams As Action)

        gLog($"Restoring {name}'s module settings to their default states", indent:=True)

        setDefaultParams()

        setHeaderText($"{name} settings have been reset to their defaults.")

    End Sub

    ''' <summary>
    ''' Denies the ability to access online-only functions if offline
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-07-22 | Code last updated: 2025-07-22

    Public Function denySettingOffline() As Boolean

        gLog("Action was unable to complete because winapp2ool is offline", isOffline)
        setHeaderText("This option is unavailable while in offline mode", True, isOffline)

        Return isOffline

    End Function

    ''' <summary>
    ''' Cycles through enum values and updates the associated setting
    ''' </summary>
    ''' 
    ''' <param name="currentValue">
    ''' The current enum value
    ''' </param>
    ''' 
    ''' <param name="enumType">
    ''' The type of the enum
    ''' </param>
    ''' 
    ''' <param name="paramText">
    ''' Descriptive text for the setting
    ''' </param>
    ''' 
    ''' <param name="mSettingsChanged">
    ''' Reference to the module's settings changed flag
    ''' </param>
    ''' 
    ''' <param name="callingModule">
    ''' The name of the calling module
    ''' </param>
    ''' 
    ''' <param name="mSettingsChangedText">
    ''' The name of the settings changed flag
    ''' </param>
    ''' 
    ''' Docs last updated: 2025-07-15 | Code last updated: 2025-07-15
    Public Sub cycleEnumSetting(Of T As Structure)(ByRef currentValue As T,
                                                     enumType As Type,
                                                     paramText As String,
                                               ByRef mSettingsChanged As Boolean,
                                                     callingModule As String,
                                                     mSettingsChangedText As String)

        Dim enumValues = [Enum].GetValues(enumType)
        Dim currentIndex = Array.IndexOf(enumValues, currentValue)
        Dim nextIndex = (currentIndex + 1) Mod enumValues.Length

        Dim nextValue As Object = enumValues.GetValue(nextIndex)
        currentValue = CType(nextValue, T)

        gLog($"Cycling {paramText} to {currentValue}", indent:=True)
        setHeaderText($"{paramText} set to {currentValue}", True, True, ConsoleColor.Green)

        mSettingsChanged = True
        updateSettings(callingModule, mSettingsChangedText, tsInvariant(mSettingsChanged))

    End Sub

End Module
