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
Imports System.Reflection

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
                                      settingChangedName As String,
                             Optional fileDesc As String = "")

        Dim curName = someFile.Name
        Dim curDir = someFile.Dir

        initModule("File Chooser", AddressOf someFile.printFileChooserMenu, AddressOf someFile.handleFileChooserInput)

        Dim fileChanged = Not someFile.Name = curName OrElse Not someFile.Dir = curDir
        If Not settingsChangedSetting Then settingsChangedSetting = fileChanged

        setHeaderText($"{If(someFile.SecondName.Length = 0, someFile.InitName, "save file")} parameters update{If(Not fileChanged, " aborted", "d")}", Not fileChanged)
        setNextMenuHeaderText($"{fileDesc} parameters update{If(Not fileChanged, " aborted", "d")}", printColor:=GetRedGreen(Not fileChanged))
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
        setHeaderText($"{paramText} {enStr(setting)}d", True, True, If(Not setting, ConsoleColor.Green, ConsoleColor.Red))
        'setNextMenuHeaderText($"{paramText} {enStr(setting)}d", printColor:=If(Not setting, ConsoleColor.Green, ConsoleColor.Red))
        setting = Not setting
        mSettingsChanged = True

        updateSettings(callingModule, settingName, setting.ToString(CultureInfo.InvariantCulture))
        updateSettings(callingModule, settingChangedName, mSettingsChanged.ToString(CultureInfo.InvariantCulture))

        Dim isSaveReadSetting = settingName = NameOf(saveSettingsToDisk) OrElse settingName = NameOf(readSettingsFromDisk)
        settingsFile.overwriteToFile(settingsFile.toString, Not IsCommandLineMode AndAlso isSaveReadSetting)

    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' 
    ''' <param name="paramText">
    ''' 
    ''' </param>
    ''' 
    ''' <param name="callingModule">
    ''' 
    ''' </param>
    ''' 
    ''' <param name="settingsModule">
    ''' 
    ''' </param>
    ''' 
    ''' <param name="settingName">
    ''' 
    ''' </param>
    ''' 
    ''' <param name="settingChangedName">
    ''' 
    ''' </param>
    Public Sub toggleModuleSetting(paramText As String,
                                   callingModule As String,
                                   settingsModule As Type,
                                   settingName As String,
                                   settingChangedName As String)

        Dim setting = CBool(settingsModule.GetProperty(settingName).GetValue(Nothing, Nothing))

        gLog($"Toggling {paramText} from {setting} to {Not setting}", indent:=True)
        setNextMenuHeaderText($"{paramText} {enStr(setting)}d", printColor:=GetRedGreen(setting))

        setting = Not setting

        settingsModule.GetProperty(settingName).SetValue(settingName, setting)
        settingsModule.GetProperty(settingChangedName).SetValue(settingChangedName, True)
        updateSettings(callingModule, settingName, setting.ToString(CultureInfo.InvariantCulture))
        updateSettings(callingModule, settingChangedName, True.ToString)

        settingsFile.overwriteToFile(settingsFile.toString, Not IsCommandLineMode AndAlso saveSettingsToDisk)

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
    Public Sub resetModuleSettings(name As String,
                                setDefaultParams As Action)

        gLog($"Restoring {name}'s module settings to their default states", indent:=True)

        setDefaultParams()

        setHeaderText($"{name} settings have been reset to their defaults.")

    End Sub

    ''' <summary>
    ''' Denies the ability to access online-only functions if offline
    ''' </summary>
    Public Function denySettingOffline() As Boolean

        gLog("An action was unable to complete because winapp2ool is offline", isOffline)
        setNextMenuHeaderText("This option is unavailable while in offline mode", printColor:=ConsoleColor.Red)

        Return isOffline

    End Function

    ''' <summary>
    ''' Cycles an enum property to its next value, marks its settings changed flag,
    ''' and updates the disk-writable settings representation
    ''' </summary>
    ''' 
    ''' <param name="propName">
    ''' The name of the Enum property as it appears in the codebase 
    ''' </param>
    ''' 
    ''' <param name="displayName">
    ''' The name of the Enum property as it should be displayed to the user
    ''' </param>
    ''' 
    ''' <param name="propertyType">
    ''' The <c> Type </c> containing the Enum property to be cycled 
    ''' </param>
    ''' 
    ''' <param name="moduleName">
    ''' The name of the module containing the Enum property
    ''' </param>
    ''' 
    ''' <param name="mSettingsChanged">
    ''' Indicates that the calling modules settings have been changed
    ''' </param>
    ''' 
    ''' <param name="settingsChangedName">
    ''' The name of <c> <paramref name="mSettingsChanged"/> </c> as it appears in the codebase
    ''' </param>
    ''' 
    ''' <param name="printColor">
    ''' The color with which to print the success message
    ''' </param>
    Public Sub CycleEnumProperty(propName As String,
                                 displayName As String,
                                 propertyType As Type,
                                 moduleName As String,
                           ByRef mSettingsChanged As Boolean,
                                 settingsChangedName As String,
                                 printColor As ConsoleColor)

        Dim p = propertyType.GetProperty(propName)

        Dim enumType = p.PropertyType
        Dim curObj = p.GetValue(Nothing)
        Dim enumValues = [Enum].GetValues(enumType)
        Dim currentIndex = Array.IndexOf(enumValues, curObj)
        Dim nextIndex = (currentIndex + 1) Mod enumValues.Length
        Dim nextValue = enumValues.GetValue(nextIndex)

        p.SetValue(Nothing, nextValue)

        mSettingsChanged = True
        updateSettings(moduleName, settingsChangedName, True.ToString)

        gLog()
        setNextMenuHeaderText($"{displayName} set to {nextValue}", printColor:=printColor)

    End Sub

    ''' <summary>
    ''' Gets the set of menu numbers associated with a dictionary of options, starting from a specified base number
    ''' </summary>
    ''' 
    ''' <typeparam name="T">
    ''' The type of the dictionary's values
    ''' </typeparam>
    ''' 
    ''' <param name="optionsDict">
    ''' A dictionary of menu options for which to generate menu numbers
    ''' </param>
    ''' 
    ''' <param name="baseNum">
    ''' The number from which to start numbering the options
    ''' </param>
    ''' 
    ''' <returns>
    ''' The set of numbers associated with the visible menu options in <c> <paramref name="optionsDict"/> </c>
    ''' </returns>
    Public Function getMenuNumbering(Of T)(optionsDict As Dictionary(Of String, T),
                                           baseNum As Integer) As List(Of String)

        Dim optNums = New List(Of String)
        For i = 0 To optionsDict.Count - 1

            Dim curNum = baseNum + i
            optNums.Add(curNum.ToString)

        Next

        Return optNums

    End Function

End Module
