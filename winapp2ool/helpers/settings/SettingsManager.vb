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
    ''' Prompts the user to change an <c>iniFileChooser</c>'s parameters, marks both settings and the chooser as having been changed
    ''' </summary>
    '''
    ''' <param name="chooser">
    ''' The <c>iniFileChooser</c> whose parameters will be changed
    ''' </param>
    '''
    ''' <param name="settingsChangedSetting">
    ''' A pointer to the boolean indicating that a module's settings have been modified from their default state
    ''' </param>
    Public Sub changeFile2Params(ByRef chooser As iniFileChooser,
                                 ByRef settingsChangedSetting As Boolean,
                                       callingModule As String,
                                       settingName As String,
                                       settingChangedName As String,
                              Optional fileDesc As String = "")

        Dim curName = chooser.Name
        Dim curDir = chooser.Dir

        initModule("File Chooser", AddressOf chooser.PrintMenu, AddressOf chooser.HandleInput)

        Dim fileChanged = Not chooser.Name = curName OrElse Not chooser.Dir = curDir

        setNextMenuHeaderText($"{fileDesc} parameters update{If(Not fileChanged, " aborted", "d")}", printColor:=GetRedGreen(Not fileChanged))
        If Not fileChanged Then Return

        saveChooserParams(chooser, settingsChangedSetting, callingModule, settingName, settingChangedName)

    End Sub

    ''' <summary>
    ''' Persists an <c> iniFileChooser </c>'s current parameters into the settings file, marking the
    ''' owning module's settings as having been changed
    ''' <br /> Called by every path which lets the user pick a file, including
    ''' <c> iniFileChooser.Load </c>'s missing-file prompt
    ''' </summary>
    '''
    ''' <param name="chooser">
    ''' The <c> iniFileChooser </c> whose parameters will be saved
    ''' </param>
    '''
    ''' <param name="settingsChangedSetting">
    ''' A pointer to the boolean indicating that a module's settings have been modified from their default state
    ''' </param>
    '''
    ''' <param name="callingModule">
    ''' The name of the module owning <paramref name="chooser"/> as it appears in the settings file
    ''' </param>
    '''
    ''' <param name="settingName">
    ''' The name of <paramref name="chooser"/> as it appears in the codebase
    ''' </param>
    '''
    ''' <param name="settingChangedName">
    ''' The name of <c> <paramref name="settingsChangedSetting"/> </c> as it appears in the codebase
    ''' </param>
    Public Sub saveChooserParams(chooser As iniFileChooser,
                           ByRef settingsChangedSetting As Boolean,
                                 callingModule As String,
                                 settingName As String,
                                 settingChangedName As String)

        settingsChangedSetting = True

        SetSetting(callingModule, $"{settingName}_Dir", chooser.Dir)
        SetSetting(callingModule, $"{settingName}_Name", chooser.Name)
        SetSetting(callingModule, settingChangedName, settingsChangedSetting.ToString(CultureInfo.InvariantCulture))

        FlushIfDirty2()

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

        gLog($"  Toggling {paramText} from {setting} to {Not setting}")
        setNextMenuHeaderText($"{paramText} {enStr(setting)}d", printColor:=GetRedGreen(setting))

        setting = Not setting

        settingsModule.GetProperty(settingName).SetValue(settingName, setting)
        settingsModule.GetProperty(settingChangedName).SetValue(settingChangedName, True)
        SetSetting(callingModule, settingName, setting.ToString(CultureInfo.InvariantCulture))
        SetSetting(callingModule, settingChangedName, True.ToString)

        FlushIfDirty2()

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

        gLog($"  Restoring {name}'s module settings to their default states")

        setDefaultParams()

        setNextMenuHeaderText($"{name} settings have been reset to their defaults.")

    End Sub

    ''' <summary>
    ''' Denies the ability to access online-only functions if offline
    ''' </summary>
    Public Function denySettingOffline() As Boolean

        gLog("An action was unable to complete because winapp2ool is offline", isOffline)
        setNextMenuHeaderText("This option is unavailable while in offline mode", cond:=isOffline, printColor:=ConsoleColor.Red)

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
        SetSetting(moduleName, propName, nextValue.ToString())
        SetSetting(moduleName, settingsChangedName, True.ToString)

        gLog()
        setNextMenuHeaderText($"{displayName} set to {nextValue}", printColor:=printColor)

    End Sub

End Module
