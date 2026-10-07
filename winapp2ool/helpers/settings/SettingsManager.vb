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
Module SettingsManager

    ''' <summary>
    ''' Opens the file chooser menu for an <c> iniFileChooser </c>. If the user changed its
    ''' name or directory, we save the new path through <see cref="saveChooserParams"/>.
    ''' Either way, the next menu header reports whether the update happened.
    ''' </summary>
    '''
    ''' <param name="chooser">
    ''' The <c> iniFileChooser </c> whose parameters will be changed
    ''' </param>
    '''
    ''' <param name="settingsChangedSetting">
    ''' The module's settings-changed flag. Set to <c> True </c> if the path changed.
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
    ''' The name of <paramref name="settingsChangedSetting"/> as it appears in the codebase
    ''' </param>
    '''
    ''' <param name="fileDesc">
    ''' A description of the file, shown in the header message <br /><br />
    ''' Optional, Default: <c> "" </c>
    ''' </param>
    Public Sub changeFileParams(ByRef chooser As iniFileChooser,
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
    ''' Writes an <c> iniFileChooser </c>'s current parameters and the owning module's
    ''' settings-changed flag into <see cref="SettingsFile"/>, then calls
    ''' <see cref="FlushIfDirty"/>, so they reach disk only when the save gate allows it.
    ''' <br /> Every path which lets the user pick a file has to come through here, including
    ''' <c> iniFileChooser.Load </c>'s missing-file prompt, or the choice is lost at exit.
    ''' </summary>
    '''
    ''' <param name="chooser">
    ''' The <c> iniFileChooser </c> whose parameters will be saved
    ''' </param>
    '''
    ''' <param name="settingsChangedSetting">
    ''' The module's settings-changed flag. Always set to <c> True </c>.
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
    ''' The name of <paramref name="settingsChangedSetting"/> as it appears in the codebase
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

        FlushIfDirty()

    End Sub

    ''' <summary>
    ''' Inverts a Boolean setting property, marks its owning module's settings as having been changed,
    ''' and records both in <see cref="SettingsFile"/>. The <see cref="FlushIfDirty"/> that follows
    ''' writes them to disk only when the save gate allows it.
    ''' </summary>
    '''
    ''' <remarks>
    ''' <paramref name="settingName"/> and <paramref name="settingChangedName"/> are both resolved
    ''' against <paramref name="settingsModule"/>, so a setting declared outside its module's settings
    ''' type cannot be toggled here
    ''' </remarks>
    '''
    ''' <param name="paramText">
    ''' The name of the setting as it should be displayed to the user
    ''' </param>
    '''
    ''' <param name="callingModule">
    ''' The name of the module owning the setting as it appears in the settings file
    ''' </param>
    '''
    ''' <param name="settingsModule">
    ''' The <c> Type </c> declaring both the setting and its changed flag
    ''' </param>
    '''
    ''' <param name="settingName">
    ''' The name of the setting property as it appears in the codebase
    ''' </param>
    '''
    ''' <param name="settingChangedName">
    ''' The name of the changed flag property as it appears in the codebase
    ''' </param>
    Public Sub toggleModuleSetting(paramText As String,
                                   callingModule As String,
                                   settingsModule As Type,
                                   settingName As String,
                                   settingChangedName As String)

        Dim settingProp = settingsModule.GetProperty(settingName)
        Dim changedProp = settingsModule.GetProperty(settingChangedName)

        If settingProp Is Nothing OrElse changedProp Is Nothing Then

            Dim missingName = If(settingProp Is Nothing, settingName, settingChangedName)

            gLog($"  {settingsModule.Name} declares no property named {missingName}")
            setNextMenuHeaderText($"{paramText} could not be changed", printColor:=ConsoleColor.Red)
            argIsInvalid($"{settingsModule.Name}.{missingName}")

            Return

        End If

        Dim setting = CBool(settingProp.GetValue(Nothing, Nothing))

        gLog($"  Toggling {paramText} from {setting} to {Not setting}")
        setNextMenuHeaderText($"{paramText} {enStr(setting)}d", printColor:=GetRedGreen(setting))

        setting = Not setting

        settingProp.SetValue(Nothing, setting)
        changedProp.SetValue(Nothing, True)
        SetSetting(callingModule, settingName, setting.ToString(CultureInfo.InvariantCulture))
        SetSetting(callingModule, settingChangedName, True.ToString)

        FlushIfDirty()

    End Sub

    ''' <summary>
    ''' Resets a module's settings to the defaults by calling <paramref name="setDefaultParams"/>,
    ''' then reports the reset in the next menu header. We don't flush here, so whatever
    ''' <paramref name="setDefaultParams"/> records in <see cref="SettingsFile"/> reaches disk only
    ''' when a later flush gets past the save gate.
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
    ''' Returns whether winapp2ool is offline, so a caller can refuse an online-only action.
    ''' When it is, we also log the refusal and set an error as the next menu header.
    ''' </summary>
    '''
    ''' <returns>
    ''' <c> True </c> if the caller should refuse the action, <c> False </c> if it can go ahead
    ''' </returns>
    Public Function denySettingOffline() As Boolean

        gLog("An action was unable to complete because winapp2ool is offline", isOffline)
        setNextMenuHeaderText("This option is unavailable while in offline mode", cond:=isOffline, printColor:=ConsoleColor.Red)

        Return isOffline

    End Function

    ''' <summary>
    ''' Cycles an enum property to its next value in numeric order, wrapping to the lowest,
    ''' marks its settings changed flag, and records both in <see cref="SettingsFile"/>.
    ''' We don't flush here, so the change reaches disk only when a later flush gets past
    ''' the save gate.
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
    ''' The calling module's settings-changed flag. Set to <c> True </c> once the property
    ''' changes.
    ''' </param>
    '''
    ''' <param name="settingsChangedName">
    ''' The name of <paramref name="mSettingsChanged"/> as it appears in the codebase
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

        If p Is Nothing Then

            gLog($"  {propertyType.Name} declares no property named {propName}")
            setNextMenuHeaderText($"{displayName} could not be changed", printColor:=ConsoleColor.Red)
            argIsInvalid($"{propertyType.Name}.{propName}")

            Return

        End If

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
