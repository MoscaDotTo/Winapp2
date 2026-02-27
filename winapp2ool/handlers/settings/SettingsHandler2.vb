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

Imports System.Reflection

''' <summary>
''' A parallel settings backend powered by <c>iniFile2</c>.
''' <br />
''' <c>SettingsFile2</c> is the single authoritative in-memory representation
''' for modules that have been migrated away from the legacy
''' <c>settingsDict</c> + <c>settingsFile</c> dual-representation.
''' <br />
''' During the transition period, <c>settingsHandler</c> keeps <c>SettingsFile2</c>
''' up to date via <c>Load2</c> (called from <c>loadSettings</c>) and
''' <c>SetValue2</c> (called from <c>updateSettings</c>).
''' Migrated modules call <c>LoadModule2</c> / <c>SaveModule2</c> instead of the
''' legacy <c>LoadModuleSettingsFromDict</c> / <c>createModuleSettingsSection</c>.
''' </summary>
'''
''' Docs last updated: 2026-02-26 | Code last updated: 2026-02-26
Public Module SettingsHandler2

    Private _dirty2 As Boolean = False

    ''' <summary>
    ''' The <c>iniFile2</c>-backed representation of winapp2ool's settings.
    ''' This is the single source of truth for modules that have been migrated
    ''' to the new settings backend.
    ''' </summary>
    Public Property SettingsFile2 As iniFile2 = iniFile2.Empty(Environment.CurrentDirectory, "winapp2ool.ini")

    ''' <summary>
    ''' Reads <c>winapp2ool.ini</c> from disk into <c>SettingsFile2</c>.
    ''' No-ops if the file does not exist.
    ''' </summary>
    Public Sub Load2()

        If Not IO.File.Exists(SettingsFile2.Path()) Then Return
        SettingsFile2 = iniFile2.FromFile(SettingsFile2.Path())

    End Sub

    ''' <summary>
    ''' Returns the value of a setting from <c>SettingsFile2</c>,
    ''' or <c>""</c> if the module section or key is not found.
    ''' </summary>
    Public Function GetValue2(moduleName As String, settingName As String) As String

        Dim section = SettingsFile2.GetSection(moduleName)
        If section Is Nothing Then Return ""
        Dim key = section.Keys.GetKey(settingName)
        Return If(key Is Nothing, "", key.Value)

    End Function

    ''' <summary>
    ''' Sets or creates a setting in <c>SettingsFile2</c>,
    ''' creating the module section and/or key if absent.
    ''' Marks the backend dirty; the write is deferred to <c>FlushIfDirty2</c>.
    ''' </summary>
    Public Sub SetValue2(moduleName As String, settingName As String, value As String)

        Dim section = SettingsFile2.GetOrCreateSection(moduleName)
        Dim key = section.Keys.GetKey(settingName)

        If key Is Nothing Then
            section.AddKey(New iniKey2($"{settingName}={value}"))
        Else
            key.Value = value
        End If

        _dirty2 = True

    End Sub

    ''' <summary>
    ''' Writes <c>SettingsFile2</c> to disk, subject to <paramref name="condition"/>.
    ''' </summary>
    Public Sub Save2(Optional condition As Boolean = True)

        SettingsFile2.OverwriteToFile(SettingsFile2.ToString(), condition)
        _dirty2 = False

    End Sub

    ''' <summary>
    ''' Writes <c>SettingsFile2</c> to disk only if it has been modified since the last save.
    ''' </summary>
    Public Sub FlushIfDirty2(Optional condition As Boolean = True)

        If _dirty2 Then Save2(condition)

    End Sub

    ''' <summary>
    ''' Populates a module's public static properties from <c>SettingsFile2</c>.
    ''' <br />
    ''' Handles <c>Boolean</c>, <c>Enum</c>, and <c>iniFileChooser</c> property types.
    ''' Silently skips properties whose keys are absent in <c>SettingsFile2</c>.
    ''' </summary>
    Public Sub LoadModule2(moduleName As String, moduleType As Type)

        Dim section = SettingsFile2.GetSection(moduleName)
        If section Is Nothing Then Return

        For Each prop As PropertyInfo In moduleType.GetProperties()

            If Not prop.CanWrite Then Continue For

            If prop.PropertyType Is GetType(iniFileChooser) Then

                Dim nameKey = section.Keys.GetKey(prop.Name & "_Name")
                Dim dirKey = section.Keys.GetKey(prop.Name & "_Dir")
                If nameKey Is Nothing OrElse dirKey Is Nothing Then Continue For

                Dim chooser = TryCast(prop.GetValue(Nothing), iniFileChooser)
                If chooser Is Nothing Then Continue For

                chooser.Name = nameKey.Value
                chooser.Dir = dirKey.Value

                Continue For

            End If

            Dim k = section.Keys.GetKey(prop.Name)
            If k Is Nothing Then Continue For

            If prop.PropertyType.IsEnum Then

                prop.SetValue(Nothing, [Enum].Parse(prop.PropertyType, k.Value))

            ElseIf prop.PropertyType Is GetType(Boolean) Then

                Dim bVal As Boolean
                If Boolean.TryParse(k.Value, bVal) Then prop.SetValue(Nothing, bVal)

            End If

        Next

    End Sub

    ''' <summary>
    ''' Writes a module's public static properties into <c>SettingsFile2</c>.
    ''' <br />
    ''' Handles <c>Boolean</c>, <c>Enum</c>, and <c>iniFileChooser</c> property types.
    ''' </summary>
    Public Sub SaveModule2(moduleName As String, moduleType As Type)

        For Each prop As PropertyInfo In moduleType.GetProperties()

            If Not prop.CanRead OrElse Not prop.CanWrite Then Continue For

            Dim value = prop.GetValue(Nothing)
            If value Is Nothing Then Continue For

            If prop.PropertyType Is GetType(iniFileChooser) Then

                Dim chooser = DirectCast(value, iniFileChooser)
                SetValue2(moduleName, prop.Name & "_Name", chooser.Name)
                SetValue2(moduleName, prop.Name & "_Dir", chooser.Dir)

                Continue For

            End If

            If prop.PropertyType Is GetType(Boolean) OrElse prop.PropertyType.IsEnum Then
                SetValue2(moduleName, prop.Name, value.ToString())
            End If

        Next

    End Sub

End Module
