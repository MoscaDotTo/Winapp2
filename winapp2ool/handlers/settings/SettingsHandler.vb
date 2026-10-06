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

Imports System.Diagnostics.Eventing.Reader
Imports System.Reflection

''' <summary>
''' The sole settings backend, powered by <c> iniFile </c>.
''' <br />
''' <see cref="SettingsFile"/> is the single authoritative in-memory representation of
''' <c> winapp2ool.ini </c>. All modules are registered in <see cref="loadAllModuleSettings"/>
''' and use <see cref="LoadModule"/> / <see cref="SaveModule"/> to move values between their
''' properties and <see cref="SettingsFile"/>. Only <see cref="SaveSettings"/> writes to disk.
''' </summary>
Public Module SettingsHandler

    ''' <summary>
    ''' Indicates whether <see cref="SettingsFile"/> has changed since the last write.
    ''' A save the gate refused leaves it set.
    ''' </summary>
    Private _dirty As Boolean = False

    ''' <summary>
    ''' The <c> iniFile </c>-backed representation of winapp2ool's settings, pointing at
    ''' <c> winapp2ool.ini </c> in the working directory.
    ''' </summary>
    Public Property SettingsFile As iniFile = iniFile.Empty(Environment.CurrentDirectory, "winapp2ool.ini")

    ''' <summary>
    ''' Reads <c> winapp2ool.ini </c> from disk into <see cref="SettingsFile"/>, then loads every
    ''' module's settings from it. When the file doesn't exist, we turn
    ''' <see cref="readSettingsFromDisk"/> and <see cref="saveSettingsToDisk"/> off and every
    ''' module keeps its defaults. When the file turns <see cref="readSettingsFromDisk"/> off,
    ''' only the <c> Winapp2ool </c> section is loaded.
    ''' </summary>
    Public Sub LoadWinapp2oolsettings()

        Using gLogScope("Loading settings")

            ' Handle the default case where winapp2ool.ini doesn't exist
            If Not System.IO.File.Exists(SettingsFile.Path) Then

                readSettingsFromDisk = False
                saveSettingsToDisk = False

                ' We still need to maintain an internal representation
                ' of the settings so create the settingsFile and settingsDict
                ' using the default winapp2ool configuration
                gLog("No settings file Found - loading default settings")
                loadAllModuleSettings()
                Return

            End If

            SettingsFile = iniFile.FromFile(SettingsFile.Path())
            loadAllModuleSettings()

        End Using

        gLog("Settings loaded", buffr:=True)

    End Sub

    ''' <summary>
    ''' Loads the settings for every module. A new module adds its <see cref="LoadModule"/>
    ''' call here. The <c> Winapp2ool </c> section loads first, and the rest load only if it
    ''' leaves <see cref="readSettingsFromDisk"/> on.
    ''' </summary>
    Private Sub loadAllModuleSettings()

        ' Winapp2ool is loaded first so readSettingsFromDisk is populated before other modules load
        LoadModule(NameOf(Winapp2ool), GetType(maintoolsettings))

        If Not readSettingsFromDisk Then Return

        LoadModule(NameOf(Diff), GetType(diffsettings))
        LoadModule(NameOf(UWPBuilder), GetType(uwpbuildersettings))
        LoadModule(NameOf(EntryBuilder), GetType(entryBuilderSettings))
        LoadModule(NameOf(BrowserBuilder), GetType(browserbuildersettings))
        LoadModule(NameOf(CC7Patcher), GetType(cc7patchersettings))
        LoadModule(NameOf(CCiniDebug), GetType(ccdebugsettings))
        LoadModule(NameOf(Combine), GetType(combinesettings))
        LoadModule(NameOf(Flavorizer), GetType(FlavorizerSettings))
        LoadModule(NameOf(Transmute), GetType(transmuteSettings))
        LoadModule(NameOf(Trim), GetType(trimsettings))
        LoadModule(NameOf(Downloader), GetType(downloadersettings))
        LoadModule(NameOf(WinappDebug), GetType(lintsettings))
        LoadLintRulesFromSettings()

    End Sub

    ''' <summary>
    ''' Returns the value of a setting from <see cref="SettingsFile"/>,
    ''' or <c> "" </c> if the module section or key is not found.
    ''' </summary>
    '''
    ''' <param name="moduleName">The name of the module's section</param>
    '''
    ''' <param name="settingName">The name of the setting's key</param>
    Public Function GetSetting(moduleName As String,
                               settingName As String) As String

        Dim section = SettingsFile.GetSection(moduleName)
        If section Is Nothing Then Return ""

        Dim key = section.Keys.GetKey(settingName)
        Return If(key Is Nothing, "", key.Value)

    End Function

    ''' <summary>
    ''' Sets a setting in <see cref="SettingsFile"/> in memory, creating the module section
    ''' and key if absent, and marks the backend dirty. Nothing reaches disk until a later
    ''' <see cref="FlushIfDirty"/> or <see cref="SaveSettings"/> gets past the save gate.
    ''' </summary>
    '''
    ''' <param name="moduleName">The name of the module's section</param>
    '''
    ''' <param name="settingName">The name of the setting's key</param>
    '''
    ''' <param name="value">The value to store</param>
    Public Sub SetSetting(moduleName As String,
                         settingName As String,
                         value As String)

        Dim section = SettingsFile.GetOrCreateSection(moduleName)
        Dim key = section.Keys.GetKey(settingName)

        If key Is Nothing Then

            section.AddKey(New iniKey($"{settingName}={value}"))

        Else

            key.Value = value

        End If

        _dirty = True

    End Sub

    ''' <summary>
    ''' Writes the whole of <see cref="SettingsFile"/> to disk and clears the dirty flag, subject
    ''' to <paramref name="condition"/> and to the global save gate
    ''' (<see cref="saveSettingsToDisk"/>, and never during a command line run). We clear the
    ''' flag even when the write itself fails.
    ''' <br />
    ''' The gate lives here rather than at the call sites because <see cref="FlushIfDirty"/> is also
    ''' invoked ungated whenever a menu or the application closes. Any <see cref="SetSetting"/> caller
    ''' which neglects to gate its own flush would otherwise have its changes persisted by one of
    ''' those, writing settings the user asked not to save.
    ''' </summary>
    '''
    ''' <param name="condition">
    ''' An additional condition which must hold for the write to occur <br /><br />
    ''' Optional, Default: <c> True </c>
    ''' </param>
    Public Sub SaveSettings(Optional condition As Boolean = True)

        If Not condition OrElse IsCommandLineMode OrElse Not saveSettingsToDisk Then

            gLog("Settings save skipped - saving to disk is disabled")
            Return

        End If

        SettingsFile.OverwriteToFile(SettingsFile.ToString())
        _dirty = False

    End Sub

    ''' <summary>
    ''' Calls <see cref="SaveSettings"/> only if <see cref="SettingsFile"/> has been modified
    ''' since the last save, so the same gate applies.
    ''' <br />
    ''' A flush suppressed by the save gate leaves the backend dirty, so enabling
    ''' <see cref="saveSettingsToDisk"/> later in the session still persists the changes made before it
    ''' was turned on.
    ''' </summary>
    '''
    ''' <param name="condition">
    ''' An additional condition which must hold for the write to occur <br /><br />
    ''' Optional, Default: <c> True </c>
    ''' </param>
    Public Sub FlushIfDirty(Optional condition As Boolean = True)

        If _dirty Then SaveSettings(condition)

    End Sub

    ''' <summary>
    ''' Populates a module's writable public properties from its section of
    ''' <see cref="SettingsFile"/>, and does nothing if the section is missing.
    ''' <br />
    ''' Handles <c> Boolean </c>, <c> Enum </c>, and <c> iniFileChooser </c> property types.
    ''' A property whose key is absent keeps its current value, and so does a
    ''' <c> Boolean </c> whose value doesn't parse. Any other property type with a key present,
    ''' or an enum value that isn't a member name or number, throws, and we don't catch it.
    ''' </summary>
    '''
    ''' <param name="moduleName">The name of the module's section</param>
    '''
    ''' <param name="moduleType">The settings module whose properties receive the values</param>
    Public Sub LoadModule(moduleName As String, moduleType As Type)

        Dim section = SettingsFile.GetSection(moduleName)
        If section Is Nothing Then Return

        Using gLogScope($"Loading settings for {moduleName}")

            gLog("")

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
                    gLog($"{prop.Name}'s parameters successfully read from disk")
                    Continue For

                End If

                Dim k = section.Keys.GetKey(prop.Name)
                If k Is Nothing Then gLog($"Could not load {prop.Name} from disk") : Continue For

                If prop.PropertyType Is GetType(Boolean) Then

                    Dim bVal As Boolean
                    If Boolean.TryParse(k.Value, bVal) Then prop.SetValue(Nothing, bVal)

                Else

                    prop.SetValue(Nothing, [Enum].Parse(prop.PropertyType, k.Value))

                End If

                gLog($"{prop.Name}'s value successfully read from disk")

            Next

        End Using

        gLog("")
    End Sub

    ''' <summary>
    ''' Copies a module's readable and writable public properties into
    ''' <see cref="SettingsFile"/> through <see cref="SetSetting"/>. This only changes memory:
    ''' the values reach disk when a later <see cref="FlushIfDirty"/> gets past the save gate,
    ''' which is off by default and never opens during a command line run.
    ''' <br />
    ''' Handles <c> Boolean </c>, <c> Enum </c>, and <c> iniFileChooser </c> property types.
    ''' Logs a warning and skips properties of any other type, and skips any property whose
    ''' value is <c> Nothing </c>.
    ''' </summary>
    '''
    ''' <param name="moduleName">The name of the module's section</param>
    '''
    ''' <param name="moduleType">The settings module whose properties we save</param>
    Public Sub SaveModule(moduleName As String, moduleType As Type)

        For Each prop As PropertyInfo In moduleType.GetProperties()

            If Not prop.CanRead OrElse Not prop.CanWrite Then Continue For

            Dim value = prop.GetValue(Nothing)
            If value Is Nothing Then Continue For

            If prop.PropertyType Is GetType(iniFileChooser) Then

                Dim chooser = DirectCast(value, iniFileChooser)
                SetSetting(moduleName, prop.Name & "_Name", chooser.Name)
                SetSetting(moduleName, prop.Name & "_Dir", chooser.Dir)
                Continue For

            End If

            If prop.PropertyType IsNot GetType(Boolean) AndAlso Not prop.PropertyType.IsEnum Then

                gLog($"SaveModule: unhandled property type '{prop.PropertyType.Name}' for '{prop.Name}' in {moduleName}. skipping.")
                Continue For

            End If

            SetSetting(moduleName, prop.Name, value.ToString())

        Next

    End Sub

End Module
