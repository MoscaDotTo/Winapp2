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
''' The sole settings backend, powered by <c>iniFile2</c>.
''' <br />
''' <c>SettingsFile2</c> is the single authoritative in-memory representation of
''' <c>winapp2ool.ini</c>. All modules are registered in <c>loadAllModuleSettings</c>
''' and use <c>LoadModule2</c> / <c>SaveModule2</c> for persistence.
''' </summary>
Public Module SettingsHandler2

    Private _dirty2 As Boolean = False

    ''' <summary>
    ''' The <c> iniFile2 </c>-backed representation of winapp2ool's settings.
    ''' </summary>
    Public Property SettingsFile2 As iniFile2 = iniFile2.Empty(Environment.CurrentDirectory, "winapp2ool.ini")

    ''' <summary>
    ''' Reads <c>winapp2ool.ini</c> from disk into <c> SettingsFile2 </c>.
    ''' </summary>
    Public Sub LoadWinapp2oolsettings()

        Using gLogScope("Loading settings")

            ' Handle the default case where winapp2ool.ini doesn't exist
            If Not System.IO.File.Exists(SettingsFile2.Path) Then

                readSettingsFromDisk = False
                saveSettingsToDisk = False

                ' We still need to maintain an internal representation
                ' of the settings so create the settingsFile and settingsDict
                ' using the default winapp2ool configuration
                gLog("No settings file Found - loading default settings")
                loadAllModuleSettings()
                Return

            End If

            SettingsFile2 = iniFile2.FromFile(SettingsFile2.Path())
            loadAllModuleSettings()

        End Using

        gLog("Settings loaded", buffr:=True)

    End Sub

    ''' <summary>
    ''' Loads settings for modules that have migrated to <c>SettingsHandler2</c>.
    ''' Each migrated module's <c>LoadModule2</c> call is added here as modules migrate.
    ''' </summary>
    Private Sub loadAllModuleSettings()

        ' Winapp2ool is loaded first so readSettingsFromDisk is populated before other modules load
        LoadModule2(NameOf(Winapp2ool), GetType(maintoolsettings))

        If Not readSettingsFromDisk Then Return

        LoadModule2(NameOf(Diff), GetType(diffsettings))
        LoadModule2(NameOf(UWPBuilder), GetType(uwpbuildersettings))
        LoadModule2(NameOf(EntryBuilder), GetType(entryBuilderSettings))
        LoadModule2(NameOf(BrowserBuilder), GetType(browserbuildersettings))
        LoadModule2(NameOf(CC7Patcher), GetType(cc7patchersettings))
        LoadModule2(NameOf(CCiniDebug), GetType(ccdebugsettings))
        LoadModule2(NameOf(Combine), GetType(combinesettings))
        LoadModule2(NameOf(Flavorizer), GetType(FlavorizerSettings))
        LoadModule2(NameOf(Transmute), GetType(transmuteSettings))
        LoadModule2(NameOf(Trim), GetType(trimsettings))
        LoadModule2(NameOf(Downloader), GetType(downloadersettings))
        LoadModule2(NameOf(WinappDebug), GetType(lintsettings))
        LoadLintRulesFromSettings2()

    End Sub

    ''' <summary>
    ''' Returns the value of a setting from <c>SettingsFile2</c>,
    ''' or <c>""</c> if the module section or key is not found.
    ''' </summary>
    Public Function GetSetting(moduleName As String,
                               settingName As String) As String

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
    Public Sub SetSetting(moduleName As String,
                         settingName As String,
                         value As String)

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
    ''' Writes <c>SettingsFile2</c> to disk, subject to <paramref name="condition"/> and to the
    ''' global save gate (<c> saveSettingsToDisk </c>, and never during a command line run).
    ''' <br />
    ''' The gate lives here rather than at the call sites because <c>FlushIfDirty2</c> is also
    ''' invoked ungated whenever a menu or the application closes. Any <c>SetSetting</c> caller
    ''' which neglects to gate its own flush would otherwise have its changes persisted by one of
    ''' those, writing settings the user asked not to save.
    ''' </summary>
    '''
    ''' <param name="condition">
    ''' An additional condition which must hold for the write to occur
    ''' <br /> Optional, Default: <c> True </c>
    ''' </param>
    Public Sub SaveSettings(Optional condition As Boolean = True)

        If Not condition OrElse IsCommandLineMode OrElse Not saveSettingsToDisk Then

            gLog("Settings save skipped - saving to disk is disabled")
            Return

        End If

        SettingsFile2.OverwriteToFile(SettingsFile2.ToString())
        _dirty2 = False

    End Sub

    ''' <summary>
    ''' Writes <c>SettingsFile2</c> to disk only if it has been modified since the last save.
    ''' <br />
    ''' A flush suppressed by the save gate leaves the backend dirty, so enabling
    ''' <c> saveSettingsToDisk </c> later in the session still persists the changes made before it
    ''' was turned on.
    ''' </summary>
    '''
    ''' <param name="condition">
    ''' An additional condition which must hold for the write to occur
    ''' <br /> Optional, Default: <c> True </c>
    ''' </param>
    Public Sub FlushIfDirty2(Optional condition As Boolean = True)

        If _dirty2 Then SaveSettings(condition)

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
    ''' Writes a module's public static properties into <c>SettingsFile2</c>.
    ''' <br />
    ''' Handles <c>Boolean</c>, <c>Enum</c>, and <c>iniFileChooser</c> property types.
    ''' Logs a warning and skips properties of any other type.
    ''' </summary>
    Public Sub SaveModule2(moduleName As String, moduleType As Type)

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

                gLog($"SaveModule2: unhandled property type '{prop.PropertyType.Name}' for '{prop.Name}' in {moduleName}. skipping.")
                Continue For

            End If

            SetSetting(moduleName, prop.Name, value.ToString())

        Next

    End Sub

End Module
