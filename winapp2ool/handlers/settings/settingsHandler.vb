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
Imports System.IO
Imports System.Reflection

''' <summary> 
''' settingsHandler manages the disk-writable representation of winapp2ool
''' module properties for persistence between sessions by deserializing the 
''' <c> winapp2ool.ini </c> file into their respective module settings,
''' monitoring any changes during the session and serializing them back to disk
''' as needed. <br /> 
''' A real-time copy of all disk-writable settings is stored in a 
''' <c> Dictionary (Of String, Dictionary (Of String, String)) </c> where the 
''' primary key is the module name and the secondary key is the setting name
''' 
''' <br /> The settings are also stored in an <c> iniFile </c>, updated 
''' alongside the dictionary, for on-disk formatting purposes and 
''' can be read from or written to disk as needed
''' </summary>
''' 
''' Docs last updated: 2025-06-25 | Code last updated 2025-06-25
Public Module settingsHandler

    ''' <summary> 
    ''' A copy of winapp2ool's current set of user configurable settings that
    ''' can be read from or written to disk in a familiar ini format <br /> 
    ''' 
    ''' If <c> saveSettingsToDisk </c> is <c> True </c>, then this file will 
    ''' be overwritten on disk every time a setting is updated 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-06-25 | Code last updated: 2025-06-25
    Public Property settingsFile As New iniFile(Environment.CurrentDirectory, "winapp2ool.ini")

    ''' <summary> 
    ''' A cleaner structure for accessing module settings which can be easily
    ''' written to <c> settingsFile </c> <br /> 
    ''' 
    ''' Keys in the primary dictionary are module names, the value dictionaries
    ''' are simple key value pairs where the key is the name of a setting 
    ''' and the value is the value of the setting
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-06-25 | Code last updated: 2025-06-25
    Public Property settingsDict As New Dictionary(Of String, Dictionary(Of String, String))

    ''' <summary> 
    ''' If saving is enabled, saves the current state of winapp2ool's
    ''' settings to disk, overwriting any existing settings 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-06-25 | Code last updated: 2025-06-25
    Public Sub saveSettingsFile()

        settingsFile.overwriteToFile(settingsFile.toString, saveSettingsToDisk)

    End Sub

    ''' <summary> 
    ''' Keeps the on-disk representation of the application's
    ''' settings in sync with the current session's settings
    ''' </summary>
    ''' 
    ''' <param name="targetModule">
    ''' The name of the module who owns the setting being updated 
    ''' </param>
    ''' 
    ''' <param name="settingName"> 
    ''' The name of the setting being updated
    ''' </param>
    ''' 
    ''' <param name="newVal"> 
    ''' The updated value held by the setting 
    ''' </param>
    ''' 
    ''' Docs last updated: 2025-06-25 | Code last updated: 2025-06-25
    Public Sub updateSettings(targetModule As String, settingName As String, newVal As String)

        settingsDict(targetModule)(settingName) = newVal

        For Each key In settingsFile.getSection(targetModule).Keys.Keys

            If Not key.Name = settingName Then Continue For

            key.Value = newVal
            Exit For

        Next

        saveSettingsFile()

    End Sub

    ''' <summary> 
    ''' Adds the <c> Keys </c> from a given <c> iniSection </c> 
    ''' to the disk-writable copy of winapp2ool's settings
    ''' for the <c> <paramref name="targetModule"/> </c>
    ''' </summary>
    ''' 
    ''' <param name="targetModule">
    ''' The module whose settings should be updated 
    ''' </param>
    ''' 
    ''' <param name="settings">
    ''' The <c> iniSection </c> containing <c> iniKey </c> 
    ''' format serialized winapp2ool settings 
    ''' </param>
    ''' 
    ''' Docs last updated: 2025-06-25 | Code last updated: 2025-06-25

    Private Sub addToSettingsDict(targetModule As String, settings As iniSection)

        Dim subDict As New Dictionary(Of String, String)
        settings.Keys.Keys.ForEach(Sub(key) If Not subDict.ContainsKey(key.Name) Then subDict.Add(key.Name, key.Value))
        settingsDict(targetModule) = subDict

    End Sub

    ''' <summary> 
    ''' Looks for the <c> settingsFile </c> 
    ''' and attempts to load it from disk. If the <c> settingsFile </c>
    ''' is empty, settings will be loaded from their default configuration 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-06-25 | Code last updated: 2025-06-25
    Public Sub loadSettings()

        gLog("Loading settings")
        gLog(ascend:=True)

        ' Handle the default case where winapp2ool.ini doesn't exist  
        If Not File.Exists(settingsFile.Path) Then

            readSettingsFromDisk = False
            saveSettingsToDisk = False

            ' We still need to maintain an internal representation
            ' of the settings so create the settingsFile and settingsDict
            ' using the default winapp2ool configuration
            gLog("No settings file Found - loading default settings")
            loadAllModuleSettings()

            Return

        End If

        ' Otherwise, try to load the file. If it's empty,
        ' lets assume the user wanted to reset their settings to default
        settingsFile.init()
        loadAllModuleSettings()

        gLog(descend:=True)
        gLog("Settings loaded")

    End Sub

    ''' <summary>
    ''' Loads all module settings into disk-writable representation
    ''' in the form of both an <c> iniFile </c> and a <c> Dictionary </c> 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-06-25 | Code last updated: 2025-06-25
    Private Sub loadAllModuleSettings()

        serializeModuleSettings(NameOf(Winapp2ool),
                                AddressOf createToolSettingsSection,
                                AddressOf getSeralizedToolSettings)

        serializeModuleSettings(NameOf(WinappDebug),
                                AddressOf CreateLintSettingsSection,
                                AddressOf getSeralizedLintSettings)

        serializeModuleSettings(NameOf(Trim),
                                AddressOf createTrimSettingsSection,
                                AddressOf getSerializedTrimSettings)

        serializeModuleSettings(NameOf(Merge),
                                AddressOf createMergeSettingsSection,
                                AddressOf getSeralizedMergeSettings)

        serializeModuleSettings(NameOf(Diff),
                                AddressOf CreateDiffSettingsSection,
                                AddressOf GetSerializedDiffSettings)

        serializeModuleSettings(NameOf(CCiniDebug),
                                AddressOf createCCDBSettingsSection,
                                AddressOf getSerializedDebugSettings)

        serializeModuleSettings(NameOf(Downloader),
                                AddressOf createDownloadSettingsSection,
                                AddressOf getSerializedDownloaderSettings)

    End Sub


    ''' <summary> 
    ''' Clears the settings belonging to the <c> <paramref name="moduleName"/>
    ''' </c> given from the settingsDict and from the respective
    ''' <c> iniSection </c> in the settingsFile 
    ''' </summary>
    ''' 
    ''' <param name="moduleName"> 
    ''' The name of the module whose settings are being cleared 
    ''' </param>
    ''' 
    ''' Docs last updated: 2025-06-25 | Code last updated: 2025-06-25
    Public Sub clearAllModuleSettings(moduleName As String)

        gLog($"Clearing {moduleName}'s settings")

        settingsDict(moduleName).Clear()
        settingsFile.getSection(moduleName).Keys.Keys.Clear()

    End Sub

    ''' <summary> 
    ''' Creates the disk-writable representation
    ''' of a module's settings, either by generating
    ''' their default values or by assigning values 
    ''' given from a previous session to the current values 
    ''' </summary>
    ''' 
    ''' <param name="moduleName"> 
    ''' The winapp2ool module whose settings 
    ''' will be either instantiated or read from disk 
    ''' </param>
    ''' 
    ''' <param name="createModuleSettings"> 
    ''' The function who creates the in-memory representation of the module's
    ''' setting in the absence of a version on disk from which to read them
    ''' </param>
    ''' 
    ''' <param name="getSerializedSettings"> 
    ''' The function who assigns values from disk to
    ''' the current active session of a module's settings 
    ''' </param>
    ''' 
    ''' Docs last updated: 2025-06-25 | Code last updated: 2025-06-25
    Private Sub serializeModuleSettings(moduleName As String,
                                        createModuleSettings As Action,
                                        getSerializedSettings As Action)

        Dim rdSttngsName = NameOf(readSettingsFromDisk)
        Dim moduleSection = settingsFile.getSection(moduleName)

        addToSettingsDict(moduleName, moduleSection)

        If settingsDict(NameOf(Winapp2ool)).ContainsKey(rdSttngsName) Then

            If CBool(settingsDict(NameOf(Winapp2ool))(rdSttngsName)) Then

                getSerializedSettings()

            Else

                ' If we're serializing the root module's settings and it tells us not to read from disk, load the defaults
                If moduleName = NameOf(Winapp2ool) Then clearAllModuleSettings(NameOf(Winapp2ool))

            End If

        Else

            ' If we don't know if we're reading settings from disk, we probably just shouldn't. 
            ' clear the settings inifile And the settingsdict for winapp2ool and load the defaults 
            ' Unless something goes terribly wrong, we won't hit this path again after initalizing the root winapp2ool module 
            clearAllModuleSettings(NameOf(Winapp2ool))

        End If

        createModuleSettings()

    End Sub

    ''' <summary> 
    ''' Creates the <c> iniSection </c> containing the default values
    ''' for variables associated with the calling module and adds 
    ''' them to the dictionary representation of the settings
    ''' </summary>
    ''' 
    ''' <param name="callingModule"> 
    ''' The name of the module whose settings are being created
    ''' </param>
    ''' 
    ''' <param name="settingTuples"> 
    ''' The array of Strings containing the
    ''' module's settings and their default values
    ''' </param>
    ''' 
    ''' <param name="numBools"> 
    ''' The number of boolean settings in the module 
    ''' </param>
    ''' 
    ''' <param name="numFiles"> 
    ''' The number of <c> iniFile </c> settings in the module 
    ''' </param>
    ''' 
    ''' <remarks>
    ''' The <c> settingTuples </c> should be formatted as follows: 
    ''' (Name, Value) pairs for <c> Boolean </c>settings and 
    ''' (Name, Filename, Dir) triplets for <c> iniFile </c> settings
    ''' </remarks>
    ''' 
    ''' Docs last updated: 2025-06-25 | Code last updated: 2025-06-25
    Public Sub createModuleSettingsSection(callingModule As String,
                                           settingTuples As List(Of String),
                                           numBools As Integer,
                                           numFiles As Integer)

        If settingTuples Is Nothing Then argIsNull(NameOf(settingTuples)) : Return
        Dim settingKeys As List(Of String)

        settingKeys = getSettingKeys(settingTuples, callingModule, numBools, numFiles)

        Dim toolSection = settingsFile.getSection(callingModule)
        Dim mustSaveFile = False

        For Each key In settingKeys

            If key.Length = 0 Then Continue For
            toolSection.Keys.add(New iniKey(key))
            mustSaveFile = True

        Next

        If Not settingsFile.hasSection(callingModule) Then settingsFile.Sections.Add(callingModule, toolSection) : mustSaveFile = True
        addToSettingsDict(callingModule, toolSection)
        If mustSaveFile Then saveSettingsFile()

    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="settingsTuples"> 
    ''' A set of module properties to process into iniKeys
    ''' </param>
    ''' 
    ''' <param name="moduleName">
    ''' The name of the module whose properties will be returned as iniKeys 
    ''' </param>
    ''' 
    ''' <param name="numBools"> 
    ''' The number of boolean settings in the module 
    ''' </param>
    ''' 
    ''' <param name="numFiles"> 
    ''' The number of <c> iniFile </c> settings in the module 
    ''' </param>
    ''' 
    ''' <returns>
    ''' A list of iniKeys representing the settings
    ''' for the specified module to be written to disk
    ''' </returns>
    ''' 
    ''' Docs last updated: 2025-06-25 | Code last updated: 2025-06-25
    Private Function getSettingKeys(settingsTuples As List(Of String),
                                    moduleName As String,
                                    numBools As Integer,
                                    numFiles As Integer) As List(Of String)

        ' Ensure that we only operate on properly formatted tuples 
        Dim expectedSettingsCount = 2 * numBools + 3 * numFiles

        Dim out As New List(Of String)

        For i = 0 To settingsTuples.Count - 2

            If i >= 3 * numFiles - 1 Then

                ' boolean settings are in (name,value) pairs
                out.Add(getSettingIniKey(moduleName, settingsTuples(i), settingsTuples(i + 1)))
                i += 1

            Else

                ' file settings are in (name, filename, dir) triplets 
                out.Add(getSettingIniKey(moduleName, settingsTuples(i), settingsTuples(i + 1), isName:=True))
                out.Add(getSettingIniKey(moduleName, settingsTuples(i), settingsTuples(i + 2), isDir:=True))
                i += 2

            End If

        Next

        gLog($"Finished loading settings for {moduleName}")

        Return out
    End Function

    ''' <summary> 
    ''' Returns the text that will comprise the <c> iniKey </c>
    ''' representation of the given <c> <paramref name="settingName"/> </c> 
    ''' iff it doesn't already exist in the <c> settingDict </c>
    ''' <br /> Returns <c> "" </c> otherwise 
    ''' </summary>
    ''' 
    ''' <param name="moduleName"> 
    ''' The name of the module whose settings are being queried 
    ''' </param>
    ''' 
    ''' <param name="settingName"> 
    ''' The name of the setting being queried as it will appear on disk 
    ''' </param>
    ''' 
    ''' <param name="settingValue"> 
    ''' The current value of the setting as a String 
    ''' </param>
    ''' 
    ''' <param name="isName"> 
    ''' Indicates that the <c> <paramref name="settingName"/> </c>
    ''' is an <c> iniFile's Name</c> property 
    ''' <br/> Optional, Default <c> False </c> 
    ''' </param>
    ''' 
    ''' <param name="isDir"> 
    ''' Indicates that the <c> <paramref name="settingName"/> </c> 
    ''' is an <c> iniFile's Dir </c> property 
    ''' <br/> Optional, Default <c> False </c>
    ''' </param>
    ''' 
    ''' <returns>
    ''' A string representing the iniKey or an empty 
    ''' string if the setting already exists in the dictionary
    ''' </returns>
    ''' 
    ''' Docs last updated: 2025-06-25 | Code last updated: 2025-06-25
    Private Function getSettingIniKey(moduleName As String,
                                      settingName As String,
                                      settingValue As String,
                                      Optional isName As Boolean = False,
                                      Optional isDir As Boolean = False) As String

        If isName Then settingName += "_Name"
        If isDir Then settingName += "_Dir"

        Dim settingInDict = settingsDict(moduleName).ContainsKey(settingName)

        gLog($"{settingName} was not found in {settingsFile.Name}", Not settingInDict, indent:=True)

        Return If(Not settingInDict, $"{settingName}={settingValue}", "")

    End Function

    ''' <summary> 
    ''' Clears the current module setting from the disk-writable representation
    '''  of the <c> <paramref name="callingModule"/>'s </c> module settings 
    ''' </summary>
    ''' 
    ''' <param name="callingModule">
    ''' The name of the module whose settings will be cleared 
    ''' </param>
    ''' 
    ''' <param name="createSection"> 
    ''' The subroutine that creates the module's disk-writable settings 
    ''' </param>
    ''' 
    ''' Docs last updated: 2025-06-25 | Code last updated: 2025-06-25
    Public Sub restoreDefaultSettings(callingModule As String,
                                      createSection As Action)

        If createSection Is Nothing Then argIsNull(NameOf(createSection)) : Return

        clearAllModuleSettings(callingModule)
        createSection()

        saveSettingsFile()

    End Sub

    ''' <summary>
    ''' Assigns a module's properties for the current 
    ''' session of a winapp2ool module by reading them 
    ''' from the disk-writable representation of the settings
    ''' </summary>
    ''' 
    ''' <param name="moduleName"> 
    ''' The name of the module whose properties will be assigned 
    ''' </param>
    ''' 
    ''' <param name="winapp2oolmodule"> 
    ''' A module containing the set of disk-writable properties for 
    ''' <c> <paramref name="moduleName"/> </c>
    ''' </param>
    ''' 
    ''' Docs last updated: 2025-06-25 | Code last updated: 2025-06-25
    Public Sub LoadModuleSettingsFromDict(moduleName As String,
                                          winapp2oolmodule As Type)

        If Not moduleName = "Winapp2ool" AndAlso Not readSettingsFromDisk Then Return

        If Not settingsDict.ContainsKey(moduleName) Then Return
        Dim dict = settingsDict(moduleName)

        For Each prop As PropertyInfo In winapp2oolmodule.GetProperties()

            If Not prop.CanWrite Then Continue For

            Dim value2 = prop.GetValue(Nothing)

            If value2.GetType().Name = "iniFile" Then

                If Not dict.ContainsKey(prop.Name & "_Name") OrElse Not dict.ContainsKey(prop.Name & "_Dir") Then Continue For

                Dim iniProp As iniFile = CType(value2, iniFile)

                iniProp.Name = dict(prop.Name & "_Name")
                iniProp.Dir = dict(prop.Name & "_Dir")

                Continue For

            End If

            If Not dict.ContainsKey(prop.Name) Then Continue For

            Dim valueStr = dict(prop.Name)
            Dim value = Convert.ChangeType(valueStr, prop.PropertyType, CultureInfo.InvariantCulture)

            prop.SetValue(winapp2oolmodule, value)

        Next

    End Sub

    ''' <summary>
    ''' Returns a list of strings representing the settings
    ''' of a module, uses reflection to access these properties 
    ''' <br /> tuple is comprised of
    ''' (Name, Value) pairs for <c> Boolean </c> properties and
    ''' (Name, Filename, Dir) triplets for <c> iniFile </c> properties
    ''' </summary>
    ''' 
    ''' <param name="moduleType"> 
    ''' The type of the module whose settings will be retrieved
    ''' </param>
    ''' 
    ''' <returns>
    ''' A <c> List (Of String) </c> containing
    ''' the disk-writable key/value pair names
    ''' </returns>
    ''' 
    ''' Docs last updated: 2025-06-25 | Code last updated: 2025-06-25
    Public Function GetSettingsTupleWithReflection(moduleType As Type) As List(Of String)

        Dim tuple As New List(Of String)

        For Each prop As PropertyInfo In moduleType.GetProperties()

            If Not prop.CanRead OrElse Not prop.CanWrite Then Continue For

            Dim value = prop.GetValue(Nothing)
            If value Is Nothing Then Continue For

            ' Boolean Properties
            If Not value.GetType().Name = "iniFile" Then

                tuple.Add(prop.Name)
                tuple.Add(value.ToString())
                Continue For

            End If

            ' iniFile properties
            Dim ini As iniFile = CType(value, iniFile)
            tuple.Add(prop.Name)
            tuple.Add($"{ini.Name}")
            tuple.Add($"{ini.Dir}")

        Next

        Return tuple

    End Function

    ''' <summary>
    ''' Returns the number of <c> iniFile </c>
    ''' properties in a given winapp2ool module type
    ''' </summary>
    ''' 
    ''' <param name="winapp2oolModule">
    ''' The winapp2ool Module whose properties are being counted
    ''' </param>
    ''' 
    ''' <returns> 
    ''' The number of <c> iniFile> </c>
    ''' properties in the specified module type
    ''' </returns>
    ''' 
    ''' Docs last updated: 2025-06-25 | Code last updated: 2025-06-25
    Public Function getNumFiles(winapp2oolModule As Type) As Integer

        Dim numFiles = 0

        For Each prop As PropertyInfo In winapp2oolModule.GetProperties()

            If prop.PropertyType Is GetType(iniFile) Then numFiles += 1

        Next

        Return numFiles

    End Function

    ''' <summary>
    ''' Returns the number of <c> Boolean </c> 
    ''' properties in a given winapp2ool module type
    ''' </summary>
    ''' 
    ''' <param name="winapp2oolModule">
    ''' The winapp2ool Module whose properties are being counted
    ''' </param>
    ''' 
    ''' <returns> 
    ''' The number of <c> Boolean </c> properties in the specified module type
    ''' </returns>
    ''' 
    ''' Docs last updated: 2025-06-25 | Code last updated: 2025-06-25
    Public Function getNumBools(winapp2oolModule As Type) As Integer

        Dim numBools = 0

        For Each prop As PropertyInfo In winapp2oolModule.GetProperties()

            If prop.PropertyType Is GetType(Boolean) Then numBools += 1

        Next

        Return numBools

    End Function

End Module