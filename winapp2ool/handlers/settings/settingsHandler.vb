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
Imports System.IO
''' <summary> settingsHandler provides an interface for winapp2ool modules to serialize some of their settings to disk using a familiar ini interface </summary>
Public Module settingsHandler
    ''' <summary> A copy of winapp2ool's user configurable settings that was read from disk <br/> 
    ''' Its contents overwritten with the contents of the settingsDict whenever settings are changed </summary>
    Public Property settingsFile As New iniFile(Environment.CurrentDirectory, "winapp2ool.ini")
    ''' <summary> A more nicely accessible structure for accessing module settings <br /> 
    ''' Keys in the primary dictionary are module names, the value dictionaries are simple key value pairs where 
    ''' the key is the name of a setting and the value is the state of the setting </summary>
    Public Property settingsDict As New Dictionary(Of String, Dictionary(Of String, String))
    '''<summary> Indicates that changes to the application's settings should be serialized back to the disk </summary>
    Public Property saveSettingsToDisk As Boolean = False
    ''' <summary> Indicates that settings who are read from the disk should override the corresponding default module settings </summary>
    Public Property readSettingsFromDisk As Boolean = False

    ''' <summary> Saves the winapp2ool settings to disk </summary>
    Public Sub saveSettingsFile()
        settingsFile.overwriteToFile(settingsFile.toString, saveSettingsToDisk)
    End Sub

    ''' <summary> Updates the disk-writable representations of the application's settings as the user makes changes to them </summary>
    ''' <param name="targetModule"> The name of the module who owns the setting being updated </param>
    ''' <param name="settingName"> The name of the setting being updated </param>
    ''' <param name="newVal"> The updated value held by the setting </param>
    Public Sub updateSettings(targetModule As String, settingName As String, newVal As String)
        settingsDict(targetModule)(settingName) = newVal
        For Each key In settingsFile.getSection(targetModule).Keys.Keys
            If key.Name = settingName Then key.Value = newVal : Exit For
        Next
        saveSettingsFile()
    End Sub

    ''' <summary> Adds the <c> Keys </c> from a given <c> iniSection </c> to the disk-writable copy of winapp2ool's settings for the <c> <paramref name="targetModule"/> </c></summary>
    ''' <param name="targetModule"> The module whose settings should be updated </param>
    ''' <param name="settings"> The <c> iniSection </c> containing <c> iniKey </c> format serialized winapp2ool settings </param>
    Private Sub addToSettingsDict(targetModule As String, settings As iniSection)
        Dim subDict As New Dictionary(Of String, String)
        settings.Keys.Keys.ForEach(Sub(key) subDict.Add(key.Name, key.Value))
        settingsDict(targetModule) = subDict
    End Sub

    ''' <summary> Looks for the <c> settingsFile </c> and attempts to load it from disk. If the <c> settingsFile </c> 
    ''' is empty, settings will be loaded from their default configuration </summary>
    Public Sub loadSettings()
        gLog("Loading settings")
        gLog(ascend:=True)
        ' Handle the default case where winapp2ool.ini doesn't exist  
        If Not File.Exists(settingsFile.Path) Then
            readSettingsFromDisk = False
            saveSettingsToDisk = False
            ' We still need to maintain an internal representation of the settings, so create the settingsfile and settingsdict using the default winapp2ool configuration
            gLog("No settings file Found - loading default settings")
            loadAllModuleSettings()
            Return
        End If
        ' Otherwise, try to load the file. If it's empty, lets assume the user wanted to reset their settings to default
        settingsFile.init()
        loadAllModuleSettings()
        gLog(descend:=True)
        gLog("Settings loaded")
    End Sub

    ''' <summary> Loads all module settings into disk-writable representation in the form of both an <c> iniFile </c> and a <c> Dictionary </c> </summary>
    Private Sub loadAllModuleSettings()
        serializeModuleSettings(NameOf(Winapp2ool), AddressOf createToolSettingsSection, AddressOf getSeralizedToolSettings)
        serializeModuleSettings(NameOf(WinappDebug), AddressOf createLintSettingsSection, AddressOf getSeralizedLintSettings)
        serializeModuleSettings(NameOf(Trim), AddressOf createTrimSettingsSection, AddressOf getSerializedTrimSettings)
        serializeModuleSettings(NameOf(Merge), AddressOf createMergeSettingsSection, AddressOf getSeralizedMergeSettings)
        serializeModuleSettings(NameOf(Diff), AddressOf createDiffSettingsSection, AddressOf getSerializedDiffSettings)
        serializeModuleSettings(NameOf(CCiniDebug), AddressOf createDebugSettingsSection, AddressOf getSerializedDebugSettings)
        serializeModuleSettings(NameOf(Downloader), AddressOf createDownloadSettingsSection, AddressOf getSerializedDownloaderSettings)
    End Sub


    ''' <summary> Clears the settings belonging to the <c> <paramref name="moduleName"/> </c> given from the settingsDict and from the respective <c> iniSection </c> in the settingsFile </summary>
    ''' <param name="moduleName"> The name of the module whose settings are being cleared </param>
    Public Sub clearAllModuleSettings(moduleName As String)
        gLog($"Clearing {moduleName}'s settings")
        settingsDict(moduleName).Clear()
        settingsFile.getSection(moduleName).Keys.Keys.Clear()
    End Sub

    ''' <summary> Creates the disk-writable representation of a module's settings, either by generating their default values or by assigning 
    ''' values given from a previous session to the current values </summary>
    ''' <param name="moduleName"> The winapp2ool module whose settings will be either instantiated or read from disk </param>
    ''' <param name="createModuleSettings"> The function who creates the in-memory representation of the module's setting in the absence of a
    ''' version on disk from which to read them </param>
    ''' <param name="getSerializedSettings"> The function who assigns values from disk to the current active session of a module's settings </param>
    Private Sub serializeModuleSettings(moduleName As String, createModuleSettings As Action, getSerializedSettings As Action)
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

    ''' <summary> Creates the <c> iniSection </c> containing the default values for variables associated with the calling module
    ''' and adds them to the dictionary representation of the settings </summary>
    ''' <param name="callingModule"> The name of the module whose settings are being created </param>
    ''' <param name="settingTuples"> The array of Strings containing the module's settings and their default values </param>
    Public Sub createModuleSettingsSection(callingModule As String, settingTuples As List(Of String), numBools As Integer, Optional numFiles As Integer = 3)
        If settingTuples Is Nothing Then argIsNull(NameOf(settingTuples)) : Return
        Dim settingKeys = getSettingKeys(settingTuples, callingModule, numBools, numFiles)
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

    Private Function getSettingKeys(settingsTuples As List(Of String), moduleName As String, numBools As Integer, Optional numFiles As Integer = 3) As List(Of String)
        ' Ensure that we only operate on properly formatted tuples 
        gLog($"The number of settings provided to {moduleName}'s settings initializer doesn't match the number expected.", Not settingsTuples.Count = 2 * numBools + 3 * numFiles)
        Dim out As New List(Of String)
        For i = 0 To settingsTuples.Count - 3
            If i < 2 * numBools Then
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
        Return out
    End Function

    ''' <summary> Returns the text that will comprise the <c> iniKey </c> representation of the given <c> <paramref name="settingName"/> </c> iff it doesn't already
    ''' exist in the <c> settingDict </c> <br /> Returns <c> "" </c> otherwise </summary>
    ''' <param name="moduleName"> The name of the module whose settings are being queried </param>
    ''' <param name="settingName"> The name of the setting being queried as it will appear on disk </param>
    ''' <param name="settingValue"> The current value of the setting as a String </param>
    ''' <param name="isName"> Indicates that the <c> <paramref name="settingName"/> </c> is an <c> iniFile's Name</c> property <br/> Optional, Default <c> False </c> </param>
    ''' <param name="isDir"> Indicates that the <c> <paramref name="settingName"/> </c> is an <c> iniFile's Dir </c> property <br/> Optional, Default <c> False </c></param>
    ''' <returns></returns>
    Private Function getSettingIniKey(moduleName As String, settingName As String, settingValue As String, Optional isName As Boolean = False, Optional isDir As Boolean = False) As String
        If isName Then settingName += "_Name"
        If isDir Then settingName += "_Dir"
        Dim settingInDict = settingsDict(moduleName).ContainsKey(settingName)
        gLog($"{settingName} was not found in the settings dictionary", Not settingInDict, indent:=True)
        Return If(Not settingInDict, $"{settingName}={settingValue}", "")
    End Function

    ''' <summary> Clears the current module setting from the disk-writable representation of the <c> <paramref name="callingModule"/>'s </c> module settings </summary>
    ''' <param name="callingModule"> The name of the module whose settings will be cleared </param>
    ''' <param name="createSection"> The subroutine that creates the module's disk-writable settings </param>
    Public Sub restoreDefaultSettings(callingModule As String, createSection As Action)
        If createSection Is Nothing Then argIsNull(NameOf(createSection)) : Return
        clearAllModuleSettings(callingModule)
        createSection()
        saveSettingsFile()
    End Sub

End Module