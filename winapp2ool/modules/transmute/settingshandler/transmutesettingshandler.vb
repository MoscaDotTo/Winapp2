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

''' <summary> 
''' Manages the settings of the Transmute module for the purpose of syncing to disk 
''' </summary>
''' 
''' Docs last updated: 2025-07-15 | Code last updated: 2025-07-15
Public Module transmuteSettingsHandler

    ''' <summary> 
    ''' Restores the default state of the module's properties 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-07-15 | Code last updated: 2025-07-15
    Public Sub initDefaultTransmuteSettings()

        TransmuteFile1.resetParams()
        TransmuteFile2.resetParams()
        TransmuteFile3.resetParams()
        Transmutator = TransmuteMode.Add
        TransmuteReplaceMode = ReplaceMode.ByKey
        TransmuteRemoveMode = RemoveMode.ByKey
        TransmuteRemoveKeyMode = RemoveKeyMode.ByName
        TransmuteModuleSettingsChanged = False
        restoreDefaultSettings(NameOf(Transmute), AddressOf createTransmuteSettingsSection)

    End Sub

    ''' <summary> 
    ''' Loads values from disk into memory for the Transmute module settings 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-07-15 | Code last updated: 2025-07-15
    Public Sub getSerializedTransmuteSettings()

        LoadModuleSettingsFromDict(NameOf(Transmute), GetType(transmuteSettings))

        Transmutator = If(TransmuteModeIsAdd, TransmuteMode.Add,
                            If(TransmuteModeIsRemove, TransmuteMode.Remove, TransmuteMode.Replace))

        TransmuteReplaceMode = If(ReplaceModeIsBySection, ReplaceMode.BySection, ReplaceMode.ByKey)

        TransmuteRemoveMode = If(RemoveModeIsBySection, RemoveMode.BySection, RemoveMode.ByKey)

        TransmuteRemoveKeyMode = If(RemoveKeyModeIsByName, RemoveKeyMode.ByName, RemoveKeyMode.ByValue)

    End Sub

    '''<summary>
    ''' Adds the current (typically default) state of the module's 
    ''' settings into the disk-writable settings representation 
    '''</summary>
    ''' 
    ''' Docs last updated: 2025-07-15 | Code last updated: 2025-07-15
    Public Sub createTransmuteSettingsSection()

        updateTransmuteEnumFlags()

        Dim settingsModule = GetType(transmuteSettings)
        Dim moduleName = NameOf(Transmute)

        createModuleSettingsSection(moduleName, settingsModule)

    End Sub

    ''' <summary>
    ''' Syncs the current state of the Transmute module's enums with their boolean representations
    ''' ensuring that they accurately reflect the current state of the module
    ''' 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-07-15 | Code last updated: 2025-07-15
    Public Sub updateTransmuteEnumFlags()

        TransmuteModeIsAdd = Transmutator = TransmuteMode.Add
        TransmuteModeIsRemove = Transmutator = TransmuteMode.Remove
        TransmuteModeIsReplace = Transmutator = TransmuteMode.Replace

        ReplaceModeIsBySection = TransmuteReplaceMode = ReplaceMode.BySection
        ReplaceModeIsByKey = TransmuteReplaceMode = ReplaceMode.ByKey

        RemoveModeIsBySection = TransmuteRemoveMode = RemoveMode.BySection
        RemoveModeIsByKey = TransmuteRemoveMode = RemoveMode.ByKey

        RemoveKeyModeIsByName = TransmuteRemoveKeyMode = RemoveKeyMode.ByName
        RemoveKeyModeIsByValue = TransmuteRemoveKeyMode = RemoveKeyMode.ByValue

    End Sub

    ''' <summary>
    ''' Updates the disk-writable representation of each of Transmute's enums 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-07-15 | Code last updated: 2025-07-15
    Public Sub updateAllTransmuteEnumSettings()

        updateTransmuteEnumFlags()

        Dim names As New List(Of String)
        Dim vals As New List(Of String)

        For Each prop In GetType(transmuteSettings).GetProperties()

            ' This will actually update all Booleans but that's probably fine
            If prop.PropertyType IsNot GetType(Boolean) Then Continue For

            Dim val = DirectCast(prop.GetValue(Nothing), Boolean)
            updateSettings(NameOf(Transmute), prop.Name, val.ToString())

        Next

    End Sub

End Module
