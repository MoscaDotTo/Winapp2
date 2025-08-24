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
''' Provides methods for managing the Combine module settings, including support methods for syncing to disk
''' Also provides the function which restores the default state of the Combine module's properties
''' </summary>
''' 
''' Docs last updated: 2025-08-20 | Code last updated: 2025-08-20
Public Module combinesettingshandler

    ''' <summary>
    ''' Restores the default state of the Combine module's properties
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-08-20 | Code last updated: 2025-08-20
    Public Sub initDefaultCombineSettings()

        CombineFile1.resetParams()
        CombineFile1.Dir = Environment.CurrentDirectory
        CombineFile3.resetParams()
        CombineFile3.Name = "combined.ini"
        CombineModuleSettingsChanged = False
        restoreDefaultSettings(NameOf(Combine), AddressOf createCombineSettingsSection)

    End Sub

    ''' <summary>
    ''' Assigns the module settings to Combine based on the current disk-writable settings representation
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-08-20 | Code last updated: 2025-08-20
    Public Sub getSerializedCombineSettings()

        LoadModuleSettingsFromDict(NameOf(Combine), GetType(combinesettings))

    End Sub

    ''' <summary>
    ''' Adds the current state of the module's settings into the disk-writable settings representation
    ''' </summary>
    ''' 
    ''' <remarks>
    ''' Most often, this is the default state of these settings
    ''' </remarks>
    ''' 
    ''' Docs last updated: 2025-08-20 | Code last updated: 2025-08-20
    Public Sub createCombineSettingsSection()

        Dim settingsModule = GetType(combinesettings)
        Dim moduleName = NameOf(Combine)

        createModuleSettingsSection(moduleName, settingsModule)

    End Sub

End Module
