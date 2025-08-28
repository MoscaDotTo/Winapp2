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
''' Handles the loading, saving, and management of Flavorizer module settings.
''' 
''' This module manages the configuration for the Flavorizer, which provides
''' a user interface for applying "flavors" (sets of modifications) to ini files
''' using the Transmute.Flavorize() function.
''' </summary>
''' 
''' Docs last updated: 2025-08-01 | Code last updated: 2025-08-01
Public Module FlavorizerSettingsHandler

    ''' <summary>
    ''' Restores the default state of the Flavorizer module's properties
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-08-01 | Code last updated: 2025-08-01
    Public Sub initDefaultFlavorizerSettings()

        FlavorizerFile1.resetParams()
        FlavorizerFile2.resetParams()
        FlavorizerFile3.resetParams()
        FlavorizerFile4.resetParams()
        FlavorizerFile5.resetParams()
        FlavorizerFile6.resetParams()
        FlavorizerFile7.resetParams()
        FlavorizerFile8.resetParams()
        FlavorizeAsWinapp = True
        FlavorizerModuleSettingsChanged = False
        restoreDefaultSettings(NameOf(Flavorizer), AddressOf createFlavorizerSettingsSection)

    End Sub

    ''' <summary>
    ''' Assigns the module settings to Flavorizer based on the current disk-writable settings representation
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-08-01 | Code last updated: 2025-08-01
    Public Sub getSerializedFlavorizerSettings()

        LoadModuleSettingsFromDict(NameOf(Flavorizer), GetType(FlavorizerSettings))

    End Sub

    ''' <summary>
    ''' Adds the current state of the module's settings into the disk-writable settings representation 
    ''' </summary>
    ''' 
    ''' <remarks>
    ''' Most often, this is the default state of these settings 
    ''' </remarks>
    ''' 
    ''' Docs last updated: 2025-08-01 | Code last updated: 2025-08-01
    Public Sub createFlavorizerSettingsSection()

        Dim settingsModule = GetType(FlavorizerSettings)
        Dim moduleName = NameOf(Flavorizer)

        createModuleSettingsSection(moduleName, settingsModule)

    End Sub

End Module