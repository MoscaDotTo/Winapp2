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
''' 
''' </summary>
Public Module browserbuildersettingshandler

    ''' <summary>
    ''' Restores the default state of the BrowserBuilder module's properties
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-07-02 | Code last updated: 2025-07-02
    Public Sub initDefaultBrowserBuilderSettings()

        BuilderFile1.resetParams()
        builderfile2.resetParams()
        builderfile3.resetParams()
        BrowserBuilderModuleSettingsChanged = False
        restoreDefaultSettings(NameOf(BrowserBuilder), AddressOf createBrowserBuilderSettingsSection)

    End Sub

    ''' <summary>
    ''' Assigns the module settings to BrowserBuilder based on the current disk-writable settings representation
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-07-02 | Code last updated: 2025-07-02
    Public Sub getSerializedBrowserBuilderSettings()

        LoadModuleSettingsFromDict(NameOf(BrowserBuilder), GetType(BrowserBuilder))

    End Sub

    ''' <summary>
    ''' Adds the current state of the module's settings into the disk-writable settings representation 
    ''' </summary>
    ''' 
    ''' <remarks>
    ''' Most often, this is the default state of these settings 
    ''' </remarks>
    ''' 
    ''' Docs last updated: 2025-07-02 | Code last updated: 2025-08-05
    Public Sub createBrowserBuilderSettingsSection()

        Dim settingsModule = GetType(browserbuildersettings)
        Dim moduleName = NameOf(BrowserBuilder)

        createModuleSettingsSection(moduleName, settingsModule)

    End Sub

End Module
