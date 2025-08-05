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
''' Holds the settings for the Flavorizer module, which provides a user interface
''' for applying "flavors" (sets of modifications) to ini files.
''' 
''' The flavorization process uses multiple correction files applied in this order:
''' 1. Section Removal (File3) -> Remove entire sections
''' 2. Key Name Removal (File4) -> Remove keys by name matching
''' 3. Key Value Removal (File5) -> Remove keys by value and keytype matching  
''' 4. Section Replacement (File6) -> Replace entire sections
''' 5. Key Replacement (File7) -> Replace individual key values
''' 6. Section and Key Additions (File8) -> Add new sections and keys
''' 
''' All correction files are optional - the flavorization will skip any that
''' are not specified or do not exist.
''' </summary>
''' 
''' Docs last updated: 2025-08-01 | Code last updated: 2025-08-01
Public Module FlavorizerSettings

    ''' <summary>
    ''' The base <c> iniFile </c> to which the flavor will be applied
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-08-01 | Code last updated: 2025-08-01
    Public Property FlavorizerFile1 As New iniFile(Environment.CurrentDirectory, "winapp2.ini")

    ''' <summary>
    ''' The location to which the Flavorizer will save its output 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-08-01 | Code last updated: 2025-08-01
    Public Property FlavorizerFile2 As New iniFile(Environment.CurrentDirectory, "winapp2-flavorized.ini")

    ''' <summary>
    ''' Section removal file - contains sections to be removed entirely from the base file. <br />
    ''' Sections will be removed regardless of their content. <br />
    ''' Applied in the first stage of flavorization.
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-08-01 | Code last updated: 2025-08-01
    Public Property FlavorizerFile3 As New iniFile(Environment.CurrentDirectory, "")

    ''' <summary>
    ''' Key name removal file - contains keys to be removed by name matching. <br />
    ''' Keys must be within sections that match exactly (case sensitive). <br/>
    ''' The values in this file are ignored - only key names matter. <br />
    ''' Applied in the second stage of flavorization.
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-08-01 | Code last updated: 2025-08-01
    Public Property FlavorizerFile4 As New iniFile(Environment.CurrentDirectory, "")

    ''' <summary>
    ''' Key value removal file - contains keys to be removed by keytype and value matching. <br />
    ''' Numbers in key names are ignored for matching purposes. <br />
    ''' Both the keytype (name without numbers) and value must match. <br />
    ''' Applied in the third stage of flavorization. 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-08-01 | Code last updated: 2025-08-01
    Public Property FlavorizerFile5 As New iniFile(Environment.CurrentDirectory, "")

    ''' <summary>
    ''' Section replacement file - contains complete sections that will replace <br />
    ''' sections of the same name in the base file. <br />
    ''' Section names must match exactly (case sensitive). <br />
    ''' This completely replaces the section content. <br/>
    ''' Applied in the fourth stage of flavorization. <br />
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-08-01 | Code last updated: 2025-08-01
    Public Property FlavorizerFile6 As New iniFile(Environment.CurrentDirectory, "")

    ''' <summary>
    ''' Key replacement file - contains individual keys that will replace <br />
    ''' keys of the same name within matching sections in the base file. <br />
    ''' Both section names and key names must match exactly (case sensitive). <br />
    ''' Applied in the fifth stage of flavorization. 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-08-01 | Code last updated: 2025-08-01
    Public Property FlavorizerFile7 As New iniFile(Environment.CurrentDirectory, "")

    ''' <summary>
    ''' Additions file - contains sections and keys to be added to the base file. <br />
    ''' New sections will be added as-is. <br />
    ''' Keys within existing sections will be added to those sections. <br />
    ''' Applied in the sixth and final stage of flavorization. <br />
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-08-01 | Code last updated: 2025-08-01
    Public Property FlavorizerFile8 As New iniFile(Environment.CurrentDirectory, "")

    ''' <summary>
    ''' Holds the "Target Directory" for the Flavorizer module which is used to automatically
    ''' detect the set of Flavor files. <br /> 
    ''' Never has a file name and is never saved to disk
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-08-05 | Code last updated: 2025-08-05
    Public Property FlavorizerFile9 As New iniFile(Environment.CurrentDirectory, "")

    ''' <summary>
    ''' Indicates whether the output should be formatted as a winapp2.ini file <br/>
    ''' Default: <c> True </c>
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-08-01 | Code last updated: 2025-08-01
    Public Property FlavorizeAsWinapp As Boolean = True

    ''' <summary>
    ''' Indicates that the module settings have been modified from their defaults <br />
    ''' Default: <c> False </c>
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-08-01 | Code last updated: 2025-08-01
    Public Property FlavorizerModuleSettingsChanged As Boolean = False

End Module