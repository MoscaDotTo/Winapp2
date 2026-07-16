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

''' <summary>
''' Holds the settings for the Flavorizer module, which provides a user interface
''' for applying "flavors" (sets of modifications) to ini files. <br /><br />
'''
''' The flavorization process uses multiple correction files applied in this order:
'''
''' <list type="number">
''' <item> Section Removal (<c> File3 </c>) — Remove entire sections </item>
''' <item> Key Name Removal (<c> File4 </c>) — Remove keys by name matching </item>
''' <item> Key Value Removal (<c> File5 </c>) — Remove keys by value and keytype matching </item>
''' <item> Section Replacement (<c> File6 </c>) — Replace entire sections </item>
''' <item> Key Replacement (<c> File7 </c>) — Replace individual key values </item>
''' <item> Section and Key Additions (<c> File8 </c>) — Add new sections and keys </item>
''' </list>
'''
''' All correction files are optional. The flavorization process skips any that are not
''' specified or do not exist.
''' </summary>
Public Module FlavorizerSettings

    ''' <summary>
    ''' The base <c> iniFileChooser </c> to which the flavor will be applied
    ''' </summary>
    Public Property FlavorizerFile1 As iniFileChooser = New iniFileChooser(Environment.CurrentDirectory, "winapp2.ini", "winapp2.ini")

    ''' <summary>
    ''' The location to which the Flavorizer will save its output
    ''' </summary>
    Public Property FlavorizerFile2 As iniFileChooser = New iniFileChooser(Environment.CurrentDirectory, "winapp2-flavorized.ini", "winapp2-flavorized.ini", mustExist:=False)

    ''' <summary>
    ''' Section removal file - contains sections to be removed entirely from the base file. <br />
    ''' Sections will be removed regardless of their content. <br />
    ''' Applied in the first stage of flavorization.
    ''' </summary>
    Public Property FlavorizerFile3 As iniFileChooser = New iniFileChooser(Environment.CurrentDirectory, "", "", mustExist:=False)

    ''' <summary>
    ''' Key name removal file - contains keys to be removed by name matching. <br />
    ''' Section and key name matching is case-insensitive. <br />
    ''' The values in this file are ignored - only key names matter. <br />
    ''' Applied in the second stage of flavorization.
    ''' </summary>
    Public Property FlavorizerFile4 As iniFileChooser = New iniFileChooser(Environment.CurrentDirectory, "", "", mustExist:=False)

    ''' <summary>
    ''' Key value removal file - contains keys to be removed by keytype and value matching. <br />
    ''' Numbers in key names are ignored for matching purposes. <br />
    ''' Both the keytype (name without numbers) and value must match. <br />
    ''' Applied in the third stage of flavorization.
    ''' </summary>
    Public Property FlavorizerFile5 As iniFileChooser = New iniFileChooser(Environment.CurrentDirectory, "", "", mustExist:=False)

    ''' <summary>
    ''' Section replacement file - contains complete sections that will replace <br />
    ''' sections of the same name in the base file. <br />
    ''' Section name matching is case-insensitive. <br />
    ''' This completely replaces the section content. <br />
    ''' Applied in the fourth stage of flavorization.
    ''' </summary>
    Public Property FlavorizerFile6 As iniFileChooser = New iniFileChooser(Environment.CurrentDirectory, "", "", mustExist:=False)

    ''' <summary>
    ''' Key replacement file - contains individual keys that will replace <br />
    ''' keys of the same name within matching sections in the base file. <br />
    ''' Section and key name matching is case-insensitive. <br />
    ''' Applied in the fifth stage of flavorization.
    ''' </summary>
    Public Property FlavorizerFile7 As iniFileChooser = New iniFileChooser(Environment.CurrentDirectory, "", "", mustExist:=False)

    ''' <summary>
    ''' Additions file - contains sections and keys to be added to the base file. <br />
    ''' New sections will be added as-is. <br />
    ''' Keys within existing sections will be added to those sections. <br />
    ''' Applied in the sixth and final stage of flavorization.
    ''' </summary>
    Public Property FlavorizerFile8 As iniFileChooser = New iniFileChooser(Environment.CurrentDirectory, "", "", mustExist:=False)

    ''' <summary>
    ''' Holds the "Target Directory" for the Flavorizer module which is used to automatically
    ''' detect the set of Flavor files. <br />
    ''' Never has a file name and is never saved to disk.
    ''' </summary>
    Public Property FlavorizerFile9 As iniFileChooser = New iniFileChooser(Environment.CurrentDirectory, "", "", mustExist:=False)

    ''' <summary>
    ''' Indicates whether the output should be formatted as a winapp2.ini file <br />
    ''' Default: <c> True </c>
    ''' </summary>
    Public Property FlavorizeAsWinapp As Boolean = True

    ''' <summary>
    ''' Indicates that the module settings have been modified from their defaults <br />
    ''' Default: <c> False </c>
    ''' </summary>
    Public Property FlavorizerModuleSettingsChanged As Boolean = False

    ''' <summary>
    ''' Restores the default state of the Flavorizer module's properties and persists them via <c> SaveModule2 </c>
    ''' </summary>
    Public Sub initDefaultFlavorizerSettings()

        FlavorizerFile1 = New iniFileChooser(Environment.CurrentDirectory, "winapp2.ini", "winapp2.ini")
        FlavorizerFile2 = New iniFileChooser(Environment.CurrentDirectory, "winapp2-flavorized.ini", "winapp2-flavorized.ini", mustExist:=False)
        FlavorizerFile3 = New iniFileChooser(Environment.CurrentDirectory, "", "", mustExist:=False)
        FlavorizerFile4 = New iniFileChooser(Environment.CurrentDirectory, "", "", mustExist:=False)
        FlavorizerFile5 = New iniFileChooser(Environment.CurrentDirectory, "", "", mustExist:=False)
        FlavorizerFile6 = New iniFileChooser(Environment.CurrentDirectory, "", "", mustExist:=False)
        FlavorizerFile7 = New iniFileChooser(Environment.CurrentDirectory, "", "", mustExist:=False)
        FlavorizerFile8 = New iniFileChooser(Environment.CurrentDirectory, "", "", mustExist:=False)
        FlavorizerFile9 = New iniFileChooser(Environment.CurrentDirectory, "", "", mustExist:=False)
        FlavorizeAsWinapp = True
        FlavorizerModuleSettingsChanged = False
        SaveModule2(NameOf(Flavorizer), GetType(FlavorizerSettings))

    End Sub

End Module
