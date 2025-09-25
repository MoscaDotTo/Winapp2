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
''' Holds the settings for the Transmute module, which makes changes to a 
''' base iniFile based on the content of a source inifile 
''' This module contains properties that will be synced to disk which define the current state of
''' the Transmutator and its sub modes, the file locations of the transmute files, 
''' and whether the settings have been changed from their defaults.
''' </summary>
Public Module transmuteSettings

    ''' <summary> 
    ''' The 'base' file for Transmute, the one whose content will be modified by the transmute process
    ''' based on the contents of the 'source' file 
    ''' </summary>
    Public Property TransmuteFile1 As New iniFile(Environment.CurrentDirectory, "winapp2.ini")

    ''' <summary> 
    ''' The 'source' file for Transmute, the one whose content will be used by the transmute process
    ''' to make changes to the 'base' file
    ''' </summary>
    Public Property TransmuteFile2 As New iniFile(Environment.CurrentDirectory, "")

    ''' <summary> 
    ''' Stores the path to which the Transmuted file should be written back to disk 
    ''' (overwrites <c> TransmuteFile1 </c> by default) 
    ''' </summary>
    Public Property TransmuteFile3 As New iniFile(Environment.CurrentDirectory, "winapp2.ini", "winapp2-transmuted.ini")

    ''' <summary> 
    ''' Indicates that module's settings have been modified from their defaults 
    ''' </summary>
    Public Property TransmuteModuleSettingsChanged As Boolean = False

    ''' <summary>
    ''' The primary transmutation mode for the Transmute module <br />
    ''' Default: <c> <b> Add </b> </c>
    ''' 
    ''' <list>
    ''' 
    ''' <item>
    ''' <c> <b> Add </b> </c>
    ''' <description> 
    ''' Adds sections from the source file to the base file. If a section exists already in the base
    ''' file, the keys from the source section will be added to the base section. Section names 
    ''' must match exactly (case sensitive)
    ''' </description>
    ''' </item>
    ''' 
    ''' <item>
    ''' <c> <b> Replace </b></c>
    ''' <description>
    ''' Contains two sub modes. Replaces sections or individual keys in the base file with content 
    ''' from the source file. Section names must match exactly (case sensitive)
    ''' </description>
    ''' </item>
    ''' 
    ''' <item>
    ''' <c> <b> Remove </b> </c>
    ''' <description>
    ''' Contains two sub modes one of which itself contains two sub modes. Removes sections or 
    ''' individual keys from the base file. Key removal can be performed by value 
    ''' (ignores key numbering but requires a key type) or by key name (works best for unnumbered
    ''' keys) 
    ''' </description>
    ''' </item>
    ''' </list>
    '''
    ''' </summary>
    Public Property Transmutator As TransmuteMode = TransmuteMode.Add

    ''' <summary>
    ''' The granularity level for the <c> Replace Transmutator </c>, has two sub modes <br />
    ''' Default: ByKey 
    ''' <list> 
    ''' 
    ''' <item>
    ''' <c> <b> BySection </b></c>
    ''' <description>
    ''' Replaces entire sections when collisions occur
    ''' </description>
    ''' </item>
    ''' 
    ''' <item>
    ''' <c> <b> ByKey </b> </c>
    ''' <description>
    ''' Replaces individual keys when collisions occur 
    ''' </description>
    ''' </item>
    ''' </list>
    ''' 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-07-30 | Code last updated: 2025-07-30
    Public Property TransmuteReplaceMode As ReplaceMode = ReplaceMode.ByKey

    ''' <summary>
    ''' The granularity level for the <c> Remove Transmutator </c>, has two sub modes <br />
    ''' Default: ByKey 
    ''' <list> 
    ''' 
    ''' <item>
    ''' <c> <b> BySection </b></c>
    ''' <description>
    ''' Removes entire sections when collisions occur
    ''' </description>
    ''' </item>
    ''' 
    ''' <item>
    ''' <c> <b> ByKey </b> </c>
    ''' <description>
    ''' Removes individual keys when collisions occur 
    ''' </description>
    ''' </item>
    ''' </list>
    ''' 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-07-30 | Code last updated: 2025-07-30
    Public Property TransmuteRemoveMode As RemoveMode = RemoveMode.ByKey

    ''' <summary>
    ''' The granularity level for the <c> Remove by Key Transmutator </c>, has two sub modes <br />
    ''' Default: ByName
    ''' <list> 
    ''' 
    ''' <item>
    ''' <c> <b> ByName </b></c>
    ''' <description>
    ''' Removes keys from the base section if they have the same name as a key in the source section <br />
    ''' Ignores provided values
    ''' </description>
    ''' </item>
    ''' 
    ''' <item>
    ''' <c> <b> ByValue </b> </c>
    ''' <description>
    ''' Removes keys from the base section if they have the same KeyType and Value <br />
    ''' Ignores numbers in the Name of the <c> iniKey </c>
    ''' </description>
    ''' </item>
    ''' </list>
    ''' 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-07-30 | Code last updated: 2025-07-30
    Public Property TransmuteRemoveKeyMode As RemoveKeyMode = RemoveKeyMode.ByName

    ''' <summary>
    ''' Indicates that <c> <cref> TransmuteFile3 </cref> </c> should be saved with winapp2.ini formatting <br />
    ''' Default: <c> True </c>
    ''' </summary>
    Public Property UseWinapp2Syntax As Boolean = True

End Module