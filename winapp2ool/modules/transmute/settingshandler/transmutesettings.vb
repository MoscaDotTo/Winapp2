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
''' Holds the settings for the Transmute module, which makes changes to a
''' base ini file based on the content of a source ini file.
''' This module contains properties that will be synced to disk which define the current state of
''' the Transmutator and its sub modes, the file locations of the transmute files,
''' and whether the settings have been changed from their defaults.
''' </summary>
Public Module transmuteSettings

    ''' <summary>
    ''' The 'base' file for Transmute, the one whose content will be modified by the transmute process
    ''' based on the contents of the 'source' file
    ''' </summary>
    Public Property TransmuteFile1 As iniFileChooser = New iniFileChooser(Environment.CurrentDirectory, "winapp2.ini", "winapp2.ini")

    ''' <summary>
    ''' The 'source' file for Transmute, the one whose content will be used by the transmute process
    ''' to make changes to the 'base' file
    ''' </summary>
    Public Property TransmuteFile2 As iniFileChooser = New iniFileChooser(Environment.CurrentDirectory, "", "", mustExist:=False)

    ''' <summary>
    ''' Stores the path to which the Transmuted file should be written back to disk
    ''' (overwrites <c> TransmuteFile1 </c> by default)
    ''' </summary>
    Public Property TransmuteFile3 As iniFileChooser = New iniFileChooser(Environment.CurrentDirectory, "winapp2-transmuted.ini", "winapp2.ini", mustExist:=False)

    ''' <summary>
    ''' Indicates that module's settings have been modified from their defaults
    ''' </summary>
    Public Property TransmuteModuleSettingsChanged As Boolean = False

    ''' <summary>
    ''' The primary transmutation mode for the Transmute module <br />
    ''' Default: <c> Add </c>
    '''
    ''' <list>
    '''
    ''' <item>
    ''' <c> Add </c>
    ''' <description>
    ''' Adds sections from the source file to the base file. If a section exists already in the base
    ''' file, the keys from the source section will be added to the base section. Section names
    ''' must match exactly (case sensitive)
    ''' </description>
    ''' </item>
    '''
    ''' <item>
    ''' <c> Replace </c>
    ''' <description>
    ''' Contains two sub modes. Replaces sections or individual keys in the base file with content
    ''' from the source file. Section names must match exactly (case sensitive)
    ''' </description>
    ''' </item>
    '''
    ''' <item>
    ''' <c> Remove </c>
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
    ''' Default: <c> ByKey </c>
    ''' <list>
    '''
    ''' <item>
    ''' <c> BySection </c>
    ''' <description>
    ''' Replaces entire sections when collisions occur
    ''' </description>
    ''' </item>
    '''
    ''' <item>
    ''' <c> ByKey </c>
    ''' <description>
    ''' Replaces individual keys when collisions occur
    ''' </description>
    ''' </item>
    ''' </list>
    '''
    ''' </summary>
    Public Property TransmuteReplaceMode As ReplaceMode = ReplaceMode.ByKey

    ''' <summary>
    ''' The granularity level for the <c> Remove Transmutator </c>, has two sub modes <br />
    ''' Default: <c> ByKey </c>
    ''' <list>
    '''
    ''' <item>
    ''' <c> BySection </c>
    ''' <description>
    ''' Removes entire sections when collisions occur
    ''' </description>
    ''' </item>
    '''
    ''' <item>
    ''' <c> ByKey </c>
    ''' <description>
    ''' Removes individual keys when collisions occur
    ''' </description>
    ''' </item>
    ''' </list>
    '''
    ''' </summary>
    Public Property TransmuteRemoveMode As RemoveMode = RemoveMode.ByKey

    ''' <summary>
    ''' The granularity level for the <c> Remove by Key Transmutator </c>, has two sub modes <br />
    ''' Default: <c> ByName </c>
    ''' <list>
    '''
    ''' <item>
    ''' <c> ByName </c>
    ''' <description>
    ''' Removes keys from the base section if they have the same name as a key in the source section <br />
    ''' Ignores provided values
    ''' </description>
    ''' </item>
    '''
    ''' <item>
    ''' <c> ByValue </c>
    ''' <description>
    ''' Removes keys from the base section if they have the same KeyType and Value <br />
    ''' Ignores numbers in the Name of the <c> iniKey </c>
    ''' </description>
    ''' </item>
    ''' </list>
    '''
    ''' </summary>
    Public Property TransmuteRemoveKeyMode As RemoveKeyMode = RemoveKeyMode.ByName

    ''' <summary>
    ''' Indicates that <c> TransmuteFile3 </c> should be saved with winapp2.ini formatting <br />
    ''' Default: <c> True </c>
    ''' </summary>
    Public Property UseWinapp2Syntax As Boolean = True

    ''' <summary>
    ''' Restores the default state of the module's properties and persists them via <c> SaveModule2 </c>
    ''' </summary>
    Public Sub initDefaultTransmuteSettings()

        TransmuteFile1.ResetParams()
        TransmuteFile2.ResetParams()
        TransmuteFile3.ResetParams()
        Transmutator = TransmuteMode.Add
        TransmuteReplaceMode = ReplaceMode.ByKey
        TransmuteRemoveMode = RemoveMode.ByKey
        TransmuteRemoveKeyMode = RemoveKeyMode.ByName
        UseWinapp2Syntax = True
        TransmuteModuleSettingsChanged = False
        SaveModule2(NameOf(Transmute), GetType(transmuteSettings))

    End Sub

End Module
