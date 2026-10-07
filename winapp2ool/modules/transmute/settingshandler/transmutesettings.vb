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
''' This module contains properties, saved through the settings system, which define the current
''' state of the Transmutator and its sub modes, the file locations of the transmute files,
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
    ''' Stores the path to which the Transmuted file should be written back to disk <br />
    ''' Default: <c> winapp2-transmuted.ini </c> in the current directory, so the base file on
    ''' disk is left untouched
    ''' </summary>
    Public Property TransmuteFile3 As iniFileChooser = New iniFileChooser(Environment.CurrentDirectory, "winapp2-transmuted.ini", "winapp2.ini", mustExist:=False)

    ''' <summary>
    ''' Indicates whether the module's settings have been modified from their defaults
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
    ''' are matched case-insensitively
    ''' </description>
    ''' </item>
    '''
    ''' <item>
    ''' <c> Replace </c>
    ''' <description>
    ''' Contains two sub modes. Replaces sections or individual keys in the base file with content
    ''' from the source file. Section names are matched case-insensitively
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
    ''' Indicates whether <see cref="TransmuteFile3"/> should be sorted and saved with winapp2.ini
    ''' formatting. When <c> False </c>, we save a plain ini file with alphabetized sections. <br />
    ''' Default: <c> True </c>
    ''' </summary>
    Public Property UseWinapp2Syntax As Boolean = True

    ''' <summary>
    ''' Indicates whether source file sentinel sections are treated as global operations: a
    ''' <c> [*] </c> section applies to every section in the base file, sections whose names
    ''' begin with <c> *Map: </c> are key mapping rules during Replace ByKey operations, and
    ''' sections whose names begin with <c> *Name: </c> apply to the base sections their filters
    ''' select <br />
    ''' When <c> False </c>, these sentinel names fall through to normal section name matching,
    ''' for generic ini files where they could be real section names <br />
    ''' Default: <c> True </c>
    ''' </summary>
    Public Property RecognizeGlobalSections As Boolean = True

    ''' <summary>
    ''' Restores the default state of the module's properties and records them in the settings
    ''' file through <see cref="SaveModule"/>. Nothing is written to disk here.
    ''' <see cref="FlushIfDirty"/> does that later, if the save gate allows it
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
        RecognizeGlobalSections = True
        TransmuteModuleSettingsChanged = False
        SaveModule(NameOf(Transmute), GetType(transmuteSettings))

    End Sub

End Module
