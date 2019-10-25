'    Copyright (C) 2018-2019 Robbie Ward
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
''' Represents a winapp2.ini format iniKey and provides direct access to the keys by their type
''' </summary>
Public Class winapp2entry
    '''<summary> The name of the entry (from the iniSection)</summary>
    Public Property Name As String
    '''<summary> The full name with [Brackets] of the entry</summary>
    Public Property FullName As String
    '''<summary> The starting line number of the iniSection used to create the winapp2entry</summary>
    Public Property LineNum As Integer
    ''' <summary>The keys in the entry with KeyType "DetectOS". Valid winapp2.ini Syntax dictates there be only one key
    ''' in this list, however this property will hold an arbitrary number</summary>
    Public Property DetectOS As keyList
    ''' <summary>The keys in the entry with KeyType "LangSecRef". Valid winapp2.ini Syntax dictates there be only one key
    ''' in this list, however this property will hold an arbitrary number</summary>
    Public Property LangSecRef As keyList
    ''' <summary>The keys in the entry with KeyType "Section". Valid winapp2.ini Syntax dictates there be only one key
    ''' in this list, however this property will hold an arbitrary number</summary>
    Public Property SectionKey As keyList
    ''' <summary>The keys in the entry with KeyType "Detect"</summary>
    Public Property Detects As keyList
    ''' <summary>The keys in the entry with KeyType "DetectFile"</summary>
    Public Property DetectFiles As keyList
    ''' <summary>The keys in the entry with the KeyType "SpecialDetect"</summary>
    Public Property SpecialDetect As keyList
    '''<summary>The keys in the entry with the KeyType "Default". Valid winapp2.ini Syntax dictates there be only one key
    ''' in this list, however this property will hold an arbitrary number</summary>
    Public Property DefaultKey As keyList
    ''' <summary>The keys in the entry with KeyType "FileKey"</summary>
    Public Property FileKeys As keyList
    '''<summary>The keys in the entry with KeyType "RegKey"</summary>
    Public Property RegKeys As keyList
    ''' <summary>The keys in the entry with KeyType "Warning". Valid winapp2.ini Syntax dictates there be only one key
    ''' in this list, however this property will hold an arbitrary number</summary>
    Public Property WarningKey As keyList
    ''' <summary>The keys in the entry with KeyType "ExcludeKey"</summary>
    Public Property ExcludeKeys As keyList
    ''' <summary> Keys in the entry with KeyTypes that aren't valid winapp2.ini commands </summary>
    Public Property ErrorKeys As keyList
    ''' <summary>A list containing all the KeyLists in winapp2.ini order</summary>
    Public Property KeyListList As List(Of keyList)

    ''' <summary>Construct a new winapp2entry object from an iniSection</summary>
    ''' <param name="section">A winapp2.ini format iniSection object</param>
    Public Sub New(ByVal section As iniSection)
        FullName = section.getFullName
        Name = section.Name
        updKeyListList()
        LineNum = section.StartingLineNumber
        DetectOS = New keyList("DetectOS")
        LangSecRef = New keyList("LangSecRef")
        SectionKey = New keyList("Section")
        SpecialDetect = New keyList("SpecialDetect")
        Detects = New keyList("Detect")
        DetectFiles = New keyList("DetectFile")
        DefaultKey = New keyList("Default")
        WarningKey = New keyList("Warning")
        FileKeys = New keyList("FileKey")
        RegKeys = New keyList("RegKey")
        ExcludeKeys = New keyList("ExcludeKey")
        ErrorKeys = New keyList("Error")
        updKeyListList()
        section.constKeyLists(KeyListList)
    End Sub

    ''' <summary>Clears and updates the keyListList with the current state of the keys</summary>
    Private Sub updKeyListList()
        KeyListList = New List(Of keyList) From {DetectOS, LangSecRef, SectionKey, SpecialDetect, Detects, DetectFiles,
                                                 DefaultKey, WarningKey, FileKeys, RegKeys, ExcludeKeys, ErrorKeys}
    End Sub

    ''' <summary>Returns the keys in each keyList back as a list of Strings in winapp2.ini (style) order</summary>
    Public Function dumpToListOfStrings() As List(Of String)
        Dim outList As New List(Of String) From {FullName}
        updKeyListList()
        KeyListList.ForEach(Sub(lst) lst.Keys.ForEach(Sub(key) outList.Add(key.toString)))
        Return outList
    End Function
End Class