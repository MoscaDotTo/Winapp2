'    Copyright (C) 2018-2019 Robbie Ward
' 
'    This file is a part of Winapp2ool
' 
'    Winapp2ool is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    Winap2ool is distributed in the hope that it will be useful,
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
    Public name As String
    Public fullName As String
    Public detectOS As New keyList("DetectOS")
    Public langSecRef As New keyList("LangSecRef")
    Public sectionKey As New keyList("Section")
    Public specialDetect As New keyList("SpecialDetect")
    Public detects As New keyList("Detect")
    Public detectFiles As New keyList("DetectFile")
    Public defaultKey As New keyList("Default")
    Public warningKey As New keyList("Warning")
    Public fileKeys As New keyList("FileKey")
    Public regKeys As New keyList("RegKey")
    Public excludeKeys As New keyList("ExcludeKey")
    Public errorKeys As New keyList("Error")
    Public keyListList As New List(Of keyList) From {detectOS, langSecRef, sectionKey, specialDetect, detects, detectFiles,
                                                    defaultKey, warningKey, fileKeys, regKeys, excludeKeys, errorKeys}
    Public lineNum As New Integer

    ''' <summary>
    ''' Construct a new winapp2entry object from an iniSection
    ''' </summary>
    ''' <param name="section">A winapp2.ini format iniSection object</param>
    Public Sub New(ByVal section As iniSection)
        fullName = section.getFullName
        name = section.name
        updKeyListList()
        lineNum = section.startingLineNumber
        section.constKeyLists(keyListList)
    End Sub

    ''' <summary>
    ''' Clears and updates the keyListList with the current state of the keys
    ''' </summary>
    Private Sub updKeyListList()
        keyListList = New List(Of keyList) From {detectOS, langSecRef, sectionKey, specialDetect, detects, detectFiles,
                                                 defaultKey, warningKey, fileKeys, regKeys, excludeKeys, errorKeys}
    End Sub

    ''' <summary>
    ''' Returns the keys in each keyList back as a list of Strings in winapp2.ini (style) order
    ''' </summary>
    ''' <returns></returns>
    Public Function dumpToListOfStrings() As List(Of String)
        Dim outList As New List(Of String) From {fullName}
        updKeyListList()
        keyListList.ForEach(Sub(lst) lst.keys.ForEach(Sub(key) outList.Add(key.toString)))
        Return outList
    End Function
End Class