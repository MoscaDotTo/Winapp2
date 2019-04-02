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
''' An object representing a section of a .ini file
''' </summary>
Public Class iniSection
    Public startingLineNumber As Integer
    Public endingLineNumber As Integer
    Public name As String
    Public keys As New Dictionary(Of Integer, iniKey)

    ''' <summary>
    ''' Sorts a section's keys into keylists based on their KeyType
    ''' </summary>
    ''' <param name="listOfKeyLists">The list of keyLists to be sorted into</param>
    ''' The last list in the keylist list holds the error keys
    Public Sub constKeyLists(ByRef listOfKeyLists As List(Of keyList))
        Dim keyTypeList As New List(Of String)
        listOfKeyLists.ForEach(Sub(kl) keyTypeList.Add(kl.keyType.ToLower))
        For Each key In keys.Values
            Dim type = key.KeyType.ToLower
            If keyTypeList.Contains(type) Then listOfKeyLists(keyTypeList.IndexOf(type)).add(key) Else listOfKeyLists.Last.add(key)
        Next
    End Sub

    ''' <summary>
    ''' Removes a series of keys from the section
    ''' </summary>
    ''' <param name="indicies"></param>
    Public Sub removeKeys(indicies As List(Of Integer))
        indicies.ForEach(Sub(ind) keys.Remove(ind))
    End Sub

    ''' <summary>
    ''' Returns the iniSection name as it would appear on disk.
    ''' </summary>
    ''' <returns></returns>
    Public Function getFullName() As String
        Return $"[{name}]"
    End Function

    ''' <summary>
    ''' Creates a new (empty) iniSection object.
    ''' </summary>
    Public Sub New()
        startingLineNumber = 0
        endingLineNumber = 0
        name = ""
    End Sub

    ''' <summary>
    ''' Creates a new iniSection object without tracking the line numbers
    ''' </summary>
    ''' <param name="listOfLines">The list of Strings comprising the iniSection</param>
    ''' <param name="listOfLineCounts">The list of line numbers associated with the lines</param>
    Public Sub New(ByVal listOfLines As List(Of String), Optional listOfLineCounts As List(Of Integer) = Nothing)
        name = listOfLines(0).Trim(CChar("["), CChar("]"))
        startingLineNumber = If(listOfLineCounts IsNot Nothing, listOfLineCounts(0), 1)
        endingLineNumber = startingLineNumber + listOfLines.Count
        If listOfLines.Count > 1 Then
            For i As Integer = 1 To listOfLines.Count - 1
                keys.Add(i - 1, New iniKey(listOfLines(i), If(listOfLineCounts Is Nothing, 0, listOfLineCounts(i))))
            Next
        End If
    End Sub

    ''' <summary>
    ''' Returns the keys in the iniSection as a list of Strings
    ''' </summary>
    ''' <returns></returns>
    Public Function getKeysAsList() As List(Of String)
        Dim out As New List(Of String)
        For Each key In Me.keys.Values
            out.Add(key.toString)
        Next
        Return out
    End Function

    ''' <summary>
    ''' Compares two iniSections, returns false if they are not the same.
    ''' </summary>
    ''' <param name="ss">The section to be compared against</param>
    ''' <param name="removedKeys">A return list on iniKey objects that appear in the iniFile object but not the given</param>
    ''' <param name="addedKeys">A return list of iniKey objects that appear in the given iniFile object but not this one</param>
    ''' <returns></returns>
    Public Function compareTo(ss As iniSection, ByRef removedKeys As keyList, ByRef addedKeys As keyList) As Boolean
        ' Create a copy of the section so we can modify it
        Dim secondSection As New iniSection With {.name = ss.name, .startingLineNumber = ss.startingLineNumber}
        For i As Integer = 0 To ss.keys.Count - 1
            secondSection.keys.Add(i, ss.keys.Values(i))
        Next
        Dim noMatch As Boolean
        Dim tmpList As New List(Of Integer)
        For Each key In keys.Values
            noMatch = True
            For i As Integer = 0 To secondSection.keys.Values.Count - 1
                Dim sKey = secondSection.keys.Values(i)
                Select Case True
                    Case key.compareTypes(sKey) And key.compareValues(sKey)
                        noMatch = False
                        tmpList.Add(i)
                        Exit For
                End Select
            Next
            ' If the key isn't found in the second (newer) section, consider it removed for now
            If noMatch Then removedKeys.add(key)
        Next
        ' Remove all matched keys
        tmpList.Reverse()
        secondSection.removeKeys(tmpList)
        ' Assume any remaining keys have been added
        For Each key In secondSection.keys.Values
            addedKeys.add(key)
        Next
        Return removedKeys.keyCount + addedKeys.keyCount = 0
    End Function

    ''' <summary>
    ''' Returns an iniSection as it would appear on disk as a String
    ''' </summary>
    ''' <returns></returns>
    Public Overrides Function ToString() As String
        Dim out As String = Me.getFullName
        For Each key In keys.Values
            out += Environment.NewLine & key.toString
        Next
        out += Environment.NewLine
        Return out
    End Function
End Class