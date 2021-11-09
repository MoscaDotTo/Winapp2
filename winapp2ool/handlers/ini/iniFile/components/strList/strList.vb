'    Copyright (C) 2018-2021 Hazel Ward
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
''' A helpful wrapper for List(Of String)s 
''' </summary>
Public Class strList
    ''' <summary> Creates a new (empty) strList</summary>
    Public Sub New()
        Items = New List(Of String)
    End Sub

    ''' <summary>The values inside the strList</summary>
    Public Property Items As List(Of String)

    ''' <summary>Returns the number of items in the strList</summary>
    Public Function Count() As Integer
        Return If(Items Is Nothing, 0, Items.Count)
    End Function

    ''' <summary>Returns the index of the given String in the list if it exists. Else -1</summary>
    ''' <param name="item">A String to search for in the list</param>
    Public Function indexOf(item As String) As Integer
        Return Items.IndexOf(item)
    End Function

    ''' <summary>Conditionally adds an item to the list </summary>
    ''' <param name="item">A string value to add to the list</param>
    ''' <param name="cond">The optional condition under which the value should be added (default: true)</param>
    Public Sub add(item As String, Optional cond As Boolean = True)
        If cond Then Items.Add(item)
    End Sub

    ''' <summary>Conditionally adds an array of items to the strlist</summary>>
    ''' <param name="items">An array of items to be added</param>
    ''' <param name="cond">The optional condition under which the items should be added (default: true)</param>
    Public Sub add(items As String(), Optional cond As Boolean = True)
        If items Is Nothing Then argIsNull(NameOf(items)) : Return
        For Each item In items
            add(item, cond)
        Next
    End Sub

    ''' <summary>Conditionally adds the contents of another strlist to the strlist</summary>
    ''' <param name="items">A strlist of items to be added</param>
    ''' <param name="cond">The optional condition under which the items should be added (default: true)</param>
    Public Sub add(items As strList, Optional cond As Boolean = True)
        If items Is Nothing Then argIsNull(NameOf(items)) : Return
        For Each item In items.Items
            add(item, cond)
        Next
    End Sub

    '''<summary>Empties the strlist</summary>
    Public Sub clear()
        Items.Clear()
    End Sub

    ''' <summary>Returns true if the list contains a given value. Case sensitive by default</summary>
    ''' <param name="givenValue">A value to search the list for</param>
    ''' <param name="ignoreCase">The optional condition specifying whether string casing should be ignored</param>
    Public Function contains(givenValue As String, Optional ignoreCase As Boolean = False) As Boolean
        Items.Contains(givenValue, StringComparer.InvariantCultureIgnoreCase)
        Return If(ignoreCase, Items.Contains(givenValue, StringComparer.InvariantCultureIgnoreCase), Items.Contains(givenValue, StringComparer.InvariantCulture))
    End Function

    ''' <summary>
    ''' Checks whether the current value appears in the given list of strings (case insensitive). Returns true if there is a duplicate,
    ''' otherwise, adds the current value to the list and returns false.
    ''' </summary>
    ''' <param name="currentValue">The current value to be audited</param>
    Public Function chkDupes(currentValue As String) As Boolean
        If currentValue Is Nothing Then argIsNull(NameOf(currentValue)) : Return False
        If currentValue.Length = 0 Then Return True
        For Each value In Items
            If currentValue.Equals(value, StringComparison.InvariantCultureIgnoreCase) Then Return True
        Next
        Items.Add(currentValue)
        Return False
    End Function

    ''' <summary>Construct a list of neighbors for strings in a list</summary>
    Public Function getNeighborList() As List(Of KeyValuePair(Of String, String))
        Dim neighborList As New List(Of KeyValuePair(Of String, String)) From {New KeyValuePair(Of String, String)("first", Items(1))}
        For i = 1 To Items.Count - 2
            neighborList.Add(New KeyValuePair(Of String, String)(Items(i - 1), Items(i + 1)))
        Next
        neighborList.Add(New KeyValuePair(Of String, String)(Items(Items.Count - 2), "last"))
        Return neighborList
    End Function

    ''' <summary>Replaces an item in a list of strings at the index of another given item</summary>
    ''' <param name="indexOfText">The text to be replaced</param>
    ''' <param name="newText">The replacement text</param>
    Public Sub replaceStrAtIndexOf(indexOfText As String, newText As String)
        Items(Items.IndexOf(indexOfText)) = newText
    End Sub
End Class