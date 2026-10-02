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

    ''' <summary>Returns the index of <paramref name="item"/> in the list, or -1 if it isn't there</summary>
    ''' <param name="item">A String to search for in the list</param>
    Public Function indexOf(item As String) As Integer
        Return Items.IndexOf(item)
    End Function

    ''' <summary>Conditionally adds an item to the list </summary>
    ''' <param name="item">A string value to add to the list</param>
    ''' <param name="cond">
    ''' Indicates whether to add the item <br /><br />
    ''' Optional, Default: <c> True </c>
    ''' </param>
    Public Sub add(item As String, Optional cond As Boolean = True)
        If cond Then Items.Add(item)
    End Sub

    ''' <summary>Conditionally adds an array of items to the strlist</summary>
    ''' <param name="items">An array of items to be added</param>
    ''' <param name="cond">
    ''' Indicates whether to add the items <br /><br />
    ''' Optional, Default: <c> True </c>
    ''' </param>
    Public Sub add(items As String(), Optional cond As Boolean = True)
        If items Is Nothing Then argIsNull(NameOf(items)) : Return
        For Each item In items
            add(item, cond)
        Next
    End Sub

    ''' <summary>Conditionally adds the contents of another strlist to the strlist</summary>
    ''' <param name="items">A strlist of items to be added</param>
    ''' <param name="cond">
    ''' Indicates whether to add the items <br /><br />
    ''' Optional, Default: <c> True </c>
    ''' </param>
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

    ''' <summary>Returns whether the list contains <paramref name="givenValue"/></summary>
    ''' <param name="givenValue">A value to search the list for</param>
    '''
    ''' <param name="ignoreCase">
    ''' Indicates whether to ignore case <br /><br />
    ''' Optional, Default: <c> False </c>
    ''' </param>
    Public Function contains(givenValue As String, Optional ignoreCase As Boolean = False) As Boolean
        Return If(ignoreCase, Items.Contains(givenValue, StringComparer.InvariantCultureIgnoreCase), Items.Contains(givenValue, StringComparer.InvariantCulture))
    End Function

    ''' <summary>Returns whether any of <paramref name="strLists"/> contains <paramref name="phrase"/> (case-sensitive)</summary>
    ''' <param name="strLists">The lists to search</param>
    ''' <param name="phrase">The value to search for</param>
    Public Shared Function IsInAny(strLists As strList(), phrase As String) As Boolean
        For i = 0 To strLists.Length - 1
            If strLists(i).contains(phrase) Then Return True
        Next
        Return False
    End Function

    ''' <summary>
    ''' Checks <paramref name="currentValue"/> against the list, ignoring case, and adds it
    ''' if it isn't there yet
    ''' </summary>
    '''
    ''' <param name="currentValue">The current value to be audited</param>
    '''
    ''' <returns>
    ''' <c> True </c> if the value was already in the list or is empty,
    ''' <c> False </c> if we just added it
    ''' </returns>
    Public Function chkDupes(currentValue As String) As Boolean
        If currentValue Is Nothing Then argIsNull(NameOf(currentValue)) : Return False
        If currentValue.Length = 0 Then Return True
        For Each value In Items
            If currentValue.Equals(value, StringComparison.InvariantCultureIgnoreCase) Then Return True
        Next
        Items.Add(currentValue)
        Return False
    End Function

    ''' <summary>
    ''' Returns one pair per item holding the items before and after it, with <c> "first" </c>
    ''' and <c> "last" </c> standing in at the ends. Throws when the list has fewer than two items.
    ''' </summary>
    Public Function getNeighborList() As List(Of KeyValuePair(Of String, String))
        Dim neighborList As New List(Of KeyValuePair(Of String, String)) From {New KeyValuePair(Of String, String)("first", Items(1))}
        For i = 1 To Items.Count - 2
            neighborList.Add(New KeyValuePair(Of String, String)(Items(i - 1), Items(i + 1)))
        Next
        neighborList.Add(New KeyValuePair(Of String, String)(Items(Items.Count - 2), "last"))
        Return neighborList
    End Function

    ''' <summary>
    ''' Replaces an item in a list of strings at the index of another given item.
    ''' Throws if <paramref name="indexOfText"/> isn't in the list.
    ''' </summary>
    ''' <param name="indexOfText">The text to be replaced</param>
    ''' <param name="newText">The replacement text</param>
    Public Sub replaceStrAtIndexOf(indexOfText As String, newText As String)
        Items(Items.IndexOf(indexOfText)) = newText
    End Sub
End Class