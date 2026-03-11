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
''' <summary>An ordered, case-insensitive-keyed collection of <c>iniKey2</c> objects</summary>
Public Class iniKeyCollection
    Implements IEnumerable(Of iniKey2)

    Private ReadOnly _ordered As New List(Of iniKey2)
    Private ReadOnly _byName As New Dictionary(Of String, iniKey2)(StringComparer.OrdinalIgnoreCase)
    Private ReadOnly _byType As New Dictionary(Of String, List(Of iniKey2))(StringComparer.OrdinalIgnoreCase)

    Private Shared ReadOnly EmptyKeyList As IReadOnlyList(Of iniKey2) = New List(Of iniKey2)().AsReadOnly()

    ''' <summary>The number of keys in the collection</summary>
    Public ReadOnly Property Count As Integer
        Get
            Return _ordered.Count
        End Get
    End Property

    ''' <summary>
    ''' Adds a key to the collection. Duplicate names are allowed — all occurrences appear
    ''' in enumeration order and in <c>GetByType</c> results. <c>GetKey</c> returns the
    ''' first occurrence for any given name.
    ''' </summary>
    ''' <param name="key">The key to add</param>
    Public Sub Add(key As iniKey2)
        If key Is Nothing Then argIsNull(NameOf(key)) : Return
        _ordered.Add(key)
        If Not _byName.ContainsKey(key.Name) Then _byName.Add(key.Name, key)
        Dim bucket As List(Of iniKey2) = Nothing
        If Not _byType.TryGetValue(key.KeyType, bucket) Then
            bucket = New List(Of iniKey2)
            _byType.Add(key.KeyType, bucket)
        End If
        bucket.Add(key)
    End Sub

    ''' <summary>Returns whether a key with the given name exists in the collection</summary>
    ''' <param name="name">The key name to search for (case-insensitive)</param>
    Public Function Contains(name As String) As Boolean
        If name Is Nothing Then argIsNull(NameOf(name)) : Return False
        Return _byName.ContainsKey(name)
    End Function

    ''' <summary>Returns the key with the given name, or <c>Nothing</c> if not found</summary>
    ''' <param name="name">The key name to look up (case-insensitive)</param>
    Public Function GetKey(name As String) As iniKey2
        If name Is Nothing Then argIsNull(NameOf(name)) : Return Nothing
        Dim result As iniKey2 = Nothing
        _byName.TryGetValue(name, result)
        Return result
    End Function

    ''' <summary>
    ''' Returns all keys whose <c>KeyType</c> equals <paramref name="keyType"/> (case-insensitive).
    ''' Returns an empty read-only list if no keys of that type exist.
    ''' The returned list is the live internal bucket — do not mutate it.
    ''' </summary>
    ''' <param name="keyType">The key type to look up, e.g. "FileKey" or "RegKey"</param>
    Public Function GetByType(keyType As String) As IReadOnlyList(Of iniKey2)
        If keyType Is Nothing Then argIsNull(NameOf(keyType)) : Return EmptyKeyList
        Dim result As List(Of iniKey2) = Nothing
        If _byType.TryGetValue(keyType, result) Then Return result
        Return EmptyKeyList
    End Function

    ''' <summary>
    ''' Removes the given key from the collection. If this key was the <c>GetKey</c>-indexed
    ''' occurrence for its name, the index is updated to the next remaining key with that name.
    ''' </summary>
    '''
    ''' <param name="key">
    ''' The key to remove
    ''' </param>
    Public Sub Remove(key As iniKey2)

        If key Is Nothing Then argIsNull(NameOf(key)) : Return
        _ordered.Remove(key)

        Dim indexed As iniKey2 = Nothing
        If _byName.TryGetValue(key.Name, indexed) AndAlso ReferenceEquals(indexed, key) Then
            _byName.Remove(key.Name)
            For Each remaining In _ordered
                If remaining.Name.Equals(key.Name, StringComparison.OrdinalIgnoreCase) Then
                    _byName.Add(remaining.Name, remaining)
                    Exit For
                End If
            Next
        End If

        Dim bucket As List(Of iniKey2) = Nothing
        If _byType.TryGetValue(key.KeyType, bucket) Then bucket.Remove(key)

    End Sub

    ''' <summary>
    ''' Returns a copy of the ordered key list
    ''' </summary>
    Public Function ToList() As List(Of iniKey2)

        Return New List(Of iniKey2)(_ordered)

    End Function

    ''' <summary>
    ''' Creates an empty collection
    ''' </summary>
    Public Sub New()

    End Sub

    ''' <summary>
    ''' Creates a collection pre-populated from the given sequence
    ''' </summary>
    '''
    ''' <param name="keys">
    ''' The keys to add (first-write-wins for duplicate names)
    ''' </param>
    Public Sub New(keys As IEnumerable(Of iniKey2))

        If keys Is Nothing Then argIsNull(NameOf(keys)) : Return

        For Each key In keys

            Add(key)

        Next

    End Sub

    Public Function GetEnumerator() As IEnumerator(Of iniKey2) Implements IEnumerable(Of iniKey2).GetEnumerator

        Return _ordered.GetEnumerator()

    End Function

    Private Function GetEnumeratorNonGeneric() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator

        Return _ordered.GetEnumerator()

    End Function

End Class
