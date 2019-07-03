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
''' A handy wrapper object for lists of iniKeys
''' </summary>
Public Class keyList

    ''' <summary> The list of iniKeys </summary>
    Public Property Keys As New List(Of iniKey)

    ''' <summary> The KeyType of keys contained in the list</summary>
    Public Property KeyType As String

    ''' <summary>Returns the number of keys in the keylist</summary>
    Public Function KeyCount() As Integer
        Return Keys.Count
    End Function

    ''' <summary>Creates a new (empty) keylist</summary>
    ''' <param name="kt">Optional String containing the expected KeyType of the keys in the list</param>
    Public Sub New(Optional kt As String = "")
        Keys = New List(Of iniKey)
        KeyType = kt
    End Sub

    ''' <summary>Creates a new keylist using an existing list of iniKeys</summary>
    ''' <param name="kl">A list of iniKeys to be inserted into the keylist</param>
    Public Sub New(kl As List(Of iniKey))
        Keys = kl
        KeyType = If(Keys.Count > 0, Keys(0).KeyType, "")
    End Sub

    ''' <summary>Conditionally adds a key to the keylist</summary>
    ''' <param name="key">The key to be added</param>
    ''' <param name="cond">The condition under which to add the key</param>
    Public Sub add(key As iniKey, Optional cond As Boolean = True)
        If cond Then Keys.Add(key)
    End Sub

    ''' <summary>Adds a list of iniKeys to the keylist</summary>
    ''' <param name="kl">The list to be added</param>
    Public Sub add(kl As List(Of iniKey))
        kl.ForEach(Sub(key) Keys.Add(key))
    End Sub

    ''' <summary>Removes a key from the keylist</summary>
    ''' <param name="key">The key to be removed</param>
    Public Sub remove(key As iniKey)
        Me.Keys.Remove(key)
    End Sub

    ''' <summary>Removes a list of keys from the keylist</summary>
    ''' <param name="kl">The list of keys to be removed</param>
    Public Sub remove(kl As List(Of iniKey))
        kl.ForEach(Sub(key) remove(key))
    End Sub

    ''' <summary> Removes the key at the provided index</summary>
    ''' <param name="ind">The index of the key to be removed</param>
    Public Sub remove(ind As Integer)
        Keys.RemoveAt(ind)
    End Sub

    ''' <summary>Returns whether or not the keyType of the list matches the input String</summary>
    ''' <param name="type">The String against which to match the keylist's type</param>
    Public Function typeIs(type As String) As Boolean
        Return If(KeyType = "", Keys(0).KeyType, KeyType) = type
    End Function

    ''' <summary>Returns a keyList as a strList</summary>
    ''' <param name="onlyGetVals">True if only requesting the values from the keys. Defaukt: false</param>
    Public Function toStrLst(Optional onlyGetVals As Boolean = False) As strList
        Dim out As New strList
        Keys.ForEach(Sub(key) out.add(If(onlyGetVals, key.Value, key.toString)))
        Return out
    End Function

    ''' <summary>Removes the last element in the key list if it exists</summary>
    Public Sub removeLast()
        If Keys.Count > 0 Then Keys.Remove(Keys.Last)
    End Sub

    ''' <summary>Renumber keys according to the sorted state of the values</summary>
    ''' <param name="sortedKeyValues"></param>
    Public Sub renumberKeys(sortedKeyValues As strList)
        gLog("Renumbering keys", indent:=True)
        For i As Integer = 0 To Me.KeyCount - 1
            Keys(i).Name = Keys(i).KeyType & i + 1
            Keys(i).Value = sortedKeyValues.Items(i)
        Next
    End Sub

    ''' <summary>Returns a list of integers containing the line numbers from the keylist</summary>
    Public Function lineNums() As List(Of Integer)
        Dim out As New List(Of Integer)
        Keys.ForEach(Sub(key) out.Add(key.LineNumber))
        Return out
    End Function
End Class