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
    Public keys As List(Of iniKey)
    Public keyType As String

    ''' <summary>
    ''' Creates a new (empty) keylist
    ''' </summary>
    ''' <param name="kt">Optional String containing the expected KeyType of the keys in the list</param>
    Public Sub New(Optional kt As String = "")
        keys = New List(Of iniKey)
        keyType = kt
    End Sub

    ''' <summary>
    ''' Creates a new keylist using an existing list of iniKeys
    ''' </summary>
    ''' <param name="kl">A list of iniKeys to be inserted into the keylist</param>
    Public Sub New(kl As List(Of iniKey))
        keys = kl
        keyType = If(keys.Count > 0, keys(0).keyType, "")
    End Sub

    ''' <summary>
    ''' Conditionally adds a key to the keylist
    ''' </summary>
    ''' <param name="key">The key to be added</param>
    ''' <param name="cond">The condition under which to add the key</param>
    Public Sub add(key As iniKey, Optional cond As Boolean = True)
        If cond Then keys.Add(key)
    End Sub

    ''' <summary>
    ''' Adds a list of iniKeys to the keylist
    ''' </summary>
    ''' <param name="kl">The list to be added</param>
    Public Sub add(kl As List(Of iniKey))
        kl.ForEach(Sub(key) keys.Add(key))
    End Sub

    ''' <summary>
    ''' Removes a key from the keylist
    ''' </summary>
    ''' <param name="key">The key to be removed</param>
    Public Sub remove(key As iniKey)
        Me.keys.Remove(key)
    End Sub

    ''' <summary>
    ''' Removes a list of keys from the keylist
    ''' </summary>
    ''' <param name="kl">The list of keys to be removed</param>
    Public Sub remove(kl As List(Of iniKey))
        kl.ForEach(Sub(key) remove(key))
    End Sub

    ''' <summary>
    ''' Returns the number of keys in the keylist
    ''' </summary>
    ''' <returns></returns>
    Public Function keyCount() As Integer
        Return keys.Count
    End Function

    ''' <summary>
    ''' Returns whether or not the keyType of the list matches the input String
    ''' </summary>
    ''' <param name="type">The String against which to match the keylist's type</param>
    ''' <returns></returns>
    Public Function typeIs(type As String) As Boolean
        Return If(keyType = "", keys(0).keyType, keyType) = type
    End Function

    ''' <summary>
    ''' Returns the keylist in the form of a list of Strings
    ''' </summary>
    ''' <param name="onlyGetVals">Optional Boolean specifying whether or not the function should return only the values from the keys</param>
    ''' <returns></returns>
    Public Function toListOfStr(Optional onlyGetVals As Boolean = False) As List(Of String)
        Dim out As New List(Of String)
        keys.ForEach(Sub(key) out.Add(If(onlyGetVals, key.value, key.toString)))
        Return out
    End Function

    ''' <summary>
    ''' Removes the last element in the key list if it exists
    ''' </summary>
    Public Sub removeLast()
        If keys.Count > 0 Then keys.Remove(keys.Last)
    End Sub

    ''' <summary>
    ''' Renumber keys according to the sorted state of the values
    ''' </summary>
    ''' <param name="sortedKeyValues"></param>
    Public Sub renumberKeys(sortedKeyValues As List(Of String))
        For i As Integer = 0 To Me.keyCount - 1
            keys(i).name = keys(i).keyType & i + 1
            keys(i).value = sortedKeyValues(i)
        Next
    End Sub

    ''' <summary>
    ''' Returns a list of integers containing the line numbers from the keylist
    ''' </summary>
    ''' <returns></returns>
    Public Function lineNums() As List(Of Integer)
        Dim out As New List(Of Integer)
        keys.ForEach(Sub(key) out.Add(key.lineNumber))
        Return out
    End Function
End Class