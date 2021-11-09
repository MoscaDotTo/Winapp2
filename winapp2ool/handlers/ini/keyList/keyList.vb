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
''' <summary> A handy wrapper object for lists of iniKeys </summary>
Public Class keyList

    ''' <summary> The list of <c> iniKeys </c> </summary>
    Public Property Keys As New List(Of iniKey)

    ''' <summary> The <c> KeyType </c> of the <c> iniKeys </c> contained in the <c> keyList </c> </summary>
    Public Property KeyType As String

    ''' <summary> Returns the number of keys in the keylist </summary>
    Public Function KeyCount() As Integer
        Return Keys.Count
    End Function

    ''' <summary> Creates a new (empty) <c> keyList </c> </summary>
    ''' <param name="kt"> The expected <c> keyType </c> of the <c> iniKeys </c> in the <c> keyList </c> <br /> Optional, Default: <c> "" </c> </param>
    Public Sub New(Optional kt As String = "")
        Keys = New List(Of iniKey)
        KeyType = kt
    End Sub

    ''' <summary> Creates a new <c> keyList </c> using an existing list of <c> iniKeys </c> </summary>
    ''' <param name="kl"> A list of <c> iniKeys </c> to be inserted into the <c> keylist </c> </param>
    Public Sub New(kl As List(Of iniKey))
        If kl Is Nothing Then argIsNull(NameOf(kl)) : Return
        KeyType = If(kl.Count > 0, kl(0).KeyType, "")
    End Sub

    ''' <summary> Conditionally adds an <c> iniKey </c> into the <c> keyList </c> </summary>
    ''' <param name="key"> The <c> iniKey </c>to be added </param>
    ''' <param name="cond"> The condition under which to add the <c> iniKey </c> </param>
    Public Sub add(key As iniKey, Optional cond As Boolean = True)
        If cond Then Keys.Add(key)
    End Sub

    ''' <summary> Adds a list of <c> iniKeys </c> to the <c> keyList </c> </summary>
    ''' <param name="kl">The list of <c> iniKeys </c> to be added </param>
    Public Sub add(kl As List(Of iniKey))
        If kl Is Nothing Then argIsNull(NameOf(kl)) : Return
        kl.ForEach(Sub(key) Keys.Add(key))
    End Sub

    ''' <summary> Removes an <c> iniKey </c> from the <c> keyList </c> </summary>
    ''' <param name="key">The <c> iniKey </c> to be removed </param>
    Public Sub remove(key As iniKey)
        Me.Keys.Remove(key)
    End Sub

    ''' <summary>Removes a list of <c> iniKeys </c>from the <c> keyList </c> </summary>
    ''' <param name="kl"> The list of <c> iniKeys </c> to be removed </param>
    Public Sub remove(kl As List(Of iniKey))
        If kl Is Nothing Then argIsNull(NameOf(kl)) : Return
        kl.ForEach(Sub(key) remove(key))
    End Sub

    ''' <summary> Removes the <c>iniKey </c>at the provided index provided by <c> <paramref name="ind"/> </c> from the <c> keyList </c> </summary>
    ''' <param name="ind"> The index of the <c> iniKey </c> to be removed </param>
    Public Sub remove(ind As Integer)
        Keys.RemoveAt(ind)
    End Sub

    ''' <summary> Returns a <c> Boolean </c> indicating whether or not the <c> keyType  </c>of the list matches the <c> keyType </c> provided by <c> <paramref name="type"/> </c> </summary>
    ''' <param name="type"> The String against which to match the <c> keylist's keyType </c> </param>
    Public Function typeIs(type As String) As Boolean
        Return If(KeyType.Length = 0, Keys(0).KeyType, KeyType) = type
    End Function

    ''' <summary> Returns a <c> keyList </c> as a <c> strList </c> </summary>
    ''' <param name="onlyGetVals"> <c> True </c> if only requesting the values from the keys. <br /> Optional, Default: <c> False </c> </param>
    Public Function toStrLst(Optional onlyGetVals As Boolean = False) As strList
        Dim out As New strList
        Keys.ForEach(Sub(key) out.add(If(onlyGetVals, key.Value, key.toString)))
        Return out
    End Function

    ''' <summary> Removes the last element in the <c> keyList </c> if it exists </summary>
    Public Sub removeLast()
        If Keys.Count > 0 Then Keys.Remove(Keys.Last)
    End Sub

    ''' <summary> Renumber keys according to the sorted state of the values provided by <c> <paramref name="sortedKeyValues"/> </c> </summary>
    ''' <param name="sortedKeyValues"> Some target sorted state of the <c> Values </c> of the <c> iniKeys </c> in the <c> keyList </c> </param>
    Public Sub renumberKeys(sortedKeyValues As strList)
        If sortedKeyValues Is Nothing Then argIsNull(NameOf(sortedKeyValues)) : Return
        gLog("Renumbering keys", indent:=True)
        For i = 0 To Me.KeyCount - 1
            Keys(i).Name = Keys(i).KeyType & i + 1
            Keys(i).Value = sortedKeyValues.Items(i)
        Next
    End Sub

    ''' <summary> Returns a <c> List(Of Integer) </c> containing the line numbers from the <c> iniKeys </c> in the <c> keyList </c> </summary>
    Public Function lineNums() As List(Of Integer)
        Dim out As New List(Of Integer)
        Keys.ForEach(Sub(key) out.Add(key.LineNumber))
        Return out
    End Function
End Class