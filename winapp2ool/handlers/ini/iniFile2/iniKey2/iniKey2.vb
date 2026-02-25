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
Imports System.Text.RegularExpressions

''' <summary>
''' An object representing a name=value pair from an ini file
''' </summary>
Public Class iniKey2

    Private _name As String
    Private _keyType As String

    ''' <summary>
    ''' The name of the key: any text on the left side of the '='.
    ''' Setting this also updates <c>KeyType</c>.
    ''' </summary>
    Public Property Name As String
        Get
            Return _name
        End Get
        Set(value As String)
            _name = value
            _keyType = stripNums(value)
        End Set
    End Property

    ''' <summary>
    ''' The value of the key: any text on the right side of the '='
    ''' </summary>
    Public Property Value As String

    ''' <summary>
    ''' The type of the key: the Name with digits removed. Updated automatically when Name is set.
    ''' </summary>
    Public ReadOnly Property KeyType As String
        Get
            Return _keyType
        End Get
    End Property

    ''' <summary>The line number from which this key was originally read</summary>
    Public ReadOnly Property LineNumber As Integer

    ''' <summary>Returns whether the key's Name equals the given string</summary>
    ''' <param name="n">The string to compare against</param>
    ''' <param name="ignoreCase">When True, comparison is case-insensitive</param>
    Public Function nameIs(n As String, Optional ignoreCase As Boolean = False) As Boolean
        Return If(ignoreCase, Name.Equals(n, StringComparison.InvariantCultureIgnoreCase), Name = n)
    End Function

    ''' <summary>Returns whether the key's KeyType equals the given string</summary>
    ''' <param name="t">The string to compare against</param>
    ''' <param name="ignoreCase">When True, comparison is case-insensitive</param>
    Public Function typeIs(t As String, Optional ignoreCase As Boolean = False) As Boolean
        Return If(ignoreCase, KeyType.Equals(t, StringComparison.InvariantCultureIgnoreCase), KeyType = t)
    End Function

    ''' <summary>Returns whether the key's Value contains the given string</summary>
    ''' <param name="txt">The string to search for</param>
    ''' <param name="ignoreCase">When True, search is case-insensitive</param>
    Public Function vHas(txt As String, Optional ignoreCase As Boolean = False) As Boolean
        Return If(ignoreCase, Value.IndexOf(txt, 0, StringComparison.CurrentCultureIgnoreCase) > -1, Value.Contains(txt))
    End Function

    ''' <summary>Returns whether the key's Value contains any of the given strings</summary>
    ''' <param name="txts">The strings to search for</param>
    ''' <param name="ignoreCase">When True, search is case-insensitive</param>
    Public Function vHasAny(txts As String(), Optional ignoreCase As Boolean = False) As Boolean
        If txts Is Nothing Then argIsNull(NameOf(txts)) : Return False
        For Each txt In txts
            If vHas(txt, ignoreCase) Then Return True
        Next
        Return False
    End Function

    ''' <summary>Returns whether the key's Value equals the given string</summary>
    ''' <param name="txt">The string to compare against</param>
    ''' <param name="ignoreCase">When True, comparison is case-insensitive</param>
    Public Function vIs(txt As String, Optional ignoreCase As Boolean = False) As Boolean
        Return If(ignoreCase, Value.Equals(txt, StringComparison.InvariantCultureIgnoreCase), Value = txt)
    End Function

    Private Function stripNums(keyName As String) As String
        Return New Regex("[\d]").Replace(keyName, "")
    End Function

    ''' <summary>Returns whether this key's Name matches the given key's Name (case-insensitive)</summary>
    ''' <param name="key">The key to compare against</param>
    Public Function compareNames(key As iniKey2) As Boolean
        If key Is Nothing Then argIsNull(NameOf(key)) : Return False
        Return nameIs(key.Name, True)
    End Function

    ''' <summary>Returns whether this key's Value matches the given key's Value (case-insensitive)</summary>
    ''' <param name="key">The key to compare against</param>
    Public Function compareValues(key As iniKey2) As Boolean
        If key Is Nothing Then argIsNull(NameOf(key)) : Return False
        Return vIs(key.Value, True)
    End Function

    ''' <summary>Returns whether this key's KeyType matches the given key's KeyType (case-insensitive)</summary>
    ''' <param name="key">The key to compare against</param>
    Public Function compareTypes(key As iniKey2) As Boolean
        If key Is Nothing Then argIsNull(NameOf(key)) : Return False
        Return typeIs(key.KeyType, True)
    End Function

    ''' <summary>Creates an iniKey2 from a name=value string</summary>
    ''' <param name="line">A string in the format name=value</param>
    ''' <param name="count">The line number for this key</param>
    Public Sub New(ByVal line As String, Optional ByVal count As Integer = 0)
        If line Is Nothing Then argIsNull(NameOf(line)) : Return
        Dim eqPos = line.IndexOf("="c)
        If eqPos >= 0 Then
            LineNumber = count
            Dim lhs = line.Substring(0, eqPos)
            Dim rhs = line.Substring(eqPos + 1)
            Select Case True
                Case lhs.Length <> 0 And rhs.Length <> 0
                    Name = lhs : Value = rhs   ' KeyType updated automatically by Name setter
                Case lhs.Length = 0 And rhs.Length <> 0
                    _name = "KeyTypeNotGiven" : Value = rhs : _keyType = "Error"
                Case lhs.Length <> 0 And rhs.Length = 0
                    _name = lhs : Value = "" : _keyType = "DeleteMe"
            End Select
        Else
            _name = line : Value = "" : _keyType = "DeleteMe"
        End If
    End Sub

    ''' <summary>Returns the key in name=value format</summary>
    Public Overrides Function ToString() As String
        Return $"{Name}={Value}"
    End Function

End Class
