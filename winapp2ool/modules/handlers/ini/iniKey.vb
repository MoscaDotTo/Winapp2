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
Imports System.Text.RegularExpressions
''' <summary>
''' An object representing the name value pairs that make up iniSections
''' </summary>
Public Class iniKey
    Private _name As String
    Private _value As String
    Public lineNumber As Integer
    Private _keyType As String

    ''' <summary>
    ''' The value of the iniKey: any text on the right side of the '='
    ''' </summary>
    ''' <returns></returns>
    Public Property Value As String
        Get
            Return _value
        End Get
        Set(value As String)
            _value = value
        End Set
    End Property

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <returns></returns>
    Public Property KeyType As String
        Get
            Return _keyType
        End Get
        Set(value As String)
            _keyType = value
        End Set
    End Property

    Public Property Name As String
        Get
            Return _name
        End Get
        Set(value As String)
            _name = value
        End Set
    End Property

    ''' <summary>
    ''' Assigns or overwrites the value of the iniKey
    ''' </summary>
    ''' <param name="newVal">The String containing the new value</param>
    Public Sub setValue(newVal As String)
        Value = newVal
    End Sub

    ''' <summary>
    ''' Assigns or overwrites the name of the iniKey
    ''' </summary>
    ''' <param name="newName">The String containing the new value</param>
    Public Sub setName(newName As String)
        Name = newName
    End Sub

    ''' <summary>
    ''' Assigns or overwrites the keyType of the iniKey 
    ''' </summary>
    ''' <param name="newType">The String containing the new value</param>
    Public Sub setKeyType(newType As String)
        KeyType = newType
    End Sub

    ''' <summary>
    ''' Returns whether or not an iniKey's name is equal to a given value
    ''' </summary>
    ''' <param name="n">The String to check equality for</param>
    ''' <param name="ignoreCase">Optional boolean specifying whether or not the casing of the strings should be ignored (default false)</param>
    ''' <returns></returns>
    Public Function nameIs(n As String, Optional ignoreCase As Boolean = False) As Boolean
        Return If(ignoreCase, Name.Equals(n, StringComparison.InvariantCultureIgnoreCase), Name = n)
        Name = n
    End Function

    ''' <summary>
    ''' Returns whether or not an iniKey's type is equal to a given value
    ''' </summary>
    ''' <param name="t">The string to check equality for</param>
    ''' <param name="ignoreCase">Optional boolean specifying whether or not the casing of the strings should be ignored (default false)</param>
    ''' <returns></returns>
    Public Function typeIs(t As String, Optional ignoreCase As Boolean = False) As Boolean
        Return If(ignoreCase, KeyType.Equals(t, StringComparison.InvariantCultureIgnoreCase), KeyType = t)
    End Function

    ''' <summary>
    ''' Returns whether or not an iniKey object's value contains a given string with conditional case sensitivity
    ''' </summary>
    ''' <param name="txt">The string to search for</param>
    ''' <param name="ignoreCase">Optional boolean specifying whether or not the casing of the strings should be ignored (default false)</param>
    ''' <returns></returns>
    Public Function vHas(txt As String, Optional ignoreCase As Boolean = False) As Boolean
        Return If(ignoreCase, Value.IndexOf(txt, 0, StringComparison.CurrentCultureIgnoreCase) > -1, Value.Contains(txt))
    End Function

    ''' <summary>
    ''' Returns whether or not an iniKey object's value contains any of a given array of strings with conditional case sensitivity
    ''' </summary>
    ''' <param name="txts">The array of search strings</param>
    ''' <param name="ignoreCase">Optional boolean specifying whether or not the casing of the strings should be ignored (default false)</param>
    ''' <returns></returns>
    Public Function vHasAny(txts As String(), Optional ignoreCase As Boolean = False) As Boolean
        For Each txt In txts
            If vHas(txt, ignoreCase) Then Return True
        Next
        Return False
    End Function

    ''' <summary>
    ''' Returns whether or not an iniKey object's value is equal to a given string with conditional case sensitivity
    ''' </summary>
    ''' <param name="txt">The string to be searched for</param>
    ''' <param name="ignoreCase">Optional boolean specifying whether or not the casing of the strings should be ignored (default false)</param>
    ''' <returns></returns>
    Public Function vIs(txt As String, Optional ignoreCase As Boolean = False) As Boolean
        Return If(ignoreCase, Value.Equals(txt, StringComparison.InvariantCultureIgnoreCase), Value = txt)
    End Function

    ''' <summary>
    ''' Returns an iniKey object's keyName field with numbers removed
    ''' </summary>
    ''' <param name="keyName">The string containing the iniKey's keyname</param>
    ''' <returns></returns>
    Private Function stripNums(keyName As String) As String
        Return New Regex("[\d]").Replace(keyName, "")
    End Function

    ''' <summary>
    ''' Compares the names of two iniKeys and returns whether or not they match
    ''' </summary>
    ''' <param name="key">The iniKey to be compared to</param>
    ''' <returns></returns>
    Public Function compareNames(key As iniKey) As Boolean
        Return nameIs(key.Name, True)
    End Function

    ''' <summary>
    ''' Compares the values of two iniKeys and returns whether or not they match
    ''' </summary>
    ''' <param name="key">The iniKey to be compared to</param>
    ''' <returns></returns>
    Public Function compareValues(key As iniKey) As Boolean
        Return vIs(key.Value, True)
    End Function

    ''' <summary>
    ''' Compares the types of two iniKeys and returns whether or not they match
    ''' </summary>
    ''' <param name="key">The iniKey to be compared to</param>
    ''' <returns></returns>
    Public Function compareTypes(key As iniKey) As Boolean
        Return typeIs(key.KeyType, True)
    End Function

    ''' <summary>
    ''' Create an iniKey object from a string containing a name value pair
    ''' </summary>
    ''' <param name="line">A string in the format name=value</param>
    ''' <param name="count">The line number for the string</param>
    Public Sub New(ByVal line As String, Optional ByVal count As Integer = 0)
        If line.Contains("=") Then
            Dim splitLine As String() = line.Split(CChar("="))
            lineNumber = count
            Select Case True
                Case splitLine(0) <> "" And splitLine(1) <> ""
                    Name = splitLine(0)
                    Value = splitLine(1)
                    KeyType = stripNums(Name)
                Case splitLine(0) = "" And splitLine(1) <> ""
                    Name = "KeyTypeNotGiven"
                    Value = splitLine(1)
                    KeyType = "Error"
                Case splitLine(0) <> "" And splitLine(1) = ""
                    Name = "DeleteMe"
                    Value = "This key was not provided with a value and will be deleted. The user should never see this, if you do, please report it as a bug on GitHub"
                    KeyType = "DeleteMe"
            End Select
        Else
            Name = line
            Value = ""
            KeyType = "Error"
        End If
    End Sub

    ''' <summary>
    ''' Returns the key in name=value format as a String
    ''' </summary>
    ''' <returns></returns>
    Public Overrides Function toString() As String
        Return $"{Name}={Value}"
    End Function
End Class