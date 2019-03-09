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
    Public name As String
    Public value As String
    Public lineNumber As Integer
    Public keyType As String

    ''' <summary>
    ''' Returns whether or not an iniKey's name is equal to a given value
    ''' </summary>
    ''' <param name="n">The string to check equality for</param>
    ''' <param name="ignoreCase">Optional boolean specifying whether or not the casing of the strings should be ignored (default false)</param>
    ''' <returns></returns>
    Public Function nameIs(n As String, Optional ignoreCase As Boolean = False) As Boolean
        Return If(ignoreCase, name.Equals(n, StringComparison.InvariantCultureIgnoreCase), name = n)
        name = n
    End Function

    ''' <summary>
    ''' Returns whether or not an iniKey's type is equal to a given value
    ''' </summary>
    ''' <param name="t">The string to check equality for</param>
    ''' <param name="ignoreCase">Optional boolean specifying whether or not the casing of the strings should be ignored (default false)</param>
    ''' <returns></returns>
    Public Function typeIs(t As String, Optional ignoreCase As Boolean = False) As Boolean
        Return If(ignoreCase, keyType.Equals(t, StringComparison.InvariantCultureIgnoreCase), keyType = t)
    End Function

    ''' <summary>
    ''' Returns whether or not an iniKey object's value contains a given string with conditional case sensitivity
    ''' </summary>
    ''' <param name="txt">The string to search for</param>
    ''' <param name="ignoreCase">Optional boolean specifying whether or not the casing of the strings should be ignored (default false)</param>
    ''' <returns></returns>
    Public Function vHas(txt As String, Optional ignoreCase As Boolean = False) As Boolean
        Return If(ignoreCase, value.IndexOf(txt, 0, StringComparison.CurrentCultureIgnoreCase) > -1, value.Contains(txt))
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
        Return If(ignoreCase, value.Equals(txt, StringComparison.InvariantCultureIgnoreCase), value = txt)
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
        Return nameIs(key.name, True)
    End Function

    ''' <summary>
    ''' Compares the values of two iniKeys and returns whether or not they match
    ''' </summary>
    ''' <param name="key">The iniKey to be compared to</param>
    ''' <returns></returns>
    Public Function compareValues(key As iniKey) As Boolean
        Return vIs(key.value, True)
    End Function

    ''' <summary>
    ''' Compares the types of two iniKeys and returns whether or not they match
    ''' </summary>
    ''' <param name="key">The iniKey to be compared to</param>
    ''' <returns></returns>
    Public Function compareTypes(key As iniKey) As Boolean
        Return typeIs(key.keyType, True)
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
                    name = splitLine(0)
                    value = splitLine(1)
                    keyType = stripNums(name)
                Case splitLine(0) = "" And splitLine(1) <> ""
                    name = "KeyTypeNotGiven"
                    value = splitLine(1)
                    keyType = "Error"
                Case splitLine(0) <> "" And splitLine(1) = ""
                    name = "DeleteMe"
                    value = "This key was not provided with a value and will be deleted. The user should never see this, if you do, please report it as a bug on GitHub"
                    keyType = "DeleteMe"
            End Select
        Else
            name = line
            value = ""
            keyType = "Error"
        End If
    End Sub

    ''' <summary>
    ''' Returns the key in name=value format as a String
    ''' </summary>
    ''' <returns></returns>
    Public Overrides Function toString() As String
        Return $"{name}={value}"
    End Function
End Class