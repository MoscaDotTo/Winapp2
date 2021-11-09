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
''' Provides a few helpful methods for dissecting winapp2key objects
''' </summary>
Public Class winapp2KeyParameters
    ''' <summary>The extracted File/Registry path from the key </summary>
    Public Property PathString As String
    ''' <summary>The list of any file names (patterns) /registry keys from the key</summary>
    Public Property ArgsList As New List(Of String)
    ''' <summary> Holds the key's flag (RECURSE, REMOVESELF, etc)</summary>
    Public Property FlagString As String
    ''' <summary>The original KeyType of the iniKey used to create the winapp2key</summary>
    Public Property KeyType As String
    ''' <summary>The number associated with the KeyType</summary>
    Public Property KeyNum As String

    ''' <summary>Creates a new keyparams object from a given iniKey object</summary>
    ''' <param name="key">The iniKey to get parameters from</param>
    Public Sub New(key As iniKey)
        If key Is Nothing Then argIsNull(NameOf(key)) : Return
        KeyType = key.KeyType
        Dim splitKey = key.Value.Split(CChar("|"))
        ArgsList = New List(Of String)
        FlagString = ""
        Select Case key.KeyType
            Case "FileKey"
                KeyNum = key.KeyType.Replace("FileKey", "")
                If splitKey.Length > 1 Then
                    PathString = splitKey(0)
                    ArgsList.AddRange(splitKey(1).Split(CChar(";")))
                    FlagString = If(splitKey.Length >= 3, splitKey.Last, "None")
                Else
                    PathString = key.Value
                End If
            Case "ExcludeKey"
                KeyNum = key.KeyType.Replace("ExcludeKey", "")
                Select Case splitKey.Length
                    Case 2
                        PathString = splitKey(1)
                        FlagString = splitKey(0)
                    Case 3
                        PathString = splitKey(1)
                        ArgsList.AddRange(splitKey(2).Split(CChar(";")))
                        FlagString = splitKey(0)
                End Select
            Case "RegKey"
                KeyNum = key.KeyType.Replace("RegKey", "")
                PathString = splitKey(0)
                If splitKey.Length > 1 Then ArgsList.Add(splitKey(1))
        End Select
    End Sub

    ''' <summary>Reconstructs a FileKey to hold the format of FileKeyX=PATH|FILE;FILE;FILE....|FLAG</summary>
    ''' <param name="key">An iniKey to be reconstructed</param>
    ''' Also trims empty comments 
    Public Sub reconstructKey(ByRef key As iniKey)
        If key Is Nothing Then argIsNull(NameOf(key)) : Return
        Dim out = ""
        out += $"{PathString}{If(ArgsList.Count > 0, "|", "")}"
        If ArgsList.Count > 1 Then
            For i = 0 To ArgsList.Count - 2
                If Not ArgsList(i).Length = 0 Then out += ArgsList(i) & ";"
            Next
        End If
        If ArgsList.Count > 0 Then out += ArgsList.Last
        If Not FlagString = "None" Then out += "|" & FlagString
        ' Small edgecase for when empty parameters lead the pipe, resulting in the above loop leaving a single one for expectation of a next element
        out = out.Replace(";|", "|")
        key.Value = out
    End Sub

    ''' <summary>Constructs a new iniKey in an attempt to merge keys together</summary>
    ''' <param name="tmpKeyStr">The string to contain the new key text</param>
    Public Sub addArgs(ByRef tmpKeyStr As String)
        tmpKeyStr += $"{KeyType}{KeyNum}={PathString}|{ArgsList(0)}"
        If ArgsList.Count > 1 Then
            For i = 1 To ArgsList.Count - 1
                tmpKeyStr += $";{ArgsList(i)}"
            Next
        End If
    End Sub

    ''' <summary>Tracks params and flags from a winapp2key</summary>
    ''' <param name="paramList">The list of params observed</param>
    ''' <param name="flagList">The list of flags observed</param>
    Public Sub trackParamAndFlags(ByRef paramList As strList, ByRef flagList As strList)
        If paramList Is Nothing Then argIsNull(NameOf(paramList)) : Return
        If flagList Is Nothing Then argIsNull(NameOf(flagList)) : Return
        paramList.add(PathString)
        flagList.add(FlagString)
    End Sub
End Class