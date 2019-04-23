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
''' Provides a few helpful methods for dissecting winapp2key objects
''' </summary>
Public Class winapp2KeyParameters
    Public pathString As String = ""
    Public argsList As New List(Of String)
    Public flagString As String = ""
    Public keyType As String = ""
    Public keyNum As String = ""

    ''' <summary>Creates a new keyparams object from a given iniKey object</summary>
    ''' <param name="key">The iniKey to get parameters from</param>
    Public Sub New(key As iniKey)
        keyType = key.KeyType
        Dim splitKey As String() = key.Value.Split(CChar("|"))
        Select Case key.KeyType
            Case "FileKey"
                keyNum = key.KeyType.Replace("FileKey", "")
                If splitKey.Count > 1 Then
                    pathString = splitKey(0)
                    argsList.AddRange(splitKey(1).Split(CChar(";")))
                    flagString = If(splitKey.Count >= 3, splitKey.Last, "None")
                Else
                    pathString = key.Value
                End If
            Case "ExcludeKey"
                Select Case splitKey.Count
                    Case 2
                        pathString = splitKey(1)
                        flagString = splitKey(0)
                    Case 3
                        pathString = splitKey(1)
                        argsList.AddRange(splitKey(2).Split(CChar(";")))
                        flagString = splitKey(0)
                End Select
            Case "RegKey"
                pathString = splitKey(0)
                If splitKey.Count > 1 Then argsList.Add(splitKey(1))
        End Select
    End Sub

    ''' <summary>Reconstructs a FileKey to hold the format of FileKeyX=PATH|FILE;FILE;FILE....|FLAG</summary>
    ''' <param name="key">An iniKey to be reconstructed</param>
    ''' Also trims empty comments 
    Public Sub reconstructKey(ByRef key As iniKey)
        Dim out As String = ""
        out += $"{pathString}{If(argsList.Count > 0, "|", "")}"
        If argsList.Count > 1 Then
            For i As Integer = 0 To argsList.Count - 2
                If Not argsList(i) = "" Then out += argsList(i) & ";"
            Next
        End If
        If argsList.Count > 0 Then out += argsList.Last
        If Not flagString = "None" Then out += "|" & flagString
        key.Value = out
    End Sub

    ''' <summary>Constructs a new iniKey in an attempt to merge keys together</summary>
    ''' <param name="tmpKeyStr">The string to contain the new key text</param>
    Public Sub addArgs(ByRef tmpKeyStr As String)
        appendStrs({$"{keyType}{keyNum}=", $"{pathString}|", argsList(0)}, tmpKeyStr)
        If argsList.Count > 1 Then
            For i = 1 To argsList.Count - 1
                tmpKeyStr += $";{argsList(i)}"
            Next
        End If
    End Sub

    ''' <summary>Tracks params and flags from a winapp2key</summary>
    ''' <param name="paramList">The list of params observed</param>
    ''' <param name="flagList">The list of flags observed</param>
    Public Sub trackParamAndFlags(ByRef paramList As strList, ByRef flagList As strList)
        paramList.add(pathString)
        flagList.add(flagString)
    End Sub
End Class