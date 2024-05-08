'    Copyright (C) 2018-2024 Hazel Ward
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
Imports Microsoft.Win32
Imports System.Text.RegularExpressions
Imports System.Text.RegularExpressions.Regex
''' <summary>
''' This module holds any functions winapp2ool might require for accessing and manipulating the windows registry
''' </summary>
Module RegistryHelper
    ''' <summary>Returns the requested key or subkey from the HKEY_LOCAL_MACHINE registry hive</summary>
    ''' <param name="subkey">An optional string specifying the path to a subkey in the hive</param>
    Public Function getLMKey(Optional subkey As String = "") As RegistryKey
        Return Registry.LocalMachine.OpenSubKey(subkey)
    End Function

    ''' <summary>Returns the requested key or subkey from the HKEY_CLASSES_ROOT registry hive</summary>
    ''' <param name="subkey">An optional string specifying the path to a subkey in the hive</param>
    Public Function getCRKey(Optional subkey As String = "") As RegistryKey
        Return Registry.ClassesRoot.OpenSubKey(subkey)
    End Function

    ''' <summary>Returns the requested key or subkey from the HKEY_CURRENT_USER registry hive</summary>
    ''' <param name="subkey">An optional string specifying the path to a subkey in the hive</param>
    Public Function getCUKey(Optional subkey As String = "") As RegistryKey
        Return Registry.CurrentUser.OpenSubKey(subkey)
    End Function

    ''' <summary>Returns the requested key or subkey from the HKEY_USERS registry hive</summary>
    ''' <param name="subkey">An optional string specifying the path to a subkey in the hive</param>
    Public Function getUserKey(Optional subkey As String = "") As RegistryKey
        Return Registry.Users.OpenSubKey(subkey)
    End Function

    Public Function getWildcardKeys(path As String) As strList
        Dim out As New strList
        Dim tmp As New strList
        Dim tmp2 As New strList
        ' This will be the path split into pieces everywhere there's a wildcard
        Dim splitPath = path.Split(CChar("*"))
        cwl("testing")
        ' This will be anything that would return under the first wildcard
        cwl($"Finding all subkeys of {splitPath(0)}")
        Dim baseList = getSubKeys(splitPath(0))
        For i = 0 To splitPath.Length - 1

        Next
        For Each item In baseList.Items
            tmp.add($"{splitPath(0)}{item}{splitPath(1)}")
        Next
        ' Process remaining wildcards


        Return out
    End Function

    Public Function getSubKeys(path As String) As strList
        Dim out As New strList
        Dim root = getFirstDir(path)
        Dim trimmed = path.Replace(root & "\", "")
        Select Case root
            Case "HKCR", "HKEY_CLASSES_ROOT"
                out.add(Registry.ClassesRoot.OpenSubKey(trimmed).GetSubKeyNames)
            Case "HKU", "HKEY_USERS"
                out.add(Registry.Users.OpenSubKey(trimmed).GetSubKeyNames)
            Case "HKCU", "HKEY_CURRENT_USER"
                out.add(Registry.CurrentUser.OpenSubKey(trimmed).GetSubKeyNames)
            Case "HKLM""HKEY_LOCAL_MACHINE"
                out.add(Registry.LocalMachine.OpenSubKey(trimmed).GetSubKeyNames)
            Case "HKCC", "HKEY_CURRENT_CONFIG"
                out.add(Registry.CurrentConfig.OpenSubKey(trimmed).GetSubKeyNames)
        End Select
        Return out
    End Function

    ' <summary>Interpret parameterized wildcards for the current system</summary>
    ' <param name="dir">A path containing a wildcard</param>
    'Private Function expandWildcard(dir As String) As Boolean
    '    ' This should handle wildcards anywhere in a path even though CCleaner only supports them at the end for DetectFiles
    '    Dim possibleDirs As New strList
    '    Dim currentPaths As New strList
    '    currentPaths.add("")
    '    ' Split the given string into sections by directory 
    '    Dim splitDir As String() = dir.Split(CChar("\"))
    '    For Each pathPart In splitDir
    '        ' If this directory parameterization includes a wildcard, expand it appropriately
    '        ' This probably wont work if a string for some reason starts with a *
    '        If pathPart.Contains("*") Then
    '            For Each currentPath In currentPaths.items
    '                Try
    '                    ' Query the existence of child paths for each current path we hold 
    '                    Dim possibilities = getSubKeys(currentPath) 'Directory.GetDirectories(currentPath, pathPart)
    '                    ' If there are any, add them to our possibility list
    '                    Dim pattern As New Regex(pathPart)
    '                    Dim newPossibleDirs As new strList
    '                    For Each subkey In possibilities.items
    '                        If pattern.IsMatch(subkey) Then possibleDirs.items.ForEach(Sub(item) newPossibleDirs.add(item & subkey))
    '                    Next
    '                Catch
    '                    ' The exception we encounter here is going to be the result of directories not existing.
    '                    ' The exception will be thrown from the getSubKeys call and will prevent us from attempting to add new
    '                    ' items to the possibility list. In this instance, we can silently fail (here). 
    '                End Try
    '            Next
    '            ' If no possibilities remain, the wildcard parameterization hasn't left us with any real paths on the system, so we may return false.
    '            If possibleDirs.count = 0 Then Return False
    '            ' Otherwise, clear the current paths and repopulate them with the possible paths 
    '            currentPaths.Clear()
    '            currentPaths.AddRange(possibleDirs)
    '            possibleDirs.Clear()
    '        Else
    '            If currentPaths.count = 0 Then
    '                currentPaths.add($"{pathPart}")
    '            Else
    '                Dim newCurPaths As New List(Of String)
    '                For Each path As String In currentPaths
    '                    If Not path.EndsWith("\") And path <> "" Then path += "\"
    '                    Dim newPath As String = $"{path}{pathPart}\"
    '                    Dim exists As Boolean = Directory.Exists(newPath)
    '                    If Directory.Exists($"{path}{pathPart}\") Then newCurPaths.Add($"{path}{pathPart}\")
    '                Next
    '                currentPaths = newCurPaths
    '            End If
    '        End If
    '    Next
    '    ' If any file/path exists, return true
    '    For Each currDir In currentPaths
    '        If Directory.Exists(currDir) Or File.Exists(currDir) Then Return True
    '    Next
    '    Return False
    'End Function




End Module