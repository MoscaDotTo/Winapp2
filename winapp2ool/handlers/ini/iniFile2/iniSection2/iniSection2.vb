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
''' <summary>An object representing a [SectionName] header and its child keys</summary>
Public Class iniSection2

    ''' <summary>The name of the section, without brackets</summary>
    Public ReadOnly Property Name As String

    ''' <summary>The line number from which this section's header was originally read</summary>
    Public ReadOnly Property StartingLineNumber As Integer

    ''' <summary>The keys contained in this section</summary>
    Public ReadOnly Property Keys As New iniKeyCollection

    ''' <summary>Returns the section name as it appears on disk, with brackets</summary>
    Public Function GetFullName() As String
        Return $"[{Name}]"
    End Function

    ''' <summary>Returns whether this section contains a key with the given name</summary>
    ''' <param name="name">The key name to search for (case-insensitive)</param>
    Public Function HasKey(name As String) As Boolean
        Return Keys.Contains(name)
    End Function

    ''' <summary>Returns the key with the given name, or <c>Nothing</c> if not found</summary>
    ''' <param name="name">The key name to look up (case-insensitive)</param>
    Public Function GetKey(name As String) As iniKey2
        Return Keys.GetKey(name)
    End Function

    ''' <summary>Adds a key to this section</summary>
    ''' <param name="key">The key to add</param>
    Public Sub AddKey(key As iniKey2)
        If key Is Nothing Then argIsNull(NameOf(key)) : Return
        Keys.Add(key)
    End Sub

    ''' <summary>Creates a new section with the given name</summary>
    ''' <param name="name">The section name (without brackets)</param>
    ''' <param name="startingLineNumber">The line number of the section header</param>
    Public Sub New(name As String, Optional startingLineNumber As Integer = 0)
        If name Is Nothing Then argIsNull(NameOf(name)) : Return
        Me.Name = name
        Me.StartingLineNumber = startingLineNumber
    End Sub

    ''' <summary>Returns the section as it would appear on disk</summary>
    Public Overrides Function ToString() As String
        Dim out = GetFullName()
        For Each key In Keys
            out += Environment.NewLine & key.ToString()
        Next
        out += Environment.NewLine
        Return out
    End Function

End Class
