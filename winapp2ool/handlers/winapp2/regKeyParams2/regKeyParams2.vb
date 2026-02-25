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

''' <summary>
''' Parses and represents the structured components of a RegKey value:
''' the registry path and the optional value name (subkey).
''' <br/><br/>
''' RegKey format: <c>registry_path[|value_name]</c>
''' </summary>
Public Class regKeyParams2

    ''' <summary>The registry path — everything before the pipe, or the whole value if no pipe</summary>
    Public ReadOnly Property Path As String

    ''' <summary>
    ''' The specific registry value name to delete — everything after the pipe.
    ''' Empty string when absent: deleting the whole key rather than a single value.
    ''' </summary>
    Public ReadOnly Property Subkey As String

    ''' <summary>Returns <c>True</c> when a specific value name is targeted rather than the whole key</summary>
    Public ReadOnly Property HasSubkey As Boolean
        Get
            Return Subkey.Length > 0
        End Get
    End Property

    ''' <summary>
    ''' Parses a raw RegKey value string into its structured components
    ''' </summary>
    ''' <param name="value">The raw value from a RegKey, e.g. <c>HKCU\Software\App|SettingName</c></param>
    Public Sub New(value As String)

        If value Is Nothing Then argIsNull(NameOf(value)) : Return

        Dim pipePos = value.IndexOf("|"c)

        If pipePos < 0 Then
            Path = value
            Subkey = ""
        Else
            Path = value.Substring(0, pipePos)
            Subkey = value.Substring(pipePos + 1)
        End If

    End Sub

    ''' <summary>
    ''' Reconstructs the RegKey value string from the parsed components.
    ''' Produces <c>path</c> or <c>path|value_name</c>.
    ''' </summary>
    Public Function Reconstruct() As String
        Return If(HasSubkey, $"{Path}|{Subkey}", Path)
    End Function

End Class
