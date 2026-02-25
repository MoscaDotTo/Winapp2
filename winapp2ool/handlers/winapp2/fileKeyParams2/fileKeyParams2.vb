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

''' <summary>The optional deletion behavior flag on a FileKey</summary>
Public Enum fileKeyFlag
    ''' <summary>No flag — delete only matching files, non-recursively</summary>
    None = 0
    ''' <summary>RECURSE — delete matching files in subdirectories too</summary>
    Recurse = 1
    ''' <summary>REMOVESELF — delete matching files and the containing folder; implies RECURSE</summary>
    RemoveSelf = 2
End Enum

''' <summary>
''' Parses and represents the structured components of a FileKey value:
''' the path, the semicolon-delimited file patterns, and the optional deletion flag.
''' <br/><br/>
''' FileKey format: <c>path|pattern[;pattern...][|FLAG]</c>
''' </summary>
Public Class fileKeyParams2

    ''' <summary>The filesystem path — everything before the first pipe</summary>
    Public ReadOnly Property Path As String

    Private ReadOnly _patterns As New List(Of String)

    ''' <summary>The file patterns — everything between the first and optional second pipe, split on semicolons</summary>
    Public ReadOnly Property Patterns As IReadOnlyList(Of String)
        Get
            Return _patterns
        End Get
    End Property

    ''' <summary>The deletion behavior flag — <c>None</c> if absent or unrecognized</summary>
    Public ReadOnly Property Flag As fileKeyFlag

    ''' <summary>
    ''' Parses a raw FileKey value string into its structured components
    ''' </summary>
    ''' <param name="value">The raw value from a FileKey, e.g. <c>%LocalAppData%\App|*.tmp;*.log|RECURSE</c></param>
    Public Sub New(value As String)

        If value Is Nothing Then argIsNull(NameOf(value)) : Return

        Dim pipe1 = value.IndexOf("|"c)

        If pipe1 < 0 Then
            Path = value
            Return
        End If

        Path = value.Substring(0, pipe1)
        Dim afterPipe1 = value.Substring(pipe1 + 1)
        Dim pipe2 = afterPipe1.IndexOf("|"c)

        If pipe2 < 0 Then
            _patterns.AddRange(afterPipe1.Split(CChar(";")))
            Return
        End If

        _patterns.AddRange(afterPipe1.Substring(0, pipe2).Split(CChar(";")))

        Select Case afterPipe1.Substring(pipe2 + 1).ToUpperInvariant()
            Case "RECURSE"
                Flag = fileKeyFlag.Recurse
            Case "REMOVESELF"
                Flag = fileKeyFlag.RemoveSelf
            Case Else
                Flag = fileKeyFlag.None
        End Select

    End Sub

    ''' <summary>
    ''' Reconstructs the FileKey value string from the parsed components.
    ''' Produces <c>path|pat1;pat2[|FLAG]</c>.
    ''' </summary>
    Public Function Reconstruct() As String

        Dim out = Path

        If _patterns.Count > 0 Then
            out &= "|" & String.Join(";", _patterns)
        End If

        If Flag <> fileKeyFlag.None Then
            out &= "|" & Flag.ToString().ToUpperInvariant()
        End If

        Return out

    End Function

End Class
