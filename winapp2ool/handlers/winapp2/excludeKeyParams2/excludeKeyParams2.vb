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

''' <summary>The exclusion type flag on an ExcludeKey</summary>
Public Enum excludeKeyFlag
    ''' <summary>FILE — exclude a specific file by exact path</summary>
    File = 0
    ''' <summary>PATH — exclude files matching a wildcard pattern under a path</summary>
    Path = 1
    ''' <summary>REG — exclude a registry path or value</summary>
    Reg = 2
    ''' <summary>Unknown — flag value not recognized; stored verbatim for round-trip fidelity</summary>
    Unknown = 3
End Enum

''' <summary>
''' Parses and represents the structured components of an ExcludeKey value:
''' the exclusion type flag, the path, and the optional semicolon-delimited patterns.
''' <br/><br/>
''' ExcludeKey format: <c>FLAG|path[|pattern[;pattern...]]</c>
''' </summary>
Public Class excludeKeyParams2

    ''' <summary>The exclusion type: FILE, PATH, or REG</summary>
    Public ReadOnly Property Flag As excludeKeyFlag

    ''' <summary>
    ''' The raw flag text as it appeared in the file.
    ''' Populated only when <c>Flag = excludeKeyFlag.Unknown</c> to preserve the original text for reconstruction.
    ''' </summary>
    Public ReadOnly Property RawFlag As String

    ''' <summary>The path — everything between the first and optional second pipe</summary>
    Public ReadOnly Property Path As String

    Private ReadOnly _patterns As New List(Of String)

    ''' <summary>
    ''' The optional patterns — everything after the second pipe, split on semicolons.
    ''' For FILE exclusions this is a filename; for PATH it is a wildcard; for REG it is a value name.
    ''' Empty when no second pipe was present.
    ''' </summary>
    Public ReadOnly Property Patterns As IReadOnlyList(Of String)
        Get
            Return _patterns
        End Get
    End Property

    ''' <summary>Returns <c>True</c> when patterns are present (second pipe existed)</summary>
    Public ReadOnly Property HasPatterns As Boolean
        Get
            Return _patterns.Count > 0
        End Get
    End Property

    ''' <summary>
    ''' Parses a raw ExcludeKey value string into its structured components
    ''' </summary>
    ''' <param name="value">The raw value from an ExcludeKey, e.g. <c>FILE|%LocalAppData%\App|important.dat</c></param>
    Public Sub New(value As String)

        If value Is Nothing Then argIsNull(NameOf(value)) : Return

        Dim pipe1 = value.IndexOf("|"c)

        If pipe1 < 0 Then
            Flag = excludeKeyFlag.Unknown
            RawFlag = ""
            Path = value
            Return
        End If

        Dim flagStr = value.Substring(0, pipe1)
        Dim afterPipe1 = value.Substring(pipe1 + 1)

        Select Case flagStr.ToUpperInvariant()
            Case "FILE"
                Flag = excludeKeyFlag.File
                RawFlag = ""
            Case "PATH"
                Flag = excludeKeyFlag.Path
                RawFlag = ""
            Case "REG"
                Flag = excludeKeyFlag.Reg
                RawFlag = ""
            Case Else
                Flag = excludeKeyFlag.Unknown
                RawFlag = flagStr
        End Select

        Dim pipe2 = afterPipe1.IndexOf("|"c)

        If pipe2 < 0 Then
            Path = afterPipe1
        Else
            Path = afterPipe1.Substring(0, pipe2)
            _patterns.AddRange(afterPipe1.Substring(pipe2 + 1).Split(CChar(";")))
        End If

    End Sub

    ''' <summary>
    ''' Reconstructs the ExcludeKey value string from the parsed components.
    ''' Produces <c>FLAG|path[|pat1;pat2]</c>.
    ''' </summary>
    Public Function Reconstruct() As String

        Dim flagText = If(Flag = excludeKeyFlag.Unknown, RawFlag, Flag.ToString().ToUpperInvariant())
        Dim out = $"{flagText}|{Path}"

        If _patterns.Count > 0 Then
            out &= "|" & String.Join(";", _patterns)
        End If

        Return out

    End Function

End Class
