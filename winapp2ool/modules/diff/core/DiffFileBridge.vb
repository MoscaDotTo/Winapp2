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
''' Converts between the legacy ini layer (<c>iniFile</c>/<c>iniSection</c>/<c>iniKey</c>)
''' and the new ini layer (<c>iniFile2</c>/<c>iniSection2</c>/<c>iniKey2</c>).
''' Used by the diff migration comparison harness.
''' </summary>
Module DiffFileBridge

    ''' <summary>Converts an <c>iniFile</c> to an <c>iniFile2</c></summary>
    Public Function ToIniFile2(f As iniFile) As iniFile2

        Dim out = iniFile2.Empty(f.Dir, f.Name)

        For Each section In f.Sections.Values

            Dim s2 As New iniSection2(section.Name)

            For Each key In section.Keys.Keys
                s2.AddKey(New iniKey2(key.toString()))
            Next

            out.AddSection(s2)

        Next

        Return out

    End Function

    ''' <summary>Converts an <c>iniFile2</c> to an <c>iniFile</c></summary>
    Public Function ToIniFile(f As iniFile2) As iniFile

        Dim out As New iniFile()

        For Each section In f
            out.Sections.Add(section.Name, ToIniSection(section))
        Next

        Return out

    End Function

    ''' <summary>Converts an <c>iniSection2</c> to an <c>iniSection</c></summary>
    Public Function ToIniSection(s As iniSection2) As iniSection

        Dim lines As New List(Of String) From {s.GetFullName()}

        For Each key In s.Keys
            lines.Add(key.ToString())
        Next

        Return New iniSection(lines)

    End Function

    ''' <summary>Converts an <c>iniKey2</c> to an <c>iniKey</c></summary>
    Public Function ToIniKey(k As iniKey2) As iniKey
        Return New iniKey(k.ToString())
    End Function

End Module
