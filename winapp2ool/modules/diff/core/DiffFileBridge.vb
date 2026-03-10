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
''' Converts an <c>iniFile</c> to an <c>iniFile2</c> for use by the Diff2 pipeline.
''' </summary>
Module DiffFileBridge

    ''' <summary>
    ''' Converts an <c>iniFile</c> to an <c>iniFile2</c>
    ''' </summary>
    '''
    ''' <param name="f">
    ''' The <c>iniFile</c> to convert
    ''' </param>
    '''
    ''' <returns>
    ''' An <c> iniFile2 </c> containing equivalent sections,
    ''' keys, and comments as <c> <paramref name="f"/> </c>
    ''' </returns>
    Public Function ToIniFile2(f As iniFile) As iniFile2

        Dim out = iniFile2.Empty(f.Dir, f.Name)

        For Each comment In f.Comments.Values : out.Comments.Add(New iniComment2(comment.Comment, comment.LineNumber)) : Next

        For Each section In f.Sections.Values

            Dim s2 As New iniSection2(section.Name)

            For Each key In section.Keys.Keys : s2.AddKey(New iniKey2(key.toString())) : Next

            out.AddSection(s2)

        Next

        Return out

    End Function

End Module
