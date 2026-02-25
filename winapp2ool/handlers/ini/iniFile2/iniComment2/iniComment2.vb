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

''' <summary>An object representing a comment line from an ini file</summary>
Public Class iniComment2

    ''' <summary>The raw comment text, including the leading semicolon</summary>
    Public ReadOnly Property Text As String

    ''' <summary>The line number from which this comment was originally read</summary>
    Public ReadOnly Property LineNumber As Integer

    ''' <summary>Creates an <c>iniComment2</c> from a comment line and its line number</summary>
    ''' <param name="text">The raw comment text</param>
    ''' <param name="lineNumber">The line number in the source file</param>
    Public Sub New(text As String, lineNumber As Integer)
        Me.Text = text
        Me.LineNumber = lineNumber
    End Sub

    ''' <summary>Returns the raw comment text</summary>
    Public Overrides Function ToString() As String
        Return Text
    End Function

End Class
