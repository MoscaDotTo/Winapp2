'    Copyright (C) 2018-2020 Robbie Ward
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
''' An object representing a comment in a .ini file
''' </summary>
Public Class iniComment
    ''' <summary> The text of the comment</summary>
    Public Property Comment As String
    ''' <summary> The line number from which the comment was originally read</summary>
    Public Property LineNumber As Integer

    ''' <summary>Creates a new iniComment object</summary>
    ''' <param name="c">The comment text</param>
    ''' <param name="l">The line number</param>
    Public Sub New(c As String, l As Integer)
        Comment = c
        LineNumber = l
    End Sub
End Class