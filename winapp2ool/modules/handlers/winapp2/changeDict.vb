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
''' Tracks the changes made to a list of Strings and allows for reverting the changes 
''' </summary>
Public Class changeDict
    Private changes As Dictionary(Of String, String)

    ''' <summary>Creates a new string change tracking dictionary</summary>
    Public Sub New()
        Changes1 = New Dictionary(Of String, String)
    End Sub

    ''' <summary>The dictionary of changes made to a list of strings</summary>
    Public Property Changes1 As Dictionary(Of String, String)
        Get
            Return changes
        End Get
        Set(value As Dictionary(Of String, String))
            changes = value
        End Set
    End Property

    ''' <summary>Tracks renames made while mutating data for string sorting.</summary>
    ''' <param name="currentValue">The current value of a string</param>
    ''' <param name="newValue">The new value of a piece of a string</param>
    Public Sub trackChanges(currentValue As String, newValue As String)
        Try
            If changes.Keys.Contains(currentValue) Then
                changes.Add(newValue, changes.Item(currentValue))
                changes.Remove(currentValue)
            Else
                changes.Add(newValue, currentValue)
            End If
        Catch ex As Exception
            ' If this exception is thrown, there is a duplicate value in the list being audited 
            ' But we are not linting duplicates (or else this value would have been removed by this point)
            ' For now, silently fail here, unless some issue crops up as a result of this silent failure. 
            ' The only other potential failure case I can think of is, if during the sorting process, some entry name or value
            ' wound up becoming the same due to character replacement (eg. someApp 07 and someApp 7 both becoming someApp 007)
            ' This isn't the case in winapp2.ini and is probably easier rectified by just "Not Doing That™" 
        End Try
    End Sub

    ''' <summary>Restores the original state of data mutated for the purposes of string sorting.</summary>
    ''' <param name="lstArray">An array containing lists of strings whose data has been modified</param>
    Public Sub undoChanges(ByRef lstArray As strList())
        For Each lst In lstArray
            For Each key In changes.Keys
                lst.replaceStrAtIndexOf(key, changes.Item(key))
            Next
        Next
    End Sub
End Class