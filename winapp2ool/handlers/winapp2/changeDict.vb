'    Copyright (C) 2018-2022 Hazel Ward
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
''' Tracks the changes made to a list of Strings and allows for reverting the changes 
''' </summary>
Public Class changeDict
    ''' <summary>Creates a new string change tracking dictionary</summary>
    Public Sub New()
        Changes = New Dictionary(Of String, String)
    End Sub

    ''' <summary>The dictionary of changes made to a list of strings</summary>
    Public Property Changes As Dictionary(Of String, String)

    ''' <summary>Tracks renames made while mutating data for string sorting.</summary>
    ''' <param name="currentValue">The current value of a string</param>
    ''' <param name="newValue">The new value of a piece of a string</param>
    Public Sub trackChanges(currentValue As String, newValue As String)
        Try
            If Changes.Keys.Contains(currentValue) Then
                Changes.Add(newValue, Changes.Item(currentValue))
                Changes.Remove(currentValue)
            Else
                Changes.Add(newValue, currentValue)
            End If
        Catch ex As ArgumentException
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
        If lstArray Is Nothing Then argIsNull(NameOf(lstArray)) : Return
        For Each lst In lstArray
            For Each key In Changes.Keys
                lst.replaceStrAtIndexOf(key, Changes.Item(key))
            Next
        Next
    End Sub
End Class