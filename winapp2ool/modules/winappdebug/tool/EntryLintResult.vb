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
''' A single error found while linting a <c>winapp2entry2</c>
''' </summary>
Public Structure LintError

    ''' <summary>
    ''' The error description shown to the user
    ''' </summary>
    Public ReadOnly Property Message As String

    ''' <summary>
    ''' Supporting detail lines shown below the error description
    ''' </summary>
    Public ReadOnly Property Details As IReadOnlyList(Of String)

    '''
    Public Sub New(message As String, details As String())
        Me.Message = message
        Me.Details = details
    End Sub

End Structure

''' <summary>
''' Accumulates errors found while linting a single <c>winapp2entry2</c>. <br/>
''' Each entry processed by <c>WinappDebug</c> gets its own <c>EntryLintResult</c>,
''' making per-entry processing independent of shared state.
''' </summary>
Public Class EntryLintResult

    ''' <summary>
    ''' The full name of the entry being linted, used as error context in output
    ''' </summary>
    Public ReadOnly Property EntryName As String

    Private ReadOnly _errors As New List(Of LintError)

    ''' <summary>
    ''' All errors recorded for this entry
    ''' </summary>
    Public ReadOnly Property Errors As IReadOnlyList(Of LintError)
        Get
            Return _errors
        End Get
    End Property

    ''' <summary>
    ''' The number of errors recorded for this entry
    ''' </summary>
    Public ReadOnly Property ErrorCount As Integer
        Get
            Return _errors.Count
        End Get
    End Property

    '''
    Public Sub New(entryName As String)
        Me.EntryName = entryName
    End Sub

    ''' <summary>
    ''' Records an error if <paramref name="cond"/> is <c> True </c>
    ''' </summary>
    '''
    ''' <param name="message">
    ''' The error description
    ''' </param>
    '''
    ''' <param name="details">
    ''' Supporting detail lines for the error
    ''' </param>
    '''
    ''' <param name="cond">
    ''' The condition under which the error should be recorded
    ''' <br /> Optional, Default: <c> True </c>
    ''' </param>
    Public Sub RecordError(message As String, details As String(), Optional cond As Boolean = True)
        If Not cond Then Return
        _errors.Add(New LintError(message, details))
    End Sub

End Class
