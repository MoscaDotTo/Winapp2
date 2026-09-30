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
''' A single error found while linting a <c> winapp2entry2 </c>
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
''' Accumulates errors found while linting a single <c> winapp2entry2 </c>. <br />
''' Every entry <c> WinappDebug </c> processes gets its own, so no part of the per-entry
''' work has to reach for shared state.
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

    ''' <summary>
    ''' Log lines <c> gLogCapture </c> picked up while this entry was being processed in
    ''' parallel. <c> EmitCaptured </c> flushes them in a deterministic order when the result
    ''' gets rendered.
    ''' </summary>
    '''
    ''' <remarks>
    ''' <c> WinappDebug.ProcessEntry </c> opens a capture scope, runs the entry through the
    ''' lint pipeline, then drops the captured <c> Lines </c> in here.
    ''' <c> WinappDebug.EmitEntryResult </c> flushes them back out at the calling thread's depth.
    ''' </remarks>
    Public Property LogLines As IReadOnlyList(Of String) = New List(Of String)()

    Private ReadOnly _deferredSections As New List(Of MenuSection)

    ''' <summary>
    ''' <c> MenuSection </c>s built during parallel processing to be rendered sequentially
    ''' </summary>
    Public ReadOnly Property DeferredSections As IReadOnlyList(Of MenuSection)
        Get
            Return _deferredSections
        End Get
    End Property

    ''' <summary>
    ''' Queues a <c> MenuSection </c> to be rendered when this result is emitted
    ''' </summary>
    '''
    ''' <param name="section">
    ''' The <c> MenuSection </c> to defer
    ''' </param>
    Public Sub DeferSection(section As MenuSection)
        _deferredSections.Add(section)
    End Sub

End Class
