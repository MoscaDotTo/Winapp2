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
Imports System.Text

''' <summary> 
''' Maintains the global log for winapp2ool, which is used to track internal operations and errors.
''' Provides methods for adding to the log, saving it to disk, and printing it to the console.
''' </summary>
Public Module logger

    ''' <summary>
    ''' Synchronizes writes to <c> GlobalLog </c>. Per-thread paths
    ''' (capture buffers) bypass this lock since they touch only thread-local state.
    ''' </summary>
    Private ReadOnly _logLock As New Object

    ''' <summary>
    ''' The global Winapp2ool log, containing everything logged during the current session
    ''' </summary>
    Public Property GlobalLog As New strList

    ''' <summary>
    ''' Per-thread indentation depth. Each thread sees its own counter so that
    ''' parallel work cannot corrupt the indentation of unrelated threads.
    ''' </summary>
    '''
    ''' <remarks>
    ''' <c> ThreadStatic </c> fields default to zero on every thread that has not
    ''' yet written to them, exactly the desired initial state. Threads spawned
    ''' inside a parallel section start at depth zero regardless of the calling
    ''' thread's depth; use <c> gLogCapture </c> to record those threads' output
    ''' for later flush under the parent's depth.
    ''' </remarks>
    <ThreadStatic>
    Private _nestCount As Integer

    ''' <summary>
    ''' When non-<c> Nothing </c>, <c> gLog </c> writes go to this thread-local
    ''' buffer instead of <c> GlobalLog </c>. Set by <c> gLogCapture </c> and
    ''' restored on dispose.
    ''' </summary>
    <ThreadStatic>
    Private _captureBuffer As List(Of String)

    ''' <summary>
    ''' The current indentation level of the global log on the calling thread.
    ''' </summary>
    Public Property nestCount As Integer
        Get
            Return _nestCount
        End Get
        Set(value As Integer)
            _nestCount = value
        End Set
    End Property

    ''' <summary> 
    ''' Adds an item into the global log 
    ''' </summary>
    ''' 
    ''' <param name="logstr"> 
    ''' The <c> String </c> to be added into the log <br /> 
    ''' Optional, Default: <c> ""</c> 
    ''' </param>
    ''' 
    ''' <param name="cond"> 
    ''' Indicates that the <c> <paramref name="logstr"/> </c> should be added into the log <br /> 
    ''' Optional, Default: <c> True </c> 
    ''' </param>
    ''' 
    ''' <param name="buffr"> 
    ''' Indicates that an empty line should be added into the log following <c> <paramref name="logstr"/> </c> 
    ''' </param>
    ''' 
    ''' <param name="leadr"> 
    ''' Indicates that an empty line should be added into the log before <c> <paramref name="logstr"/> </c> 
    ''' </param>
    Public Sub gLog(Optional logstr As String = "",
                    Optional cond As Boolean = True,
                    Optional buffr As Boolean = False,
                    Optional leadr As Boolean = False)

        If Not cond Then Return

        Dim capture = _captureBuffer

        If capture IsNot Nothing Then

            If leadr Then capture.Add("")

            If logstr IsNot Nothing Then

                Dim pad = New String(" "c, Math.Max(_nestCount * 2, 0))
                capture.Add(pad & logstr)

            End If

            If buffr Then capture.Add("")

            Return

        End If

        SyncLock _logLock

            If leadr Then GlobalLog.add("")

            If logstr IsNot Nothing Then

                Dim pad = New String(" "c, Math.Max(_nestCount * 2, 0))
                GlobalLog.add(pad & logstr)

            End If

            If buffr Then GlobalLog.add("")

        End SyncLock

    End Sub

    ''' <summary>
    ''' Opens a nested logging scope. The optional <paramref name="message"/> is logged at the
    ''' current depth, then the depth is increased by <paramref name="amount"/> until the
    ''' returned <c> IDisposable </c> is disposed. Use with <c> Using </c> to make
    ''' ascend/descend imbalance impossible by construction.
    ''' </summary>
    '''
    ''' <param name="message">
    ''' Optional header line written before ascending. Pass <c> "" </c> for an anonymous scope
    ''' </param>
    '''
    ''' <param name="amount">
    ''' Number of indentation levels to ascend. Optional, default: <c> 1 </c>
    ''' </param>
    '''
    ''' <example>
    ''' <code>
    ''' Using gLogScope("Beginning lint")
    '''     ' nested gLog calls are indented one level deeper here
    ''' End Using
    ''' </code>
    ''' </example>
    Public Function gLogScope(Optional message As String = "",
                              Optional amount As Integer = 1) As IDisposable

        If message IsNot Nothing AndAlso message.Length > 0 Then gLog(message)

        Return New LogScope(amount)

    End Function

    ''' <summary>
    ''' Redirects the calling thread's <c> gLog </c> writes into a thread-local buffer
    ''' until the returned object is disposed. Captured lines preserve their relative
    ''' indentation (depth resets to zero on entry and restores on exit), so they can
    ''' be replayed later under any parent depth via <c> EmitCaptured </c>.
    ''' </summary>
    '''
    ''' <returns>
    ''' A <c> LogCapture </c> whose <c> Lines </c> contains the buffered output. The capture
    ''' is active until <c> Dispose </c> is called (typically via <c> Using </c>)
    ''' </returns>
    '''
    ''' <remarks>
    ''' Designed for parallel sections: each parallel task captures its own log slice,
    ''' and the orchestrating thread emits the slices in deterministic order after the
    ''' parallel work finishes. This eliminates the need for per-module deferred-log
    ''' workarounds (see <c> EntryLintResult.DiagLines </c>).
    ''' </remarks>
    Public Function gLogCapture() As LogCapture

        Return New LogCapture()

    End Function

    ''' <summary>
    ''' Appends previously-captured lines to <c> GlobalLog </c>, indented at the calling
    ''' thread's current depth. Acquires the log lock once for the whole batch.
    ''' </summary>
    '''
    ''' <param name="lines">
    ''' The captured lines, as returned by <c> LogCapture.Lines </c>. <c> Nothing </c> is a no-op
    ''' </param>
    Public Sub EmitCaptured(lines As IEnumerable(Of String))

        If lines Is Nothing Then Return

        Dim pad = New String(" "c, Math.Max(_nestCount * 2, 0))
        Dim target = _captureBuffer

        If target IsNot Nothing Then

            For Each line In lines
                target.Add(pad & line)
            Next

            Return

        End If

        SyncLock _logLock

            For Each line In lines
                GlobalLog.add(pad & line)
            Next

        End SyncLock

    End Sub

    ''' <summary>
    ''' Disposable handle returned by <c> gLogScope </c>. On dispose, restores the
    ''' nesting depth that was in effect at construction.
    ''' </summary>
    '''
    ''' <remarks>
    ''' Created and disposed on the same thread. Passing the handle to another thread
    ''' for disposal would corrupt that thread's depth — don't do it.
    ''' </remarks>
    Public NotInheritable Class LogScope

        Implements IDisposable

        Private ReadOnly _amount As Integer
        Private _disposed As Boolean

        Friend Sub New(amount As Integer)

            _amount = amount
            _nestCount += amount

        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose

            If _disposed Then Return
            _disposed = True
            _nestCount -= _amount

        End Sub

    End Class

    ''' <summary>
    ''' Disposable handle returned by <c> gLogCapture </c>. While alive, redirects the
    ''' calling thread's <c> gLog </c> writes into <c> Lines </c>. Restores the previous
    ''' capture buffer and nesting depth on dispose.
    ''' </summary>
    '''
    ''' <remarks>
    ''' Captures nest correctly: an inner capture saves and restores the outer capture's
    ''' buffer, so its lines flow back to the outer buffer rather than to <c> GlobalLog </c>.
    ''' </remarks>
    Public NotInheritable Class LogCapture

        Implements IDisposable

        Private ReadOnly _previousBuffer As List(Of String)
        Private ReadOnly _previousNest As Integer
        Private ReadOnly _buffer As New List(Of String)
        Private _disposed As Boolean

        Friend Sub New()

            _previousBuffer = _captureBuffer
            _previousNest = _nestCount
            _captureBuffer = _buffer
            _nestCount = 0

        End Sub

        ''' <summary>
        ''' The captured lines, with relative indentation preserved.
        ''' </summary>
        Public ReadOnly Property Lines As IReadOnlyList(Of String)
            Get
                Return _buffer
            End Get
        End Property

        Public Sub Dispose() Implements IDisposable.Dispose

            If _disposed Then Return
            _disposed = True
            _captureBuffer = _previousBuffer
            _nestCount = _previousNest

        End Sub

    End Class

    ''' <summary> 
    ''' Saves the global log to disk if the given <c> <paramref name="cond"/> </c> is met 
    ''' </summary>
    ''' 
    ''' <param name="cond">
    ''' Indicates that the global log should be saved to disk 
    ''' <br /> Optional, Default: <c> False </c>
    ''' </param>
    Public Sub saveGlobalLog(Optional cond As Boolean = True)

        If cond Then IO.File.WriteAllText(GlobalLogFile.Path(), logger.toString)

    End Sub

    ''' <summary>
    ''' Returns the log as a single <c> String </c>
    ''' </summary>
    Public Function toString() As String

        Dim sb As New StringBuilder()

        GlobalLog.Items.ForEach(Sub(line) sb.AppendLine(line))

        Return sb.ToString()

    End Function

    ''' <summary> 
    ''' Prints the winapp2ool log to the user and waits for the 'enter' key to be pressed before returning to the calling menu 
    ''' </summary>
    Public Sub printLog()

        cwl("Printing the winapp2ool log, this may take a moment")

        Dim out = logger.toString

        clrConsole()

        cwl(out)
        cwl()
        cwl($"End of log. {pressEnterStr}")

        Console.ReadLine()

    End Sub


    '''<summary> 
    ''' Gets the most recent segment of the global log contained by two phrases (ie. a module name or subroutine) 
    ''' As a <c> String </c> <br /> Can be used to simply fetch the logs from modules for saving to disk or displaying to the user after a run.
    ''' </summary>
    ''' 
    ''' <param name="startingPhrase"> 
    ''' The starting phrase of the requested log slice 
    ''' </param>
    ''' 
    ''' <param name="endingPhrase"> 
    ''' The ending phrase of the requested log slice 
    ''' </param>
    ''' 
    ''' <returns> 
    ''' The set of log lines between the most recent incidences of the provided phrases from the log 
    ''' </returns>
    Public Function getLogSliceFromGlobal(startingPhrase As String, endingPhrase As String) As String

        Dim startInd = -1
        Dim endInd = -1

        For i = GlobalLog.Items.Count - 1 To 0 Step -1


            If GlobalLog.Items(i) Is Nothing OrElse Not GlobalLog.Items(i).Contains(startingPhrase) Then Continue For

            startInd = i
            Exit For

        Next

        If startInd = -1 Then Return ""

        For i = startInd To GlobalLog.Items.Count - 1

            If Not GlobalLog.Items(i).EndsWith(endingPhrase, StringComparison.InvariantCultureIgnoreCase) Then Continue For

            endInd = i
            Exit For

        Next

        If endInd = -1 OrElse endInd < startInd Then Return ""

        Dim toTrim = 0

        For Each c In GlobalLog.Items(startInd)

            If Not c = CChar(" ") Then Exit For
            toTrim += 1

        Next

        Dim sb As New StringBuilder()
        For i = startInd To endInd

            Dim line = GlobalLog.Items(i)


            Dim leadingSpaces = 0
            For Each ch In line

                If ch <> " "c Then Exit For
                leadingSpaces += 1
                If leadingSpaces >= toTrim Then Exit For

            Next

            sb.AppendLine(line.Substring(leadingSpaces))

        Next

        Return sb.ToString()

    End Function

    ''' <summary> 
    ''' Prints a slice of the global log to the user and waits for them to press the 'enter' key 
    ''' </summary>
    '''
    ''' <param name="slice">
    ''' A portion of the global log to be printed to the user 
    ''' </param>
    Public Sub printSlice(slice As String)

        clrConsole()
        cwl(slice)
        Console.ReadLine()

    End Sub

End Module