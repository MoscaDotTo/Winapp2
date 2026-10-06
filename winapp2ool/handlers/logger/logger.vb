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
    Public Property GlobalLog As New List(Of String)

    ''' <summary>
    ''' Indicates whether the global log is written to disk as the application exits, even
    ''' on a clean silent-mode run. The global <c> -writelog </c> command line flag sets it.
    ''' A nonzero process exit code saves the log independently of this flag, so scripted and
    ''' CI runs always retain the diagnostics from a failed build.
    ''' </summary>
    Public Property SaveGlobalLogOnExit As Boolean = False

    ''' <summary>
    ''' Per-thread indentation depth. Each thread sees its own counter so that
    ''' parallel work cannot corrupt the indentation of unrelated threads.
    ''' </summary>
    '''
    ''' <remarks>
    ''' <c> ThreadStatic </c> fields default to zero on every thread that has not
    ''' yet written to them, exactly the desired initial state. Threads spawned
    ''' inside a parallel section start at depth zero regardless of the calling
    ''' thread's depth; use <see cref="gLogCapture"/> to record those threads' output
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
    ''' Adds a line to the global log, indented two spaces per level of the calling thread's
    ''' depth. While the thread has a <see cref="gLogCapture"/> open, the line goes to that
    ''' capture's buffer instead.
    ''' </summary>
    '''
    ''' <param name="logstr">
    ''' The <c> String </c> to be added into the log. <c> Nothing </c> adds no line, though
    ''' <paramref name="buffr"/> and <paramref name="leadr"/> still add theirs. <br /><br />
    ''' Optional, Default: <c> "" </c>
    ''' </param>
    '''
    ''' <param name="cond">
    ''' Indicates whether to log anything at all <br /><br />
    ''' Optional, Default: <c> True </c>
    ''' </param>
    '''
    ''' <param name="buffr">
    ''' Indicates whether to add an empty line after <paramref name="logstr"/> <br /><br />
    ''' Optional, Default: <c> False </c>
    ''' </param>
    '''
    ''' <param name="leadr">
    ''' Indicates whether to add an empty line before <paramref name="logstr"/> <br /><br />
    ''' Optional, Default: <c> False </c>
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

            If leadr Then GlobalLog.Add("")

            If logstr IsNot Nothing Then

                Dim pad = New String(" "c, Math.Max(_nestCount * 2, 0))
                GlobalLog.Add(pad & logstr)

            End If

            If buffr Then GlobalLog.Add("")

        End SyncLock

    End Sub

    ''' <summary>
    ''' Adds a line to the global log like the other overload, but builds the line only when
    ''' <paramref name="cond"/> is <c> True </c>, so a suppressed call costs no string
    ''' interpolation.
    ''' </summary>
    '''
    ''' <param name="messageFactory">
    ''' Factory invoked to produce the log line. Not called when <paramref name="cond"/> is
    ''' <c> False </c>, and <c> Nothing </c> logs nothing at all
    ''' </param>
    '''
    ''' <param name="cond">
    ''' Indicates whether to log anything at all
    ''' </param>
    '''
    ''' <param name="buffr">
    ''' Indicates whether to add an empty line after the message <br /><br />
    ''' Optional, Default: <c> False </c>
    ''' </param>
    '''
    ''' <param name="leadr">
    ''' Indicates whether to add an empty line before the message <br /><br />
    ''' Optional, Default: <c> False </c>
    ''' </param>
    Public Sub gLog(messageFactory As Func(Of String),
                    cond As Boolean,
                    Optional buffr As Boolean = False,
                    Optional leadr As Boolean = False)

        If Not cond Then Return
        If messageFactory Is Nothing Then Return

        gLog(messageFactory(), True, buffr, leadr)

    End Sub

    ''' <summary>
    ''' Opens a nested logging scope. The optional <paramref name="message"/> is logged at the
    ''' current depth, then the depth is increased by <paramref name="amount"/> until the
    ''' returned <c> IDisposable </c> is disposed. Open it with <c> Using </c> so the depth
    ''' always comes back down.
    ''' </summary>
    '''
    ''' <param name="message">
    ''' Header line written before ascending. <c> "" </c> or <c> Nothing </c> opens a scope
    ''' with no header. <br /><br />
    ''' Optional, Default: <c> "" </c>
    ''' </param>
    '''
    ''' <param name="amount">
    ''' Number of indentation levels to ascend <br /><br />
    ''' Optional, Default: <c> 1 </c>
    ''' </param>
    '''
    ''' <returns>A <see cref="LogScope"/> that lowers the depth again when disposed</returns>
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
    ''' be replayed later under any parent depth via <see cref="EmitCaptured"/>.
    ''' </summary>
    '''
    ''' <returns>
    ''' A <see cref="LogCapture"/> whose <c> Lines </c> contains the buffered output. The capture
    ''' is active until <c> Dispose </c> is called (typically via <c> Using </c>), and its
    ''' lines reach the log only if someone passes them to <see cref="EmitCaptured"/>
    ''' </returns>
    '''
    ''' <remarks>
    ''' Designed for parallel sections: each parallel task captures its own log slice,
    ''' and the orchestrating thread emits the slices in deterministic order after the
    ''' parallel work finishes.
    ''' </remarks>
    Public Function gLogCapture() As LogCapture

        Return New LogCapture()

    End Function

    ''' <summary>
    ''' Appends previously-captured lines to <see cref="GlobalLog"/>, indented at the calling
    ''' thread's current depth, taking the log lock once for the whole batch. If the calling
    ''' thread has a capture open itself, the lines go to that capture's buffer instead.
    ''' </summary>
    '''
    ''' <param name="lines">
    ''' The captured lines, as returned by <see cref="LogCapture.Lines"/>. <c> Nothing </c> is a no-op
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
                GlobalLog.Add(pad & line)
            Next

        End SyncLock

    End Sub

    ''' <summary>
    ''' Disposable handle returned by <see cref="gLogScope"/>. On dispose, lowers the
    ''' calling thread's depth by the amount the constructor raised it.
    ''' </summary>
    '''
    ''' <remarks>
    ''' Create and dispose it on the same thread. Disposing it on another thread lowers that
    ''' thread's depth instead, so don't pass it across threads.
    ''' </remarks>
    Public NotInheritable Class LogScope

        Implements IDisposable

        Private ReadOnly _amount As Integer
        Private _disposed As Boolean

        ''' <summary>Creates a new <c> LogScope </c>, raising the calling thread's depth by <paramref name="amount"/></summary>
        ''' <param name="amount">The number of indentation levels to ascend</param>
        Friend Sub New(amount As Integer)

            _amount = amount
            _nestCount += amount

        End Sub

        ''' <summary>Lowers the depth by the amount the constructor raised it. Later calls do nothing.</summary>
        Public Sub Dispose() Implements IDisposable.Dispose

            If _disposed Then Return
            _disposed = True
            _nestCount -= _amount

        End Sub

    End Class

    ''' <summary>
    ''' Disposable handle returned by <see cref="gLogCapture"/>. While alive, redirects the
    ''' calling thread's <c> gLog </c> writes into <c> Lines </c>. Restores the previous
    ''' capture buffer and nesting depth on dispose.
    ''' </summary>
    '''
    ''' <remarks>
    ''' Captures nest: an inner capture saves the outer capture's buffer and restores it on
    ''' dispose. The inner capture's lines stay in its own <c> Lines </c>, and passing them to
    ''' <see cref="EmitCaptured"/> while the outer capture is open adds them to the outer buffer
    ''' rather than to <see cref="GlobalLog"/>.
    ''' </remarks>
    Public NotInheritable Class LogCapture

        Implements IDisposable

        Private ReadOnly _previousBuffer As List(Of String)
        Private ReadOnly _previousNest As Integer
        Private ReadOnly _buffer As New List(Of String)
        Private _disposed As Boolean

        ''' <summary>
        ''' Creates a new <c> LogCapture </c>, pointing the calling thread's writes at its
        ''' buffer and resetting the thread's depth to zero
        ''' </summary>
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

        ''' <summary>
        ''' Restores the capture buffer and depth the thread had before this capture opened.
        ''' Later calls do nothing.
        ''' </summary>
        Public Sub Dispose() Implements IDisposable.Dispose

            If _disposed Then Return
            _disposed = True
            _captureBuffer = _previousBuffer
            _nestCount = _previousNest

        End Sub

    End Class

    ''' <summary>
    ''' Writes the global log to <see cref="GlobalLogFile"/>, replacing whatever the file held.
    ''' A write error goes uncaught to the caller.
    ''' </summary>
    '''
    ''' <param name="cond">
    ''' Indicates whether to save the log at all <br /><br />
    ''' Optional, Default: <c> True </c>
    ''' </param>
    Public Sub saveGlobalLog(Optional cond As Boolean = True)

        If cond Then IO.File.WriteAllText(GlobalLogFile.Path(), logger.toString)

    End Sub

    ''' <summary>
    ''' Returns the log as a single <c> String </c>, with a newline after every line
    ''' </summary>
    Public Function toString() As String

        Dim sb As New StringBuilder()

        GlobalLog.ForEach(Sub(line) sb.AppendLine(line))

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

        crl()

    End Sub


    ''' <summary>
    ''' Returns the most recent stretch of the global log that a module run left between two
    ''' phrases, such as its start and end messages. The slice starts at the last line that
    ''' contains <paramref name="startingPhrase"/> and ends at the first line from there on that
    ''' ends with <paramref name="endingPhrase"/>. We strip up to the starting line's indentation
    ''' from each line, but always at least one leading space, so when the starting line isn't
    ''' indented every indented line below it loses one space.
    ''' </summary>
    '''
    ''' <param name="startingPhrase">
    ''' Text the first line must contain (case-sensitive)
    ''' </param>
    '''
    ''' <param name="endingPhrase">
    ''' Text the last line must end with (case-insensitive). The starting line itself can match.
    ''' </param>
    '''
    ''' <returns>
    ''' The slice with a newline after every line, or <c> "" </c> if either phrase isn't found
    ''' </returns>
    Public Function getLogSliceFromGlobal(startingPhrase As String, endingPhrase As String) As String

        Dim startInd = -1
        Dim endInd = -1

        For i = GlobalLog.Count - 1 To 0 Step -1


            If GlobalLog(i) Is Nothing OrElse Not GlobalLog(i).Contains(startingPhrase) Then Continue For

            startInd = i
            Exit For

        Next

        If startInd = -1 Then Return ""

        For i = startInd To GlobalLog.Count - 1

            If Not GlobalLog(i).EndsWith(endingPhrase, StringComparison.InvariantCultureIgnoreCase) Then Continue For

            endInd = i
            Exit For

        Next

        If endInd = -1 OrElse endInd < startInd Then Return ""

        Dim toTrim = 0

        For Each c In GlobalLog(startInd)

            If Not c = CChar(" ") Then Exit For
            toTrim += 1

        Next

        Dim sb As New StringBuilder()
        For i = startInd To endInd

            Dim line = GlobalLog(i)


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
        crl()

    End Sub

End Module