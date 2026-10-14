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
Imports System.Threading

''' <summary>
''' Grows the console window, both in conhost and in terminals that host winapp2ool through a
''' pseudoconsole, such as Windows Terminal. Nothing here shrinks the window, so space the user
''' gave it is never taken away.
''' </summary>
'''
''' <remarks>
''' Under a pseudoconsole the <c> Console </c> size setters succeed without resizing anything the
''' user can see: they resize ConPTY's copy of the screen, which then disagrees with the
''' terminal's window and breaks drawing. There we send the xterm resize escape
''' (<c> ESC [ 8 ; rows ; cols t </c>), which Windows Terminal honors, and wait for the terminal
''' to apply it.
''' </remarks>
Friend Module ConsoleWindow

    ''' <summary>
    ''' How long, in milliseconds, we wait for a terminal to apply a resize request before
    ''' treating it as refused
    ''' </summary>
    Private Const ResizeWaitMs As Integer = 500

    ''' <summary>
    ''' The window's size, as <c> "{columns}x{rows}" </c>, when a pseudoconsole terminal last
    ''' left a resize request unhonored, or <c> Nothing </c>. A terminal may refuse every resize
    ''' (a maximized window, split panes, a terminal without the escape), and each refusal costs
    ''' a <see cref="ResizeWaitMs"/> wait before the menu is drawn. So while the window keeps that
    ''' size, we don't ask again for anything at least as large as the refused request.
    ''' </summary>
    Private _refusedFrom As String = Nothing

    ''' <summary>The columns asked for by the request <see cref="_refusedFrom"/> records</summary>
    Private _refusedWidth As Integer

    ''' <summary>The rows asked for by the request <see cref="_refusedFrom"/> records</summary>
    Private _refusedHeight As Integer

    ''' <summary>
    ''' Grows the console window to at least <paramref name="minWidth"/> columns and
    ''' <paramref name="minHeight"/> rows. A dimension already that large is left alone. Does
    ''' nothing in silent mode or when output is redirected.
    ''' </summary>
    '''
    ''' <param name="minWidth">
    ''' The fewest columns the window should have
    ''' </param>
    '''
    ''' <param name="minHeight">
    ''' The fewest rows the window should have
    ''' </param>
    Public Sub EnsureMinimumSize(minWidth As Integer, minHeight As Integer)

        If SuppressOutput OrElse Console.IsOutputRedirected Then Return

        Try

            Dim width = Console.WindowWidth
            Dim height = Console.WindowHeight
            Dim targetWidth = Math.Max(width, minWidth)
            Dim targetHeight = Math.Max(height, minHeight)

            If targetWidth = width AndAlso targetHeight = height Then Return

            If TerminalCapabilities.IsPseudoConsole Then
                requestTerminalResize(targetWidth, targetHeight, width, height)
            Else
                ' A window stretched past the reported largest size keeps its width or height
                Console.SetWindowSize(Math.Max(width, Math.Min(targetWidth, Console.LargestWindowWidth)),
                                      Math.Max(height, Math.Min(targetHeight, Console.LargestWindowHeight)))
            End If

        Catch ex As IO.IOException
            ' The host reports a window but won't let us measure or size it
        Catch ex As ArgumentOutOfRangeException
            ' The host rejected the size; the window keeps the one it had
        Catch ex As PlatformNotSupportedException
        End Try

    End Sub

    ''' <summary>
    ''' Asks a pseudoconsole terminal to resize its window, then waits up to
    ''' <see cref="ResizeWaitMs"/> for the new size to reach us
    ''' </summary>
    '''
    ''' <param name="targetWidth">The number of columns to ask for</param>
    '''
    ''' <param name="targetHeight">The number of rows to ask for</param>
    '''
    ''' <param name="width">The window's current number of columns</param>
    '''
    ''' <param name="height">The window's current number of rows</param>
    Private Sub requestTerminalResize(targetWidth As Integer, targetHeight As Integer, width As Integer, height As Integer)

        If Not TerminalCapabilities.HasVT Then Return

        If $"{width}x{height}" = _refusedFrom AndAlso
           targetWidth >= _refusedWidth AndAlso targetHeight >= _refusedHeight Then Return

        Console.Out.Write($"{ChrW(&H1B)}[8;{targetHeight};{targetWidth}t")

        ' Anything written before the terminal applies the resize is laid out for the old size
        Dim waited = Stopwatch.StartNew()

        Do

            Thread.Sleep(10)
            width = Console.WindowWidth
            height = Console.WindowHeight

            If width >= targetWidth AndAlso height >= targetHeight Then Return

        Loop While waited.ElapsedMilliseconds < ResizeWaitMs

        _refusedFrom = $"{width}x{height}"
        _refusedWidth = targetWidth
        _refusedHeight = targetHeight

    End Sub

    ''' <summary>
    ''' Enlarges conhost's scrollback buffer to 32000 rows so long output isn't lost. Does nothing
    ''' under a pseudoconsole, where the terminal keeps its own scrollback and enlarging ConPTY's
    ''' buffer only makes it disagree with the terminal, or in silent mode, or when output is
    ''' redirected.
    ''' </summary>
    Public Sub ExpandScrollback()

        If SuppressOutput OrElse Console.IsOutputRedirected OrElse TerminalCapabilities.IsPseudoConsole Then Return

        Try

            Console.SetBufferSize(Console.BufferWidth, Math.Max(Console.BufferHeight, 32000))

        Catch ex As IO.IOException
        Catch ex As ArgumentOutOfRangeException
            ' The window is already taller than the requested buffer
        Catch ex As PlatformNotSupportedException
        End Try

    End Sub

End Module
