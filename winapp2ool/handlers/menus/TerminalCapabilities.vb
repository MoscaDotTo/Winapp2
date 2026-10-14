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
Imports System.Runtime.InteropServices

''' <summary>
''' Probes whether the attached console supports virtual terminal (ANSI/VT)
''' escape sequences. On Windows 10 1607 (Aug 2016) and newer, conhost and
''' Windows Terminal both support VT once
''' <c> ENABLE_VIRTUAL_TERMINAL_PROCESSING </c> is set on stdout, and the probe
''' sets it if it isn't already. Older versions (Windows 7, 8.1, Windows 10
''' before 1607) return failure from <c> SetConsoleMode </c> when the flag is requested.
''' A redirected stdout counts as unsupported. It also probes whether the console is a
''' pseudoconsole (<see cref="IsPseudoConsole"/>).
'''
''' Each result is cached after the first probe. Tests may force a value via
''' <see cref="SetHasVTForTesting"/> to exercise either code path.
''' </summary>
Friend Module TerminalCapabilities

    Private Const STD_OUTPUT_HANDLE As Integer = -11
    Private Const ENABLE_VIRTUAL_TERMINAL_PROCESSING As UInteger = 4

    Private _probed As Boolean = False
    Private _hasVT As Boolean = False

    Private _pseudoConsoleProbed As Boolean = False
    Private _isPseudoConsole As Boolean = False

    ''' <summary>
    ''' Indicates whether the console is a pseudoconsole (ConPTY), which is how Windows Terminal,
    ''' VS Code and other modern terminals host console programs, probing on first access. Under
    ''' ConPTY, the <c> Console </c> window and buffer size setters succeed but resize only
    ''' ConPTY's own copy of the screen, never the terminal's window, so the two disagree about
    ''' the screen's size and drawing breaks. <see cref="ConsoleWindow"/> resizes with a VT escape
    ''' instead when this is <c> True </c>.
    ''' </summary>
    '''
    ''' <remarks>
    ''' ConPTY gives the process a hidden stand-in window of class <c> PseudoConsoleWindow </c>,
    ''' where conhost gives it a real <c> ConsoleWindowClass </c> window. Checking that class covers
    ''' every ConPTY host, unlike the <c> WT_SESSION </c> variable, which only Windows Terminal sets.
    ''' </remarks>
    Public ReadOnly Property IsPseudoConsole As Boolean
        Get
            If Not _pseudoConsoleProbed Then ProbePseudoConsole()
            Return _isPseudoConsole
        End Get
    End Property

    Private Sub ProbePseudoConsole()

        _pseudoConsoleProbed = True

        Dim hwnd = GetConsoleWindow()
        If hwnd = IntPtr.Zero Then Return

        Dim className As New Text.StringBuilder(64)
        If GetClassName(hwnd, className, className.Capacity) = 0 Then Return

        _isPseudoConsole = className.ToString() = "PseudoConsoleWindow"

    End Sub

    ''' <summary>
    ''' Indicates whether the current console supports virtual terminal
    ''' escape sequences, probing on first access. <c> True </c> means inline ANSI
    ''' sequences (color, cursor positioning, clear) are safe to emit; <c> False </c>
    ''' means the caller must fall back to the Win32 console API.
    ''' </summary>
    Public ReadOnly Property HasVT As Boolean
        Get
            If Not _probed Then Probe()
            Return _hasVT
        End Get
    End Property

    ''' <summary>
    ''' Forces the cached <see cref="HasVT"/> result to a specific value, bypassing
    ''' the actual probe. Intended for tests and benchmarks that need to
    ''' exercise both code paths deterministically.
    ''' </summary>
    '''
    ''' <param name="value">The value <see cref="HasVT"/> reports until <see cref="ResetForTesting"/></param>
    Friend Sub SetHasVTForTesting(value As Boolean)
        _hasVT = value
        _probed = True
    End Sub

    ''' <summary>
    ''' Resets the cached probe result so the next access to <see cref="HasVT"/>
    ''' re-probes. Intended for tests that want to restore real behavior
    ''' after a forced override.
    ''' </summary>
    Friend Sub ResetForTesting()
        _hasVT = False
        _probed = False
    End Sub

    Private Sub Probe()

        _probed = True

        Try

            ' If stdout isn't actually a console (redirected to file/pipe,
            ' running under a test harness, etc.) VT escapes would corrupt
            ' the output, so treat it as unsupported regardless of OS.
            If Console.IsOutputRedirected Then Return

            Dim hOut = GetStdHandle(STD_OUTPUT_HANDLE)
            If hOut = IntPtr.Zero OrElse hOut = New IntPtr(-1) Then Return

            Dim mode As UInteger
            If Not GetConsoleMode(hOut, mode) Then Return

            ' Already enabled (Windows Terminal, recent conhost): nothing to do.
            If (mode And ENABLE_VIRTUAL_TERMINAL_PROCESSING) <> 0 Then _hasVT = True : Return
            ' Try to enable; failure means the OS doesn't understand the flag,
            ' which is our signal that this is a legacy console.
            If SetConsoleMode(hOut, mode Or ENABLE_VIRTUAL_TERMINAL_PROCESSING) Then _hasVT = True

        Catch ex As Exception
            ' Defensive: capability detection must never crash the app.
            ' Any failure → fall back to the legacy path.
        End Try

    End Sub

    <DllImport("kernel32.dll", SetLastError:=True)>
    Private Function GetStdHandle(nStdHandle As Integer) As IntPtr
    End Function

    <DllImport("kernel32.dll", SetLastError:=True)>
    Private Function GetConsoleMode(hConsoleHandle As IntPtr, ByRef lpMode As UInteger) As Boolean
    End Function

    <DllImport("kernel32.dll", SetLastError:=True)>
    Private Function SetConsoleMode(hConsoleHandle As IntPtr, dwMode As UInteger) As Boolean
    End Function

    <DllImport("kernel32.dll")>
    Private Function GetConsoleWindow() As IntPtr
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Unicode)>
    Private Function GetClassName(hWnd As IntPtr, lpClassName As Text.StringBuilder, nMaxCount As Integer) As Integer
    End Function

End Module
