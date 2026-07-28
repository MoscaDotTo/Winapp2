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
''' <c> ENABLE_VIRTUAL_TERMINAL_PROCESSING </c> is set on stdout. Older
''' versions (XP / Vista / 7 / pre-1607 Win10) return failure from
''' <c> SetConsoleMode </c> when the flag is requested
'''
''' Result is cached after first probe. Tests may force a value via
''' <c> SetHasVTForTesting </c> to exercise either code path.
''' </summary>
Friend Module TerminalCapabilities

    Private Const STD_OUTPUT_HANDLE As Integer = -11
    Private Const ENABLE_VIRTUAL_TERMINAL_PROCESSING As UInteger = 4

    Private _probed As Boolean = False
    Private _hasVT As Boolean = False

    ''' <summary>
    ''' Indicates whether the current console supports virtual terminal
    ''' escape sequences. <c> True </c> means inline ANSI sequences (color,
    ''' cursor positioning, clear) are safe to emit; <c> False </c> means the
    ''' caller must fall back to the legacy Win32 console API.
    ''' </summary>
    Public ReadOnly Property HasVT As Boolean
        Get
            If Not _probed Then Probe()
            Return _hasVT
        End Get
    End Property

    ''' <summary>
    ''' Forces the cached <c> HasVT </c> result to a specific value, bypassing
    ''' the actual probe. Intended for tests and benchmarks that need to
    ''' exercise both code paths deterministically.
    ''' </summary>
    Friend Sub SetHasVTForTesting(value As Boolean)
        _hasVT = value
        _probed = True
    End Sub

    ''' <summary>
    ''' Resets the cached probe result so the next access to <c> HasVT </c>
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
            If (mode And ENABLE_VIRTUAL_TERMINAL_PROCESSING) <> 0 Then
                _hasVT = True
                Return
            End If

            ' Try to enable; failure means the OS doesn't understand the flag,
            ' which is our signal that this is a legacy console.
            If SetConsoleMode(hOut, mode Or ENABLE_VIRTUAL_TERMINAL_PROCESSING) Then
                _hasVT = True
            End If

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

End Module
