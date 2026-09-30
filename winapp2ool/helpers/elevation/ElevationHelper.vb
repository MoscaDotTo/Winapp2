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

Imports System.ComponentModel
Imports System.IO
Imports System.Security.Principal
Imports System.Text.RegularExpressions

''' <summary>
''' Gets winapp2ool administrator rights only when the folders it saves to need them.
''' <br /><br />
'''
''' Winapp2ool runs as the user who started it. Most users keep it beside winapp2.ini, which for CCleaner
''' means Program Files, where writing needs administrator rights. At startup we check the working folder and
''' every folder named on the command line, and relaunch once through a UAC prompt if we can't write to one of
''' them. A save that is refused later in the menu offers the same restart. A folder that needs administrator
''' rights to write is also one other programs can't tamper with, so we never run elevated from a folder
''' anyone can write to.
''' </summary>
Module ElevationHelper

    ''' <summary>
    ''' The Win32 error Windows reports when the user declines a UAC prompt
    ''' </summary>
    Private Const ErrorCancelled As Integer = 1223

    ''' <summary>
    ''' Indicates whether the user has already turned down administrator rights this session, so we
    ''' don't ask again
    ''' </summary>
    Private _declined As Boolean = False

    ''' <summary>
    ''' Returns whether this process holds an elevated administrator token
    ''' </summary>
    Friend Function isElevated() As Boolean

        Using identity = WindowsIdentity.GetCurrent()

            Return New WindowsPrincipal(identity).IsInRole(WindowsBuiltInRole.Administrator)

        End Using

    End Function

    ''' <summary>
    ''' Returns whether creating a file in <paramref name="folder"/> is refused for lack of permission.
    ''' A folder that doesn't exist yet is judged by the nearest folder above it that does, since that is
    ''' where it would be created
    ''' </summary>
    '''
    ''' <param name="folder">
    ''' The folder to test
    ''' </param>
    '''
    ''' <returns>
    ''' <c> True </c> if Windows denies us write access to <paramref name="folder"/>, <br />
    ''' <c> False </c> if we can write there, or if the write fails for a reason administrator rights won't fix
    ''' </returns>
    Friend Function needsElevationToWrite(folder As String) As Boolean

        Dim target = folder

        Try

            target = Path.GetFullPath(target)

            While Not Directory.Exists(target)

                target = Path.GetDirectoryName(target)
                If String.IsNullOrEmpty(target) Then Return False

            End While

            Dim probe = Path.Combine(target, $"winapp2ool-write-test-{Guid.NewGuid():N}.tmp")

            Using New FileStream(probe, FileMode.CreateNew, FileAccess.Write, FileShare.None, 1, FileOptions.DeleteOnClose)
            End Using

            Return False

        Catch ex As UnauthorizedAccessException

            Return True

        Catch ex As IOException

            Return False

        Catch ex As ArgumentException

            Return False

        Catch ex As NotSupportedException

            Return False

        End Try

    End Function

    ''' <summary>
    ''' Returns the folders named on the command line with <c> -1d </c>, <c> -2d </c> and so on, resolved the
    ''' way the command line handler resolves them: a leading backslash is relative to the working folder, and a
    ''' final segment containing a dot is a file name
    ''' </summary>
    '''
    ''' <param name="args">
    ''' The command line arguments, not including the executable's own path
    ''' </param>
    Friend Function commandLineFolders(args As IList(Of String)) As List(Of String)

        Dim folders As New List(Of String)

        For i = 0 To args.Count - 2

            If Not Regex.IsMatch(args(i), "^-\d+d$") Then Continue For

            Dim value = args(i + 1).Trim(""""c)
            If value.Length = 0 Then Continue For

            Dim folder = If(value.StartsWith("\", StringComparison.Ordinal), Environment.CurrentDirectory & value, value)
            Dim lastSegment = folder.Split("\"c).Last()
            If lastSegment.Contains(".") Then folder = folder.Substring(0, folder.Length - lastSegment.Length)

            folders.Add(folder)

        Next

        Return folders

    End Function

    ''' <summary>
    ''' Relaunches winapp2ool with administrator rights when it can't write to its working folder or to a folder
    ''' named on the command line, then waits for that copy to finish
    ''' </summary>
    '''
    ''' <returns>
    ''' The elevated copy's exit code, which the caller should exit with, <br />
    ''' <c> Nothing </c> if this process should carry on: every folder is writable, we're already elevated,
    ''' or the user declined the UAC prompt
    ''' </returns>
    Friend Function relaunchElevatedIfNeeded() As Integer?

        If isElevated() Then Return Nothing

        Dim folders = New List(Of String) From {Environment.CurrentDirectory}
        folders.AddRange(commandLineFolders(Environment.GetCommandLineArgs().Skip(1).ToList()))

        Dim protectedFolder = folders.FirstOrDefault(AddressOf needsElevationToWrite)
        If protectedFolder Is Nothing Then Return Nothing

        gLog($"{protectedFolder} isn't writable without administrator rights, relaunching elevated")

        ' -s hasn't been consumed yet, so SuppressOutput can't tell us whether to print
        Dim silent = Environment.GetCommandLineArgs().Skip(1).Any(Function(a) a.Equals("-s", StringComparison.OrdinalIgnoreCase))

        Try

            Using elevated = startElevated(BuildArgumentString(Environment.GetCommandLineArgs().Skip(1)))

                If elevated Is Nothing Then Return Nothing

                cwl("Winapp2ool needs administrator rights to save files there, and is continuing in a new window.", Not silent)

                elevated.WaitForExit()
                Return elevated.ExitCode

            End Using

        Catch ex As Win32Exception

            noteDeclined(ex)
            setNextMenuHeaderText("Winapp2ool can't save files in this folder without administrator rights", printColor:=ConsoleColor.Red)

            Return Nothing

        End Try

    End Function

    ''' <summary>
    ''' Offers to restart winapp2ool with administrator rights after a save to <paramref name="folder"/> was refused.
    ''' Does nothing in silent mode, when we're already elevated, or once the user has turned the offer down
    ''' </summary>
    '''
    ''' <param name="folder">
    ''' The folder the refused save was aimed at
    ''' </param>
    Friend Sub offerElevatedRestart(folder As String)

        If SuppressOutput OrElse _declined OrElse isElevated() Then Return

        cwl()
        cwl($"Winapp2ool needs administrator rights to save files in {folder}")
        cwl(If(saveSettingsToDisk, "Your settings are saved and will carry over.", "Settings you changed this session won't carry over, because Saving Settings is off."))
        cwl("Restart winapp2ool as administrator now? (y/N)")

        Dim answer = Console.ReadLine()

        If answer Is Nothing OrElse Not answer.Trim().Equals("y", StringComparison.OrdinalIgnoreCase) Then

            _declined = True
            Return

        End If

        If saveSettingsToDisk Then FlushIfDirty2()

        Try

            Using startElevated("")
            End Using

            Environment.Exit(0)

        Catch ex As Win32Exception

            noteDeclined(ex)

        End Try

    End Sub

    ''' <summary>
    ''' Starts another copy of winapp2ool through a UAC prompt, in the current working folder
    ''' </summary>
    '''
    ''' <param name="arguments">
    ''' The command line to pass on
    ''' </param>
    '''
    ''' <returns>
    ''' The elevated process, or <c> Nothing </c> if Windows didn't report one
    ''' </returns>
    Private Function startElevated(arguments As String) As Process

        Dim startInfo As New ProcessStartInfo(runningExePath()) With {
            .UseShellExecute = True,
            .Verb = "runas",
            .WorkingDirectory = Environment.CurrentDirectory,
            .Arguments = arguments
        }

        Return Process.Start(startInfo)

    End Function

    ''' <summary>
    ''' Records that administrator rights were refused, so we don't ask again this session
    ''' </summary>
    '''
    ''' <param name="ex">
    ''' The error from the elevated launch
    ''' </param>
    Private Sub noteDeclined(ex As Win32Exception)

        _declined = True

        Dim reason = If(ex.NativeErrorCode = ErrorCancelled, "the UAC prompt was declined", ex.Message)
        gLog($"Continuing without administrator rights: {reason}")

    End Sub

End Module
