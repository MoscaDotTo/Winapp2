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
''' The main event loop for winapp2ool  
''' </summary>
Public Module launcher

    ''' <summary>
    ''' Performs startup checks, loads settings and handles the command line, then opens the
    ''' main menu. A silent (<c> -s </c>) run exits after the command line with the run's exit
    ''' code instead, saving the global log first when that code is nonzero or
    ''' <c> -writelog </c> was given. A command line run without <c> -s </c> opens the main
    ''' menu once its module returns.
    ''' </summary>
    '''
    ''' <remarks>
    ''' The startup checks may relaunch winapp2ool elevated, decide whether we're offline
    ''' (skipped under <c> -offline </c>), and check the installed .NET Framework against the one
    ''' this build targets. An out-of-date framework, or running from the <c> %temp% </c>
    ''' folder, sets <see cref="cantDownloadExecutable"/>.
    ''' </remarks>
    Public Sub main()

        AddHandler AppDomain.CurrentDomain.UnhandledException,
            Sub(sender As Object, e As UnhandledExceptionEventArgs)
                Dim ex = TryCast(e.ExceptionObject, Exception)
                If ex IsNot Nothing Then exc(ex)
                saveGlobalLog()
            End Sub

        gLog($"Starting application")

        Try

            tryResizeWindow()

            Dim elevatedExitCode = relaunchElevatedIfNeeded()
            If elevatedExitCode.HasValue Then Environment.Exit(elevatedExitCode.Value)

            ' don't bother checking checking the connection if we know we want to be offline
            If Environment.GetCommandLineArgs().Any(Function(a) a.Equals("-offline", StringComparison.OrdinalIgnoreCase)) Then
                isOffline = True
                gLog("Found argument: -offline (skipping connection check)")
            Else
                chkOfflineMode()
            End If

            DotNetFrameworkOutOfDate = Not frameworkRequirementMet(runningTargetFramework()).GetValueOrDefault(True)
            gLog($".NET Framework is out of date. Found release {installedFrameworkRelease()}", DotNetFrameworkOutOfDate)

            Dim curDirIsTemp As Boolean = Environment.CurrentDirectory.Equals(Environment.GetEnvironmentVariable("temp"), StringComparison.InvariantCultureIgnoreCase)
            cantDownloadExecutable = curDirIsTemp OrElse DotNetFrameworkOutOfDate

            LoadWinapp2oolsettings()

            processCommandLineArgs()

            If SuppressOutput Then

                ' A nonzero exit code (e.g. a generator's lint-reconciliation gate trip) or an
                ' explicit -writelog persists the global log to disk before exiting, so headless
                ' and CI runs retain diagnostics that otherwise live only in memory
                saveGlobalLog(Environment.ExitCode <> 0 OrElse SaveGlobalLogOnExit)
                Environment.Exit(Environment.ExitCode)

            End If

            currentVersion = FileVersionInfo.GetVersionInfo(Environment.GetCommandLineArgs(0)).FileVersion
            Console.Title = $"Winapp2ool v{currentVersion}"
            Dim launchHeader = $"Winapp2ool v{currentVersion} - A multitool for winapp2.ini"
            setNextMenuHeaderText(launchHeader, printColor:=ConsoleColor.Cyan)
            initModule(launchHeader, AddressOf buildToolMainMenu)

            FlushIfDirty()

            ' Honor -writelog for interactive runs too; silent mode has already exited above
            saveGlobalLog(SaveGlobalLogOnExit)

        Catch ex As Exception

            exc(ex)

        End Try

    End Sub

    ''' <summary>
    ''' Widens the console window to at least 130 columns for interactive use, then enlarges
    ''' conhost's scrollback buffer. Each menu grows the window's height to fit itself as it's
    ''' drawn (see <see cref="initModule"/>), so we set no height here.
    ''' </summary>
    '''
    ''' <remarks>
    ''' This runs before the command line (and therefore <c> SuppressOutput </c>) is parsed, so we
    ''' pre-scan the raw args for <c> -s </c>, mirroring the <c> -offline </c> pre-scan in
    ''' <see cref="main"/>, to avoid resizing a window that silent mode is about to abandon
    ''' </remarks>
    Private Sub tryResizeWindow()

        If Environment.GetCommandLineArgs().Any(Function(a) a.Equals("-s", StringComparison.OrdinalIgnoreCase)) Then Return

        EnsureMinimumSize(130, 0)
        ExpandScrollback()

    End Sub

End Module
