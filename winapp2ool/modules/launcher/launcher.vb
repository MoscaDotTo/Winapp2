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
    ''' Performs startup checks and then initializes the winapp2ool main menu module 
    ''' </summary>
    ''' 
    ''' <remarks> 
    ''' Winapp2ool requires an internet connection for some functions 
    ''' .NET 4.6 or higher is required to update the executable 
    ''' When run from the temporary folder, winapp2ool.exe update functionality is disabled
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

            If Not SuppressOutput Then
                Console.WindowWidth = 130
                Console.WindowHeight = 35
                tryExpandScrollback()
            End If

            ' don't bother checking checking the connection if we know we want to be offline
            If Environment.GetCommandLineArgs().Any(Function(a) a.Equals("-offline", StringComparison.OrdinalIgnoreCase)) Then
                isOffline = True
                gLog("Found argument: -offline (skipping connection check)")
            Else
                chkOfflineMode()
            End If

            If Not Environment.Version.ToString = "4.0.30319.42000" Then DotNetFrameworkOutOfDate = True
            gLog($".NET Framework is out of date. Found {Environment.Version}", DotNetFrameworkOutOfDate)

            Dim curDirIsTemp As Boolean = Environment.CurrentDirectory.Equals(Environment.GetEnvironmentVariable("temp"), StringComparison.InvariantCultureIgnoreCase)
            cantDownloadExecutable = curDirIsTemp OrElse DotNetFrameworkOutOfDate

            LoadWinapp2oolsettings()

            processCommandLineArgs()

            If SuppressOutput Then Environment.Exit(0)

            currentVersion = FileVersionInfo.GetVersionInfo(Environment.GetCommandLineArgs(0)).FileVersion
            Console.Title = $"Winapp2ool v{currentVersion}"
            Dim launchHeader = $"Winapp2ool v{currentVersion} - A multitool for winapp2.ini"
            setNextMenuHeaderText(launchHeader, printColor:=ConsoleColor.Cyan)
            initModule(launchHeader, AddressOf printToolMainMenu, AddressOf handleToolMainUserInput)

            FlushIfDirty2()

        Catch ex As Exception

            exc(ex)

        End Try

    End Sub

    ''' <summary>
    ''' Attempts to enlarge the console scrollback buffer so long output isn't truncated
    ''' </summary>
    '''
    ''' <remarks>
    ''' Only effective on classic conhost. Windows Terminal, VS Code's integrated terminal,
    ''' and other modern hosts control their own scrollback and will throw <c> IOException </c>
    ''' or <c> PlatformNotSupportedException </c>; in those cases we silently fall through
    ''' since the host's own scrollback is already in effect
    ''' </remarks>
    Private Sub tryExpandScrollback()

        Try

            Dim targetHeight = Math.Max(Console.WindowHeight, 32000)
            Console.SetBufferSize(Console.BufferWidth, targetHeight)

        Catch ex As IO.IOException
            ' Modern terminal host buffer is host-controlled, nothing to do
        Catch ex As PlatformNotSupportedException
            ' Non-Windows or hosted environment without a conhost buffer
        Catch ex As ArgumentOutOfRangeException
            ' Window is already larger than the requested buffer; not fatal
        End Try

    End Sub

End Module
