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

        gLog($"Starting application")

        If Not SuppressOutput Then Console.WindowWidth = 130 : Console.WindowHeight = 35

        ' don't bother checking checking the connection if we know we want to be offline
        If Environment.GetCommandLineArgs().Any(Function(a) a.Equals("-offline", StringComparison.OrdinalIgnoreCase)) Then
            isOffline = True
            invertSettingAndRemoveArg(True, "-offline")
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

    End Sub

End Module
