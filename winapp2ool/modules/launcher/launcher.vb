'    Copyright (C) 2018-2024 Hazel Ward
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
    ''' 
    ''' Docs last updated: 2020-07-14 | Code last updated: 2020-07-14
    Public Sub main()

        gLog($"Starting application")

        chkOfflineMode()

        If Not Environment.Version.ToString = "4.0.30319.42000" Then DotNetFrameworkOutOfDate = True
        gLog($".NET Framework is out of date. Found {Environment.Version}", DotNetFrameworkOutOfDate)

        cantDownloadExecutable = Environment.CurrentDirectory.Equals(Environment.GetEnvironmentVariable("temp"), StringComparison.InvariantCultureIgnoreCase)

        loadSettings()

        processCommandLineArgs()

        If SuppressOutput Then Environment.Exit(1)

        currentVersion = FileVersionInfo.GetVersionInfo(Environment.GetCommandLineArgs(0)).FileVersion
        Console.Title = $"Winapp2ool v{currentVersion}"
        Console.WindowWidth = 126

        initModule($"Winapp2ool v{currentVersion} - A multitool for winapp2.ini", AddressOf printToolMainMenu, AddressOf handleToolMainUserInput)

    End Sub

End Module