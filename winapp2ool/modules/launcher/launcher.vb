'    Copyright (C) 2018-2022 Hazel Ward
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
''' <summary> The main event loop for winapp2ool  </summary>
Public Module launcher

    ''' <summary> Performs startup checks and then initializes the winapp2ool main menu module </summary>
    ''' Docs last updated: 2020-07-14 | Code last updated: 2020-07-14
    Public Sub main()
        gLog($"Starting application")
        ' winapp2ool requires internet access for some functions
        chkOfflineMode()
        currentVersion = FileVersionInfo.GetVersionInfo(Environment.GetCommandLineArgs(0)).FileVersion
        ' Set the console stage 
        Console.Title = $"Winapp2ool v{currentVersion}"
        Console.WindowWidth = 126
        ' winapp2ool requires .NET 4.6 or higher for full functionality, all versions of which report the following version
        If Not Environment.Version.ToString = "4.0.30319.42000" Then DotNetFrameworkOutOfDate = True
        gLog($".NET Framework is out of date. Found {Environment.Version}", DotNetFrameworkOutOfDate)
        ' If winapp2ool is run from the temporary folder the executable cannot be downloaded as all downloads are initally staged to the temporary folder
        cantDownloadExecutable = Environment.CurrentDirectory.Equals(Environment.GetEnvironmentVariable("temp"), StringComparison.InvariantCultureIgnoreCase)
        loadSettings()
        processCommandLineArgs()
        If SuppressOutput Then Environment.Exit(1)
        initModule($"Winapp2ool v{currentVersion} - A multitool for winapp2.ini", AddressOf printToolMainMenu, AddressOf handleToolMainUserInput)
    End Sub

End Module