'    Copyright (C) 2018-2025 Hazel Ward
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
''' Holds the settings for the Combine module, which is responsible for combining multiple ini files
''' potentially spread across numerous subdirectories into a single ini file <br />
''' This module contains properties that define the target directory containing files to be combined 
''' and the output file location for the combined ini file <br />
''' It also tracks whether the module settings have been modified from their defaults. 
''' </summary>
''' 
''' Docs last updated: 2025-08-23 
Public Module combinesettings

    ''' <summary>
    ''' The target directory containing files to be combined <br />
    ''' Default: <c> current directory </c>
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-08-20 | Code last updated: 2025-08-20
    Public Property CombineFile1 As New iniFile(Environment.CurrentDirectory, "")

    ''' <summary>
    ''' The output file location for the combined ini file <br />
    ''' Default: <c> combined.ini </c> in the <c> current directory </c>
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-08-20 | Code last updated: 2025-08-20
    Public Property CombineFile3 As New iniFile(Environment.CurrentDirectory, "combined.ini")

    ''' <summary>
    ''' Indicates that the module settings have been modified from their defaults 
    ''' <br />
    ''' Default: <c> False </c>
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-08-20 | Code last updated: 2025-08-20
    Public Property CombineModuleSettingsChanged As Boolean = False

End Module
