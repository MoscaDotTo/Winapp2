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
''' Holds the settings for the Trim module, which is responsible for trimming the winapp2.ini file.
''' This module contains properties that define the input and output files, as well as flags for includes and excludes.
''' It also tracks whether the module settings have been modified from their defaults.
''' </summary>
''' Docs last updated: 2025-06-25 | Code last updated: 2025-06-25

Public Module trimsettings

    ''' <summary> 
    ''' The winapp2.ini file that will be trimmed 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-06-25 | Code last updated: 2025-06-25
    Public Property TrimFile1 As New iniFile(Environment.CurrentDirectory, "winapp2.ini", mExist:=True)

    ''' <summary> 
    ''' Holds the path of an iniFile containing the names of Sections who should never be trimmed (an Includes File)
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-06-25 | Code last updated: 2025-06-25
    Public Property TrimFile2 As New iniFile(Environment.CurrentDirectory, "includes.ini")

    ''' <summary> 
    ''' Holds the path where the output file will be saved to disk. Overwrites the input file by default 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-06-25 | Code last updated: 2025-06-25
    Public Property TrimFile3 As New iniFile(Environment.CurrentDirectory, "winapp2.ini", "winapp2-trimmed.ini")

    ''' <summary> 
    ''' Holds the path of an iniFile containing the names of Sections who should always be trimmed (an Excludes file) 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-06-25 | Code last updated: 2025-06-25
    Public Property TrimFile4 As New iniFile(Environment.CurrentDirectory, "excludes.ini")

    ''' <summary> 
    ''' Indicates that we are downloading a winapp2.ini from GitHub 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-06-25 | Code last updated: 2025-06-25
    Public Property DownloadFileToTrim As Boolean = False

    ''' <summary> 
    ''' Indicates that the includes should be consulted while trimming, 
    ''' automatically retaining any entries whose name matches a section within <c> TrimFile2 </c> (the "Includes File")
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-06-25 | Code last updated: 2025-06-25
    Public Property UseTrimIncludes As Boolean = False

    ''' <summary> 
    ''' Indicates that the Excludes should be consulted while trimming, 
    ''' automatically removing any entries whose name matches a section within <c> TrimFile4 (the "Excludes File") </c>
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-06-25 | Code last updated: 2025-06-25
    Public Property UseTrimExcludes As Boolean = False

    ''' <summary> 
    ''' Indicates that the module settings have been modified from their defaults 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-06-25 | Code last updated: 2025-06-25
    Public Property TrimModuleSettingsChanged As Boolean = False

End Module
