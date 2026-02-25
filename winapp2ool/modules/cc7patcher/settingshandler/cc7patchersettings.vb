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
''' Holds the settings for the CC7Patcher module
''' </summary>
Public Module cc7patchersettings

    ''' <summary>
    ''' The winapp2.ini file to be used as input for patching ccleaner.ini
    ''' </summary>
    Public Property CC7PatcherFile1 As New iniFile(Environment.CurrentDirectory, "winapp2.ini", mExist:=True)

    ''' <summary>
    ''' The ccleaner.ini file to be patched
    ''' </summary>
    Public Property CC7PatcherFile2 As New iniFile(Environment.CurrentDirectory, "ccleaner.ini", mExist:=True)

    ''' <summary>
    ''' The output file location (overwrites ccleaner.ini by default)
    ''' </summary>
    Public Property CC7PatcherFile3 As New iniFile(Environment.CurrentDirectory, "ccleaner.ini")

    ''' <summary>
    ''' Indicates that winapp2.ini should be downloaded from GitHub
    ''' </summary>
    Public Property DownloadWinapp2 As Boolean = True

    ''' <summary>
    ''' Indicates that winapp2.ini should be trimmed before patching
    ''' </summary>
    Public Property TrimBeforePatching As Boolean = False

    ''' <summary>
    ''' Indicates that the module settings have been modified from their defaults
    ''' </summary>
    Public Property CC7PatcherModuleSettingsChanged As Boolean = False

End Module