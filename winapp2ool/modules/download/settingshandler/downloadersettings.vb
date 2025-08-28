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
''' Holds the settings for the Downloader module, which is responsible for downloading files.
''' This module contains properties that define the file to be downloaded and whether the settings have been changed from their defaults.
''' </summary>
''' 
''' Docs last updated: 2025-06-25 | Code last updated: 2025-06-25
Public Module downloadersettings

    '''<summary> 
    '''Holds the path of any files to be saved by the Downloader 
    '''</summary>
    '''
    ''' Docs last updated: 2020-09-14 | Code last updated: 2020-09-14
    Public Property downloadFile As iniFile = New iniFile(Environment.CurrentDirectory, "")

    ''' <summary> 
    ''' Indicates that the Downloader module's settings have been changed from their defaults 
    '''</summary>
    '''
    ''' Docs last updated: 2020-09-14 | Code last updated: 2020-09-14
    Public Property DownloadModuleSettingsChanged As Boolean = False

End Module
