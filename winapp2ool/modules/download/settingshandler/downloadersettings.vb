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
''' Holds the settings for the Downloader module, which provides a simple interface
''' for downloading project files from the winapp2 GitHub.
''' </summary>
Public Module downloadersettings

    ''' <summary>
    ''' The directory in which downloaded files are saved
    ''' </summary>
    Public Property downloadFile As iniFileChooser = New iniFileChooser(Environment.CurrentDirectory, "", mustExist:=False)

    ''' <summary>
    ''' Indicates that the Downloader module's settings have been changed from their defaults
    ''' </summary>
    Public Property DownloadModuleSettingsChanged As Boolean = False

    ''' <summary>
    ''' Restores all Downloader settings to their defaults and persists the reset to disk
    ''' </summary>
    Public Sub InitDefaultDownloadSettings()

        downloadFile = New iniFileChooser(Environment.CurrentDirectory, "", mustExist:=False)
        DownloadModuleSettingsChanged = False
        SaveModule2(NameOf(Downloader), GetType(downloadersettings))

    End Sub

End Module
