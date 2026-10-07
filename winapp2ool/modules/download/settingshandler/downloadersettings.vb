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
    ''' Where downloaded files are saved. Only <c> Dir </c> holds the user's choice: every
    ''' download sets <c> Name </c> before it starts, and on the command line <c> -1f </c>
    ''' then overrides it.
    ''' </summary>
    Public Property downloadFile As iniFileChooser = New iniFileChooser(Environment.CurrentDirectory, "", mustExist:=False)

    ''' <summary>
    ''' Indicates whether the Downloader module's settings have been changed from their defaults
    ''' </summary>
    Public Property DownloadModuleSettingsChanged As Boolean = False

    ''' <summary>
    ''' Restores all Downloader settings to their defaults and records them in the settings
    ''' file through <see cref="SaveModule"/>. Nothing is written to disk here.
    ''' <see cref="FlushIfDirty"/> does that later, if the save gate allows it
    ''' </summary>
    Public Sub InitDefaultDownloadSettings()

        downloadFile = New iniFileChooser(Environment.CurrentDirectory, "", mustExist:=False)
        DownloadModuleSettingsChanged = False
        SaveModule(NameOf(Downloader), GetType(downloadersettings))

    End Sub

End Module
