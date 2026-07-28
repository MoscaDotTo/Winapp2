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
''' Holds the settings for the CC7Patcher module
''' </summary>
Public Module cc7patchersettings

    ''' <summary>
    ''' The winapp2.ini file to be used as input for patching ccleaner.ini
    ''' </summary>
    Public Property CC7PatcherFile1 As iniFileChooser = New iniFileChooser(Environment.CurrentDirectory, "winapp2.ini", "winapp2.ini", mustExist:=True)

    ''' <summary>
    ''' The ccleaner.ini file to be patched
    ''' </summary>
    Public Property CC7PatcherFile2 As iniFileChooser = New iniFileChooser(Environment.CurrentDirectory, "ccleaner.ini", "ccleaner.ini", mustExist:=True)

    ''' <summary>
    ''' The output file location (overwrites ccleaner.ini by default)
    ''' </summary>
    Public Property CC7PatcherFile3 As iniFileChooser = New iniFileChooser(Environment.CurrentDirectory, "ccleaner.ini", "ccleaner.ini", mustExist:=False)

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

    ''' <summary>
    ''' Restores all CC7Patcher settings to their defaults and persists the reset to disk
    ''' </summary>
    Public Sub InitDefaultCC7PatcherSettings()

        DownloadWinapp2 = True
        TrimBeforePatching = False
        CC7PatcherModuleSettingsChanged = False
        CC7PatcherFile1 = New iniFileChooser(Environment.CurrentDirectory, "winapp2.ini", "winapp2.ini", mustExist:=True)
        CC7PatcherFile2 = New iniFileChooser(Environment.CurrentDirectory, "ccleaner.ini", "ccleaner.ini", mustExist:=True)
        CC7PatcherFile3 = New iniFileChooser(Environment.CurrentDirectory, "ccleaner.ini", "ccleaner.ini", mustExist:=False)
        SaveModule2(NameOf(CC7Patcher), GetType(cc7patchersettings))

    End Sub

End Module
