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
''' Provides methods for managing the Diff module settings, including support methods for syncing to disk
''' Also provides the function which restores the default state of the Diff module's properties
''' </summary>
Module diffsettingshandler

    ''' <summary> 
    ''' Restores the default state of the Diff module's properties 
    ''' </summary>
    Public Sub InitDefaultDiffSettings()

        DownloadDiffFile = Not isOffline
        TrimRemoteFile = Not isOffline
        ShowFullEntries = False
        SaveDiffLog = False
        DiffModuleSettingsChanged = False
        DiffFile1 = New iniFileChooser(Environment.CurrentDirectory, "winapp2.ini", "winapp2-old.ini", mustExist:=True)
        DiffFile2 = New iniFileChooser(Environment.CurrentDirectory, "", "winapp2.ini", mustExist:=True)
        DiffFile3 = New iniFileChooser(Environment.CurrentDirectory, "diff.txt", "diff.txt", mustExist:=False)
        SaveModule2(NameOf(Diff), GetType(diffsettings))

    End Sub

    ''' <summary>
    ''' Loads the Diff module's settings from <c>SettingsFile2</c> into the module's properties.
    ''' </summary>
    Public Sub GetSerializedDiffSettings()

        LoadModule2(NameOf(Diff), GetType(diffsettings))

    End Sub

    ''' <summary>
    ''' Writes the Diff module's current property values into <c>SettingsFile2</c>.
    ''' </summary>
    Public Sub CreateDiffSettingsSection()

        SaveModule2(NameOf(Diff), GetType(diffsettings))

    End Sub

End Module
