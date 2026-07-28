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
''' Holds the settings for the Diff module, which is responsible for assessing changes between
''' two versions of winapp2.ini.
''' This module contains properties that define the input and output files, as well as flags to
''' download the diff file and save the diff log.
''' It also tracks whether the module settings have been modified from their defaults and whether
''' the remote file should be trimmed when downloading.
''' </summary>
Public Module diffsettings

    ''' <summary> 
    ''' The "old" version of winapp2.ini, against which <c> DiffFile2 </c> will be compared 
    ''' <br /> When downloading, this is the local version of winapp2.ini
    ''' </summary>
    Public Property DiffFile1 As iniFileChooser = New iniFileChooser(Environment.CurrentDirectory, "winapp2.ini", "winapp2.ini", mustExist:=True)

    ''' <summary>
    ''' The "new" version of winapp2.ini against which <c> DiffFile1 </c> will be compared
    ''' <br /> When downloading, this is the remote version of winapp2.ini
    ''' </summary>
    Public Property DiffFile2 As iniFileChooser = New iniFileChooser(Environment.CurrentDirectory, "", mustExist:=True)

    ''' <summary>
    ''' The path for the Diff log
    ''' </summary>
    Public Property DiffFile3 As iniFileChooser = New iniFileChooser(Environment.CurrentDirectory, "diff.txt", "diff.txt", mustExist:=False)

    ''' <summary> 
    ''' Indicates that a remote winapp2.ini should be downloaded to use as <c> DiffFile2 </c> 
    ''' </summary>
    Public Property DownloadDiffFile As Boolean = Not isOffline

    ''' <summary> 
    ''' Indicates that the diff output from winapp2ool's global log should be saved to disk 
    ''' </summary>
    Public Property SaveDiffLog As Boolean = False

    ''' <summary> 
    ''' Indicates that the module settings have been modified from their defaults 
    ''' </summary>
    Public Property DiffModuleSettingsChanged As Boolean = False

    ''' <summary> 
    ''' Indicates that the remote (online) should be trimmed for the local system before beginning the Diff 
    ''' </summary>
    Public Property TrimRemoteFile As Boolean = Not isOffline

    ''' <summary>
    ''' Indicates that full entries should be printed in the Diff output. <br/>
    ''' Called "verbose mode" in the menu
    ''' </summary>
    Public Property ShowFullEntries As Boolean = False

    ''' <summary>
    ''' Restores all Diff settings to their defaults and persists the reset to disk
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

End Module
