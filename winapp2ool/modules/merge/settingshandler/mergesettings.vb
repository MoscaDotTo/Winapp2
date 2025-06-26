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
''' Holds the settings for the Merge module, which is responsible for conflict resolution when merging two iniFiles together 
''' This module contains properties that define the merge mode, the files to be merged, and whether the settings have been changed from their defaults.
''' </summary>
''' 
''' Docs last updated: 2025-06-25 | Code last updated: 2025-06-25
Public Module mergesettings

    ''' <summary> 
    ''' The file with whose contents <c> MergeFile2 </c> will be merged 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2020-09-14 | Code last updated: 2020-09-14
    Public Property MergeFile1 As New iniFile(Environment.CurrentDirectory, "winapp2.ini")

    ''' <summary> 
    ''' The file whose contents will be merged into <c> MergeFile1 </c>
    ''' </summary>
    ''' 
    ''' Docs last updated: 2020-09-14 | Code last updated: 2020-09-14
    Public Property MergeFile2 As New iniFile(Environment.CurrentDirectory, "")

    ''' <summary> 
    ''' Stores the path to which the merged file should be written back to disk (overwrites <c> MergeFile1 </c> by default) 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2020-09-14 | Code last updated: 2020-09-14
    Public Property MergeFile3 As New iniFile(Environment.CurrentDirectory, "winapp2.ini", "winapp2-merged.ini")

    ''' <summary>
    ''' <c> True </c> if replacing collisions, <c> False </c> if removing them 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2020-09-14 | Code last updated: 2020-09-14
    Public Property mergeMode As Boolean = True

    ''' <summary> 
    ''' Indicates that module's settings have been modified from their defaults 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2020-09-14 | Code last updated: 2020-09-14
    Public Property MergeModuleSettingsChanged As Boolean = False

End Module
