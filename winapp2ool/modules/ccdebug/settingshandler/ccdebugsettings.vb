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
''' Holds the settings for the CCiniDebug module, which is responsible for debugging ccleaner.ini files.
''' This module contains properties that define the input and output files, as well as flags to prune stale entries, sort the file, and save the debugged file.
''' It also tracks whether the module settings have been modified from their defaults.
''' </summary>
''' 
''' Docs last updated: 2025-06-25 | Code last updated: 2025-06-25
Public Module ccdebugsettings

    ''' <summary> 
    ''' The winapp2.ini file that ccleaner.ini may optionally be checked against 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2020-07-18 | Code last updated: 2020-07-18
    Public Property CCDebugFile1 As New iniFile(Environment.CurrentDirectory, "winapp2.ini")

    ''' <summary> 
    ''' The ccleaner.ini file to be debugged 
    ''' </summary>
    '''
    ''' Docs last updated: 2020-07-18 | Code last updated: 2020-07-18
    Public Property CCDebugFile2 As New iniFile(Environment.CurrentDirectory, "ccleaner.ini", mExist:=True)

    ''' <summary> 
    ''' Holds the path for the debugged file that will be saved to disk. Overwrites ccleaner.ini by default 
    ''' </summary>
    '''
    ''' Docs last updated: 2020-07-18 | Code last updated: 2020-07-18
    Public Property CCDebugFile3 As New iniFile(Environment.CurrentDirectory, "ccleaner.ini", "ccleaner-debugged.ini")

    ''' <summary>
    ''' Indicates that stale winapp2.ini entries should be pruned from ccleaner.ini 
    ''' <br/> Default: <c> True </c>
    ''' </summary>
    '''
    ''' <remarks> 
    ''' A "stale" entry is one that appears in an (app) key in <c> CCDebugFile2 </c> but does not have a corresponding section in <c> CCDebugFile1</c> 
    ''' </remarks>
    '''
    ''' Docs last updated: 2020-07-18 | Code last updated: 2020-07-18
    Public Property PruneStaleEntries As Boolean = True

    '''<summary> 
    '''Indicates that the debugged file should be saved back to disk 
    '''<br/> Default: <c> True </c> 
    '''</summary>
    '''
    ''' Docs last updated: 2020-07-18 | Code last updated: 2020-07-18
    Public Property SaveDebuggedFile As Boolean = True

    '''<summary> 
    '''Indicates that the contents of ccleaner.ini should be sorted alphabetically 
    '''<br /> Default: <c> True </c>
    '''</summary>
    '''
    ''' Docs last updated: 2020-07-18 | Code last updated: 2020-07-18
    Public Property SortFileForOutput As Boolean = True

    '''<summary> 
    '''Indicates that the module's settings have been modified from their defaults 
    '''</summary>
    '''
    ''' Docs last updated: 2020-07-18 | Code last updated: 2020-07-18
    Public Property CCDBSettingsChanged As Boolean = False

End Module
