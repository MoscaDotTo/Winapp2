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
''' Holds the settings for the Transmute module, which makes changes to a 
''' base iniFile based on the content of a source inifile 
''' This module contains properties that will be synced to disk which define the current state of
''' the Transmutator and its sub modes, the file locations of the transmute files, 
''' and whether the settings have been changed from their defaults.
''' </summary>
Public Module transmuteSettings

    ''' <summary> 
    ''' The 'base' file for Transmute, the one whose content will be modified by the transmute process
    ''' based on the contents of the 'source' file 
    ''' </summary>
    Public Property TransmuteFile1 As New iniFile(Environment.CurrentDirectory, "winapp2.ini")

    ''' <summary> 
    ''' The 'source' file for Transmute, the one whose content will be used by the transmute process
    ''' to make changes to the 'base' file
    ''' </summary>
    Public Property TransmuteFile2 As New iniFile(Environment.CurrentDirectory, "")

    ''' <summary> 
    ''' Stores the path to which the Transmuted file should be written back to disk 
    ''' (overwrites <c> TransmuteFile1 </c> by default) 
    ''' </summary>
    Public Property TransmuteFile3 As New iniFile(Environment.CurrentDirectory, "winapp2.ini", "winapp2-transmuted.ini")

    ''' <summary> 
    ''' Indicates that module's settings have been modified from their defaults 
    ''' </summary>
    Public Property TransmuteModuleSettingsChanged As Boolean = False

    ''' <summary>
    ''' Indicates the saved <c> Transmutator </c> is <c> Add </c> <br />
    ''' Default: <c> False </c>
    ''' </summary>
    Public Property TransmuteModeIsAdd As Boolean = True

    ''' <summary>
    ''' Indicates the saved <c> Transmutator </c> is <c> Remove </c> <br />
    ''' Default: <c> True </c>
    ''' </summary>
    Public Property TransmuteModeIsRemove As Boolean = False

    ''' <summary>
    ''' Indicates the saved <c> Transmutator </c> is <c> Replace </c> <br />
    ''' Default: <c> False </c>
    ''' </summary>
    Public Property TransmuteModeIsReplace As Boolean = False

    ''' <summary>
    ''' Indicates the saved <c> Replace Transmutator </c> sub mode is by Section <br />
    ''' Default: <c> False </c>
    ''' </summary>
    Public Property ReplaceModeIsBySection As Boolean = False

    ''' <summary>
    ''' Indicates the saved <c> Replace Transmutator </c> sub mode is by Key <br />
    ''' Default: <c> True </c>
    ''' </summary>
    Public Property ReplaceModeIsByKey As Boolean = True

    ''' <summary>
    ''' Indicates the saved <c> Remove Transmutator </c> sub mode is by Section <br />
    ''' Default: <c> False </c>
    ''' </summary>
    Public Property RemoveModeIsBySection As Boolean = False

    ''' <summary>
    ''' Indicates the saved <c> Remove Transmutator </c> sub mode is by Key <br />
    ''' Default: <c> True </c>
    ''' </summary>
    Public Property RemoveModeIsByKey As Boolean = False

    ''' <summary>
    ''' Indicates the saved <c> Remove by Key Transmutator </c> sub mode is by Name <br />
    ''' Default: <c> True </c>
    ''' </summary>
    Public Property RemoveKeyModeIsByName As Boolean = True

    ''' <summary>
    ''' Indicates the saved <c> Remove by Key Transmutator </c> sub mode is by Section <br />
    ''' Default: <c> False </c>
    ''' </summary>
    Public Property RemoveKeyModeIsByValue As Boolean = False

    ''' <summary>
    ''' Indicates that <c> <cref> TransmuteFile3 </cref> </c> should be saved with winapp2.ini formatting <br />
    ''' Default: <c> True </c>
    ''' </summary>
    Public Property UseWinapp2Syntax As Boolean = True

End Module