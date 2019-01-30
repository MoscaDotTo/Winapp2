'    Copyright (C) 2018-2019 Robbie Ward
' 
'    This file is a part of Winapp2ool
' 
'    Winapp2ool is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    Winap2ool is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with Winapp2ool.  If not, see <http://www.gnu.org/licenses/>.
Option Strict On
Imports Microsoft.Win32
''' <summary>
''' This module holds any functions winapp2ool might require for accessing and manipulating the windows registry
''' </summary>
Module RegistryHelper
    ''' <summary>
    ''' Returns the requested key or subkey from the HKEY_LOCAL_MACHINE registry hive
    ''' </summary>
    ''' <param name="subkey">An optional string specifying the path to a subkey in the hive</param>
    ''' <returns></returns>
    Public Function getLMKey(Optional subkey As String = "") As RegistryKey
        Return Registry.LocalMachine.OpenSubKey(subkey)
    End Function

    ''' <summary>
    ''' Returns the requested key or subkey from the HKEY_CLASSES_ROOT registry hive
    ''' </summary>
    ''' <param name="subkey">An optional string specifying the path to a subkey in the hive</param>
    ''' <returns></returns>
    Public Function getCRKey(Optional subkey As String = "") As RegistryKey
        Return Registry.ClassesRoot.OpenSubKey(subkey)
    End Function

    ''' <summary>
    ''' Returns the requested key or subkey from the HKEY_CURRENT_USER registry hive
    ''' </summary>
    ''' <param name="subkey">An optional string specifying the path to a subkey in the hive</param>
    ''' <returns></returns>
    Public Function getCUKey(Optional subkey As String = "") As RegistryKey
        Return Registry.CurrentUser.OpenSubKey(subkey)
    End Function

    ''' <summary>
    ''' Returns the requested key or subkey from the HKEY_USERS registry hive
    ''' </summary>
    ''' <param name="subkey">An optional string specifying the path to a subkey in the hive</param>
    ''' <returns></returns>
    Public Function getUserKey(Optional subkey As String = "") As RegistryKey
        Return Registry.Users.OpenSubKey(subkey)
    End Function
End Module