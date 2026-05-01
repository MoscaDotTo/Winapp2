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
''' This is the top level module for winapp2ool, through which all other user-facing modules are accessed. The "main menu" 
''' </summary>
Public Module Winapp2ool

    ''' <summary>
    ''' The different flavors of winapp2.ini that winapp2ool supports
    ''' </summary>
    Public Enum WinappFlavor

        ''' <summary>
        ''' Designed for use with CCleaner 
        ''' </summary>
        CCleaner = 0

        ''' <summary>
        ''' The base flavor of winapp2.ini from which all others are derived
        ''' </summary>
        NonCCleaner = 1

        ''' <summary>
        ''' Designed to pass BleachBit's santity checker 
        ''' </summary>
        BleachBit = 2

        ''' <summary>
        ''' Designed to overcome Detection limitations in System Ninja
        ''' </summary>
        SystemNinja = 3

        ''' <summary>
        ''' Captures upstream changes made by Tron to the CCleaner flavor 
        ''' </summary>
        Tron = 4

        ''' <summary>
        ''' Designed for use with CCleaner 7.x and later
        ''' </summary>
        CCleaner7 = 5

    End Enum

    ''' <summary> 
    ''' Indicates that the .NET Framework installed on the current machine is below the targeted version (.NET Framework 4.5)
    ''' </summary>
    Public Property DotNetFrameworkOutOfDate As Boolean = False

    ''' <summary> 
    ''' Indicates that winapp2ool currently has access to the internet
    ''' </summary>
    Public Property isOffline As Boolean = False

    ''' <summary> 
    ''' Indicates that we're unable to download the executable 
    ''' </summary>
    Public Property cantDownloadExecutable As Boolean = False

    ''' <summary>
    ''' Indicates that winapp2ool.exe has already been downloaded during this session and prevents us from redownloading it 
    ''' </summary>
    Public Property alreadyDownloadedExecutable As Boolean = False

    ''' <summary> 
    ''' Checks the version of Windows on the current system and returns it as a Double 
    ''' </summary>
    ''' 
    ''' <returns> 
    ''' The Windows version running on the machine, <br /> 
    ''' <c> 0.0 </c> if the windows version cannot be determined 
    ''' </returns>
    Public Function getWinVer() As Double

        gLog("Checking Windows version")

        Dim osVersion = System.Environment.OSVersion.ToString().Replace("Microsoft Windows NT ", "")
        Dim ver = osVersion.Split(CChar("."))
        Dim out = Val($"{ver(0)}.{ver(1)}")

        gLog($"Found Windows {out}")

        Return out

    End Function

    ''' <summary> 
    ''' Returns the first portion of a registry or filepath parameterization 
    ''' </summary>
    ''' 
    ''' <param name="val"> 
    ''' A Windows filesystem or registry path from which the root should be returned 
    ''' </param>
    ''' 
    ''' <returns> 
    ''' The root directory given by <paramref name="val"/> 
    ''' </returns>
    Public Function getFirstDir(val As String) As String

        Return val.Split(CChar("\"))(0)

    End Function

    ''' <summary> 
    ''' Returns an invariant string representation of a boolean 
    ''' </summary>
    ''' 
    ''' <param name="bool"> 
    ''' A boolean value to return as a string 
    ''' </param>
    Public Function tsInvariant(bool As Boolean) As String

        Return bool.ToString(System.Globalization.CultureInfo.InvariantCulture)

    End Function

    ''' <summary>
    ''' Ensures that an <c>iniFile2</c> has content and informs the user if it does not.
    ''' Unlike the <c>iniFile</c> overload, this does not trigger validation or the File Chooser;
    ''' the caller is responsible for loading the file before calling this.
    ''' </summary>
    '''
    ''' <param name="iFile">
    ''' An <c>iniFile2</c> to be checked for content
    ''' </param>
    '''
    ''' <returns>
    ''' <c>True</c> if the <c>iniFile2</c> has content,
    ''' <br /><c>False</c> otherwise
    ''' </returns>
    Public Function enforceFileHasContent(iFile As iniFile2) As Boolean

        If iFile IsNot Nothing AndAlso iFile.Count > 0 Then Return True

        Dim fileName = If(iFile?.Name, "File")
        Dim out = $"{fileName} was empty or not found"
        setNextMenuHeaderText(out, printColor:=ConsoleColor.DarkRed)
        gLog($"  {out}")

        Return False

    End Function

End Module
