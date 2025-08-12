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
''' This is the top level module for winapp2ool, through which all other user-facing modules are accessed. The "main menu" 
''' </summary>
''' 
''' Docs last updated: 2023-07-19 | Code last updated: 2023-07-19
Module Winapp2ool

    ''' <summary> 
    ''' Indicates that the .NET Framework installed on the current machine is below the targeted version (.NET Framework 4.5)
    ''' </summary>
    ''' 
    ''' Docs last updated: 2020-07-14 | Code last updated: 2020-07-14
    Public Property DotNetFrameworkOutOfDate As Boolean = False

    ''' <summary> 
    ''' Indicates that winapp2ool currently has access to the internet
    ''' </summary>
    ''' 
    ''' Docs last updated: 2020-07-14 | Code last updated: 2020-07-14
    Public Property isOffline As Boolean = False

    ''' <summary> 
    ''' Indicates that we're unable to download the executable 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2020-07-14 | Code last updated: 2020-07-14
    Public Property cantDownloadExecutable As Boolean = False

    ''' <summary>
    ''' Indicates that winapp2ool.exe has already been downloaded during this session and prevents us from redownloading it 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2020-07-14 | Code last updated: 2020-07-14
    Public Property alreadyDownloadedExecutable As Boolean = False

    ''' <summary> 
    ''' Checks the version of Windows on the current system and returns it as a Double 
    ''' </summary>
    ''' 
    ''' <returns> 
    ''' The Windows version running on the machine, <br /> 
    ''' <c> 0.0 </c> if the windows version cannot be determined 
    ''' </returns>
    ''' 
    ''' Docs last updated: 2020-07-14 | Code last updated: 2020-07-14
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
    ''' 
    ''' Docs last updated: 2020-07-14 | Code last updated: 2020-07-14
    Public Function getFirstDir(val As String) As String

        Return val.Split(CChar("\"))(0)

    End Function

    ''' <summary> 
    ''' Ensures that an <c> iniFile </c> has content and informs the user if it does not.
    ''' </summary>
    ''' 
    ''' <param name="iFile">
    ''' An <c> iniFile </c> to be checked for content 
    ''' </param>
    ''' 
    ''' <returns> 
    ''' <c> True </c> if the <c> iniFile </c> has content, 
    ''' <br /> <c> False </c>otherwise
    ''' </returns>
    ''' Docs last updated: 2020-07-14 | Code last updated: 2020-07-14
    Public Function enforceFileHasContent(iFile As iniFile) As Boolean

        iFile.validate()

        If iFile.Sections.Count = 0 Then

            setHeaderText($"{iFile.Name} was empty or not found", True)
            gLog($"{iFile.Name} was empty or not found", indent:=True)

            Return False

        End If

        Return True

    End Function

    ''' <summary> 
    ''' Returns an invariant string representation of a boolean 
    ''' </summary>
    ''' 
    ''' <param name="bool"> 
    ''' A boolean value to return as a string 
    ''' </param>
    ''' 
    ''' Docs last updated: 2023-07-19 | Code last updated: 2023-07-19
    Public Function tsInvariant(bool As Boolean) As String

        Return bool.ToString(System.Globalization.CultureInfo.InvariantCulture)

    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' 
    ''' <param name="flavorFile">
    ''' 
    ''' </param>
    ''' 
    ''' <returns>
    ''' 
    ''' </returns>
    ''' 
    ''' Docs last updated: 2025-08-01 | Code last updated: 2025-08-06
    Public Function getFileMenuColor(menuFile As iniFile) As ConsoleColor

        Return If(menuFile.Name.Length > 0, ConsoleColor.Green, ConsoleColor.Red)

    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' 
    ''' <param name="menuFile">
    ''' 
    ''' </param>
    ''' 
    ''' <returns>
    ''' 
    ''' </returns>
    ''' 
    ''' Docs last updated: 2025-08-01 | Code last updated: 2025-08-06
    Public Function getFileMenuName(menuFile As iniFile) As String

        Return If(menuFile.Name.Length > 0, replDir(menuFile.Path), "Not specified")

    End Function

End Module