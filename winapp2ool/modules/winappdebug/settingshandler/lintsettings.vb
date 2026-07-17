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
Public Module lintsettings

    ''' <summary> 
    ''' The winapp2.ini file that will be linted 
    ''' </summary>
    Public Property winappDebugFile1 As New iniFileChooser(Environment.CurrentDirectory, "winapp2.ini", mustExist:=True)

    ''' <summary>
    ''' The save path for the linted file
    ''' <br /> Default: <c> winapp2-debugged.ini </c> in the current directory — the input file is not overwritten unless explicitly selected as the save target
    ''' </summary>
    Public Property winappDebugFile3 As New iniFileChooser(Environment.CurrentDirectory, "winapp2-debugged.ini", "winapp2-debugged.ini", mustExist:=False)

    ''' <summary> 
    ''' Indicates that some but not all repairs will run 
    ''' </summary>
    Public Property RepairSomeErrsFound As Boolean = False

    ''' <summary>
    ''' Indicates that the scan settings have been modified from their defaults 
    ''' <br /> Default: <c> False </c> 
    ''' </summary>
    Public Property ScanSettingsChanged As Boolean = False

    ''' <summary>
    ''' Indicates that the module settings have been modified from their defaults 
    ''' <br/> Default: <c> False </c> 
    ''' </summary>
    Public Property LintModuleSettingsChanged As Boolean = False

    ''' <summary> 
    ''' Indicates that the any changes made by the linter should be saved back to disk 
    ''' <br/> Default: <c> False </c> 
    ''' </summary>
    Public Property SaveChanges As Boolean = False

    ''' <summary> 
    ''' Indicates that the linter should attempt to repair errors it finds 
    ''' <br/> Default: <c> True </c>
    ''' </summary>
    Public Property RepairErrsFound As Boolean = True

    ''' <summary> 
    ''' Indicates that Default keys should have their values auited instead of being considered invalid for existing 
    ''' <br /> Default: <c> False </c> 
    ''' </summary>
    Public Property overrideDefaultVal As Boolean = False

    ''' <summary>
    ''' The expected value for Default keys when auditing their values
    ''' <br/> Default: <c> False </c>
    ''' </summary>
    Public Property expectedDefaultValue As Boolean = False

End Module
