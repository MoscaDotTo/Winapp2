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
''' Holds the WinappDebug module's settings. <see cref="SaveModule"/> stores every Boolean and
''' file chooser property here under its own name.
''' </summary>
Public Module lintsettings

    ''' <summary>
    ''' The winapp2.ini file that will be linted
    ''' <br /> Default: <c> winapp2.ini </c> in the current directory, which must exist
    ''' </summary>
    Public Property winappDebugFile1 As New iniFileChooser(Environment.CurrentDirectory, "winapp2.ini", mustExist:=True)

    ''' <summary>
    ''' The save path for the linted file
    ''' <br /> Default: <c> winapp2-debugged.ini </c> in the current directory
    ''' </summary>
    Public Property winappDebugFile3 As New iniFileChooser(Environment.CurrentDirectory, "winapp2-debugged.ini", "winapp2-debugged.ini", mustExist:=False)

    ''' <summary>
    ''' Indicates whether any rule had its repair on when <c> determineScanSettings </c> last
    ''' ran (until then it holds the saved or default value). While <see cref="RepairErrsFound"/> is
    ''' <c> False </c>, this is what lets each rule's own repair toggle apply.
    ''' <br /> Default: <c> False </c>
    ''' </summary>
    Public Property RepairSomeErrsFound As Boolean = False

    ''' <summary>
    ''' Indicates whether the scan settings have been modified from their defaults
    ''' <br /> Default: <c> False </c> 
    ''' </summary>
    Public Property ScanSettingsChanged As Boolean = False

    ''' <summary>
    ''' Indicates whether the module settings have been modified from their defaults
    ''' <br /> Default: <c> False </c> 
    ''' </summary>
    Public Property LintModuleSettingsChanged As Boolean = False

    ''' <summary>
    ''' Indicates whether the linted file is written to <see cref="winappDebugFile3"/> after the lint
    ''' <br /> Default: <c> False </c>
    ''' </summary>
    Public Property SaveChanges As Boolean = False

    ''' <summary>
    ''' Indicates whether every rule repairs what it finds, whatever its own repair toggle says.
    ''' The Scan Settings menu clears it when a changed rule has its repair off.
    ''' <br /> Default: <c> True </c>
    ''' </summary>
    Public Property RepairErrsFound As Boolean = True

    ''' <summary>
    ''' Indicates whether every entry must have a Default key holding <see cref="expectedDefaultValue"/>.
    ''' When <c> True </c>, the Defaults rule stops removing Default keys and reports a wrong value
    ''' instead, and a missing Default key is reported and added.
    ''' <br /> Default: <c> False </c>
    ''' </summary>
    Public Property overrideDefaultVal As Boolean = False

    ''' <summary>
    ''' The expected value for Default keys when auditing their values
    ''' <br /> Default: <c> False </c>
    ''' </summary>
    Public Property expectedDefaultValue As Boolean = False

    ''' <summary>
    ''' Indicates whether the Defaults rule leaves existing Default keys alone instead of
    ''' reporting and removing them. The <c> -keepdefaults </c> CLI flag sets it, for flavors
    ''' that manage their own Default values, such as FluentCleaner. Unlike
    ''' <see cref="overrideDefaultVal"/>, it doesn't check their values or require them to exist.
    ''' <br /> Default: <c> False </c>
    ''' </summary>
    Public Property PreserveDefaultKeys As Boolean = False

End Module
