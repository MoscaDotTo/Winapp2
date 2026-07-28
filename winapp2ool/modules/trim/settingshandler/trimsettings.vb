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
''' Holds the settings for the <c> Trim </c> module, which trims a winapp2.ini file
''' such that it contains only entries relevant to the current system.
''' <br /><br />
'''
''' Summary of Trim files and their expected content:
'''
''' <list>
'''
''' <item>
''' <b><c> TrimFile1 </c></b>
''' <description>
''' The winapp2.ini file to be trimmed. Ignored when
''' <c> DownloadFileToTrim </c> is <c> True </c>.
''' </description>
''' </item>
'''
''' <item>
''' <b><c> TrimFile2 </c></b>
''' <description>
''' Includes file — entry names listed here are never removed, regardless of
''' detection. Only consulted when <c> UseTrimIncludes </c> is <c> True </c>.
''' </description>
''' </item>
'''
''' <item>
''' <b><c> TrimFile3 </c></b>
''' <description>
''' Output file — where the trimmed winapp2.ini is saved. Defaults to overwriting
''' the input file.
''' </description>
''' </item>
'''
''' <item>
''' <b><c> TrimFile4 </c></b>
''' <description>
''' Excludes file — entry names listed here are always removed. Only consulted
''' when <c> UseTrimExcludes </c> is <c> True </c>.
''' </description>
''' </item>
'''
''' </list>
'''
''' </summary>
Public Module trimsettings

    ''' <summary>
    ''' The winapp2.ini file to be trimmed.
    ''' Ignored when <c> DownloadFileToTrim </c> is <c> True </c>.
    ''' <br/> Default: <c> winapp2.ini</c>
    ''' </summary>
    Public Property TrimFile1 As iniFileChooser = New iniFileChooser(Environment.CurrentDirectory, "winapp2.ini")

    ''' <summary>
    ''' Includes file: sections listed here are never trimmed, regardless of detection criteria.
    ''' Only consulted when <c> UseTrimIncludes </c> is <c> True </c>.
    ''' <br /> Default: <c> includes.ini </c>
    ''' </summary>
    Public Property TrimFile2 As iniFileChooser = New iniFileChooser(Environment.CurrentDirectory, "includes.ini", mustExist:=False)

    ''' <summary>
    ''' Output file: Location on disk to which the output will be saved 
    ''' <br /> Default: <c> winapp2.ini </c>
    ''' <br /> Default rename: <c> winapp2-trimmed.ini </c>
    ''' </summary>
    Public Property TrimFile3 As iniFileChooser = New iniFileChooser(Environment.CurrentDirectory, "winapp2.ini", "winapp2-trimmed.ini", mustExist:=False)

    ''' <summary>
    ''' Excludes file — sections listed here are always trimmed, regardless of detection criteria.
    ''' Only consulted when <c> UseTrimExcludes </c> is <c> True </c>.
    ''' <br /> 
    ''' </summary>
    Public Property TrimFile4 As iniFileChooser = New iniFileChooser(Environment.CurrentDirectory, "excludes.ini", mustExist:=False)

    ''' <summary>
    ''' Indicates that the latest winapp2.ini should be downloaded from GitHub as the input file
    ''' <br /> Default: <c> False </c>
    ''' </summary>
    Public Property DownloadFileToTrim As Boolean = False

    ''' <summary>
    ''' Indicates that the includes file is consulted during trimming,
    ''' automatically retaining entries whose name appears in <c> TrimFile2 </c>
    ''' <br /> Default: <c> False </c>
    ''' </summary>
    Public Property UseTrimIncludes As Boolean = False

    ''' <summary>
    ''' Indicates that the excludes file is consulted during trimming,
    ''' automatically removing entries whose name appears in <c> TrimFile4 </c>
    ''' <br /> Default: <c> False </c>
    ''' </summary>
    Public Property UseTrimExcludes As Boolean = False

    ''' <summary>
    ''' Indicates that the module settings have been modified from their defaults
    ''' </summary>
    Public Property TrimModuleSettingsChanged As Boolean = False

    ''' <summary>
    ''' Restores all <c> Trim </c> settings to their defaults and persists the reset to disk
    ''' </summary>
    Public Sub InitDefaultTrimSettings()

        TrimFile1.ResetParams()
        TrimFile2.ResetParams()
        TrimFile3.ResetParams()
        TrimFile4.ResetParams()
        DownloadFileToTrim = False
        UseTrimIncludes = False
        UseTrimExcludes = False
        TrimModuleSettingsChanged = False
        SaveModule2(NameOf(Trim), GetType(trimsettings))

    End Sub

End Module
