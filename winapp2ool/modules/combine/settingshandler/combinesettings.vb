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
''' Holds the settings for the Combine module, which merges all ini files from a target
''' directory and its subdirectories into a single output file.
''' <br /><br />
'''
''' Summary of Combine files:
'''
''' <list>
'''
''' <item>
''' <b><c> CombineFile1 </c></b>
''' <description>
''' Target directory — the root folder to scan for <c> .ini </c> files.
''' Only the <c> Dir </c> property is used; <c> Name </c> is intentionally empty.
''' </description>
''' </item>
'''
''' <item>
''' <b><c> CombineFile3 </c></b>
''' <description>
''' Output file — the combined <c> .ini </c> file written after all source files are merged.
''' </description>
''' </item>
'''
''' </list>
'''
''' </summary>
Public Module combinesettings

    ''' <summary>
    ''' The target directory to scan for <c> .ini </c> files to combine. <br />
    ''' Only the <c> Dir </c> property is used; <c> Name </c> is intentionally empty. <br />
    ''' Default: <c> current directory </c>
    ''' </summary>
    Public Property CombineFile1 As iniFileChooser = New iniFileChooser(Environment.CurrentDirectory, "", "", mustExist:=False)

    ''' <summary>
    ''' The output file to which the combined ini sections are saved. <br />
    ''' Default: <c> combined.ini </c> in the <c> current directory </c>
    ''' </summary>
    Public Property CombineFile3 As iniFileChooser = New iniFileChooser(Environment.CurrentDirectory, "combined.ini", "combined.ini", mustExist:=False)

    ''' <summary>
    ''' When <c> True </c>, a section name appearing in more than one input file is
    ''' treated as an error: the collisions are reported, the output is not saved, and
    ''' a nonzero process exit code is set so scripted builds fail. When <c> False </c>
    ''' (the default), same-named sections are merged key-wise as before, with a
    ''' warning listing the collisions. Intended for the winapp2.ini build pipeline,
    ''' where the staged input files are expected to be disjoint and any collision is a
    ''' maintainer error. CLI: <c> -strict </c>
    ''' </summary>
    Public Property CombineStrictNames As Boolean = False

    ''' <summary>
    ''' Indicates that the module settings have been modified from their defaults
    ''' </summary>
    Public Property CombineModuleSettingsChanged As Boolean = False

    ''' <summary>
    ''' Restores all Combine settings to their defaults and persists the reset to disk
    ''' </summary>
    Public Sub InitDefaultCombineSettings()

        CombineFile1.ResetParams()
        CombineFile3.ResetParams()
        CombineStrictNames = False
        CombineModuleSettingsChanged = False
        SaveModule2(NameOf(Combine), GetType(combinesettings))

    End Sub

End Module
