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
''' Holds the settings for the UWPBuilder module, which generates winapp2.ini entries
''' for Universal Windows Platform applications from a small templating DSL.
''' <br /><br />
'''
''' Summary of UWPBuilder files and their expected content:
'''
''' <list>
'''
''' <item>
''' <b><c> UWPFile1 </c></b>
''' <description>
''' Source directory — the root folder containing the UWP scaffold template
''' (<c> UWP.ini </c>) and an <c> AppInfo\ </c> subdirectory whose per-letter
''' <c> *.ini </c> files are combined at runtime to produce the full app list.
''' Only the <c> Dir </c> of this chooser is used; the <c> Name </c> is ignored.
''' </description>
''' </item>
'''
''' <item>
''' <b><c> UWPFile2 </c></b>
''' <description>
''' Output file — where the generated UWP entries are written.
''' This file is consumed by the build pipeline to produce the final winapp2.ini.
''' </description>
''' </item>
'''
''' </list>
'''
''' </summary>
Public Module uwpbuildersettings

    ''' <summary>
    ''' The source directory containing the UWP scaffold template (<c> UWP.ini </c>)
    ''' and the <c> AppInfo\ </c> subfolder with per-letter app definitions.
    ''' Only the <c> Dir </c> property is used; <c> Name </c> is ignored.
    ''' </summary>
    Public Property UWPFile1 As iniFileChooser = New iniFileChooser(Environment.CurrentDirectory, "", "")

    ''' <summary>
    ''' The output file to which the generated UWP entries are saved.
    ''' This file is consumed by the build pipeline to produce the final winapp2.ini.
    ''' </summary>
    Public Property UWPFile2 As iniFileChooser = New iniFileChooser(Environment.CurrentDirectory, "uwp.ini", "uwp.ini", mustExist:=False)

    ''' <summary>
    ''' Indicates that the module settings have been modified from their defaults
    ''' </summary>
    Public Property UWPBuilderModuleSettingsChanged As Boolean = False

    ''' <summary>
    ''' Restores all UWPBuilder settings to their defaults and persists the reset to disk
    ''' </summary>
    Public Sub InitDefaultUWPBuilderSettings()

        UWPFile1 = New iniFileChooser(Environment.CurrentDirectory, "", "")
        UWPFile2 = New iniFileChooser(Environment.CurrentDirectory, "uwp.ini", "uwp.ini", mustExist:=False)
        UWPBuilderModuleSettingsChanged = False
        SaveModule2(NameOf(UWPBuilder), GetType(uwpbuildersettings))

    End Sub

End Module
