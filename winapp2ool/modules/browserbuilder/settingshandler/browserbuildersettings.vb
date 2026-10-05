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
''' Holds the settings for the Browser Builder module, which allows winapp2ool to build bespoke entries
''' for individual web browsers through a simple scripting interface. <br /><br />
'''
''' Browser Builder reads all of its input files from a single source directory:
'''
''' <list type="bullet">
''' 
''' <item>
''' <b> <c> chromium.ini </c> </b>: 
''' Generative rulesets for Chromium-based browsers
''' (BrowserInfo + EntryScaffold sections). The build runs when either this or
''' <c> gecko.ini </c> has sections, but a missing one still fails the run.
''' </item>
''' 
''' <item>
''' <b> <c> gecko.ini </c> </b>: 
''' Generative rulesets for Gecko/Goanna-based browsers
''' (BrowserInfo + EntryScaffold sections). The same rule as <c> chromium.ini </c> applies.
''' </item>
''' 
''' <item>
''' <b> <c> browser_additions.ini </c> </b>: 
''' Sections and keys to add after generation (optional)
''' </item>
''' 
''' <item>
''' <b><c> browser_section_removals.ini </c></b>: Sections to remove after generation (optional)
''' </item>
''' <item>
''' 
''' <b><c> browser_name_removals.ini </c></b>: Keys to remove by name after generation (optional)
''' </item>
''' <item>
''' <b><c> browser_value_removals.ini </c></b>: Keys to remove by value after generation (optional)
''' </item>
''' <item>
''' <b><c> browser_section_replacements.ini </c></b>: Sections to replace after generation (optional)
''' </item>
''' <item>
''' <b><c> browser_key_replacements.ini </c></b>: Keys to replace by value after generation (optional)
''' </item>
''' </list>
'''
''' </summary>
Public Module browserbuildersettings

    ''' <summary>
    ''' The source directory containing <c> chromium.ini </c>, <c> gecko.ini </c>, and all
    ''' optional flavor correction files. Only the <c> Dir </c> property is used; <c> Name </c> is ignored.
    ''' </summary>
    Public Property BuilderFile1 As iniFileChooser = New iniFileChooser(Environment.CurrentDirectory, "", "")

    ''' <summary>
    ''' The output file to which Browser Builder saves its generated browser entries.
    ''' </summary>
    Public Property BuilderFile2 As iniFileChooser = New iniFileChooser(Environment.CurrentDirectory, "browsers.ini", mustExist:=False)

    ''' <summary>
    ''' Indicates whether the module settings have been modified from their defaults
    ''' </summary>
    Public Property BrowserBuilderModuleSettingsChanged As Boolean = False

    ''' <summary>
    ''' Restores all BrowserBuilder settings to their defaults and records them in the settings
    ''' file through <see cref="SaveModule"/>. Nothing is written to disk here.
    ''' <see cref="FlushIfDirty"/> does that later, if the save gate allows it
    ''' </summary>
    Public Sub InitDefaultBrowserBuilderSettings()

        BuilderFile1.ResetParams()
        BuilderFile2.ResetParams()
        BrowserBuilderModuleSettingsChanged = False
        SaveModule(NameOf(BrowserBuilder), GetType(browserbuildersettings))

    End Sub

End Module
