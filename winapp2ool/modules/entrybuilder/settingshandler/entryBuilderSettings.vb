
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
''' Holds the settings for the EntryBuilder module, which generates winapp2.ini
''' entries from a shorthand DSL (pass-through plus winapp2ool-private shorthand
''' keys expanded into standard winapp2 form).
''' <br /><br />
'''
''' Summary of EntryBuilder files and their expected content:
'''
''' <list>
'''
''' <item>
''' <b><c> EntryBuilderFile1 </c></b>
''' <description>
''' Source directory: the folder containing per-letter <c> *.ini </c> source files.
''' Each section in those files describes one output entry; the section header is the
''' output entry name. The top-level <c> *.ini </c> files are combined in sorted order
''' at runtime. Only the <c> Dir </c> of this chooser is used; the <c> Name </c> is
''' ignored.
''' </description>
''' </item>
'''
''' <item>
''' <b><c> EntryBuilderFile2 </c></b>
''' <description>
''' Output file: where the generated entries are written. With
''' <see cref="EntryBuilderSplitOutput"/> only its <c> Dir </c> is used, and the 27
''' per-letter files written there are what the build pipeline consumes.
''' </description>
''' </item>
'''
''' <item>
''' <b><c> EntryBuilderFile3 </c></b>
''' <description>
''' Shared scaffold directory (typically <c> Assembler\Scaffolds </c>). Every
''' <c> *.ini </c> in it is loaded at once and each scaffold family is derived from its
''' section headers (<c> [WebViewScaffold: ...] </c>, <c> [QtWebEngineScaffold: ...] </c>,
''' <c> [ElectronScaffold: ...] </c>), consumed when expanding entries that declare the
''' matching root key. Only the <c> Dir </c> of this chooser is used; the <c> Name </c> is
''' ignored. A missing directory warns once and the run continues with no scaffold
''' FileKeys. A family with no catalog in the directory yields none either, and each
''' scaffold an entry requests from it warns as unknown.
''' </description>
''' </item>
'''
''' </list>
'''
''' </summary>
Public Module entryBuilderSettings

    ''' <summary>
    ''' The source directory containing per-letter <c> *.ini </c> shorthand source files.
    ''' Only the <c> Dir </c> property is used; <c> Name </c> is ignored.
    ''' </summary>
    Public Property EntryBuilderFile1 As iniFileChooser = New iniFileChooser(Environment.CurrentDirectory, "", "")

    ''' <summary>
    ''' The output file to which the generated entries are saved. With
    ''' <see cref="EntryBuilderSplitOutput"/> only its <c> Dir </c> is used, as the folder for
    ''' the per-letter files.
    ''' </summary>
    Public Property EntryBuilderFile2 As iniFileChooser = New iniFileChooser(Environment.CurrentDirectory, "entrybuilder.ini", "entrybuilder.ini", mustExist:=False)

    ''' <summary>
    ''' The shared scaffold directory consumed by both UWPBuilder and EntryBuilder, holding
    ''' one catalog file per engine family. Typically <c> Assembler\Scaffolds </c>.
    ''' Only the <c> Dir </c> property is used; <c> Name </c> is ignored.
    ''' </summary>
    Public Property EntryBuilderFile3 As iniFileChooser = New iniFileChooser(Environment.CurrentDirectory, "", "")

    ''' <summary>
    ''' Indicates whether output is written as per-letter artifact files
    ''' (<c> #.ini </c>, <c> A.ini </c> ... <c> Z.ini </c>) in the save target's directory,
    ''' bucketed by each entry name's first character, instead of a single output file.
    ''' The build pipeline uses it to produce the committed <c> Assembler\Entries </c>
    ''' artifacts. CLI: <c> -split </c>, which flips the current value.
    ''' </summary>
    Public Property EntryBuilderSplitOutput As Boolean = False

    ''' <summary>
    ''' Indicates whether the module settings have been modified from their defaults
    ''' </summary>
    Public Property EntryBuilderModuleSettingsChanged As Boolean = False

    ''' <summary>
    ''' Restores all EntryBuilder settings to their defaults and records them in the in-memory
    ''' settings file via <see cref="SaveModule"/>. This doesn't write to disk by itself.
    ''' </summary>
    Public Sub InitDefaultEntryBuilderSettings()

        EntryBuilderFile1 = New iniFileChooser(Environment.CurrentDirectory, "", "")
        EntryBuilderFile2 = New iniFileChooser(Environment.CurrentDirectory, "entrybuilder.ini", "entrybuilder.ini", mustExist:=False)
        EntryBuilderFile3 = New iniFileChooser(Environment.CurrentDirectory, "", "")
        EntryBuilderSplitOutput = False
        EntryBuilderModuleSettingsChanged = False
        SaveModule(NameOf(EntryBuilder), GetType(entryBuilderSettings))

    End Sub

End Module
