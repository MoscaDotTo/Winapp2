
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
''' Source directory — the folder containing per-letter <c> *.ini </c> source files.
''' Each section in those files describes one output entry; the section header is the
''' output entry name. All files in the directory are combined in alphabetical order
''' at runtime. Only the <c> Dir </c> of this chooser is used; the <c> Name </c> is
''' ignored.
''' </description>
''' </item>
'''
''' <item>
''' <b><c> EntryBuilderFile2 </c></b>
''' <description>
''' Output file — where the generated entries are written. This file is consumed by
''' the build pipeline to produce the final winapp2.ini.
''' </description>
''' </item>
'''
''' <item>
''' <b><c> EntryBuilderFile3 </c></b>
''' <description>
''' Shared WebView scaffold catalog (typically <c> Assembler\Scaffolds\webview.ini </c>).
''' Source for <c> [WebViewScaffold: ...] </c> sections consumed when expanding
''' entries that declare <c> WebViewRoot= </c>. If the file is missing or empty,
''' generation continues with zero scaffold FileKeys emitted and a warning logged.
''' </description>
''' </item>
'''
''' <item>
''' <b><c> EntryBuilderFile4 </c></b>
''' <description>
''' Shared QtWebEngine scaffold catalog (typically <c> Assembler\Scaffolds\qtwebengine.ini </c>).
''' Source for <c> [QtWebEngineScaffold: ...] </c> sections consumed when expanding
''' entries that declare <c> QtWebEngineRoot= </c>. If the file is missing or empty,
''' generation continues with zero QtWebEngine scaffold FileKeys emitted and a warning logged.
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
    ''' The output file to which the generated entries are saved.
    ''' This file is consumed by the build pipeline to produce the final winapp2.ini.
    ''' </summary>
    Public Property EntryBuilderFile2 As iniFileChooser = New iniFileChooser(Environment.CurrentDirectory, "entrybuilder.ini", "entrybuilder.ini", mustExist:=False)

    ''' <summary>
    ''' The shared WebView scaffold catalog consumed by both UWPBuilder and EntryBuilder.
    ''' Typically <c> Assembler\Scaffolds\webview.ini </c>.
    ''' </summary>
    Public Property EntryBuilderFile3 As iniFileChooser = New iniFileChooser(Environment.CurrentDirectory, "webview.ini", "webview.ini", mustExist:=False)

    ''' <summary>
    ''' The shared QtWebEngine scaffold catalog consumed by both UWPBuilder and EntryBuilder.
    ''' Typically <c> Assembler\Scaffolds\qtwebengine.ini </c>.
    ''' </summary>
    Public Property EntryBuilderFile4 As iniFileChooser = New iniFileChooser(Environment.CurrentDirectory, "qtwebengine.ini", "qtwebengine.ini", mustExist:=False)

    ''' <summary>
    ''' Indicates that the module settings have been modified from their defaults
    ''' </summary>
    Public Property EntryBuilderModuleSettingsChanged As Boolean = False

    ''' <summary>
    ''' Restores all EntryBuilder settings to their defaults and persists the reset to disk
    ''' </summary>
    Public Sub InitDefaultEntryBuilderSettings()

        EntryBuilderFile1 = New iniFileChooser(Environment.CurrentDirectory, "", "")
        EntryBuilderFile2 = New iniFileChooser(Environment.CurrentDirectory, "entrybuilder.ini", "entrybuilder.ini", mustExist:=False)
        EntryBuilderFile3 = New iniFileChooser(Environment.CurrentDirectory, "webview.ini", "webview.ini", mustExist:=False)
        EntryBuilderFile4 = New iniFileChooser(Environment.CurrentDirectory, "qtwebengine.ini", "qtwebengine.ini", mustExist:=False)
        EntryBuilderModuleSettingsChanged = False
        SaveModule2(NameOf(EntryBuilder), GetType(entryBuilderSettings))

    End Sub

End Module
