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

Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports winapp2ool

''' <summary>
''' Pins which option an empty line runs in each module menu, built from the module's default
''' settings: option 1, the Run option, in every tool menu, and nothing in the settings,
''' download, Minefield and file chooser menus. The main menu is left out because building it
''' checks for updates online.
''' </summary>
<TestClass>
Public Class MenuDefaultsTests

    ''' <summary>Clears the pending header, which every complete menu's builder consumes</summary>
    <TestInitialize>
    Public Sub Reset()

        MenuHeaderText = ""

    End Sub

    <TestMethod>
    Public Sub ToolMenus_DefaultToOptionOne()

        InitDefaultLintSettings()
        InitDefaultTrimSettings()
        initDefaultTransmuteSettings()
        initDefaultFlavorizerSettings()
        InitDefaultDiffSettings()
        initDefaultCCDBSettings()
        InitDefaultBrowserBuilderSettings()
        InitDefaultUWPBuilderSettings()
        InitDefaultEntryBuilderSettings()
        InitDefaultCombineSettings()
        InitDefaultCC7PatcherSettings()

        Dim menus As New Dictionary(Of String, Func(Of MenuSection)) From {
            {"WinappDebug", AddressOf buildLintMenu},
            {"Trim", AddressOf buildTrimMenu},
            {"Transmute", AddressOf buildTransmuteMenu},
            {"Flavorizer", AddressOf buildFlavorizerMenu},
            {"Diff", AddressOf buildDiffMenu},
            {"CCiniDebug", AddressOf buildCCDBMenu},
            {"Entry Lab", AddressOf buildEntryLabMenu},
            {"BrowserBuilder", AddressOf buildBrowserBuilderMenu},
            {"UWPBuilder", AddressOf buildUWPBuilderMenu},
            {"EntryBuilder", AddressOf buildEntryBuilderMenu},
            {"Combine", AddressOf buildCombineMenu},
            {"CC7Patcher", AddressOf buildCC7PatcherMenu}
        }

        For Each menu In menus

            Assert.AreEqual(1, menu.Value()().DefaultOption, menu.Key)

        Next

    End Sub

    <TestMethod>
    Public Sub OtherMenus_HaveNoDefault()

        InitDefaultToolSettings()
        InitDefaultDownloadSettings()
        resetScanSettings()

        Dim chooser As New iniFileChooser(Environment.CurrentDirectory, "winapp2.ini", "winapp2.ini")

        Dim menus As New Dictionary(Of String, Func(Of MenuSection)) From {
            {"Winapp2ool Settings", AddressOf buildMainToolSettingsMenu},
            {"Scan Settings", AddressOf buildAdvSettingsMenu},
            {"Downloader", AddressOf buildDownloadMenu},
            {"Advanced Downloads", AddressOf buildAdvMenu},
            {"Minefield", AddressOf buildMinefieldMenu},
            {"File Chooser", AddressOf chooser.buildMenu}
        }

        For Each menu In menus

            Assert.AreEqual(0, menu.Value()().DefaultOption, menu.Key)

        Next

    End Sub

End Class
