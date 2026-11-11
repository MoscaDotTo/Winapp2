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

Imports System.IO

''' <summary>
''' Tests for which local file and which flavor the winapp2.ini update check compares, and
''' which remote file each flavor links to
''' </summary>
<TestClass()> Public Class UpdateTargetTests

    Private _dir As String
    Private _savedFlavor As winapp2ool.WinappFlavor

    <TestInitialize()> Public Sub MakeScratch()

        _dir = Path.Combine(Path.GetTempPath(), $"winapp2ool-updatetarget-{Guid.NewGuid():N}")
        Directory.CreateDirectory(_dir)
        _savedFlavor = winapp2ool.CurrentWinappFlavor

    End Sub

    <TestCleanup()> Public Sub DeleteScratch()

        winapp2ool.CurrentWinappFlavor = _savedFlavor
        If Directory.Exists(_dir) Then Directory.Delete(_dir, recursive:=True)

    End Sub

    ''' <summary>Writes a small winapp2 file named <paramref name="name"/> whose header names <paramref name="flavor"/>, or none</summary>
    ''' <param name="name">The file name</param>
    ''' <param name="flavor">The flavor line's text, or <c> Nothing </c> for none</param>
    Private Sub WriteLocal(name As String, flavor As String)

        Dim flavorLine = If(flavor Is Nothing, "", $"; Flavor: {flavor}" & vbCrLf)
        File.WriteAllText(Path.Combine(_dir, name),
            "; Version: 260219" & vbCrLf & "; # of entries: 1" & vbCrLf & flavorLine & ";" & vbCrLf & vbCrLf &
            "[Test Entry A *]" & vbCrLf & "LangSecRef=3021" & vbCrLf)

    End Sub

    ''' <summary>The local file's own flavor wins over the saved setting</summary>
    <TestMethod()> Public Sub HeaderFlavorWinsOverSetting()

        winapp2ool.CurrentWinappFlavor = winapp2ool.WinappFlavor.CCleaner
        WriteLocal("winapp2.ini", "Tron")

        winapp2ool.findLocalWinapp2(_dir)

        Assert.AreEqual(winapp2ool.WinappFlavor.Tron, winapp2ool.localWa2Flavor)
        Assert.AreEqual("winapp2.ini", winapp2ool.localWa2Name)

    End Sub

    ''' <summary>A winapp2.ini naming no flavor goes by the saved setting</summary>
    <TestMethod()> Public Sub NoHeaderFlavorFallsBackToSetting()

        winapp2ool.CurrentWinappFlavor = winapp2ool.WinappFlavor.BleachBit
        WriteLocal("winapp2.ini", Nothing)

        winapp2ool.findLocalWinapp2(_dir)

        Assert.AreEqual(winapp2ool.WinappFlavor.BleachBit, winapp2ool.localWa2Flavor)

    End Sub

    ''' <summary>A lone winapp2.rules naming no flavor is System Ninja's, whatever the setting says</summary>
    <TestMethod()> Public Sub RulesFileIsSystemNinja()

        winapp2ool.CurrentWinappFlavor = winapp2ool.WinappFlavor.CCleaner
        WriteLocal("winapp2.rules", Nothing)

        winapp2ool.findLocalWinapp2(_dir)

        Assert.AreEqual(winapp2ool.WinappFlavor.SystemNinja, winapp2ool.localWa2Flavor)
        Assert.AreEqual("winapp2.rules", winapp2ool.localWa2Name)

    End Sub

    ''' <summary>With both files present, winapp2.ini is the one compared</summary>
    <TestMethod()> Public Sub IniWinsOverRules()

        WriteLocal("winapp2.ini", "Base")
        WriteLocal("winapp2.rules", "System Ninja")

        winapp2ool.findLocalWinapp2(_dir)

        Assert.AreEqual("winapp2.ini", winapp2ool.localWa2Name)
        Assert.AreEqual(winapp2ool.WinappFlavor.Base, winapp2ool.localWa2Flavor)

    End Sub

    ''' <summary>With no local copy, the setting's flavor and its file name are used</summary>
    <TestMethod()> Public Sub NoLocalFileUsesSetting()

        winapp2ool.CurrentWinappFlavor = winapp2ool.WinappFlavor.SystemNinja

        winapp2ool.findLocalWinapp2(_dir)

        Assert.AreEqual(winapp2ool.WinappFlavor.SystemNinja, winapp2ool.localWa2Flavor)
        Assert.AreEqual("winapp2.rules", winapp2ool.localWa2Name)

    End Sub

    ''' <summary>
    ''' The default flavor is base, so a fresh winapp2ool with no local copy compares and
    ''' downloads the base winapp2.ini
    ''' </summary>
    <TestMethod()> Public Sub DefaultFlavorIsBase()

        winapp2ool.InitDefaultToolSettings()

        winapp2ool.findLocalWinapp2(_dir)

        Assert.AreEqual(winapp2ool.WinappFlavor.Base, winapp2ool.CurrentWinappFlavor)
        Assert.AreEqual(winapp2ool.WinappFlavor.Base, winapp2ool.localWa2Flavor)
        Assert.AreEqual("https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2.ini", winapp2ool.getWinappLink())
        Assert.AreEqual(winapp2ool.WinappFlavor.Base, New winapp2ool.LintOptions().Flavor)

    End Sub

    ''' <summary>
    ''' Each flavor's remote link points at a file of the kind its local copy is saved as, so
    ''' an update never writes a .rules file over a winapp2.ini or the reverse
    ''' </summary>
    <TestMethod()> Public Sub EveryFlavorDownloadsItsOwnKindOfFile()

        For Each flavor As winapp2ool.WinappFlavor In [Enum].GetValues(GetType(winapp2ool.WinappFlavor))

            Dim link = winapp2ool.getWinappLink(flavor)
            Dim isRules = winapp2ool.WinappFlavorInfo.Lookup(flavor).LocalFileName = "winapp2.rules"

            Assert.AreEqual(isRules, link.EndsWith(".rules", StringComparison.OrdinalIgnoreCase), flavor.ToString())

        Next

    End Sub

    ''' <summary>
    ''' Every flavor downloads its own asset from the latest release, so no two flavors share
    ''' a link
    ''' </summary>
    <TestMethod()> Public Sub EveryFlavorDownloadsItsReleaseAsset()

        Const prefix = "https://github.com/MoscaDotTo/Winapp2/releases/latest/download/"
        Dim links As New HashSet(Of String)

        For Each flavor As winapp2ool.WinappFlavor In [Enum].GetValues(GetType(winapp2ool.WinappFlavor))

            Dim link = winapp2ool.getWinappLink(flavor)

            Assert.AreEqual(prefix & winapp2ool.WinappFlavorInfo.Lookup(flavor).AssetName, link, flavor.ToString())
            links.Add(link)

        Next

        Assert.AreEqual(7, links.Count)

    End Sub

End Class
