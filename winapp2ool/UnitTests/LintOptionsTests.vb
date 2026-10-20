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
Imports System.Text

''' <summary>
''' Tests for <see cref="winapp2ool.LintOptions"/>: the repair flags it derives from the rule
''' toggles, and the independence of its rule copies from the user's
''' </summary>
<TestClass()> Public Class LintOptionsTests

    ''' <summary>An entry whose only error is the casing of its LangSecRef key name</summary>
    Private Const MiscasedEntry As String = "[App *]" & vbCrLf &
                                            "langsecref=3021" & vbCrLf &
                                            "Detect=HKCU\Software\App" & vbCrLf &
                                            "FileKey1=%AppData%\App|*.log" & vbCrLf

    ''' <summary>Returns an <c> iniFile </c> parsed from <paramref name="text"/></summary>
    ''' <param name="text">Literal ini text</param>
    Private Shared Function MakeIni(text As String) As winapp2ool.iniFile

        Using reader As New StreamReader(New MemoryStream(Encoding.UTF8.GetBytes(text)))
            Return winapp2ool.iniFile.FromStream(reader, "", "winapp2ool.ini")
        End Using

    End Function

    ''' <summary>Starts each test from the default lint settings, whatever an earlier test left</summary>
    <TestInitialize()> Public Sub StartFromDefaults()

        winapp2ool.InitDefaultLintSettings()

    End Sub

    ''' <summary>Restores the default lint settings after each test</summary>
    <TestCleanup()> Public Sub ResetLintSettings()

        winapp2ool.InitDefaultLintSettings()

    End Sub

    ''' <summary>
    ''' A rule saved with its repair off stays unrepaired after a reload. The settings file of a
    ''' build before <c> LintOptions </c> also holds the stale cached flag
    ''' <c> RepairErrsFound=True </c>, which used to force every repair on until the user next
    ''' visited Scan Settings.
    ''' </summary>
    <TestMethod()> Public Sub SavedRepairOff_HonoredAfterReload()

        Dim savedFile = winapp2ool.SettingsHandler.SettingsFile

        Try

            winapp2ool.SettingsHandler.SettingsFile = MakeIni("[WinappDebug]" & vbCrLf &
                                                              "RepairErrsFound=True" & vbCrLf &
                                                              "Casing_Scan=True" & vbCrLf &
                                                              "Casing_Repair=False" & vbCrLf)
            winapp2ool.LoadModule(NameOf(winapp2ool.WinappDebug), GetType(winapp2ool.lintsettings))
            winapp2ool.LoadLintRulesFromSettings()

            Dim options = winapp2ool.LintOptions.FromSettings()

            Assert.IsFalse(options.RepairAll)
            Assert.IsFalse(options.Casing.fixFormat(), "Casing repairs although its repair is saved off")
            Assert.IsTrue(options.Alpha.fixFormat(), "Other rules keep their repairs")

            Dim wa2 As New winapp2ool.winapp2file(MakeIni(MiscasedEntry))
            winapp2ool.WinappDebug.Debug(wa2, options)

            Assert.IsTrue(winapp2ool.WinappDebug.ErrorsFound > 0, "The casing error should still be reported")
            CollectionAssert.Contains(wa2.Entries(0).ToIniSection().Keys.Select(Function(k) k.ToString()).ToList(), "langsecref=3021")

        Finally

            winapp2ool.SettingsHandler.SettingsFile = savedFile

        End Try

    End Sub

    ''' <summary>
    ''' Gives Scan Settings each input in turn against an empty settings file, then restores the
    ''' rules to their defaults and reloads them from that file, as a restart does
    ''' </summary>
    ''' <param name="inputs">What the user types into Scan Settings, one line each</param>
    Private Shared Sub RunScanSettingsThenRestart(ParamArray inputs As String())

        Dim savedFile = winapp2ool.SettingsHandler.SettingsFile

        Try

            winapp2ool.SettingsHandler.SettingsFile = MakeIni("[WinappDebug]" & vbCrLf)

            For Each line In inputs
                winapp2ool.buildAdvSettingsMenu().HandleInput(line)
            Next

            winapp2ool.resetScanSettings()
            winapp2ool.LoadLintRulesFromSettings()

        Finally

            winapp2ool.SettingsHandler.SettingsFile = savedFile

        End Try

    End Sub

    ''' <summary>Turning a scan off also turns its repair off, and both survive a restart</summary>
    <TestMethod()> Public Sub ScanTurnedOff_RepairStaysOffAfterRestart()

        Dim rule = winapp2ool.WinappDebug.Rules(0)

        RunScanSettingsThenRestart("1")

        Assert.IsFalse(rule.ShouldScan, $"{rule.LintName}'s scan came back on")
        Assert.IsFalse(rule.ShouldRepair, $"{rule.LintName}'s repair came back on")

    End Sub

    ''' <summary>Turning a repair on also turns its scan on, and both survive a restart</summary>
    <TestMethod()> Public Sub RepairTurnedOn_ScanStaysOnAfterRestart()

        Dim rule = winapp2ool.WinappDebug.Rules(0)
        Dim count = winapp2ool.WinappDebug.Rules.Count

        ' Scan off (which turns the repair off too), then the repair back on
        RunScanSettingsThenRestart("1", CStr(count + 1))

        Assert.IsTrue(rule.ShouldRepair, $"{rule.LintName}'s repair came back off")
        Assert.IsTrue(rule.ShouldScan, $"{rule.LintName}'s scan came back off")

    End Sub

    ''' <summary>The hidden <c> alloff </c> command survives a restart</summary>
    <TestMethod()> Public Sub AllOff_SurvivesRestart()

        RunScanSettingsThenRestart("alloff")

        Dim stillOn = winapp2ool.WinappDebug.Rules.Where(Function(r) r.ShouldScan OrElse r.ShouldRepair).Select(Function(r) r.LintName)
        Assert.IsFalse(stillOn.Any(), "Came back on: " & String.Join(", ", stillOn))

    End Sub

    ''' <summary>Scan Settings' Reset option survives a restart</summary>
    <TestMethod()> Public Sub ScanSettingsReset_SurvivesRestart()

        Dim rule = winapp2ool.WinappDebug.Rules(0)
        Dim count = winapp2ool.WinappDebug.Rules.Count

        ' Reset is the option after every scan and repair toggle
        RunScanSettingsThenRestart("1", CStr(2 * count + 1))

        Assert.IsTrue(rule.ShouldScan AndAlso rule.ShouldRepair, $"{rule.LintName} came back as it was before the reset")

    End Sub

    ''' <summary>
    ''' A new <c> LintOptions </c> holds the defaults whatever the user's toggles are, and
    ''' changing its rules leaves the user's toggles alone
    ''' </summary>
    <TestMethod()> Public Sub NewOptions_AreDefaultsAndIndependent()

        winapp2ool.WinappDebug.Rules.ForEach(Sub(rule) rule.turnOff())

        Dim options As New winapp2ool.LintOptions

        Assert.IsTrue(options.RepairAll)
        Assert.IsTrue(options.Casing.ShouldScan AndAlso options.Casing.ShouldRepair)
        Assert.IsFalse(options.Opti.ShouldScan OrElse options.Opti.ShouldRepair)

        options.Opti.turnOn()

        Assert.IsFalse(winapp2ool.lintOpti.ShouldScan, "Turning on a copy changed the user's toggle")
        Assert.IsTrue(winapp2ool.WinappDebug.Rules.All(Function(rule) Not rule.ShouldScan AndAlso Not rule.ShouldRepair))

    End Sub

    ''' <summary>
    ''' The repair flags follow the formula the Scan Settings menu used: every repair runs until
    ''' a changed rule has its repair off, and rules with their repair on keep repairing after that
    ''' </summary>
    <TestMethod()> Public Sub FromSettings_DerivesRepairFlags()

        Dim defaults = winapp2ool.LintOptions.FromSettings()
        Assert.IsTrue(defaults.RepairAll)
        Assert.IsTrue(defaults.Opti.fixFormat(), "RepairAll covers rules whose own repair is off")

        winapp2ool.WinappDebug.Rules(0).ShouldRepair = False

        Dim oneOff = winapp2ool.LintOptions.FromSettings()
        Assert.IsFalse(oneOff.RepairAll)
        Assert.IsTrue(oneOff.RepairSome)
        Assert.IsFalse(oneOff.Casing.fixFormat())
        Assert.IsTrue(oneOff.Alpha.fixFormat())

        winapp2ool.WinappDebug.Rules.ForEach(Sub(rule) rule.turnOff())

        Dim allOff = winapp2ool.LintOptions.FromSettings()
        Assert.IsFalse(allOff.RepairAll OrElse allOff.RepairSome)
        Assert.IsTrue(allOff.Rules.All(Function(rule) Not rule.fixFormat()))

    End Sub

End Class
