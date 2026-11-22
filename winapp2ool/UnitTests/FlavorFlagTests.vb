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

Imports winapp2ool

''' <summary>Tests for which flavor each command line flavor flag selects, and that every one leaves <c> cmdargs </c></summary>
<TestClass()> Public Class FlavorFlagTests

    Private _savedArgs As List(Of String)
    Private _savedFlavor As WinappFlavor
    Private _savedFlag As WinappFlavor?

    ' Written out by hand rather than read from WinappFlavorInfo, so a flag dropped from the table fails here
    Private Shared ReadOnly ExpectedFlags As New Dictionary(Of String, WinappFlavor) From {
        {"-base", WinappFlavor.Base}, {"-ncc", WinappFlavor.Base},
        {"-ccleaner", WinappFlavor.CCleaner}, {"-cc", WinappFlavor.CCleaner},
        {"-ccleaner7", WinappFlavor.CCleaner7}, {"-cc7", WinappFlavor.CCleaner7},
        {"-bleachbit", WinappFlavor.BleachBit}, {"-bb", WinappFlavor.BleachBit},
        {"-fluentcleaner", WinappFlavor.FluentCleaner}, {"-fc", WinappFlavor.FluentCleaner},
        {"-systemninja", WinappFlavor.SystemNinja}, {"-sn", WinappFlavor.SystemNinja},
        {"-tron", WinappFlavor.Tron}
    }

    <TestInitialize()> Public Sub SaveState()

        _savedArgs = commandLineHandler.cmdargs
        _savedFlavor = CurrentWinappFlavor
        _savedFlag = commandLineHandler.FlavorFromCommandLine

    End Sub

    <TestCleanup()> Public Sub RestoreState()

        commandLineHandler.cmdargs = _savedArgs
        CurrentWinappFlavor = _savedFlavor
        commandLineHandler.FlavorFromCommandLine = _savedFlag

    End Sub

    ''' <summary>Runs <c> processFlavorArgs </c> on <paramref name="args"/> and returns the flavor it picked</summary>
    ''' <param name="args">The command line, without the executable's path</param>
    Private Shared Function RunFlavorArgs(ParamArray args As String()) As WinappFlavor?

        commandLineHandler.FlavorFromCommandLine = Nothing
        commandLineHandler.cmdargs = args.ToList()
        commandLineHandler.processFlavorArgs()
        Return commandLineHandler.FlavorFromCommandLine

    End Function

    ''' <summary>Each flag, alias included, selects its flavor and is removed from <c> cmdargs </c></summary>
    <TestMethod()> Public Sub EachFlag_SelectsItsFlavor()

        For Each pair In ExpectedFlags

            Assert.AreEqual(pair.Value, RunFlavorArgs(pair.Key), pair.Key)
            Assert.AreEqual(pair.Value, CurrentWinappFlavor, pair.Key)
            Assert.AreEqual(0, commandLineHandler.cmdargs.Count, $"{pair.Key} was left in cmdargs")

        Next

    End Sub

    ''' <summary>No flavor flag is left behind when several are given, and other args stay</summary>
    <TestMethod()> Public Sub AllFlags_AreRemoved()

        RunFlavorArgs(ExpectedFlags.Keys.Concat({"1"}).ToArray())

        CollectionAssert.AreEqual({"1"}, commandLineHandler.cmdargs)

    End Sub

    ''' <summary>With two flavor flags, the one later in the flavor table wins, whatever the command line order</summary>
    <TestMethod()> Public Sub TwoFlags_LaterInTableWins()

        Assert.AreEqual(WinappFlavor.CCleaner, RunFlavorArgs("-cc", "-base"))
        Assert.AreEqual(WinappFlavor.CCleaner, RunFlavorArgs("-base", "-cc"))

    End Sub

    ''' <summary>A command line with no flavor flag picks no flavor</summary>
    <TestMethod()> Public Sub NoFlag_PicksNone()

        Assert.IsFalse(RunFlavorArgs("1", "-c").HasValue)
        CollectionAssert.AreEqual({"1", "-c"}, commandLineHandler.cmdargs)

    End Sub

End Class
