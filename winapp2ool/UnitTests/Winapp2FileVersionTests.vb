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

Imports System.Text

''' <summary>
''' Tests for the survival of the version comment and the flavor line across the
''' <c> iniFile </c> to <c> winapp2file </c> round trip. Diff reads the version off the
''' first comment of the file <see cref="winapp2ool.winapp2file.ToIni"/> returns to build its header
''' </summary>
<TestClass()> Public Class Winapp2FileVersionTests

    ''' <summary>Returns an <c> iniFile </c> named <c> winapp2.ini </c> parsed from <paramref name="text"/></summary>
    ''' <param name="text">Literal ini text</param>
    Private Shared Function MakeIni(text As String) As winapp2ool.iniFile

        Dim bytes = Encoding.UTF8.GetBytes(text)
        Using ms As New IO.MemoryStream(bytes)
            Using reader As New IO.StreamReader(ms)
                Return winapp2ool.iniFile.FromStream(reader, "", "winapp2.ini")
            End Using
        End Using

    End Function

    Private Const VersionedFile As String =
        "; Version: 260219" & vbCrLf &
        "; # of entries: 2" & vbCrLf &
        vbCrLf &
        "[Test Entry A *]" & vbCrLf &
        "LangSecRef=3021" & vbCrLf &
        "Detect=HKCU\Software\TestA" & vbCrLf &
        "FileKey1=%AppData%\TestA|*.log" & vbCrLf &
        vbCrLf &
        "[Test Entry B *]" & vbCrLf &
        "LangSecRef=3021" & vbCrLf &
        "Detect=HKCU\Software\TestB" & vbCrLf &
        "FileKey1=%AppData%\TestB|*.log" & vbCrLf

    ''' <summary>
    ''' A file's version comment must still be readable off the <c> iniFile </c> produced
    ''' by <c> ToIni </c>, which is the form Diff receives when it trims the remote file
    ''' </summary>
    <TestMethod()> Public Sub ToIniPreservesVersionComment()

        Dim wa2 As New winapp2ool.winapp2file(MakeIni(VersionedFile))
        Dim roundTripped = wa2.ToIni()

        Assert.AreNotEqual(0, roundTripped.Comments.Count, "ToIni dropped the version comment entirely")
        Assert.AreEqual("; Version: 260219", roundTripped.Comments(0).Text)

    End Sub

    ''' <summary>
    ''' The round trip must be idempotent: rebuilding a <c> winapp2file </c> from a
    ''' <c> ToIni </c> result must find the same version rather than falling back to 000000
    ''' </summary>
    <TestMethod()> Public Sub VersionSurvivesRepeatedRoundTrips()

        Dim first As New winapp2ool.winapp2file(MakeIni(VersionedFile))
        Dim second As New winapp2ool.winapp2file(first.ToIni())
        Dim third As New winapp2ool.winapp2file(second.ToIni())

        Assert.AreEqual("; Version: 260219", third.Version)

    End Sub

    ''' <summary>
    ''' A file with no version comment still round trips, carrying the 000000 placeholder
    ''' rather than producing a file with no comments at all
    ''' </summary>
    <TestMethod()> Public Sub UnversionedFileCarriesPlaceholder()

        Dim noVersion = VersionedFile.Replace("; Version: 260219" & vbCrLf, "")
        Dim wa2 As New winapp2ool.winapp2file(MakeIni(noVersion))

        Assert.AreEqual("; Version: 000000", wa2.ToIni().Comments(0).Text)

    End Sub

    ''' <summary>Returns <see cref="VersionedFile"/> with <paramref name="line"/> after its entry count</summary>
    ''' <param name="line">A preamble comment line, without its line break</param>
    Private Shared Function WithPreambleLine(line As String) As String

        Return VersionedFile.Replace("; # of entries: 2" & vbCrLf, "; # of entries: 2" & vbCrLf & line & vbCrLf)

    End Function

    ''' <summary>
    ''' The flavor must survive the round trip too, since it selects the header text and links
    ''' <c> ToWinapp2String </c> writes, so losing it unlabels the file
    ''' </summary>
    <TestMethod()> Public Sub FlavorSurvivesRepeatedRoundTrips()

        Dim first As New winapp2ool.winapp2file(MakeIni(WithPreambleLine("; Flavor: BleachBit")))
        Assert.AreEqual(winapp2ool.WinappFlavor.BleachBit, first.Flavor)

        Dim third As New winapp2ool.winapp2file(New winapp2ool.winapp2file(first.ToIni()).ToIni())
        Assert.AreEqual(winapp2ool.WinappFlavor.BleachBit, third.Flavor, "ToIni dropped the flavor")

        Dim text = third.ToWinapp2String()
        StringAssert.Contains(text, "; Flavor: BleachBit")
        StringAssert.Contains(text, "/releases/latest/download/winapp2-bleachbit.ini")
        StringAssert.Contains(text, "/releases/latest/download/changelog-bleachbit.txt")

    End Sub

    ''' <summary>
    ''' The flavor line accepts the display name or the enum name, in any case and spacing
    ''' </summary>
    <TestMethod()> Public Sub FlavorNamesParse()

        Assert.AreEqual(winapp2ool.WinappFlavor.NonCCleaner, winapp2ool.WinappFlavorInfo.ParseName("Base"))
        Assert.AreEqual(winapp2ool.WinappFlavor.NonCCleaner, winapp2ool.WinappFlavorInfo.ParseName("NonCCleaner"))
        Assert.AreEqual(winapp2ool.WinappFlavor.SystemNinja, winapp2ool.WinappFlavorInfo.ParseName("system ninja"))
        Assert.AreEqual(winapp2ool.WinappFlavor.SystemNinja, winapp2ool.WinappFlavorInfo.ParseName("SystemNinja"))
        Assert.AreEqual(winapp2ool.WinappFlavor.CCleaner7, winapp2ool.WinappFlavorInfo.ParseName("CCleaner 7"))
        Assert.IsFalse(winapp2ool.WinappFlavorInfo.ParseName("Winapp3").HasValue)

    End Sub

    ''' <summary>
    ''' Every flavor's display name parses back to that flavor, so a new flavor can't ship a
    ''' header its own file can't read
    ''' </summary>
    <TestMethod()> Public Sub EveryFlavorHasHeaderInfo()

        For Each flavor As winapp2ool.WinappFlavor In [Enum].GetValues(GetType(winapp2ool.WinappFlavor))

            Dim info = winapp2ool.WinappFlavorInfo.Lookup(flavor)
            Assert.AreEqual(flavor, winapp2ool.WinappFlavorInfo.ParseName(info.DisplayName), flavor.ToString())
            Assert.AreNotEqual(0, info.Description.Count, flavor.ToString())

        Next

    End Sub

    ''' <summary>
    ''' An unknown flavor name is treated as no flavor, and a file without one gets the header
    ''' that names none and gains no flavor line on a round trip
    ''' </summary>
    <TestMethod()> Public Sub NoFlavorGivesTheGenericHeader()

        For Each fileText In {VersionedFile, WithPreambleLine("; Flavor: Winapp3")}

            Dim first As New winapp2ool.winapp2file(MakeIni(fileText))
            Dim second As New winapp2ool.winapp2file(first.ToIni())

            Assert.IsFalse(first.Flavor.HasValue)
            Assert.IsFalse(second.Flavor.HasValue)
            Assert.AreEqual(1, first.ToIni().Comments.Count)

            Dim header = second.ToWinapp2String()
            Assert.IsFalse(header.Contains("; Flavor:"))
            StringAssert.Contains(header, "/releases/latest" & vbCrLf)

        Next

    End Sub

    ''' <summary>
    ''' A <c> ; Flavor: </c> comment below the first entry isn't part of the preamble
    ''' </summary>
    <TestMethod()> Public Sub FlavorLineAfterFirstSectionIsIgnored()

        Dim fileText = VersionedFile.Replace("[Test Entry B *]", "; Flavor: Tron" & vbCrLf & "[Test Entry B *]")
        Dim wa2 As New winapp2ool.winapp2file(MakeIni(fileText))

        Assert.IsFalse(wa2.Flavor.HasValue)

    End Sub

    ''' <summary>
    ''' The version and the entry count stay the header's first two lines, with or without a
    ''' flavor, since the updater, Diff and the build read them there
    ''' </summary>
    <TestMethod()> Public Sub HeaderStartsWithVersionAndCount()

        For Each fileText In {VersionedFile, WithPreambleLine("; Flavor: Tron")}

            Dim lines = New winapp2ool.winapp2file(MakeIni(fileText)).ToWinapp2String().Split({vbCrLf}, StringSplitOptions.None)

            Assert.AreEqual("; Version: 260219", lines(0))
            Assert.AreEqual("; # of entries: 2", lines(1))

        Next

    End Sub

    ''' <summary>
    ''' Trimming every entry out of a file must not take the version with it
    ''' </summary>
    <TestMethod()> Public Sub VersionSurvivesEntryRemoval()

        Dim wa2 As New winapp2ool.winapp2file(MakeIni(VersionedFile))

        For Each entry In wa2.Entries.ToList()
            wa2.RemoveEntry(entry)
        Next

        Assert.AreEqual(0, wa2.Count)
        Assert.AreEqual("; Version: 260219", wa2.ToIni().Comments(0).Text)

    End Sub

End Class
