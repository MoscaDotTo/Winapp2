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
''' Tests for the survival of the version comment, the flavor line and the two file markers across
''' the <c> iniFile </c> to <c> winapp2file </c> round trip. Diff reads the version off the
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

        Assert.AreEqual(winapp2ool.WinappFlavor.Base, winapp2ool.WinappFlavorInfo.ParseName("Base"))
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
    ''' Sections added from another file keep the line numbers they had there, so one from that
    ''' file's line 1 must not cut the flavor out of the preamble, as a Transmute add did
    ''' </summary>
    <TestMethod()> Public Sub AddedSectionFromLineOneKeepsFlavor()

        Dim baseIni = MakeIni(WithPreambleLine("; Flavor: BleachBit"))
        Dim source = MakeIni("[Added Entry *]" & vbCrLf & "LangSecRef=3021" & vbCrLf & "Detect=HKCU\Software\Added" & vbCrLf)
        baseIni.AddSection(source.First())

        Dim wa2 As New winapp2ool.winapp2file(baseIni)

        Assert.AreEqual(winapp2ool.WinappFlavor.BleachBit, wa2.Flavor)

    End Sub

    ''' <summary>
    ''' A flavor stamped on a file that had no header survives <c> ToIni </c>, though the file's
    ''' first section was on line 1, below which <c> ToIni </c> writes the flavor comment
    ''' </summary>
    <TestMethod()> Public Sub StampedFlavorSurvivesToIniOfHeaderlessFile()

        Dim first As New winapp2ool.winapp2file(MakeIni(VersionedFile.Substring(VersionedFile.IndexOf("[", StringComparison.Ordinal))))
        first.Flavor = winapp2ool.WinappFlavor.Tron

        Dim second As New winapp2ool.winapp2file(first.ToIni())

        Assert.AreEqual(winapp2ool.WinappFlavor.Tron, second.Flavor)

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
    ''' The archive marker survives repeated round trips, in any case, and gives the archive's
    ''' header: the marker as its third line, how to add the archive, and its own links
    ''' </summary>
    <TestMethod()> Public Sub ArchiveMarkerSurvivesRepeatedRoundTrips()

        Dim first As New winapp2ool.winapp2file(MakeIni(WithPreambleLine("; archive: Retired winapp2.ini entries")))
        Assert.AreEqual(winapp2ool.WinappFileKind.Archive, first.Kind)

        Dim third As New winapp2ool.winapp2file(New winapp2ool.winapp2file(first.ToIni()).ToIni())
        Assert.AreEqual(winapp2ool.WinappFileKind.Archive, third.Kind, "ToIni dropped the archive marker")
        Assert.IsFalse(third.Flavor.HasValue)

        Dim text = third.ToWinapp2String()
        Dim lines = text.Split({vbCrLf}, StringSplitOptions.None)
        Assert.AreEqual(winapp2ool.winapp2file.ArchiveMarkerLine, lines(2))
        StringAssert.Contains(text, ";   winapp2ool -transmute -1f winapp2.ini -2f archived.ini -3f winapp2.ini" & vbCrLf)
        StringAssert.Contains(text, ";   https://github.com/MoscaDotTo/Winapp2/releases/latest/download/archived.ini" & vbCrLf)
        StringAssert.Contains(text, ";   https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/Sources" & vbCrLf)
        StringAssert.Contains(text, "/blob/master/CONTRIBUTING.md")
        Assert.IsFalse(text.Contains("; Flavor:"))
        Assert.IsFalse(text.Contains("every published file"))

    End Sub

    ''' <summary> A file without the marker gets none, from its header or from a round trip </summary>
    <TestMethod()> Public Sub UnmarkedFileGainsNoMarker()

        Dim wa2 As New winapp2ool.winapp2file(New winapp2ool.winapp2file(MakeIni(WithPreambleLine("; Flavor: Tron"))).ToIni())

        Assert.AreEqual(winapp2ool.WinappFileKind.None, wa2.Kind)
        Assert.IsFalse(wa2.ToWinapp2String().Contains("; Archive:"))
        Assert.IsFalse(wa2.ToWinapp2String().Contains("; Winapp3:"))

    End Sub

    ''' <summary>
    ''' The archive's header keeps the version and the entry count as its first two lines
    ''' </summary>
    <TestMethod()> Public Sub ArchiveHeaderStartsWithVersionAndCount()

        Dim wa2 As New winapp2ool.winapp2file(MakeIni(VersionedFile))
        wa2.Kind = winapp2ool.WinappFileKind.Archive

        Dim lines = wa2.ToWinapp2String().Split({vbCrLf}, StringSplitOptions.None)

        Assert.AreEqual("; Version: 260219", lines(0))
        Assert.AreEqual("; # of entries: 2", lines(1))
        Assert.AreEqual(winapp2ool.winapp2file.ArchiveMarkerLine, lines(2))

    End Sub

    ''' <summary>
    ''' The winapp3.ini marker survives repeated round trips, in any case, and gives winapp3.ini's
    ''' header: the version, the count and the marker, the warning, how to add the file, and its
    ''' own links
    ''' </summary>
    <TestMethod()> Public Sub Winapp3MarkerSurvivesRepeatedRoundTrips()

        Dim first As New winapp2ool.winapp2file(MakeIni(WithPreambleLine("; winapp3: Aggressive entries for power users")))
        Assert.AreEqual(winapp2ool.WinappFileKind.Winapp3, first.Kind)

        Dim third As New winapp2ool.winapp2file(New winapp2ool.winapp2file(first.ToIni()).ToIni())
        Assert.AreEqual(winapp2ool.WinappFileKind.Winapp3, third.Kind, "ToIni dropped the winapp3.ini marker")
        Assert.IsFalse(third.Flavor.HasValue)

        Dim expected = String.Join(vbCrLf,
            "; Version: 260219",
            "; # of entries: 2",
            "; Winapp3: aggressive entries for power users",
            ";",
            "; These entries are more aggressive than the ones in winapp2.ini",
            "; Some may delete files their software or your system needs, and some take a very long time to scan",
            "; Back up your system first, and only enable an entry when you understand what it removes",
            ";",
            "; To add them to your winapp2.ini:",
            ";   winapp2ool -transmute -1f winapp2.ini -2f winapp3.ini -3f winapp2.ini",
            ";",
            "; Latest copy of this file:",
            ";   https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp3.ini",
            "; Changelog for the latest release:",
            ";   https://github.com/MoscaDotTo/Winapp2/releases/latest/download/changelog-winapp3.txt",
            "; Each changelog compares one release with the one before it",
            "; Source folder:",
            ";   https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/Sources/Winapp3",
            ";",
            "; About winapp2.ini, and how to install it for your cleaner:",
            "")

        Dim text = third.ToWinapp2String()
        Assert.IsTrue(text.StartsWith(expected, StringComparison.Ordinal), text)
        StringAssert.Contains(text, "/blob/master/CONTRIBUTING.md")
        Assert.IsFalse(text.Contains("; Flavor:"))
        Assert.IsFalse(text.Contains("; Archive:"))
        Assert.IsFalse(text.Contains("every published file"))

    End Sub

    ''' <summary>
    ''' A file carrying both marker lines is flagged and refused, whichever comes first, and a
    ''' file with one marker and no flavor isn't
    ''' </summary>
    <TestMethod()> Public Sub BothMarkers_AreAConflict()

        Dim lines = {winapp2ool.winapp2file.ArchiveMarkerLine, winapp2ool.winapp2file.Winapp3MarkerLine}

        For Each pair In {lines, lines.Reverse().ToArray()}

            Dim wa2 As New winapp2ool.winapp2file(MakeIni(WithPreambleLine(pair(0) & vbCrLf & pair(1))))
            Assert.IsTrue(wa2.HasConflictingMarkers)
            Assert.AreNotEqual("", wa2.MarkerConflict())

        Next

        For Each line In lines

            Dim wa2 As New winapp2ool.winapp2file(MakeIni(WithPreambleLine(line)))
            Assert.IsFalse(wa2.HasConflictingMarkers)
            Assert.AreEqual("", wa2.MarkerConflict())

        Next

    End Sub

    ''' <summary>
    ''' A marked file with a flavor, from its header or set afterwards, is refused
    ''' </summary>
    <TestMethod()> Public Sub MarkedFileWithFlavor_IsAConflict()

        Dim fromHeader As New winapp2ool.winapp2file(MakeIni(WithPreambleLine("; Flavor: Tron" & vbCrLf & winapp2ool.winapp2file.Winapp3MarkerLine)))
        Assert.AreNotEqual("", fromHeader.MarkerConflict())

        Dim setLater As New winapp2ool.winapp2file(MakeIni(VersionedFile))
        setLater.Kind = winapp2ool.WinappFileKind.Winapp3
        setLater.Flavor = winapp2ool.WinappFlavor.Base
        Assert.AreNotEqual("", setLater.MarkerConflict())

    End Sub

    ''' <summary>
    ''' A file with no flavor and no marker links the latest release under the label that names
    ''' no product
    ''' </summary>
    <TestMethod()> Public Sub UnflavoredUnmarkedFile_LinksTheLatestReleaseNeutrally()

        Dim text = New winapp2ool.winapp2file(MakeIni(VersionedFile)).ToWinapp2String()

        StringAssert.Contains(text, "; # of entries: 2" & vbCrLf & ";" & vbCrLf &
                                    "; Latest copy of every published file:" & vbCrLf &
                                    ";   https://github.com/MoscaDotTo/Winapp2/releases/latest" & vbCrLf & ";" & vbCrLf)
        Assert.IsFalse(text.Contains("in every flavor"))

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
