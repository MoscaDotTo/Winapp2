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
''' Tests for the version comment's survival across the <c> iniFile2 </c> to
''' <c> winapp2file2 </c> round trip. Trim, Transmute, WinappDebug, and CC7Patcher all
''' pass files back through <c> ToIni </c>, and Diff reads the version off the resulting
''' file's first comment to build its header
''' </summary>
<TestClass()> Public Class Winapp2FileVersionTests

    ''' <summary>
    ''' Helper: parse an <c> iniFile2 </c> from literal ini text
    ''' </summary>
    Private Shared Function MakeIni(text As String) As winapp2ool.iniFile2

        Dim bytes = Encoding.UTF8.GetBytes(text)
        Using ms As New IO.MemoryStream(bytes)
            Using reader As New IO.StreamReader(ms)
                Return winapp2ool.iniFile2.FromStream(reader, "", "winapp2.ini")
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
    ''' A file's version comment must still be readable off the <c> iniFile2 </c> produced
    ''' by <c> ToIni </c>, which is the form Diff receives when it trims the remote file
    ''' </summary>
    <TestMethod()> Public Sub ToIniPreservesVersionComment()

        Dim wa2 As New winapp2ool.winapp2file2(MakeIni(VersionedFile))
        Dim roundTripped = wa2.ToIni()

        Assert.AreNotEqual(0, roundTripped.Comments.Count, "ToIni dropped the version comment entirely")
        Assert.AreEqual("; Version: 260219", roundTripped.Comments(0).Text)

    End Sub

    ''' <summary>
    ''' The round trip must be idempotent: rebuilding a <c> winapp2file2 </c> from a
    ''' <c> ToIni </c> result must find the same version rather than falling back to 000000
    ''' </summary>
    <TestMethod()> Public Sub VersionSurvivesRepeatedRoundTrips()

        Dim first As New winapp2ool.winapp2file2(MakeIni(VersionedFile))
        Dim second As New winapp2ool.winapp2file2(first.ToIni())
        Dim third As New winapp2ool.winapp2file2(second.ToIni())

        Assert.AreEqual("; Version: 260219", third.Version)

    End Sub

    ''' <summary>
    ''' A file with no version comment still round trips, carrying the 000000 placeholder
    ''' rather than producing a file with no comments at all
    ''' </summary>
    <TestMethod()> Public Sub UnversionedFileCarriesPlaceholder()

        Dim noVersion = VersionedFile.Replace("; Version: 260219" & vbCrLf, "")
        Dim wa2 As New winapp2ool.winapp2file2(MakeIni(noVersion))

        Assert.AreEqual("; Version: 000000", wa2.ToIni().Comments(0).Text)

    End Sub

    ''' <summary>
    ''' The non-CCleaner marker must survive the round trip too — it selects the header text
    ''' and license link written by <c> ToWinapp2String </c>, so losing it mislabels the file
    ''' </summary>
    <TestMethod()> Public Sub ToIniPreservesNCCMarker()

        Dim nccText = VersionedFile.Replace("; # of entries: 2" & vbCrLf,
            "; # of entries: 2" & vbCrLf &
            "; This is the non-CCleaner version of Winapp2 that contains extra entries that were removed due to them being added to CCleaner." & vbCrLf)

        Dim first As New winapp2ool.winapp2file2(MakeIni(nccText))
        Assert.IsTrue(first.IsNCC, "the fixture was not recognized as non-CCleaner")

        Dim second As New winapp2ool.winapp2file2(first.ToIni())
        Assert.IsTrue(second.IsNCC, "ToIni dropped the non-CCleaner marker")
        Assert.IsTrue(second.ToWinapp2String().Contains("Winapp2 (Non-CCleaner version)"))

    End Sub

    ''' <summary>
    ''' A CCleaner-variant file must not pick up the marker from its own round trip
    ''' </summary>
    <TestMethod()> Public Sub RoundTripDoesNotInventNCCMarker()

        Dim first As New winapp2ool.winapp2file2(MakeIni(VersionedFile))
        Dim second As New winapp2ool.winapp2file2(first.ToIni())

        Assert.IsFalse(first.IsNCC)
        Assert.IsFalse(second.IsNCC)
        Assert.AreEqual(1, first.ToIni().Comments.Count)

    End Sub

    ''' <summary>
    ''' Trimming every entry out of a file must not take the version with it
    ''' </summary>
    <TestMethod()> Public Sub VersionSurvivesEntryRemoval()

        Dim wa2 As New winapp2ool.winapp2file2(MakeIni(VersionedFile))

        For Each entry In wa2.Entries.ToList()
            wa2.RemoveEntry(entry)
        Next

        Assert.AreEqual(0, wa2.Count)
        Assert.AreEqual("; Version: 260219", wa2.ToIni().Comments(0).Text)

    End Sub

End Class
