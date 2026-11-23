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
''' Ties <see cref="winapp2ool.WinappFlavorInfo"/>'s table to <c> .github/scripts/flavors.csv </c>,
''' the list the build scripts publish from. The project links the CSV into the test output folder.
''' </summary>
<TestClass()> Public Class FlavorManifestTests

    Private Const ArchiveKey As String = "Archive"

    Private Const Winapp3Key As String = "Winapp3"

    ''' <summary>The keys of the rows that publish a file that isn't a flavor</summary>
    Private Shared ReadOnly NonFlavorKeys As New HashSet(Of String) From {ArchiveKey, Winapp3Key}

    ''' <summary>Reads the manifest's rows as column name to value, header row excluded</summary>
    Private Shared Function ReadManifest() As List(Of Dictionary(Of String, String))

        Dim manifestPath = Path.Combine(Path.GetDirectoryName(GetType(FlavorManifestTests).Assembly.Location), "flavors.csv")
        Assert.IsTrue(File.Exists(manifestPath), $"The linked manifest is missing from the test output: {manifestPath}")

        Dim lines = File.ReadAllLines(manifestPath).Where(Function(line) line.Length > 0).ToList()
        Dim header = lines(0).Split(","c)
        Dim rows As New List(Of Dictionary(Of String, String))

        For Each line In lines.Skip(1)

            Dim cells = line.Split(","c)
            Assert.AreEqual(header.Length, cells.Length, $"Wrong number of cells: {line}")

            Dim row As New Dictionary(Of String, String)
            For i = 0 To header.Length - 1
                row(header(i)) = cells(i)
            Next
            rows.Add(row)

        Next

        Return rows

    End Function

    ''' <summary>Each flavor row names a flavor and carries the asset and changelog names the tool downloads</summary>
    <TestMethod()> Public Sub FlavorRows_MatchTheTable()

        For Each row In ReadManifest().Where(Function(r) Not NonFlavorKeys.Contains(r("Key")))

            Dim flavor As winapp2ool.WinappFlavor
            Assert.IsTrue([Enum].TryParse(row("Key"), flavor) AndAlso [Enum].IsDefined(GetType(winapp2ool.WinappFlavor), flavor) AndAlso
                          flavor.ToString() = row("Key"), $"'{row("Key")}' is not a WinappFlavor member")

            Dim info = winapp2ool.WinappFlavorInfo.Lookup(flavor)
            Assert.AreEqual(info.AssetName, row("Asset"), $"{flavor}'s asset")
            Assert.AreEqual(info.ChangelogAssetName, row("Changelog"), $"{flavor}'s changelog")

        Next

    End Sub

    ''' <summary>Every flavor has exactly one row in the manifest, and every flavor is in the table</summary>
    <TestMethod()> Public Sub EveryFlavor_HasOneRow()

        Dim keys = ReadManifest().Select(Function(r) r("Key")).ToList()

        For Each flavor In [Enum].GetValues(GetType(winapp2ool.WinappFlavor)).Cast(Of winapp2ool.WinappFlavor)()

            Assert.AreEqual(1, keys.Where(Function(k) k = flavor.ToString()).Count(), $"Rows for {flavor}")
            Assert.AreEqual(1, winapp2ool.WinappFlavorInfo.All.Where(Function(info) info.Flavor = flavor).Count(), $"Table entries for {flavor}")

        Next

    End Sub

    ''' <summary>The archive's row publishes the file the tool downloads as the archive</summary>
    <TestMethod()> Public Sub ArchiveRow_MatchesArchiveFileName()

        Dim archiveRows = ReadManifest().Where(Function(r) r("Key") = ArchiveKey).ToList()

        Assert.AreEqual(1, archiveRows.Count)
        Assert.AreEqual(winapp2ool.winapp2file.ArchiveFileName, archiveRows(0)("Asset"))

    End Sub

    ''' <summary>
    ''' winapp3.ini's row publishes the file and the changelog the published winapp3.ini header
    ''' links. Inconclusive until the build gains the row.
    ''' </summary>
    <TestMethod()> Public Sub Winapp3Row_MatchesTheLinkedNames()

        Dim winapp3Rows = ReadManifest().Where(Function(r) r("Key") = Winapp3Key).ToList()

        If winapp3Rows.Count = 0 Then

            Assert.Inconclusive("flavors.csv has no Winapp3 row yet")

        End If

        Assert.AreEqual(1, winapp3Rows.Count)
        Assert.AreEqual(winapp2ool.winapp2file.Winapp3FileName, winapp3Rows(0)("Asset"))
        Assert.AreEqual(winapp2ool.winapp2file.Winapp3ChangelogFileName, winapp3Rows(0)("Changelog"))

    End Sub

End Class
