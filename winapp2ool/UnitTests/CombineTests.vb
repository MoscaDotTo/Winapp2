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
''' Tests for how Combine names its input files in its output
''' </summary>
<TestClass()> Public Class CombineTests

    ''' <summary>
    ''' A file is named by its path below the target directory, so a top-level <c> A.ini </c>
    ''' and <c> UWP\A.ini </c> read differently, with or without a trailing slash on the target
    ''' </summary>
    <TestMethod()> Public Sub DisplayName_IsRelativeToTarget()

        Assert.AreEqual("A.ini", winapp2ool.Combine.displayName("C:\Assembler\Entries", "C:\Assembler\Entries\A.ini"))
        Assert.AreEqual("UWP\A.ini", winapp2ool.Combine.displayName("C:\Assembler\Entries", "C:\Assembler\Entries\UWP\A.ini"))
        Assert.AreEqual("UWP\A.ini", winapp2ool.Combine.displayName("C:\Assembler\Entries\", "C:\Assembler\Entries\UWP\A.ini"))
        Assert.AreEqual("UWP\A.ini", winapp2ool.Combine.displayName("c:\assembler\entries", "C:\Assembler\Entries\UWP\A.ini"))

    End Sub

    ''' <summary>
    ''' A path that isn't under the target directory, including a sibling whose name starts with
    ''' the target's, falls back to the file name alone
    ''' </summary>
    <TestMethod()> Public Sub DisplayName_OutsideTargetGivesFileName()

        Assert.AreEqual("A.ini", winapp2ool.Combine.displayName("C:\Assembler\Entries", "C:\Elsewhere\A.ini"))
        Assert.AreEqual("A.ini", winapp2ool.Combine.displayName("C:\Assembler\Entries", "C:\Assembler\Entries2\A.ini"))

    End Sub

    ''' <summary>
    ''' The paths <c> Directory.GetFiles </c> returns for a real directory start with the target
    ''' as given, so every file found, nested or not, gets its relative name
    ''' </summary>
    <TestMethod()> Public Sub DisplayName_MatchesDirectoryGetFiles()

        Dim root = Path.Combine(Path.GetTempPath(), $"CombineTests_{Guid.NewGuid():N}")
        Directory.CreateDirectory(Path.Combine(root, "UWP"))

        Try

            File.WriteAllText(Path.Combine(root, "A.ini"), "")
            File.WriteAllText(Path.Combine(root, "UWP", "A.ini"), "")

            For Each target In {root, root & "\"}

                Dim names = Directory.GetFiles(target, "*.ini", SearchOption.AllDirectories).
                    Select(Function(f) winapp2ool.Combine.displayName(target, f)).
                    OrderBy(Function(n) n, StringComparer.Ordinal).ToArray()

                CollectionAssert.AreEqual({"A.ini", "UWP\A.ini"}, names, $"target: {target}")

            Next

        Finally

            Directory.Delete(root, recursive:=True)

        End Try

    End Sub

End Class
