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
''' Tests for <c> SourceDirectory.Combine </c>, which merges EntryBuilder's sources and
''' UWPBuilder's <c> AppInfo\ </c>, and for the parser's <c> iniFile.DuplicateSections </c> list it
''' relies on
''' </summary>
<TestClass()> Public Class SourceDirectoryTests

    ''' <summary>
    ''' The parser keeps the first section of a name and lists a later one, in any case and with
    ''' its keys and line number, in <c> DuplicateSections </c>
    ''' </summary>
    <TestMethod()> Public Sub Parser_ListsRepeatedSections()

        Dim text = "[Alpha *]" & vbCrLf & "Package=First" & vbCrLf &
                   "[ALPHA *]" & vbCrLf & "Package=Second" & vbCrLf

        Using reader As New IO.StreamReader(New IO.MemoryStream(Encoding.UTF8.GetBytes(text)))

            Dim parsed = winapp2ool.iniFile.FromStream(reader)

            Assert.AreEqual(1, parsed.Count)
            Assert.AreEqual("First", parsed.GetSection("alpha *").Keys.GetKey("Package").Value)
            Assert.AreEqual(1, parsed.DuplicateSections.Count)
            Assert.AreEqual("ALPHA *", parsed.DuplicateSections(0).Name)
            Assert.AreEqual(3, parsed.DuplicateSections(0).StartingLineNumber)
            Assert.AreEqual("Second", parsed.DuplicateSections(0).Keys.GetKey("Package").Value)

        End Using

    End Sub

    ''' <summary>
    ''' A section repeated in the same file or a later one is dropped with a warning naming the
    ''' repeat's file and line and the file whose definition is kept, and the run fails. Files
    ''' are read in name order, the first definition wins, and names match in any case.
    ''' </summary>
    <TestMethod()> Public Sub Combine_WarnsAboutRepeatsAndKeepsTheFirst()

        Dim dir = IO.Path.Combine(IO.Path.GetTempPath(), $"sourcedir-dupes-{Guid.NewGuid():N}")
        IO.Directory.CreateDirectory(dir)
        Dim savedExitCode = Environment.ExitCode

        Try

            IO.File.WriteAllText(IO.Path.Combine(dir, "B.ini"),
                                 "[Alpha *]" & vbCrLf & "Package=B.Third" & vbCrLf &
                                 "[Beta *]" & vbCrLf & "Package=B.Beta" & vbCrLf)
            IO.File.WriteAllText(IO.Path.Combine(dir, "A.ini"),
                                 "[Alpha *]" & vbCrLf & "Package=A.First" & vbCrLf &
                                 "[alpha *]" & vbCrLf & "Package=A.Second" & vbCrLf)

            Environment.ExitCode = 0
            Dim combined As winapp2ool.iniFile
            Dim log As List(Of String)

            Using cap = winapp2ool.gLogCapture()
                combined = winapp2ool.SourceDirectory.Combine(dir, "AppInfo", New winapp2ool.MenuSection)
                log = cap.Lines.ToList()
            End Using

            Dim allLog = String.Join("; ", log)
            Assert.AreEqual(1, Environment.ExitCode, allLog)
            Assert.AreEqual(2, combined.Count)
            Assert.AreEqual("A.First", combined.GetSection("Alpha *").Keys.GetKey("Package").Value)

            Dim warnings = log.Where(Function(l) l.Contains("Duplicate AppInfo section")).ToList()
            Assert.AreEqual(2, warnings.Count, allLog)
            Assert.IsTrue(warnings.Any(Function(w) w.Contains("[alpha *] at A.ini line 3") AndAlso w.EndsWith("in A.ini", StringComparison.Ordinal)), allLog)
            Assert.IsTrue(warnings.Any(Function(w) w.Contains("[Alpha *] at B.ini line 1") AndAlso w.EndsWith("in A.ini", StringComparison.Ordinal)), allLog)

        Finally

            Environment.ExitCode = savedExitCode
            IO.Directory.Delete(dir, recursive:=True)

        End Try

    End Sub

    ''' <summary>
    ''' Files with no repeated names leave the exit code alone
    ''' </summary>
    <TestMethod()> Public Sub Combine_NoRepeats_DoesNotFail()

        Dim dir = IO.Path.Combine(IO.Path.GetTempPath(), $"sourcedir-clean-{Guid.NewGuid():N}")
        IO.Directory.CreateDirectory(dir)
        Dim savedExitCode = Environment.ExitCode

        Try

            IO.File.WriteAllText(IO.Path.Combine(dir, "A.ini"), "[Alpha *]" & vbCrLf & "Package=A" & vbCrLf)
            IO.File.WriteAllText(IO.Path.Combine(dir, "B.ini"), "[Beta *]" & vbCrLf & "Package=B" & vbCrLf)

            Environment.ExitCode = 0
            Dim combined = winapp2ool.SourceDirectory.Combine(dir, "EntryBuilder", New winapp2ool.MenuSection)

            Assert.AreEqual(2, combined.Count)
            Assert.AreEqual(0, Environment.ExitCode)

        Finally

            Environment.ExitCode = savedExitCode
            IO.Directory.Delete(dir, recursive:=True)

        End Try

    End Sub

    ''' <summary>
    ''' A missing directory yields an empty file without a warning
    ''' </summary>
    <TestMethod()> Public Sub Combine_MissingDirectory_IsEmpty()

        Dim dir = IO.Path.Combine(IO.Path.GetTempPath(), $"sourcedir-missing-{Guid.NewGuid():N}")

        Dim combined = winapp2ool.SourceDirectory.Combine(dir, "EntryBuilder", New winapp2ool.MenuSection)

        Assert.AreEqual(0, combined.Count)

    End Sub

End Class
