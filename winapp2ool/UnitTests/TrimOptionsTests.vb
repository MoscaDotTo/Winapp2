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
''' Tests that <see cref="winapp2ool.Trim.trimFile"/> takes its include and exclude lists from
''' the <see cref="winapp2ool.TrimOptions"/> it is given, never from the Trim module's saved
''' settings, so another module's trim can't pick up the user's Trim lists
''' </summary>
<TestClass()> Public Class TrimOptionsTests

    ''' <summary>
    ''' An entry whose only detection is a registry key no machine has, so detection alone
    ''' always discards it
    ''' </summary>
    Private Const UndetectedEntry As String = "[Ghost *]" & vbCrLf &
                                              "LangSecRef=3021" & vbCrLf &
                                              "Detect=HKCU\Software\winapp2ool-unit-test-ghost-entry" & vbCrLf &
                                              "FileKey1=%AppData%\winapp2ool-unit-test-ghost|*.*" & vbCrLf

    ''' <summary>
    ''' An entry with no detection keys, which detection alone always keeps
    ''' </summary>
    Private Const UndetectableEntry As String = "[Always *]" & vbCrLf &
                                                "LangSecRef=3021" & vbCrLf &
                                                "FileKey1=%AppData%\winapp2ool-unit-test-always|*.*" & vbCrLf

    Private _dir As String

    ''' <summary>Creates a scratch folder holding an includes and an excludes file</summary>
    <TestInitialize()> Public Sub MakeListFiles()

        _dir = Path.Combine(Path.GetTempPath(), $"winapp2ool-trimoptions-{Guid.NewGuid():N}")
        Directory.CreateDirectory(_dir)
        File.WriteAllText(Path.Combine(_dir, "includes.ini"), "[Ghost *]" & vbCrLf)
        File.WriteAllText(Path.Combine(_dir, "excludes.ini"), "[Always *]" & vbCrLf)

    End Sub

    ''' <summary>Deletes the scratch folder</summary>
    <TestCleanup()> Public Sub DeleteListFiles()

        Directory.Delete(_dir, recursive:=True)

    End Sub

    ''' <summary>Trims <paramref name="text"/> with <paramref name="options"/> and returns the surviving entry names</summary>
    ''' <param name="text">Literal winapp2.ini text</param>
    ''' <param name="options">The lists to trim with, or <c> Nothing </c> for none</param>
    Private Shared Function TrimNames(text As String,
                                      options As winapp2ool.TrimOptions) As List(Of String)

        Dim ini As winapp2ool.iniFile

        Using reader As New StreamReader(New MemoryStream(Encoding.UTF8.GetBytes(text)))
            ini = winapp2ool.iniFile.FromStream(reader, "", "winapp2.ini")
        End Using

        Dim wa2 As New winapp2ool.winapp2file(ini)
        winapp2ool.Trim.trimFile(wa2, options)

        Return wa2.Entries.Select(Function(e) e.Name).ToList()

    End Function

    ''' <summary>Returns a chooser for <paramref name="name"/> in the scratch folder</summary>
    ''' <param name="name">The file name</param>
    Private Function Chooser(name As String) As winapp2ool.iniFileChooser

        Return New winapp2ool.iniFileChooser(_dir, name, mustExist:=False)

    End Function

    ''' <summary>
    ''' With Trim's saved lists switched on, a trim given no options still uses no lists: the
    ''' included entry is discarded by detection and the excluded one kept
    ''' </summary>
    <TestMethod()> Public Sub NoOptions_IgnoresTrimsSavedLists()

        Dim savedIncludes = winapp2ool.trimsettings.UseTrimIncludes
        Dim savedExcludes = winapp2ool.trimsettings.UseTrimExcludes
        Dim savedFile2 = winapp2ool.trimsettings.TrimFile2
        Dim savedFile4 = winapp2ool.trimsettings.TrimFile4

        Try

            winapp2ool.trimsettings.UseTrimIncludes = True
            winapp2ool.trimsettings.UseTrimExcludes = True
            winapp2ool.trimsettings.TrimFile2 = Chooser("includes.ini")
            winapp2ool.trimsettings.TrimFile4 = Chooser("excludes.ini")

            Dim names = TrimNames(UndetectedEntry & UndetectableEntry, Nothing)

            CollectionAssert.AreEqual({"Always *"}, names)

        Finally

            winapp2ool.trimsettings.UseTrimIncludes = savedIncludes
            winapp2ool.trimsettings.UseTrimExcludes = savedExcludes
            winapp2ool.trimsettings.TrimFile2 = savedFile2
            winapp2ool.trimsettings.TrimFile4 = savedFile4

        End Try

    End Sub

    ''' <summary>
    ''' Lists given in the options apply while Trim's own lists are off, which is how
    ''' CC7Patcher trims with its own: the included entry survives and the excluded one goes
    ''' </summary>
    <TestMethod()> Public Sub GivenLists_ApplyWithTrimsListsOff()

        Assert.IsFalse(winapp2ool.trimsettings.UseTrimIncludes)
        Assert.IsFalse(winapp2ool.trimsettings.UseTrimExcludes)

        Dim options As New winapp2ool.TrimOptions With {
            .IncludesFile = Chooser("includes.ini"),
            .ExcludesFile = Chooser("excludes.ini")
        }

        Dim names = TrimNames(UndetectedEntry & UndetectableEntry, options)

        CollectionAssert.AreEqual({"Ghost *"}, names)

    End Sub

    ''' <summary>
    ''' <see cref="winapp2ool.TrimOptions.FromSettings"/> carries Trim's lists only while their
    ''' toggles are on
    ''' </summary>
    <TestMethod()> Public Sub FromSettings_FollowsTheToggles()

        Dim savedIncludes = winapp2ool.trimsettings.UseTrimIncludes
        Dim savedExcludes = winapp2ool.trimsettings.UseTrimExcludes

        Try

            winapp2ool.trimsettings.UseTrimIncludes = True
            winapp2ool.trimsettings.UseTrimExcludes = False

            Dim options = winapp2ool.TrimOptions.FromSettings()

            Assert.AreSame(winapp2ool.trimsettings.TrimFile2, options.IncludesFile)
            Assert.IsNull(options.ExcludesFile)

        Finally

            winapp2ool.trimsettings.UseTrimIncludes = savedIncludes
            winapp2ool.trimsettings.UseTrimExcludes = savedExcludes

        End Try

    End Sub

End Class
