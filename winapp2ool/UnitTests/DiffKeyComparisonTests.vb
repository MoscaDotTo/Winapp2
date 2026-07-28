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
''' Tests for the Diff module's key-level equivalence in <c> KeyModificationAnalyzer2 </c>.
''' A FileKey's semicolon-delimited pattern list is an unordered OR-set, so a pure reordering
''' of its patterns must not be reported as a change; any genuine pattern, path, or flag
''' difference must still surface as a modification.
''' </summary>
<TestClass()> Public Class DiffKeyComparisonTests

    ''' <summary>
    ''' Helper: parse an <c> iniFile2 </c> from literal ini text
    ''' </summary>
    Private Shared Function MakeIni(text As String) As winapp2ool.iniFile2

        Dim bytes = Encoding.UTF8.GetBytes(text)
        Using ms As New IO.MemoryStream(bytes)
            Using reader As New IO.StreamReader(ms)
                Return winapp2ool.iniFile2.FromStream(reader, "", "test.ini")
            End Using
        End Using

    End Function

    ''' <summary>
    ''' Helper: build a one-entry section named <c> [App *] </c> from a list of key lines
    ''' </summary>
    Private Shared Function MakeSection(ParamArray keyLines As String()) As winapp2ool.iniSection2

        Dim text = "[App *]" & vbCrLf & String.Join(vbCrLf, keyLines) & vbCrLf
        Return MakeIni(text).GetSection("App *")

    End Function

    ''' <summary>
    ''' Helper: run key-level modification analysis on a single entry present in both versions,
    ''' returning whether the entry was recorded as modified
    ''' </summary>
    Private Shared Function EntryModified(oldSection As winapp2ool.iniSection2,
                                          newSection As winapp2ool.iniSection2) As Boolean

        Dim state As New winapp2ool.DiffState
        Dim analyzer As New winapp2ool.KeyModificationAnalyzer2(state)
        analyzer.FindModifications(oldSection, newSection)

        Return state.ModifiedEntries.ModifiedEntryNames.Contains(newSection.Name)

    End Function

    ''' <summary>
    ''' Reordering a FileKey's patterns (and renumbering the key) is not a change
    ''' </summary>
    <TestMethod()> Public Sub FileKeyPatternReorder_IsNotAChange()

        Dim oldSection = MakeSection("LangSecRef=3021",
                                     "Detect=HKCU\Software\App",
                                     "FileKey9=%ProgramFiles%\ASUS\*|*.log;*log*.txt;*.tmp|RECURSE")

        Dim newSection = MakeSection("LangSecRef=3021",
                                     "Detect=HKCU\Software\App",
                                     "FileKey12=%ProgramFiles%\ASUS\*|*.log;*.tmp;*log*.txt|RECURSE")

        Assert.IsFalse(EntryModified(oldSection, newSection))

    End Sub

    ''' <summary>
    ''' A reorder that also adds a new pattern is still a real change
    ''' </summary>
    <TestMethod()> Public Sub FileKeyPatternReorderWithAddition_IsAChange()

        Dim oldSection = MakeSection("LangSecRef=3021",
                                     "Detect=HKCU\Software\App",
                                     "FileKey1=%ProgramFiles%\ASUS\*|*.log;*.tmp|RECURSE")

        Dim newSection = MakeSection("LangSecRef=3021",
                                     "Detect=HKCU\Software\App",
                                     "FileKey1=%ProgramFiles%\ASUS\*|*.tmp;*.log;*.bak|RECURSE")

        Assert.IsTrue(EntryModified(oldSection, newSection))

    End Sub

    ''' <summary>
    ''' Same patterns but a changed deletion flag is a real change
    ''' </summary>
    <TestMethod()> Public Sub FileKeyFlagChange_IsAChange()

        Dim oldSection = MakeSection("LangSecRef=3021",
                                     "Detect=HKCU\Software\App",
                                     "FileKey1=%ProgramFiles%\ASUS\*|*.log;*.tmp|RECURSE")

        Dim newSection = MakeSection("LangSecRef=3021",
                                     "Detect=HKCU\Software\App",
                                     "FileKey1=%ProgramFiles%\ASUS\*|*.tmp;*.log")

        Assert.IsTrue(EntryModified(oldSection, newSection))

    End Sub

    ''' <summary>
    ''' Same patterns and flag but a changed path is a real change
    ''' </summary>
    <TestMethod()> Public Sub FileKeyPathChange_IsAChange()

        Dim oldSection = MakeSection("LangSecRef=3021",
                                     "Detect=HKCU\Software\App",
                                     "FileKey1=%ProgramFiles%\ASUS\*|*.log;*.tmp|RECURSE")

        Dim newSection = MakeSection("LangSecRef=3021",
                                     "Detect=HKCU\Software\App",
                                     "FileKey1=%ProgramFiles%\ASUS\Logs\*|*.tmp;*.log|RECURSE")

        Assert.IsTrue(EntryModified(oldSection, newSection))

    End Sub

    ''' <summary>
    ''' A duplicated pattern on only one side is not absorbed as a reorder — the multiset differs
    ''' </summary>
    <TestMethod()> Public Sub FileKeyDuplicatePatternAsymmetry_IsAChange()

        Dim oldSection = MakeSection("LangSecRef=3021",
                                     "Detect=HKCU\Software\App",
                                     "FileKey1=%ProgramFiles%\ASUS\*|*.log;*.log;*.tmp|RECURSE")

        Dim newSection = MakeSection("LangSecRef=3021",
                                     "Detect=HKCU\Software\App",
                                     "FileKey1=%ProgramFiles%\ASUS\*|*.tmp;*.log|RECURSE")

        Assert.IsTrue(EntryModified(oldSection, newSection))

    End Sub

End Class
