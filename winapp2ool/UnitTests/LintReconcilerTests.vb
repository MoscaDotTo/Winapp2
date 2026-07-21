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
''' Tests for the lint reconciliation gate (<c> LintReconciler </c>): the semantic-unit
''' comparison around the generative modules' WinappDebug normalization pass must be
''' blind to the three sanctioned optimization classes (key reordering/renumbering,
''' FileKey pattern alphabetization, same-path FileKey merges) while reporting every
''' semantic loss, rewrite, invented unit, or dropped entry. Also covers EntryBuilder's
''' per-letter artifact bucketing rule.
''' </summary>
<TestClass()> Public Class LintReconcilerTests

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
    ''' Helper: reconcile two literal ini texts and return the findings
    ''' </summary>
    Private Shared Function Reconcile(preText As String, postText As String) As List(Of String)

        Return winapp2ool.LintReconciler.FindSemanticLosses(
            winapp2ool.LintReconciler.CollectSemanticUnits(MakeIni(preText)),
            winapp2ool.LintReconciler.CollectSemanticUnits(MakeIni(postText)))

    End Function

    ''' <summary>
    ''' Merging two FileKeys that share a path and flag into one semicolon-delimited key
    ''' is a sanctioned optimization and must produce zero findings
    ''' </summary>
    <TestMethod()> Public Sub FileKeyMerge_IsInvisible()

        Dim pre = "[App *]" & vbCrLf &
                  "LangSecRef=3021" & vbCrLf &
                  "Detect=HKCU\Software\App" & vbCrLf &
                  "FileKey1=%AppData%\App|*.log" & vbCrLf &
                  "FileKey2=%AppData%\App|*.tmp" & vbCrLf

        Dim post = "[App *]" & vbCrLf &
                   "LangSecRef=3021" & vbCrLf &
                   "Detect=HKCU\Software\App" & vbCrLf &
                   "FileKey1=%AppData%\App|*.log;*.tmp" & vbCrLf

        Dim losses = Reconcile(pre, post)

        Assert.AreEqual(0, losses.Count, String.Join("; ", losses))

    End Sub

    ''' <summary>
    ''' Key renumbering and reordering changes neither KeyType nor value and must
    ''' produce zero findings
    ''' </summary>
    <TestMethod()> Public Sub RenumberAndReorder_IsInvisible()

        Dim pre = "[App *]" & vbCrLf &
                  "LangSecRef=3021" & vbCrLf &
                  "FileKey1=%ProgramData%\App\Logs|*.log|RECURSE" & vbCrLf &
                  "FileKey2=%AppData%\App|*.ini" & vbCrLf &
                  "RegKey1=HKCU\Software\App|MRU" & vbCrLf

        Dim post = "[App *]" & vbCrLf &
                   "LangSecRef=3021" & vbCrLf &
                   "FileKey1=%AppData%\App|*.ini" & vbCrLf &
                   "FileKey2=%ProgramData%\App\Logs|*.log|RECURSE" & vbCrLf &
                   "RegKey1=HKCU\Software\App|MRU" & vbCrLf

        Dim losses = Reconcile(pre, post)

        Assert.AreEqual(0, losses.Count, String.Join("; ", losses))

    End Sub

    ''' <summary>
    ''' Alphabetizing a FileKey's semicolon-delimited pattern list does not change the
    ''' pattern set and must produce zero findings
    ''' </summary>
    <TestMethod()> Public Sub PatternAlphabetization_IsInvisible()

        Dim pre = "[App *]" & vbCrLf & "LangSecRef=3021" & vbCrLf &
                  "FileKey1=%AppData%\App|*.log;*.err" & vbCrLf

        Dim post = "[App *]" & vbCrLf & "LangSecRef=3021" & vbCrLf &
                   "FileKey1=%AppData%\App|*.err;*.log" & vbCrLf

        Dim losses = Reconcile(pre, post)

        Assert.AreEqual(0, losses.Count, String.Join("; ", losses))

    End Sub

    ''' <summary>
    ''' A FileKey pattern present before the pass but absent afterwards is data loss
    ''' and must be reported
    ''' </summary>
    <TestMethod()> Public Sub DroppedPattern_IsReported()

        Dim pre = "[App *]" & vbCrLf & "LangSecRef=3021" & vbCrLf &
                  "FileKey1=%AppData%\App|*.err;*.log" & vbCrLf

        Dim post = "[App *]" & vbCrLf & "LangSecRef=3021" & vbCrLf &
                   "FileKey1=%AppData%\App|*.err" & vbCrLf

        Dim losses = Reconcile(pre, post)

        Assert.AreEqual(1, losses.Count, String.Join("; ", losses))
        StringAssert.Contains(losses(0), "[App *]")
        StringAssert.Contains(losses(0), "lost content")

    End Sub

    ''' <summary>
    ''' An entry removed outright by the pass must be reported
    ''' </summary>
    <TestMethod()> Public Sub DroppedEntry_IsReported()

        Dim pre = "[App *]" & vbCrLf & "LangSecRef=3021" & vbCrLf &
                  "FileKey1=%AppData%\App|*.log" & vbCrLf & vbCrLf &
                  "[Gone *]" & vbCrLf & "LangSecRef=3021" & vbCrLf &
                  "FileKey1=%AppData%\Gone|*.log" & vbCrLf

        Dim post = "[App *]" & vbCrLf & "LangSecRef=3021" & vbCrLf &
                   "FileKey1=%AppData%\App|*.log" & vbCrLf

        Dim losses = Reconcile(pre, post)

        Assert.AreEqual(1, losses.Count, String.Join("; ", losses))
        StringAssert.Contains(losses(0), "[Gone *]")
        StringAssert.Contains(losses(0), "removed entirely")

    End Sub

    ''' <summary>
    ''' A rewritten value (here a RegKey) with an unambiguous one-lost/one-gained
    ''' correspondence is reported as a single <c> old → new </c> rewrite finding
    ''' rather than a lost/gained pair
    ''' </summary>
    <TestMethod()> Public Sub RewrittenValue_IsReportedAsSingleRewrite()

        Dim pre = "[App *]" & vbCrLf & "LangSecRef=3021" & vbCrLf &
                  "RegKey1=HKCU\Software\App|MRU 1" & vbCrLf

        Dim post = "[App *]" & vbCrLf & "LangSecRef=3021" & vbCrLf &
                   "RegKey1=HKCU\Software\App|1" & vbCrLf

        Dim losses = Reconcile(pre, post)

        Assert.AreEqual(1, losses.Count, String.Join("; ", losses))
        StringAssert.Contains(losses(0), "rewritten")
        StringAssert.Contains(losses(0), "HKCU\Software\App|MRU 1")
        StringAssert.Contains(losses(0), "HKCU\Software\App|1")

    End Sub

    ''' <summary>
    ''' A FileKey whose pattern is rewritten (path and flag unchanged) pairs into a
    ''' single rewrite finding
    ''' </summary>
    <TestMethod()> Public Sub FileKeyPatternRewrite_IsReportedAsSingleRewrite()

        Dim pre = "[App *]" & vbCrLf & "LangSecRef=3021" & vbCrLf &
                  "FileKey1=%AppData%\App|*.log" & vbCrLf

        Dim post = "[App *]" & vbCrLf & "LangSecRef=3021" & vbCrLf &
                   "FileKey1=%AppData%\App|*.txt" & vbCrLf

        Dim losses = Reconcile(pre, post)

        Assert.AreEqual(1, losses.Count, String.Join("; ", losses))
        StringAssert.Contains(losses(0), "rewritten")
        StringAssert.Contains(losses(0), "*.log")
        StringAssert.Contains(losses(0), "*.txt")

    End Sub

    ''' <summary>
    ''' A FileKey whose path is rewritten (pattern and flag unchanged) pairs into a
    ''' single rewrite finding
    ''' </summary>
    <TestMethod()> Public Sub FileKeyPathRewrite_IsReportedAsSingleRewrite()

        Dim pre = "[App *]" & vbCrLf & "LangSecRef=3021" & vbCrLf &
                  "FileKey1=%AppData%\App|*.log" & vbCrLf

        Dim post = "[App *]" & vbCrLf & "LangSecRef=3021" & vbCrLf &
                   "FileKey1=%LocalAppData%\App|*.log" & vbCrLf

        Dim losses = Reconcile(pre, post)

        Assert.AreEqual(1, losses.Count, String.Join("; ", losses))
        StringAssert.Contains(losses(0), "rewritten")
        StringAssert.Contains(losses(0), "%AppData%\App")
        StringAssert.Contains(losses(0), "%LocalAppData%\App")

    End Sub

    ''' <summary>
    ''' A FileKey gaining a flag (path and pattern unchanged) pairs into a single
    ''' rewrite finding rather than being misread as a pattern change
    ''' </summary>
    <TestMethod()> Public Sub FileKeyFlagChange_IsReportedAsSingleRewrite()

        Dim pre = "[App *]" & vbCrLf & "LangSecRef=3021" & vbCrLf &
                  "FileKey1=%AppData%\App|*.log" & vbCrLf

        Dim post = "[App *]" & vbCrLf & "LangSecRef=3021" & vbCrLf &
                   "FileKey1=%AppData%\App|*.log|RECURSE" & vbCrLf

        Dim losses = Reconcile(pre, post)

        Assert.AreEqual(1, losses.Count, String.Join("; ", losses))
        StringAssert.Contains(losses(0), "rewritten")
        StringAssert.Contains(losses(0), "RECURSE")

    End Sub

    ''' <summary>
    ''' When the lost/gained correspondence is ambiguous (two lost patterns, one gained),
    ''' nothing is paired and every difference stays reported as separate lost/gained
    ''' findings
    ''' </summary>
    <TestMethod()> Public Sub AmbiguousRewrite_StaysLostAndGained()

        Dim pre = "[App *]" & vbCrLf & "LangSecRef=3021" & vbCrLf &
                  "FileKey1=%AppData%\App|*.log;*.err" & vbCrLf

        Dim post = "[App *]" & vbCrLf & "LangSecRef=3021" & vbCrLf &
                   "FileKey1=%AppData%\App|*.txt" & vbCrLf

        Dim losses = Reconcile(pre, post)

        Assert.AreEqual(3, losses.Count, String.Join("; ", losses))
        Assert.IsFalse(losses.Any(Function(l) l.Contains("rewritten")), String.Join("; ", losses))
        Assert.AreEqual(2, losses.Where(Function(l) l.Contains("lost content")).Count, String.Join("; ", losses))
        Assert.AreEqual(1, losses.Where(Function(l) l.Contains("gained unexpected content")).Count, String.Join("; ", losses))

    End Sub

    ''' <summary>
    ''' A lost key of one type alongside a gained key of another type is never paired —
    ''' the gate does not guess at cross-type conversions
    ''' </summary>
    <TestMethod()> Public Sub CrossTypeChange_IsNotPaired()

        Dim pre = "[App *]" & vbCrLf & "LangSecRef=3021" & vbCrLf &
                  "Detect=HKCU\Software\App" & vbCrLf &
                  "FileKey1=%AppData%\App|*.log" & vbCrLf

        Dim post = "[App *]" & vbCrLf & "LangSecRef=3021" & vbCrLf &
                   "DetectFile=%AppData%\App" & vbCrLf &
                   "FileKey1=%AppData%\App|*.log" & vbCrLf

        Dim losses = Reconcile(pre, post)

        Assert.AreEqual(2, losses.Count, String.Join("; ", losses))
        Assert.IsFalse(losses.Any(Function(l) l.Contains("rewritten")), String.Join("; ", losses))
        Assert.IsTrue(losses.Any(Function(l) l.Contains("lost content")), String.Join("; ", losses))
        Assert.IsTrue(losses.Any(Function(l) l.Contains("gained unexpected content")), String.Join("; ", losses))

    End Sub

    ''' <summary>
    ''' Case-only differences are cosmetic and must produce zero findings
    ''' </summary>
    <TestMethod()> Public Sub CaseOnlyChange_IsInvisible()

        Dim pre = "[App *]" & vbCrLf & "LangSecRef=3021" & vbCrLf &
                  "FileKey1=%appdata%\App|*.LOG|recurse" & vbCrLf

        Dim post = "[App *]" & vbCrLf & "LangSecRef=3021" & vbCrLf &
                   "FileKey1=%AppData%\App|*.log|RECURSE" & vbCrLf

        Dim losses = Reconcile(pre, post)

        Assert.AreEqual(0, losses.Count, String.Join("; ", losses))

    End Sub

    ''' <summary>
    ''' An entry present only after the pass (invented by the linter) must be reported
    ''' </summary>
    <TestMethod()> Public Sub IntroducedEntry_IsReported()

        Dim pre = "[App *]" & vbCrLf & "LangSecRef=3021" & vbCrLf &
                  "FileKey1=%AppData%\App|*.log" & vbCrLf

        Dim post = pre & vbCrLf &
                   "[New *]" & vbCrLf & "LangSecRef=3021" & vbCrLf &
                   "FileKey1=%AppData%\New|*.log" & vbCrLf

        Dim losses = Reconcile(pre, post)

        Assert.AreEqual(1, losses.Count, String.Join("; ", losses))
        StringAssert.Contains(losses(0), "[New *]")
        StringAssert.Contains(losses(0), "introduced")

    End Sub

    ''' <summary>
    ''' Removing an exact duplicate FileKey pattern is sanctioned deduplication and
    ''' must produce zero findings
    ''' </summary>
    <TestMethod()> Public Sub ExactDuplicateRemoval_IsInvisible()

        Dim pre = "[App *]" & vbCrLf & "LangSecRef=3021" & vbCrLf &
                  "FileKey1=%AppData%\App|*.log;*.log" & vbCrLf &
                  "FileKey2=%AppData%\App|*.log" & vbCrLf

        Dim post = "[App *]" & vbCrLf & "LangSecRef=3021" & vbCrLf &
                   "FileKey1=%AppData%\App|*.log" & vbCrLf

        Dim losses = Reconcile(pre, post)

        Assert.AreEqual(0, losses.Count, String.Join("; ", losses))

    End Sub

    ''' <summary>
    ''' The per-letter artifact bucketing rule: uppercased ASCII first letter, with
    ''' digits, punctuation, and anything non-alphabetic routed to <c> # </c>
    ''' </summary>
    <TestMethod()> Public Sub LetterBucketFor_ClassifiesByFirstCharacter()

        Assert.AreEqual("D", winapp2ool.EntryBuilder.LetterBucketFor("Discord *"))
        Assert.AreEqual("D", winapp2ool.EntryBuilder.LetterBucketFor("discord *"))
        Assert.AreEqual("#", winapp2ool.EntryBuilder.LetterBucketFor("3DMark *"))
        Assert.AreEqual("#", winapp2ool.EntryBuilder.LetterBucketFor(".NET Framework *"))
        Assert.AreEqual("#", winapp2ool.EntryBuilder.LetterBucketFor(""))

    End Sub

End Class
