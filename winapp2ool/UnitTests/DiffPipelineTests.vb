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
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports winapp2ool

''' <summary>
''' Runs whole files through the Diff pipeline (<c> Diff.CompareFiles </c>) and checks the rendered
''' changelog and the outcome counts
''' </summary>
<TestClass>
Public Class DiffPipelineTests

    ''' <summary>
    ''' Two removed entries that each qualify as a rename of <c> Gamma * </c>. Alpha sorts first, so
    ''' its rename is recorded and then demoted to a merger when Beta's is refused.
    ''' </summary>
    Private Const ContestedRenameOld As String =
        "[Alpha *]" & vbCrLf &
        "LangSecRef=3021" & vbCrLf &
        "DetectFile=%AppData%\Foo" & vbCrLf &
        "FileKey1=%AppData%\Foo|*.log" & vbCrLf & vbCrLf &
        "[Beta *]" & vbCrLf &
        "LangSecRef=3021" & vbCrLf &
        "DetectFile=%AppData%\Bar" & vbCrLf &
        "FileKey1=%AppData%\Foo|*.log" & vbCrLf

    ''' <summary>
    ''' Returns an <c> iniFile </c> parsed from <paramref name="text"/>
    ''' </summary>
    '''
    ''' <param name="text">
    ''' Literal ini text
    ''' </param>
    Private Shared Function MakeIni(text As String) As iniFile

        Using reader As New StreamReader(New MemoryStream(Encoding.UTF8.GetBytes(text)))
            Return iniFile.FromStream(reader, "", "test.ini")
        End Using

    End Function

    ''' <summary>
    ''' Diffs <paramref name="oldText"/> against <paramref name="newText"/> with the console
    ''' redirected, and returns everything the changelog printed
    ''' </summary>
    '''
    ''' <param name="oldText">
    ''' The old file's ini text
    ''' </param>
    '''
    ''' <param name="newText">
    ''' The new file's ini text
    ''' </param>
    Private Shared Function RunDiff(oldText As String, newText As String) As String

        Dim original = Console.Out

        Try

            Using captured As New StringWriter

                Console.SetOut(captured)
                For Each section In Diff.CompareFiles(MakeIni(oldText), MakeIni(newText))
                    section.Print()
                Next
                Return captured.ToString()

            End Using

        Finally

            Console.SetOut(original)

        End Try

    End Function

    ''' <summary>
    ''' Counts the occurrences of <paramref name="text"/> in <paramref name="output"/>
    ''' </summary>
    '''
    ''' <param name="output">
    ''' The rendered changelog
    ''' </param>
    '''
    ''' <param name="text">
    ''' The text to count
    ''' </param>
    Private Shared Function CountOf(output As String, text As String) As Integer

        Dim count = 0
        Dim at = output.IndexOf(text, StringComparison.Ordinal)

        While at >= 0

            count += 1
            at = output.IndexOf(text, at + text.Length, StringComparison.Ordinal)

        End While

        Return count

    End Function

    ''' <summary>
    ''' Gamma's Warning comes from no merged source, so Delta losing the same Warning is a move
    ''' into Gamma, credited to Delta in the merger output rather than called novel
    ''' </summary>
    <TestMethod>
    Public Sub DemotedRename_KeyMovesIntoMergerTarget()

        Dim oldText =
            "[Alpha *]" & vbCrLf & "LangSecRef=3021" & vbCrLf & "DetectFile=%AppData%\Foo" & vbCrLf &
            "FileKey1=%AppData%\Foo|*.log" & vbCrLf & vbCrLf &
            "[Beta *]" & vbCrLf & "LangSecRef=3021" & vbCrLf & "DetectFile=%AppData%\Foo" & vbCrLf &
            "FileKey1=%AppData%\Foo|*.tmp" & vbCrLf & vbCrLf &
            "[Delta *]" & vbCrLf & "LangSecRef=3021" & vbCrLf & "DetectFile=%AppData%\Delta" & vbCrLf &
            "FileKey1=%AppData%\Delta|*.dat" & vbCrLf & "Warning=Closes Foo first" & vbCrLf

        Dim newText =
            "[Delta *]" & vbCrLf & "LangSecRef=3021" & vbCrLf & "DetectFile=%AppData%\Delta" & vbCrLf &
            "FileKey1=%AppData%\Delta|*.dat" & vbCrLf & vbCrLf &
            "[Gamma *]" & vbCrLf & "LangSecRef=3021" & vbCrLf & "DetectFile=%AppData%\Foo" & vbCrLf &
            "FileKey1=%AppData%\Foo|*.log;*.tmp" & vbCrLf & "Warning=Closes Foo first" & vbCrLf

        Dim output = RunDiff(oldText, newText)

        Dim outcome = Diff.MostRecentDiffOutcome
        Assert.AreEqual(2, outcome.MergedEntries)
        Assert.AreEqual(0, outcome.RenamedEntries)
        Assert.AreEqual(1, outcome.MovedKeys)
        Assert.AreEqual(0, outcome.RemovedKeys)
        StringAssert.Contains(output, "Warning=Closes Foo first moved to [Gamma *]")
        StringAssert.Contains(output, "Warning=Closes Foo first (from [Delta *])")
        Assert.IsFalse(output.Contains("novel key"), output)

    End Sub

    ''' <summary>
    ''' Gamma's Warning is carried over from Alpha and also left Delta, so it moves and both get
    ''' the credit. A demoted rename's leftover records would measure Gamma against Alpha alone,
    ''' drop the shared Warning from the candidates, and report no move.
    ''' </summary>
    <TestMethod>
    Public Sub DemotedRename_SharedKeyStillMoves()

        Dim oldText =
            "[Alpha *]" & vbCrLf & "LangSecRef=3021" & vbCrLf & "DetectFile=%AppData%\Foo" & vbCrLf &
            "FileKey1=%AppData%\Foo|*.log" & vbCrLf & "Warning=Closes Foo first" & vbCrLf & vbCrLf &
            "[Beta *]" & vbCrLf & "LangSecRef=3021" & vbCrLf & "DetectFile=%AppData%\Foo" & vbCrLf &
            "FileKey1=%AppData%\Foo|*.tmp" & vbCrLf & vbCrLf &
            "[Delta *]" & vbCrLf & "LangSecRef=3021" & vbCrLf & "DetectFile=%AppData%\Delta" & vbCrLf &
            "FileKey1=%AppData%\Delta|*.dat" & vbCrLf & "Warning=Closes Foo first" & vbCrLf

        Dim newText =
            "[Delta *]" & vbCrLf & "LangSecRef=3021" & vbCrLf & "DetectFile=%AppData%\Delta" & vbCrLf &
            "FileKey1=%AppData%\Delta|*.dat" & vbCrLf & vbCrLf &
            "[Gamma *]" & vbCrLf & "LangSecRef=3021" & vbCrLf & "DetectFile=%AppData%\Foo" & vbCrLf &
            "FileKey1=%AppData%\Foo|*.log;*.tmp" & vbCrLf & "Warning=Closes Foo first" & vbCrLf

        Dim output = RunDiff(oldText, newText)

        Dim outcome = Diff.MostRecentDiffOutcome
        Assert.AreEqual(2, outcome.MergedEntries)
        Assert.AreEqual(1, outcome.MovedKeys)
        StringAssert.Contains(output, "Warning=Closes Foo first (from [Alpha *], [Delta *])")

    End Sub

    ''' <summary>
    ''' A key that leaves a surviving entry for an entry that didn't exist before has moved
    ''' </summary>
    <TestMethod>
    Public Sub KeyMovesIntoNovelEntry()

        Dim oldText =
            "[Delta *]" & vbCrLf & "LangSecRef=3021" & vbCrLf & "DetectFile=%AppData%\Delta" & vbCrLf &
            "FileKey1=%AppData%\Delta|*.dat" & vbCrLf & "FileKey2=%AppData%\Delta\Sync|*" & vbCrLf

        Dim newText =
            "[Delta *]" & vbCrLf & "LangSecRef=3021" & vbCrLf & "DetectFile=%AppData%\Delta" & vbCrLf &
            "FileKey1=%AppData%\Delta|*.dat" & vbCrLf & vbCrLf &
            "[Delta Sync *]" & vbCrLf & "LangSecRef=3022" & vbCrLf & "DetectFile=%AppData%\Delta" & vbCrLf &
            "FileKey1=%AppData%\Delta\Sync|*" & vbCrLf

        Dim output = RunDiff(oldText, newText)

        Dim outcome = Diff.MostRecentDiffOutcome
        Assert.AreEqual(1, outcome.MovedKeys)
        Assert.AreEqual(0, outcome.RemovedKeys)
        StringAssert.Contains(output, "FileKey2=%AppData%\Delta\Sync|* moved to [Delta Sync *]")
        StringAssert.Contains(output, "Moved 1 FileKey")
        StringAssert.Contains(output, "+ 1 key moved into 1 added entry")

    End Sub

    ''' <summary>
    ''' A key moving between two entries that both exist shows in each one's itemized output,
    ''' and an entry whose only change is a move still gets itemized
    ''' </summary>
    <TestMethod>
    Public Sub KeyMovesIntoModifiedEntry_ItemizedOnBothSides()

        Dim oldText =
            "[Delta *]" & vbCrLf & "LangSecRef=3021" & vbCrLf & "DetectFile=%AppData%\Delta" & vbCrLf &
            "FileKey1=%AppData%\Delta|*.dat" & vbCrLf & "FileKey2=%AppData%\Shared|*.log" & vbCrLf & vbCrLf &
            "[Epsilon *]" & vbCrLf & "LangSecRef=3021" & vbCrLf & "DetectFile=%AppData%\Epsilon" & vbCrLf &
            "FileKey1=%AppData%\Epsilon|*.dat" & vbCrLf

        Dim newText =
            "[Delta *]" & vbCrLf & "LangSecRef=3021" & vbCrLf & "DetectFile=%AppData%\Delta" & vbCrLf &
            "FileKey1=%AppData%\Delta|*.dat" & vbCrLf & vbCrLf &
            "[Epsilon *]" & vbCrLf & "LangSecRef=3021" & vbCrLf & "DetectFile=%AppData%\Epsilon" & vbCrLf &
            "FileKey1=%AppData%\Epsilon|*.dat" & vbCrLf & "FileKey2=%AppData%\Shared|*.log" & vbCrLf

        Dim output = RunDiff(oldText, newText)

        Assert.AreEqual(1, Diff.MostRecentDiffOutcome.MovedKeys)
        StringAssert.Contains(output, "2 modified entries")
        StringAssert.Contains(output, "Moved 1 FileKey")
        StringAssert.Contains(output, "FileKey2=%AppData%\Shared|*.log moved to [Epsilon *]")
        StringAssert.Contains(output, "Moved in 1 FileKey")
        StringAssert.Contains(output, "FileKey2=%AppData%\Shared|*.log (from [Delta *])")
        StringAssert.Contains(output, "~ 1 key moved into 1 modified entry")
        StringAssert.Contains(output, "1 key moved to other entries:")
        StringAssert.Contains(output, "1 key moved in from other entries:")
        StringAssert.Contains(output, "Delta * moved 1 key:")

    End Sub

    ''' <summary>
    ''' Delta's unchanged Detect shares its value with the RegKey that moves in from Epsilon, and
    ''' must not be listed as moved in alongside it
    ''' </summary>
    <TestMethod>
    Public Sub MovedIn_ListsOnlyTheKeyThatMoved()

        Dim oldText =
            "[Delta *]" & vbCrLf & "LangSecRef=3021" & vbCrLf & "Detect=HKCU\Software\Shared" & vbCrLf &
            "RegKey1=HKCU\Software\Delta" & vbCrLf & vbCrLf &
            "[Epsilon *]" & vbCrLf & "LangSecRef=3021" & vbCrLf & "Detect=HKCU\Software\Epsilon" & vbCrLf &
            "RegKey1=HKCU\Software\Epsilon" & vbCrLf & "RegKey2=HKCU\Software\Shared" & vbCrLf

        Dim newText =
            "[Delta *]" & vbCrLf & "LangSecRef=3021" & vbCrLf & "Detect=HKCU\Software\Shared" & vbCrLf &
            "RegKey1=HKCU\Software\Delta" & vbCrLf & "RegKey2=HKCU\Software\Shared" & vbCrLf & vbCrLf &
            "[Epsilon *]" & vbCrLf & "LangSecRef=3021" & vbCrLf & "Detect=HKCU\Software\Epsilon" & vbCrLf &
            "RegKey1=HKCU\Software\Epsilon" & vbCrLf

        Dim output = RunDiff(oldText, newText)

        Assert.AreEqual(1, Diff.MostRecentDiffOutcome.MovedKeys)
        StringAssert.Contains(output, "RegKey2=HKCU\Software\Shared (from [Epsilon *])")
        Assert.IsFalse(output.Contains("Moved in 1 Detect"), output)
        Assert.IsFalse(output.Contains("Detect=HKCU\Software\Shared (from"), output)

    End Sub

    ''' <summary>
    ''' Delta absorbs the removed Rho and gives a key away to the new Zeta. The merger block's
    ''' re-diff must list that key as moved, not removed.
    ''' </summary>
    <TestMethod>
    Public Sub MergerTargetGivingAKeyAway_ListsItAsMoved()

        Dim oldText =
            "[Delta *]" & vbCrLf & "LangSecRef=3021" & vbCrLf & "DetectFile=%AppData%\Delta" & vbCrLf &
            "FileKey1=%AppData%\Delta|*.dat" & vbCrLf & "FileKey2=%AppData%\Delta\Sync|*" & vbCrLf & vbCrLf &
            "[Rho *]" & vbCrLf & "LangSecRef=3021" & vbCrLf & "DetectFile=%AppData%\Rho" & vbCrLf &
            "FileKey1=%AppData%\Rho|*.log" & vbCrLf

        Dim newText =
            "[Delta *]" & vbCrLf & "LangSecRef=3021" & vbCrLf & "DetectFile1=%AppData%\Delta" & vbCrLf &
            "DetectFile2=%AppData%\Rho" & vbCrLf & "FileKey1=%AppData%\Delta|*.dat" & vbCrLf &
            "FileKey2=%AppData%\Rho|*.log" & vbCrLf & vbCrLf &
            "[Zeta *]" & vbCrLf & "LangSecRef=3022" & vbCrLf & "DetectFile=%AppData%\Delta" & vbCrLf &
            "FileKey1=%AppData%\Delta\Sync|*" & vbCrLf

        Dim output = RunDiff(oldText, newText)

        Assert.AreEqual(1, Diff.MostRecentDiffOutcome.MergedEntries)
        Assert.AreEqual(1, Diff.MostRecentDiffOutcome.MovedKeys)
        StringAssert.Contains(output, "1 modified entry incorporating merged content")
        StringAssert.Contains(output, "Moved 1 FileKey")
        Assert.IsFalse(output.Contains("Removed 1 FileKey"), output)

    End Sub

    ''' <summary>
    ''' Delta absorbs the removed Rho. Delta's own lost FileKey is a removal, but Rho's Detect,
    ''' which Delta never had, is listed as a merged key that didn't survive.
    ''' </summary>
    <TestMethod>
    Public Sub MergerIntoModifiedEntry_SeparatesDroppedSourceKeys()

        Dim oldText =
            "[Delta *]" & vbCrLf & "LangSecRef=3021" & vbCrLf & "DetectFile=%AppData%\Delta" & vbCrLf &
            "FileKey1=%AppData%\Delta|*.dat" & vbCrLf & "FileKey2=%AppData%\Delta|*.old" & vbCrLf & vbCrLf &
            "[Rho *]" & vbCrLf & "LangSecRef=3021" & vbCrLf & "Detect=HKCU\Software\Rho" & vbCrLf &
            "FileKey1=%AppData%\Rho|*.log" & vbCrLf

        Dim newText =
            "[Delta *]" & vbCrLf & "LangSecRef=3021" & vbCrLf & "DetectFile=%AppData%\Delta" & vbCrLf &
            "FileKey1=%AppData%\Delta|*.dat" & vbCrLf & "FileKey2=%AppData%\Rho|*.log" & vbCrLf

        Dim output = RunDiff(oldText, newText)

        StringAssert.Contains(output, "1 modified entry incorporating merged content")

        Dim ownRemoval = output.IndexOf("FileKey2=%AppData%\Delta|*.old", StringComparison.Ordinal)
        Dim header = output.IndexOf("1 key from merged entries not in this entry:", StringComparison.Ordinal)
        Dim dropped = output.IndexOf("Detect=HKCU\Software\Rho (from [Rho *])", StringComparison.Ordinal)

        Assert.IsTrue(ownRemoval >= 0 AndAlso header > ownRemoval AndAlso dropped > header, output)

        Dim modifiedHeader = output.IndexOf("Delta * has been modified", StringComparison.Ordinal)
        Dim sources = output.IndexOf("This entry contains merged keys from 1 removed entry:", StringComparison.Ordinal)
        Assert.IsTrue(modifiedHeader >= 0 AndAlso sources > modifiedHeader AndAlso ownRemoval > sources, output)

        Dim removedHeading = output.IndexOf("1 key removed:", StringComparison.Ordinal)
        Assert.IsTrue(removedHeading > sources AndAlso ownRemoval > removedHeading, output)
        StringAssert.Contains(output, "Dropped 1 Detect")

    End Sub

    ''' <summary>
    ''' A rename whose only key change is a key moved into it isn't a name-only rename. The moved
    ''' key is a Warning because gaining a FileKey would make Alpha a merger instead.
    ''' </summary>
    <TestMethod>
    Public Sub RenameReceivingAMovedKey_IsNotNameOnly()

        Dim oldText =
            "[Alpha *]" & vbCrLf & "LangSecRef=3021" & vbCrLf & "DetectFile=%AppData%\Alpha" & vbCrLf &
            "FileKey1=%AppData%\Alpha|*.log" & vbCrLf & vbCrLf &
            "[Delta *]" & vbCrLf & "LangSecRef=3021" & vbCrLf & "DetectFile=%AppData%\Delta" & vbCrLf &
            "FileKey1=%AppData%\Delta|*.dat" & vbCrLf & "Warning=Closes Alpha first" & vbCrLf

        Dim newText =
            "[Alpha Renamed *]" & vbCrLf & "LangSecRef=3021" & vbCrLf & "DetectFile=%AppData%\Alpha" & vbCrLf &
            "FileKey1=%AppData%\Alpha|*.log" & vbCrLf & "Warning=Closes Alpha first" & vbCrLf & vbCrLf &
            "[Delta *]" & vbCrLf & "LangSecRef=3021" & vbCrLf & "DetectFile=%AppData%\Delta" & vbCrLf &
            "FileKey1=%AppData%\Delta|*.dat" & vbCrLf

        Dim output = RunDiff(oldText, newText)

        Assert.AreEqual(1, Diff.MostRecentDiffOutcome.RenamedEntries)
        Assert.AreEqual(1, Diff.MostRecentDiffOutcome.MovedKeys)
        StringAssert.Contains(output, "Minor changes to 1 renamed entry")
        StringAssert.Contains(output, "Moved in 1 Warning")
        StringAssert.Contains(output, "1 key moved in from other entries:")
        StringAssert.Contains(output, "& 1 key moved into 1 renamed entry")
        Assert.IsFalse(output.Contains("name-only"), output)

    End Sub

    ''' <summary>
    ''' A moved key lands in an exact copy of itself before a broader key that captures it, even
    ''' when the broader key's entry sorts first
    ''' </summary>
    <TestMethod>
    Public Sub KeyMovement_PrefersExactMatch()

        Dim oldText =
            "[Delta *]" & vbCrLf & "LangSecRef=3021" & vbCrLf & "DetectFile=%AppData%\Delta" & vbCrLf &
            "FileKey1=%AppData%\Delta|*.dat" & vbCrLf & "FileKey2=%AppData%\Foo|*.log" & vbCrLf

        Dim newText =
            "[Alpha *]" & vbCrLf & "LangSecRef=3022" & vbCrLf & "DetectFile=%AppData%\Alpha" & vbCrLf &
            "FileKey1=%AppData%\Foo|*" & vbCrLf & vbCrLf &
            "[Delta *]" & vbCrLf & "LangSecRef=3021" & vbCrLf & "DetectFile=%AppData%\Delta" & vbCrLf &
            "FileKey1=%AppData%\Delta|*.dat" & vbCrLf & vbCrLf &
            "[Zeta *]" & vbCrLf & "LangSecRef=3022" & vbCrLf & "DetectFile=%AppData%\Zeta" & vbCrLf &
            "FileKey1=%AppData%\Foo|*.log" & vbCrLf

        Dim output = RunDiff(oldText, newText)

        Assert.AreEqual(1, Diff.MostRecentDiffOutcome.MovedKeys)
        StringAssert.Contains(output, "FileKey2=%AppData%\Foo|*.log moved to [Zeta *]")

    End Sub

    ''' <summary>
    ''' A classifier shared by many entries never moves: Delta's dropped LangSecRef stays a removal
    ''' even though a new entry uses the same value
    ''' </summary>
    <TestMethod>
    Public Sub KeyMovement_IgnoresClassifiers()

        Dim oldText =
            "[Delta *]" & vbCrLf & "LangSecRef=3021" & vbCrLf & "Section=Delta" & vbCrLf &
            "DetectFile=%AppData%\Delta" & vbCrLf & "FileKey1=%AppData%\Delta|*.dat" & vbCrLf

        Dim newText =
            "[Delta *]" & vbCrLf & "Section=Delta" & vbCrLf &
            "DetectFile=%AppData%\Delta" & vbCrLf & "FileKey1=%AppData%\Delta|*.dat" & vbCrLf & vbCrLf &
            "[Zeta *]" & vbCrLf & "LangSecRef=3021" & vbCrLf & "DetectFile=%AppData%\Zeta" & vbCrLf &
            "FileKey1=%AppData%\Zeta|*.log" & vbCrLf

        RunDiff(oldText, newText)

        Assert.AreEqual(0, Diff.MostRecentDiffOutcome.MovedKeys)

    End Sub

    ''' <summary>
    ''' A Detect on a parent path captures Delta's dropped child-path Detect, but that isn't
    ''' where the key went, so it stays a removal
    ''' </summary>
    <TestMethod>
    Public Sub KeyMovement_DetectNeedsExactMatch()

        Dim oldText =
            "[Delta *]" & vbCrLf & "LangSecRef=3021" & vbCrLf & "Detect1=HKCU\Software\Delta" & vbCrLf &
            "Detect2=HKLM\Software\Microsoft\Windows\CurrentVersion\Uninstall\Delta" & vbCrLf &
            "FileKey1=%AppData%\Delta|*.dat" & vbCrLf

        Dim newText =
            "[Delta *]" & vbCrLf & "LangSecRef=3021" & vbCrLf & "Detect=HKCU\Software\Delta" & vbCrLf &
            "FileKey1=%AppData%\Delta|*.dat" & vbCrLf & vbCrLf &
            "[Zeta *]" & vbCrLf & "LangSecRef=3025" & vbCrLf & "Detect=HKLM\Software\Microsoft\Windows" & vbCrLf &
            "FileKey1=%WinDir%\INF|INFCACHE.1" & vbCrLf

        RunDiff(oldText, newText)

        Assert.AreEqual(0, Diff.MostRecentDiffOutcome.MovedKeys)
        Assert.AreEqual(1, Diff.MostRecentDiffOutcome.RemovedKeys)

    End Sub

    ''' <summary>
    ''' When the merged sources' combined keys equal the target's, the demoted rename's name change
    ''' must not show up inside the merger output
    ''' </summary>
    <TestMethod>
    Public Sub DemotedRename_PrintsNoStrayNameChange()

        Dim newText =
            "[Gamma *]" & vbCrLf & "LangSecRef=3021" & vbCrLf & "DetectFile1=%AppData%\Bar" & vbCrLf &
            "DetectFile2=%AppData%\Foo" & vbCrLf & "FileKey1=%AppData%\Foo|*.log" & vbCrLf

        Dim output = RunDiff(ContestedRenameOld, newText)

        Assert.AreEqual(2, Diff.MostRecentDiffOutcome.MergedEntries)
        StringAssert.Contains(output, "Gamma * has been added")
        StringAssert.Contains(output, "This entry contains merged keys from 2 removed entries:")
        Assert.IsFalse(output.Contains("Entry Name has been modified"), output)

    End Sub

    ''' <summary>
    ''' A merger into an added entry with nothing novel still counts the keys it carried over.
    ''' Neither source matches every FileKey, so no rename is involved.
    ''' </summary>
    <TestMethod>
    Public Sub MergerIntoAddedEntry_CountsCarriedOverKeys()

        Dim oldText =
            "[Alpha *]" & vbCrLf & "LangSecRef=3021" & vbCrLf & "DetectFile=%AppData%\Foo" & vbCrLf &
            "FileKey1=%AppData%\Foo|*.log" & vbCrLf & vbCrLf &
            "[Beta *]" & vbCrLf & "LangSecRef=3021" & vbCrLf & "DetectFile=%AppData%\Bar" & vbCrLf &
            "FileKey1=%AppData%\Bar|*.log" & vbCrLf

        Dim newText =
            "[Gamma *]" & vbCrLf & "LangSecRef=3021" & vbCrLf & "DetectFile1=%AppData%\Bar" & vbCrLf &
            "DetectFile2=%AppData%\Foo" & vbCrLf & "FileKey1=%AppData%\Bar|*.log" & vbCrLf &
            "FileKey2=%AppData%\Foo|*.log" & vbCrLf

        Dim output = RunDiff(oldText, newText)

        Assert.AreEqual(2, Diff.MostRecentDiffOutcome.MergedEntries)
        StringAssert.Contains(output, "1 entry contains 5 keys carried over unchanged from merged sources")
        StringAssert.Contains(output, "1 entry consolidates content from")

    End Sub

    ''' <summary>
    ''' Split is divided between two added entries. Neither lists a key as dropped: each names its
    ''' sibling as the home of the other half, and points to Split's own section for the Warning
    ''' neither kept, which is listed there once.
    ''' </summary>
    <TestMethod>
    Public Sub SplitEntry_SiblingKeysAreNotDropped()

        Dim oldText =
            "[Split *]" & vbCrLf & "LangSecRef=3021" & vbCrLf & "DetectFile=%AppData%\Split" & vbCrLf & "Warning=Gone" & vbCrLf &
            "FileKey1=%AppData%\Split\A|*.log" & vbCrLf & "FileKey2=%AppData%\Split\B|*.log" & vbCrLf

        Dim newText =
            "[Split A *]" & vbCrLf & "LangSecRef=3021" & vbCrLf & "DetectFile=%AppData%\Split" & vbCrLf &
            "FileKey1=%AppData%\Split\A|*.log" & vbCrLf & vbCrLf &
            "[Split B *]" & vbCrLf & "LangSecRef=3021" & vbCrLf & "DetectFile=%AppData%\Split" & vbCrLf &
            "FileKey1=%AppData%\Split\B|*.log" & vbCrLf

        Dim output = RunDiff(oldText, newText)

        Assert.AreEqual(0, CountOf(output, "from merged entries not in this entry:"), output)
        Assert.AreEqual(1, CountOf(output, "Warning=Gone"), output)
        Assert.IsTrue(output.IndexOf("1 key not kept by any of these entries:", StringComparison.Ordinal) < output.IndexOf("Warning=Gone", StringComparison.Ordinal), output)
        StringAssert.Contains(output, "Lost 1 Warning")
        StringAssert.Contains(output, "1 other key from merged entries went to [Split B *]")
        StringAssert.Contains(output, "1 other key from merged entries went to [Split A *]")
        Assert.AreEqual(2, CountOf(output, "1 other key that no entry kept is listed under [Split *]"), output)
        Assert.IsFalse(output.Contains("FileKey2=%AppData%\Split\B|*.log"), output)
        Assert.IsFalse(output.Contains("dropped"), output)
        StringAssert.Contains(output, "- 1 split entry lost 1 key that no entry kept")

    End Sub

    ''' <summary>
    ''' The summary counts the old keys a capturing key replaced as the block itemizes them, not
    ''' every source key that resurfaced. Gamma's third FileKey keeps Alpha from being a rename.
    ''' </summary>
    <TestMethod>
    Public Sub MergerIntoAddedEntry_CountsCapturedKeysAsItemized()

        Dim oldText =
            "[Alpha *]" & vbCrLf & "LangSecRef=3021" & vbCrLf & "DetectFile=%AppData%\Foo" & vbCrLf &
            "FileKey1=%AppData%\Foo|*.log" & vbCrLf & "FileKey2=%AppData%\Foo|*.tmp" & vbCrLf & vbCrLf &
            "[Beta *]" & vbCrLf & "LangSecRef=3021" & vbCrLf & "DetectFile=%AppData%\Bar" & vbCrLf &
            "FileKey1=%AppData%\Bar|*.log" & vbCrLf

        Dim newText =
            "[Gamma *]" & vbCrLf & "LangSecRef=3021" & vbCrLf & "DetectFile1=%AppData%\Bar" & vbCrLf &
            "DetectFile2=%AppData%\Foo" & vbCrLf & "FileKey1=%AppData%\Bar|*.log" & vbCrLf &
            "FileKey2=%AppData%\Baz|*.log" & vbCrLf & "FileKey3=%AppData%\Foo|*.*" & vbCrLf

        Dim output = RunDiff(oldText, newText)

        StringAssert.Contains(output, "1 key capturing 2 removed keys")

    End Sub

    ''' <summary>
    ''' Keep absorbs half of Split and loses nothing of its own, so its block is only the line
    ''' naming where the rest of Split went. The block still prints.
    ''' </summary>
    <TestMethod>
    Public Sub SplitIntoModifiedEntry_KeepsBlockWhenEveryDropWentToASibling()

        Dim oldText =
            "[Keep *]" & vbCrLf & "LangSecRef=3021" & vbCrLf & "DetectFile=%AppData%\Keep" & vbCrLf &
            "FileKey1=%AppData%\Keep|*.log" & vbCrLf & vbCrLf &
            "[Split *]" & vbCrLf & "LangSecRef=3021" & vbCrLf & "DetectFile=%AppData%\Split" & vbCrLf &
            "FileKey1=%AppData%\Split\A|*.log" & vbCrLf & "FileKey2=%AppData%\Split\B|*.log" & vbCrLf

        Dim newText =
            "[Keep *]" & vbCrLf & "LangSecRef=3021" & vbCrLf & "DetectFile=%AppData%\Keep" & vbCrLf &
            "FileKey1=%AppData%\Keep|*.log" & vbCrLf & "FileKey2=%AppData%\Split\A|*.log" & vbCrLf & vbCrLf &
            "[Split B *]" & vbCrLf & "LangSecRef=3021" & vbCrLf & "DetectFile=%AppData%\Split" & vbCrLf &
            "FileKey1=%AppData%\Split\B|*.log" & vbCrLf

        Dim output = RunDiff(oldText, newText)

        StringAssert.Contains(output, "Keep * has been modified")
        StringAssert.Contains(output, "2 other keys from merged entries went to [Split B *]")

    End Sub

    ''' <summary>
    ''' Both halves of Split keep its <c> *.log </c> wildcard, which covers <c> cache.log </c>, but
    ''' neither block's re-diff can show that as a capture. The key must still appear somewhere
    ''' rather than each block saying it went to the other.
    ''' </summary>
    <TestMethod>
    Public Sub SplitEntry_KeyCoveredByEveryTarget_StillAppears()

        Dim oldText =
            "[Split *]" & vbCrLf & "LangSecRef=3021" & vbCrLf & "DetectFile=%AppData%\Split" & vbCrLf &
            "FileKey1=%AppData%\Split|*.log" & vbCrLf & "FileKey2=%AppData%\Split|cache.log" & vbCrLf &
            "FileKey3=%AppData%\Split\A|*.tmp" & vbCrLf & "FileKey4=%AppData%\Split\B|*.tmp" & vbCrLf

        Dim newText =
            "[Split A *]" & vbCrLf & "LangSecRef=3021" & vbCrLf & "DetectFile=%AppData%\Split" & vbCrLf &
            "FileKey1=%AppData%\Split|*.log" & vbCrLf & "FileKey2=%AppData%\Split\A|*.tmp" & vbCrLf & vbCrLf &
            "[Split B *]" & vbCrLf & "LangSecRef=3021" & vbCrLf & "DetectFile=%AppData%\Split" & vbCrLf &
            "FileKey1=%AppData%\Split|*.log" & vbCrLf & "FileKey2=%AppData%\Split\B|*.tmp" & vbCrLf

        Dim output = RunDiff(oldText, newText)

        StringAssert.Contains(output, "%AppData%\Split|cache.log")
        StringAssert.Contains(output, "1 other key from merged entries went to [Split B *]")

    End Sub

    ''' <summary>
    ''' Gamma's block lists its novel key and the key that moved in from Delta as added, and the
    ''' keys it kept from Alpha and Beta in their own group, matching the summary's two counts.
    ''' </summary>
    <TestMethod>
    Public Sub MergerIntoAddedEntry_GroupsCarriedOverKeysApart()

        Dim oldText =
            "[Alpha *]" & vbCrLf & "LangSecRef=3021" & vbCrLf & "DetectFile=%AppData%\Foo" & vbCrLf &
            "FileKey1=%AppData%\Foo|*.log" & vbCrLf & vbCrLf &
            "[Beta *]" & vbCrLf & "LangSecRef=3021" & vbCrLf & "DetectFile=%AppData%\Bar" & vbCrLf &
            "FileKey1=%AppData%\Bar|*.log" & vbCrLf & vbCrLf &
            "[Delta *]" & vbCrLf & "LangSecRef=3022" & vbCrLf & "DetectFile=%AppData%\Delta" & vbCrLf &
            "FileKey1=%AppData%\Delta|*.dat" & vbCrLf & "FileKey2=%AppData%\Delta\Sync|*" & vbCrLf

        Dim newText =
            "[Delta *]" & vbCrLf & "LangSecRef=3022" & vbCrLf & "DetectFile=%AppData%\Delta" & vbCrLf &
            "FileKey1=%AppData%\Delta|*.dat" & vbCrLf & vbCrLf &
            "[Gamma *]" & vbCrLf & "LangSecRef=3021" & vbCrLf & "DetectFile1=%AppData%\Bar" & vbCrLf &
            "DetectFile2=%AppData%\Foo" & vbCrLf & "FileKey1=%AppData%\Bar|*.log" & vbCrLf &
            "FileKey2=%AppData%\Delta\Sync|*" & vbCrLf & "FileKey3=%AppData%\Foo|*.log" & vbCrLf &
            "FileKey4=%AppData%\Gamma|*.tmp" & vbCrLf

        Dim output = RunDiff(oldText, newText)

        Dim addedHeader = output.IndexOf("2 keys added:", StringComparison.Ordinal)
        Dim novel = output.IndexOf("FileKey4=%AppData%\Gamma|*.tmp (novel)", StringComparison.Ordinal)
        Dim movedIn = output.IndexOf("FileKey2=%AppData%\Delta\Sync|* (from [Delta *])", StringComparison.Ordinal)
        Dim carriedHeader = output.IndexOf("5 keys carried over from merged entries:", StringComparison.Ordinal)
        Dim carried = output.IndexOf("FileKey3=%AppData%\Foo|*.log (from [Alpha *])", StringComparison.Ordinal)

        Assert.IsTrue(addedHeader >= 0 AndAlso novel > addedHeader AndAlso movedIn > addedHeader, output)
        Assert.IsTrue(carriedHeader > novel AndAlso carriedHeader > movedIn AndAlso carried > carriedHeader, output)
        StringAssert.Contains(output, "Carried over 2 FileKeys")
        Assert.IsFalse(output.Contains("added or carried over"), output)
        StringAssert.Contains(output, "1 entry contains 1 novel key")
        StringAssert.Contains(output, "1 entry contains 5 keys carried over unchanged from merged sources")

    End Sub
    ''' <summary>
    ''' Each list in a modified entry's block has a heading with its total, ahead of its count lines
    ''' </summary>
    <TestMethod>
    Public Sub ModifiedEntry_EveryListHasAHeadingWithItsTotal()

        Dim oldText =
            "[Delta *]" & vbCrLf & "LangSecRef=3021" & vbCrLf & "DetectFile=%AppData%\Delta" & vbCrLf &
            "FileKey1=%AppData%\Delta|*.dat" & vbCrLf & "FileKey2=%AppData%\Delta\Logs|*.log" & vbCrLf &
            "FileKey3=%AppData%\Delta\Old|*" & vbCrLf

        Dim newText =
            "[Delta *]" & vbCrLf & "LangSecRef=3021" & vbCrLf & "DetectFile=%AppData%\Delta" & vbCrLf &
            "FileKey1=%AppData%\Delta|*.dat" & vbCrLf & "FileKey2=%AppData%\Delta\Logs|*.log;*.txt" & vbCrLf &
            "FileKey3=%AppData%\Delta\New|*" & vbCrLf & "Warning=Closes Delta first" & vbCrLf

        Dim output = RunDiff(oldText, newText)

        For Each pair In {Tuple.Create("2 keys added:", "Added 1 Warning"),
                          Tuple.Create("1 key removed:", "Removed 1 FileKey"),
                          Tuple.Create("1 key modified:", "Modified 1 FileKey")}

            Dim heading = output.IndexOf(pair.Item1, StringComparison.Ordinal)
            Dim countLine = output.IndexOf(pair.Item2, StringComparison.Ordinal)
            Assert.IsTrue(heading >= 0 AndAlso countLine > heading, $"{pair.Item1}{vbCrLf}{output}")

        Next

    End Sub

    ''' <summary>
    ''' Builds the old and new text of a removed entry merged into a modified one
    ''' </summary>
    '''
    ''' <param name="withNovelKey">
    ''' Indicates whether the merge target also gains a key no source had
    ''' </param>
    '''
    ''' <returns>
    ''' The old text, then the new text
    ''' </returns>
    Private Shared Function MergeIntoModifiedTexts(withNovelKey As Boolean) As Tuple(Of String, String)

        Dim oldText =
            "[Alpha Tool *]" & vbCrLf & "LangSecRef=3021" & vbCrLf & "DetectFile=%ProgramFiles%\Alpha" & vbCrLf &
            "FileKey1=%AppData%\Alpha\Logs|*.log" & vbCrLf & vbCrLf &
            "[Alpha Tool Cache *]" & vbCrLf & "LangSecRef=3021" & vbCrLf & "DetectFile=%ProgramFiles%\Alpha" & vbCrLf &
            "FileKey1=%LocalAppData%\Alpha\Cache|*|RECURSE" & vbCrLf

        Dim newText =
            "[Alpha Tool *]" & vbCrLf & "LangSecRef=3021" & vbCrLf & "DetectFile=%ProgramFiles%\Alpha" & vbCrLf &
            "FileKey1=%AppData%\Alpha\Logs|*.log" & vbCrLf &
            If(withNovelKey, "FileKey2=%AppData%\Alpha\Temp|*.tmp" & vbCrLf, "") &
            $"FileKey{If(withNovelKey, 3, 2)}=%LocalAppData%\Alpha\Cache|*|RECURSE" & vbCrLf

        Return Tuple.Create(oldText, newText)

    End Function

    ''' <summary>
    ''' A modified entry that took in a removed entry's key and gained one of its own: the summary
    ''' counts the one key its merger block lists as added, not the merged key as well
    ''' </summary>
    <TestMethod>
    Public Sub MergerIntoModifiedEntry_SummaryCountsWhatTheBlockLists()

        Dim texts = MergeIntoModifiedTexts(withNovelKey:=True)
        Dim output = RunDiff(texts.Item1, texts.Item2)

        Assert.AreEqual(1, Diff.MostRecentDiffOutcome.MergedEntries)
        Assert.AreEqual(1, Diff.MostRecentDiffOutcome.AddedKeys)
        StringAssert.Contains(output, "1 key added:")
        StringAssert.Contains(output, "FileKey2=%AppData%\Alpha\Temp|*.tmp (novel)")
        StringAssert.Contains(output, "+ 1 added key across 1 entry")
        Assert.IsFalse(output.Contains("2 added keys"), output)

    End Sub

    ''' <summary>
    ''' A modified entry whose only change is taking in a removed entry's key lists no added key,
    ''' as when it also gains a key of its own. Its block lists the merged key as carried over
    ''' instead, and the keys both entries already had not at all.
    ''' </summary>
    <TestMethod>
    Public Sub PureMergerIntoModifiedEntry_ListsMergedKeyAsCarriedOver()

        Dim texts = MergeIntoModifiedTexts(withNovelKey:=False)
        Dim output = RunDiff(texts.Item1, texts.Item2)

        Assert.AreEqual(1, Diff.MostRecentDiffOutcome.MergedEntries)
        Assert.AreEqual(1, Diff.MostRecentDiffOutcome.ModifiedEntries)
        Assert.AreEqual(0, Diff.MostRecentDiffOutcome.AddedKeys)
        StringAssert.Contains(output, "Alpha Tool * has been modified")
        StringAssert.Contains(output, "1 removed entry:")
        Assert.IsFalse(output.Contains("Added 1 FileKey"), output)
        StringAssert.Contains(output, "1 key carried over from merged entries:")
        StringAssert.Contains(output, "Carried over 1 FileKey")
        StringAssert.Contains(output, "FileKey2=%LocalAppData%\Alpha\Cache|*|RECURSE (from [Alpha Tool Cache *])")
        Assert.IsFalse(output.Contains("Carried over 1 DetectFile"), output)
        StringAssert.Contains(output, "= 1 entry contains 1 key carried over unchanged from merged sources")

    End Sub

    ''' <summary>
    ''' A modified entry that took in a removed entry's key and gained one of its own lists the
    ''' first as carried over and the second as added, and the summary counts each on its own line
    ''' </summary>
    <TestMethod>
    Public Sub MergerIntoModifiedEntry_ListsCarriedOverKeysApartFromAdded()

        Dim texts = MergeIntoModifiedTexts(withNovelKey:=True)
        Dim output = RunDiff(texts.Item1, texts.Item2)

        Dim added = output.IndexOf("1 key added:", StringComparison.Ordinal)
        Dim carried = output.IndexOf("1 key carried over from merged entries:", StringComparison.Ordinal)
        Dim carriedKey = output.IndexOf("FileKey3=%LocalAppData%\Alpha\Cache|*|RECURSE (from [Alpha Tool Cache *])", StringComparison.Ordinal)

        Assert.IsTrue(added >= 0 AndAlso carried > added AndAlso carriedKey > carried, output)
        Assert.AreEqual(1, Diff.MostRecentDiffOutcome.AddedKeys)
        StringAssert.Contains(output, "+ 1 added key across 1 entry")
        StringAssert.Contains(output, "= 1 entry contains 1 key carried over unchanged from merged sources")

    End Sub

    ''' <summary>
    ''' A key carried over from a source is listed even when the entry's updated Detect has the
    ''' same value, since a Detect and a RegKey often share one
    ''' </summary>
    <TestMethod>
    Public Sub MergerIntoModifiedEntry_CarriesOverKeySharingAnUpdatedKeysValue()

        Dim oldText =
            "[Alpha Tool *]" & vbCrLf & "LangSecRef=3021" & vbCrLf & "Detect=HKCU\Software\AlphaSoft\Alpha Tool" & vbCrLf &
            "FileKey1=%AppData%\Alpha\Logs|*.log" & vbCrLf & vbCrLf &
            "[Alpha Tool Cache *]" & vbCrLf & "LangSecRef=3021" & vbCrLf & "DetectFile=%LocalAppData%\Alpha" & vbCrLf &
            "FileKey1=%LocalAppData%\Alpha\Cache|*|RECURSE" & vbCrLf & "RegKey1=HKCU\Software\AlphaSoft\Alpha" & vbCrLf

        Dim newText =
            "[Alpha Tool *]" & vbCrLf & "LangSecRef=3021" & vbCrLf & "Detect=HKCU\Software\AlphaSoft\Alpha" & vbCrLf &
            "FileKey1=%AppData%\Alpha\Logs|*.log" & vbCrLf & "FileKey2=%LocalAppData%\Alpha\Cache|*|RECURSE" & vbCrLf &
            "RegKey1=HKCU\Software\AlphaSoft\Alpha" & vbCrLf

        Dim output = RunDiff(oldText, newText)

        StringAssert.Contains(output, "2 keys carried over from merged entries:")
        StringAssert.Contains(output, "RegKey1=HKCU\Software\AlphaSoft\Alpha (from [Alpha Tool Cache *])")
        StringAssert.Contains(output, "= 1 entry contains 2 keys carried over unchanged from merged sources")

    End Sub

    ''' <summary>
    ''' The summary's modified added and removed key totals equal the sums of the modified
    ''' blocks' list headings, for a plain modified entry and a merger block together. That holds
    ''' only while no key moves into a merger target: its block lists a moved-in key as added,
    ''' and the summary counts it on the moved line instead.
    ''' </summary>
    <TestMethod>
    Public Sub ModifiedTotals_EqualTheBlocksHeadings()

        Dim texts = MergeIntoModifiedTexts(withNovelKey:=True)

        Dim oldText = texts.Item1 & vbCrLf &
            "[Delta *]" & vbCrLf & "LangSecRef=3021" & vbCrLf & "DetectFile=%AppData%\Delta" & vbCrLf &
            "FileKey1=%AppData%\Delta|*.dat" & vbCrLf & "FileKey2=%AppData%\Delta\Old|*.old" & vbCrLf

        Dim newText = texts.Item2 & vbCrLf &
            "[Delta *]" & vbCrLf & "LangSecRef=3021" & vbCrLf & "DetectFile=%AppData%\Delta" & vbCrLf &
            "FileKey1=%AppData%\Delta|*.dat" & vbCrLf & "FileKey2=%AppData%\Delta\Logs|*.log" & vbCrLf &
            "Warning=Closes Delta first" & vbCrLf

        Dim output = RunDiff(oldText, newText)

        Dim sumHeadings = Function(listName As String) RegularExpressions.Regex.Matches(output, $"(\d+) keys? {listName}:").
                                                           Cast(Of RegularExpressions.Match)().
                                                           Sum(Function(m) Integer.Parse(m.Groups(1).Value, Globalization.CultureInfo.InvariantCulture))

        Assert.AreEqual(3, Diff.MostRecentDiffOutcome.AddedKeys, output)
        Assert.AreEqual(sumHeadings("added"), Diff.MostRecentDiffOutcome.AddedKeys, output)
        Assert.AreEqual(sumHeadings("removed"), Diff.MostRecentDiffOutcome.RemovedKeys, output)

    End Sub

End Class
