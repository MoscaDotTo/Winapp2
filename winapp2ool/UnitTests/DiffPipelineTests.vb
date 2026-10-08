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
        StringAssert.Contains(output, "+ 1 into 1 added entry")

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
        StringAssert.Contains(output, "~ 1 into 1 modified entry")

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
        StringAssert.Contains(output, "& 1 into 1 renamed entry")
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
        StringAssert.Contains(output, "Gamma * has been added (consolidating 2 removed entries)")
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
        StringAssert.Contains(output, "1 entry contain 5 keys carried over unchanged from merged sources")

    End Sub

End Class
