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
    ''' A demoted rename must not leave key records behind: Gamma's Warning comes from no merged
    ''' source, so Delta losing the same Warning is a removal, not a move into Gamma
    ''' </summary>
    <TestMethod>
    Public Sub DemotedRename_LeavesNoKeyMovement()

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

        RunDiff(oldText, newText)

        Dim outcome = Diff.MostRecentDiffOutcome
        Assert.AreEqual(2, outcome.MergedEntries)
        Assert.AreEqual(0, outcome.RenamedEntries)
        Assert.AreEqual(0, outcome.MovedKeys)
        Assert.AreEqual(1, outcome.RemovedKeys)

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
