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
''' Tests for which writers put a flavor on the file they save: Flavorizer and WinappDebug
''' replace it when told one, and keep the input's otherwise. Also covers WinappDebug's
''' <c> -archive </c> and <c> -winapp3 </c> markers, which a file never carries beside a flavor
''' or each other, and Transmute's merge of a marked file into winapp2.ini.
''' </summary>
<TestClass()> Public Class FlavorStampTests

    Private _dir As String

    Private Const Entries As String =
        "[Test Entry A *]" & vbCrLf &
        "LangSecRef=3021" & vbCrLf &
        "Detect=HKCU\Software\TestA" & vbCrLf &
        "FileKey1=%AppData%\TestA|*.log" & vbCrLf

    ''' <summary>Returns a small winapp2.ini whose preamble names <paramref name="flavor"/>, or no flavor</summary>
    ''' <param name="flavor">The text of the <c> ; Flavor: </c> line, or <c> Nothing </c> for none</param>
    Private Shared Function FileText(flavor As String) As String

        Dim flavorLine = If(flavor Is Nothing, "", $"; Flavor: {flavor}" & vbCrLf)
        Return "; Version: 260219" & vbCrLf & "; # of entries: 1" & vbCrLf & flavorLine & vbCrLf & Entries

    End Function

    ''' <summary>Returns a small winapp2.ini whose preamble carries the archive marker</summary>
    Private Shared Function ArchiveText() As String

        Return MarkedText(winapp2ool.winapp2file.ArchiveMarkerLine)

    End Function

    ''' <summary>Returns a small winapp2.ini whose preamble carries the given marker lines</summary>
    ''' <param name="markerLines">The marker lines, in preamble order</param>
    Private Shared Function MarkedText(ParamArray markerLines As String()) As String

        Return "; Version: 260219" & vbCrLf & "; # of entries: 1" & vbCrLf & String.Join("", markerLines.Select(Function(l) l & vbCrLf)) & vbCrLf & Entries

    End Function

    <TestInitialize()> Public Sub MakeScratch()

        _dir = Path.Combine(Path.GetTempPath(), $"winapp2ool-flavorstamp-{Guid.NewGuid():N}")
        Directory.CreateDirectory(_dir)

    End Sub

    <TestCleanup()> Public Sub DeleteScratch()

        If Directory.Exists(_dir) Then Directory.Delete(_dir, recursive:=True)

    End Sub

    ''' <summary>Flavorizes a base file naming <paramref name="baseFlavor"/> and returns the saved text</summary>
    ''' <param name="baseFlavor">The base file's flavor line text, or <c> Nothing </c></param>
    ''' <param name="stampFlavor">The flavor to stamp, or <c> Nothing </c></param>
    ''' <param name="baseText">The base file's whole text, used in place of <paramref name="baseFlavor"/> when given</param>
    Private Function RunFlavorize(baseFlavor As String, stampFlavor As winapp2ool.WinappFlavor?, Optional baseText As String = Nothing) As String

        File.WriteAllText(Path.Combine(_dir, "base.ini"), If(baseText, FileText(baseFlavor)))

        Dim baseFile = winapp2ool.iniFile.FromFile(Path.Combine(_dir, "base.ini"))
        Dim outputFile = winapp2ool.iniFile.Empty(_dir, "out.ini")

        winapp2ool.Flavorize(baseFile, outputFile, New winapp2ool.MenuSection, stampFlavor:=stampFlavor)

        Dim outPath = Path.Combine(_dir, "out.ini")
        Return If(File.Exists(outPath), File.ReadAllText(outPath), Nothing)

    End Function

    ''' <summary>
    ''' A stamped flavor replaces the base file's, as the Tron flavor must replace the
    ''' CCleaner flavor it's built from
    ''' </summary>
    <TestMethod()> Public Sub Flavorize_StampReplacesBaseFlavor()

        Dim text = RunFlavorize("CCleaner", winapp2ool.WinappFlavor.Tron)

        StringAssert.Contains(text, "; Flavor: Tron" & vbCrLf)
        Assert.IsFalse(text.Contains("; Flavor: CCleaner"))

    End Sub

    ''' <summary>Without a stamp, the output keeps the base file's flavor</summary>
    <TestMethod()> Public Sub Flavorize_WithoutStampKeepsBaseFlavor()

        StringAssert.Contains(RunFlavorize("BleachBit", Nothing), "; Flavor: BleachBit" & vbCrLf)

    End Sub

    ''' <summary>
    ''' Runs WinappDebug's command line on a file naming <paramref name="inputFlavor"/>, as if
    ''' a flavor flag had named <paramref name="flagFlavor"/>, and returns the saved text
    ''' </summary>
    ''' <param name="inputFlavor">The input's flavor line text, or <c> Nothing </c></param>
    ''' <param name="flagFlavor">The flavor the command line named, or <c> Nothing </c></param>
    ''' <param name="markArchive">Indicates whether to pass <c> -archive </c></param>
    ''' <param name="inputText">The input's whole text, used in place of <paramref name="inputFlavor"/> when given</param>
    ''' <param name="markWinapp3">Indicates whether to pass <c> -winapp3 </c></param>
    ''' <returns>The saved text, or <c> Nothing </c> when nothing was saved</returns>
    Private Function RunLint(inputFlavor As String, flagFlavor As winapp2ool.WinappFlavor?,
                    Optional markArchive As Boolean = False,
                    Optional inputText As String = Nothing,
                    Optional markWinapp3 As Boolean = False) As String

        Dim savedSuppress = winapp2ool.SuppressOutput
        Dim savedMode = winapp2ool.commandLineHandler.IsCommandLineMode
        Dim savedFlag = winapp2ool.commandLineHandler.FlavorFromCommandLine

        Try

            winapp2ool.SuppressOutput = True
            winapp2ool.commandLineHandler.IsCommandLineMode = True
            winapp2ool.commandLineHandler.FlavorFromCommandLine = flagFlavor

            File.WriteAllText(Path.Combine(_dir, "in.ini"), If(inputText, FileText(inputFlavor)))

            Dim args = {"-1d", _dir, "-1f", "in.ini", "-3d", _dir, "-3f", "out.ini", "-c"}
            If markArchive Then args = args.Concat({"-archive"}).ToArray()
            If markWinapp3 Then args = args.Concat({"-winapp3"}).ToArray()
            setCmdLineArgs(AddressOf winapp2ool.WinappDebug.HandleLintCmdLine, args)

            Dim outPath = Path.Combine(_dir, "out.ini")
            Return If(File.Exists(outPath), File.ReadAllText(outPath), Nothing)

        Finally

            winapp2ool.SuppressOutput = savedSuppress
            winapp2ool.commandLineHandler.IsCommandLineMode = savedMode
            winapp2ool.commandLineHandler.FlavorFromCommandLine = savedFlag
            winapp2ool.InitDefaultLintSettings()

        End Try

    End Function

    ''' <summary>A flavor named on the command line stamps a file that named none, as the base build needs</summary>
    <TestMethod()> Public Sub Lint_CommandLineFlavorStamps()

        StringAssert.Contains(RunLint(Nothing, winapp2ool.WinappFlavor.Base), "; Flavor: Base" & vbCrLf)

    End Sub

    ''' <summary>
    ''' With no flavor on the command line, the lint keeps the file's own flavor, whatever the
    ''' saved flavor setting says
    ''' </summary>
    <TestMethod()> Public Sub Lint_WithoutFlagKeepsFileFlavor()

        Dim text = RunLint("Tron", Nothing)

        StringAssert.Contains(text, "; Flavor: Tron" & vbCrLf)

    End Sub

    ''' <summary>A trimmed file keeps the flavor of the file it was trimmed from</summary>
    <TestMethod()> Public Sub Trim_KeepsFileFlavor()

        Dim savedSuppress = winapp2ool.SuppressOutput

        Try

            winapp2ool.SuppressOutput = True
            File.WriteAllText(Path.Combine(_dir, "winapp2.ini"), FileText("FluentCleaner"))

            winapp2ool.Trim.remoteTrim(New winapp2ool.iniFileChooser(_dir, "winapp2.ini"),
                                       New winapp2ool.iniFileChooser(_dir, "trimmed.ini", mustExist:=False),
                                       d:=False)

            StringAssert.Contains(File.ReadAllText(Path.Combine(_dir, "trimmed.ini")), "; Flavor: FluentCleaner" & vbCrLf)

        Finally

            winapp2ool.SuppressOutput = savedSuppress

        End Try

    End Sub

    ''' <summary>Runs <paramref name="action"/> from a zero exit code and returns the exit code it leaves, restoring the old one</summary>
    ''' <param name="action">The work to run</param>
    Private Shared Function ExitCodeAfter(action As Action) As Integer

        Dim savedExitCode = Environment.ExitCode

        Try

            Environment.ExitCode = 0
            action()
            Return Environment.ExitCode

        Finally

            Environment.ExitCode = savedExitCode

        End Try

    End Function

    ''' <summary><c> -archive </c> marks the saved file as the archive and gives it the archive's header</summary>
    <TestMethod()> Public Sub Lint_ArchiveFlagMarksTheFile()

        Dim text As String = Nothing

        Assert.AreEqual(0, ExitCodeAfter(Sub() text = RunLint(Nothing, Nothing, markArchive:=True)))
        StringAssert.Contains(text, "; # of entries: 1" & vbCrLf & winapp2ool.winapp2file.ArchiveMarkerLine & vbCrLf)
        Assert.IsFalse(text.Contains("; Flavor:"))

    End Sub

    ''' <summary>A lint of a marked file keeps the marker without being told again</summary>
    <TestMethod()> Public Sub Lint_KeepsTheFilesMarker()

        StringAssert.Contains(RunLint(Nothing, Nothing, inputText:=ArchiveText()), winapp2ool.winapp2file.ArchiveMarkerLine & vbCrLf)

    End Sub

    ''' <summary><c> -archive </c> beside a flavor flag saves nothing and fails the run</summary>
    <TestMethod()> Public Sub Lint_ArchiveWithFlavorFlag_IsRefused()

        Dim text As String = "unset"

        Assert.AreEqual(1, ExitCodeAfter(Sub() text = RunLint(Nothing, winapp2ool.WinappFlavor.Base, markArchive:=True)))
        Assert.IsNull(text)

    End Sub

    ''' <summary>A flavor-flagged lint of a file that carries the marker saves nothing and fails the run</summary>
    <TestMethod()> Public Sub Lint_FlavorFlagOnMarkedFile_IsRefused()

        Dim text As String = "unset"

        Assert.AreEqual(1, ExitCodeAfter(Sub() text = RunLint(Nothing, winapp2ool.WinappFlavor.CCleaner, inputText:=ArchiveText())))
        Assert.IsNull(text)

    End Sub

    ''' <summary>A flavor run never gains the marker, from the lint or from Flavorizer</summary>
    <TestMethod()> Public Sub FlavorRuns_NeverGainTheMarker()

        For Each marker In {"; Archive:", "; Winapp3:"}

            Assert.IsFalse(RunLint(Nothing, winapp2ool.WinappFlavor.Base).Contains(marker))
            Assert.IsFalse(RunFlavorize(Nothing, winapp2ool.WinappFlavor.Tron).Contains(marker))

        Next

    End Sub

    ''' <summary>Flavorizing a marked file with a flavor stamp saves nothing and fails the run</summary>
    <TestMethod()> Public Sub Flavorize_StampOnMarkedFile_IsRefused()

        Dim text As String = "unset"

        Assert.AreEqual(1, ExitCodeAfter(Sub() text = RunFlavorize(Nothing, winapp2ool.WinappFlavor.Tron, baseText:=ArchiveText())))
        Assert.IsNull(text)

    End Sub

    ''' <summary><c> -winapp3 </c> marks the saved file as winapp3.ini and gives it winapp3.ini's header</summary>
    <TestMethod()> Public Sub Lint_Winapp3FlagMarksTheFile()

        Dim text As String = Nothing

        Assert.AreEqual(0, ExitCodeAfter(Sub() text = RunLint(Nothing, Nothing, markWinapp3:=True)))
        StringAssert.Contains(text, "; # of entries: 1" & vbCrLf & winapp2ool.winapp2file.Winapp3MarkerLine & vbCrLf)
        Assert.IsFalse(text.Contains("; Flavor:"))
        Assert.IsFalse(text.Contains("; Archive:"))

    End Sub

    ''' <summary>A lint of a file marked as winapp3.ini keeps the marker without being told again, and a matching flag is fine</summary>
    <TestMethod()> Public Sub Lint_KeepsTheWinapp3Marker()

        Dim marked = MarkedText(winapp2ool.winapp2file.Winapp3MarkerLine)
        Dim text As String = Nothing

        StringAssert.Contains(RunLint(Nothing, Nothing, inputText:=marked), winapp2ool.winapp2file.Winapp3MarkerLine & vbCrLf)
        Assert.AreEqual(0, ExitCodeAfter(Sub() text = RunLint(Nothing, Nothing, inputText:=marked, markWinapp3:=True)))
        StringAssert.Contains(text, winapp2ool.winapp2file.Winapp3MarkerLine & vbCrLf)

    End Sub

    ''' <summary>
    ''' Each of these saves nothing and fails the run: <c> -winapp3 </c> beside <c> -archive </c>
    ''' or a flavor flag, a flavor flag on a file marked as winapp3.ini, a marker flag that
    ''' disagrees with the file's own marker, and a file that carries both markers
    ''' </summary>
    <TestMethod()> Public Sub Lint_Winapp3Conflicts_AreRefused()

        Dim winapp3Text = MarkedText(winapp2ool.winapp2file.Winapp3MarkerLine)
        Dim bothText = MarkedText(winapp2ool.winapp2file.ArchiveMarkerLine, winapp2ool.winapp2file.Winapp3MarkerLine)

        Dim runs As New Dictionary(Of String, Func(Of String)) From {
            {"-winapp3 with -archive", Function() RunLint(Nothing, Nothing, markArchive:=True, markWinapp3:=True)},
            {"-winapp3 with a flavor flag", Function() RunLint(Nothing, winapp2ool.WinappFlavor.Base, markWinapp3:=True)},
            {"a flavor flag on a winapp3.ini", Function() RunLint(Nothing, winapp2ool.WinappFlavor.CCleaner, inputText:=winapp3Text)},
            {"-winapp3 on an archive", Function() RunLint(Nothing, Nothing, inputText:=ArchiveText(), markWinapp3:=True)},
            {"-archive on a winapp3.ini", Function() RunLint(Nothing, Nothing, markArchive:=True, inputText:=winapp3Text)},
            {"both markers", Function() RunLint(Nothing, Nothing, inputText:=bothText)},
            {"both markers, flagged as the first", Function() RunLint(Nothing, Nothing, markArchive:=True, inputText:=bothText)}
        }

        For Each run In runs

            Dim text As String = "unset"
            Assert.AreEqual(1, ExitCodeAfter(Sub() text = run.Value.Invoke()), run.Key)
            Assert.IsNull(text, run.Key)

        Next

    End Sub

    ''' <summary>
    ''' Flavorizing a file marked as winapp3.ini with a flavor stamp, or a file that carries both
    ''' markers at all, saves nothing and fails the run
    ''' </summary>
    <TestMethod()> Public Sub Flavorize_Winapp3Conflicts_AreRefused()

        Dim runs As New Dictionary(Of String, Func(Of String)) From {
            {"stamped winapp3.ini", Function() RunFlavorize(Nothing, winapp2ool.WinappFlavor.Tron, baseText:=MarkedText(winapp2ool.winapp2file.Winapp3MarkerLine))},
            {"both markers", Function() RunFlavorize(Nothing, Nothing, baseText:=MarkedText(winapp2ool.winapp2file.Winapp3MarkerLine, winapp2ool.winapp2file.ArchiveMarkerLine))}
        }

        For Each run In runs

            Dim text As String = "unset"
            Assert.AreEqual(1, ExitCodeAfter(Sub() text = run.Value.Invoke()), run.Key)
            Assert.IsNull(text, run.Key)

        Next

    End Sub

    ''' <summary>
    ''' Merging a marked winapp3.ini into winapp2.ini, as winapp3.ini's header says to, keeps
    ''' winapp2.ini's header and gains no marker
    ''' </summary>
    <TestMethod()> Public Sub Transmute_MergingWinapp3_KeepsTheBaseHeader()

        Dim savedSuppress = winapp2ool.SuppressOutput

        Try

            winapp2ool.SuppressOutput = True
            File.WriteAllText(Path.Combine(_dir, "winapp2.ini"), FileText("Base"))
            File.WriteAllText(Path.Combine(_dir, "winapp3.ini"),
                              MarkedText(winapp2ool.winapp2file.Winapp3MarkerLine).Replace("Test Entry A", "Test Entry W").Replace("TestA", "TestW"))

            Dim baseFile = winapp2ool.iniFile.FromFile(Path.Combine(_dir, "winapp2.ini"))
            Dim sourceFile = winapp2ool.iniFile.FromFile(Path.Combine(_dir, "winapp3.ini"))
            Dim outputFile = winapp2ool.iniFile.Empty(_dir, "winapp2.ini")

            Assert.IsTrue(winapp2ool.RemoteTransmute(baseFile, sourceFile, outputFile, True, New winapp2ool.MenuSection))

            Dim text = File.ReadAllText(Path.Combine(_dir, "winapp2.ini"))
            StringAssert.Contains(text, "; # of entries: 2" & vbCrLf & "; Flavor: Base" & vbCrLf)
            StringAssert.Contains(text, "[Test Entry W *]")
            Assert.IsFalse(text.Contains("; Winapp3:"), text)
            Assert.IsFalse(text.Contains("aggressive"), text)

        Finally

            winapp2ool.SuppressOutput = savedSuppress

        End Try

    End Sub

    ''' <summary>With no flavor anywhere, the lint writes the header that names none</summary>
    <TestMethod()> Public Sub Lint_NoFlavorAnywhereStaysFlavorless()

        Assert.IsFalse(RunLint(Nothing, Nothing).Contains("; Flavor:"))

    End Sub

End Class
