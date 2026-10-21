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
''' replace it when told one, and keep the input's otherwise
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
    Private Function RunFlavorize(baseFlavor As String, stampFlavor As winapp2ool.WinappFlavor?) As String

        File.WriteAllText(Path.Combine(_dir, "base.ini"), FileText(baseFlavor))

        Dim baseFile = winapp2ool.iniFile.FromFile(Path.Combine(_dir, "base.ini"))
        Dim outputFile = winapp2ool.iniFile.Empty(_dir, "out.ini")

        winapp2ool.Flavorize(baseFile, outputFile, New winapp2ool.MenuSection, stampFlavor:=stampFlavor)

        Return File.ReadAllText(Path.Combine(_dir, "out.ini"))

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
    Private Function RunLint(inputFlavor As String, flagFlavor As winapp2ool.WinappFlavor?) As String

        Dim savedSuppress = winapp2ool.SuppressOutput
        Dim savedMode = winapp2ool.commandLineHandler.IsCommandLineMode
        Dim savedFlag = winapp2ool.commandLineHandler.FlavorFromCommandLine

        Try

            winapp2ool.SuppressOutput = True
            winapp2ool.commandLineHandler.IsCommandLineMode = True
            winapp2ool.commandLineHandler.FlavorFromCommandLine = flagFlavor

            File.WriteAllText(Path.Combine(_dir, "in.ini"), FileText(inputFlavor))

            setCmdLineArgs(AddressOf winapp2ool.WinappDebug.HandleLintCmdLine,
                           {"-1d", _dir, "-1f", "in.ini", "-3d", _dir, "-3f", "out.ini", "-c"})

            Return File.ReadAllText(Path.Combine(_dir, "out.ini"))

        Finally

            winapp2ool.SuppressOutput = savedSuppress
            winapp2ool.commandLineHandler.IsCommandLineMode = savedMode
            winapp2ool.commandLineHandler.FlavorFromCommandLine = savedFlag
            winapp2ool.InitDefaultLintSettings()

        End Try

    End Function

    ''' <summary>A flavor named on the command line stamps a file that named none, as the base build needs</summary>
    <TestMethod()> Public Sub Lint_CommandLineFlavorStamps()

        StringAssert.Contains(RunLint(Nothing, winapp2ool.WinappFlavor.NonCCleaner), "; Flavor: Base" & vbCrLf)

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

    ''' <summary>With no flavor anywhere, the lint writes the header that names none</summary>
    <TestMethod()> Public Sub Lint_NoFlavorAnywhereStaysFlavorless()

        Assert.IsFalse(RunLint(Nothing, Nothing).Contains("; Flavor:"))

    End Sub

End Class
