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
Imports Microsoft.Win32

''' <summary>
''' Tests <see cref="winapp2ool.Trim"/>'s detection against a scratch folder and a scratch
''' registry key, and that a trim with no usable input writes nothing and fails the run
''' </summary>
<TestClass()> Public Class TrimDetectionTests

    ''' <summary>Where the VirtualStore tests create their registry key, under HKCU</summary>
    Private Const VirtualStoreRoot As String = "Software\Classes\VirtualStore\MACHINE\SOFTWARE"

    Private _dir As String
    Private _vsKeyName As String

    ''' <summary>Creates a scratch folder and picks a unique VirtualStore key name</summary>
    <TestInitialize()> Public Sub MakeScratch()

        _dir = Path.Combine(Path.GetTempPath(), $"winapp2ool-trimdetect-{Guid.NewGuid():N}")
        Directory.CreateDirectory(_dir)
        _vsKeyName = $"winapp2ool-unit-test-vs-{Guid.NewGuid():N}"

    End Sub

    ''' <summary>Deletes the scratch folder and the VirtualStore key, if a test made one</summary>
    <TestCleanup()> Public Sub DeleteScratch()

        Directory.Delete(_dir, recursive:=True)
        Registry.CurrentUser.DeleteSubKeyTree($"{VirtualStoreRoot}\{_vsKeyName}", throwOnMissingSubKey:=False)

    End Sub

    ''' <summary>Parses <paramref name="text"/>, trims it with no lists and returns the result</summary>
    ''' <param name="text">Literal winapp2.ini text</param>
    Private Shared Function TrimText(text As String) As winapp2ool.winapp2file

        Dim ini As winapp2ool.iniFile

        Using reader As New StreamReader(New MemoryStream(Encoding.UTF8.GetBytes(text)))
            ini = winapp2ool.iniFile.FromStream(reader, "", "winapp2.ini")
        End Using

        Dim wa2 As New winapp2ool.winapp2file(ini)
        winapp2ool.Trim.trimFile(wa2)

        Return wa2

    End Function

    ''' <summary>Returns whether an entry detected by <paramref name="detectFile"/> survives a trim</summary>
    ''' <param name="detectFile">The entry's DetectFile value</param>
    Private Shared Function IsKept(detectFile As String) As Boolean

        Dim text = "[Probe *]" & vbCrLf &
                   "LangSecRef=3021" & vbCrLf &
                   $"DetectFile={detectFile}" & vbCrLf &
                   "FileKey1=%AppData%\winapp2ool-unit-test-probe|*.*" & vbCrLf

        Return TrimText(text).Count = 1

    End Function

    ''' <summary>
    ''' A wildcard in the last DetectFile segment matches a file, not only a folder
    ''' </summary>
    <TestMethod()> Public Sub WildcardLastSegment_MatchesAFile()

        File.WriteAllText(Path.Combine(_dir, "app.exe"), "")

        Assert.IsTrue(IsKept($"{_dir}\*.exe"))
        Assert.IsFalse(IsKept($"{_dir}\*.dll"))

    End Sub

    ''' <summary>
    ''' A file named after a wildcard folder segment is found
    ''' </summary>
    <TestMethod()> Public Sub FileAfterWildcardFolder_IsFound()

        Directory.CreateDirectory(Path.Combine(_dir, "App 2"))
        File.WriteAllText(Path.Combine(_dir, "App 2", "app.exe"), "")

        Assert.IsTrue(IsKept($"{_dir}\App*\app.exe"))
        Assert.IsFalse(IsKept($"{_dir}\App*\other.exe"))

    End Sub

    ''' <summary>
    ''' A wildcard in the segment after the drive letter is matched at the drive's root, not in
    ''' the current folder on that drive
    ''' </summary>
    <TestMethod()> Public Sub WildcardAfterDriveLetter_MatchesAtTheRoot()

        File.WriteAllText(Path.Combine(_dir, "app.exe"), "")

        ' C:\Users\...\scratch becomes C:\Users*\...\scratch\app.exe
        Dim root = Path.GetPathRoot(_dir)
        Dim rest = _dir.Substring(root.Length)
        Dim first = rest.Split("\"c)(0)

        Assert.IsTrue(IsKept($"{root}{first}*{rest.Substring(first.Length)}\app.exe"))

    End Sub

    ''' <summary>
    ''' A folder segment's wildcard still matches folders only, so a file doesn't stand in for one
    ''' </summary>
    <TestMethod()> Public Sub WildcardFolderSegment_SkipsFiles()

        File.WriteAllText(Path.Combine(_dir, "App 2"), "")

        Assert.IsFalse(IsKept($"{_dir}\App*\app.exe"))

    End Sub

    ''' <summary>
    ''' A malformed environment variable keeps the entry without waiting for a key press, which
    ''' would throw here, since a test has no console to read
    ''' </summary>
    <TestMethod()> Public Sub MalformedVariable_KeepsEntryWithoutPrompting()

        Assert.IsTrue(IsKept("%AppData\winapp2ool-unit-test-malformed"))

    End Sub

    ''' <summary>Trims an always-detected entry holding <paramref name="regKey"/> and returns its RegKey values</summary>
    ''' <param name="regKey">The entry's RegKey1 value</param>
    Private Shared Function RegKeysAfterTrim(regKey As String) As List(Of String)

        Dim text = "[Probe *]" & vbCrLf &
                   "LangSecRef=3021" & vbCrLf &
                   "Detect=HKCU\Software" & vbCrLf &
                   $"RegKey1={regKey}" & vbCrLf

        Return TrimText(text).Entries(0).RegKeys.Select(Function(k) k.Value).ToList()

    End Function

    ''' <summary>
    ''' A RegKey naming a value gets a VirtualStore counterpart when the key holding the value
    ''' exists there
    ''' </summary>
    <TestMethod()> Public Sub VirtualStore_RegKeyWithValueName_GetsCounterpart()

        Registry.CurrentUser.CreateSubKey($"{VirtualStoreRoot}\{_vsKeyName}").Dispose()

        Dim values = RegKeysAfterTrim($"HKLM\Software\{_vsKeyName}|Setting")

        CollectionAssert.Contains(values, $"HKCU\{VirtualStoreRoot}\{_vsKeyName}|Setting")
        Assert.AreEqual(2, values.Count)

    End Sub

    ''' <summary>
    ''' A RegKey spelling <c> HKLM\SOFTWARE </c> in capitals still gets its VirtualStore counterpart
    ''' </summary>
    <TestMethod()> Public Sub VirtualStore_MatchIgnoresCase()

        Registry.CurrentUser.CreateSubKey($"{VirtualStoreRoot}\{_vsKeyName}").Dispose()

        Dim values = RegKeysAfterTrim($"HKLM\SOFTWARE\{_vsKeyName}")

        CollectionAssert.Contains(values, $"HKCU\{VirtualStoreRoot}\{_vsKeyName}")
        Assert.AreEqual(2, values.Count)

    End Sub

    ''' <summary>
    ''' A download asked for in offline mode writes no output and fails the run
    ''' </summary>
    <TestMethod()> Public Sub OfflineDownload_WritesNothingAndFails()

        Dim savedOffline = winapp2ool.Winapp2ool.isOffline
        Dim savedExitCode = Environment.ExitCode

        Try

            winapp2ool.Winapp2ool.isOffline = True
            Environment.ExitCode = 0

            winapp2ool.Trim.remoteTrim(New winapp2ool.iniFileChooser(_dir, "winapp2.ini", mustExist:=False),
                                       New winapp2ool.iniFileChooser(_dir, "trimmed.ini", mustExist:=False),
                                       d:=True)

            Assert.IsFalse(File.Exists(Path.Combine(_dir, "trimmed.ini")))
            Assert.AreEqual(1, Environment.ExitCode)

        Finally

            winapp2ool.Winapp2ool.isOffline = savedOffline
            Environment.ExitCode = savedExitCode

        End Try

    End Sub

    ''' <summary>
    ''' Trims an empty input file with <c> IsCommandLineMode </c> set to <paramref name="commandLine"/>
    ''' and returns the exit code it leaves, asserting that nothing was written
    ''' </summary>
    ''' <param name="commandLine">Indicates whether the trim runs as a command line run</param>
    Private Function ExitCodeForEmptyInput(commandLine As Boolean) As Integer

        Dim savedExitCode = Environment.ExitCode
        Dim savedMode = winapp2ool.commandLineHandler.IsCommandLineMode

        Try

            File.WriteAllText(Path.Combine(_dir, "winapp2.ini"), "")
            Environment.ExitCode = 0
            winapp2ool.commandLineHandler.IsCommandLineMode = commandLine

            winapp2ool.Trim.remoteTrim(New winapp2ool.iniFileChooser(_dir, "winapp2.ini", mustExist:=False),
                                       New winapp2ool.iniFileChooser(_dir, "trimmed.ini", mustExist:=False),
                                       d:=False)

            Assert.IsFalse(File.Exists(Path.Combine(_dir, "trimmed.ini")))
            Return Environment.ExitCode

        Finally

            Environment.ExitCode = savedExitCode
            winapp2ool.commandLineHandler.IsCommandLineMode = savedMode

        End Try

    End Function

    ''' <summary>
    ''' On the command line, an input file with no entries writes no output and fails the run
    ''' </summary>
    <TestMethod()> Public Sub EmptyInput_CommandLine_WritesNothingAndFails()

        Assert.AreEqual(1, ExitCodeForEmptyInput(commandLine:=True))

    End Sub

    ''' <summary>
    ''' In the menu, where a cancelled File Chooser ends the same way, an empty input writes no
    ''' output and leaves the exit code alone
    ''' </summary>
    <TestMethod()> Public Sub EmptyInput_Menu_WritesNothingAndDoesNotFail()

        Assert.AreEqual(0, ExitCodeForEmptyInput(commandLine:=False))

    End Sub

End Class
