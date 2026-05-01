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
Imports System.Reflection

''' <summary>
''' Data-driven repair tests for WinappDebug.
''' Adding a test requires only adding matching sections to WinappDebugInputs.ini
''' and WinappDebugOutputs.ini — no code changes needed.
'''
''' Input section conventions:
''' - Must include a <c> Rule= </c> metadata key (stripped before the section reaches WinappDebug):
'''   <c> Rule=N </c> isolates lint rule N; <c> Rule=None </c> disables all rules (unconditional checks only);
'''   <c> Rule=All </c> leaves all rules at their defaults
''' - May include a <c> ScanOnly=True </c> metadata key to assert error detection without applying repairs.
'''   The expected output section must match the (unchanged) input for such tests.
''' - May include an <c> ExpectedErrors= </c> metadata key to assert the initial scan error count
''' - May include a <c> Flavor= </c> metadata key to set the winapp2.ini flavor (default: <c> NonCCleaner </c>)
''' - May include a <c> Group=name </c> metadata key to combine multiple input sections into one
'''   <c> winapp2file2 </c> for testing file-level checks (e.g. entry alphabetization).
'''   All sections sharing the same <c> Group= </c> value are combined into a single test.
'''   Only the FIRST section in a group carries <c> Rule= </c>, <c> ScanOnly= </c>,
'''   <c> ExpectedErrors= </c>, and <c> Flavor= </c>; those keys are ignored on subsequent members.
''' - All other keys must appear in winapp2.ini declaration order
''' - Name should end in <c> *] </c> for normal entries; omit <c> *</c> to test missing-star detection
'''
''' Output section conventions:
''' - Name matches the input section name exactly
''' - Keys represent the expected post-repair state in declaration order
''' - For <c> ScanOnly=True </c> tests the output section should be identical to the input
''' - For grouped tests with entry sorting, output sections must be in the expected post-sort order
''' </summary>
<TestClass()> Public Class WinappDebugDataTests

    Public Property TestContext As TestContext

    ''' <summary>
    ''' Yields one test case per input section (or group of sections sharing a <c> Group= </c> key)
    ''' that has matching output sections.
    ''' Strips all metadata keys from input sections before yielding.
    ''' Skips sections with a missing or unparseable <c> Rule= </c> key, or no matching output section.
    ''' Yields nothing when either data file does not exist.
    ''' </summary>
    Private Shared Iterator Function GetRepairTestCases() As IEnumerable(Of Object())

        Dim inputPath = IO.Path.Combine(Environment.CurrentDirectory, "WinappDebugInputs.ini")
        Dim outputPath = IO.Path.Combine(Environment.CurrentDirectory, "WinappDebugOutputs.ini")
        If Not IO.File.Exists(inputPath) OrElse Not IO.File.Exists(outputPath) Then
            Yield {Nothing, Nothing, Integer.MinValue, -1, winapp2ool.Winapp2ool.WinappFlavor.NonCCleaner, False, Nothing}
            Return
        End If

        Dim inputFile = winapp2ool.iniFile2.FromFile(inputPath)
        Dim outputFile = winapp2ool.iniFile2.FromFile(outputPath)

        ' Group accumulation — keyed by group name, preserving first-occurrence order
        Dim pendingInputSections As New Dictionary(Of String, List(Of winapp2ool.iniSection2))
        Dim pendingExpectedSections As New Dictionary(Of String, List(Of winapp2ool.iniSection2))
        Dim groupOrder As New List(Of String)
        Dim groupRuleIndex As New Dictionary(Of String, Integer)
        Dim groupExpectedErrors As New Dictionary(Of String, Integer)
        Dim groupFlavor As New Dictionary(Of String, winapp2ool.WinappFlavor)
        Dim groupScanOnly As New Dictionary(Of String, Boolean)

        For Each section As winapp2ool.iniSection2 In inputFile

            Dim groupKey = section.Keys.GetKey("Group")
            Dim groupName As String = Nothing
            If groupKey IsNot Nothing Then
                groupName = groupKey.Value.Trim()
                section.Keys.Remove(groupKey)
            End If

            ' Non-first members of a group: skip metadata parsing, just collect
            If groupName IsNot Nothing AndAlso pendingInputSections.ContainsKey(groupName) Then
                pendingInputSections(groupName).Add(section)
                pendingExpectedSections(groupName).Add(outputFile.GetSection(section.Name))
                Continue For
            End If

            ' Parse shared metadata from the first (or only) section of each test
            Dim ruleKey = section.Keys.GetKey("Rule")

            If ruleKey Is Nothing Then
                Yield SingleErrorCase(section, $"Missing Rule= key in '{section.Name}'")
                Continue For
            End If

            Dim ruleIndex As Integer
            Select Case ruleKey.Value.Trim().ToUpperInvariant()
                Case "NONE" : ruleIndex = -1
                Case "ALL" : ruleIndex = -2
                Case Else
                    If Not Integer.TryParse(ruleKey.Value, ruleIndex) Then
                        Yield SingleErrorCase(section, $"Unparseable Rule= value '{ruleKey.Value}' in '{section.Name}'")
                        Continue For
                    End If
            End Select
            section.Keys.Remove(ruleKey)

            Dim errorKey = section.Keys.GetKey("ExpectedErrors")
            Dim expectedErrors As Integer = -1
            If errorKey IsNot Nothing Then
                Integer.TryParse(errorKey.Value, expectedErrors)
                section.Keys.Remove(errorKey)
            End If

            Dim flavor As winapp2ool.WinappFlavor = winapp2ool.WinappFlavor.NonCCleaner
            Dim flavorKey = section.Keys.GetKey("Flavor")
            If flavorKey IsNot Nothing Then
                [Enum].TryParse(flavorKey.Value.Trim(), True, flavor)
                section.Keys.Remove(flavorKey)
            End If

            Dim scanOnly As Boolean = False
            Dim scanOnlyKey = section.Keys.GetKey("ScanOnly")
            If scanOnlyKey IsNot Nothing Then
                Boolean.TryParse(scanOnlyKey.Value, scanOnly)
                section.Keys.Remove(scanOnlyKey)
            End If

            Dim expectedSection = outputFile.GetSection(section.Name)

            If groupName IsNot Nothing Then

                ' First section of a new group
                If expectedSection Is Nothing Then
                    Yield SingleErrorCase(section, $"No matching output section for '{section.Name}'")
                    Continue For
                End If

                pendingInputSections.Add(groupName, New List(Of winapp2ool.iniSection2) From {section})
                pendingExpectedSections.Add(groupName, New List(Of winapp2ool.iniSection2) From {expectedSection})
                groupOrder.Add(groupName)
                groupRuleIndex.Add(groupName, ruleIndex)
                groupExpectedErrors.Add(groupName, expectedErrors)
                groupFlavor.Add(groupName, flavor)
                groupScanOnly.Add(groupName, scanOnly)

            Else

                ' Ungrouped — yield immediately
                If expectedSection Is Nothing Then
                    Yield SingleErrorCase(section, $"No matching output section for '{section.Name}'")
                    Continue For
                End If

                Yield {New winapp2ool.iniSection2() {section},
                       New winapp2ool.iniSection2() {expectedSection},
                       ruleIndex, expectedErrors, flavor, scanOnly, Nothing}

            End If

        Next

        ' Yield accumulated group tests
        For Each gName In groupOrder

            Dim secs = pendingInputSections(gName).ToArray()
            Dim exps = pendingExpectedSections(gName).ToArray()

            Dim missingNames As New List(Of String)
            For i = 0 To secs.Length - 1
                If exps(i) Is Nothing Then missingNames.Add(secs(i).Name)
            Next

            If missingNames.Count > 0 Then
                Yield {secs, Nothing, 0, -1, groupFlavor(gName), False,
                       $"No matching output section(s) for group '{gName}': {String.Join(", ", missingNames)}"}
            Else
                Yield {secs, exps, groupRuleIndex(gName), groupExpectedErrors(gName), groupFlavor(gName), groupScanOnly(gName), Nothing}
            End If

        Next

    End Function

    ''' <summary>Convenience helper for error-case yields where only a single input section is available</summary>
    Private Shared Function SingleErrorCase(section As winapp2ool.iniSection2, msg As String) As Object()
        Return {New winapp2ool.iniSection2() {section}, Nothing, 0, -1, winapp2ool.WinappFlavor.NonCCleaner, False, msg}
    End Function

    ''' <summary>
    ''' Loads both data files and confirms they contain at least one section each <br />
    ''' Placed first alphabetically to eliminate I/O from any of the tests which follow
    ''' </summary>
    <TestMethod()> Public Sub AALoadDataFiles_Success()

        Dim inputPath = IO.Path.Combine(Environment.CurrentDirectory, "WinappDebugInputs.ini")
        Dim outputPath = IO.Path.Combine(Environment.CurrentDirectory, "WinappDebugOutputs.ini")

        If Not IO.File.Exists(inputPath) OrElse Not IO.File.Exists(outputPath) Then
            Assert.Inconclusive("WinappDebugInputs.ini and/or WinappDebugOutputs.ini not present")
            Return
        End If

        Dim inputFile = winapp2ool.iniFile2.FromFile(inputPath)
        Dim outputFile = winapp2ool.iniFile2.FromFile(outputPath)

        Assert.IsTrue(inputFile.Count > 0, "WinappDebugInputs.ini has no sections")
        Assert.IsTrue(outputFile.Count > 0, "WinappDebugOutputs.ini has no sections")

    End Sub

    ''' <summary>Returns the first input section name as the displayed test name in the runner</summary>
    Public Shared Function GetRepairTestDisplayName(methodInfo As MethodInfo, data As Object()) As String
        If data(0) Is Nothing Then Return "No data files"
        Return DirectCast(data(0), winapp2ool.iniSection2())(0).Name
    End Function

    ''' <summary>
    ''' Runs WinappDebug on one or more input sections (combined into a single <c> winapp2file2 </c>)
    ''' with the specified rule selection, optionally applies repairs, and asserts that each
    ''' repaired entry matches its expected output section key-for-key.
    ''' When <paramref name="scanOnly"/> is <c> True </c>, the repair pass is skipped and
    ''' the expected output must match the unmodified input.
    ''' </summary>
    <TestMethod()>
    <DynamicData(NameOf(GetRepairTestCases), DynamicDataSourceType.Method,
                 DynamicDataDisplayName:=NameOf(GetRepairTestDisplayName))>
    Public Sub debug_Repair(sections As winapp2ool.iniSection2(),
                            expectedSections As winapp2ool.iniSection2(),
                            ruleIndex As Integer,
                            expectedErrors As Integer,
                            flavor As winapp2ool.Winapp2ool.WinappFlavor,
                            scanOnly As Boolean,
                            skipReason As String)

        If skipReason IsNot Nothing Then
            Assert.Fail(skipReason)
            Return
        End If

        If ruleIndex = Integer.MinValue Then
            Assert.Inconclusive("Test load failure")
            Return
        End If

        Dim captured As New IO.StringWriter()
        Dim previousOut = Console.Out
        Console.SetOut(captured)

        Try

            setCmdLineArgs(AddressOf winapp2ool.WinappDebug.HandleLintCmdLine, Array.Empty(Of String)(), True)
            winapp2ool.CurrentWinappFlavor = flavor
            applyRuleSelection(ruleIndex, scanOnly)

            Dim testFile = buildTestFile(sections)

            ' First pass: scan without repairs to assert error count
            winapp2ool.WinappDebug.Debug(testFile)

            If expectedErrors >= 0 Then Assert.AreEqual(expectedErrors, winapp2ool.WinappDebug.ErrorsFound, "Wrong error count on initial scan")

            ' Second pass: apply repairs (skipped for ScanOnly tests)
            If Not scanOnly Then
                winapp2ool.lintsettings.RepairSomeErrsFound = True
                winapp2ool.WinappDebug.Debug(testFile)
            End If

            ' Assert each expected section against its matching repaired entry (matched by name)
            For Each expected In expectedSections
                Dim match As winapp2ool.winapp2entry2 = Nothing
                For Each e In testFile.Entries
                    If e.Name.Equals(expected.Name, StringComparison.OrdinalIgnoreCase) Then
                        match = e
                        Exit For
                    End If
                Next
                If match Is Nothing Then
                    Assert.Fail($"No repaired entry found matching expected section '{expected.Name}'")
                    Return
                End If
                Dim repairedKeys = match.ToIniSection().Keys.Select(Function(k) k.ToString()).ToList()
                Dim expectedKeys = expected.Keys.Select(Function(k) k.ToString()).ToList()
                assertKeysEqual(expectedKeys, repairedKeys)
            Next

        Finally

            Console.SetOut(previousOut)
            TestContext.WriteLine(captured.ToString())

        End Try

    End Sub

    Private Shared Sub assertKeysEqual(expected As List(Of String), actual As List(Of String))

        If expected.SequenceEqual(actual) Then Return

        Dim sb As New System.Text.StringBuilder
        sb.AppendLine($"Key count: expected {expected.Count}, got {actual.Count}")
        For i = 0 To Math.Max(expected.Count, actual.Count) - 1
            Dim exp = If(i < expected.Count, expected(i), "<missing>")
            Dim got = If(i < actual.Count, actual(i), "<extra>")
            If exp = got Then
                sb.AppendLine($"  [{i}] {exp}")
            Else
                sb.AppendLine($"* [{i}] expected: {exp}")
                sb.AppendLine($"       actual:   {got}")
            End If
        Next
        Assert.Fail(sb.ToString())

    End Sub

    Private Shared Function buildTestFile(sections As winapp2ool.iniSection2()) As winapp2ool.winapp2file2

        Dim f = winapp2ool.iniFile2.Empty("", "")
        For Each s In sections
            f.AddSection(s)
        Next
        Return New winapp2ool.winapp2file2(f)

    End Function

    ''' <summary>
    ''' Configures the active rule set for a test run.
    ''' <c> Rule=None </c> (-1) disables all rules so only unconditional checks fire.
    ''' <c> Rule=All </c> (-2) leaves all rules at their defaults.
    ''' A non-negative index enables only that rule.
    ''' When <paramref name="scanOnly"/> is <c> True </c>, all repair flags are cleared
    ''' after rule selection so that errors are detected but not applied.
    ''' </summary>
    Private Shared Sub applyRuleSelection(ruleIndex As Integer, scanOnly As Boolean)

        winapp2ool.lintsettings.RepairErrsFound = False
        winapp2ool.lintsettings.RepairSomeErrsFound = False

        Select Case ruleIndex

            Case -2

                ' Leave every rule at its default

            Case -1

                For Each rule In winapp2ool.WinappDebug.Rules : rule.turnOff() : Next

            Case Else

                For i = 0 To winapp2ool.WinappDebug.Rules.Count - 1

                    If i = ruleIndex Then
                        winapp2ool.WinappDebug.Rules(i).turnOn()
                    Else
                        winapp2ool.WinappDebug.Rules(i).turnOff()
                    End If

                Next

        End Select

        If scanOnly Then
            For Each rule In winapp2ool.WinappDebug.Rules
                rule.ShouldRepair = False
            Next
        End If

    End Sub

    Private Sub setDebugStage(args As String(), Optional addHalt As Boolean = False)
        setCmdLineArgs(AddressOf winapp2ool.WinappDebug.HandleLintCmdLine, args, addHalt)
    End Sub

    ''' <summary>Tests CLI default state — no args leaves both files at their initial values and autocorrect off</summary>
    <TestMethod()> Public Sub handleCmdLine_NoInputSuccess()
        setDebugStage(Array.Empty(Of String)(), True)
        Assert.AreEqual(winapp2ool.winappDebugFile1.Dir, winapp2ool.winappDebugFile3.Dir)
        Assert.AreEqual("winapp2.ini", winapp2ool.winappDebugFile1.Name)
        Assert.AreEqual("", winapp2ool.winappDebugFile1.SecondName)
        Assert.AreEqual("winapp2-debugged.ini", winapp2ool.winappDebugFile3.Name)
        Assert.AreEqual("winapp2-debugged.ini", winapp2ool.winappDebugFile3.SecondName)
        Assert.IsFalse(winapp2ool.SaveChanges)
    End Sub

    ''' <summary>Tests that <c> -Nf </c> and <c> -Nd </c> args update the correct file properties</summary>
    <TestMethod()> Public Sub handleCmdLine_ChangeFileParamsSuccess()
        ' Case 1: -1f changes File1 name; File3 resets to its InitName
        setDebugStage({"-1f", "winapp2debugged.ini"}, True)
        Assert.AreEqual("winapp2debugged.ini", winapp2ool.winappDebugFile1.Name)
        Assert.AreEqual(Environment.CurrentDirectory, winapp2ool.winappDebugFile1.Dir)
        Assert.AreEqual("winapp2-debugged.ini", winapp2ool.winappDebugFile3.Name)
        ' Case 2: leading relative subdirectory in -3f splits into Dir suffix and Name
        setDebugStage({"-1f", "winapp2debugged.ini", "-3f", "\subdir\winapp2debugged.ini"}, True)
        Assert.AreEqual("winapp2debugged.ini", winapp2ool.winappDebugFile1.Name)
        Assert.AreEqual(Environment.CurrentDirectory, winapp2ool.winappDebugFile1.Dir)
        Assert.AreEqual("winapp2debugged.ini", winapp2ool.winappDebugFile3.Name)
        Assert.AreEqual(Environment.CurrentDirectory & "\subdir", winapp2ool.winappDebugFile3.Dir)
        ' Case 3: -1d sets directory; -1f sets name independently
        setDebugStage({"-1d", "C:\Test Directory\", "-1f", "winapp2-test2.ini"}, True)
        Assert.AreEqual("C:\Test Directory", winapp2ool.winappDebugFile1.Dir)
        Assert.AreEqual("winapp2-test2.ini", winapp2ool.winappDebugFile1.Name)
    End Sub

    ''' <summary>Tests that <c> -c </c> enables autocorrect</summary>
    <TestMethod()> Public Sub handleCmdLine_EnableAutoCorrectSuccess()
        setDebugStage({"-c"}, True)
        Assert.IsTrue(winapp2ool.SaveChanges)
    End Sub

    ''' <summary>
    ''' Tests that <c> -usedate </c> is consumed from cmdargs.
    ''' Full assertion requires <c> UseCurrentDate </c> to be moved from <c> Private </c>
    ''' in the <c> WinappDebug </c> module to <c> lintsettings </c>.
    ''' </summary>
    <TestMethod()> Public Sub handleCmdLine_EnableUseDateSuccess()
        setDebugStage({"-usedate"}, True)
        Assert.IsFalse(winapp2ool.commandLineHandler.cmdargs.Contains("-usedate"))
    End Sub

End Class
