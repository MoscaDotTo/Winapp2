'    Copyright (C) 2018-2021 Hazel Ward
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

''' <summary>
''' Unit and integration tests for winapp2ool's WinappDebug module
''' Test naming convention: methodName_testState_expectedBehavior
''' </summary>
<TestClass()> Public Class WinappDebugTests

    Private Property WDUTFile1 As New winapp2ool.iniFile(Environment.CurrentDirectory, "WinappDebugUnitTests.ini")

    ''' <summary>Initializes WinappDebug with provided commandline args</summary>
    ''' <param name="args">An array of args to pass to WinappDebug</param>
    ''' <param name="addHalt">Optional Boolean specifying whether or not the halting flag should be added to the args</param>
    Private Sub setDebugStage(args As String(), Optional addHalt As Boolean = False)
        setCmdLineArgs(AddressOf winapp2ool.WinappDebug.HandleLintCmdLine, args, addHalt)
    End Sub

    ''' <summary>Tests the commandline handling for WinappDebug to ensure success under no input conditions</summary>
    <TestMethod()> Public Sub handleCmdLine_NoInputSuccess()
        ' Test case: Do nothing, expect our default values
        setDebugStage(Array.Empty(Of String)(), True)
        Assert.AreEqual(winapp2ool.winappDebugFile1.Path, winapp2ool.winappDebugFile3.Path)
        Assert.AreEqual("winapp2-debugged.ini", winapp2ool.winappDebugFile3.SecondName)
        Assert.AreNotEqual(winapp2ool.winappDebugFile1.SecondName, winapp2ool.winappDebugFile3.SecondName)
    End Sub

    ''' <summary>Tests the commandline handling for WinappDebug to ensure success when changing file directory or name parameters</summary>
    <TestMethod()> Public Sub handleCmdLine_ChangeFileParamsSuccess()
        ' First test case: Change only the first file name parameter
        setDebugStage({"-1f", "winapp2debugged.ini"}, True)
        Assert.AreEqual("winapp2debugged.ini", winapp2ool.winappDebugFile1.Name)
        Assert.AreEqual($"{Environment.CurrentDirectory}\{winapp2ool.winappDebugFile1.Name}", winapp2ool.winappDebugFile1.Path)
        Assert.AreEqual("winapp2.ini", winapp2ool.winappDebugFile3.Name)
        ' Second test case: Change the first and third file name parameters, also tests setting of a subdirectory through the file command
        setDebugStage({"-1f", "winapp2debugged.ini", "-3f", "\subdir\winapp2debugged.ini"}, True)
        Assert.AreEqual("winapp2debugged.ini", winapp2ool.winappDebugFile1.Name)
        Assert.AreEqual(Environment.CurrentDirectory, winapp2ool.winappDebugFile1.Dir)
        Assert.AreEqual("winapp2debugged.ini", winapp2ool.winappDebugFile3.Name)
        Assert.AreEqual(Environment.CurrentDirectory & "\subdir", winapp2ool.winappDebugFile3.Dir)
        ' Third test case: Change directory and name parameters of a single file
        setDebugStage({"-1d", "C:\Test Directory\", "-1f", "winapp2-test2.ini"}, True)
        Assert.AreEqual("C:\Test Directory", winapp2ool.winappDebugFile1.Dir)
        Assert.AreEqual("winapp2-test2.ini", winapp2ool.winappDebugFile1.Name)
    End Sub

    ''' <summary>
    ''' Tests the commandline handling for WinappDebug to ensure success when enabling autocorrect
    ''' </summary>
    <TestMethod()> Public Sub handleCmdLine_EnableAutoCorrectSuccess()
        ' Test case: Enable autocorrect 
        setDebugStage({"-c"}, True)
        Assert.AreEqual(winapp2ool.SaveChanges, True)
    End Sub

    ' Tests below this point until the marked point test that individual scans and repairs work within WinappDebug's debug method along happy paths

    ''' <summary>Toggles off all the tests except a specific one</summary>
    ''' <param name="lintRuleIndex">The index of the rule to be enabled</param>
    Private Sub disableAllTestsExcept(lintRuleIndex As Integer)
        ' Don't correct all formatting
        winapp2ool.RepairErrsFound = False
        winapp2ool.RepairSomeErrsFound = False
        For i = 0 To winapp2ool.WinappDebug.Rules.Count - 1
            If Not i = lintRuleIndex Then
                ' Turn off all rules by default 
                winapp2ool.WinappDebug.Rules(i).turnOff()
            Else
                ' Enable the rule we want
                winapp2ool.WinappDebug.Rules(i).turnOn()
            End If
        Next
    End Sub

    ''' <summary>Returns a winapp2file object containing the requested test</summary>
    ''' <param name="testNum">The test index to return from the unit test file</param>
    Private Function getSingleTestFile(testNum As Integer) As winapp2ool.winapp2file
        If Not WDUTFile1.Sections.Count > 0 Then WDUTFile1.validate()
        Dim testSection = WDUTFile1.Sections.Values(testNum)
        Dim testFile As New winapp2ool.iniFile
        testFile.Sections.Add(testSection.Name, testSection)
        Return New winapp2ool.winapp2file(testFile)
    End Function

    ''' <summary>Tests the debug function in winapp2ool using tests from WinappDebugUnitTests.ini</summary>
    ''' <param name="testNum">The test number to request from the file</param>
    ''' <param name="expectedErrsWithoutRepair">The expected number of errors to be found on the first run</param>
    ''' <param name="expectedErrsWithRepair">The expected number of errors to be found after repairs are run</param>
    Public Function debug_ErrorFindAndRepair_Success(testNum As Integer, expectedErrsWithoutRepair As Integer, expectedErrsWithRepair As Integer, lintRuleIndex As Integer) As winapp2ool.winapp2entry
        ' Initalize the default state of the module
        setDebugStage(Array.Empty(Of String)(), True)
        ' Disable all the lint rules we're not currently testing
        disableAllTestsExcept(lintRuleIndex)
        Dim test As winapp2ool.winapp2file = getSingleTestFile(testNum)
        ' Confirm the errors are found without autocorrect on
        winapp2ool.WinappDebug.debug(test)
        Assert.AreEqual(expectedErrsWithoutRepair, winapp2ool.WinappDebug.ErrorsFound)
        ' Enable repairs
        winapp2ool.WinappDebug.RepairSomeErrsFound = True
        winapp2ool.WinappDebug.debug(test)
        ' Confirm the errors are still found (ie. not erroneously corrected during the first test)
        Assert.AreEqual(expectedErrsWithoutRepair, winapp2ool.WinappDebug.ErrorsFound)
        winapp2ool.WinappDebug.debug(test)
        ' Confirm fixable errors are fixed
        Assert.AreEqual(expectedErrsWithRepair, winapp2ool.WinappDebug.ErrorsFound)
        ' Return the entry for further assessment
        Return New winapp2ool.winapp2entry(test.EntrySections.Last.Sections.Values.First)
    End Function

    '''<summary>Confirms that the unit test file was initialzed correctly</summary>
    '''<remarks>Test is named this way to make sure it is first on the list of tests, and prevent the overhead of loading the test file
    '''from being reflected in other tests runtimes</remarks>
    <TestMethod> Public Sub debug_AAGetTestFile_success()
        Dim init = getSingleTestFile(0)
        Assert.IsTrue(WDUTFile1.Sections.Count = 12)
    End Sub

    ''' <summary>Runs tests to ensure that keys with duplicate values are detected and removed</summary>
    <TestMethod> Public Sub debug_DuplicateKeyValue_FindAndRepair_Success()
        ' Disable all the tests except the duplicate checks
        Dim testOutput = debug_ErrorFindAndRepair_Success(0, 12, 0, 7)
        ' We expect that the returned object should have 1 key in each keylist except the last
        For i = 0 To testOutput.KeyListList.Count - 2
            Dim lst = testOutput.KeyListList(i)
            Assert.AreEqual(1, lst.KeyCount)
        Next
        ' The last keylist (the error keys) should be empty
        Assert.AreEqual(0, testOutput.KeyListList.Last.KeyCount)
    End Sub

    ''' <summary>Runs tests to ensure that key with names that are improperly numbered are detected and repaired</summary>
    <TestMethod> Public Sub debug_KeyNumberingError_FindAndRepair_Sucess()
        Dim testOutput = debug_ErrorFindAndRepair_Success(1, 8, 0, 2)
        For i = 0 To testOutput.KeyListList.Count - 2
            Dim lst = testOutput.KeyListList(i)
            Select Case lst.KeyType
                Case "Detect", "DetectFile", "ExcludeKey", "FileKey", "RegKey"
                    Dim curKeys = lst.Keys
                    ' We expect 2 keys in each of these lists, having the keyNumbers 1 and 2 respectively
                    Assert.AreEqual(curKeys.First.Name, $"{lst.KeyType}1")
                    Assert.AreEqual(curKeys.Last.Name, $"{lst.KeyType}2")
            End Select
        Next
    End Sub

    ''' <summary>Runs tests to ensure that keys in situations where there should be no numbering are detected and repaired</summary>
    <TestMethod> Public Sub debug_keyNumberingUneededError_FindAndRepair_Success()
        Dim testOutput = debug_ErrorFindAndRepair_Success(2, 8, 0, 8)
        For Each lst In testOutput.KeyListList
            If lst.KeyCount > 0 Then
                Select Case lst.KeyType
                    ' RegKeys, FileKeys, and ExcludeKeys always require a trailing number in the name, even if there is only one
                    Case "RegKey", "FileKey", "ExcludeKey"
                        Assert.AreEqual(lst.Keys.First.KeyType & 1, lst.Keys.First.Name)
                    Case Else
                        Assert.AreEqual(lst.Keys.First.KeyType, lst.Keys.First.Name)
                End Select
            End If
        Next
    End Sub

    ''' <summary>Runs test to ensure that incorrect alphabetization is caught and repaired</summary>
    <TestMethod> Public Sub debug_keyAlphabetization_FindAndRepair_Success()
        Dim testOutput = debug_ErrorFindAndRepair_Success(3, 4, 0, 1)
        Assert.AreEqual("HKCU\Software3", testOutput.Detects.Keys.First.Value)
        Assert.AreEqual("HKCU\Software4", testOutput.Detects.Keys(1).Value)
        Assert.AreEqual("HKCU\Software20", testOutput.Detects.Keys(2).Value)
        Assert.AreEqual("HKCU\Software30", testOutput.Detects.Keys(3).Value)
    End Sub

    <TestMethod> Public Sub debug_forwardSlashes_FindAndRepair_Success()
        Dim testOutput = debug_ErrorFindAndRepair_Success(4, 4, 0, 5)
        For Each lst In testOutput.KeyListList
            ' If the test was successful, none of the keys should have any forward slashes
            If lst.KeyCount > 0 Then Assert.AreEqual(True, Not lst.Keys.First.Value.Contains(CChar("/")))
        Next
    End Sub

    <TestMethod> Public Sub debug_trailingSemiColons_FindAndRepair_Success()
        Dim testOutput = debug_ErrorFindAndRepair_Success(5, 2, 0, 13)
        Assert.IsTrue(Not testOutput.FileKeys.Keys.First.Value.EndsWith(";"))
        Assert.IsTrue(Not testOutput.FileKeys.Keys(1).Value.Contains(";|"))
    End Sub

    <TestMethod> Public Sub debug_multipleBackSlashes_FindAndRepair_Success()
        Dim testOutput = debug_ErrorFindAndRepair_Success(11, 2, 0, 5)
        Assert.IsTrue(Not testOutput.Detects.Keys.First.Value.Contains("\\"))
        Assert.IsTrue(Not testOutput.FileKeys.Keys.First.Value.Contains("\\"))
    End Sub
End Class