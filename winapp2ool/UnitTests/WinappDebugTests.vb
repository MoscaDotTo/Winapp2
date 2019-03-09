'    Copyright (C) 2018-2019 Robbie Ward
' 
'    This file is a part of Winapp2ool
' 
'    Winapp2ool is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    Winap2ool is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with Winapp2ool.  If not, see <http://www.gnu.org/licenses/>.
Option Strict On
Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
''' <summary>
''' Unit and integration tests for winapp2ool
''' 's WinappDebug module
''' Test naming convention: methodName_testState_expectedBehavior
''' </summary>
<TestClass()> Public Class WinappDebugTests

    ''' <summary>
    ''' Initializes WinappDebug with provided commandline args
    ''' </summary>
    ''' <param name="args">An array of args to pass to WinappDebug</param>
    ''' <param name="addHalt">Optional Boolean specifying whether or not the halting flag should be added to the args</param>
    Private Sub setDebugStage(args As String(), Optional addHalt As Boolean = False)
        setCmdLineArgs(AddressOf winapp2ool.WinappDebug.handleCmdLine, args, addHalt)
    End Sub

    ''' <summary>
    ''' Tests the commandline handling for WinappDebug to ensure success under no input conditions
    ''' </summary>
    <TestMethod()> Public Sub handleCmdLine_NoInputSuccess()
        ' Test case: Do nothing, expect our default values
        setDebugStage({}, True)
        Assert.AreEqual(winapp2ool.winappDebugFile1.path, winapp2ool.winappDebugFile3.path)
        Assert.AreEqual("winapp2-debugged.ini", winapp2ool.winappDebugFile3.secondName)
        Assert.AreNotEqual(winapp2ool.winappDebugFile1.secondName, winapp2ool.winappDebugFile3.secondName)
    End Sub

    ''' <summary>
    ''' Tests the commandline handling for WinappDebug to ensure success when changing file directory or name parameters
    ''' </summary>
    <TestMethod()> Public Sub handleCmdLine_ChangeFileParamsSuccess()
        ' First test case: Change only the first file name parameter
        setDebugStage({"-1f", "winapp2debugged.ini"}, True)
        Assert.AreEqual("winapp2debugged.ini", winapp2ool.winappDebugFile1.name)
        Assert.AreEqual($"{Environment.CurrentDirectory}\{winapp2ool.winappDebugFile1.name}", winapp2ool.winappDebugFile1.path)
        Assert.AreEqual("winapp2.ini", winapp2ool.winappDebugFile3.name)
        ' Second test case: Change the first and third file name parameters, also tests setting of a subdirectory through the file command
        setDebugStage({"-1f", "winapp2debugged.ini", "-3f", "\subdir\winapp2debugged.ini"}, True)
        Assert.AreEqual("winapp2debugged.ini", winapp2ool.winappDebugFile1.name)
        Assert.AreEqual(Environment.CurrentDirectory, winapp2ool.winappDebugFile1.dir)
        Assert.AreEqual("winapp2debugged.ini", winapp2ool.winappDebugFile3.name)
        Assert.AreEqual(Environment.CurrentDirectory & "\subdir", winapp2ool.winappDebugFile3.dir)
        ' Third test case: Change directory and name parameters of a single file
        setDebugStage({"-1d", "C:\Test Directory\", "-1f", "winapp2-test2.ini"}, True)
        Assert.AreEqual("C:\Test Directory", winapp2ool.winappDebugFile1.dir)
        Assert.AreEqual("winapp2-test2.ini", winapp2ool.winappDebugFile1.name)
    End Sub

    ''' <summary>
    ''' Tests the commandline handling for WinappDebug to ensure success when enabling autocorrect
    ''' </summary>
    <TestMethod()> Public Sub handleCmdLine_EnableAutoCorrectSuccess()
        ' Test case: Enable autocorrect 
        setDebugStage({"-c"}, True)
        Assert.AreEqual(winapp2ool.saveChanges, True)
    End Sub

    ' Tests below this point until the marked point test that individual scans and repairs work within WinappDebug's debug method

    ''' <summary>
    ''' Toggles off all the tests except a specific one
    ''' </summary>
    ''' <param name="lintRuleIndex">The index of the rule to be enabled</param>
    Private Sub disableAllTestsExcept(lintRuleIndex As Integer)
        ' Don't correct all formatting
        winapp2ool.CorrectFormatting1 = False
        winapp2ool.CorrectSomeFormatting1 = False
        For i As Integer = 0 To winapp2ool.WinappDebug.Rules1.Count - 1
            If Not i = lintRuleIndex Then
                ' Turn off all rules by default 
                winapp2ool.WinappDebug.Rules1(i).turnOff()
            Else
                ' Enable the rule we want
                winapp2ool.WinappDebug.Rules1(i).turnOn()
            End If
        Next
    End Sub

    ''' <summary>
    ''' Returns a winapp2file object containing the requested test
    ''' </summary>
    ''' <param name="testNum">The test index to return from the unit test file</param>
    ''' <returns></returns>
    Private Function getSingleTestFile(testNum As Integer) As winapp2ool.winapp2file
        Dim unitFile As New winapp2ool.iniFile(Environment.CurrentDirectory, "WinappDebugUnitTests.ini")
        unitFile.validate()
        Dim testFile As New winapp2ool.iniFile
        Dim testSection = unitFile.sections.Values(testNum)
        testFile.sections.Add(testSection.name, testSection)
        Return New winapp2ool.winapp2handler.winapp2file(testFile)
    End Function

    ''' <summary>
    ''' Tests the debug function in winapp2ool using tests from WinappDebugUnitTests.ini
    ''' </summary>
    ''' <param name="testNum">The test number to request from the file</param>
    ''' <param name="expectedErrsWithoutRepair">The expected number of errors to be found on the first run</param>
    ''' <param name="expectedErrsWithRepair">The expected number of errors to be found after repairs are run</param>
    ''' <returns></returns>
    Public Function debug_ErrorFindAndRepair_Success(testNum As Integer, expectedErrsWithoutRepair As Integer, expectedErrsWithRepair As Integer, lintRuleIndex As Integer) As winapp2ool.winapp2entry
        ' Initalize the default state of the module
        setDebugStage({}, True)
        ' Disable all the lint rules we're not currently testing
        disableAllTestsExcept(lintRuleIndex)
        Dim test As winapp2ool.winapp2file = getSingleTestFile(testNum)
        ' Confirm the errors are found without autocorrect on
        winapp2ool.WinappDebug.debug(test)
        Assert.AreEqual(expectedErrsWithoutRepair, winapp2ool.WinappDebug.errorsFound)
        ' Enable repairs
        winapp2ool.WinappDebug.CorrectSomeFormatting1 = True
        winapp2ool.WinappDebug.debug(test)
        ' Confirm the errors are still found (ie. not erroneously corrected during the first test)
        Assert.AreEqual(expectedErrsWithoutRepair, winapp2ool.WinappDebug.errorsFound)
        winapp2ool.WinappDebug.debug(test)
        ' Confirm fixable errors are fixed
        Assert.AreEqual(expectedErrsWithRepair, winapp2ool.WinappDebug.errorsFound)
        ' Return the entry for further assessment
        Return New winapp2ool.winapp2handler.winapp2entry(test.entrySections.Last.sections.Values.First)
    End Function

    ''' <summary>
    ''' Runs tests to ensure that keys with duplicate values are detected and removed
    ''' </summary>
    <TestMethod> Public Sub debug_DuplicateKeyValue_FindAndRepair_Success()
        ' Disable all the tests except the duplicate checks
        Dim testOutput = debug_ErrorFindAndRepair_Success(0, 12, 0, 7)
        ' We expect that the returned object should have 1 key in each keylist except the last
        For i = 0 To testOutput.keyListList.Count - 2
            Dim lst = testOutput.keyListList(i)
            Assert.AreEqual(1, lst.keyCount)
        Next
        ' The last keylist (the error keys) should be empty
        Assert.AreEqual(0, testOutput.keyListList.Last.keyCount)
    End Sub

    ''' <summary>
    ''' Runs tests to ensure that keys that are improperly numbered are detected and repaired
    ''' </summary>
    <TestMethod> Public Sub debug_KeyNumberingError_FindAndRepair_Sucess()
        Dim testOutput = debug_ErrorFindAndRepair_Success(1, 8, 0, 2)
        For i = 0 To testOutput.keyListList.Count - 2
            Dim lst = testOutput.keyListList(i)
            Select Case lst.keyType
                Case "Detect", "DetectFile", "ExcludeKey", "FileKey", "RegKey"
                    Dim curKeys = lst.keys
                    ' We expect 2 keys in each of these lists, having the keyNumbers 1 and 2 respectively
                    Assert.AreEqual(curKeys.First.name, $"{lst.keyType}1")
                    Assert.AreEqual(curKeys.Last.name, $"{lst.keyType}2")
            End Select
        Next
    End Sub

End Class