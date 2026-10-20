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
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports winapp2ool

''' <summary>
''' Covers when <see cref="iniFileChooser"/>'s menu closes after a directory change. Every choice
''' closes it only when the file exists or <see cref="iniFileChooser.MustExist"/> is off, which is
''' what lets <see cref="iniFileChooser.Load"/> keep the chooser open until the user finds a file.
''' </summary>
<TestClass>
Public Class IniFileChooserTests

    ''' <summary>The chooser's filename, unique so that no folder we visit already has it</summary>
    Private Shared ReadOnly FileName As String = $"chooser-{Guid.NewGuid():N}.ini"

    ''' <summary>A scratch folder, with a <c> child </c> folder inside it</summary>
    Private _root As String

    ''' <summary>The <c> child </c> folder inside <see cref="_root"/></summary>
    Private _child As String

    ''' <summary>Creates the scratch folders and clears the pending header and the exit flag</summary>
    <TestInitialize>
    Public Sub Setup()

        _root = Path.Combine(Path.GetTempPath(), $"winapp2ool-chooser-{Guid.NewGuid():N}")
        _child = Path.Combine(_root, "child")
        Directory.CreateDirectory(_child)
        MenuHeaderText = ""
        ExitPending = False

    End Sub

    ''' <summary>Deletes the scratch folders</summary>
    <TestCleanup>
    Public Sub Cleanup()

        If Directory.Exists(_root) Then Directory.Delete(_root, recursive:=True)
        ExitPending = False

    End Sub

    ''' <summary>
    ''' Builds the chooser's menu and gives it <paramref name="input"/>. With no second name,
    ''' option 2 is "Default directory" and option 3 is "Parent folder".
    ''' </summary>
    '''
    ''' <param name="chooser">
    ''' The chooser whose menu gets the input
    ''' </param>
    '''
    ''' <param name="input">
    ''' The line the user typed
    ''' </param>
    Private Shared Sub Enter(chooser As iniFileChooser, input As String)

        chooser.buildMenu().HandleInput(input)

    End Sub

    <TestMethod>
    Public Sub ParentFolder_WithoutTheFile_StaysOpen()

        Dim chooser As New iniFileChooser(_child, FileName)

        Enter(chooser, "3")

        Assert.AreEqual(_root, chooser.Dir)
        Assert.IsFalse(ExitPending)
        StringAssert.Contains(MenuHeaderText, "does not exist")

    End Sub

    <TestMethod>
    Public Sub ParentFolder_WithTheFile_Closes()

        File.WriteAllText(Path.Combine(_root, FileName), "")
        Dim chooser As New iniFileChooser(_child, FileName)

        Enter(chooser, "3")

        Assert.AreEqual(_root, chooser.Dir)
        Assert.IsTrue(ExitPending)

    End Sub

    <TestMethod>
    Public Sub ParentFolder_MustExistOff_ClosesWithoutTheFile()

        Dim chooser As New iniFileChooser(_child, FileName, mustExist:=False)

        Enter(chooser, "3")

        Assert.AreEqual(_root, chooser.Dir)
        Assert.IsTrue(ExitPending)

    End Sub

    <TestMethod>
    Public Sub ParentFolder_AtADriveRoot_LeavesTheDirectory()

        Dim driveRoot = Path.GetPathRoot(_root)
        Dim chooser As New iniFileChooser(driveRoot, FileName)

        Enter(chooser, "3")

        Assert.AreEqual(driveRoot, chooser.Dir)
        Assert.IsFalse(ExitPending)
        StringAssert.Contains(MenuHeaderText, "no parent folder")

    End Sub

    <TestMethod>
    Public Sub DefaultDirectory_WithoutTheFile_StaysOpen()

        Dim chooser As New iniFileChooser(_child, FileName)

        Enter(chooser, "2")

        Assert.AreEqual(Environment.CurrentDirectory, chooser.Dir)
        Assert.IsFalse(ExitPending)

    End Sub

    <TestMethod>
    Public Sub TypedDirectory_WithoutTheFile_StaysOpen()

        Dim chooser As New iniFileChooser(_child, FileName)

        Enter(chooser, _root & "\")

        Assert.AreEqual(_root & "\", chooser.Dir)
        Assert.IsFalse(ExitPending)
        StringAssert.Contains(MenuHeaderText, "does not exist")

    End Sub

    <TestMethod>
    Public Sub TypedDirectory_WithTheFile_Closes()

        File.WriteAllText(Path.Combine(_root, FileName), "")
        Dim chooser As New iniFileChooser(_child, FileName)

        Enter(chooser, _root & "\")

        Assert.IsTrue(ExitPending)

    End Sub

End Class
