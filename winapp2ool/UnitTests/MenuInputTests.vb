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

Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports winapp2ool

''' <summary>
''' Covers <see cref="MenuSection.HandleInput"/>, the one routine every menu's input goes
''' through: <c> 0 </c> exits, a number runs its option, an empty line runs the default option,
''' other text goes to the <see cref="MenuSection.OnText"/> handler, and anything else sets the
''' invalid-input header
''' </summary>
<TestClass>
Public Class MenuInputTests

    ''' <summary>How many times each option's handler has run, by option number</summary>
    Private _runs As Integer()

    ''' <summary>The inputs the text handler has received</summary>
    Private _texts As List(Of String)

    ''' <summary>Clears the counters, the pending header and the exit flag</summary>
    <TestInitialize>
    Public Sub Reset()

        _runs = New Integer(3) {}
        _texts = New List(Of String)
        MenuHeaderText = ""
        ExitPending = False

    End Sub

    ''' <summary>
    ''' Builds a menu with three options that count their runs. Option 2 is added only when
    ''' <paramref name="showSecond"/> is <c> True </c>, so the third handler then answers to 2.
    ''' </summary>
    '''
    ''' <param name="defaultOption">
    ''' The option, by its position in the builder (1 to 3), that is the default, or <c> 0 </c> for none
    ''' </param>
    '''
    ''' <param name="showSecond">
    ''' The condition the second option is added with
    ''' </param>
    Private Function BuildMenu(Optional defaultOption As Integer = 0,
                               Optional showSecond As Boolean = True) As MenuSection

        Return New MenuSection() _
            .AddDispatchedOption("One", "first", Sub() _runs(1) += 1, isDefault:=defaultOption = 1) _
            .AddDispatchedToggle("Two", "second", False, Sub() _runs(2) += 1, showSecond, isDefault:=defaultOption = 2) _
            .AddDispatchedColoredOption("Three", "third", ConsoleColor.Yellow, Sub() _runs(3) += 1, isDefault:=defaultOption = 3)

    End Function

    ''' <summary>
    ''' A text handler that records its input and handles only <c> "known" </c>
    ''' </summary>
    '''
    ''' <param name="input">
    ''' The user's input
    ''' </param>
    Private Function KnownOnly(input As String) As Boolean

        _texts.Add(input)
        Return input = "known"

    End Function

    ''' <summary>Asserts that the pending header is the invalid-input one</summary>
    Private Shared Sub AssertInvalid()

        Assert.AreEqual(invInpStr, MenuHeaderText)
        Assert.AreEqual(ConsoleColor.Red, MenuHeaderTextColor)

    End Sub

    ''' <summary>Asserts that no option's handler has run</summary>
    Private Sub AssertNothingRan()

        CollectionAssert.AreEqual({0, 0, 0, 0}, _runs)

    End Sub

    <TestMethod>
    Public Sub Zero_Exits()

        For Each zero In {"0", "00", "+0", " 0"}

            ExitPending = False
            BuildMenu(defaultOption:=1).HandleInput(zero)
            Assert.IsTrue(ExitPending, zero)

        Next

        AssertNothingRan()

    End Sub

    <TestMethod>
    Public Sub Number_RunsThatOption()

        BuildMenu().HandleInput("3")

        CollectionAssert.AreEqual({0, 0, 0, 1}, _runs)
        Assert.AreEqual("", MenuHeaderText)
        Assert.IsFalse(ExitPending)

    End Sub

    <TestMethod>
    Public Sub NumberOutOfRange_IsInvalid()

        For Each number In {"4", "-1", "99"}

            BuildMenu().HandleInput(number)

        Next

        AssertNothingRan()
        AssertInvalid()

    End Sub

    <TestMethod>
    Public Sub Number_SkipsAHiddenOption()

        BuildMenu(showSecond:=False).HandleInput("2")

        CollectionAssert.AreEqual({0, 0, 0, 1}, _runs)

    End Sub

    <TestMethod>
    Public Sub Empty_RunsTheDefault()

        Dim menu = BuildMenu(defaultOption:=3)
        menu.HandleInput("")

        CollectionAssert.AreEqual({0, 0, 0, 1}, _runs)
        Assert.AreEqual(3, menu.DefaultOption)
        Assert.AreEqual("", MenuHeaderText)

    End Sub

    <TestMethod>
    Public Sub Empty_WithoutADefault_IsInvalid()

        Dim menu = BuildMenu()
        menu.HandleInput("")

        Assert.AreEqual(0, menu.DefaultOption)
        AssertNothingRan()
        AssertInvalid()

    End Sub

    <TestMethod>
    Public Sub HiddenDefault_IsNoDefault()

        Dim menu = BuildMenu(defaultOption:=2, showSecond:=False)
        menu.HandleInput("")

        Assert.AreEqual(0, menu.DefaultOption)
        AssertNothingRan()
        AssertInvalid()

    End Sub

    <TestMethod>
    Public Sub DefaultAfterAHiddenOption_KeepsItsNumber()

        Assert.AreEqual(2, BuildMenu(defaultOption:=3, showSecond:=False).DefaultOption)

    End Sub

    <TestMethod>
    Public Sub SecondDefault_Throws()

        Dim menu = New MenuSection().AddDispatchedOption("One", "first", Sub() _runs(1) += 1, isDefault:=True)

        Assert.ThrowsExactly(Of InvalidOperationException)(
            Sub() menu.AddDispatchedOption("Two", "second", Sub() _runs(2) += 1, isDefault:=True))

    End Sub

    <TestMethod>
    Public Sub Text_Handled()

        BuildMenu().OnText(AddressOf KnownOnly).HandleInput("known")

        CollectionAssert.AreEqual({"known"}, _texts)
        Assert.AreEqual("", MenuHeaderText)
        AssertNothingRan()

    End Sub

    <TestMethod>
    Public Sub Text_NotHandled_IsInvalid()

        BuildMenu().OnText(AddressOf KnownOnly).HandleInput("x")

        CollectionAssert.AreEqual({"x"}, _texts)
        AssertInvalid()

    End Sub

    <TestMethod>
    Public Sub Text_WithoutAHandler_IsInvalid()

        BuildMenu().HandleInput("x")

        AssertNothingRan()
        AssertInvalid()

    End Sub

    <TestMethod>
    Public Sub Number_NeverReachesTheTextHandler()

        BuildMenu().OnText(AddressOf KnownOnly).HandleInput("2")

        Assert.AreEqual(0, _texts.Count)
        CollectionAssert.AreEqual({0, 0, 1, 0}, _runs)

    End Sub

    <TestMethod>
    Public Sub Empty_ReachesTheTextHandler_OnlyWithoutADefault()

        BuildMenu(defaultOption:=1).OnText(AddressOf KnownOnly).HandleInput("")
        Assert.AreEqual(0, _texts.Count)

        BuildMenu().OnText(AddressOf KnownOnly).HandleInput("")
        CollectionAssert.AreEqual({""}, _texts)

    End Sub

    <TestMethod>
    Public Sub EndOfInput_IsInvalid()

        BuildMenu(defaultOption:=1).OnText(AddressOf KnownOnly).HandleInput(Nothing)

        AssertNothingRan()
        Assert.AreEqual(0, _texts.Count)
        Assert.IsFalse(ExitPending)
        AssertInvalid()

    End Sub

End Class
