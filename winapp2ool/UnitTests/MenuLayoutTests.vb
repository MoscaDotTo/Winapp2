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
''' Covers the geometry of rendered <see cref="MenuSection"/> output: that every line in a box is
''' the same width, that centered text is actually centered, that lines of any length start in
''' the same column and an overlong one keeps its left frame, and that a section reports itself
''' empty, and renders nothing, exactly when it has nothing in it
''' </summary>
'''
''' <remarks>
''' There is no console in the test host, so <c> GetConsoleWidth </c> falls back to its cached
''' default of 120. That is fine, because these assertions are about the relationship between the
''' border and the text, not about any particular width
''' </remarks>
<TestClass>
Public Class MenuLayoutTests

    ''' <summary>
    ''' Renders a <c> MenuSection </c> with the console redirected, returning the lines it wrote
    ''' </summary>
    '''
    ''' <param name="section">
    ''' The section to render
    ''' </param>
    '''
    ''' <returns>
    ''' Every non-empty line the section printed, in order
    ''' </returns>
    Private Shared Function Render(section As MenuSection) As String()

        Dim original = Console.Out

        Try

            Using captured As New StringWriter

                Console.SetOut(captured)
                section.Print()

                Return captured.ToString().
                    Split({Environment.NewLine}, StringSplitOptions.None).
                    Where(Function(line) line.Length > 0).
                    ToArray()

            End Using

        Finally

            Console.SetOut(original)

        End Try

    End Function

    ''' <summary>
    ''' Measures how far a bordered line's text sits from true center
    ''' </summary>
    '''
    ''' <param name="line">
    ''' A rendered line made of one space, a border glyph, the interior, and a closing border glyph
    ''' </param>
    '''
    ''' <returns>
    ''' Leading interior spaces minus trailing interior spaces: <c> 0 </c> when perfectly
    ''' centered, positive when the text sits right of center
    ''' </returns>
    Private Shared Function CenteringOffset(line As String) As Integer

        Dim interior = line.Substring(2, line.Length - 3)

        Return (interior.Length - interior.TrimStart().Length) -
               (interior.Length - interior.TrimEnd().Length)

    End Function

    <TestMethod>
    Public Sub CenteredText_IsActuallyCentered()

        ' Both parities: an odd-length run cannot split its remainder evenly, so allow one column
        For Each caption In {"Diff Summary", "Press Enter to continue", "abcde", "ab"}

            Dim section As New MenuSection
            section.AddTopBorder().AddColoredLine(caption, ConsoleColor.White, centered:=True).AddBottomBorder()

            Dim body = Render(section).First(Function(line) line.Contains(caption))
            Dim offset = CenteringOffset(body)

            Assert.IsTrue(Math.Abs(offset) <= 1,
                          $"'{caption}' is off center by {offset} columns: [{body}]")

        Next

    End Sub

    <TestMethod>
    Public Sub AllRenderedLines_ShareTheSameWidth()

        Dim section As New MenuSection
        section.AddTopBorder() _
               .AddColoredLine("Diff Summary", ConsoleColor.White, centered:=True) _
               .AddDivider() _
               .AddLine("No changes detected across 3715 entries") _
               .AddDivider(solid:=False) _
               .AddBottomBorder()

        Dim rendered = Render(section)
        Dim widths = rendered.Select(Function(line) line.Length).Distinct().ToArray()

        Assert.AreEqual(1, widths.Length,
                        $"box lines disagree on width ({String.Join(", ", widths)}):" & Environment.NewLine &
                        String.Join(Environment.NewLine, rendered))

    End Sub

    ''' <summary>
    ''' Renders a single left-aligned line of the given length and returns it
    ''' </summary>
    '''
    ''' <param name="length">
    ''' The number of characters of body text to render
    ''' </param>
    '''
    ''' <returns>
    ''' The rendered line
    ''' </returns>
    Private Shared Function RenderLineOfLength(length As Integer) As String

        Dim section As New MenuSection
        section.AddLine(New String("x"c, length))

        Return Render(section).Single()

    End Function

    <TestMethod>
    Public Sub OverlongLine_KeepsTheLeftFrameAndIndent()

        ' Well past any console width, so this is the "no closing border possible" path
        Dim overlong = RenderLineOfLength(400)
        Dim normal = RenderLineOfLength(10)

        Assert.AreEqual(normal.IndexOf("x"c), overlong.IndexOf("x"c),
                        "an overlong line should start in the same column as a normal one")
        Assert.AreEqual(" ║ ", overlong.Substring(0, 3),
                        "the opener and indent should survive even when the closer cannot")

    End Sub

    <TestMethod>
    Public Sub EveryLineLength_StartsInTheSameColumn()

        ' Sweeps the boundary between framed and unframed. The band just under the console width
        ' used to jam the closing border against the text; nothing may shift the left edge
        Dim expected = RenderLineOfLength(1).IndexOf("x"c)

        For length = 1 To 200

            Dim rendered = RenderLineOfLength(length)

            Assert.AreEqual(expected, rendered.IndexOf("x"c),
                            $"a {length}-character line starts in the wrong column: [{rendered}]")

        Next

    End Sub

    <TestMethod>
    Public Sub FramedLines_NeverRunPastTheBorder()

        ' Any line that still carries a closing border must respect the box width
        Dim framedWidth = RenderLineOfLength(1).Length

        For length = 1 To 200

            Dim rendered = RenderLineOfLength(length)

            If rendered.EndsWith("║", StringComparison.Ordinal) AndAlso rendered.Length > 3 Then

                Assert.AreEqual(framedWidth, rendered.Length,
                                $"a {length}-character line closed its border at the wrong width")

            End If

        Next

    End Sub

    <TestMethod>
    Public Sub EmptySection_ReportsEmptyAndRendersNothing()

        Dim empty As New MenuSection

        Assert.IsTrue(empty.IsEmpty)
        Assert.AreEqual(0, Render(empty).Length)

    End Sub

    <TestMethod>
    Public Sub ScreenRows_CountsOneRowPerWidthFilled()

        Assert.AreEqual(1, ScreenRows(Nothing, 120), "an empty line still takes a row")
        Assert.AreEqual(1, ScreenRows(New String("x"c, 120), 120))
        Assert.AreEqual(2, ScreenRows(New String("x"c, 121), 120))
        Assert.AreEqual(4, ScreenRows(New String("x"c, 404), 120))

    End Sub

    <TestMethod>
    Public Sub PassRows_CountsEveryRowAPrintWrote()

        ' initModule sizes the window from PassRows, so a row written around cwl would leave the
        ' menu taller than the window it fits
        Dim menu = MenuSection.CreateCompleteMenu("Layout", {"A description line"}, ConsoleColor.Cyan)
        menu.AddOption("First", "An option") _
            .AddBlank() _
            .AddColoredLine("A colored line", ConsoleColor.Green, centered:=True) _
            .AddLine(New String("x"c, 400))

        Dim original = Console.Out

        Try

            Using captured As New StringWriter

                Console.SetOut(captured)
                menu.Print()

                Dim lines = captured.ToString().Split({Environment.NewLine}, StringSplitOptions.None)
                ' The headless console is 120 columns wide, see the class remarks
                Dim expected = lines.Take(lines.Length - 1).Sum(Function(line) Math.Max(1, (line.Length + 119) \ 120))

                Assert.AreEqual(expected, PassRows)

            End Using

        Finally

            Console.SetOut(original)

        End Try

    End Sub

    <TestMethod>
    Public Sub NestedPass_HoldsOutputUntilTheOutermostFlush()

        ' initModule wraps the menu's own Print in a pass so it can size the window before
        ' anything reaches the console
        Dim section As New MenuSection
        section.AddTopBorder().AddLine("held back").AddBottomBorder()

        Dim original = Console.Out

        Try

            Using captured As New StringWriter

                Console.SetOut(captured)

                BeginBuffered()
                section.Print()
                Assert.AreEqual("", captured.ToString(), "the inner Print ended the outer pass")

                FlushBuffered()
                Assert.IsTrue(captured.ToString().Contains("held back"), "the outer flush wrote nothing")

            End Using

        Finally

            Console.SetOut(original)

        End Try

    End Sub

    <TestMethod>
    Public Sub SectionWithAnyItem_IsNotEmpty()

        Dim divider As New MenuSection
        divider.AddDivider(solid:=False)

        Assert.IsFalse(divider.IsEmpty, "a lone divider still renders a line, so it is not empty")

    End Sub

End Class
