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
Imports System.Text

''' <summary> 
''' MenuMaker is a driver module for powering dynamic finite 
''' state console applications with variable numbered menus 
''' </summary>
Module MenuMaker

    ''' <summary>
    ''' Names the menu frame types, numbered the same as the <c> frameNum </c> and
    ''' <c> borderInd </c> integers the frame builders take
    ''' </summary>
    Public Enum FrameType

        ''' <summary>Vertical borders, <c> ║ </c> on both ends</summary>
        Vertical = 0

        ''' <summary>A top border that opens a box, <c> ╔ </c> and <c> ╗ </c></summary>
        Top = 1

        ''' <summary>A bottom border that closes a box, <c> ╚ </c> and <c> ╝ </c></summary>
        Bottom = 2

        ''' <summary>A divider inside a box, <c> ╠ </c> and <c> ╣ </c></summary>
        Conjoin = 3

    End Enum

    ''' <summary>
    ''' An instruction to press the Enter button to continue 
    ''' </summary>
    Public ReadOnly Property pressEnterStr As String = "Press Enter to continue"

    ''' <summary>
    ''' An instruction to press any key to return to the previous menu 
    ''' </summary>
    Public ReadOnly Property anyKeyStr As String = "Press any key to return to the menu."

    ''' <summary> 
    ''' An error message informing the user their input was invalid 
    ''' </summary>
    Public ReadOnly Property invInpStr As String = "Invalid input. Please try again."

    ''' <summary> 
    ''' An instruction for the user to provide input
    ''' </summary>
    Public ReadOnly Property promptStr As String = "Enter a number, or leave blank to run the default: "

    ''' <summary>
    ''' The width to which <see cref="printMenuOpt"/> pads the <c> #. Name </c> half of an option,
    ''' so descriptions line up. A longer name isn't cut, it pushes its description to the right.
    ''' </summary>
    Private Property menuItemLength As Integer

    ''' <summary>
    ''' Indicates whether the menu header should be printed with color. <see cref="setNextMenuHeaderText"/>
    ''' sets it, but nothing reads it.
    ''' </summary>
    Public Property ColorHeader As Boolean

    ''' <summary>
    ''' The color most recently passed to <see cref="setNextMenuHeaderText"/>. Nothing reads it.
    ''' Menus take the header color from <see cref="MenuHeaderTextColor"/>.
    ''' </summary>
    Public Property HeaderColor As ConsoleColor

    ''' <summary>
    ''' Indicates whether winapp2ool runs silently. When <c> True </c>, <see cref="cwl"/>,
    ''' <see cref="crk"/>, <see cref="crl"/> and <see cref="clrConsole"/> do nothing, so exception
    ''' reports reach only the log, and <see cref="initModule"/> ends the process.
    ''' <br /> Default: <c> False </c>
    ''' </summary>
    Public Property SuppressOutput As Boolean = False

    ''' <summary>
    ''' Indicates whether the current menu should close at the end of its loop iteration
    ''' </summary>
    Public Property ExitPending As Boolean

    ''' <summary>
    ''' A pending header that replaces the next complete menu's own header, or <c> "" </c> when
    ''' none is pending. <see cref="MenuSection.CreateCompleteMenu"/> uses it and clears it.
    ''' </summary>
    Public Property MenuHeaderText As String

    ''' <summary>The color of the pending header in <see cref="MenuHeaderText"/></summary>
    Public Property MenuHeaderTextColor As ConsoleColor = ConsoleColor.Red

    ''' <summary>
    ''' The number the next printed menu option gets. Exit is option <c> 0 </c>.
    ''' </summary>
    Private Property OptNum As Integer = 0

    ''' <summary>
    ''' Frame characters used to open a menu line 
    ''' </summary>
    Private ReadOnly Property Openers As String() = {"║", "╔", "╚", "╠"}

    ''' <summary> 
    ''' Frame characters used to close a menu line 
    ''' </summary>
    Private ReadOnly Property Closers As String() = {"║", "╗", "╝", "╣"}

    ''' <summary>
    ''' The cached console window width, used to avoid unneeded calls to
    ''' <c> Console.WindowWidth </c>. It stays at <c> 120 </c> when there is no console to measure.
    ''' </summary>
    Private _cachedWindowWidth As Integer = 120

    ''' <summary>
    ''' The time at which the console window width was last checked
    ''' </summary>
    Private _lastWidthCheckTime As DateTime = DateTime.Now

    ''' <summary>
    ''' When non-<c> Nothing </c>, <see cref="cwl"/> writes, VT color escapes and the VT clear
    ''' escape are appended to this buffer instead of going to <c> Console.Out </c> directly. The
    ''' buffer is written as a single <c> Write </c> at the end of a render pass. On a console
    ''' without VT support we also flush it before each color change, so the color applies to
    ''' the right characters.
    ''' </summary>
    Private _outputBuffer As StringBuilder = Nothing

    ''' <summary>
    ''' The console width held constant for the duration of a render pass, or <c> Nothing </c>
    ''' when no pass is in progress. Set by <see cref="BeginBuffered"/> and cleared by
    ''' <see cref="FlushBuffered"/> so every line of one box is measured against the same width
    ''' </summary>
    Private _pinnedWidth As Integer? = Nothing

    ''' <summary>
    ''' Indicates whether render passes are buffered. When <c> False </c>, <see cref="BeginBuffered"/>
    ''' still pins the width but doesn't start a buffer, so each line goes straight to
    ''' <c> Console.Out </c>.
    ''' </summary>
    Public Property BufferingEnabled As Boolean = True

    ''' <summary>
    ''' Returns the console window width. During a render pass we return the pinned width.
    ''' Otherwise we read <c> Console.WindowWidth </c> at most once every 500 milliseconds and
    ''' keep the cached value if the read fails.
    ''' </summary>
    Private Function GetConsoleWidth() As Integer

        ' A render pass pins the width for its whole duration. Without that, the 500ms refresh
        ' below can land between two lines of the same box and leave it drawn at two different
        ' widths
        If _pinnedWidth.HasValue Then Return _pinnedWidth.Value

        ' The width is extremely unlikely to change during the printing process
        ' if ever, so only check it every 500 milliseconds at most 
        If DateTime.Now.Subtract(_lastWidthCheckTime).TotalMilliseconds > 500 Then

            Try
                _cachedWindowWidth = Console.WindowWidth
            Catch e As IO.IOException
            End Try
            _lastWidthCheckTime = DateTime.Now

        End If

        Return _cachedWindowWidth

    End Function

    ''' <summary>
    ''' Runs a menu loop: clears the console, shows the menu, reads a line and passes it to
    ''' <paramref name="handleInput"/>, until <see cref="exitModule"/> is called. Exiting returns
    ''' exactly one level up, to the menu that called this one. On entry the next header is set
    ''' to <paramref name="name"/>, and on exit to <c> "{name} closed" </c>. On exit we also call
    ''' <c> FlushIfDirty </c>, which writes changed settings only when <c> saveSettingsToDisk </c>
    ''' is on (it's off by default) and this isn't a command line run.
    ''' <br /><br />
    ''' In silent mode there's no input to read, so we save the global log and end the process
    ''' with exit code <c> 1 </c>.
    ''' </summary>
    '''
    ''' <param name="name">
    ''' The name of the module as it will be displayed to the user
    ''' </param>
    '''
    ''' <param name="showMenu">
    ''' The subroutine that prints the module's menu
    ''' </param>
    '''
    ''' <param name="handleInput">
    ''' The subroutine that handles the module's input
    ''' </param>
    '''
    ''' <param name="itmLen">
    ''' The width the <c> #. Name </c> half of each option is padded to <br /><br />
    ''' Optional, Default: <c> 35 </c>
    ''' </param>
    Public Sub initModule(name As String,
                          showMenu As Action,
                          handleInput As Action(Of String),
                 Optional itmLen As Integer = 35)

        Using gLogScope($"Loading module {name}")


            If SuppressOutput Then

                gLog($"Interactive menu '{name}' cannot run in silent mode (no input available); aborting")
                saveGlobalLog()
                Environment.Exit(1)

            End If

            ExitPending = False
            setNextMenuHeaderText(name)

            menuItemLength = itmLen

            Do Until ExitPending

                clrConsole()
                showMenu()
                Console.Write(Environment.NewLine & promptStr)
                handleInput(Console.ReadLine)

            Loop

            ExitPending = False

            FlushIfDirty()

            setNextMenuHeaderText($"{name} closed")

        End Using

        gLog($"Exited {name}", leadr:=True)

    End Sub

    ''' <summary>
    ''' Writes an empty line straight to the console. Unlike <see cref="cwl"/>, it ignores
    ''' <see cref="SuppressOutput"/> and skips the render buffer, so inside a buffered pass the
    ''' line comes out before everything buffered so far.
    ''' </summary>
    '''
    ''' <param name="condition">
    ''' Indicates whether the new line should be printed <br /><br />
    ''' Optional, Default: <c> True </c>
    ''' </param>
    Public Sub PrintNewLine(Optional condition As Boolean = True)

        If condition Then Console.WriteLine()

    End Sub

    ''' <summary>
    ''' Prints a line to the console if output isn't suppressed and <paramref name="cond"/> is
    ''' <c> True </c>. During a buffered render pass the line goes into the buffer instead.
    ''' </summary>
    '''
    ''' <param name="msg">
    ''' The string to be printed <br /><br />
    ''' Optional, Default: <c> Nothing </c> <br />
    ''' prints an empty line
    ''' </param>
    '''
    ''' <param name="cond">
    ''' Indicates whether the line should be printed <br /><br />
    ''' Optional, Default: <c> True </c>
    ''' </param>
    Public Sub cwl(Optional msg As String = Nothing,
                   Optional cond As Boolean = True)

        If Not cond OrElse SuppressOutput Then Return

        If _outputBuffer IsNot Nothing Then

            If msg IsNot Nothing Then _outputBuffer.Append(msg)
            _outputBuffer.Append(Environment.NewLine)
            Return

        End If

        Console.WriteLine(msg)

    End Sub

    ''' <summary>
    ''' Begins a render pass. We pin the console width for the whole pass and, unless buffering
    ''' is off or output is suppressed, start a buffer that later <see cref="cwl"/> calls append to.
    ''' <see cref="FlushBuffered"/> writes the buffer as one <c> Write </c> and ends the pass.
    ''' Passes don't nest: a call during a pass does nothing, and the first
    ''' <see cref="FlushBuffered"/> ends it.
    ''' </summary>
    Public Sub BeginBuffered()

        ' Pinned before the buffering guards below, so the width stays stable for the pass even
        ' when buffering is switched off or output is suppressed
        If Not _pinnedWidth.HasValue Then _pinnedWidth = GetConsoleWidth()

        If Not BufferingEnabled OrElse SuppressOutput Then Return
        If _outputBuffer IsNot Nothing Then Return

        _outputBuffer = New StringBuilder(4096)

    End Sub

    ''' <summary>
    ''' Ends a render pass, writing the accumulated buffer to <c> Console.Out </c> in a single
    ''' call and releasing the pinned console width.
    ''' </summary>
    Public Sub FlushBuffered()

        Dim sb = _outputBuffer
        _outputBuffer = Nothing
        _pinnedWidth = Nothing

        If sb Is Nothing Then Return

        If sb.Length > 0 Then Console.Out.Write(sb.ToString())

    End Sub

    ''' <summary>
    ''' Writes out the buffer, if any, without ending the pass or releasing the pinned width,
    ''' so that a Win32 color change applies only to text printed after it.
    ''' </summary>
    Private Sub FlushBufferIfActive()

        If _outputBuffer Is Nothing OrElse _outputBuffer.Length = 0 Then Return

        Console.Out.Write(_outputBuffer.ToString())
        _outputBuffer.Length = 0

    End Sub

    ''' <summary>
    ''' Waits for the user to press a key if output
    ''' is not currently being suppressed
    ''' </summary>
    Public Sub crk()

        If SuppressOutput Then Return

        Console.ReadKey()

    End Sub

    ''' <summary>
    ''' Waits for the user to press Enter if output
    ''' is not currently being suppressed
    ''' </summary>
    Public Sub crl()

        If SuppressOutput Then Return

        Console.ReadLine()

    End Sub

    ''' <summary>
    ''' VT escape that homes the cursor and erases the entire screen. One
    ''' atomic write to a VT-capable terminal replaces the three Win32 calls
    ''' (<c> SetConsoleCursorPosition </c>, <c> FillConsoleOutputCharacter </c>,
    ''' <c> FillConsoleOutputAttribute </c>) that <c> Console.Clear </c>
    ''' performs internally.
    ''' </summary>
    Private ReadOnly VtClearScreen As String = ChrW(&H1B) & "[H" & ChrW(&H1B) & "[2J"

    ''' <summary>
    ''' Clears the console when <paramref name="cond"/> is <c> True </c> and output isn't
    ''' suppressed. On a console without VT support we also skip the clear when output is redirected.
    ''' </summary>
    '''
    ''' <remarks>
    ''' On a VT-capable terminal (Windows 10 1607+ conhost, Windows Terminal)
    ''' this emits the VT clear escape, which is a single write. If a buffered
    ''' render is active, the escape is appended to the buffer so the clear and
    ''' the new menu arrive at the terminal in one write, without flicker.
    '''
    ''' On older consoles we fall back to <c> Console.Clear </c>. When output is
    ''' redirected (test runner, piped to a file) we skip that call, because clearing
    ''' a non-console sink would corrupt downstream output.
    ''' </remarks>
    '''
    ''' <param name="cond">
    ''' Indicates whether the console should be cleared <br /><br />
    ''' Optional, Default: <c> True </c>
    ''' </param>
    Public Sub clrConsole(Optional cond As Boolean = True)

        If Not cond OrElse SuppressOutput Then Return

        ' VT path: an escape sequence written through Console.Out is safe even
        ' if output happens to be redirected to a pipe/file. The caller opted
        ' in to VT by virtue of HasVT being True (which the production probe
        ' only sets when stdout is a real VT-capable console).
        If TerminalCapabilities.HasVT Then

            If _outputBuffer IsNot Nothing Then
                _outputBuffer.Append(VtClearScreen)
            Else
                Console.Out.Write(VtClearScreen)
            End If

            Return

        End If

        ' Legacy path: Console.Clear writes nowhere meaningful when stdout
        ' isn't a real console (test runners, pipelines, redirection) and will
        ' usually throw IOException. Skip it outright in that case.
        If Console.IsOutputRedirected Then Return

        Try
            Console.Clear()
        Catch e As IO.IOException
            ' Belt and braces: IsOutputRedirected should already cover this.
        End Try

    End Sub

    ''' <summary> 
    ''' Returns an empty menu line, or a variety of filled menu lines 
    ''' </summary>
    ''' 
    ''' <param name="frameNum"> 
    ''' Indicates which frame should be returned <br />
    ''' 
    ''' <list type="bullet">
    ''' 
    ''' <item>
    ''' <description>
    ''' 0: Vertical frames <c> ║     ║ </c>
    ''' </description>
    ''' </item>
    ''' 
    ''' <item>
    ''' <description> 
    ''' 1: Downward opening 90° angle frames <c> ╔ ═ ═ ═ ═ ═╗ </c>
    ''' </description> 
    ''' </item>
    ''' 
    ''' <item>
    ''' <description> 
    ''' 2: Upward opening 90° angle frames <c> ╚ ═ ═ ═ ═ ═╝ </c>
    ''' </description> 
    ''' </item>
    ''' 
    ''' <item> 
    ''' <description> 
    ''' 3: Inward facing T-frames <c> ╠ ═ ═ ═ ═ ═ ╣ </c> 
    ''' </description> 
    ''' </item>
    ''' 
    ''' </list>
    ''' 
    ''' <br /> Optional, Default: <c> 0 </c>
    ''' </param>
    '''
    ''' <param name="fillFrame">
    ''' Indicates whether to fill the line with <c> ═ </c> instead of spaces <br /><br />
    ''' Optional, Default: <c> False </c>
    ''' </param>
    '''
    ''' <returns>
    ''' A full-width line holding the frame requested by <paramref name="frameNum"/>
    ''' </returns>
    Private Function getFrame(Optional frameNum As Integer = 0,
                              Optional fillFrame As Nullable(Of Boolean) = False) As String

        Return mkMenuLine("", 2, frameNum, fillFrame)

    End Function



    ''' <summary>
    ''' Replaces the next complete menu's header with <paramref name="txt"/>, for delivering a
    ''' status update or error message between menus or modules. The next
    ''' <see cref="MenuSection.CreateCompleteMenu"/> call uses it and clears it, and a later call
    ''' to this sub before then overwrites it.
    ''' </summary>
    '''
    ''' <param name="txt">The header text to show</param>
    '''
    ''' <param name="cond">
    ''' Indicates whether to set the header <br /><br />
    ''' Optional, Default: <c> True </c>
    ''' </param>
    '''
    ''' <param name="printColor">
    ''' The color of the header <br /><br />
    ''' Optional, Default: <c> ConsoleColor.White </c>
    ''' </param>
    Public Sub setNextMenuHeaderText(txt As String,
                            Optional cond As Boolean = True,
                            Optional printColor As ConsoleColor = ConsoleColor.White)

        If Not cond Then Return

        MenuHeaderText = txt
        MenuHeaderTextColor = printColor
        ColorHeader = True
        HeaderColor = printColor

    End Sub


    ''' <summary>
    ''' Sets <paramref name="errText"/> as the next menu header when <paramref name="cond"/> is
    ''' <c> True </c>, and returns <paramref name="cond"/>
    ''' </summary>
    '''
    ''' <param name="cond">
    ''' Indicates whether the action should be denied
    ''' </param>
    '''
    ''' <param name="errText">
    ''' The error text to be printed in the menu header
    ''' </param>
    '''
    ''' <returns>
    ''' <c> True </c> if the action is denied and the caller should not proceed,
    ''' <c> False </c> if it may go ahead
    ''' </returns>
    Public Function denyActionWithHeader(cond As Boolean,
                                         errText As String) As Boolean

        setNextMenuHeaderText(errText, cond)

        Return cond

    End Function

    ''' <summary>
    ''' Returns the action that would flip a setting, as a String
    ''' </summary>
    '''
    ''' <param name="setting">
    ''' The current state of a module setting
    ''' </param>
    '''
    ''' <returns>
    ''' <c> "Disable" </c> if <paramref name="setting"/> is <c> True </c>,
    ''' <br /> <c> "Enable" </c> if it is <c> False </c> or <c> Nothing </c>
    ''' </returns>
    Public Function enStr(setting As Nullable(Of Boolean)) As String

        Return If(setting, "Disable", "Enable")

    End Function

    ''' <summary>
    ''' Makes <see cref="initModule"/> close the current menu once the current input has
    ''' been handled
    ''' </summary>
    Public Sub exitModule()

        ExitPending = True

    End Sub

    ''' <summary>Resets option numbering so the next printed option is number <c> 0 </c></summary>
    Public Sub resetMenuNumbering()

        OptNum = 0

    End Sub

    ''' <summary>
    ''' Prints a line bounded by vertical menu frames, or an empty menu line
    ''' if <paramref name="lineString"/> is <c> Nothing </c> or empty
    ''' </summary>
    '''
    ''' <param name="lineString">
    ''' The text to be printed <br /><br />
    ''' Optional, Default: <c> Nothing </c>
    ''' </param>
    '''
    ''' <param name="isCentered">
    ''' Indicates whether the printed text should be centered <br /><br />
    ''' Optional, Default: <c> False </c>
    ''' </param>
    '''
    ''' <param name="cond">
    ''' Indicates whether the line should be printed <br /><br />
    ''' Optional, Default: <c> True </c>
    ''' </param>
    Private Sub printMenuLine(Optional lineString As String = Nothing,
                              Optional isCentered As Boolean = False,
                              Optional cond As Boolean = True)

        If Not cond Then Return

        ' A blank row is already a finished full-width line, so it goes straight out. Feeding it
        ' back through mkMenuLine would frame the frame
        If lineString = Nothing Then printRenderedLine(getFrame()) : Return

        cwl(mkMenuLine(lineString, If(isCentered, 0, 1)))

    End Sub

    ''' <summary>
    ''' Prints a line that <see cref="getFrame"/> already rendered to the full console width: a
    ''' border, divider, or blank row. These bypass <see cref="mkMenuLine"/> because they are not
    ''' content to be framed, they are the frame
    ''' </summary>
    '''
    ''' <param name="line">
    ''' A complete, already-rendered menu line
    ''' </param>
    '''
    ''' <param name="cond">
    ''' Indicates whether the line should be printed <br /><br />
    ''' Optional, Default: <c> True </c>
    ''' </param>
    Private Sub printRenderedLine(line As String,
                         Optional cond As Boolean = True)

        If Not cond Then Return

        cwl(line)

    End Sub

    ''' <summary>
    ''' Prints a numbered menu option, padding <c> #. Name </c> to <see cref="menuItemLength"/>
    ''' before the description. An option that isn't printed doesn't use up a number.
    ''' </summary>
    '''
    ''' <param name="lineString1">
    ''' The name of the menu option
    ''' </param>
    '''
    ''' <param name="lineString2">
    ''' The description of the menu option
    ''' </param>
    '''
    ''' <param name="cond">
    ''' Indicates whether the option should be printed <br /><br />
    ''' Optional, Default: <c> True </c>
    ''' </param>
    Private Sub printMenuOpt(lineString1 As String,
                             lineString2 As String,
                    Optional cond As Boolean = True)

        If Not cond Then Return

        Dim sb As New StringBuilder($"{OptNum}. {lineString1}")
        padToEnd(sb, menuItemLength, "")
        cwl(mkMenuLine($"{sb}- {lineString2}", 1))
        OptNum += 1

    End Sub

    ''' <summary>
    ''' Returns the longest run of text that still fits between a menu line's two borders
    ''' </summary>
    '''
    ''' <remarks>
    ''' A left-aligned line spends two columns on the leading space and opener, one on the indent,
    ''' and two more on the trailing space and closer, so the text itself may occupy at most
    ''' <c> GetConsoleWidth() - 5 </c> columns
    ''' </remarks>
    '''
    ''' <returns>
    ''' The maximum framed line length, or <c> 0 </c> on a console too narrow to frame anything
    ''' </returns>
    Private Function MaxFramedLineLength() As Integer

        Return Math.Max(0, GetConsoleWidth() - 5)

    End Function

    ''' <summary>
    ''' Returns a menu line fit to the width of the console. Text longer than
    ''' <see cref="MaxFramedLineLength"/> gets the opener and the one-space indent but no closer,
    ''' whatever <paramref name="align"/> asks for, and runs past the right edge of the box.
    ''' </summary>
    ''' 
    ''' <param name="line">
    ''' The text to be printed 
    ''' </param>
    ''' 
    ''' <param name="align"> 
    ''' The alignment of the line to be printed: <br /> 
    ''' 
    ''' <list type="bullet">
    ''' 
    ''' <item>
    ''' <description> 
    ''' 0: centers the string 
    ''' </description> 
    ''' </item>
    ''' 
    ''' <item>
    ''' <description> 
    ''' 1: left-aligns the string
    ''' </description> 
    ''' </item>
    ''' 
    ''' <item>
    ''' <description> 
    ''' 2: builds a menu frame, ignoring <paramref name="line"/>
    ''' </description>
    ''' </item>
    ''' 
    ''' </list> 
    ''' </param>
    ''' 
    ''' <param name="borderInd"> 
    ''' Determines which characters should
    ''' create the border for the menuline: <br />
    ''' 
    ''' <list type="bullet">
    ''' 
    ''' <item>
    ''' <description> 
    ''' 0: Vertical lines 
    ''' </description> 
    ''' </item>
    ''' 
    ''' <item> 
    ''' <description> 
    ''' 1: Ceiling brackets 
    ''' </description> 
    ''' </item>
    ''' 
    ''' <item> 
    ''' <description> 
    ''' 2: Floor brackets 
    ''' </description>
    ''' </item>
    ''' 
    ''' <item>
    ''' <description> 
    ''' 3: Conjoining brackets 
    ''' </description> 
    ''' </item> 
    ''' 
    ''' </list>
    ''' 
    ''' <br /> Optional, Default: <c> 0 </c> 
    ''' </param>
    ''' 
    ''' <param name="fillBorder">
    ''' Indicates whether a frame line (<paramref name="align"/> <c> 2 </c>) is filled with
    ''' <c> ═ </c> rather than spaces <br /><br />
    ''' Optional, Default: <c> True </c>
    ''' </param>
    Private Function mkMenuLine(line As String,
                                align As Integer,
                                Optional borderInd As Integer = 0,
                                Optional fillBorder As Nullable(Of Boolean) = True) As String

        Dim out As New StringBuilder($" {Openers(borderInd)}")

        ' Text with nowhere to put a closing border still gets the opener and the standard
        ' one-space indent, so an overlong line starts in the same column as every other line in
        ' the box and the left edge stays unbroken. It simply runs past where the right border
        ' would have been. Centering is meaningless at this width, so these fall back to the
        ' left-aligned indent regardless of the requested alignment
        If line.Length > MaxFramedLineLength() Then Return out.Append(" "c).Append(line).ToString()

        Select Case align

            Case 0

                ' The interior spans the two columns after the opener through GetConsoleWidth() - 3,
                ' so centering means starting the text at 2 + (interiorWidth - line.Length) \ 2 --
                ' which reduces to the expression below. Integer division keeps odd-width cases
                ' landing consistently one column left rather than alternating with CInt's rounding
                padToEnd(out, (GetConsoleWidth() - line.Length) \ 2, Closers(borderInd))
                out.Append(line)
                padToEnd(out, GetConsoleWidth() - 2, Closers(borderInd))

            Case 1

                out.Append(" " & line)
                padToEnd(out, GetConsoleWidth() - 2, Closers(borderInd))

            Case 2

                padToEnd(out, GetConsoleWidth() - 2, Closers(borderInd), If(fillBorder, "═", " "))

        End Select

        Return out.ToString

    End Function

    ''' <summary>
    ''' Pads <paramref name="out"/> until it is <paramref name="targetLen"/> long, then appends
    ''' <paramref name="endline"/>, but only when <paramref name="targetLen"/> is the full line
    ''' width of <c> GetConsoleWidth() - 2 </c>
    ''' </summary>
    '''
    ''' <param name="out">
    ''' The text to be padded, changed in place
    ''' </param>
    ''' 
    ''' <param name="targetLen"> 
    ''' The length to which the text should be padded 
    ''' </param>
    ''' 
    ''' <param name="endline"> 
    ''' The closer character for the type of frame being built 
    ''' </param>
    ''' 
    ''' <param name="padStr">
    ''' The character(s) with which to pad the text <br /><br />
    ''' Optional, Default: <c> " " </c>
    ''' </param>

    Private Sub padToEnd(ByRef out As StringBuilder,
                               targetLen As Integer,
                               endline As String,
                      Optional padStr As String = " ")

        While out.Length < targetLen

            out.Append(padStr)

        End While

        If targetLen = GetConsoleWidth() - 2 Then out.Append(endline)

    End Sub

    ''' <summary> 
    ''' Replaces instances of the current directory in a path string with <c> ".." </c>
    ''' </summary>
    ''' 
    ''' <param name="dirStr">
    ''' A windows filesystem path 
    ''' </param>
    ''' 
    ''' <returns>
    ''' <paramref name="dirStr"/> with instances of the
    ''' current directory replaced with <c> ".." </c>
    ''' </returns>
    Public Function replDir(dirStr As String) As String

        Return dirStr.Replace(Environment.CurrentDirectory, "..")

    End Function

    ''' <summary>
    ''' Returns, as a String, <paramref name="defaultNumber"/> plus the weight of every
    ''' component in <paramref name="weightedComponents"/> that is <c> True </c>
    ''' </summary>
    '''
    ''' <param name="defaultNumber">
    ''' The menu number associated with the option
    ''' in winapp2ool's default, online configuration
    ''' </param>
    '''
    ''' <param name="weightedComponents">
    ''' A set of conditions which shift the
    ''' position of a menu option in the menu
    ''' </param>
    '''
    ''' <param name="weights">
    ''' The weight of each condition in <paramref name="weightedComponents"/>, at the same index.
    ''' It must be at least as long as <paramref name="weightedComponents"/>.
    ''' </param>
    Public Function computeMenuNumber(defaultNumber As Integer,
                                      weightedComponents As Boolean(),
                                      weights As Integer()) As String

        Dim out = defaultNumber

        For i = 0 To weightedComponents.Length - 1

            If weightedComponents(i) Then out += weights(i)

        Next

        Return out.ToString

    End Function


    ''' <summary>
    ''' Prints a framed menu line with text. An empty <paramref name="text"/> prints a blank line.
    ''' </summary>
    '''
    ''' <param name="text">The text to print</param>
    '''
    ''' <param name="centered">
    ''' Indicates whether the text should be centered <br /><br />
    ''' Optional, Default: <c> False </c>
    ''' </param>
    '''
    ''' <param name="condition">
    ''' Indicates whether the line should be printed <br /><br />
    ''' Optional, Default: <c> True </c>
    ''' </param>
    Public Sub PrintLine(text As String,
                Optional centered As Boolean = False,
                Optional condition As Boolean = True)

        printMenuLine(text, centered, condition)

    End Sub

    ''' <summary>
    ''' Prints a blank menu line
    ''' </summary>
    '''
    ''' <param name="condition">
    ''' Indicates whether the line should be printed <br /><br />
    ''' Optional, Default: <c> True </c>
    ''' </param>
    Public Sub PrintBlank(Optional condition As Boolean = True)

        printMenuLine(Nothing, cond:=condition)

    End Sub

    ''' <summary>
    ''' Prints a numbered menu option. An option that isn't printed doesn't use up a number.
    ''' </summary>
    '''
    ''' <param name="name">The name of the menu option</param>
    '''
    ''' <param name="description">The description of the menu option</param>
    '''
    ''' <param name="condition">
    ''' Indicates whether the option should be printed <br /><br />
    ''' Optional, Default: <c> True </c>
    ''' </param>
    Public Sub PrintOption(name As String,
                           description As String,
                  Optional condition As Boolean = True)

        printMenuOpt(name, description, condition)

    End Sub

    ''' <summary>
    ''' Prints a numbered menu option in a specific color
    ''' </summary>
    '''
    ''' <param name="name">
    ''' The name of the menu option
    ''' </param>
    '''
    ''' <param name="description">
    ''' The description of the menu option
    ''' </param>
    '''
    ''' <param name="color">
    ''' The color with which to print the option
    ''' </param>
    '''
    ''' <param name="condition">
    ''' Indicates whether the option should be printed <br /><br />
    ''' Optional, Default: <c> True </c>
    ''' </param>
    Public Sub PrintColoredOption(name As String,
                                  description As String,
                                  color As ConsoleColor,
                         Optional condition As Boolean = True)

        Dim startingColor = getForegroundColor()
        setForegroundColor(color)
        printMenuOpt(name, description, condition)
        setForegroundColor(startingColor)

    End Sub


    ''' <summary>
    ''' Prints a toggle as a numbered option whose description starts with <c> "Disable" </c>
    ''' when <paramref name="isEnabled"/> is <c> True </c> and <c> "Enable" </c> otherwise. It
    ''' prints green when enabled and red when not.
    ''' </summary>
    '''
    ''' <param name="name">The name of the toggle</param>
    '''
    ''' <param name="description">The description, printed after the Enable or Disable word</param>
    '''
    ''' <param name="isEnabled">Indicates whether the setting the toggle controls is enabled</param>
    '''
    ''' <param name="condition">
    ''' Indicates whether the toggle should be printed <br /><br />
    ''' Optional, Default: <c> True </c>
    ''' </param>
    Public Sub PrintToggle(name As String,
                           description As String,
                           isEnabled As Boolean,
                  Optional condition As Boolean = True)

        PrintColoredOption(name, enStr(isEnabled) & " " & description, GetRedGreen(Not isEnabled), condition)

    End Sub

    ''' <summary>
    ''' Prints a left-aligned menu line in yellow
    ''' </summary>
    '''
    ''' <param name="text">The warning text</param>
    '''
    ''' <param name="condition">
    ''' Indicates whether the warning should be printed <br /><br />
    ''' Optional, Default: <c> True </c>
    ''' </param>
    Public Sub PrintWarning(text As String,
                   Optional condition As Boolean = True)


        Dim startingColor = getForegroundColor()
        setForegroundColor(ConsoleColor.Yellow)
        printMenuLine(text, cond:=condition)
        setForegroundColor(startingColor)

    End Sub

    ''' <summary>
    ''' VT foreground color codes indexed by <c> ConsoleColor </c>'s underlying
    ''' integer value (0..15). The first eight, <c> Black </c> through <c> Gray </c>, map
    ''' to the base ANSI 30-37, and the last eight, <c> DarkGray </c> through <c> White </c>,
    ''' to the bright 90-97. <c> ConsoleColor.Yellow </c> is the bright variant, ANSI 93,
    ''' so what looks yellow on Windows comes out yellow on a VT terminal too.
    ''' </summary>
    Private ReadOnly VtForegroundCodes As Integer() = {
        30, 34, 32, 36, 31, 35, 33, 37,
        90, 94, 92, 96, 91, 95, 93, 97
    }

    ''' <summary>
    ''' VT escape that resets the foreground color to the terminal's default.
    ''' </summary>
    Private ReadOnly VtResetForeground As String = ChrW(&H1B) & "[39m"

    ''' <summary>
    ''' Tracked current foreground color, read once from <c> Console.ForegroundColor </c> on
    ''' first use and <c> Gray </c> if that read fails. On the VT path we cannot read it
    ''' back from the terminal, so we mirror the state ourselves. Without VT this also
    ''' avoids a <c> Console.ForegroundColor </c> read (one P/Invoke into
    ''' <c> GetConsoleScreenBufferInfo </c>) per save call.
    ''' </summary>
    Private _currentForeground As ConsoleColor = ConsoleColor.Gray
    Private _currentForegroundInitialized As Boolean = False

    Private Sub ensureForegroundInitialized()

        If _currentForegroundInitialized Then Return

        Try
            _currentForeground = Console.ForegroundColor
        Catch ex As IO.IOException
            ' Reading ForegroundColor can fail when stdout isn't a real
            ' console; fall back to the assumed default.
        End Try

        _currentForegroundInitialized = True

    End Sub

    ''' <summary>
    ''' Builds the VT foreground escape for a given <c> ConsoleColor </c>.
    ''' Returns the default-foreground reset for any value outside 0..15.
    ''' </summary>
    Private Function VtForegroundEscape(color As ConsoleColor) As String

        Dim idx = CInt(color)
        If idx < 0 OrElse idx > 15 Then Return VtResetForeground
        Return ChrW(&H1B) & "[" & VtForegroundCodes(idx).ToString() & "m"

    End Function

    Private Sub setForegroundColor(color As ConsoleColor)

        ensureForegroundInitialized()
        _currentForeground = color

        If TerminalCapabilities.HasVT Then

            ' Inline VT escape, with no flush and no syscall. It sits between the
            ' surrounding text in the active buffer (or in Console.Out
            ' directly when buffering is off).
            Dim escape = VtForegroundEscape(color)

            If _outputBuffer IsNot Nothing Then
                _outputBuffer.Append(escape)
            Else
                Console.Out.Write(escape)
            End If

            Return

        End If

        ' Legacy path: real Win32 attribute set. Must flush any buffered
        ' output first so the attribute applies to characters yet to come,
        ' not characters already in the in-memory buffer.
        FlushBufferIfActive()
        Console.ForegroundColor = color

    End Sub

    Private Function getForegroundColor() As ConsoleColor

        ensureForegroundInitialized()
        Return _currentForeground

    End Function

    ''' <summary>
    ''' Prints a framed menu line in the given color, then restores the previous color
    ''' </summary>
    '''
    ''' <param name="text">The text to print</param>
    '''
    ''' <param name="color">The color of the line</param>
    '''
    ''' <param name="centered">
    ''' Indicates whether the text should be centered <br /><br />
    ''' Optional, Default: <c> False </c>
    ''' </param>
    '''
    ''' <param name="condition">
    ''' Indicates whether the line should be printed <br /><br />
    ''' Optional, Default: <c> True </c>
    ''' </param>
    Public Sub PrintColored(text As String,
                            color As ConsoleColor,
                   Optional centered As Boolean = False,
                   Optional condition As Boolean = True)

        Dim startingColor = getForegroundColor()
        setForegroundColor(color)
        printMenuLine(text, centered, condition)
        setForegroundColor(startingColor)

    End Sub

    ''' <summary>
    ''' Opens a menu with top border
    ''' </summary>
    '''
    ''' <param name="solid">
    ''' Indicates whether the border is filled with <c> ═ </c> rather than spaces <br /><br />
    ''' Optional, Default: <c> True </c>
    ''' </param>
    Public Sub BeginMenu(Optional solid As Boolean = True)

        printRenderedLine(getFrame(1, solid))

    End Sub

    ''' <summary>
    ''' Prints the top of a complete menu: a top border, the centered module name, a divider,
    ''' the centered lines of <paramref name="centeredMenuText"/>, the menu prompt, and option
    ''' <c> 0 </c>, Exit. Throws if <paramref name="centeredMenuText"/> is <c> Nothing </c>.
    ''' </summary>
    '''
    ''' <param name="moduleName">The name shown in the header</param>
    '''
    ''' <param name="headerColor">The color of the header</param>
    '''
    ''' <param name="centeredMenuText">
    ''' Description lines printed centered under the header <br /><br />
    ''' Optional, Default: <c> Nothing </c> <br />
    ''' but leaving it out throws, so always pass an array
    ''' </param>
    Public Sub OpenMenu(moduleName As String,
                        headerColor As ConsoleColor,
               Optional centeredMenuText As String() = Nothing)

        BeginMenu()
        PrintColored(moduleName, headerColor, True)
        PrintDivider()

        For Each line In centeredMenuText

            PrintLine(line, True)

        Next

        PrintBlank()
        PrintLine("Menu: Enter a number to select", True)
        PrintBlank()
        OptNum = 0
        PrintOption("Exit", "Return to the previous menu", True)

    End Sub

    ''' <summary>
    ''' Closes a menu with bottom border
    ''' </summary>
    '''
    ''' <param name="filled">
    ''' Indicates whether the border is filled with <c> ═ </c> rather than spaces <br /><br />
    ''' Optional, Default: <c> True </c>
    ''' </param>
    Public Sub EndMenu(Optional filled As Boolean = True)

        printRenderedLine(getFrame(2, filled))

    End Sub

    ''' <summary>
    ''' Prints a section divider
    ''' </summary>
    '''
    ''' <param name="solid">
    ''' Indicates whether the divider is filled with <c> ═ </c> rather than spaces <br /><br />
    ''' Optional, Default: <c> True </c>
    ''' </param>
    Public Sub PrintDivider(Optional solid As Boolean = True)

        printRenderedLine(getFrame(3, solid))

    End Sub

    ''' <summary>
    ''' Returns <c> ConsoleColor.Red </c> if <paramref name="cond"/> is <c> True </c>,
    ''' <c> ConsoleColor.Green </c> otherwise
    ''' </summary>
    '''
    ''' <param name="cond">Indicates whether to return red, as for a missing file</param>
    Public Function GetRedGreen(cond As Boolean) As ConsoleColor

        Return If(cond, ConsoleColor.Red, ConsoleColor.Green)

    End Function

End Module
