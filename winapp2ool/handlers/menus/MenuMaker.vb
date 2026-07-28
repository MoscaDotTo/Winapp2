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
    ''' Defines the types of menu frames available
    ''' </summary>
    Public Enum FrameType

        ''' <summary>
        ''' 
        ''' </summary>
        Vertical = 0

        ''' <summary>
        ''' 
        ''' </summary>
        Top = 1

        ''' <summary>
        ''' 
        ''' </summary>
        Bottom = 2

        ''' <summary>
        ''' 
        ''' </summary>
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
    ''' The maximum length of the 'Name' half of a 
    ''' '#. Name - Description' style menu option
    ''' </summary>
    Private Property menuItemLength As Integer

    ''' <summary> 
    ''' Indicates that the menu header should be printed with color
    ''' </summary>
    Public Property ColorHeader As Boolean

    ''' <summary> 
    ''' The color with which the next header should be 
    ''' printed if <c> ColorHeader </c> is <c> True </c>
    ''' </summary>
    Public Property HeaderColor As ConsoleColor

    ''' <summary> 
    ''' Indicates that the application should not output or ask
    ''' input from the user except when encountering exceptions
    ''' <br/> Default: <c> False </c>
    ''' </summary>
    Public Property SuppressOutput As Boolean = False

    ''' <summary> 
    ''' Indicates that an exit from the current menu is pending 
    ''' </summary>
    Public Property ExitPending As Boolean

    ''' <summary> 
    ''' The text that appears in the top block of the menu 
    ''' </summary>
    Public Property MenuHeaderText As String

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <returns></returns>
    Public Property MenuHeaderTextColor As ConsoleColor = ConsoleColor.Red

    ''' <summary>
    ''' The number associated with the next
    ''' <c> Menu Option </c> that will be printed (if any)
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
    ''' The cached console window width, used to
    ''' avoid unneeded calls to <c> Console.WindowWidth </c>
    ''' </summary>
    Private _cachedWindowWidth As Integer = 120

    ''' <summary>
    ''' The time at which the console window width was last checked
    ''' </summary>
    Private _lastWidthCheckTime As DateTime = DateTime.Now

    ''' <summary>
    ''' When non-<c> Nothing </c>, <c> cwl </c> writes are appended to this buffer
    ''' instead of going to <c> Console.Out </c> directly. The buffer is flushed
    ''' as a single <c> Write </c> at the end of a render pass, and also flushed
    ''' on color changes so colored runs still apply to the right characters.
    ''' </summary>
    Private _outputBuffer As StringBuilder = Nothing

    ''' <summary>
    ''' Master switch for the buffered render path. When <c> False </c>, <c> BeginBuffered </c>
    ''' is a no-op and rendering reverts to the legacy per-line <c> Console.WriteLine </c> path.
    ''' Used by benchmarks and as a kill switch.
    ''' </summary>
    Public Property BufferingEnabled As Boolean = True

    ''' <summary>
    ''' Returns the current console window width, caching it for
    ''' 500 milliseconds at a time so as to avoid unneeded calls
    ''' </summary>
    ''' 
    ''' <returns>
    ''' The current console window width if not within the timeout
    ''' <br/> Otherwise, the cached console window width
    ''' </returns>
    Private Function GetConsoleWidth() As Integer

        ' The width is extremely unlikely to change during the printing process
        ' if ever, so only check it every 500 milliseconds at the most frequent
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
    ''' Displays a menu to and passes the user's input
    ''' over to be handled until the exit command is given 
    ''' 
    ''' <br/> Exiting a menu returns exactly one level 
    ''' up in the stack to the menu that called it 
    ''' 
    ''' <br/> Effectively the main event loop 
    ''' for anything built with <c> MenuMaker </c>
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
    ''' Indicates the maximum length of menu option names
    ''' <br/> Optional, Default: <c> 35 </c>
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

            FlushIfDirty2()

            setNextMenuHeaderText($"{name} closed")

        End Using

        gLog($"Exited {name}", leadr:=True)

    End Sub

    ''' <summary>
    ''' Prints a new line to the console
    ''' </summary>
    ''' 
    ''' <param name="condition">
    ''' Indicates whether the new line should be printed
    ''' Optional, Default: <c> True </c>
    ''' </param>
    Public Sub PrintNewLine(Optional condition As Boolean = True)

        If condition Then Console.WriteLine()

    End Sub

    ''' <summary> 
    ''' Prints a line to the console window if output is not currently being
    ''' suppressed and the given <c> <paramref name="cond"/> </c> is met
    ''' </summary>
    ''' 
    ''' <param name="msg"> 
    ''' The string to be printed
    ''' 
    ''' <br/> Optional, Default: <c> Nothing </c> 
    ''' </param>
    ''' 
    ''' <param name="cond"> 
    ''' Indicates the line should be printed 
    ''' 
    ''' <br/> Optional, Default: <c> True </c> 
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
    ''' Begins a buffered render pass. Subsequent <c> cwl </c> calls append to an
    ''' in-memory buffer instead of going to <c> Console.Out </c>. The buffer is
    ''' emitted as one <c> Write </c> when <c> FlushBuffered </c> is called, with
    ''' intermediate flushes around color changes so colored runs still apply to
    ''' the intended characters.
    ''' </summary>
    Public Sub BeginBuffered()

        If Not BufferingEnabled OrElse SuppressOutput Then Return
        If _outputBuffer IsNot Nothing Then Return

        _outputBuffer = New StringBuilder(4096)

    End Sub

    ''' <summary>
    ''' Ends a buffered render pass, writing the accumulated buffer to
    ''' <c> Console.Out </c> in a single call.
    ''' </summary>
    Public Sub FlushBuffered()

        If _outputBuffer Is Nothing Then Return

        Dim sb = _outputBuffer
        _outputBuffer = Nothing

        If sb.Length > 0 Then Console.Out.Write(sb.ToString())

    End Sub

    ''' <summary>
    ''' Flushes the in-flight buffer (if any) without ending the buffered pass.
    ''' Used so that subsequent state changes (color, cursor) apply to the
    ''' correct terminal position.
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
    ''' Waits for the users to press Enter if output
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
    ''' Clears the console when the given <c> <paramref name="cond"/> </c>
    ''' is <c> True </c>, output is not redirected, and output is not suppressed.
    ''' </summary>
    '''
    ''' <param name="cond">
    ''' Indicates that the console should be cleared
    '''
    ''' <br/> Optional, Default: <c> True </c>
    ''' </param>
    '''
    ''' <remarks>
    ''' On a VT-capable terminal (Windows 10 1607+ conhost, Windows Terminal,
    ''' anything UNIX-ish) this emits the VT clear escape, which is a single
    ''' write. If a buffered render is active, the escape is appended to the
    ''' buffer so the clear and the new menu arrive at the terminal in one
    ''' atomic write — no flicker.
    '''
    ''' On legacy consoles (XP / Vista / 7 / pre-1607 Win10) falls back to
    ''' <c> Console.Clear </c>. When output is redirected (test runner,
    ''' pipeline silent mode, piped to a file) the call is a no-op — clearing
    ''' a non-console sink would corrupt downstream output.
    ''' </remarks>
    Public Sub clrConsole(Optional cond As Boolean = True)

        If Not cond OrElse SuppressOutput Then Return

        ' VT path: an escape sequence written through Console.Out is safe even
        ' if output happens to be redirected to a pipe/file — the caller opted
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
            ' Belt-and-braces — IsOutputRedirected should already cover this.
        End Try

    End Sub

    ''' <summary> 
    ''' Returns an empty menu line, or a variety of filled menu lines 
    ''' </summary>
    ''' 
    ''' <param name="frameNum"> 
    ''' Indicates which frame should be returned <br/>
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
    ''' <br/> Optional, Default: <c> 0 </c>
    ''' </param>
    ''' 
    ''' <returns> 
    ''' A String containing the menuFrame requested
    ''' by <c> <paramref name="frameNum"/> </c>
    ''' </returns>
    Private Function getFrame(Optional frameNum As Integer = 0,
                              Optional fillFrame As Nullable(Of Boolean) = False) As String

        Return mkMenuLine("", 2, frameNum, fillFrame)

    End Function



    ''' <summary>
    ''' Overrides the next menu's default header with with <c> <paramref name="txt"/> </c>
    ''' <br />
    ''' Useful for delivering a status update or error message between menues or modules
    ''' </summary>
    ''' 
    ''' <param name="txt">
    ''' 
    ''' </param>
    ''' 
    ''' <param name="cond">
    ''' 
    ''' </param>
    ''' 
    ''' <param name="printColor">
    ''' 
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
    ''' Informs a user when an action is unable to proceed due to a condition
    ''' </summary>
    ''' 
    ''' <param name="cond"> 
    ''' Indicates that an action should be denied 
    ''' </param>
    ''' 
    ''' <param name="errText"> 
    ''' The error text to be printed in the menu header 
    ''' </param>
    Public Function denyActionWithHeader(cond As Boolean,
                                         errText As String) As Boolean

        setNextMenuHeaderText(errText, cond)

        Return cond

    End Function

    ''' <summary> 
    ''' Returns the inverse state of a given boolean as a String
    ''' </summary>
    ''' 
    ''' <param name="setting">
    ''' A <c> module setting </c> whose state will be observed 
    ''' </param>
    ''' 
    ''' <returns> 
    ''' <c> "Disable" </c> if
    ''' <c> <paramref name="setting"/> </c> is
    ''' <c> True </c>,
    ''' 
    ''' <br/> <c> "Enable" </c> otherwise 
    ''' </returns>
    Public Function enStr(setting As Nullable(Of Boolean)) As String

        Return If(setting, "Disable", "Enable")

    End Function

    ''' <summary> 
    ''' Enforces that <c> initMenu </c> exit the current 
    ''' level in the stack on the next iteration of its loop
    ''' </summary>
    Public Sub exitModule()

        ExitPending = True

    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    Public Sub resetMenuNumbering()

        OptNum = 0

    End Sub

    ''' <summary> 
    ''' Prints a line bounded by vertical menu frames, or an empty menu line
    ''' if <c> <paramref name="lineString"/> </c> is <c> Nothing </c>
    ''' </summary>
    ''' 
    ''' <param name="lineString"> 
    ''' The text to be printed 
    ''' 
    ''' <br/> Optional, Default: <c> Nothing </c> 
    ''' </param>
    ''' 
    ''' <param name="isCentered"> 
    ''' Indicates that the printed text should be centered 
    ''' 
    ''' <br/> Optional, Default: <c> False </c> 
    ''' </param>
    ''' 
    ''' <param name="cond">
    ''' Indicates that the line should be printed <br />
    ''' Optional, Default: <c> True </c>
    ''' </param>
    Private Sub printMenuLine(Optional lineString As String = Nothing,
                              Optional isCentered As Boolean = False,
                              Optional cond As Boolean = True)

        If Not cond Then Return

        If lineString = Nothing Then lineString = getFrame()
        cwl(mkMenuLine(lineString, If(isCentered, 0, 1)))

    End Sub

    ''' <summary> 
    ''' Prints a numbered menu option after padding it to a set length 
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
    ''' Indicates that the option should be printed <br />
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
    ''' Constructs a menu line fit to the width of the console 
    ''' </summary>
    ''' 
    ''' <param name="line">
    ''' The text to be printed 
    ''' </param>
    ''' 
    ''' <param name="align"> 
    ''' The alignment of the line to be printed: <br/> 
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
    ''' 1: leftaligns the string 
    ''' </description> 
    ''' </item>
    ''' 
    ''' <item>
    ''' <description> 
    ''' 2: prints a menu frame 
    ''' </description> 
    ''' </item>
    ''' 
    ''' </list> 
    ''' </param>
    ''' 
    ''' <param name="borderInd"> 
    ''' Determines which characters should
    ''' create the border for the menuline: <br/>
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
    ''' <br/> Optional, Default: <c> 0 </c> 
    ''' </param>
    ''' 
    ''' <param name="fillBorder"> 
    ''' Indicates that top and bottom borders 
    ''' should be printed when printing menuframes
    ''' </param>
    Private Function mkMenuLine(line As String,
                                align As Integer,
                                Optional borderInd As Integer = 0,
                                Optional fillBorder As Nullable(Of Boolean) = True) As String

        If line.Length >= GetConsoleWidth() - 1 Then Return line
        Dim out As New StringBuilder($" {Openers(borderInd)}")

        Select Case align

            Case 0

                padToEnd(out, CInt((((GetConsoleWidth() - line.Length) / 2) + 2)), Closers(borderInd))
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
    ''' Pads a given string until it is a given length 
    ''' </summary>
    ''' 
    ''' <param name="out"> 
    ''' The text to be padded 
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
    ''' The character(s) with which to pad the text 
    ''' <br/> Default: <c> " " </c> (space character)
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
    ''' <c> <paramref name="dirStr"/> </c> with instances of the
    ''' current directory replaced with <c> ".." </c> 
    ''' </returns>
    ''' 
    ''' Docs last updated: 2020-09-04 | Code last updated: 2020-09-04
    Public Function replDir(dirStr As String) As String

        Return dirStr.Replace(Environment.CurrentDirectory, "..")

    End Function

    ''' <summary> 
    ''' Determines the number currently associated 
    ''' with a particular menu option
    ''' </summary>
    '''
    ''' <param name="defaultNumber"> 
    ''' The menu number associated with the option 
    ''' in winapp2ool's default, online configuration
    ''' </param>
    '''
    ''' <param name="weightedComponents"> 
    ''' A set of parameters which influence the 
    ''' position of a menu option in the menu 
    ''' </param>
    '''
    ''' <param name="weights"> 
    ''' The weights correlating to each <c>Component</c>
    ''' in <c><paramref name="weightedComponents"/> </c> 
    ''' </param>
    ''' 
    ''' Docs last updated: 2022-11-21 | Code last updated: 2022-11-21
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
    ''' Prints a simple menu line with text
    ''' </summary>
    Public Sub PrintLine(text As String,
                Optional centered As Boolean = False,
                Optional condition As Boolean = True)

        printMenuLine(text, centered, condition)

    End Sub

    ''' <summary>
    ''' Prints a blank menu line
    ''' </summary>
    Public Sub PrintBlank(Optional condition As Boolean = True)

        printMenuLine(Nothing, cond:=condition)

    End Sub

    ''' <summary>
    ''' Prints a numbered menu option
    ''' </summary>
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
    ''' Indicates whether the option should be printed <br />
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
    ''' Prints a toggle option (Enable/Disable)
    ''' </summary>
    Public Sub PrintToggle(name As String,
                           description As String,
                           isEnabled As Boolean,
                  Optional condition As Boolean = True)

        PrintColoredOption(name, enStr(isEnabled) & " " & description, GetRedGreen(Not isEnabled), condition)

    End Sub

    ''' <summary>
    ''' Prints a colored warning line
    ''' </summary>
    Public Sub PrintWarning(text As String,
                   Optional condition As Boolean = True)


        Dim startingColor = getForegroundColor()
        setForegroundColor(ConsoleColor.Yellow)
        printMenuLine(text, condition)
        setForegroundColor(startingColor)

    End Sub

    ''' <summary>
    ''' VT foreground color codes indexed by <c> ConsoleColor </c>'s underlying
    ''' integer value (0..15). Windows' "Dark" variants map to ANSI 30-37
    ''' (the base 8); the non-dark variants map to 90-97 (bright). Note that
    ''' Windows' <c> ConsoleColor.Yellow </c> is the bright variant — ANSI 93 —
    ''' so what looks "yellow" on Windows comes out yellow on a VT terminal too.
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
    ''' Tracked current foreground color. On the VT path we cannot read it
    ''' back from the terminal, so we mirror the state ourselves. On the
    ''' legacy path this also avoids a <c> Console.ForegroundColor </c> read
    ''' (one P/Invoke into <c> GetConsoleScreenBufferInfo </c>) per save call.
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
    ''' Returns the default-foreground reset when the color is the sentinel
    ''' <c> -1 </c> used to indicate "restore default."
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

            ' Inline VT escape — no flush, no syscall. Sits between the
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
    ''' Prints colored text
    ''' </summary>
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
    ''' Indicates whether the divider line should be solid
    ''' Optional, Default: <c> True </c> (solid)
    ''' </param>
    ''' 
    ''' Docs last updated: 2025-08-06 | Code last updated: 2025-08-06
    Public Sub BeginMenu(Optional solid As Boolean = True)

        printMenuLine(getFrame(1, solid), cond:=True)

    End Sub

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
    Public Sub EndMenu(Optional filled As Boolean = True)

        printMenuLine(getFrame(2, filled), cond:=True)

    End Sub

    ''' <summary>
    ''' Prints a section divider
    ''' </summary>
    '''
    ''' <param name="solid">
    ''' Indicates whether the divider line should be solid
    ''' Optional, Default: <c> True </c> (solid)
    ''' </param>
    Public Sub PrintDivider(Optional solid As Boolean = True)

        printMenuLine(getFrame(3, solid))

    End Sub

    ''' <summary>
    '''
    ''' </summary>
    ''' 
    ''' <param name="cond">
    ''' 
    ''' </param>
    ''' <returns></returns>
    Public Function GetRedGreen(cond As Boolean) As ConsoleColor

        Return If(cond, ConsoleColor.Red, ConsoleColor.Green)

    End Function

End Module
