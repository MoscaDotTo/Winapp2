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
''' <summary>
''' Provides functions and methods for presenting and maintaining menus to a user
''' </summary>
Module MenuMaker
    '''<summary>String literal containing an instruction to press any key to return </summary>
    Public ReadOnly Property anyKeyStr As String = "Press any key to return to the menu."
    '''<summary>String literal containing an error message informing the user their input was invalid</summary>
    Public ReadOnly Property invInpStr As String = "Invalid input. Please try again."
    '''<summary>String literal containing an instruction for the user to provide input</summary>
    Public ReadOnly Property promptStr As String = "Enter a number, or leave blank to run the default: "
    '''<summary>The maximum length of the portion of the first half of a '#. Option - Description' style menu line</summary>
    Private Property menuItemLength As Integer
    '''<summary>Indicates that the menu header should be printed with color </summary>
    Public Property colorHeader As Boolean
    '''<summary>The color with which the next header should be printed if color is used</summary>
    Public Property headerColor As ConsoleColor

    '''<summary>Holds the current option number for the menu instance</summary>
    Private Property optNum As Integer = 0
    ''' <summary>When enabled, prevents winapp2ool from outputting to the console or asking for input (usually)</summary>
    Public Property SuppressOutput As Boolean = False
    ''' <summary> True if there is a pending exit from the menu</summary>
    Public Property ExitCode As Boolean
    ''' <summary> Holds the text that appears in the top block of the menu </summary>
    Public Property MenuHeaderText As String
    '''<summary>Frame characters used to open a menu line</summary>
    Private ReadOnly Property openers As String() = {"║", "╔", "╚", "╠"}
    '''<summary>Frame characters used to close a menu line</summary>
    Private ReadOnly Property closers As String() = {"║", "╗", "╝", "╣"}

    '''<summary>Returns a menuframe</summary>
    '''<param name="frameNum">Indicates which frame should be returned. <br />
    '''0: empty line, 1: top, 2: bottom, 3: conjoiner <br /> Default: 0</param>
    Public Function getFrame(Optional frameNum As Integer = 0) As String
        Return mkMenuLine("", "f", frameNum)
    End Function

    ''' <summary>Inserts text into the menu header</summary>
    ''' <param name="txt">The text to appear in the header</param>
    ''' <param name="cHeader">Optional boolean indicating whether or not the text should be colored <br />Default: False</param>
    ''' <param name="cond">Optional boolean indicating whether or not the header should be set <br /> Default: True</param>
    ''' <param name="printColor">Optional ConsoleColor with which the header should be colored <br/> Default: Red</param>
    Public Sub setHeaderText(txt As String, Optional cHeader As Boolean = False, Optional cond As Boolean = True, Optional printColor As ConsoleColor = ConsoleColor.Red)
        If cond Then
            MenuHeaderText = txt
            colorHeader = cHeader
            headerColor = printColor
        End If
    End Sub

    ''' <summary>Initializes the menu</summary>
    ''' <param name="topper">The text to be displayed at the top of the menu screen</param>
    ''' <param name="itemlen">The length in characters that should comprise the first bloc of options in the menu <br />Default: 35</param>
    Public Sub initMenu(topper As String, Optional itemlen As Integer = 35)
        ExitCode = False
        setHeaderText(topper)
        menuItemLength = itemlen
    End Sub

    ''' <summary>Sets the menu header to an error string conditionally, returns the given condition</summary>
    ''' <param name="cond">The condition under which the error text should be printed</param>
    ''' <param name="errText">The error text to be printed in the menu header</param>
    Public Function denyActionWithTopper(cond As Boolean, errText As String) As Boolean
        setHeaderText(errText, True, cond)
        Return cond
    End Function

    ''' <summary>Prints a menu line, option, or reset string, conditionally</summary>
    ''' <param name="cond">The optional condition under which to print (default: true)</param>
    ''' <param name="printType">The type of menu information to print. <br /> 0: line, 1: opt, 2: Reset Settings, 3: Box w/ centered text, 
    ''' 4: red error string, 5: enable/disable string</param>
    ''' <param name="str1">The first string or half string to be printed</param>
    ''' <param name="optString">The second half string to be printed for menu options <br /> Default: ""</param>
    ''' <param name="leadingBlank">Optional condition under which the print should be buffered with a leading blank line <br /> Default: False</param>
    ''' <param name="trailingBlank">Optional condition under which the print should be buffered with a trailing blank line <br /> Default: False</param>
    ''' <param name="isCentered">Optional condition specifying whether or not the text should be centered <br /> Default: False</param>
    ''' <param name="closeMenu">Optional condition specifying whether a menu is waiting to be closed <br /> Default: False</param>
    ''' <param name="enStrCond">Optional condition that helps color Enabled/Disabled lines <br /> Default: False</param>
    ''' <param name="colorLine">Optional condition that indicates we want to color lines without Enabled/Disabled strings in them <br /> Default: False</param>
    ''' <param name="arbitraryColor">Optional ConsoleColor to be used when printing with colorLine but wanting to use a color outside red/green <br /> Default: Nothing</param>
    ''' <param name="useArbitraryColor">Optional condition specifying whether or not the line should be colored using the value provided by arbitraryColor <br /> Default: False</param>
    Public Sub print(printType As Integer, str1 As String, Optional optString As String = "", Optional cond As Boolean = True,
                     Optional leadingBlank As Boolean = False, Optional trailingBlank As Boolean = False, Optional isCentered As Boolean = False,
                     Optional closeMenu As Boolean = False, Optional enStrCond As Boolean = False, Optional colorLine As Boolean = False,
                     Optional useArbitraryColor As Boolean = False, Optional arbitraryColor As ConsoleColor = Nothing)
        If Not cond Then Exit Sub
        print(0, Nothing, cond:=leadingBlank)
        If colorLine Then Console.ForegroundColor = If(useArbitraryColor, arbitraryColor, If(enStrCond, ConsoleColor.Green, ConsoleColor.Red))
        Select Case printType
            ' Prints lines
            Case 0
                printMenuLine(str1, isCentered)
            ' Prints options
            Case 1
                printMenuOpt(str1, optString)
            ' Prints the Reset Settings option
            Case 2
                print(1, "Reset Settings", $"Restore {str1}'s settings to their default state", leadingBlank:=True)
            ' Prints a box with centered text
            Case 3
                print(0, tmenu(str1), isCentered:=True, closeMenu:=True)
            ' Prints a red error string into the menu
            Case 4
                Console.ForegroundColor = ConsoleColor.Red
                print(0, $"{str1}, some functions will not be available.", trailingBlank:=True, isCentered:=True)
            ' Colored line printing for enable/disable menu options
            Case 5
                Console.ForegroundColor = If(enStrCond, ConsoleColor.Green, ConsoleColor.Red)
                print(1, str1, $"{enStr(enStrCond)} {optString}")
        End Select
        Console.ResetColor()
        print(0, Nothing, cond:=trailingBlank)
        print(0, getFrame(2), cond:=closeMenu)
    End Sub

    ''' <summary>Returns the inverse state of a given setting as a String</summary>
    ''' <param name="setting">The setting whose state will be reported</param>
    Public Function enStr(setting As Boolean) As String
        Return If(setting, "Disable", "Enable")
    End Function

    ''' <summary>Exits a menu or module by flipping the exitCode to true</summary>
    Public Sub exitModule()
        ExitCode = True
    End Sub

    ''' <summary>Prints the top of the menu (containing the topper), any description text provided, the menu prompt, and the exit option</summary>
    ''' <param name="descriptionItems">Items describing the menu</param>
    ''' <param name="printExit">Optional boolean indicating whether an option to exit should be printed <br /> Default: True</param>
    Public Sub printMenuTop(descriptionItems As String(), Optional printExit As Boolean = True)
        print(0, tmenu(MenuHeaderText), isCentered:=True, colorLine:=colorHeader, useArbitraryColor:=True, arbitraryColor:=headerColor)
        print(0, getFrame(3), colorLine:=colorHeader, useArbitraryColor:=True, arbitraryColor:=headerColor)
        Console.ResetColor()
        For Each line In descriptionItems
            print(0, line, isCentered:=True)
        Next
        print(0, getFrame() & Environment.NewLine & mkMenuLine("Menu: Enter a number to select", "c") & Environment.NewLine & getFrame())
        optNum = 0
        print(1, "Exit", "Return to the menu", printExit)
    End Sub

    ''' <summary>Prints a line from the menu that does not contain an option</summary>
    ''' <param name="lineString">The text to be printed. <br/> Default: Nothing, will print an empty menu frame with no text in it</param>
    ''' <param name="isCentered">Optional boolean indicating whether or not the text should be centered <br /> Default: False</param>
    Private Sub printMenuLine(Optional lineString As String = Nothing, Optional isCentered As Boolean = False)
        If lineString = Nothing Then lineString = getFrame()
        cwl(mkMenuLine(lineString, If(isCentered, "c", "l")))
    End Sub

    ''' <summary>Prints a numbered menu option after padding it to a set length</summary>
    ''' <param name="lineString1">The first part of the menu option</param>
    ''' <param name="lineString2">The second part of the menu option</param>
    Private Sub printMenuOpt(lineString1 As String, lineString2 As String)
        lineString1 = $"{optNum}. {lineString1}"
        padToEnd(lineString1, menuItemLength, "")
        cwl(mkMenuLine($"{lineString1}- {lineString2}", "l"))
        optNum += 1
    End Sub

    ''' <summary>Flips the exitCode boolean so we can return to the previous menu when desired</summary>
    Public Sub revertMenu()
        ExitCode = Not ExitCode
    End Sub

    ''' <summary>Constructs a menu line fit to the width of the console</summary>
    ''' <param name="line">The line to be printed</param>
    ''' <param name="align">The alignment of the line to be printed: <br /> 'l' for Left, 'c' for Centre, 'f' for Frame</param>
    ''' <param name="borderInd">Determines which characters should create the border for the menuline: <br />
    ''' 0: Vertical lines, 1: ceiling brackets, 2: floor brackets, 3: conjoining brackets</param>
    Private Function mkMenuLine(line As String, align As String, Optional borderInd As Integer = 0) As String
        If line.Length >= Console.WindowWidth - 1 Then Return line
        Dim out = $" {openers(borderInd)}"
        Select Case align
            Case "c"
                padToEnd(out, CInt((((Console.WindowWidth - line.Length) / 2) + 2)), closers(borderInd))
                out += line
                padToEnd(out, Console.WindowWidth - 2, closers(borderInd))
            Case "l"
                out += " " & line
                padToEnd(out, Console.WindowWidth - 2, closers(borderInd))
            Case "f"
                padToEnd(out, Console.WindowWidth - 2, closers(borderInd), If(borderInd = 0, " ", "═"))
        End Select
        Return out
    End Function

    ''' <summary>Pads a given string with spaces</summary>
    ''' <param name="out">The string to be padded</param>
    ''' <param name="targetLen">The end length to which the string should be padded</param>
    ''' <param name="endline">The closer character for the type of frame being built</param>
    ''' <param name="padChar">The character with which to pad the line <br /> Default: " " (space character)</param>
    Private Sub padToEnd(ByRef out As String, targetLen As Integer, endline As String, Optional padChar As String = " ")
        While out.Length < targetLen
            out += padChar
        End While
        If targetLen = Console.WindowWidth - 2 Then out += endline
    End Sub

    ''' <summary>Returns a String containing the topmost part of the menu with no bottom</summary>
    ''' <param name="text">The String to be printed in the faux menu header</param>
    Public Function tmenu(text As String) As String
        Dim out = getFrame(1) & Environment.NewLine
        out += mkMenuLine(text, "c")
        Return out
    End Function

    ''' <summary>Replaces instances of the current directory in a path string ".."</summary>
    ''' <param name="dirStr">A String containing a windows path</param>
    Public Function replDir(dirStr As String) As String
        Return dirStr.Replace(Environment.CurrentDirectory, "..")
    End Function

    ''' <summary>Prints a line with a string if we're not suppressing output.</summary>
    ''' <param name="msg">The string to be printed <br />Default: Nothing</param>
    ''' <param name="cond">The optional condition under which to print the line <br /> Default: True</param>
    Public Sub cwl(Optional msg As String = Nothing, Optional cond As Boolean = True)
        If cond And Not SuppressOutput Then Console.WriteLine(msg)
    End Sub

    ''' <summary>Clears the console conditionally when not running unit tests</summary>
    ''' <param name="cond">Optional Boolean specifying whether or not the console should be cleared <br /> Default: True</param>
    Public Sub clrConsole(Optional cond As Boolean = True)
        ' Do not clear the console during unit tests because there isnt one and the invalid handler throws an IO Exception
        If cond And Not SuppressOutput And Not Console.Title.Contains("testhost.x86") Then Console.Clear()
    End Sub

    ''' <summary>Initializes a module's menu, prints it, and handles the user input. Effectively the main event loop for anything built with MenuMaker</summary>
    ''' <param name="name">The name of the module</param>
    ''' <param name="showMenu">The function that prints the module's menu</param>
    ''' <param name="handleInput">The function that handles the module's input</param>
    Public Sub initModule(name As String, showMenu As Action, handleInput As Action(Of String))
        gLog("", ascend:=True)
        gLog($"Loading module {name}")
        initMenu(name)
        Try
            Do Until ExitCode
                clrConsole()
                showMenu()
                Console.Write(Environment.NewLine & promptStr)
                handleInput(Console.ReadLine)
            Loop
            revertMenu()
            setHeaderText($"{name} closed")
            gLog($"Exiting {name}", descend:=True, leadr:=True)
        Catch ex As Exception
            exc(ex)
        End Try
    End Sub
End Module