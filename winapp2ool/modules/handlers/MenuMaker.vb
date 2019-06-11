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
    ' Basic menu frames & Strings
    Public Const menuStr00 As String = " ╔══════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╗"
    Public Const menuStr01 As String = " ║                                                                                                                          ║"
    Public Const menuStr02 As String = " ╚══════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╝"
    Public Const menuStr03 As String = " ╠══════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╣"
    Public menuStr04 As String = aNL(menu(menuStr01)) & aNL(mkMenuLine("Menu: Enter a number to select", "c")) & mkMenuLine(menuStr01, "")
    Public Const anyKeyStr As String = "Press any key to return to the menu."
    Public Const invInpStr As String = "Invalid input. Please try again."
    Public Const promptStr As String = "Enter a number, or leave blank to run the default: "
    ' The maximum length of the portion of the first half of a '#. Option - Description' style menu line
    Dim menuItemLength As Integer
    '''<summary>Indicates that the last operation was unsucessful</summary>
    Public Property lastOpWasErr As Boolean
    ' Holds the current option number at any given moment
    Dim optNum As Integer = 0
    ''' <summary>When enabled, prevents winapp2ool from outputting to the console or asking for input (usually)</summary>
    Public Property SuppressOutput As Boolean = False
    ''' <summary> True if there is a pending exit from the menu</summary>
    Public Property ExitCode As Boolean
    ''' <summary> Holds the text that appears in the top block of the menu </summary>
    Public Property MenuHeaderText As String

    ''' <summary>Inserts text into the menu header</summary>
    ''' <param name="txt">The text to appear in the header</param>
    ''' <param name="hasErr">The boolean indicating whether or not the text should be colored red</param>
    Public Sub setHeaderText(txt As String, Optional hasErr As Boolean = False)
        MenuHeaderText = txt
        lastOpWasErr = hasErr
    End Sub

    ''' <summary>Initializes the menu</summary>
    ''' <param name="topper">The text to be displayed at the top of the menu screen</param>
    ''' <param name="itemlen">The length in characters that should comprise the first bloc of options in the menu</param>
    Public Sub initMenu(topper As String, Optional itemlen As Integer = 35)
        ExitCode = False
        setHeaderText(topper)
        menuItemLength = itemlen
    End Sub

    ''' <summary>Sets the menu header to an error string conditionally, returns the given condition</summary>
    ''' <param name="cond">The condition under which the error text should be printed</param>
    ''' <param name="errText">The error text to be printed in the menu header</param>
    Public Function denyActionWithTopper(cond As Boolean, errText As String) As Boolean
        If cond Then setHeaderText(errText, True)
        Return cond
    End Function

    ''' <summary>Prints a menu line, option, or reset string, conditionally</summary>
    ''' <param name="cond">The optional condition under which to print (default: true)</param>
    ''' <param name="printType">The type of menu information to print</param>
    ''' <param name="str1">The first string or half string to be printed</param>
    ''' <param name="optString">The second half string to be printed for menu options</param>
    ''' <param name="leadingBlank">Optional condition under which the print should be buffered with a leading blank line (default: false)</param>
    ''' <param name="trailingBlank">Optional condition under which the print should be buffered with a trailing blank line (default: false)</param>
    ''' <param name="isCentered">Optional condition specifying whether or not the text should be centered (default: false)</param>
    ''' <param name="closeMenu">Optional condition specifying whether a menu is waiting to be closed (default: false)</param>
    ''' <param name="enStrCond">Optional condition that helps color Enabled/Disabled lines</param>
    ''' <param name="colorLine">Optional condition that indicates we want to color lines without Enabled/Disabled strings in them</param>
    Public Sub print(printType As Integer, str1 As String, Optional optString As String = "", Optional cond As Boolean = True,
                     Optional leadingBlank As Boolean = False, Optional trailingBlank As Boolean = False, Optional isCentered As Boolean = False,
                     Optional closeMenu As Boolean = False, Optional enStrCond As Boolean = False, Optional colorLine As Boolean = False)
        If cond And leadingBlank Then printMenuLine()
        If colorLine Then Console.ForegroundColor = If(enStrCond, ConsoleColor.Green, ConsoleColor.Red)
        Select Case True
            ' Prints lines
            Case cond And printType = 0
                printMenuLine(str1, isCentered)
            ' Prints options
            Case cond And printType = 1
                printMenuOpt(str1, optString)
            ' Prints the Reset Settings option
            Case cond And printType = 2
                print(1, "Reset Settings", $"Restore {str1}'s settings to their default state", leadingBlank:=True)
            ' Prints a box with centered text
            Case cond And printType = 3
                print(0, tmenu(str1), isCentered:=True, closeMenu:=True)
            ' Prints a red error string into the menu
            Case cond And printType = 4
                Console.ForegroundColor = ConsoleColor.Red
                print(0, $"{str1}, some functions will not be available.", trailingBlank:=True, isCentered:=True)
            ' Colored line printing for enable/disable menu options
            Case cond And printType = 5
                Console.ForegroundColor = If(enStrCond, ConsoleColor.Green, ConsoleColor.Red)
                print(1, str1, $"{enStr(enStrCond)} {optString}")
        End Select
        Console.ResetColor()
        If cond And trailingBlank Then printMenuLine()
        If cond And closeMenu Then printMenuLine(menuStr02)
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
    ''' <param name="printExit">The boolean representing whether an option to exit should be printed</param>
    Public Sub printMenuTop(descriptionItems As String(), Optional printExit As Boolean = True)
        If lastOpWasErr Then Console.ForegroundColor = ConsoleColor.Red
        printMenuLine(tmenu(MenuHeaderText))
        printMenuLine(menuStr03)
        Console.ResetColor()
        For Each line In descriptionItems
            print(0, line, isCentered:=True)
        Next
        printMenuLine(menuStr04)
        optNum = 0
        print(1, "Exit", "Return to the menu", printExit)
    End Sub

    ''' <summary>Constructs and returns to the calling function a new menu String</summary>
    ''' <param name="lineString">A string to be made into a menu line</param>
    Public Function menu(lineString As String) As String
        Return mkMenuLine(lineString, "l")
    End Function

    ''' <summary>Prints a line in a menu</summary>
    ''' <param name="lineString"></param>
    Public Sub printMenuLine(Optional lineString As String = menuStr01)
        cwl(menu(lineString))
    End Sub

    ''' <summary>Constructs a menu string with a given alignment</summary>
    ''' <param name="lineString">The line to be printed</param>
    ''' <param name="isCentered">The boolean indicating whether the line text should be centered</param>
    Public Function menu(lineString As String, Optional isCentered As Boolean = False) As String
        Return mkMenuLine(lineString, If(isCentered, "c", "l"))
    End Function

    ''' <summary>Prints a menu string with a given alignment</summary>
    ''' <param name="lineString"></param>
    ''' <param name="isCenteredLine">The boolean indicating whether the line text should be centered</param>
    Public Sub printMenuLine(lineString As String, Optional isCenteredLine As Boolean = False)
        cwl(menu(lineString, isCenteredLine))
    End Sub

    ''' <summary>Prints a numbered menu option after padding it to a set length</summary>
    ''' <param name="lineString1">The first part of the menu option</param>
    ''' <param name="lineString2">The second part of the menu option</param>
    Public Sub printMenuOpt(lineString1 As String, lineString2 As String)
        lineString1 = $"{optNum}. {lineString1}"
        While lineString1.Length < menuItemLength
            lineString1 += " "
        End While
        cwl(menu($"{lineString1}- {lineString2}"))
        optNum += 1
    End Sub

    ''' <summary>Flips the exitCode boolean so we can return to the menu when desired</summary>
    Public Sub revertMenu()
        ExitCode = Not ExitCode
    End Sub

    ''' <summary>Forces the exitCode to be False</summary>
    Public Sub undoAnyPendingExits()
        ExitCode = False
    End Sub

    ''' <summary>Constructs a menu line fit to the width of the console</summary>
    ''' <param name="line">The line to be printed</param>
    ''' <param name="align">The alignment of the line to be printed. 'l' for Left or 'c' for Centre</param>
    Public Function mkMenuLine(line As String, align As String) As String
        If line.Length >= 125 Then Return line
        Dim out = " ║"
        Select Case align
            Case "c"
                padToEnd(out, CInt((((124 - line.Length) / 2) + 2)))
                out += line
                padToEnd(out, 124)
            Case "l"
                out += " " & line
                padToEnd(out, 124)
        End Select
        Return out
    End Function

    ''' <summary>Pads a given string with spaces</summary>
    ''' <param name="out">The string to be padded</param>
    ''' <param name="targetLen">The end length to which the string should be padded</param>
    Private Sub padToEnd(ByRef out As String, targetLen As Integer)
        While out.Length < targetLen
            out += " "
        End While
        If targetLen = 124 Then out += "║"
    End Sub

    ''' <summary>Prints a box with a single message inside it</summary>
    ''' <param name="text">The string to be printed in the box</param>
    Public Function bmenu(text As String) As String
        Dim out = aNL(menu(menuStr00))
        out += aNL(menu(text, True))
        out += menu(menuStr02)
        Return out
    End Function

    ''' <summary>Prints the topmost part of the menu with no bottom</summary>
    ''' <param name="text">The String to be printed in the faux menu header</param>
    Public Function tmenu(text As String) As String
        Dim out As String = aNL(menu(menuStr00))
        out += menu(text, True)
        Return out
    End Function

    ''' <summary>Replaces instances of the current directory in a path string ".."</summary>
    ''' <param name="dirStr">A String containing a windows path</param>
    Public Function replDir(dirStr As String) As String
        Return dirStr.Replace(Environment.CurrentDirectory, "..")
    End Function

    ''' <summary>Appends a newline (or two) to a given String</summary>
    ''' <param name="line">The string to be appended</param>
    ''' <param name="cond">Optional condition under which to append two newlines</param>
    Public Function aNL(line As String, Optional cond As Boolean = False) As String
        Return line & Environment.NewLine
    End Function

    ''' <summary>Prepends a newline (or two) to a given String</summary>
    ''' <param name="cond">The parameter under which to return two newlines</param>
    Public Function pNL(line As String, Optional cond As Boolean = False) As String
        Return If(cond, Environment.NewLine & Environment.NewLine, Environment.NewLine) & line
    End Function

    ''' <summary>Prints a line with a string if we're not suppressing output.</summary>
    ''' <param name="msg">The string to be printed</param>
    Public Sub cwl(Optional msg As String = Nothing)
        If Not SuppressOutput Then Console.WriteLine(msg)
    End Sub

    ''' <summary>Clears the console conditionally when not running unit tests</summary>
    ''' <param name="cond">Optional Boolean specifying whether or not the console should be cleared</param>
    Public Sub clrConsole(Optional cond As Boolean = True)
        ' Do not clear the console during unit tests because there isnt one and the invalid handler throws an IO Exception
        If cond And Not Console.Title.Contains("testhost.x86") Then Console.Clear()
    End Sub
End Module