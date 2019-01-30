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
    Public Const menuStr00 As String = " ╔════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╗"
    Public Const menuStr01 As String = " ║                                                                                                                    ║"
    Public Const menuStr02 As String = " ╚════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╝"
    Public Const menuStr03 As String = " ╠════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╣"
    Public menuStr04 As String = appendNewLine(menu(menuStr01)) & appendNewLine(mkMenuLine("Menu: Enter a number to select", "c")) & mkMenuLine(menuStr01, "")
    Public Const anyKeyStr As String = "Press any key to return to the menu."
    Public Const invInpStr As String = "Invalid input. Please try again."
    Public Const promptStr As String = "Enter a number, or leave blank to run the default: "
    ' This boolean will prevent us from printing output or asking for input under most circumstances, triggered by the -s command line argument 
    Public suppressOutput As Boolean = False
    ' The maximum length of the portion of the first half of a '#. Option - Description' style menu line
    Dim menuItemLength As Integer
    ' Indicates whether or not we are pending an exit from the menu
    Public exitCode As Boolean
    ' Holds the text that appears in the top block of the menu
    Public menuHeaderText As String
    ' Holds the current option number at any given moment
    Dim optNum As Integer = 0

    ''' <summary>
    ''' Initializes the menu 
    ''' </summary>
    ''' <param name="topper">The text to be displayed at the top of the menu screen</param>
    ''' <param name="itemlen">The length in characters that should comprise the first bloc of options in the menu</param>
    Public Sub initMenu(topper As String, Optional itemlen As Integer = 35)
        exitCode = False
        menuHeaderText = topper
        menuItemLength = itemlen
    End Sub

    ''' <summary>
    ''' Sets the menu header to an error string conditionally, returns the given condition
    ''' </summary>
    ''' <param name="cond">The condition under which the error text should be printed</param>
    ''' <param name="errText">The error text to be printed in the menu header</param>
    ''' <returns></returns>
    Public Function denyActionWithTopper(cond As Boolean, errText As String) As Boolean
        If cond Then menuHeaderText = errText
        Return cond
    End Function

    ''' <summary>
    ''' Prints a menu line, option, or reset string, conditionally
    ''' </summary>
    ''' <param name="cond">The condition under which to print</param>
    ''' <param name="printType">The type of menu information to print</param>
    ''' <param name="str1">The first string or half string to be printed</param>
    ''' <param name="optString">The alginment or second half string to be printed</param>
    Public Sub print(printType As Integer, str1 As String, Optional optString As String = "", Optional cond As Boolean = True, Optional leadingBlank As Boolean = False, Optional trailingBlank As Boolean = False, Optional isCentered As Boolean = False, Optional closeMenu As Boolean = False)
        If cond And leadingBlank Then printMenuLine()
        Select Case True
            Case cond And printType = 0
                printMenuLine(str1, isCentered)
            Case cond And printType = 1
                printMenuOpt(str1, optString)
            Case cond And printType = 2
                print(1, "Reset Settings", $"Restore {str1}'s settings to their default state", leadingBlank:=True)
            Case cond And printType = 3
                print(0, tmenu(str1), isCentered:=True, closeMenu:=True)
            Case cond And printType = 4
                print(0, $"{str1}, some functions will not be available.", trailingBlank:=True, isCentered:=True)
        End Select
        If cond And trailingBlank Then printMenuLine()
        If cond And closeMenu Then printMenuLine(menuStr02)
    End Sub

    ''' <summary>
    ''' Returns the inverse state of a given setting as a String
    ''' </summary>
    ''' <param name="setting">The setting whose state will be reported</param>
    ''' <returns></returns>
    Public Function enStr(setting As Boolean) As String
        Return If(setting, "Disable", "Enable")
    End Function

    ''' <summary>
    ''' Exits a menu or module by flipping the exitCode to true 
    ''' </summary>
    Public Sub exitModule()
        exitCode = True
    End Sub

    ''' <summary>
    ''' Prints the top of the menu (containing the topper), any description text provided, the menu prompt, and the exit option
    ''' </summary>
    ''' <param name="descriptionItems">Items describing the menu</param>
    ''' <param name="printExit">The boolean representing whether an option to exit should be printed</param>
    Public Sub printMenuTop(descriptionItems As String(), Optional printExit As Boolean = True)
        printMenuLine(tmenu(menuHeaderText))
        printMenuLine(menuStr03)
        descriptionItems.ToList.ForEach(Sub(line) printMenuLine(line, True))
        printMenuLine(menuStr04)
        optNum = 0
        print(1, "Exit", "Return to the menu", printExit)
    End Sub

    ''' <summary>
    ''' Constructs and returns to the calling function a new menu String
    ''' </summary>
    ''' <param name="lineString">A string to be made into a menu line</param>
    ''' <returns></returns>
    Public Function menu(lineString As String) As String
        Return mkMenuLine(lineString, "l")
    End Function

    ''' <summary>
    ''' Prints a line in a menu
    ''' </summary>
    ''' <param name="lineString"></param>
    Public Sub printMenuLine(Optional lineString As String = menuStr01)
        cwl(menu(lineString))
    End Sub

    ''' <summary>
    ''' Constructs a menu string with a given alignment
    ''' </summary>
    ''' <param name="lineString">The line to be printed</param>
    ''' <param name="isCentered">The boolean indicating whether the line text should be centered</param>
    ''' <returns></returns>
    Public Function menu(lineString As String, Optional isCentered As Boolean = False) As String
        Return mkMenuLine(lineString, If(isCentered, "c", "l"))
    End Function

    ''' <summary>
    ''' Prints a menu string with a given alignment
    ''' </summary>
    ''' <param name="lineString"></param>
    ''' <param name="isCenteredLine">The boolean indicating whether the line text should be centered</param>
    Public Sub printMenuLine(lineString As String, Optional isCenteredLine As Boolean = False)
        cwl(menu(lineString, isCenteredLine))
    End Sub

    ''' <summary>
    ''' Prints a numbered menu option after padding it to a set length
    ''' </summary>
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

    ''' <summary>
    ''' Flips the exitCode boolean so we can return to the menu when desired
    ''' </summary>
    Public Sub revertMenu()
        exitCode = Not exitCode
    End Sub

    ''' <summary>
    ''' Forces the exitCode to be False
    ''' </summary>
    Public Sub undoAnyPendingExits()
        exitCode = False
    End Sub

    ''' <summary>
    ''' Constructs a menu line fit to the width of the console
    ''' </summary>
    ''' <param name="line">The line to be printed</param>
    ''' <param name="align">The alignment of the line to be printed. 'l' for Left or 'c' for Centre</param>
    ''' <returns></returns>
    Public Function mkMenuLine(line As String, align As String) As String
        If line.Length >= 119 Then Return line
        Dim out As String = " ║"
        Select Case align
            Case "c"
                padToEnd(out, CInt((((118 - line.Length) / 2) + 2)))
                out += line
                padToEnd(out, 118)
            Case "l"
                out += " " & line
                padToEnd(out, 118)
        End Select
        Return out
    End Function

    ''' <summary>
    ''' Pads a given string with spaces 
    ''' </summary>
    ''' <param name="out">The string to be padded</param>
    ''' <param name="targetLen">The end length to which the string should be padded</param>
    Private Sub padToEnd(ByRef out As String, targetLen As Integer)
        While out.Length < targetLen
            out += " "
        End While
        If targetLen = 118 Then out += "║"
    End Sub

    ''' <summary>
    ''' Prints a box with a single message inside it 
    ''' </summary>
    ''' <param name="text">The string to be printed in the box</param>
    ''' <returns></returns>
    Public Function bmenu(text As String) As String
        Dim out As String = appendNewLine(menu(menuStr00))
        out += appendNewLine(menu(text, True))
        out += menu(menuStr02)
        Return out
    End Function

    ''' <summary>
    ''' Prints the topmost part of the menu with no bottom
    ''' </summary>
    ''' <param name="text">The String to be printed in the faux menu header</param>
    ''' <returns></returns>
    Public Function tmenu(text As String) As String
        Dim out As String = appendNewLine(menu(menuStr00))
        out += menu(text, True)
        Return out
    End Function

    ''' <summary>
    ''' Replaces instances of the current directory in a path string ".."
    ''' </summary>
    ''' <param name="dirStr">A String containing a windows path</param>
    ''' <returns></returns>
    Public Function replDir(dirStr As String) As String
        Return dirStr.Replace(Environment.CurrentDirectory, "..")
    End Function

    ''' <summary>
    ''' Appends a newline to a given String
    ''' </summary>
    ''' <param name="line">The string to be appended</param>
    ''' <returns></returns>
    Public Function appendNewLine(line As String) As String
        Return line & Environment.NewLine
    End Function

    ''' <summary>
    ''' Prints a line with a string if we're not surpressing output.
    ''' </summary>
    ''' <param name="msg">The string to be printed</param>
    Public Sub cwl(Optional msg As String = Nothing)
        If Not suppressOutput Then Console.WriteLine(msg)
    End Sub

    ''' <summary>
    ''' Clears the console if there's a pending exit, returns exitCode
    ''' </summary>
    ''' <returns></returns>
    Public Function pendingExit() As Boolean
        If exitCode Then Console.Clear()
        Return exitCode
    End Function
End Module