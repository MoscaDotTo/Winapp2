'    Copyright (C) 2018 Robbie Ward
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

    'basic menu frames & strings
    Public menuStr00 As String = " ╔════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╗"
    Public menuStr01 As String = " ║                                                                                                                    ║"
    Public menuStr02 As String = " ╚════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╝"
    Public menuStr03 As String = " ╠════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╣"
    Public menuStr04 As String = appendNewLine(menu(menuStr01)) & appendNewLine(mkMenuLine("Menu: Enter a number to select", "c")) & mkMenuLine(menuStr01, "")
    Public anyKeyStr As String = "Press any key to return to the winapp2ool menu."
    Public invInpStr As String = "Invalid input. Please try again."
    Public promptStr As String = "Enter a number, or leave blank to run the default: "

    'the maximum length of the portion of the first half of a '#. Option - Description' style menu line
    Dim menuItemLength As Integer

    'indicates whether or not we are pending an exit from the menu
    Public exitCode As Boolean

    'Holds the output text at the top of the menu
    Public menuTopper As String

    Dim optNum As Integer = 0

    ''' <summary>
    ''' Initializes the menu 
    ''' </summary>
    ''' <param name="topper"></param>
    ''' <param name="itemlen"></param>
    Public Sub initMenu(topper As String, itemlen As Integer)
        exitCode = False
        menuTopper = topper
        menuItemLength = itemlen
    End Sub

    ''' <summary>
    ''' Prints a blank menu line
    ''' </summary>
    Public Sub printBlankMenuLine()
        printMenuLine(menuStr01)
    End Sub

    ''' <summary>
    ''' Sets the menuTopper to an error string conditionally, returns the given condition
    ''' </summary>
    ''' <param name="cond">The boolean representing the circumstance under which the error text should be printed</param>
    ''' <param name="errText">The text to be printed to the menu topper</param>
    ''' <returns></returns>
    Public Function denyActionWithTopper(cond As Boolean, errText As String) As Boolean
        If cond Then menuTopper = errText
        Return cond
    End Function

    ''' <summary>
    ''' Prints a menu, option, or reset line conditionally
    ''' </summary>
    ''' <param name="cond">The condition under which to print</param>
    ''' <param name="printType">The type of menu information to print</param>
    ''' <param name="str1">The first string or half string to be printed</param>
    ''' <param name="str2">The alginment or second half string to be printed</param>
    Public Sub printIf(cond As Boolean, printType As String, str1 As String, str2 As String)
        Select Case True
            Case cond And printType = "line"
                printMenuLine(str1, str2)
            Case cond And printType = "opt"
                printMenuOpt(str1, str2)
            Case cond And printType = "reset"
                printResetStr(str1)
        End Select
    End Sub

    ''' <summary>
    ''' Returns the inverse state of a given setting for 
    ''' </summary>
    ''' <param name="setting">A given setting to observe the state of</param>
    ''' <returns></returns
    Public Function enStr(setting As Boolean) As String
        Return If(setting, "Disable", "Enable")
    End Function

    ''' <summary>
    ''' Print the 'Reset Settings' menu option for any given module
    ''' </summary>
    ''' <param name="moduleName"></param>
    Public Sub printResetStr(moduleName As String)
        printBlankMenuLine()
        printMenuOpt("Reset Settings", "Restore " & moduleName & "'s settings to their deault state")
    End Sub

    'Print the top of the menu containing the topper, any description text, the menu prompt, and the exit option
    ''' <summary>
    ''' Prints the top of the menu (containing the topper), any description text provided, the menu prompt, and the exit option
    ''' </summary>
    ''' <param name="descriptionItems">Items describing the menu</param>
    ''' <param name="printExit">The boolean representing whether an option to exit should be printed</param>
    Public Sub printMenuTop(descriptionItems As String(), printExit As Boolean)
        printMenuLine(tmenu(menuTopper))
        printMenuLine(menuStr03)
        For Each line In descriptionItems
            printMenuLine(line, "c")
        Next
        printMenuLine(menuStr04)
        optNum = 0
        If printExit Then printMenuOpt("Exit", "Return to the menu")
    End Sub

    ''' <summary>
    ''' Constructs and returns to the calling function a new menu String
    ''' </summary>
    ''' <param name="lineString">A string to be made into a menu line</param>
    ''' <returns></returns>
    Public Function menu(lineString As String) As String
        Return mkMenuLine(lineString, "")
    End Function

    ''' <summary>
    ''' Prints a menu string
    ''' </summary>
    ''' <param name="lineString"></param>
    Public Sub printMenuLine(lineString As String)
        cwl(menu(lineString))
    End Sub

    ''' <summary>
    ''' Constructs a menu string with a given alignment
    ''' </summary>
    ''' <param name="lineString">The line to be printed</param>
    ''' <param name="align">The anchoring ('l' or 'c')</param>
    ''' <returns></returns>
    Public Function menu(lineString As String, align As String) As String
        Return mkMenuLine(lineString, align)
    End Function

    ''' <summary>
    ''' Prints a menu string with a given alignment
    ''' </summary>
    ''' <param name="lineString"></param>
    ''' <param name="align"></param>
    Public Sub printMenuLine(lineString As String, align As String)
        cwl(menu(lineString, align))
    End Sub

    ''' <summary>
    ''' Prints a numbered menu option after padding it to a set length
    ''' </summary>
    ''' <param name="lineString1">The first part of the menu option</param>
    ''' <param name="lineString2">The second part of the menu option</param>
    Public Sub printMenuOpt(lineString1 As String, lineString2 As String)
        lineString1 = optNum & ". " & lineString1
        While lineString1.Length < menuItemLength
            lineString1 += " "
        End While
        cwl(menu(lineString1 & "- " & lineString2, "l"))
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
    ''' <param name="text"></param>
    ''' <param name="align"></param>
    ''' <returns></returns>
    Public Function bmenu(text As String, align As String) As String
        Dim out As String = menu(menuStr00) & Environment.NewLine
        out += menu(text, align) & Environment.NewLine
        out += menu(menuStr02)
        Return out
    End Function

    ''' <summary>
    ''' Prints the topmost part of the menu with no bottom
    ''' </summary>
    ''' <param name="text"></param>
    ''' <returns></returns>
    Public Function tmenu(text As String) As String
        Dim out As String = menu(menuStr00) & Environment.NewLine
        out += menu(text, "c")
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
End Module