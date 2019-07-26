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
''' MenuMaker is a singleton driver module for powering dyanamic finite state console applications with numbered menus
''' </summary>
Module MenuMaker
    '''<summary>An instruction to press any key to return to the previous menu </summary>
    Public ReadOnly Property anyKeyStr As String = "Press any key to return to the menu."
    '''<summary>An error message informing the user their input was invalid</summary>
    Public ReadOnly Property invInpStr As String = "Invalid input. Please try again."
    '''<summary>An instruction for the user to provide input</summary>
    Public ReadOnly Property promptStr As String = "Enter a number, or leave blank to run the default: "
    '''<summary>The maximum length of the portion of the first half of a '#. Option - Description' style menu line</summary>
    Private Property menuItemLength As Integer
    '''<summary>Indicates that the menu header should be printed with color </summary>
    Public Property ColorHeader As Boolean
    '''<summary>The color with which the next header should be printed if <c>ColorHeader</c> is <c>True</c></summary>
    Public Property HeaderColor As ConsoleColor
    '''<summary>Holds the current option number for the menu instance</summary>
    Private Property OptNum As Integer = 0
    ''' <summary>Indicates that the application should not output or ask input from the user except when encountering exceptions <br /> Default: False </summary>
    Public Property SuppressOutput As Boolean = False
    ''' <summary> Indicates that an exit from the current menu is pending </summary>
    Public Property ExitPending As Boolean
    ''' <summary> Holds the text that appears in the top block of the menu </summary>
    Public Property MenuHeaderText As String

    '''<summary>Frame characters used to open a menu line</summary>
    Private ReadOnly Property Openers As String() = {"║", "╔", "╚", "╠"}
    '''<summary>Frame characters used to close a menu line</summary>
    Private ReadOnly Property Closers As String() = {"║", "╗", "╝", "╣"}

    ''' <summary>Initializes a module's menu, prints it, and handles the user input. Effectively the main event loop for anything built with MenuMaker</summary>
    ''' <param name="name">The name of the module as it will be displayed to the user</param>
    ''' <param name="showMenu">The subroutine prints the module's menu</param>
    ''' <param name="handleInput">The subroutine that handles the module's input</param>
    ''' <param name="itmLen">Indicates the maximum length of menu option names<br/>Optional, Default: 35</param>
    Public Sub initModule(name As String, showMenu As Action, handleInput As Action(Of String), Optional itmLen As Integer = 35)
        gLog("", ascend:=True)
        gLog($"Loading module {name}")
        ExitPending = False
        setHeaderText(name)
        menuItemLength = itmLen
        Do Until ExitPending
            clrConsole()
            showMenu()
            Console.Write(Environment.NewLine & promptStr)
            handleInput(Console.ReadLine)
        Loop
        ExitPending = False
        setHeaderText($"{name} closed")
        gLog($"Exiting {name}", descend:=True, leadr:=True)
    End Sub

    ''' <summary>Prints menu lines, options, and frames fit to the current console window width</summary>
    ''' <param name="printType">The type of menu information to print <br /> 
    ''' <list type="bullet">
    ''' <item><description>0: Line </description></item>
    ''' <item><description>1: Option</description></item>
    ''' <item><description>2: Option with a "Reset Settings" prompt </description></item>
    ''' <item><description>3: Box with centered text</description></item>
    ''' <item><description>4: Menu top</description></item>
    ''' <item><description>5: Option with an Enable/Disable prompt</description></item>
    ''' </list></param>
    ''' <param name="menuText">The text to be printed. <br /> When <paramref name="printType"/> is  1 or 5, <paramref name="menuText"/> contains the name of the menu option
    ''' <br/>When <paramref name="printType"/> is 3, <paramref name="menuText"/> contains the name of the module whose settings are being reset</param>
    ''' <param name="optString">The description of the menu option<br /> Optional, Default: ""</param>
    ''' <param name="cond">Indicates that the line should be printed. <br /> Optional, Default: True</param>
    ''' <param name="leadingBlank">Indicates that a blank menu line should be printed immediately before the printed line <br /> Optional, Default: False</param>
    ''' <param name="trailingBlank">Indicates that a blank menu line should be printed immediately after the printed line <br /> Optional, Default: False</param>
    ''' <param name="isCentered">Indicates that the printed text should be centered <br /> Optional, Default: False</param>
    ''' <param name="closeMenu">Indicates that the bottom menu frame should be printed <br /> Optional, Default: False</param>
    ''' <param name="openMenu">Indicates that the top menu frame should be printed <br />Optional, Default: False</param>
    ''' <param name="enStrCond">A module setting whose menu text will include an Enable/Disable toggle <br /> 
    ''' If <paramref name="colorLine"/> is True and <paramref name="useArbitraryColor"/> is False, the line will be printed 
    ''' Green if <paramref name="enStrCond"/> is True, Red if False<br /> Optional, Default: False (Red) </param>
    ''' <param name="colorLine">Indicates we want to color any lines we print<br /> Optional, Default: False</param>
    ''' <param name="useArbitraryColor">Indicates that the line should be colored using the value provided by <paramref name="arbitraryColor"/><br /> Optional, Default: False</param>
    ''' <param name="arbitraryColor">Foreground ConsoleColor to be used when printing with when <paramref name="colorLine"/> is true, but wanting to use a color other than red/green <br /> Optional, Default: Nothing</param>
    ''' <param name="buffr">Indicates that a leading newline should be printed before the menu lines<br />Optional, Default: False</param>
    ''' <param name="trailr">Indicates that a trailing newline should be printed after the menu lines<br />Optional, Default: False</param>
    ''' <param name="conjoin">Indicates that a conjoining menu frame should be printed <br /> Optional, Default: False</param>
    Public Sub print(printType As Integer, menuText As String, Optional optString As String = "", Optional cond As Boolean = True,
                     Optional leadingBlank As Boolean = False, Optional trailingBlank As Boolean = False, Optional isCentered As Boolean = False,
                     Optional closeMenu As Boolean = False, Optional openMenu As Boolean = False, Optional enStrCond As Boolean = False,
                     Optional colorLine As Boolean = False, Optional useArbitraryColor As Boolean = False, Optional arbitraryColor As ConsoleColor = Nothing,
                     Optional buffr As Boolean = False, Optional trailr As Boolean = False, Optional conjoin As Boolean = False)
        If Not cond Then Return
        cwl(cond:=buffr)
        If colorLine Then Console.ForegroundColor = If(useArbitraryColor, arbitraryColor, If(enStrCond, ConsoleColor.Green, ConsoleColor.Red))
        print(0, Nothing, cond:=leadingBlank)
        print(0, getFrame(1), cond:=openMenu)
        Select Case printType
            ' Prints lines
            Case 0
                printMenuLine(menuText, isCentered)
            ' Prints options
            Case 1
                printMenuOpt(menuText, optString)
            ' Prints the Reset Settings option
            Case 2
                print(1, "Reset Settings", $"Restore {menuText}'s settings to their default state", leadingBlank:=True)
            ' Prints a box with centered text
            Case 3
                print(4, menuText, closeMenu:=True)
            ' The top of a menu with a header
            Case 4
                print(0, menuText, isCentered:=True, openMenu:=True)
            ' Colored line printing for enable/disable menu options
            Case 5
                print(1, menuText, $"{enStr(enStrCond)} {optString}", colorLine:=True, enStrCond:=enStrCond)
        End Select
        print(0, getFrame(3), cond:=conjoin)
        print(0, Nothing, cond:=trailingBlank)
        print(0, getFrame(2), cond:=closeMenu)
        If colorLine Then Console.ResetColor()
        cwl(cond:=trailr)
    End Sub

    ''' <summary>Prints a line to the console window if <c>SuppressOutput</c> and <paramref name="cond"/> are true </summary>
    ''' <param name="msg">The string to be printed <br />Default: Nothing</param>
    ''' <param name="cond">Indicates the line should be printed <br />Optional, Default: True</param>
    Public Sub cwl(Optional msg As String = Nothing, Optional cond As Boolean = True)
        If cond And Not SuppressOutput Then Console.WriteLine(msg)
    End Sub

    ''' <summary>Waits for the user to press a key if <c>SuppressOutput</c> is <c>False</c></summary>
    Public Sub crk()
        If Not SuppressOutput Then Console.ReadKey()
    End Sub

    ''' <summary>Clears the console if <paramref name="cond"/> is <c>True</c> and we're not unit testing</summary>
    ''' <remarks>When unit testing, the console window doesn't belong to us and trying to clear the console throws an IO Exception, so we don't do that</remarks>
    ''' <param name="cond">Indicates that the console should be cleared <br /> Optional, Default: True</param>
    Public Sub clrConsole(Optional cond As Boolean = True)
        If cond And Not SuppressOutput And Not Console.Title.Contains("testhost.x86") Then Console.Clear()
    End Sub

    '''<summary>Returns a menuframe</summary>
    '''<param name="frameNum">Indicates which frame should be returned. <br />
    '''<list type="bullet">
    '''<item><description>0: Empty menu line with vertical frames</description></item>
    '''<item><description>1: Filled menu line with downward opening 90° angle frames</description></item>
    '''<item><description>2: Filled menu line with upward opening 90° angle frames</description></item>
    '''<item><description>3: Filled menu line with inward facing T-frames </description></item>
    ''' </list>
    '''0: empty line, 1: top, 2: bottom, 3: conjoiner <br /> Default: 0</param>
    '''<returns>A menu frame based on the value of <paramref name="frameNum"/></returns>
    Private Function getFrame(Optional frameNum As Integer = 0) As String
        Return mkMenuLine("", "f", frameNum)
    End Function

    ''' <summary>Saves a menu header to be printed atop the next menu, optionally with color</summary>
    ''' <param name="txt">The text to appear in the header</param>
    ''' <param name="cHeader">Indicates that the header should be colored using the color given by <paramref name="printColor"/><br />Optional, Default: False</param>
    ''' <param name="cond">Indicates that the header text should be assigned the value given by <paramref name="txt"/><br /> Optional, Default: True</param>
    ''' <param name="printColor">ConsoleColor with which the header should be colored when <paramref name="cHeader"/> is True <br/> Optional, Default: Red</param>
    Public Sub setHeaderText(txt As String, Optional cHeader As Boolean = False, Optional cond As Boolean = True, Optional printColor As ConsoleColor = ConsoleColor.Red)
        If Not cond Then Return
        MenuHeaderText = txt
        ColorHeader = cHeader
        HeaderColor = printColor
    End Sub

    ''' <summary>Informs a user when an action is unable to proceed due to a condition</summary>
    ''' <param name="cond">Indicates that an action should be denied</param>
    ''' <param name="errText">The error text to be printed in the menu header</param>
    Public Function denyActionWithHeader(cond As Boolean, errText As String) As Boolean
        setHeaderText(errText, True, cond)
        Return cond
    End Function

    ''' <summary>Returns the inverse state of a given boolean as a String</summary>
    ''' <param name="setting">A module setting whose state will be observed</param>
    ''' <returns><c>"Disable"</c> if <paramref name="setting"/>is True, <c>"Enable"</c>otherwise</returns>
    Public Function enStr(setting As Boolean) As String
        Return If(setting, "Disable", "Enable")
    End Function

    ''' <summary>Enforces that initMenu exit the current level in the stack on the next iteration of its loop</summary>
    Public Sub exitModule()
        ExitPending = True
    End Sub

    ''' <summary>Prints the top of the menu, the header, a conjoiner, any description text provided, the menu prompt, and the exit option</summary>
    ''' <param name="descriptionItems">Text describing the current menu or module functions being presented to the user, each array will be displayed on a separate line</param>
    ''' <param name="printExit">Indicates that an option to exit to the previous menu should be printed <br /> Optional, Default: True</param>
    Public Sub printMenuTop(descriptionItems As String(), Optional printExit As Boolean = True)
        print(4, MenuHeaderText, colorLine:=ColorHeader, useArbitraryColor:=True, arbitraryColor:=HeaderColor, conjoin:=True)
        For Each line In descriptionItems
            print(0, line, isCentered:=True)
        Next
        print(0, "Menu: Enter a number to select", leadingBlank:=True, trailingBlank:=True, isCentered:=True)
        OptNum = 0
        print(1, "Exit", "Return to the menu", printExit)
    End Sub

    ''' <summary>Prints a line bounded by vertical menu frames, an empty menu line if <paramref name="lineString"/> is <c>Nothing</c></summary>
    ''' <param name="lineString">The text to be printed. <br/> Optional, Default: Nothing</param>
    ''' <param name="isCentered">Indicates that the printed text should be centered <br /> Optional, Default: False</param>
    Private Sub printMenuLine(Optional lineString As String = Nothing, Optional isCentered As Boolean = False)
        If lineString = Nothing Then lineString = getFrame()
        cwl(mkMenuLine(lineString, If(isCentered, "c", "l")))
    End Sub

    ''' <summary>Prints a numbered menu option after padding it to a set length</summary>
    ''' <param name="lineString1">The name of the menu option</param>
    ''' <param name="lineString2">The description of the menu option</param>
    Private Sub printMenuOpt(lineString1 As String, lineString2 As String)
        lineString1 = $"{OptNum}. {lineString1}"
        padToEnd(lineString1, menuItemLength, "")
        cwl(mkMenuLine($"{lineString1}- {lineString2}", "l"))
        OptNum += 1
    End Sub

    ''' <summary>Constructs a menu line fit to the width of the console</summary>
    ''' <param name="line">The text to be printed</param>
    ''' <param name="align">The alignment of the line to be printed: <br /> 
    ''' <list type="bullet">
    ''' <item><description>'c': centers the string </description></item>
    ''' <item><description>'l': leftaligns the string</description></item>
    ''' <item><description>'f': prints a menu frame</description></item>
    ''' </list></param>
    ''' <param name="borderInd">Determines which characters should create the border for the menuline: <br />
    ''' <list type="bullet">
    ''' <item><description>0: Vertical lines</description></item>
    ''' <item><description>1: Ceiling brackets</description></item>
    ''' <item><description>2: Floor brackets</description></item>
    ''' <item><description>3: Conjoining brackets </description></item></list> <br />Optional, Default: 0</param>
    Private Function mkMenuLine(line As String, align As String, Optional borderInd As Integer = 0) As String
        If line.Length >= Console.WindowWidth - 1 Then Return line
        Dim out = $" {Openers(borderInd)}"
        Select Case align
            Case "c"
                padToEnd(out, CInt((((Console.WindowWidth - line.Length) / 2) + 2)), Closers(borderInd))
                out += line
                padToEnd(out, Console.WindowWidth - 2, Closers(borderInd))
            Case "l"
                out += " " & line
                padToEnd(out, Console.WindowWidth - 2, Closers(borderInd))
            Case "f"
                padToEnd(out, Console.WindowWidth - 2, Closers(borderInd), If(borderInd = 0, " ", "═"))
        End Select
        Return out
    End Function

    ''' <summary>Pads a given string with spaces until it is a target length</summary>
    ''' <param name="out">The text to be padded</param>
    ''' <param name="targetLen">The length to which the text should be padded</param>
    ''' <param name="endline">The closer character for the type of frame being built</param>
    ''' <param name="padChar">The character with which to pad the text <br /> Default: " " (space character)</param>
    Private Sub padToEnd(ByRef out As String, targetLen As Integer, endline As String, Optional padChar As String = " ")
        While out.Length < targetLen
            out += padChar
        End While
        If targetLen = Console.WindowWidth - 2 Then out += endline
    End Sub

    ''' <summary>Replaces instances of the current directory in a path string ".."</summary>
    ''' <param name="dirStr">A windows filesystem path</param>
    ''' <returns><paramref name="dirStr"/> with instances of the current directory replaced with <c>".."</c></returns>
    Public Function replDir(dirStr As String) As String
        Return dirStr.Replace(Environment.CurrentDirectory, "..")
    End Function
End Module