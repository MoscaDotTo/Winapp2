'    Copyright (C) 2018-2025 Hazel Ward
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
''' Maintains the global log for winapp2ool, which is used to track internal operations and errors.
''' Provides methods for adding to the log, saving it to disk, and printing it to the console.
''' </summary>
''' 
''' Docs last updated: 2025-06-19 | Code last updated: 2025-06-19
Public Module logger

    ''' <summary> 
    ''' The global Winapp2ool log, containing everything logged during the current session 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-06-19 | Code last updated: 2025-06-19
    Public Property GlobalLog As New strList

    '''<summary> 
    '''Holds the filesystem location to which the log file will optionally be saved.
    '''</summary>
    '''
    ''' Docs last updated: 2025-06-19 | Code last updated: 2025-06-19
    Public Property GlobalLogFile As New iniFile(Environment.CurrentDirectory, "winapp2ool.log")

    '''<summary> 
    '''The current indentation level of the global log.
    '''</summary>
    '''
    ''' Docs last updated: 2025-06-19 | Code last updated: 2025-06-19
    Public Property nestCount As Integer = 0

    ''' <summary> 
    ''' Adds an item into the global log 
    ''' </summary>
    ''' 
    ''' <param name="logstr"> 
    ''' The <c> String </c> to be added into the log <br /> 
    ''' Optional, Default: <c> ""</c> 
    ''' </param>
    ''' 
    ''' <param name="cond"> 
    ''' Indicates that the <c> <paramref name="logstr"/> </c> should be added into the log <br /> 
    ''' Optional, Default: <c> True </c> 
    ''' </param>
    ''' 
    ''' <param name="ascend">
    ''' Indicates that the line should be indented. (Generally) requires a corresponding <c> <paramref name="descend"/></c> 
    ''' to "undo." Useful for blocking groups of related log items without needing to call <c> <paramref name="indent"/> </c> on each 
    ''' <br /> Optional, Default: <c> False </c>
    ''' </param>
    ''' 
    ''' <param name="descend">
    ''' Indicates that the line should be unindented. (Generally) follows a corresponding <c> <paramref name="ascend"/> </c>. 
    ''' Useful for blocking groups of related log items without needing to call <c> <paramref name="indent"/> </c> on each 
    ''' <br /> Optional, Default: <c> False </c> 
    ''' </param>
    '''  
    ''' <param name="indent">
    ''' Indicates that the line should be indented individually, without affecting the indentation of following lines 
    ''' <br/> Optional, Default: <c> False </c> 
    ''' </param>
    ''' 
    ''' <param name="indAmt">
    ''' The number of times by which the line should be indented given <c> <paramref name="indent"/> </c> is <c> True </c> 
    ''' <br /> Optional, Default: <c> 1 </c> 
    ''' </param>
    ''' 
    ''' <param name="descAmt">
    ''' The number of times *fewer* by which following lines should be indented 
    ''' <br /> Optional, Default: <c> 1 </c> 
    ''' </param>
    ''' 
    ''' <param name="ascAmt"> The number of times by which this and following lines should be indented 
    ''' <br /> Optional, Default: <c> 1 </c> 
    ''' </param>
    ''' 
    ''' <param name="buffr"> 
    ''' Indicates that an empty line should be added into the log following <c> <paramref name="logstr"/> </c> 
    ''' </param>
    ''' 
    ''' <param name="leadr"> 
    ''' Indicates that an empty line should be added into the log before <c> <paramref name="logstr"/> </c> 
    ''' </param>
    ''' 
    ''' Docs last updated: 2025-06-24 | Code last updated: 2025-06-19
    Public Sub gLog(Optional logstr As String = "",
                    Optional cond As Boolean = True,
                    Optional ascend As Boolean = False,
                    Optional descend As Boolean = False,
                    Optional indent As Boolean = False,
                    Optional indAmt As Integer = 1,
                    Optional descAmt As Integer = 1,
                    Optional ascAmt As Integer = 1,
                    Optional buffr As Boolean = False,
                    Optional leadr As Boolean = False)

        If Not cond Then Return


        If leadr Then GlobalLog.add("")
        If indent Then nestCount += indAmt
        If ascend Then nestCount += ascAmt

        Dim buffer = New String(" "c, nestCount * 2)

        logstr = buffer + logstr
        GlobalLog.add(logstr)

        If indent Then nestCount -= indAmt
        If descend Then nestCount -= descAmt
        If buffr Then GlobalLog.add("")

    End Sub

    ''' <summary>
    ''' Adds a given message to the global log and also prints it to the console, this is just a wrapper function to avoid having to call both 
    ''' glog and print for the same message sometimes 
    ''' </summary>
    ''' 
    ''' <param name="menuText"> 
    ''' The <c> String </c> to be added into the log and also printed to the user <br /> 
    ''' Optional, Default: <c> ""</c> 
    ''' </param>
    ''' 
    ''' <param name="cond"> 
    ''' Indicates that the <c> <paramref name="menuText"/> </c> should be printed to the user <br /> 
    ''' Optional, Default: <c> True </c> 
    ''' </param>
    ''' 
    ''' <param name="ascend">
    ''' Indicates that the line should be indented. (Generally) requires a corresponding <c> <paramref name="descend"/></c> 
    ''' to "undo." Useful for blocking groups of related log items without needing to call <c> <paramref name="indent"/> </c> on each 
    ''' <br /> Optional, Default: <c> False </c>
    ''' </param>
    ''' 
    ''' <param name="descend">
    ''' Indicates that the line should be unindented. (Generally) follows a corresponding <c> <paramref name="ascend"/> </c>. 
    ''' Useful for blocking groups of related log items without needing to call <c> <paramref name="indent"/> </c> on each 
    ''' <br /> Optional, Default: <c> False </c> 
    ''' </param>
    '''  
    ''' <param name="indent">
    ''' Indicates that the line should be indented individually, without affecting the indentation of following lines 
    ''' <br/> Optional, Default: <c> False </c> 
    ''' </param>
    ''' 
    ''' <param name="indAmt">
    ''' The number of times by which the line should be indented given <c> <paramref name="indent"/> </c> is <c> True </c> 
    ''' <br /> Optional, Default: <c> 1 </c> 
    ''' </param>
    ''' 
    ''' <param name="descAmt">
    ''' The number of times *fewer* by which following lines should be indented 
    ''' <br /> Optional, Default: <c> 1 </c> 
    ''' </param>
    ''' 
    ''' <param name="ascAmt"> The number of times by which this and following lines should be indented 
    ''' <br /> Optional, Default: <c> 1 </c> 
    ''' </param>
    ''' 
    ''' <param name="buffr"> 
    ''' Indicates that an empty line should be added into the log following <c> <paramref name="menuText"/> </c> 
    ''' </param>
    ''' 
    ''' <param name="leadr"> 
    ''' Indicates that an empty line should be added into the log before <c> <paramref name="menuText"/> </c> 
    ''' </param>
    ''' 
    ''' <param name="printType"> 
    ''' The type of menu information to print <br/> 
    ''' <list type="bullet">
    ''' <item> <description> <c> 0 </c>: Line </description> </item>
    ''' <item> <description> <c> 1 </c>: Option </description> </item>
    ''' <item> <description> <c> 2 </c>: Option with a "Reset Settings" prompt </description> </item>
    ''' <item> <description> <c> 3 </c>: Box with centered text </description> </item>
    ''' <item> <description> <c> 4 </c>: Menu top </description> </item>
    ''' <item> <description> <c> 5 </c>: Option with an Enable/Disable prompt </description> </item>
    ''' </list>
    ''' </param>
    ''' 
    ''' <param name="optString"> 
    ''' The description of the menu option
    ''' <br/> Optional, Default: <c> "" </c> 
    ''' </param>
    ''' 
    ''' <param name="leadingBlank"> 
    ''' Indicates that a blank menu line should be printed immediately before the printed line
    ''' <br/> Optional, Default: <c> False </c> 
    ''' </param>
    ''' 
    ''' <param name="trailingBlank"> 
    ''' Indicates that a blank menu line should be printed immediately after the printed line
    ''' <br/> Optional, Default: <c> False </c> 
    ''' </param>
    ''' 
    ''' <param name="isCentered"> 
    ''' Indicates that the printed text should be centered
    ''' <br/> Optional, Default: <c> False </c> 
    ''' </param>
    ''' 
    ''' <param name="closeMenu"> 
    ''' Indicates that the bottom menu frame should be printed 
    ''' <br/> Optional, Default: <c> False </c> 
    ''' </param>
    ''' 
    ''' <param name="openMenu"> 
    ''' Indicates that the top menu frame should be printed 
    ''' <br/> Optional, Default: <c> False </c> 
    ''' </param>
    ''' 
    ''' <param name="enStrCond"> 
    ''' A module setting whose menu text will include an Enable/Disable toggle <br/> <br/>
    ''' If lines are being colored without an <c> <paramref name="arbitraryColor"/> </c>, 
    ''' they will be printed <c> Green </c> if <c> <paramref name="enStrCond"/> </c> is <c> True </c>,
    ''' otherwise they will be printed <c> Red </c>
    ''' <br/> Optional, Default: <c> False (Red) </c> 
    ''' </param>
    ''' 
    ''' <param name="colorLine">
    ''' Indicates that lines should be printed using color 
    ''' <br/> Optional, Default: <c> False </c> 
    ''' </param>
    ''' 
    ''' <param name="useArbitraryColor"> 
    ''' Indicates that the line should be colored using the value provided by <c> <paramref name="arbitraryColor"/> </c> 
    ''' <br/> Optional, Default: <c> False </c> 
    ''' </param>
    ''' 
    ''' <param name="arbitraryColor"> 
    ''' Foreground <c> ConsoleColor </c> to be used when printing with when <c> <paramref name="colorLine"/> </c> is <c> True </c>, 
    ''' but wanting to use a color other than <c> Red </c> or <c> Green </c> 
    ''' <br/> Optional, Default: <c> Nothing </c>
    ''' </param>
    '''
    ''' <param name="trailr"> 
    ''' Indicates that a trailing newline should be printed after the menu lines 
    ''' <br/> Optional, Default: <c> False </c> 
    ''' </param>
    ''' 
    ''' <param name="conjoin"> 
    ''' Indicates that a conjoining menu frame should be printed after the printed lines 
    ''' <br/> Optional, Default: <c> False </c>
    ''' </param>
    ''' 
    ''' <param name="fillBorder"> 
    ''' Indicates whether or not any menu frames should be filled or be empty 
    ''' <br /> Optional, Default: <c> True (filled) </c>
    ''' </param>
    ''' 
    ''' <param name="logBuffr"> 
    ''' Indicates whether to log buffered messages
    ''' <br/> Optional, Default: <c> False </c>
    ''' </param>
    ''' 
    ''' <param name="logCond">
    ''' Indicates that the <c> <paramref name="menuText"/> </c> should be added into the global log 
    ''' <br /> Optional, Default: <c> True </c>
    ''' </param>
    ''' 
    ''' Docs last updated: 2024-05-18 | Code last updated: 2024-05-18
    Public Sub LogAndPrint(printType As Integer,
                           menuText As String,
                           Optional logCond As Boolean = True,
                           Optional ascend As Boolean = False,
                           Optional descend As Boolean = False,
                           Optional indent As Boolean = False,
                           Optional indAmt As Integer = 1,
                           Optional descAmt As Integer = 1,
                           Optional ascAmt As Integer = 1,
                           Optional logBuffr As Boolean = False,
                           Optional leadr As Boolean = False,
                           Optional optString As String = "",
                           Optional cond As Boolean = True,
                           Optional leadingBlank As Boolean = False,
                           Optional trailingBlank As Boolean = False,
                           Optional isCentered As Boolean = False,
                           Optional closeMenu As Boolean = False,
                           Optional openMenu As Boolean = False,
                           Optional enStrCond As Boolean = False,
                           Optional colorLine As Boolean = False,
                           Optional useArbitraryColor As Boolean = False,
                           Optional arbitraryColor As ConsoleColor = Nothing,
                           Optional buffr As Boolean = False,
                           Optional trailr As Boolean = False,
                           Optional conjoin As Boolean = False,
                           Optional fillBorder As Boolean = True)

        gLog(menuText, logCond, ascend, descend, indent, indAmt, descAmt, ascAmt, logBuffr, leadr)
        print(printType, menuText, optString, cond, leadingBlank, trailingBlank, isCentered, closeMenu, openMenu, enStrCond, colorLine, useArbitraryColor, arbitraryColor, buffr, trailr, conjoin, fillBorder)

    End Sub

    ''' <summary> 
    ''' Saves the global log to disk if the given <c> <paramref name="cond"/> </c> is met 
    ''' </summary>
    ''' 
    ''' <param name="cond">
    ''' Indicates that the global log should be saved to disk 
    ''' <br /> Optional, Default: <c> False </c>
    ''' </param>
    ''' 
    ''' Docs last updated: 2025-06-19 | Code last updated: 2024-05-18
    Public Sub saveGlobalLog(Optional cond As Boolean = True)

        GlobalLogFile.overwriteToFile(logger.toString, cond)

    End Sub

    ''' <summary>
    ''' Returns the log as a single <c> String </c>
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-06-19  | Code last updated: 2024-05-18
    Public Function toString() As String

        Dim sb As New StringBuilder()

        GlobalLog.Items.ForEach(Sub(line) sb.AppendLine(line))

        Return sb.ToString()

    End Function

    ''' <summary> 
    ''' Prints the winapp2ool log to the user and waits for the 'enter' key to be pressed before returning to the calling menu 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-06-19 | Code last updated: 2024-05-18
    Public Sub printLog()

        cwl("Printing the winapp2ool log, this may take a moment")

        Dim out = logger.toString

        clrConsole()

        cwl(out)
        cwl()
        cwl($"End of log. {pressEnterStr}")

        Console.ReadLine()

    End Sub


    '''<summary> 
    ''' Gets the most recent segment of the global log contained by two phrases (ie. a module name or subroutine) 
    ''' As a <c> String </c> <br /> Can be used to simply fetch the logs from modules for saving to disk or displaying to the user after a run.
    ''' </summary>
    ''' 
    ''' <param name="startingPhrase"> 
    ''' The starting phrase of the requested log slice 
    ''' </param>
    ''' 
    ''' <param name="endingPhrase"> 
    ''' The ending phrase of the requested log slice 
    ''' </param>
    ''' 
    ''' <returns> 
    ''' The set of log lines between the most recent incidences of the provided phrases from the log 
    ''' </returns>
    ''' 
    ''' Docs last updated: 2025-06-19 | Code last updated: 2025-06-19
    Public Function getLogSliceFromGlobal(startingPhrase As String, endingPhrase As String) As String

        Dim startInd = -1
        Dim endInd = -1

        For i = GlobalLog.Items.Count - 1 To 0 Step -1

            If Not GlobalLog.Items(i).Contains(startingPhrase) Then Continue For

            startInd = i
            Exit For

        Next

        If startInd = -1 Then Return ""

        For i = startInd To GlobalLog.Items.Count - 1

            If Not GlobalLog.Items(i).EndsWith(endingPhrase, StringComparison.InvariantCultureIgnoreCase) Then Continue For

            endInd = i
            Exit For

        Next

        If endInd = -1 OrElse endInd < startInd Then Return ""

        ' The global log has nesting based on the depth of the winapp2ool fsm, we trim this to make the requested slice depth=0 

        Dim toTrim = 0

        For Each c In GlobalLog.Items(startInd)

            If Not c = CChar(" ") Then Exit For
            toTrim += 1

        Next

        Dim sb As New StringBuilder()
        For i = startInd To endInd

            If GlobalLog.Items(i).Length <= toTrim Then

                sb.AppendLine("")
                Continue For

            End If

            sb.AppendLine(GlobalLog.Items(i).Substring(toTrim))

        Next

        Return sb.ToString()

    End Function

    ''' <summary> 
    ''' Prints a slice of the global log to the user and waits for them to press the 'enter' key 
    ''' </summary>
    '''
    ''' <param name="slice">
    ''' A portion of the global log to be printed to the user 
    ''' </param>
    '''
    ''' Docs last updated: 2024-05-18 | Code last updated: 2024-05-18
    Public Sub printSlice(slice As String)

        clrConsole()
        cwl(slice)
        Console.ReadLine()

    End Sub

End Module