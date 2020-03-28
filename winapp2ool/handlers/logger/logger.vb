'    Copyright (C) 2018-2020 Robbie Ward
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
''' <summary> This module is an internal logger for winapp2ool </summary>
Public Module logger
    ''' <summary> Holds the contents of the winapp2ool log </summary>
    Public Property GlobalLog As New strList
    '''<summary> Holds the filesystem location to which the log file will be saved </summary>
    Public Property GlobalLogFile As New iniFile(Environment.CurrentDirectory, "winapp2ool.log")
    '''<summary> Indicates the current nesting level </summary>
    Public Property nestCount As Integer = 0

    ''' <summary> Adds an item into the global log </summary>
    ''' <param name="logstr"> The <c> String </c> to be added into the log <br /> Optional, Default: <c> ""</c> </param>
    ''' <param name="cond"> Indicates that the <c> <paramref name="logstr"/> </c> should be added into the log <br /> Optional, Default: <c> True </c> </param>
    ''' <param name="ascend"> Indicates that the line should be indented. (Generally) requires a corresponding <c> <paramref name="descend"/></c> to "undo." Usefor for blocking 
    ''' groups of related log items without needing to call <c> <paramref name="indent"/> </c> on each <br /> Optional, Default: <c> False </c></param>
    ''' <param name="descend"> Indicates that the line should be unindented. (Generally) follows a corresponding <c> <paramref name="ascend"/> </c>. Useful for blocking 
    ''' groups of related log items without needing to call <c> <paramref name="indent"/> </c> on each <br /> Optional, Default: <c> False </c> </param>
    ''' <param name="indent"> Indicates that the line should be indented individually, without affecting the indentation of following lines <br/> Optional, Default: <c> False </c> </param>
    ''' <param name="indAmt"> The number of times by which the line should be indented given <c> <paramref name="indent"/> </c> is <c> True </c> <br /> Optional, Default: <c> 1 </c> </param>
    ''' <param name="descAmt"> The number of times *fewer* by which following lines should be indented <br /> Optional, Default: <c> 1 </c> </param>
    ''' <param name="ascAmt"> The number of times by which this and following lines should be indented <br /> Optional, Default: <c> 1 </c> </param>
    ''' <param name="buffr"> Indicates that an empty line should be added into the log following <c> <paramref name="logstr"/> </c> </param>
    ''' <param name="leadr"> Indicates that an empty line should be added into the log before <c> <paramref name="logstr"/> </c> </param>
    Public Sub gLog(Optional logstr As String = "", Optional cond As Boolean = True, Optional ascend As Boolean = False,
                    Optional descend As Boolean = False, Optional indent As Boolean = False, Optional indAmt As Integer = 1,
                    Optional descAmt As Integer = 1, Optional ascAmt As Integer = 1, Optional buffr As Boolean = False, Optional leadr As Boolean = False)
        If cond Then
            If leadr Then gLog()
            If indent Then nestCount += indAmt
            If ascend Then nestCount += ascAmt
            Dim buffer = ""
            For i = 0 To nestCount - 1
                buffer += "  "
            Next
            logstr = buffer + logstr
            GlobalLog.add(logstr)
            If indent Then nestCount -= indAmt
            If descend Then nestCount -= descAmt
            If buffr Then gLog()
        End If
    End Sub

    ''' <summary> Saves the global log to disk if the given <c> <paramref name="cond"/> </c> is met </summary>
    ''' <param name="cond"> Indicates that the global log should be saved to disk <br /> Optional, Default: <c> False </c> </param>
    Public Sub saveGlobalLog(Optional cond As Boolean = True)
        GlobalLogFile.overwriteToFile(logger.toString, cond)
    End Sub

    ''' <summary> Returns the log as a single <c> String </c> </summary>
    Public Function toString() As String
        Dim out = ""
        GlobalLog.Items.ForEach(Sub(line) out += line & Environment.NewLine)
        Return out
    End Function

    ''' <summary> Prints the winapp2ool log to the user and waits for the 'enter' key to be pressed before returning to the calling menu </summary>
    Public Sub printLog()
        cwl("Printing the winapp2ool log, this may take a moment")
        Dim out = logger.toString
        clrConsole()
        cwl(out)
        cwl()
        cwl($"End of log. {pressEnterStr}")
        Console.ReadLine()
    End Sub

    '''<summary> Gets the most recent segment of the global log contained by two phrases (ie. a module name or subroutine) 
    ''' As a <c> String </c> <br /> Can be used to simply fetch the logs from modules for saving to disk or displaying to the user after a run.</summary>
    ''' <param name="containedPhrase"> The starting phrase of the requested log slice </param>
    ''' <param name="endingPhrase"> The ending phrase of the requested log slice </param>
    Public Function getLogSliceFromGlobal(containedPhrase As String, endingPhrase As String) As String
        Dim startind, endind As Integer
        For Each line In GlobalLog.Items
            If line.Contains(containedPhrase) Then startind = GlobalLog.Items.LastIndexOf(line)
            If line.EndsWith(endingPhrase, StringComparison.InvariantCultureIgnoreCase) Then endind = GlobalLog.Items.LastIndexOf(line)
        Next
        ' The global log has nesting based on the depth of the winapp2ool fsm, we trim this to make the requested slice depth=0 
        Dim toTrim = ""
        For Each c In GlobalLog.Items(startind)
            If Not c = CChar(" ") Then Exit For
            toTrim += " "
        Next
        Dim out = ""
        For i = startind To endind
            out += GlobalLog.Items(i).Substring(toTrim.Length) & Environment.NewLine
        Next
        Return out
    End Function

    '''<summary> Prints a slice of the global log to the user and waits for them to press the 'enter' key </summary>
    '''<param name="slice"> A portion of the global log to be printed to the user </param>
    Public Sub printSlice(slice As String)
        clrConsole()
        cwl(slice)
        Console.ReadLine()
    End Sub
End Module