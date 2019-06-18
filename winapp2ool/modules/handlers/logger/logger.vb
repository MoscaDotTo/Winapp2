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
''' This module handles logging for winapp2ool. 
''' </summary>
Module logger
    ''' <summary> Holds the contents of the winapp2ool log</summary>
    Public Property GlobalLog As New strList
    '''<summary>Holds the save path information for the winapp2ool log</summary>
    Public Property GlobalLogFile As New iniFile(Environment.CurrentDirectory, "winapp2ool.log")
    '''<summary>Indicates the current nesting level within the state machine </summary>
    Public Property nestCount As Integer = 0
    ''' <summary>Adds an item into the global log</summary>
    Public Sub gLog(logstr As String, Optional cond As Boolean = True, Optional ascend As Boolean = False,
                    Optional descend As Boolean = False, Optional indent As Boolean = False, Optional indAmt As Integer = 1,
                    Optional descAmt As Integer = 1, Optional ascAmt As Integer = 1, Optional buffr As Boolean = False, Optional leadr As Boolean = False)
        If cond Then
            If leadr Then gLog("")
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
            If buffr Then gLog("")
        End If
    End Sub

    ''' <summary>Returns the log as a String</summary>
    Public Function toString() As String
        Dim out = ""
        GlobalLog.Items.ForEach(Sub(line) out += line & Environment.NewLine)
        Return out
    End Function

    Public Sub printLog()
        clrConsole()
        cwl(logger.toString)
        cwl()
        cwl()
        cwl("End of log. Press any key to continue.")
        Console.ReadLine()
    End Sub

    '''<summary>Gets the most recent  segment of the global log contained by two phrases (ie. a module name or subroutine) 
    '''As a string. Can be used to simply fetch the logs from modules for saving to disk or displaying to the user after a run.</summary>
    ''' <param name="containedPhrase">The starting phrase of the requested log slice</param>
    ''' <param name="endingPhrase">The ending phease of the requested log slice</param>
    Public Function getLogSliceFromGlobal(containedPhrase As String, endingPhrase As String) As String
        Dim startind, endind As Integer
        For Each line In GlobalLog.Items
            If line.Contains(containedPhrase) Then startind = GlobalLog.Items.LastIndexOf(line)
            If line.EndsWith(endingPhrase) Then endind = GlobalLog.Items.LastIndexOf(line)
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
End Module