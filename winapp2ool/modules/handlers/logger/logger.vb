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
                    Optional descend As Boolean = False, Optional indent As Boolean = False, Optional buffr As Boolean = False, Optional leadr As Boolean = False)
        If cond Then
            If leadr Then gLog("")
            If indent Then nestCount += 1
            If ascend Then nestCount += 1
            Dim buffer = ""
            For i = 0 To nestCount - 1
                buffer += "  "
            Next
            logstr = buffer + logstr
            GlobalLog.add(logstr)
            If indent Then nestCount -= 1
            If descend Then nestCount -= 1
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
End Module