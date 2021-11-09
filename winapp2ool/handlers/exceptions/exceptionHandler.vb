'    Copyright (C) 2018-2021 Hazel Ward
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
Imports System.IO
''' <summary> Handles the processing of errors caught throughout the operation of winapp2ool, hopefully gracefully. </summary>
Module exceptionHandler
    ''' <summary> Catches general exceptions and logs them for debugging purposes  </summary>
    ''' <param name="ex"> Any given exception captured during winapp2ool's execution </param>
    Public Sub exc(ByRef ex As Exception)
        printAndLogExceptionForUser(ex.ToString, ex.GetType.ToString, True)
    End Sub

    ''' <summary> Enters Exceptions caused by a lack of internet access into the global log </summary>
    ''' <param name="ex"> An exception of type <c> Net.WebException </c> </param>
    Public Sub handleWebException(ex As Net.WebException)
        gLog($"winapp2ool was unable to establish a connection to GitHub and will now enter Offline Mode", leadr:=True)
        gLog("Check your network connection and try again. If you feel this is an error, please report the following information on GitHub", buffr:=True)
        gLog(ex.ToString, indent:=True, buffr:=True)
    End Sub

    ''' <summary> Enters Exceptions caused by being unable to access files into the global log </summary>
    ''' <param name="ex"> An Exception of type <c> IOException </c> </param>
    Public Sub handleIOException(ex As IOException)
        gLog("winapp2ool was unable to access a file and thus cannot complete its work. This is usually caused by another program accessing the file at the same time.", leadr:=True)
        gLog("Make sure you've closed any other applications who may be accessing the file and try again. If you feel this is an error, please report the following information on GitHub", buffr:=True)
        gLog(ex.ToString, indent:=True, buffr:=True)
        saveGlobalLog()
    End Sub

    ''' <summary> Creates a new <c> ArgumentNullException </c> when a public member is passed a <c> Null </c> parameter </summary>
    ''' <param name="argName"> The name of the argument whose value is <c> Null </c> </param>
    Public Sub argIsNull(argName As String)
        handleNullArgException(New ArgumentNullException(argName))
    End Sub

    ''' <summary> Creates a new <c> ArgumentException </c> when a public member is passed an invalid parameter </summary>
    ''' <param name="argName"> The name of the argument whose value is invalid </param>
    Public Sub argIsInvalid(argName As String)
        handleInvalidArgException(New ArgumentException(argName))
    End Sub

    ''' <summary> Passes off exceptions caused by invalid arugments to be logged and displayed to the user </summary>
    ''' <param name="ex"> An Exception of type <c> ArgumentException </c> </param>
    Public Sub handleInvalidArgException(ex As ArgumentException, Optional forceAck As Boolean = True, Optional onlyLog As Boolean = False)
        printAndLogExceptionForUser(ex.ToString, ex.GetType.ToString, forceAck, onlyLog)
    End Sub

    ''' <summary> Informs the user that an exception occured and records it in the winapp2ool log </summary>
    ''' <param name="exTxt"> The full text of the exception </param>
    ''' <param name="exType"> The <c> Type </c> of the exception </param>
    ''' <param name="forceAcknowledge"> Indicates that the user should be forced to press enter before the application continues <br /> Optional, Default: <c> False </c> </param>
    Public Sub printAndLogExceptionForUser(exTxt As String, exType As String, Optional forceAcknowledge As Boolean = False, Optional onlyLog As Boolean = False)
        gLog($"{exType} Encountered!")
        gLog("Please report the following information on GitHub: ", ascend:=True, buffr:=True)
        gLog(exTxt)
        gLog("The winapp2ool GitHub is https://github.com/MoscaDotTo/Winapp2")
        gLog("The old winapp2.ini website, https://www.winapp2.com also redirects to GitHub for your convenience")
        gLog("A link to our GitHub can be found in the winapp2ool settings as well!")
        gLog(descend:=True)
        If Not onlyLog Then
            cwl($"Error: {exType} Encountered")
            cwl(exTxt)
            cwl("Please report this error on GitHub. It will be saved to winapp2ool.log in the same folder as winapp2ool.")
            saveGlobalLog()
            If forceAcknowledge Then cwl(pressEnterStr) : Console.ReadLine()
        End If
    End Sub

    ''' <summary> Passes off exceptions caused by Null arguments to be logged and displayed to the user </summary>
    ''' <param name="ex"> An Exception of type <c> ArgumentNullException </c> </param>
    Public Sub handleNullArgException(ex As ArgumentNullException)
        printAndLogExceptionForUser(ex.ToString, ex.GetType.ToString, True)
    End Sub
End Module