'    Copyright (C) 2018-2019 Robbie Ward
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
''' <summary>
''' Handles the processing of errors caught throughout the operation of winapp2ool, hopefully gracefully.
''' </summary>
Module exceptionHandler
    ''' <summary>Prints out exceptions and any other information related to them that a use may need.</summary>
    ''' <param name="ex">A given exception captured during winapp2ool's execution</param>
    Public Sub exc(ByRef ex As Exception)
        gLog("Exception Encountered!", indent:=True, ascend:=True)
        Select Case True
            Case ex.GetType.FullName = "System.Net.WebException"
                gLog(ex.Message, ascend:=True)
                ' If we can't connect to GitHub we may as well be offline since that's where all our online resources 
                gLog($"winapp2ool's current online status is detected as: {Not isOffline}. It will now be set to False", indent:=True, descend:=True)
                isOffline = True
            ' Out of date .NET presents us from downloading securely from GitHub as is required for executables.
            Case ex.Message.Contains("SSL/TLS")
                printDotNetOutOfDateError()
            Case Else
                gLog(ex.ToString)
                cwl("Error: " & ex.ToString)
                cwl("Please report this error on GitHub")
                cwl("Press Enter to continue")
                Console.ReadLine()
        End Select
        gLog("", descend:=True)
    End Sub

    ''' <summary>Prints output to the user instructing them to update their .NET Framework</summary>
    Private Sub printDotNetOutOfDateError()
        cwl("Error: download could not be completed.")
        cwl("This issue is caused by an out of date .NET Framework.")
        cwl("Please update .NET Framework to version 4.6 or higher and try again.")
        cwl("If the issue persists after updating .NET Framework, please report this error on GitHub.")
    End Sub
End Module