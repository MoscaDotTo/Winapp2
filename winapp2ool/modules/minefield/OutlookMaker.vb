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
Module OutlookMaker
    Public Property OutlookerFile1 As New iniFile(Environment.CurrentDirectory, "winapp2.ini", "winapp2-outlooker-merged.ini")
    Public Property OutlookerFile2 As New iniFile(Environment.CurrentDirectory, "outlook.ini")


    Public Sub printOutlookMenu()
        printMenuTop({"This tool will attempt to generate winapp2.ini entries with support for custom outlook profiles"})
        print(1, "Run (Default)", "Run the tool", closeMenu:=True)
    End Sub

    Public Sub handleOutlookInput(input As String)
        Select Case input
            Case "0"
                ExitCode = True
            Case "1"
                initOutlooker()
            Case Else
                setHeaderText(invInpStr, True)
        End Select
    End Sub

    Private Sub initOutlooker()
        OutlookerFile2.validate()
        cwl("Right now we support only attempting to generate an entry for recent searches")
        cwl()
        cwl("Scanning for valid detects")
        Dim recents = OutlookerFile2.Sections.First.Value
        cwl($"Scanning for {recents.Keys.Keys(1).Value}")
        Dim tmp = getWildcardKeys(recents.Keys.Keys(1).Value)
        Dim tmp2 = getWildcardKeys(recents.Keys.Keys(2).Value)

    End Sub
End Module
