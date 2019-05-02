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
''' This module holds some methods for managing WinappDebug's scan/repair states
''' </summary>
Public Module advSettings

    ''' <summary>Resets the individual scan settings to their defaults</summary>
    Public Sub resetScanSettings()
        For Each rule In Rules
            rule.resetParams()
        Next
        ScanSettingsChanged = False
        RepairSomeErrsFound = False
    End Sub

    ''' <summary>Determines which if any lint rules have been modified and whether or not only some repairs are scheduled to run</summary>
    Private Sub determineScanSettings()
        Dim repairAll = True
        Dim repairAny = False
        For Each rule In Rules
            If rule.hasBeenChanged Then
                ScanSettingsChanged = True
                If Not rule.ShouldRepair Then repairAll = False
            End If
            If rule.ShouldRepair Then repairAny = True
        Next
        If Not repairAll And repairAny Then
            RepairErrsFound = False
            RepairSomeErrsFound = True
        End If
    End Sub

    ''' <summary>Handles the user input for the scan settings menu</summary>
    ''' <param name="input">The String containing the user's input</param>
    Public Sub handleUserInput(input As String)
        ' Determine the current state of the lint rules
        determineScanSettings()
        Dim scanNums = {"1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14"}
        Dim repNums = {"15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28"}
        Select Case True
            Case input = "0"
                If ScanSettingsChanged Then WinappDebug.ModuleSettingsChanged = True
                ExitCode = True
            ' Enable/Disable individual scans
            Case scanNums.Contains(input)
                Dim ind = scanNums.ToList.IndexOf(input)
                toggleSettingParam(Rules(ind).ShouldScan, "Scan", ScanSettingsChanged)
                ' Force repair off if the scan is off
                If Not Rules(ind).ShouldScan Then Rules(ind).turnOff()
            ' Enable/Disable individual repairs
            Case repNums.Contains(input)
                Dim ind = repNums.ToList.IndexOf(input)
                toggleSettingParam(Rules(ind).ShouldRepair, "Repair", ScanSettingsChanged)
                ' Force scan on if the repair is on
                If Rules(ind).ShouldRepair Then Rules(ind).turnOn()
            Case input = "29" And ScanSettingsChanged
                resetScanSettings()
                setHeaderText("Settings Reset")
            ' This isn't documented anywhere and is mostly intended as a debugging shortcut
            Case input = "alloff"
                For Each rule In WinappDebug.Rules
                    rule.turnOff()
                Next
                ScanSettingsChanged = True
            Case Else
                setHeaderText(invInpStr, True)
        End Select
    End Sub

    ''' <summary>Prints the menu for individual scans and their repairs to the user</summary>
    Public Sub printMenu()
        Console.WindowHeight = 48
        printMenuTop({"Enable or disable specific scans or repairs"})
        print(0, "Scan Options", leadingBlank:=True, trailingBlank:=True)
        Dim curRules = WinappDebug.Rules
        For Each rule In curRules
            print(5, rule.LintName, rule.ScanText, enStrCond:=rule.ShouldScan)
        Next
        print(0, "Repair Options", leadingBlank:=True, trailingBlank:=True)
        For i = 0 To curRules.Count - 2
            Dim rule = curRules(i)
            print(5, rule.LintName, rule.RepairText, enStrCond:=rule.ShouldRepair)
        Next
        ' Special case for the last repair option (closemenu flag)
        Dim lastRule = curRules.Last
        print(5, lastRule.LintName, lastRule.RepairText, closeMenu:=Not ScanSettingsChanged, enStrCond:=lastRule.ShouldRepair)
        print(2, "Scan And Repair", cond:=ScanSettingsChanged, closeMenu:=True)
    End Sub
End Module