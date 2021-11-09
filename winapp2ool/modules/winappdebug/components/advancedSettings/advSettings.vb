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
''' <summary> This module holds some methods for managing WinappDebug's scan/repair states </summary>
Public Module advSettings

    ''' <summary> Prints the scan/repair management menu to the user </summary>
    Public Sub printMenu()
        Console.WindowHeight = 51
        printMenuTop({"Enable or disable specific scans or repairs"})
        print(0, "Scan Options", leadingBlank:=True, trailingBlank:=True)
        Rules.ForEach(Sub(rule) print(5, rule.LintName, rule.ScanText, enStrCond:=rule.ShouldScan))
        ' Print all repairs except the last one
        print(0, "Repair Options", leadingBlank:=True, trailingBlank:=True)
        For i = 0 To Rules.Count - 2
            Dim rule = Rules(i)
            print(5, rule.LintName, rule.RepairText, enStrCond:=rule.ShouldRepair)
        Next
        ' Special case for the last repair option (closemenu flag)
        Dim lastRule = Rules.Last
        print(5, lastRule.LintName, lastRule.RepairText, closeMenu:=Not ScanSettingsChanged, enStrCond:=lastRule.ShouldRepair)
        print(2, "Scan And Repair", cond:=ScanSettingsChanged, closeMenu:=True)
    End Sub

    ''' <summary> Handles the user input for the scan/repair management menu </summary>
    ''' <param name="input"> The String containing the user's input </param>
    Public Sub handleUserInput(input As String)
        Dim lints As New List(Of String) From {"Casing", "Alphabetization", "Improper Numbering", "Parameters", "Flags", "Slashes", "Defaults", "Duplicates", "Unneeded Numbering",
                "Multiples", "Invalid Values", "Syntax Errors", "Path Validity", "Semicolons", "Optimizations", "Potential Duplicate Keys Between Entries"}
        ' Determine the current state of the lint rules
        determineScanSettings()
        ' Get the input as an integer so we can index it against our rules
        Dim intInput = -1
        If Not Integer.TryParse(input, intInput) Then
            ' This isn't an error since we have the "alloff" command, but the compiler throws a warning if we don't check that this is successful
            ' If the "alloff" debugging command is removed though, this will be a case of an invalid input since there is no default option in this menu
        End If
        ' The index of the rule assoicated with the user's input
        Dim ind = intInput - 1
        Select Case True
            Case input = "0"
                If ScanSettingsChanged Then WinappDebug.ModuleSettingsChanged = True
                exitModule()
        ' Enable/Disable individual scans
            Case intInput > 0 And intInput <= Rules.Count
                toggleSettingParam(Rules(ind).ShouldScan, "Scan", ScanSettingsChanged, NameOf(WinappDebug), lints(ind) & "_Scan", NameOf(ScanSettingsChanged))
                ' Force repair off if the scan is off
                If Not Rules(ind).ShouldScan Then Rules(ind).turnOff()
        ' Enable/Disable individual repairs
            Case intInput > Rules.Count And intInput <= 2 * Rules.Count
                ind -= (Rules.Count)
                toggleSettingParam(Rules(ind).ShouldRepair, "Repair", ScanSettingsChanged, NameOf(WinappDebug), lints(ind) & "_Repair", NameOf(ScanSettingsChanged))
                ' Force scan on if the repair is on
                If Rules(ind).ShouldRepair Then Rules(ind).turnOn()
            Case intInput = 2 * Rules.Count + 1 And ScanSettingsChanged
                resetScanSettings()
                setHeaderText("Settings Reset")
        ' This isn't documented anywhere and is mostly intended as a debugging shortcut
            Case input = "alloff"
                Rules.ForEach(Sub(rule) rule.turnOff())
                ScanSettingsChanged = True
            Case Else
                setHeaderText(invInpStr, True)
        End Select
    End Sub

    ''' <summary> Determines which if any lint rules have been modified and whether or not only some repairs are scheduled to run </summary>
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

    ''' <summary> Resets the individual scan/repair settings to their defaults </summary>
    Public Sub resetScanSettings()
        Rules.ForEach(Sub(rule) rule.resetParams())
        ScanSettingsChanged = False
        RepairSomeErrsFound = False
    End Sub
End Module