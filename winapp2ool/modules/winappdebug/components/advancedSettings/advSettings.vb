'    Copyright (C) 2018-2026 Hazel Ward
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

    ''' <summary>
    ''' Builds the scan/repair settings menu with all toggles and their dispatch handlers registered inline.
    ''' Called by both <c> printMenu </c> (to render) and <c> handleUserInput </c>
    ''' (to dispatch), so the displayed option numbers and the dispatch table are always in sync.
    ''' </summary>
    Private Function buildAdvSettingsMenu() As MenuSection

        Dim menu = MenuSection.CreateCompleteMenu("Scan Settings",
                                                {"Enable or disable specific scans or repairs"})

        menu.AddBlank()
        menu.AddLine("Scan Options", centered:=True)
        menu.AddDivider(solid:=False)

        For Each rule In Rules

            Dim r = rule
            menu.AddDispatchedToggle(r.LintName, r.ScanText, r.ShouldScan,
                Sub()
                    Dim prev = r.ShouldScan
                    gLog($"  Toggling Scan from {prev} to {Not prev}")
                    setNextMenuHeaderText($"Scan {enStr(prev)}d", printColor:=If(Not prev, ConsoleColor.Green, ConsoleColor.Red))
                    r.ShouldScan = Not prev
                    ScanSettingsChanged = True
                    SetSetting(NameOf(WinappDebug), r.LintName & "_Scan", tsInvariant(r.ShouldScan))
                    FlushIfDirty2()
                    If Not r.ShouldScan Then r.turnOff()
                End Sub)

        Next

        menu.AddBlank()
        menu.AddLine("Repair Options", centered:=True)
        menu.AddDivider(solid:=False)

        For Each rule In Rules

            Dim r = rule
            menu.AddDispatchedToggle(r.LintName, r.RepairText, r.ShouldRepair,
                Sub()
                    Dim prev = r.ShouldRepair
                    gLog($"  Toggling Repair from {prev} to {Not prev}")
                    setNextMenuHeaderText($"Repair {enStr(prev)}d", printColor:=If(Not prev, ConsoleColor.Green, ConsoleColor.Red))
                    r.ShouldRepair = Not prev
                    ScanSettingsChanged = True
                    SetSetting(NameOf(WinappDebug), r.LintName & "_Repair", tsInvariant(r.ShouldRepair))
                    FlushIfDirty2()
                    If r.ShouldRepair Then r.turnOn()
                End Sub)

        Next

        menu.AddBlank(ScanSettingsChanged)
        menu.AddDispatchedResetOpt("Scan And Repair", ScanSettingsChanged,
            Sub()
                resetScanSettings()
                setNextMenuHeaderText("Settings Reset", printColor:=ConsoleColor.Yellow)
            End Sub)

        Return menu

    End Function

    ''' <summary> 
    ''' Prints the scan/repair management menu to the user 
    ''' </summary>
    Public Sub printMenu()

        Console.WindowHeight = 52
        buildAdvSettingsMenu().Print()

    End Sub

    ''' <summary> 
    ''' Handles the user input for the scan/repair management menu 
    ''' </summary>
    ''' 
    ''' <param name="input">
    ''' The String containing the user's input 
    ''' </param>
    Public Sub handleUserInput(input As String)

        determineScanSettings()

        If input = "alloff" Then

            Rules.ForEach(Sub(rule) rule.turnOff())
            ScanSettingsChanged = True
            Return

        End If

        Dim intInput As Integer

        If Not Integer.TryParse(input, intInput) Then

            setNextMenuHeaderText(invInpStr, printColor:=ConsoleColor.Red)
            Return

        End If

        If intInput = 0 Then

            If ScanSettingsChanged Then LintModuleSettingsChanged = True
            exitModule()
            Return

        End If

        If Not buildAdvSettingsMenu().Dispatch(intInput) Then setNextMenuHeaderText(invInpStr, printColor:=ConsoleColor.Red)

    End Sub

    ''' <summary> 
    ''' Determines which if any lint rules have been modified and whether or not only some repairs are scheduled to run 
    ''' </summary>
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

        RepairErrsFound = repairAll
        RepairSomeErrsFound = repairAny

    End Sub

    ''' <summary>
    ''' Resets the individual scan/repair settings to their defaults 
    ''' </summary>
    Public Sub resetScanSettings()

        Rules.ForEach(Sub(rule) rule.resetParams())
        ScanSettingsChanged = False
        RepairSomeErrsFound = False

    End Sub

End Module
