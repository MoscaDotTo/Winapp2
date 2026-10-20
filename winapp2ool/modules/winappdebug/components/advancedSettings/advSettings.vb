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
''' <summary>
''' Runs WinappDebug's Scan Settings menu, where the user sets each rule's scan and repair
''' toggles. <see cref="LintOptions.FromSettings"/> derives a run's repair flags from them.
''' </summary>
Public Module advSettings

    ''' <summary>
    ''' Builds the scan/repair settings menu. Turning a rule's scan off also turns its repair
    ''' off, and turning a repair on also turns its scan on. We first check whether any rule
    ''' differs from its defaults, so the Reset option shows whenever there's something to reset.
    ''' The hidden <c> alloff </c> command turns every scan and repair off.
    ''' </summary>
    Friend Function buildAdvSettingsMenu() As MenuSection

        determineScanSettings()

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
                    If Not r.ShouldScan Then r.turnOff()
                    saveRules()
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
                    If r.ShouldRepair Then r.turnOn()
                    saveRules()
                End Sub)

        Next

        menu.AddBlank(ScanSettingsChanged)
        menu.AddDispatchedResetOpt("Scan And Repair", ScanSettingsChanged,
            Sub()
                resetScanSettings()
                saveRules()
                setNextMenuHeaderText("Settings Reset", printColor:=ConsoleColor.Yellow)
            End Sub)

        menu.OnText(
            Function(input)
                If input <> "alloff" Then Return False
                Rules.ForEach(Sub(rule) rule.turnOff())
                ScanSettingsChanged = True
                saveRules()
                Return True
            End Function)

        Return menu

    End Function

    ''' <summary>
    ''' Saves every rule's scan and repair, since a toggle can change both, and writes the
    ''' settings to disk if saving is on
    ''' </summary>
    Private Sub saveRules()

        SaveLintRules()
        FlushIfDirty()

    End Sub

    ''' <summary>
    ''' Sets <see cref="ScanSettingsChanged"/> when any rule differs from its defaults (it never
    ''' clears it)
    ''' </summary>
    Private Sub determineScanSettings()

        If Rules.Any(Function(rule) rule.hasBeenChanged) Then ScanSettingsChanged = True

    End Sub

    ''' <summary>
    ''' Resets the individual scan/repair settings to their defaults and clears
    ''' <see cref="ScanSettingsChanged"/>
    ''' </summary>
    Public Sub resetScanSettings()

        Rules.ForEach(Sub(rule) rule.resetParams())
        ScanSettingsChanged = False

    End Sub

End Module
