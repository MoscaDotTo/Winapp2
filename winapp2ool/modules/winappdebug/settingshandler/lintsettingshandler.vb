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
''' Manages the settings of the WinappDebug module for the purpose of syncing to disk
''' </summary>
Module lintsettingshandler

    ''' <summary>
    ''' Restores the default state of the WinappDebug module's properties
    ''' </summary>
    Public Sub InitDefaultLintSettings()

        winappDebugFile1.ResetParams()
        winappDebugFile3.ResetParams()
        LintModuleSettingsChanged = False
        RepairErrsFound = True
        SaveChanges = False
        overrideDefaultVal = False
        expectedDefaultValue = False

        resetScanSettings()

        SaveLintSettings()

    End Sub

    ''' <summary>
    ''' Saves all WinappDebug settings to <c> SettingsFile2 </c>,
    ''' including the per-rule scan and repair flags.
    ''' </summary>
    Public Sub SaveLintSettings()

        SaveModule2(NameOf(WinappDebug), GetType(lintsettings))

        For Each rule In Rules
            SetSetting(NameOf(WinappDebug), rule.LintName & "_Scan", tsInvariant(rule.ShouldScan))
            SetSetting(NameOf(WinappDebug), rule.LintName & "_Repair", tsInvariant(rule.ShouldRepair))
        Next

    End Sub

    ''' <summary>
    ''' Loads per-rule scan and repair flags for WinappDebug from <c> SettingsFile2 </c>.
    ''' </summary>
    Public Sub LoadLintRulesFromSettings2()

        For Each rule In Rules
            Dim scanVal = GetSetting(NameOf(WinappDebug), rule.LintName & "_Scan")
            If scanVal <> "" Then rule.ShouldScan = CBool(scanVal)
            Dim repairVal = GetSetting(NameOf(WinappDebug), rule.LintName & "_Repair")
            If repairVal <> "" Then rule.ShouldRepair = CBool(repairVal)
        Next

    End Sub

End Module
