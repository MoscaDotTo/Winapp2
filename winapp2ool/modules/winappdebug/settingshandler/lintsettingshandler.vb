'    Copyright (C) 2018-2025 Hazel Ward
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
''' 
''' Docs last updated: 2024-05-08
Module lintsettingshandler

    ''' <summary>
    ''' The names of each type of error supported by WinappDebug
    ''' </summary>
    ''' 
    ''' Docs last updated: 2023-07-20 | Code last updated: 2023-07-20
    Private ReadOnly Property Lints As New List(Of String) From {
        "Casing",
        "Alphabetization",
        "Improper Numbering",
        "Parameters",
        "Flags",
        "Slashes",
        "Defaults",
        "Duplicates",
        "Unneeded Numbering",
        "Multiples",
        "Invalid Values",
        "Syntax Errors",
        "Path Validity",
        "Semicolons",
        "Optimizations",
        "Potential Duplicate Keys Between Entries"
    }

    ''' <summary> 
    ''' Restores the default state of the WinappDebug module's properties 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2024-05-08 | Code last updated: 2023-07-19
    Public Sub InitDefaultLintSettings()

        winappDebugFile1.resetParams()
        winappDebugFile3.resetParams()
        LintModuleSettingsChanged = False
        RepairErrsFound = True
        SaveChanges = False
        overrideDefaultVal = False
        expectedDefaultValue = False

        resetScanSettings()

        restoreDefaultSettings(NameOf(WinappDebug), AddressOf CreateLintSettingsSection)

    End Sub

    ''' <summary> 
    ''' Assigns the module settings to WinappDebug based on the current disk-writable settings representation
    ''' </summary>
    ''' 
    ''' Docs last updated: 2024-05-08 | Code last updated: 2024-05-08
    Public Sub getSerializedLintSettings()

        loadLintRulesFromDict()
        LoadModuleSettingsFromDict(NameOf(WinappDebug), GetType(lintsettings))

    End Sub

    ''' <summary>
    ''' Handles the checking of the scan and repair settings for each lint type and rejects invalid settings 
    ''' </summary>
    ''' 
    ''' <param name="kvp">
    ''' A <c> KeyValuePair </c> containing a setting name and a value for that setting 
    ''' <br /> In this case, we expect the setting name to be a lint type (either a scan or repair setting) and the value to be a boolean
    ''' </param>
    ''' 
    ''' Docs last updated: 2023-07-20 | Code last updated: 2023-07-20
    Private Sub CheckScanRepair(kvp As KeyValuePair(Of String, String))

        Dim lintType = kvp.Key.Replace("_Scan", "").Replace("_Repair", "")
        Dim ind = Lints.IndexOf(lintType)

        Try

            If kvp.Key.Contains("_Scan") Then Rules(ind).ShouldScan = CBool(kvp.Value)
            If kvp.Key.Contains("_Repair") Then Rules(ind).ShouldRepair = CBool(kvp.Value)

        Catch ex As ArgumentOutOfRangeException

            gLog($"{kvp.Key} doesn't seem to be an actual setting, perhaps it is misnamed or the setting name has changed. This value will be ignored")

        End Try

    End Sub

    ''' <summary>
    ''' Adds the current state of the module's settings into the disk-writable settings representation 
    ''' </summary>
    ''' 
    ''' <remarks>
    ''' Most often, this is the default state of these settings 
    ''' </remarks>
    ''' 
    ''' Docs last updated: 2024-05-08 | Code last updated: 2025-06-25
    Public Sub CreateLintSettingsSection()

        Dim settingsModule = GetType(lintsettings)
        Dim moduleName = NameOf(WinappDebug)

        Dim lintSettingsTuples = GetSettingsTupleWithReflection(settingsModule)

        ' We add the individual lint scan settings to the tuple and provide it to the initalizer
        For i = 0 To Lints.Count - 1

            lintSettingsTuples.Add($"{Lints(i)}_Scan")
            lintSettingsTuples.Add(tsInvariant(Rules(i).ShouldScan))
            lintSettingsTuples.Add($"{Lints(i)}_Repair")
            lintSettingsTuples.Add(tsInvariant(Rules(i).ShouldRepair))

        Next

        Dim addlBools = Lints.Count * 2
        createModuleSettingsSection(moduleName, settingsModule, lintSettingsTuples, addlBools)

    End Sub

    ''' <summary>
    ''' Loads the <c>LintRule</c>s from the disk-writable settings representation into the module's settings
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-06-25 | Code last updated: 2025-06-25
    Private Sub loadLintRulesFromDict()

        If Not settingsDict.ContainsKey("WinappDebug") Then Return

        Dim ruleDict = settingsDict("WinappDebug")

        For Each rule In Rules

            If ruleDict.ContainsKey($"{rule.LintName}_Scan") Then rule.ShouldScan = CBool(ruleDict($"{rule.LintName}_Scan"))

            If ruleDict.ContainsKey($"{rule.LintName}_Repair") Then rule.ShouldRepair = CBool(ruleDict($"{rule.LintName}_Repair"))

        Next

    End Sub

End Module