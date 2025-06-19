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
Module lintsettingshandler

    ''' <summary>
    ''' The names of each type of error supported by WinappDebug
    ''' </summary>
    ''' 
    ''' Docs last updated: 2023-07-20 | Code last updated: 2023-07-20
    Private Property Lints As New List(Of String) From {"Casing", "Alphabetization", "Improper Numbering", "Parameters",
                                                        "Flags", "Slashes", "Defaults", "Duplicates", "Unneeded Numbering",
                                                        "Multiples", "Invalid Values", "Syntax Errors", "Path Validity",
                                                        "Semicolons", "Optimizations", "Potential Duplicate Keys Between Entries"}

    ''' <summary> 
    ''' Restore the default state of all of the module's parameters, undoing any changes the user may have made to them 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2023-07-19 | Code last updated: 2023-07-19
    Public Sub InitDefaultLintSettings()

        winappDebugFile1.resetParams()
        winappDebugFile3.resetParams()
        ModuleSettingsChanged = False
        RepairErrsFound = True
        SaveChanges = False
        overrideDefaultVal = False
        expectedDefaultValue = False

        resetScanSettings()

        restoreDefaultSettings(NameOf(WinappDebug), AddressOf CreateLintSettingsSection)

    End Sub

    ''' <summary> 
    ''' Loads the WinappDebug settings from disk and loads them into memory, overriding the default settings 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2021-11-13 | Code last updated: 2021-11-13
    Public Sub getSeralizedLintSettings()

        If Not readSettingsFromDisk Then Return

        For Each kvp In settingsDict(NameOf(WinappDebug))

            Select Case kvp.Key

                Case NameOf(winappDebugFile1) & "_Dir"

                    winappDebugFile1.Dir = kvp.Value

                Case NameOf(winappDebugFile1) & "_Name"

                    winappDebugFile1.Name = kvp.Value

                Case NameOf(winappDebugFile3) & "_Dir"

                    winappDebugFile3.Dir = kvp.Value

                Case NameOf(winappDebugFile3) & "_Name"

                    winappDebugFile3.Name = kvp.Value

                Case NameOf(RepairSomeErrsFound)

                    RepairSomeErrsFound = CBool(kvp.Value)

                Case NameOf(ScanSettingsChanged)

                    ScanSettingsChanged = CBool(kvp.Value)

                Case NameOf(ModuleSettingsChanged)

                    ModuleSettingsChanged = CBool(kvp.Value)

                Case NameOf(SaveChanges)

                    SaveChanges = CBool(kvp.Value)

                Case NameOf(RepairErrsFound)

                    RepairErrsFound = CBool(kvp.Value)

                Case NameOf(overrideDefaultVal)

                    overrideDefaultVal = CBool(kvp.Value)


                Case NameOf(expectedDefaultValue)

                    expectedDefaultValue = CBool(kvp.Value)

                Case Else

                    CheckScanRepair(kvp)

            End Select

        Next

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

    '''<summary>
    '''Adds the current (typically default) state of the module's settings into the disk-writable settings representation 
    '''</summary>
    '''
    ''' Docs last updated: 2021-11-13 | Code last updated: 2021-11-13
    Public Sub CreateLintSettingsSection()

        Dim settingsKeys As New List(Of String) From {
                                                      NameOf(RepairSomeErrsFound), tsInvariant(RepairSomeErrsFound),
                                                      NameOf(ScanSettingsChanged), tsInvariant(ScanSettingsChanged),
                                                      NameOf(ModuleSettingsChanged), tsInvariant(ModuleSettingsChanged),
                                                      NameOf(SaveChanges), tsInvariant(SaveChanges),
                                                      NameOf(RepairErrsFound), tsInvariant(RepairErrsFound),
                                                      NameOf(overrideDefaultVal), tsInvariant(overrideDefaultVal),
                                                      NameOf(expectedDefaultValue), tsInvariant(expectedDefaultValue)}

        For i = 0 To Lints.Count - 1

            settingsKeys.Add($"{Lints(i)}_Scan")
            settingsKeys.Add(tsInvariant(Rules(i).ShouldScan))
            settingsKeys.Add($"{Lints(i)}_Repair")
            settingsKeys.Add(tsInvariant(Rules(i).ShouldRepair))

        Next

        settingsKeys.AddRange({NameOf(winappDebugFile1), winappDebugFile1.Name, winappDebugFile1.Dir, NameOf(winappDebugFile3), winappDebugFile3.Name, winappDebugFile3.Dir})
        createModuleSettingsSection(NameOf(WinappDebug), settingsKeys, 39, 2)

    End Sub

End Module