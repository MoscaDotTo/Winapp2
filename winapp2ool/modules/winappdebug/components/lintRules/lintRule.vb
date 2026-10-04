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
''' <summary> Holds information about whether or not individual types of scans and repairs should run </summary>
Public Class lintRule
    ''' <summary> The scan state with which the rule was initially instantiated </summary>
    Private Property initScanState As Boolean
    ''' <summary> The repair state with which the rule was initially instantiated </summary>
    Private Property initRepairState As Boolean

    ''' <summary>
    ''' Indicates whether this rule's checks run. Most of its repairs only fire on an error the
    ''' check reported, so turning this off usually stops them too.
    ''' </summary>
    Public Property ShouldScan As Boolean

    ''' <summary>
    ''' Indicates whether this rule's repairs should run. Most repairs go through
    ''' <see cref="fixFormat"/>, which ignores this while <see cref="RepairErrsFound"/> is <c> True </c>.
    ''' </summary>
    Public Property ShouldRepair As Boolean

    ''' <summary> Describes to the user which type of scan routines this rule controls, starting with <c> detecting </c> </summary>
    Public Property ScanText As String

    ''' <summary> Describes to the user which type of repairs this rule controls </summary>
    Public Property RepairText As String

    ''' <summary> The name of the rule as it will appear in menus </summary>
    Public Property LintName As String

    ''' <summary> Returns whether the current scan/repair settings differ from the ones the rule was created with </summary>
    Public Function hasBeenChanged() As Boolean
        Return Not ShouldScan = initScanState Or Not ShouldRepair = initRepairState
    End Function

    ''' <summary> Restores the initial scan/repair states </summary>
    Public Sub resetParams()
        ShouldScan = initScanState
        ShouldRepair = initRepairState
    End Sub

    ''' <summary> Creates a new <c> lintRule </c> and keeps its scan and repair states for <see cref="resetParams"/> to restore </summary>
    ''' <param name="scan"> Indicates whether the rule's scan is on by default </param>
    ''' <param name="repair"> Indicates whether the rule's repair is on by default </param>
    ''' <param name="name"> The name of the rule as it will appear in menus and in the saved settings </param>
    ''' <param name="scTxt"> The description of what the rule scans for. We prefix it with <c> detecting </c> for the menu </param>
    ''' <param name="rpTxt"> The description of what the rule repairs as it will appear in menus </param>
    Public Sub New(scan As Boolean, repair As Boolean, name As String, scTxt As String, rpTxt As String)
        ShouldScan = scan
        initScanState = scan
        ShouldRepair = repair
        initRepairState = repair
        LintName = name
        ScanText = $"detecting {scTxt}"
        RepairText = rpTxt
    End Sub

    ''' <summary> Enables both the scan and repair for the rule </summary>
    Public Sub turnOn()
        ShouldScan = True
        ShouldRepair = True
    End Sub

    ''' <summary> Disables both the scan and repair for the rule </summary>
    Public Sub turnOff()
        ShouldScan = False
        ShouldRepair = False
    End Sub

    ''' <summary>
    ''' Returns whether the repairs gated by this rule should run: always while
    ''' <see cref="RepairErrsFound"/> is <c> True </c>, otherwise only when both
    ''' <see cref="RepairSomeErrsFound"/> and <see cref="ShouldRepair"/> are
    ''' </summary>
    Public Function fixFormat() As Boolean
        Return RepairErrsFound Or (RepairSomeErrsFound And ShouldRepair)
    End Function
End Class