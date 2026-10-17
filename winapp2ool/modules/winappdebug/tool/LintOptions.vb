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
''' The rules and settings a single lint run uses. <see cref="WinappDebug.Debug"/> reads these
''' instead of the module's settings, so a caller from another module builds its own and never
''' touches what the user saved. The rules are copies, so changing them changes this run only.
''' A new instance holds the defaults, and <see cref="FromSettings"/> copies the user's
''' current choices.
''' </summary>
Public Class LintOptions

    ''' <summary> This run's copies of the lint rules, in <see cref="WinappDebug.Rules"/> order </summary>
    Public ReadOnly Property Rules As IReadOnlyList(Of lintRule)

    ''' <summary> Controls scan/repairs for CamelCasing issues </summary>
    Public ReadOnly Property Casing As lintRule

    ''' <summary> Controls scan/repairs for alphabetization issues </summary>
    Public ReadOnly Property Alpha As lintRule

    ''' <summary> Controls scan/repairs for incorrectly numbered keys </summary>
    Public ReadOnly Property WrongNums As lintRule

    ''' <summary> Controls scan/repairs for parameters inside of FileKeys </summary>
    Public ReadOnly Property Params As lintRule

    ''' <summary> Controls scan/repairs for flags in ExcludeKeys and FileKeys </summary>
    Public ReadOnly Property Flags As lintRule

    ''' <summary> Controls scan/repairs for improper slash usage </summary>
    Public ReadOnly Property Slashes As lintRule

    ''' <summary>
    ''' Controls scan/repairs for Default keys of any value, or for the value check when
    ''' <see cref="OverrideDefaultValue"/> is on
    ''' </summary>
    Public ReadOnly Property Defaults As lintRule

    ''' <summary> Controls scan/repairs for duplicate values </summary>
    Public ReadOnly Property Dupes As lintRule

    ''' <summary> Controls scan/repairs for keys with numbers they shouldn't have </summary>
    Public ReadOnly Property ExtraNums As lintRule

    ''' <summary> Controls scan/repairs for keys which should only occur once </summary>
    Public ReadOnly Property Multi As lintRule

    ''' <summary> Controls scan/repairs for keys with invalid values </summary>
    Public ReadOnly Property Invalid As lintRule

    ''' <summary> Controls scan/repairs for winapp2.ini syntax errors </summary>
    Public ReadOnly Property Syntax As lintRule

    ''' <summary> Controls scan/repairs for invalid file or registry paths </summary>
    Public ReadOnly Property PathValidity As lintRule

    ''' <summary> Controls scan/repairs for improper use of semicolons </summary>
    Public ReadOnly Property Semis As lintRule

    ''' <summary> Controls the FileKey merge in <see cref="cOptimization"/> </summary>
    Public ReadOnly Property Opti As lintRule

    ''' <summary>
    ''' Indicates whether every rule's repairs run, whatever its own repair toggle says. See
    ''' <see cref="lintRule.fixFormat"/>.
    ''' </summary>
    Public Property RepairAll As Boolean

    ''' <summary>
    ''' Indicates whether the repairs of rules whose repair toggle is on run while
    ''' <see cref="RepairAll"/> is off
    ''' </summary>
    Public Property RepairSome As Boolean

    ''' <summary>
    ''' Indicates whether every entry must have a Default key with the value
    ''' <see cref="ExpectedDefaultValue"/>, instead of having none
    ''' <br /> Default: <c> False </c>
    ''' </summary>
    Public Property OverrideDefaultValue As Boolean = False

    ''' <summary>
    ''' The value Default keys must have while <see cref="OverrideDefaultValue"/> is on
    ''' <br /> Default: <c> False </c>
    ''' </summary>
    Public Property ExpectedDefaultValue As Boolean = False

    ''' <summary>
    ''' Indicates whether existing Default keys are left alone instead of reported and removed
    ''' <br /> Default: <c> False </c>
    ''' </summary>
    Public Property PreserveDefaultKeys As Boolean = False

    ''' <summary>
    ''' The flavor whose own checks apply, such as BleachBit's ExcludeKey REG flag check
    ''' <br /> Default: <c> WinappFlavor.CCleaner </c>
    ''' </summary>
    Public Property Flavor As WinappFlavor = WinappFlavor.CCleaner

    ''' <summary>
    ''' Creates options holding the default rules and settings: every rule on except
    ''' Optimizations, with every repair running
    ''' </summary>
    Public Sub New()

        Me.New(resetToDefaults:=True)

    End Sub

    ''' <summary>
    ''' Copies <see cref="WinappDebug.Rules"/> for this run, optionally back to their initial
    ''' states, and derives the repair flags from them
    ''' </summary>
    '''
    ''' <param name="resetToDefaults">
    ''' Indicates whether the copies go back to their initial scan and repair states
    ''' </param>
    Private Sub New(resetToDefaults As Boolean)

        Dim copies = WinappDebug.Rules.Select(Function(rule) rule.CloneFor(Me)).ToList()
        If resetToDefaults Then copies.ForEach(Sub(rule) rule.resetParams())

        Rules = copies
        Casing = copies(0)
        Alpha = copies(1)
        WrongNums = copies(2)
        Params = copies(3)
        Flags = copies(4)
        Slashes = copies(5)
        Defaults = copies(6)
        Dupes = copies(7)
        ExtraNums = copies(8)
        Multi = copies(9)
        Invalid = copies(10)
        Syntax = copies(11)
        PathValidity = copies(12)
        Semis = copies(13)
        Opti = copies(14)

        ' Every repair runs until the user turns one off. Rules still in their initial state
        ' don't count, so Optimizations, off by default, leaves RepairAll on
        RepairAll = Not copies.Any(Function(rule) rule.hasBeenChanged() AndAlso Not rule.ShouldRepair)
        RepairSome = copies.Any(Function(rule) rule.ShouldRepair)

    End Sub

    ''' <summary>
    ''' Returns options holding the user's current rule toggles and WinappDebug settings, with
    ''' the repair flags derived from the toggles
    ''' </summary>
    Public Shared Function FromSettings() As LintOptions

        Return New LintOptions(resetToDefaults:=False) With {
            .OverrideDefaultValue = overrideDefaultVal,
            .ExpectedDefaultValue = lintsettings.expectedDefaultValue,
            .PreserveDefaultKeys = lintsettings.PreserveDefaultKeys,
            .Flavor = CurrentWinappFlavor
        }

    End Function

End Class
