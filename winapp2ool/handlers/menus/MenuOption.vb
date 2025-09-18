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
''' Represents a numbered menu option with its display text and associated action
''' </summary>
Public Class MenuOption

    ''' <summary>
    ''' The menu number assigned to this option
    ''' </summary>
    Public ReadOnly Property Number As Integer

    ''' <summary>
    ''' The display text for this menu option
    ''' </summary>
    Public ReadOnly Property DisplayText As String

    ''' <summary>
    ''' The action to execute when this option is selected
    ''' </summary>
    Public ReadOnly Property Action As Action

    ''' <summary>
    ''' Whether this option should be displayed to the user
    ''' </summary>
    Public ReadOnly Property IsVisible As Boolean

    ''' <summary>
    ''' Optional description text for the option
    ''' </summary>
    Public ReadOnly Property Description As String

    ''' <summary>
    ''' Initializes a new MenuOption
    ''' </summary>
    ''' <param name="number">The menu number for this option</param>
    ''' <param name="displayText">The text to display for this option</param>
    ''' <param name="action">The action to execute when selected</param>
    ''' <param name="isVisible">Whether this option should be visible</param>
    ''' <param name="description">Optional description text</param>
    Public Sub New(number As Integer,
                   displayText As String,
                   action As Action,
                   Optional isVisible As Boolean = True,
                   Optional description As String = "")
        Me.Number = number
        Me.DisplayText = displayText
        Me.Action = action
        Me.IsVisible = isVisible
        Me.Description = description
    End Sub

End Class

''' <summary>
''' Helper class for building and managing collections of menu options
''' </summary>
Public Class MenuOptionsBuilder

    Private _options As New List(Of MenuOption)
    Private _currentNumber As Integer = 2  ' Start after Exit(0) and Run(1)

    ''' <summary>
    ''' Adds a toggle option to the menu
    ''' </summary>
    Public Function AddToggle(displayText As String,
                             settingName As String,
                             moduleName As String,
                             settingsType As Type,
                             settingsChangedName As String,
                             Optional condition As Boolean = True) As MenuOptionsBuilder

        If Not condition Then Return Me

        Dim action As Action = Sub() toggleModuleSetting(displayText, moduleName, settingsType, settingName, settingsChangedName)
        _options.Add(New MenuOption(_currentNumber, displayText, action, condition))
        _currentNumber += 1
        Return Me
    End Function

    ''' <summary>
    ''' Adds a file selector option to the menu
    ''' </summary>
    Public Function AddFileSelector(displayText As String,
                                   fileObj As iniFile,
                                   settingsChanged As Boolean,
                                   moduleName As String,
                                   settingName As String,
                                   settingChangedName As String,
                                   Optional condition As Boolean = True) As MenuOptionsBuilder

        If Not condition Then Return Me

        Dim action As Action = Sub() changeFileParams(fileObj, settingsChanged, moduleName, settingName, settingChangedName)
        _options.Add(New MenuOption(_currentNumber, displayText, action, condition))
        _currentNumber += 1
        Return Me
    End Function

    ''' <summary>
    ''' Adds a custom action option to the menu
    ''' </summary>
    Public Function AddAction(displayText As String,
                             action As Action,
                             Optional condition As Boolean = True,
                             Optional description As String = "") As MenuOptionsBuilder

        If Not condition Then Return Me

        _options.Add(New MenuOption(_currentNumber, displayText, action, condition, description))
        _currentNumber += 1
        Return Me
    End Function

    ''' <summary>
    ''' Adds a reset settings option (typically the last option)
    ''' </summary>
    Public Function AddResetOption(moduleName As String,
                                  resetAction As Action,
                                  Optional condition As Boolean = True) As MenuOptionsBuilder

        If Not condition Then Return Me

        _options.Add(New MenuOption(_currentNumber, "Reset Settings", resetAction, condition,
                                   $"Restore {moduleName}'s settings to their default state"))
        Return Me
    End Function

    ''' <summary>
    ''' Gets the built list of menu options
    ''' </summary>
    Public Function Build() As List(Of MenuOption)
        Return _options.Where(Function(opt) opt.IsVisible).ToList()
    End Function

    ''' <summary>
    ''' Creates a lookup dictionary for input handling
    ''' </summary>
    Public Function BuildInputLookup() As Dictionary(Of String, Action)
        Return Build().ToDictionary(Function(opt) opt.Number.ToString(), Function(opt) opt.Action)
    End Function

End Class