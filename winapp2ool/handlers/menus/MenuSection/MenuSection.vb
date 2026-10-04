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
''' A section of a MenuMaker menu: a list of options, toggles and lines of text that are
''' printed together, or a complete module menu with a header. Items are stored as print
''' actions and nothing reaches the console until <see cref="Print"/>.
''' </summary>
Public Class MenuSection

    ''' <summary>
    ''' The title of the menu section. This is displayed above the set of items in this section
    ''' when printing. When it is blank, no title is printed
    ''' </summary>
    Private _title As String

    ''' <summary>
    ''' The set of menu items in this section. Each item is an Action delegate that
    ''' represents a menu option, toggle, or line of text to be printed to the user
    ''' </summary>
    Private _items As New List(Of Action)

    ''' <summary>
    ''' Indicates whether this section has no items, so a caller assembling sections can
    ''' skip it rather than emitting a divider or a border around empty space. A title alone
    ''' doesn't count as an item. <see cref="AddOption"/>, <see cref="AddToggle"/>,
    ''' <see cref="AddLine"/>, <see cref="AddBlank"/>, <see cref="AddColoredLine"/>,
    ''' <see cref="AddFileInfo"/> and <see cref="AddColoredFileInfo"/> keep their item when
    ''' its condition is <c> False </c>, so a section of those isn't empty even when it prints nothing.
    ''' </summary>
    Public ReadOnly Property IsEmpty As Boolean
        Get
            Return _items.Count = 0
        End Get
    End Property

    ''' <summary>
    ''' Registered handlers for dispatching user input, in the order they were added
    ''' (index 0 = option 1, the first selectable item after Exit). Only the
    ''' <c> AddDispatched* </c> methods register here, so the indexes match the printed numbers
    ''' only when every numbered option in the menu was added through one of them.
    ''' </summary>
    Private _actions As New List(Of Action)

    ''' <summary>
    ''' The color of the title when printed, or <c> Nothing </c> for the default console color.
    ''' Nothing in the class sets it, so titles always print in the default color.
    ''' </summary>
    Private _titleColor As ConsoleColor? = Nothing

    ''' <summary>
    ''' Creates a new <c> MenuSection </c> with the given <paramref name="title"/>
    ''' </summary>
    '''
    ''' <param name="title">
    ''' The title of the menu section. If <c> "" </c>, no title is printed <br /><br />
    ''' Optional, Default: <c> "" </c>
    ''' </param>
    Public Sub New(Optional title As String = "")

        _title = title

    End Sub

    ''' <summary>
    ''' Indicates whether this MenuSection represents a complete menu with header
    ''' </summary>
    Private _isCompleteMenu As Boolean = False

    ''' <summary>
    ''' Indicates whether this MenuSection is the root menu of the application rather than a
    ''' submenu, which changes the wording of the Exit option
    ''' </summary>
    Private _isRootMenu As Boolean = False

    ''' <summary>
    ''' The main header text for a complete menu
    ''' </summary>
    Private _menuHeader As String = ""

    ''' <summary>
    ''' The color for the main menu header
    ''' </summary>
    Private _menuHeaderColor As ConsoleColor? = Nothing

    ''' <summary>
    ''' Description lines that appear under the header
    ''' </summary>
    Private _descriptionLines As New List(Of String)

    ''' <summary>
    ''' Adds an option to the menu section. <br />
    ''' Options consist of a name and a description, and are printed as selectable (numbered) items
    ''' in the menu. No handler is registered, so <see cref="Dispatch"/> can't reach it. <br />
    ''' If <paramref name="condition"/> is <c> False </c>, the option is not printed and doesn't
    ''' use up a number
    ''' </summary>
    ''' 
    ''' <param name="name">
    ''' The name of the menu option, appears to the left of the description
    ''' </param>
    ''' 
    ''' <param name="description">
    ''' The description of the menu option, appears to the right of the name
    ''' </param>
    ''' 
    ''' <param name="condition">
    ''' Indicates whether the option should be printed <br /><br />
    ''' Optional, Default: <c> True </c>
    ''' </param>
    '''
    ''' <returns>This section, for chaining</returns>
    Public Function AddOption(name As String,
                              description As String,
                     Optional condition As Boolean = True) As MenuSection

        _items.Add(Sub() MenuMaker.PrintOption(name, description, condition))
        Return Me

    End Function

    ''' <summary>
    ''' Builds a MenuSection that represents a complete menu with header and descriptions.
    ''' If <see cref="MenuMaker.MenuHeaderText"/> holds a pending header, we use it and its color
    ''' in place of <paramref name="menuHeader"/> and <paramref name="headerColor"/>, then clear it.
    ''' </summary>
    '''
    ''' <param name="menuHeader">
    ''' The text to appear in the header of the menu. The header
    ''' visually approximates a window title bar within the MenuMaker system
    ''' </param>
    '''
    ''' <param name="descriptionLines">
    ''' A set of text lines that describe the menu's purpose, printed
    ''' under the header but before any menu items
    ''' </param>
    '''
    ''' <param name="headerColor">
    ''' The color with which the header of the complete menu should be printed to the user <br /><br />
    ''' Optional, Default: <c> Nothing </c> <br />
    ''' prints in the default console color
    ''' </param>
    '''
    ''' <param name="isRootMenu">
    ''' Indicates whether this MenuSection is the root menu of the application rather than a
    ''' submenu. When <c> True </c>, the Exit option reads "Exit the application" instead of
    ''' "Return to the previous menu". <br /><br />
    ''' Optional, Default: <c> False </c>
    ''' </param>
    ''' 
    ''' <returns>
    ''' A MenuSection configured as a complete menu with the provided header and descriptions, 
    ''' ready to be populated with options, toggles, and other menu items
    ''' </returns>
    Public Shared Function CreateCompleteMenu(menuHeader As String,
                                              descriptionLines As String(),
                                              Optional headerColor As ConsoleColor? = Nothing,
                                              Optional isRootMenu As Boolean = False) As MenuSection

        Dim section As New MenuSection With {
            ._isCompleteMenu = True,
            ._isRootMenu = isRootMenu,
            ._menuHeader = If(MenuHeaderText = "", menuHeader, MenuHeaderText),
            ._menuHeaderColor = If(MenuHeaderText = "", headerColor, MenuHeaderTextColor)
        }

        MenuHeaderText = ""

        section._descriptionLines.AddRange(descriptionLines)

        Return section

    End Function

    ''' <summary>
    ''' Adds an Enable/Disable toggle to the menu section as a numbered option. We print the
    ''' name as <c> "Toggle {name}" </c> and start the description with <c> "Disable" </c> when
    ''' <paramref name="isEnabled"/> is <c> True </c> and <c> "Enable" </c> otherwise. <br />
    ''' Toggles are printed <c> Green </c> when <paramref name="isEnabled"/>
    ''' is <c> True </c>, <c> Red </c> otherwise. No handler is registered. <br />
    ''' If <paramref name="condition"/> is <c> False </c>, the toggle is not printed and doesn't
    ''' use up a number
    ''' </summary>
    '''
    ''' <param name="name">
    ''' The name of the toggle, appears to the left of the description
    ''' </param>
    '''
    ''' <param name="description">
    ''' The description of the toggle, appears to the right of the name
    ''' </param>
    '''
    ''' <param name="isEnabled">
    ''' Indicates whether the setting the toggle controls is currently enabled
    ''' </param>
    '''
    ''' <param name="condition">
    ''' Indicates whether the toggle should be printed <br /><br />
    ''' Optional, Default: <c> True </c>
    ''' </param>
    '''
    ''' <returns>This section, for chaining</returns>
    Public Function AddToggle(name As String,
                              description As String,
                              isEnabled As Boolean,
                              Optional condition As Boolean = True) As MenuSection

        _items.Add(Sub() MenuMaker.PrintToggle("Toggle " & name, description, isEnabled, condition))
        Return Me

    End Function

    ''' <summary>
    ''' Adds a line of text to the menu section. <br />
    ''' The line is printed as a simple text line, optionally centered. <br />
    ''' If <paramref name="condition"/> is <c> False </c>, the line will not be printed
    ''' </summary>
    '''
    ''' <param name="text">
    ''' The text to be added to the menu section
    ''' </param>
    '''
    ''' <param name="centered">
    ''' Indicates whether the text should be centered in the console window <br /><br />
    ''' Optional, Default: <c> False </c> <br />
    ''' text is left-aligned
    ''' </param>
    '''
    ''' <param name="condition">
    ''' Indicates whether the line should be printed <br /><br />
    ''' Optional, Default: <c> True </c>
    ''' </param>
    '''
    ''' <returns>This section, for chaining</returns>
    Public Function AddLine(text As String,
                   Optional centered As Boolean = False,
                   Optional condition As Boolean = True) As MenuSection

        _items.Add(Sub() MenuMaker.PrintLine(text, centered, condition))
        Return Me

    End Function

    ''' <summary>
    ''' Adds a blank framed line to the menu section, for spacing. <br />
    ''' If <paramref name="condition"/> is <c> False </c>, the blank line will not be printed
    ''' </summary>
    '''
    ''' <param name="condition">
    ''' Indicates whether the blank line should be printed <br /><br />
    ''' Optional, Default: <c> True </c>
    ''' </param>
    '''
    ''' <returns>This section, for chaining</returns>
    Public Function AddBlank(Optional condition As Boolean = True) As MenuSection

        _items.Add(Sub() MenuMaker.PrintBlank(condition))
        Return Me

    End Function

    ''' <summary>
    ''' Adds a colored line of text to the menu section
    ''' </summary>
    ''' 
    ''' <param name="text">
    ''' The text to be displayed
    ''' </param>
    ''' 
    ''' <param name="color">
    ''' The color to display the text in
    ''' </param>
    ''' 
    ''' <param name="centered">
    ''' Indicates whether the text should be centered <br /><br />
    ''' Optional, Default: <c> False </c>
    ''' </param>
    '''
    ''' <param name="condition">
    ''' Indicates whether the line should be printed <br /><br />
    ''' Optional, Default: <c> True </c>
    ''' </param>
    '''
    ''' <returns>This section, for chaining</returns>
    Public Function AddColoredLine(text As String,
                                  color As ConsoleColor,
                                  Optional centered As Boolean = False,
                                  Optional condition As Boolean = True) As MenuSection

        _items.Add(Sub() MenuMaker.PrintColored(text, color, centered, condition))

        Return Me

    End Function

    ''' <summary>
    ''' Adds a line showing <paramref name="text"/> followed by <paramref name="filePath"/>, with the
    ''' current directory shortened to <c> .. </c>. The line is <c> Green </c> if the file exists
    ''' and <c> Red </c> otherwise. We check the file when the line is added, not when it prints.
    ''' <br /><br />
    ''' If <paramref name="condition"/> is <c> False </c>, the line will not be printed
    ''' </summary>
    '''
    ''' <param name="text">
    ''' A label printed directly before the path, with no separator added
    ''' </param>
    '''
    ''' <param name="filePath">
    ''' The path on disk to the file whose existence is being indicated by color
    ''' </param>
    '''
    ''' <param name="condition">
    ''' Indicates whether the line should be printed <br /><br />
    ''' Optional, Default: <c> True </c>
    ''' </param>
    '''
    ''' <returns>This section, for chaining</returns>
    Public Function AddFileInfo(text As String,
                                filePath As String,
                       Optional condition As Boolean = True) As MenuSection


        AddColoredLine(text & replDir(filePath), GetRedGreen(Not System.IO.File.Exists(filePath)), condition:=condition)

        Return Me

    End Function

    ''' <summary>
    ''' Adds a line showing <paramref name="text"/> followed by <paramref name="filePath"/>, with the
    ''' current directory shortened to <c> .. </c>, in a fixed color. Unlike
    ''' <see cref="AddFileInfo"/>, the color doesn't depend on whether the file exists.
    ''' </summary>
    '''
    ''' <param name="text">A label printed directly before the path, with no separator added</param>
    '''
    ''' <param name="filePath">The path to show</param>
    '''
    ''' <param name="color">The color of the line</param>
    '''
    ''' <param name="condition">
    ''' Indicates whether the line should be printed <br /><br />
    ''' Optional, Default: <c> True </c>
    ''' </param>
    '''
    ''' <returns>This section, for chaining</returns>
    Public Function AddColoredFileInfo(text As String,
                                filePath As String,
                                color As ConsoleColor,
                       Optional condition As Boolean = True) As MenuSection

        AddColoredLine(text & replDir(filePath), color, condition:=condition)
        Return Me

    End Function

    ''' <summary>
    ''' Adds a numbered <c> Reset Settings </c> option to the current menu. No handler is
    ''' registered. When <paramref name="condition"/> is <c> False </c>, nothing is added.
    ''' </summary>
    '''
    ''' <param name="moduleName">
    ''' The name of the module whose reset settings option will be created
    ''' </param>
    '''
    ''' <param name="condition">
    ''' Indicates whether the option should be added, usually whether the settings have been changed
    ''' </param>
    '''
    ''' <returns>This section, for chaining</returns>
    Public Function AddResetOpt(moduleName As String,
                                condition As Boolean) As MenuSection

        If Not condition Then Return Me

        _items.Add(Sub() MenuMaker.PrintOption("Reset Settings", $"Restore {moduleName}'s settings to their default state", condition))

        Return Me

    End Function

    ''' <summary>
    ''' Adds a colored numbered option to the menu section. No handler is registered.
    ''' When <paramref name="condition"/> is <c> False </c>, nothing is added.
    ''' </summary>
    '''
    ''' <param name="name">
    ''' The name of the menu option (left text)
    ''' </param>
    '''
    ''' <param name="description">
    ''' A description of the menu option's function (right text)
    ''' </param>
    '''
    ''' <param name="color">
    ''' The color with which the option should be printed to the user
    ''' </param>
    '''
    ''' <param name="condition">
    ''' Indicates whether the option should be added <br /><br />
    ''' Optional, Default: <c> True </c>
    ''' </param>
    '''
    ''' <returns>This section, for chaining</returns>
    Public Function AddColoredOption(name As String,
                                   description As String,
                                   color As ConsoleColor,
                                   Optional condition As Boolean = True) As MenuSection

        If Not condition Then Return Me

        _items.Add(Sub() MenuMaker.PrintColoredOption(name, description, color))

        Return Me

    End Function

    ''' <summary>
    ''' Adds a warning line, printed as left-aligned yellow text, to the menu section.
    ''' When <paramref name="condition"/> is <c> False </c>, nothing is added.
    ''' </summary>
    '''
    ''' <param name="text">
    ''' The warning text to display
    ''' </param>
    '''
    ''' <param name="condition">
    ''' Indicates whether the warning should be added <br /><br />
    ''' Optional, Default: <c> True </c>
    ''' </param>
    '''
    ''' <returns>This section, for chaining</returns>
    Public Function AddWarning(text As String,
                      Optional condition As Boolean = True) As MenuSection

        If Not condition Then Return Me

        _items.Add(Sub() MenuMaker.PrintWarning(text, condition))

        Return Me

    End Function

    ''' <summary>
    ''' Adds a box of three lines: a top border, <paramref name="text"/> centered, and a bottom
    ''' border. When <paramref name="condition"/> is <c> False </c>, nothing is added.
    ''' </summary>
    '''
    ''' <param name="text">The text inside the box</param>
    '''
    ''' <param name="condition">
    ''' Indicates whether the box should be added <br /><br />
    ''' Optional, Default: <c> True </c>
    ''' </param>
    '''
    ''' <returns>This section, for chaining</returns>
    Public Function AddBoxWithText(text As String,
                          Optional condition As Boolean = True) As MenuSection

        If Not condition Then Return Me

        Me.AddTopBorder()
        Me.AddLine(text, centered:=True, condition)
        Me.AddBottomBorder()

        Return Me

    End Function

    ''' <summary>
    ''' Adds a top border to the menu section. <br />
    ''' Visually separates content within a section into a new box <br />
    ''' If <paramref name="condition"/> is <c> False </c>, nothing is added
    ''' </summary>
    '''
    ''' <param name="condition">
    ''' Indicates whether the border should be added <br /><br />
    ''' Optional, Default: <c> True </c>
    ''' </param>
    '''
    ''' <returns>This section, for chaining</returns>
    Public Function AddTopBorder(Optional condition As Boolean = True) As MenuSection

        If Not condition Then Return Me

        _items.Add(Sub() MenuMaker.BeginMenu())

        Return Me

    End Function

    ''' <summary>
    ''' Adds a bottom border that closes the current box. If <paramref name="condition"/> is
    ''' <c> False </c>, nothing is added.
    ''' </summary>
    '''
    ''' <param name="condition">
    ''' Indicates whether the border should be added <br /><br />
    ''' Optional, Default: <c> True </c>
    ''' </param>
    '''
    ''' <returns>This section, for chaining</returns>
    Public Function AddBottomBorder(Optional condition As Boolean = True) As MenuSection

        If Not condition Then Return Me

        _items.Add(Sub() MenuMaker.EndMenu())

        Return Me

    End Function

    ''' <summary>
    ''' Adds a T-frame conjoiner divider to the menu section. <br />
    ''' A conjoiner uses frame type 3 (<c> ╠ ╣ </c>), by default filled with <c> ═ </c>.
    ''' If <paramref name="condition"/> is <c> False </c>, nothing is added.
    ''' </summary>
    '''
    ''' <param name="condition">
    ''' Indicates whether the divider should be added <br /><br />
    ''' Optional, Default: <c> True </c>
    ''' </param>
    '''
    ''' <param name="solid">
    ''' Indicates whether the divider should be filled with <c> ═ </c> rather than left empty <br /><br />
    ''' Optional, Default: <c> True </c>
    ''' </param>
    '''
    ''' <returns>This section, for chaining</returns>
    Public Function AddDivider(Optional condition As Boolean = True,
                               Optional solid As Boolean = True) As MenuSection

        If Not condition Then Return Me

        _items.Add(Sub() MenuMaker.PrintDivider(solid))

        Return Me

    End Function

    ''' <summary>
    ''' Adds an unframed empty line, written by <see cref="MenuMaker.PrintNewLine"/>. That skips the
    ''' render buffer, so during a buffered <see cref="Print"/> the line reaches the console ahead
    ''' of the lines buffered before it. If <paramref name="condition"/> is <c> False </c>, nothing
    ''' is added.
    ''' </summary>
    '''
    ''' <param name="condition">
    ''' Indicates whether the line should be added <br /><br />
    ''' Optional, Default: <c> True </c>
    ''' </param>
    '''
    ''' <returns>This section, for chaining</returns>
    Public Function AddNewLine(Optional condition As Boolean = True) As MenuSection

        If Not condition Then Return Me

        _items.Add(Sub() MenuMaker.PrintNewLine(condition))

        Return Me

    End Function

    ''' <summary>
    ''' Adds the unframed line <c> "Press any key to return to the menu." </c> between two
    ''' <see cref="AddNewLine"/> lines. During a buffered <c> Print </c>, both blank lines skip
    ''' the buffer and reach the console ahead of the prompt. It only prints the prompt: waiting
    ''' for the key is up to the caller. If <paramref name="condition"/> is <c> False </c>, nothing is added.
    ''' </summary>
    '''
    ''' <param name="condition">
    ''' Indicates whether the prompt should be added <br /><br />
    ''' Optional, Default: <c> True </c>
    ''' </param>
    '''
    ''' <returns>This section, for chaining</returns>
    Public Function AddAnyKeyPrompt(Optional condition As Boolean = True) As MenuSection

        If Not condition Then Return Me

        AddNewLine()
        _items.Add(Sub() MenuMaker.cwl(anyKeyStr))
        AddNewLine()

        Return Me

    End Function

    ''' <summary>
    ''' Adds a numbered option, registers a dispatch handler, and returns <c> Me </c> for chaining. <br />
    ''' When <paramref name="condition"/> is <c> False </c>, neither the option nor the
    ''' handler is registered, and the option number sequence is unaffected.
    ''' </summary>
    '''
    ''' <param name="name">
    ''' The name of the option as it appears on the menu
    ''' </param>
    '''
    ''' <param name="description">
    ''' A description of what the option does
    ''' </param>
    '''
    ''' <param name="handler">
    ''' The action to invoke when the user selects this option
    ''' </param>
    '''
    ''' <param name="condition">
    ''' Indicates whether the option should be shown and dispatched <br /><br />
    ''' Optional, Default: <c> True </c>
    ''' </param>
    '''
    ''' <returns>This section, for chaining</returns>
    Public Function AddDispatchedOption(name As String,
                                        description As String,
                                        handler As Action,
                               Optional condition As Boolean = True) As MenuSection

        If Not condition Then Return Me

        AddOption(name, description)
        _actions.Add(handler)

        Return Me

    End Function

    ''' <summary>
    ''' Adds a toggle option, registers a dispatch handler, and returns <c> Me </c> for chaining. <br />
    ''' When <paramref name="condition"/> is <c> False </c>, neither the toggle nor the
    ''' handler is registered.
    ''' </summary>
    '''
    ''' <param name="name">
    ''' The name of the toggle as it appears on the menu
    ''' </param>
    '''
    ''' <param name="description">
    ''' A description of what the toggle controls
    ''' </param>
    '''
    ''' <param name="isEnabled">
    ''' Indicates whether the setting this toggle controls is enabled
    ''' </param>
    '''
    ''' <param name="handler">
    ''' The action to invoke when the user selects this toggle
    ''' </param>
    '''
    ''' <param name="condition">
    ''' Indicates whether the toggle should be shown and dispatched <br /><br />
    ''' Optional, Default: <c> True </c>
    ''' </param>
    '''
    ''' <returns>This section, for chaining</returns>
    Public Function AddDispatchedToggle(name As String,
                                        description As String,
                                        isEnabled As Boolean,
                                        handler As Action,
                               Optional condition As Boolean = True) As MenuSection

        If Not condition Then Return Me

        AddToggle(name, description, isEnabled)
        _actions.Add(handler)

        Return Me

    End Function

    ''' <summary>
    ''' Adds a colored option, registers a dispatch handler, and returns <c> Me </c> for chaining. <br />
    ''' When <paramref name="condition"/> is <c> False </c>, neither the option nor the
    ''' handler is registered.
    ''' </summary>
    '''
    ''' <param name="name">
    ''' The name of the option as it appears on the menu
    ''' </param>
    '''
    ''' <param name="description">
    ''' A description of what the option does
    ''' </param>
    '''
    ''' <param name="color">
    ''' The color to print this option in
    ''' </param>
    '''
    ''' <param name="handler">
    ''' The action to invoke when the user selects this option
    ''' </param>
    '''
    ''' <param name="condition">
    ''' Indicates whether the option should be shown and dispatched <br /><br />
    ''' Optional, Default: <c> True </c>
    ''' </param>
    '''
    ''' <returns>This section, for chaining</returns>
    Public Function AddDispatchedColoredOption(name As String,
                                               description As String,
                                               color As ConsoleColor,
                                               handler As Action,
                                      Optional condition As Boolean = True) As MenuSection

        If Not condition Then Return Me

        AddColoredOption(name, description, color)
        _actions.Add(handler)

        Return Me

    End Function

    ''' <summary>
    ''' Adds a Reset Settings option, registers a dispatch handler, and returns <c> Me </c> for chaining. <br />
    ''' When <paramref name="condition"/> is <c> False </c>, the option is not shown and
    ''' no handler is registered.
    ''' </summary>
    '''
    ''' <param name="moduleName">
    ''' The name of the module whose settings will be reset
    ''' </param>
    '''
    ''' <param name="condition">
    ''' Indicates whether the option should be shown, usually whether the settings have been changed
    ''' </param>
    '''
    ''' <param name="handler">
    ''' The action to invoke when the user selects this option
    ''' </param>
    '''
    ''' <returns>This section, for chaining</returns>
    Public Function AddDispatchedResetOpt(moduleName As String,
                                          condition As Boolean,
                                          handler As Action) As MenuSection
        If Not condition Then Return Me

        AddResetOpt(moduleName, True)
        _actions.Add(handler)

        Return Me

    End Function

    ''' <summary>
    ''' Dispatches the user's integer input to the registered handler for the selected option. <br />
    ''' Option 0 (Exit) is NOT dispatched, the caller handles it. <br />
    ''' Options 1..N map to handlers registered via <c> AddDispatched* </c> calls in order, which
    ''' matches the printed numbers only if no numbered option was added without a handler.
    ''' </summary>
    '''
    ''' <param name="intInput">
    ''' The integer the user entered at the menu prompt
    ''' </param>
    '''
    ''' <returns>
    ''' <c> True </c> if a handler was found and invoked; <c> False </c> if the input
    ''' was out of range (caller should show an invalid-input error)
    ''' </returns>
    Public Function Dispatch(intInput As Integer) As Boolean

        Dim idx = intInput - 1

        If idx < 0 OrElse idx >= _actions.Count Then Return False

        _actions(idx)()
        Return True

    End Function

    ''' <summary>
    ''' Prints the section as one buffered render pass, as a complete menu or as a plain section
    ''' </summary>
    '''
    ''' <param name="withDivider">
    ''' Indicates whether a plain section's title is followed by a solid divider. Complete menus
    ''' and untitled sections ignore it. <br /><br />
    ''' Optional, Default: <c> True </c>
    ''' </param>
    Public Sub Print(Optional withDivider As Boolean = True)

        MenuMaker.BeginBuffered()

        Try

            If _isCompleteMenu Then PrintCompleteMenu() : Return

            PrintSection(withDivider)

        Finally

            MenuMaker.FlushBuffered()

        End Try

    End Sub

    ''' <summary>
    ''' Prints a menu section that isn't itself a complete menu: its centered title, if it has one,
    ''' and then its items. Option numbering carries on from whatever printed before.
    ''' </summary>
    '''
    ''' <param name="withDivider">
    ''' Indicates whether to print a solid divider after the title
    ''' </param>
    Private Sub PrintSection(withDivider As Boolean)

        Dim hasTitle = Not String.IsNullOrEmpty(_title)
        Dim hasColor = _titleColor.HasValue

        If hasTitle Then

            MenuMaker.PrintColored(_title, _titleColor.Value, centered:=True, hasColor)

            MenuMaker.PrintLine(_title, centered:=True, Not hasColor)

            If withDivider Then MenuMaker.PrintDivider()

        End If

        For Each item In _items

            item()

        Next

    End Sub

    ''' <summary>
    ''' Prints a complete menu in one box: header, centered description lines, the menu prompt,
    ''' then the items. We reset option numbering first, so Exit is option <c> 0 </c>.
    ''' </summary>
    Private Sub PrintCompleteMenu()

        BeginMenu()

        If Not String.IsNullOrEmpty(_menuHeader) Then

            Dim hasColor = _menuHeaderColor.HasValue
            Dim headerColor = If(hasColor, CType(_menuHeaderColor.Value, ConsoleColor), ConsoleColor.White)

            MenuMaker.PrintColored(_menuHeader, headerColor, centered:=True, hasColor)

            MenuMaker.PrintLine(_menuHeader, centered:=True, Not hasColor)

        End If

        MenuMaker.PrintDivider()

        For Each description In _descriptionLines

            MenuMaker.PrintLine(description, centered:=True)

        Next

        Dim exitText = If(_isRootMenu, "Exit the application", "Return to the previous menu")

        PrintBlank()
        PrintLine("Menu: Enter a number to select", centered:=True)
        PrintBlank()
        resetMenuNumbering()
        PrintOption("Exit", exitText)

        For Each item In _items

            item()

        Next

        EndMenu()

    End Sub

End Class
