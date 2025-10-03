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
''' MenuSection is a MenuMaker helper class that represents a "section" of a menu 
''' It allows easy grouping of related menu options, toggles, and lines of text 
''' as well as enabling the simple creation of entire module menus with a more easily understood 
''' interface than the pre-2025 MenuMaker
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
    ''' The color of the title when printed. If set, the title will be printed in this color. 
    ''' If not set, the title is printed in the default console color. 
    ''' This allows for easy setting of accent colors for modules
    ''' </summary>
    Private _titleColor As ConsoleColor? = Nothing

    ''' <summary>
    ''' Initializes a new instance of the MenuSection class with the specified 
    ''' <c> <paramref name="title"/> </c>
    ''' </summary>
    ''' 
    ''' <param name="title">
    ''' The title of the menu section. If <c> "" </c>, no title is printed
    ''' </param> 
    Public Sub New(Optional title As String = "")

        _title = title

    End Sub

    ''' <summary>
    ''' Indicates whether or not this MenuSection represents a complete menu with header
    ''' </summary>
    Private _isCompleteMenu As Boolean = False

    ''' <summary>
    ''' Indicates whether or not this MenuSection is the root menu of an application 
    ''' (as opposed to a submenu) <br />
    ''' This affects certain printing behaviors, such as the phrasing of the exit option in the menu
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
    ''' Options comprise of a name and a description, and are printed as selectable (numbered) items 
    ''' in the menu. <br />
    ''' If the <c> <paramref name="condition"/> </c> is false, the option will not be printed
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
    ''' Indicates whether or not the option should be printed <br />
    ''' Optional, Default: <c> True </c>
    ''' </param>
    Public Function AddOption(name As String,
                              description As String,
                     Optional condition As Boolean = True) As MenuSection

        _items.Add(Sub() MenuMaker.PrintOption(name, description, condition))
        Return Me

    End Function

    ''' <summary>
    ''' Creates a MenuSection that represents a complete menu with header and descriptions
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
    ''' 
    ''' </param>
    ''' 
    ''' <param name="headerColor">
    ''' The color with which the header of the complete menu should be printed to the user 
    ''' <br />
    ''' Optional, Default: <c> Nothing </c> (prints in default console color)
    ''' </param>
    ''' 
    ''' <param name="isRootMenu">
    ''' Indicates whether or not this MenuSection is the root menu of an application 
    ''' (as opposed to a submenu) <br />
    ''' This affects certain printing behaviors, such as the phrasing of the exit option in the menu <br />
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

        ' MenuMaker stores a replacement header text as a form of menu output for the user 
        ' If this exists, we should show it instead of the module name which is the typical
        ' menu header 
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
    ''' Adds an Enable/Disable toggle to the menu section <br />
    ''' Toggles are printed as an item with a name and a description, and colored based on their state  <br />
    ''' Toggles are printed <c> Green </c> when <c> <paramref name="isEnabled"/> </c>
    ''' is <c> True </c>, <c> Red </c> otherwise <br />
    ''' If the <c> <paramref name="condition"/> </c> is false, the toggle will not be printed
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
    ''' Indicates whether the setting the toggle controls is currently enabled or disabled. <br />
    ''' If <c> True </c>, the toggle is printed in green, indicating it is enabled. <br />
    ''' If <c> False </c>, the toggle is printed in red, indicating it is disabled.
    ''' </param>
    ''' 
    ''' <param name="condition">
    ''' Indicates whether or not the toggle should be printed <br />
    ''' Optional, Default: <c> True </c>
    ''' </param>
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
    ''' If the <c> <paramref name="condition"/> </c> is false, the line will not be printed
    ''' </summary>
    ''' 
    ''' <param name="text">
    ''' The text to be added to the menu section
    ''' </param>
    ''' 
    ''' <param name="centered">
    ''' Indicates whether the text should be centered in the console window. <br />
    ''' Optional, Default: <c> False </c> (text is left-aligned)
    ''' </param>
    ''' 
    ''' <param name="condition">
    ''' Indicates whether or not the line should be printed <br />
    ''' Optional, Default: <c> True </c>
    ''' </param>
    Public Function AddLine(text As String,
                   Optional centered As Boolean = False,
                   Optional condition As Boolean = True) As MenuSection

        _items.Add(Sub() MenuMaker.PrintLine(text, centered, condition))
        Return Me

    End Function

    ''' <summary>
    ''' Adds a blank line to the menu section. <br />
    ''' A blank line is simply a line with no text, used for spacing in the menu. <br />
    ''' If the <c> <paramref name="condition"/> </c> is false, the blank line will not be printed
    ''' </summary>
    ''' 
    ''' <param name="condition">
    ''' Indicates whether or not the blank line should be printed <br />
    ''' Optional, Default: <c> True </c>
    ''' </param>
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
    ''' Whether the text should be centered
    ''' </param>
    ''' 
    ''' <param name="condition">
    ''' Whether the line should be printed
    ''' </param>
    Public Function AddColoredLine(text As String,
                                  color As ConsoleColor,
                                  Optional centered As Boolean = False,
                                  Optional condition As Boolean = True) As MenuSection

        _items.Add(Sub() MenuMaker.PrintColored(text, color, centered, condition))

        Return Me

    End Function

    ''' <summary>
    ''' Adds a <c> MenuLine </c> containing information about a file (ie. its path) to the menu,
    ''' colored based on the file's existence on disk. <br />
    ''' <br /> The line will be displayed in <br />
    ''' <c> Green </c> if the file exists, <br /> <c> Red </c> otherwise <br />
    ''' <br />
    ''' If the given <c> <paramref name="condition"/> 
    ''' </c> is <c> False </c>, the file <c> MenuLine </c> will not be printed
    ''' </summary>
    ''' 
    ''' <param name="text">
    ''' A description of the file whose path is being displayed as it will appear to the user
    ''' </param>
    ''' 
    ''' <param name="filePath">
    ''' The path on disk to the file whose existence is being indicated by color
    ''' <br />
    ''' The path will be displayed in green if the file exists, red if it does not
    ''' </param>
    ''' 
    ''' <param name="condition">
    ''' The condition under which the file info line should be printed
    ''' <br />
    ''' Optional, Default: <c> True </c>
    ''' </param>
    Public Function AddFileInfo(text As String,
                                filePath As String,
                       Optional condition As Boolean = True) As MenuSection


        AddColoredLine(text & replDir(filePath), GetRedGreen(Not System.IO.File.Exists(filePath)), condition:=condition)

        Return Me

    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="text"></param>
    ''' <param name="filePath"></param>
    ''' <param name="color"></param>
    ''' <param name="condition"></param>
    ''' <returns></returns>
    Public Function AddColoredFileInfo(text As String,
                                filePath As String,
                                color As ConsoleColor,
                       Optional condition As Boolean = True) As MenuSection

        AddColoredLine(text & replDir(filePath), color, condition:=condition)
        Return Me

    End Function

    ''' <summary>
    ''' Adds a Reset Module Settings option to the current menu 
    ''' </summary>
    ''' 
    ''' <param name="moduleName">
    ''' The name of the module whose reset settings option will be created
    ''' </param>
    ''' 
    ''' <param name="condition">
    ''' Indicates whether or not the option should be printed (i.e. whether settings have been changed)
    ''' </param>
    Public Function AddResetOpt(moduleName As String,
                                condition As Boolean) As MenuSection

        If Not condition Then Return Me

        _items.Add(Sub() MenuMaker.PrintOption("Reset Settings", $"Restore {moduleName}'s settings to their default state", condition))

        Return Me

    End Function

    ''' <summary>
    ''' Adds a colored option to the menu section
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
    ''' Whether the option should be printed
    ''' </param>
    Public Function AddColoredOption(name As String,
                                   description As String,
                                   color As ConsoleColor,
                                   Optional condition As Boolean = True) As MenuSection

        If Not condition Then Return Me

        _items.Add(Sub() MenuItem(name, description, condition).WithColor(color).Print())

        Return Me

    End Function

    ''' <summary>
    ''' Adds a warning line (yellow text with /!\ markers /!\) to the menu section
    ''' </summary>
    ''' 
    ''' <param name="text">
    ''' The warning text to display
    ''' </param>
    ''' 
    ''' <param name="condition">
    ''' Whether the warning should be printed
    ''' </param>
    Public Function AddWarning(text As String,
                      Optional condition As Boolean = True) As MenuSection

        If Not condition Then Return Me

        _items.Add(Sub() MenuMaker.PrintWarning(text, condition))

        Return Me

    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="text"></param>
    ''' <param name="condition"></param>
    ''' <returns></returns>
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
    ''' If the given <c> <paramref name="condition"/> </c> is false, the border will not be printed
    ''' </summary>
    ''' 
    ''' <param name="condition">
    ''' Indicates whether or not the border should be printed <br />
    ''' Optional, Default: <c> True </c>
    ''' </param>
    Public Function AddTopBorder(Optional condition As Boolean = True) As MenuSection

        If Not condition Then Return Me

        _items.Add(Sub() MenuMaker.BeginMenu())

        Return Me

    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="condition"></param>
    ''' <returns></returns>
    Public Function AddBottomBorder(Optional condition As Boolean = True) As MenuSection

        If Not condition Then Return Me

        _items.Add(Sub() MenuMaker.EndMenu())

        Return Me

    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="condition"></param>
    ''' <returns></returns>
    Public Function AddDivider(Optional condition As Boolean = True) As MenuSection

        If Not condition Then Return Me

        _items.Add(Sub() MenuMaker.PrintDivider())

        Return Me

    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="condition"></param>
    ''' <returns></returns>
    Public Function AddNewLine(Optional condition As Boolean = True) As MenuSection

        If Not condition Then Return Me

        _items.Add(Sub() MenuMaker.PrintNewLine(condition))

        Return Me

    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="condition"></param>
    ''' <returns></returns>
    Public Function AddAnyKeyPrompt(Optional condition As Boolean = True) As MenuSection

        If Not condition Then Return Me

        AddNewLine()
        _items.Add(Sub() MenuMaker.cwl(anyKeyStr))
        AddNewLine()

        Return Me

    End Function

    ''' <summary>
    ''' Hands off printing of the MenuSection to the appropriate function based on whether or not 
    ''' it is a complete menu 
    ''' </summary>
    ''' 
    ''' <param name="withDivider">
    ''' Indicates that the section should be printed with a solid border dividing it from its options
    ''' <br /> 
    ''' Optional, Default: <c> True </c>
    ''' </param>
    Public Sub Print(Optional withDivider As Boolean = True)

        If _isCompleteMenu Then PrintCompleteMenu() : Return

        PrintSection(withDivider)

    End Sub

    ''' <summary>
    ''' Prints a menu section that isn't itself a complete menu
    ''' </summary>
    ''' 
    ''' <param name="withDivider">
    ''' Indicates that the section should be printed with a solid border dividing it from its options
    ''' <br />
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
    ''' Prints a complete menu with header, descriptions, and options beginning with an exit option
    ''' <br />
    ''' This is used for the main menus of modules that require a header and multiple options
    ''' </summary>
    Private Sub PrintCompleteMenu()

        BeginMenu()

        If Not String.IsNullOrEmpty(_menuHeader) Then

            Dim hasColor = _menuHeaderColor.HasValue

            MenuMaker.PrintColored(_menuHeader, _menuHeaderColor.Value, centered:=True, hasColor)

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
