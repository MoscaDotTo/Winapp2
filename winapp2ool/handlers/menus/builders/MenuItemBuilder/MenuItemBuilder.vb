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
''' MenuItemBuilder is a MenuMaker helper class that represents an individual option on the menu <br />
''' Can be a regular option or specifically a toggle <br />
''' Options are menu items presented in a #. Name - Description format and are generally used to 
''' access features within a module's settings or to access modules themselves
''' <br />
''' Toggles are a type of menu option that include specific handling for options which can
''' be enabled or disabled. This is provided to simplify creation of Options for managing Booleans
''' </summary>
''' 
''' Docs last updated: 2025-08-11 | Code last updated: 2025-08-11
Public Class MenuItemBuilder

    ''' <summary>
    ''' The name of a menu option or toggle as it will be displayed to the user <br />
    ''' This is the text on the left side of a menu line, before the description
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-08-11 | Code last updated: 2025-08-11
    Private _name As String

    ''' <summary>
    ''' The description of the menu option or toggle as it will be displayed to the user <br />
    ''' This is the text on the right side of a menu line, following the name 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-08-11 | Code last updated: 2025-08-11
    Private _description As String

    ''' <summary>
    ''' Indicates whether or not this particular menu item should be printed to the user 
    ''' <br />
    ''' Default: <c> True </c>
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-08-11 | Code last updated: 2025-08-11
    Private _condition As Boolean = True

    ''' <summary>
    ''' Indicates whether or not this particular menu item should be printed as an enable/disable 
    ''' toggle for a particular setting. 
    ''' <br />
    ''' Default: <c> False </c>
    ''' </summary>
    ''' <remarks> 
    ''' Toggles differ from regular options in that they contain an 
    ''' Enable/Disable string as a hard coded feature 
    ''' </remarks>
    ''' 
    ''' Docs last updated: 2025-08-11 | Code last updated: 2025-08-11
    Private _isToggle As Boolean = False

    ''' <summary>
    ''' Indicates the Enabled/Disabled state of a toggle. <br />
    ''' Typically, this is provided in the form of the setting being toggled 
    ''' <br />
    ''' Default: <c> False </c>
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-08-11 | Code last updated: 2025-08-11
    Private _toggleState As Boolean = False

    ''' <summary>
    ''' Indicates whether or not the menu item should be printed in color 
    ''' <br />
    ''' Default: <c> False </c>
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-08-11 | Code last updated: 2025-08-11
    Private _colored As Boolean = False

    ''' <summary>
    ''' The color which will be used to print when printing with color 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-08-11 | Code last updated: 2025-08-11
    Private _color As ConsoleColor

    ''' <summary>
    ''' 
    ''' </summary>
    ''' Creates a new MenuITemBuilder with the provided <c> ><paramref name="name"/> </c> and 
    ''' <c> <paramref name="description"/> </c>
    ''' 
    ''' <param name="name">
    ''' The name of the menu option or toggle as it will appear to the user
    ''' </param>
    ''' 
    ''' <param name="description">
    ''' A description of the menu option or toggle as it will appear to the user
    ''' </param>
    ''' 
    ''' Docs last updated: 2025-08-11 | Code last updated: 2025-08-11
    Public Sub New(name As String,
                   description As String,
          Optional cond As Boolean = True)

        _name = name
        _description = description
        _condition = cond

    End Sub

    ''' <summary>
    ''' Enables printing in color for the menu item and sets the color with which to print
    ''' </summary>
    ''' 
    ''' <param name="color">
    ''' The color with which the menu item will be printed to the user 
    ''' </param>
    ''' 
    ''' <returns> 
    ''' An updates menu item with coloring enabled and configured 
    ''' </returns>
    ''' 
    ''' Docs last updated: 2025-08-11 | Code last updated: 2025-08-11
    Public Function WithColor(color As ConsoleColor) As MenuItemBuilder

        _colored = True
        _color = color
        Return Me

    End Function

    ''' <summary>
    ''' Prints the menu item to the user 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-08-11 | Code last updated: 2025-08-11
    Public Sub Print()

        ' Toggles
        MenuMaker.print(5, _name, _description, enStrCond:=_toggleState, cond:=_condition AndAlso _isToggle,
                        colorLine:=_colored, useArbitraryColor:=_colored, arbitraryColor:=_color)

        ' Options
        MenuMaker.print(1, _name, _description, cond:=_condition AndAlso Not _isToggle,
                        colorLine:=_colored, useArbitraryColor:=_colored, arbitraryColor:=_color)

    End Sub

End Class
