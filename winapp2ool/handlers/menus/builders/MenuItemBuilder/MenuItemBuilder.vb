''' <summary>
''' Fluent builder for menu items
''' </summary>
Public Class MenuItemBuilder

    ''' <summary>
    ''' 
    ''' </summary>
    Private _name As String

    ''' <summary>
    ''' 
    ''' </summary>
    Private _description As String

    ''' <summary>
    ''' 
    ''' </summary>
    Private _condition As Boolean = True

    ''' <summary>
    ''' 
    ''' </summary>
    Private _isToggle As Boolean = False

    ''' <summary>
    ''' 
    ''' </summary>
    Private _toggleState As Boolean = False

    ''' <summary>
    ''' 
    ''' </summary>
    Private _colored As Boolean = False

    ''' <summary>
    ''' 
    ''' </summary>
    Private _color As ConsoleColor

    Public Sub New(name As String, description As String)

        _name = name
        _description = description

    End Sub

    Public Function WithColor(color As ConsoleColor) As MenuItemBuilder

        _colored = True
        _color = color
        Return Me

    End Function

    Public Sub Print()

        If _isToggle Then
            MenuMaker.print(5, _name, _description, enStrCond:=_toggleState, cond:=_condition,
                           colorLine:=_colored, useArbitraryColor:=_colored, arbitraryColor:=_color)
        Else

            MenuMaker.print(1, _name, _description, cond:=_condition,
                           colorLine:=_colored, useArbitraryColor:=_colored, arbitraryColor:=_color)
        End If

    End Sub

End Class
