Option Strict On
Module MenuMaker

    'basic menu frames
    Public menuStr00 As String = " ╔════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╗"
    Public menuStr01 As String = " ║                                                                                                                    ║"
    Public menuStr02 As String = " ╚════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╝"
    Public menuStr03 As String = " ╠════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╣"
    Public menuStr04 As String = menu(menuStr01) & Environment.NewLine & mkMenuLine("Menu: Enter a number to select", "c") & Environment.NewLine & mkMenuLine(menuStr01, "")

    'constructs a menu string 
    Public Function menu(lineString As String) As String
        Return mkMenuLine(lineString, "")
    End Function

    'prints a menu string
    Public Sub printMenuLine(lineString As String)
        Console.WriteLine(menu(lineString))
    End Sub

    'constructs a menu string with alignment
    Public Function menu(lineString As String, align As String) As String
        Return mkMenuLine(lineString, align)
    End Function

    'prints a menu string with alignment
    Public Sub printMenuLine(lineString As String, align As String)
        Console.WriteLine(menu(lineString, align))
    End Sub

    'Flip the exitCode status so we can return to menu when desired
    Public Sub revertMenu(ByRef exitCode As Boolean)
        If exitCode Then
            exitCode = False
        Else
            exitCode = True
        End If
    End Sub

    'Construct a menu line properly fit to the width of the console
    Public Function mkMenuLine(line As String, align As String) As String
        If line.Length >= 119 Then
            Return line
        End If
        Dim out As String = " ║"
        Select Case align
            Case "c"
                Dim contentLength As Integer = line.Length
                Dim remainder As Integer = 118 - contentLength
                Dim lhsNum As Integer = CInt(remainder / 2)
                While out.Length < lhsNum + 2
                    out += " "
                End While
                out += line
                While out.Length < 118
                    out += " "
                End While
                out += "║"
            Case "l"
                out += " " & line
                While out.Length < 118
                    out += " "
                End While
                out += "║"
        End Select
        Return out
    End Function

    'print a box with a single message inside it
    Public Function bmenu(text As String, align As String) As String
        Dim out As String = menu(menuStr00) & Environment.NewLine
        out += menu(text, align) & Environment.NewLine
        out += menu(menuStr02)
        Return out
    End Function

    'print the top most part of the menu with no bottom
    Public Function tmenu(text As String) As String
        Dim out As String = menu(menuStr00) & Environment.NewLine
        out += menu(text, "c")
        Return out
    End Function

    'print out a middle menu box with a bottom bar
    Public Function mMenu(text As String) As String
        Dim out As String = menu(menuStr03) & Environment.NewLine
        out += menu(text, "c") & Environment.NewLine
        out += menu(menuStr03)
        Return out
    End Function

    'print out a middle menu box with no bottom bar
    Public Function moMenu(text As String) As String
        Dim out As String = menu(menuStr03) & Environment.NewLine
        out += menu(text, "c")
        Return out
    End Function

End Module
