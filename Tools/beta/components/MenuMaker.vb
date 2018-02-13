Option Strict On
Module MenuMaker

    Public menuStr00 As String = " ╔════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╗"
    Public menuStr01 As String = " ║                                                                                                                    ║"
    Public menuStr02 As String = " ╚════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╝"
    Public menuStr03 As String = " ╠════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╣"
    Public menuStr04 As String = mkMenuLine(menuStr01, "") & Environment.NewLine & mkMenuLine("Menu: Enter a number to select", "c") & Environment.NewLine & mkMenuLine(menuStr01, "")

    Public Sub menu(lineString As String)
        Console.WriteLine(mkMenuLine(lineString, ""))
    End Sub

    Public Sub menu(lineString As String, align As String)
        Console.WriteLine(mkMenuLine(lineString, align))
    End Sub

    Public Sub revertMenu(ByRef exitCode As Boolean)
        If exitCode Then
            exitCode = False
        Else
            exitCode = True
        End If
    End Sub

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
    Public Sub bmenu(text As String, align As String)
        menu(menuStr00)
        menu(text, align)
        menu(menuStr02)
    End Sub

    'print the top most part of the menu with no bottom
    Public Sub tmenu(text As String)
        menu(menuStr00)
        menu(text, "c")
    End Sub

    'print out a middle menu box with a bottom bar
    Public Sub mMenu(text As String)
        menu(menuStr03)
        menu(text, "c")
        menu(menuStr03)
    End Sub

    'print out a middle menu box with no bottom bar
    Public Sub moMenu(text As String)
        menu(menuStr03)
        menu(text, "c")
    End Sub

End Module
