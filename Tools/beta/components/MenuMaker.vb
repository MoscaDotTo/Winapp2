Option Strict On
Module MenuMaker

    'basic menu frames & strings
    Public menuStr00 As String = " ╔════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╗"
    Public menuStr01 As String = " ║                                                                                                                    ║"
    Public menuStr02 As String = " ╚════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╝"
    Public menuStr03 As String = " ╠════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╣"
    Public menuStr04 As String = menu(menuStr01) & Environment.NewLine & mkMenuLine("Menu: Enter a number to select", "c") & Environment.NewLine & mkMenuLine(menuStr01, "")
    Public anyKeyStr As String = "Press any key to return to the winapp2ool menu."
    Public invInpStr As String = "Invalid input. Please try again."
    Public promptStr As String = "Enter a number, or leave blank to run the default: "

    'the maximum length of the portion of the first half of a '#. Option - Description' style menu line
    Dim menuItemLength As Integer

    'indicates whether or not we are pending an exit from the menu
    Public exitCode As Boolean

    'Holds the output text at the top of the menu
    Public menuTopper As String

    Dim optNum As Integer = 0

    'Initalize the menu 
    Public Sub initMenu(topper As String, itemlen As Integer)
        exitCode = False
        menuTopper = topper
        menuItemLength = itemlen
    End Sub

    'Print an empty menu line
    Public Sub printBlankMenuLine()
        printMenuLine(menuStr01)
    End Sub

    'Print a menu option/line conditionally 
    Public Sub printIf(cond As Boolean, printType As String, str1 As String, str2 As String)
        Select Case True
            Case cond And printType = "line"
                printMenuLine(str1, str2)
            Case cond And printType = "opt"
                printMenuOpt(str1, str2)
            Case cond And printType = "reset"
                printResetStr(str1)
        End Select
    End Sub

    'Return the inverse state of a given setting for enable/disable purposes 
    Public Function enStr(setting As Boolean) As String
        Return If(setting, "Disable", "Enable")
    End Function

    'Print the reset settings menu option for any given module
    Public Sub printResetStr(moduleName As String)
        printBlankMenuLine()
        printMenuOpt("Reset Settings", "Restore " & moduleName & "'s settings to their deault state")
    End Sub

    'Print the top of the menu containing the topper, any description text, the menu prompt, and the exit option
    Public Sub printMenuTop(descriptionItems As String(), printExit As Boolean)
        printMenuLine(tmenu(menuTopper))
        printMenuLine(menuStr03)
        For Each line In descriptionItems
            printMenuLine(line, "c")
        Next
        printMenuLine(menuStr04)
        optNum = 0
        If printExit Then printMenuOpt("Exit", "Return to the menu")
    End Sub

    'constructs a menu string 
    Public Function menu(lineString As String) As String
        Return mkMenuLine(lineString, "")
    End Function

    'prints a menu string
    Public Sub printMenuLine(lineString As String)
        cwl(menu(lineString))
    End Sub

    'constructs a menu string with alignment
    Public Function menu(lineString As String, align As String) As String
        Return mkMenuLine(lineString, align)
    End Function

    'prints a menu string with alignment
    Public Sub printMenuLine(lineString As String, align As String)
        cwl(menu(lineString, align))
    End Sub

    'Prints a numbered menu option to a set length 
    Public Sub printMenuOpt(lineString1 As String, lineString2 As String)
        lineString1 = optNum & ". " & lineString1
        While lineString1.Length < menuItemLength
            lineString1 += " "
        End While
        cwl(menu(lineString1 & "- " & lineString2, "l"))
        optNum += 1
    End Sub

    'Flip the exitCode status so we can return to menu when desired
    Public Sub revertMenu()
        exitCode = Not exitCode
    End Sub

    Public Sub undoAnyPendingExits()
        If exitCode Then exitCode = False
    End Sub

    'Construct a menu line properly fit to the width of the console
    Public Function mkMenuLine(line As String, align As String) As String
        If line.Length >= 119 Then
            Return line
        End If
        Dim out As String = " ║"
        Select Case align
            Case "c"
                While out.Length < (((118 - line.Length)/ 2) + 2)
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

    'Replace instances of Environment.CurrentDirectory in a path with "..\"
    Public Function replDir(dirStr As String) As String
        If dirStr.Contains(Environment.CurrentDirectory) Then
            dirStr = dirStr.Replace(Environment.CurrentDirectory, "..")
        End If

        Return dirStr
    End Function

    'Observe which preceding options in the menu are enabled and return the proper offset from a minumum menu number
    Public Function getMenuNumber(valList As Boolean(), lowestStartingNum As Integer) As Integer
        For Each setting In valList
            If setting Then lowestStartingNum += 1
        Next
        Return lowestStartingNum
    End Function

End Module
