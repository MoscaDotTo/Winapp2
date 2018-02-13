Option Strict On
Imports System.IO

Module Diff

    Dim firstFile As iniFile
    Dim secondFile As iniFile
    Dim dir As String = Environment.CurrentDirectory
    Dim dir2 As String = Environment.CurrentDirectory
    Dim exitCode As Boolean = False
    Dim menuHasTopper As Boolean = False

    Private Sub printMenu()
        If Not menuHasTopper Then
            menuHasTopper = True
            tmenu("Diff")
        End If
        menu(menuStr03, "")
        menu("This tool will output the diff between two winapp2 files", "c")
        menu(menuStr01, "")
        menu("Menu: Enter a number to select", "c")
        menu(menuStr01, "")
        menu("0. Exit                        - Return to the winapp2ool menu", "l")
        menu("1. Run (default)               - Run the diff tool", "l")
        menu(menuStr02, "")
    End Sub

    Private Sub revertToMenu()
        If exitCode Then
            exitCode = False
            menuHasTopper = False
            Console.Clear()
        Else
            exitCode = True
        End If
    End Sub

    Public Sub main()
        Console.Clear()
        exitCode = False
        menuHasTopper = False
        Do Until exitCode
            printMenu()
            Console.WriteLine()
            Console.Write("Enter a number, or leave blank to run the default: ")

            Dim input As String = Console.ReadLine()
            Try
                Select Case input
                    Case "0"
                        Console.WriteLine("Exiting diff...")
                        exitCode = True
                    Case "1", ""
                        loadFiles()
                        differ()
                        revertToMenu()
                    Case Else
                        Console.Clear()
                        tmenu("Invalid input. Please try again.")
                End Select
            Catch ex As Exception
                Console.WriteLine("Error: " & ex.ToString)
                Console.WriteLine("Please report this error on GitHub")
                Console.WriteLine()
            End Try
        Loop
    End Sub

    Private Sub loadFiles()
        Console.Clear()
        Try
            tmenu("Diff file Loader")
            menu(menuStr03)
            menu("Options", "c")
            menu("Enter the name of the older file", "l")
            menu("Enter '0' to return to the menu", "l")
            menu("Leave blank to open the file/directory chooser", "l")
            menu(menuStr02)
            Dim oName As String = "\" & Console.ReadLine()
            If oName = "\" Then
                fChooser(dir, oName, exitCode, "\winapp2.ini", "")
            ElseIf oName = "\0" Then
                exitCode = True
                Exit Sub
            End If
            validate(dir, oName, exitCode, "\winapp2.ini", "")

            If exitCode Then
                Exit Sub
            End If

            tmenu("Diff file Loader")
            menu(menuStr03)
            menu("Options", "c")
            menu("Enter the name of the newer file", "l")
            menu("Enter '0' to return to the menu", "l")
            menu("Leave blank to open the file/directory chooser", "l")
            menu(menuStr02)
            Dim nName As String = "\" & Console.ReadLine()
            If nName = "\" Then
                fChooser(dir2, nName, exitCode, "\winapp2.ini", "")
            ElseIf nName = "\0" Then
                exitCode = True
                Exit Sub
            End If
            validate(dir2, nName, exitCode, "\winapp2.ini", "")

            If exitCode Then
                Exit Sub
            End If
            firstFile = New iniFile(dir, oName)
            secondFile = New iniFile(dir2, nName)
        Catch ex As Exception
            Console.WriteLine("Error: " & ex.ToString)
            Console.WriteLine("Please report this error on GitHub")
            Console.WriteLine()
        End Try

    End Sub

    Private Sub differ()
        If exitCode = True Then
            Exit Sub
        End If

        Console.Clear()
        Try
            Dim fver As String = firstFile.comments(0).comment.ToString
            fver = fver.Split(Convert.ToChar(";"))(1)
            Dim sver As String = secondFile.comments(0).comment.ToString
            sver = sver.Split(Convert.ToChar(";"))(1)
            tmenu("Changes made between" & fver & " and" & sver)
            menu(menuStr02, "")
            Console.WriteLine()
            menu(menuStr00, "")
            Dim outList As List(Of String) = firstFile.compareTo(secondFile)
            Dim remCt As Integer = 0
            Dim modCt As Integer = 0
            Dim addCt As Integer = 0
            For Each change As String In outList

                If change.Contains("has been added.") Then
                    addCt += 1
                ElseIf change.Contains("has been removed") Then
                    remCt += 1
                Else
                    modCt += 1
                End If
                Console.WriteLine(change)
            Next

            menu("Diff complete.", "c")
            menu(menuStr03, "")
            menu("Added entries: " & addCt, "l")
            menu("Modified entries: " & modCt, "l")
            menu("Removed entries: " & remCt, "l")
        Catch ex As Exception
            If ex.Message = "The given key was not present in the dictionary." Then
                Console.WriteLine("Error encountered during diff: " & ex.Message)
                Console.WriteLine("This error is typically caused by invalid file names, please double check your input and try again.")
                Console.WriteLine()
            Else
                Console.WriteLine("Error: " & ex.ToString)
                Console.WriteLine("Please report this error on GitHub")
                Console.WriteLine()
            End If

        End Try
        menu(menuStr01)
        menu("Press any key to return to the winapp2ool menu.", "l")
        menu(menuStr02, "")
        Console.ReadKey()
    End Sub

End Module