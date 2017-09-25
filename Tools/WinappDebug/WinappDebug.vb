Imports System.IO
Module Module1

    Sub Main()

        'Check if the winapp2.ini file is in the current directory. End program if it isn't.
        If File.Exists(Environment.CurrentDirectory & "\winapp2.ini") = False Then
            Console.WriteLine("winapp2.ini file could not be located in the current working directory (" & Environment.CurrentDirectory & ")")
            Console.ReadKey()
            End
        End If

        'Create some variables for later use
        Dim linecount As Integer = 0
        Dim command As String
        Dim number_of_errors As Integer = 0

        'Create a list of supported environmental variables
        Dim envir_vars As New List(Of String)
        envir_vars.AddRange(New String() {"%userprofile%", "%ProgramFiles%", "%rootdir%", "%windir%", "%appdata%", "%systemdrive%"})

        'Create a list of entry titles so we can look for duplicates
        Dim entry_titles As New List(Of String)

        'Read the winapp2.ini file line-by-line
        Dim r As IO.StreamReader
        Try

            r = New IO.StreamReader(Environment.CurrentDirectory & "\winapp2.ini")
            Do While (r.Peek() > -1)

                'Update the line that is being tested
                command = (r.ReadLine.ToString)

                'Increment the lineocunt value
                linecount = linecount + 1

                'Whitespace checks
                If command = "" = False And command.StartsWith(";") = False Then

                    'Check for trailing whitespace
                    If command.EndsWith(" ") Then
                        Console.WriteLine("Line: " & linecount & " Error: Detected unwanted whitespace at end of line." & Environment.NewLine & "Command: " & command & Environment.NewLine)
                        number_of_errors = number_of_errors + 1
                    End If

                    'Check for ending whitespace
                    If command.StartsWith(" ") Then
                        Console.WriteLine("Line: " & linecount & " Error: Detected unwanted whitespace at beginning of line." & Environment.NewLine & "Command: " & command & Environment.NewLine)
                        number_of_errors = number_of_errors + 1
                    End If

                End If

                'Check for duplicate titles
                If command.StartsWith("[") Then

                    'Check if it's already in the list
                    If entry_titles.Contains(command) Then
                        Console.WriteLine("Line: " & linecount & " Duplicate entry." & Environment.NewLine & "Command: " & command & Environment.NewLine)
                        number_of_errors = number_of_errors + 1
                    Else

                        'Not already in the list. Add it.
                        entry_titles.Add(command)

                    End If

                End If

                'Check for spelling errors in "LangSecRef"
                If command.StartsWith("Lan") Then
                    If command.Contains("LangSecRef=") = False Then
                        Console.WriteLine("Line: " & linecount & " Error: 'LangSecRef' entry is incorrectly spelled or formatted." & Environment.NewLine & "Command: " & command & Environment.NewLine)
                        number_of_errors = number_of_errors + 1
                    End If
                End If

                If command = "Default=True" Or command = "default=true" Then
                    Console.WriteLine("Line: " & linecount & " Error: All entries should be disabled by default." & Environment.NewLine & "Command: " & command & Environment.NewLine)
                    number_of_errors = number_of_errors + 1
                End If

                'Check for environmental variable spacing errors
                If command.Contains("%Program Files%") Then
                    Console.WriteLine("Line: " & linecount & " Error: '%ProgramFiles%' variable should not have spacing." & Environment.NewLine & "Command: " & command & Environment.NewLine)
                    number_of_errors = number_of_errors + 1
                End If

                'Check for cleaning command spelling errors (files)
                If command.StartsWith("Fi") Then
                    If command.Contains("FileKey") = False Then
                        Console.WriteLine("Line: " & linecount & " Error: 'FileKey' entry is incorrectly spelled or formatted. Spelling should be CamelCase." & Environment.NewLine & "Command: " & command & Environment.NewLine)
                        number_of_errors = number_of_errors + 1
                    End If
                End If

                'Check for cleaning command spelling errors (registry keys)
                If command.StartsWith("Re") Then
                    If command.Contains("RegKey") = False Then
                        Console.WriteLine("Line: " & linecount & " Error: 'RegKey' entry is incorrectly spelled or formatted. Spelling should be CamelCase." & Environment.NewLine & "Command: " & command & Environment.NewLine)
                        number_of_errors = number_of_errors + 1
                    End If
                End If

                'Check for missing numbers next to cleaning commands
                If command.StartsWith("FileKey=") Or command.StartsWith("RegKey=") Then
                    Console.WriteLine("Line: " & linecount & " Error: Cleaning path entry needs to have a trailing number." & Environment.NewLine & "Command: " & command & Environment.NewLine)
                    number_of_errors = number_of_errors + 1
                End If

                'Check for Detect that contains a filepath instead of a registry path
                If command.StartsWith("Detect=%") Or command.StartsWith("Detect=C:\") Then
                    Console.WriteLine("Line: " & linecount & " Error: 'Detect' can only be used for registry key paths." & Environment.NewLine & "Command: " & command & Environment.NewLine)
                    number_of_errors = number_of_errors + 1
                End If

                'Check for detectfile that contains a registry path
                If command.StartsWith("DetectFile=HKLM") Or command.StartsWith("DetectFile=HKCU") Or command.StartsWith("DetectFile=HKC") Or command.StartsWith("DetectFile=HKCR") Then
                    Console.WriteLine("Line: " & linecount & " Error: 'DetectFile' can only be used for filesystem paths." & Environment.NewLine & "Command: " & command & Environment.NewLine)
                    number_of_errors = number_of_errors + 1
                End If

                'Check for missing backslashes on environmental variables
                For Each var As String In envir_vars
                    If command.Contains(var) = True Then
                        If command.Contains(var & "\") = False Then
                            Console.WriteLine("Line: " & linecount & " Error: Environmental variables must have a trailing backslash." & Environment.NewLine & "Command: " & command & Environment.NewLine)
                            number_of_errors = number_of_errors + 1
                        End If
                    End If
                Next


            Loop

            'Close unmanaged code calls
            r.Close()


                Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try



        'Stop the program from closing on completion
        Console.WriteLine("***********************************************" & Environment.NewLine & "Completed analysis of winapp2.ini. " & number_of_errors & " errors were detected. Press any key to close.")
        Console.ReadKey()


    End Sub

End Module
