Option Strict On
Imports System.IO

Module Diff

    Dim firstFile As iniFile
    Dim secondFile As iniFile
    Dim dir As String = Environment.CurrentDirectory
    Dim dir2 As String = Environment.CurrentDirectory
    Dim name1 As String = "\winapp2.ini"
    Dim name2 As String = "\winapp2.ini"
    Dim exitCode As Boolean = False
    Dim menuHasTopper As Boolean = False
    Dim outputToFile As String

    Private Sub printMenu()
        If Not menuHasTopper Then
            menuHasTopper = True
            printMenuLine(tmenu("Diff"))
        End If
        printMenuLine(menuStr03)
        printMenuLine("This tool will output the diff between two winapp2 files", "c")
        printMenuLine(menuStr01)
        printMenuLine("Menu: Enter a number to select", "c")
        printMenuLine(menuStr01)
        printMenuLine("0. Exit                        - Return to the winapp2ool menu", "l")
        printMenuLine("1. Run (default)               - Run the diff tool", "l")
        printMenuLine("2. Run (online)                - Diff your local winapp2.ini against the latest", "l")
        printMenuLine("3. Run (online non ccleaner)   - Diff your local non-ccleaner winapp2.ini against the latest", "l")
        printMenuLine(menuStr02)
    End Sub

    Public Sub main()
        outputToFile = ""
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
                        firstFile = Nothing
                        secondFile = Nothing
                        loadFiles()
                        differ()
                        revertMenu(exitCode)
                    Case "2"
                        firstFile = Nothing
                        secondFile = getRemoteIniFile(Downloader.wa2Link, "\winapp2.ini")
                        loadFiles()
                        differ()
                        revertMenu(exitCode)
                    Case "3"
                        firstFile = Nothing
                        secondFile = getRemoteIniFile(Downloader.nonccLink, "\winapp2.ini")
                        loadFiles()
                        differ()
                        revertMenu(exitCode)
                    Case Else
                        Console.Clear()
                        printMenuLine(tmenu("Invalid input. Please try again."))
                End Select
            Catch ex As Exception
                exc(ex)
            End Try
        Loop
    End Sub

    'Print out the menu for selecting the files to diff
    Private Sub printFileLoader(ageType As String, ByRef path As String, ByRef name As String)
        printMenuLine(tmenu("Diff file Loader"))
        printMenuLine(menuStr03)
        printMenuLine("Options", "c")
        printMenuLine("Enter the name of the " & ageType & " file", "l")
        printMenuLine("Enter '0' to return to the menu", "l")
        printMenuLine("Leave blank to open the file/directory chooser", "l")
        printMenuLine(menuStr02)
        Console.WriteLine()
        name = "\" & Console.ReadLine()
        If name = "\" Then
            fChooser(path, name, exitCode, "\winapp2.ini", "")
        ElseIf name = "\0" Then
            exitCode = True
        End If
        validate(path, name, exitCode, "\winapp2.ini", "")
    End Sub

    'load up the files to diff
    Private Sub loadFiles()
        Console.Clear()
        Try
            'Always collect the older file
            printFileLoader("older", dir, name1)
            firstFile = New iniFile(dir, name1)
            If exitCode Then Exit Sub

            'Only collect the file if it doesn't exist (which it will if we're comparing to an online ini)
            If secondFile Is Nothing Then
                printFileLoader("newer", dir2, name2)
                secondFile = New iniFile(dir2, name2)
            End If
            If exitCode Then Exit Sub
        Catch ex As Exception
            exc(ex)
        End Try
    End Sub

    Private Sub differ()
        If exitCode Then Exit Sub
        Console.Clear()
        Try
            'collect the version #s and print them out
            Dim fver As String = firstFile.comments(0).comment.ToString
            fver = fver.TrimStart(CChar(";")).Replace("Version:", "version")
            Dim sver As String = secondFile.comments(0).comment.ToString
            sver = sver.TrimStart(CChar(";")).Replace("Version:", "version")
            outputToFile += tmenu("Changes made between" & fver & " and" & sver) & Environment.NewLine
            outputToFile += menu(menuStr02) & Environment.NewLine
            Console.WriteLine()
            outputToFile += menu(menuStr00) & Environment.NewLine

            'compare the files and then ennumerate their changes
            Dim outList As List(Of String) = compareTo(firstFile, secondFile)
            Dim remCt As Integer = 0
            Dim modCt As Integer = 0
            Dim addCt As Integer = 0
            For Each change In outList

                If change.Contains("has been added.") Then
                    addCt += 1
                ElseIf change.Contains("has been removed") Then
                    remCt += 1
                Else
                    modCt += 1
                End If
                outputToFile += change & Environment.NewLine
            Next

            outputToFile += menu("Diff complete.", "c") & Environment.NewLine
            outputToFile += menu(menuStr03) & Environment.NewLine
            outputToFile += menu("Added entries: " & addCt, "l") & Environment.NewLine
            outputToFile += menu("Modified entries: " & modCt, "l") & Environment.NewLine
            outputToFile += menu("Removed entries: " & remCt, "l") & Environment.NewLine
            Console.Write(outputToFile)
            printMenuLine(menuStr01)
            outputToFile += menu(menuStr02)
        Catch ex As Exception
            If ex.Message = "The given key was not present in the dictionary." Then
                Console.WriteLine("Error encountered during diff: " & ex.Message)
                Console.WriteLine("This error is typically caused by invalid file names, please double check your input and try again.")
                Console.WriteLine()
            Else
                exc(ex)
            End If

        End Try

        'offer to save to file
        printMenuLine(mMenu("Would you like to save the results to file?"))
        printMenuLine("Options", "c")
        printMenuLine("Enter '1' to save diff.txt to the current directory.", "l")
        printMenuLine("Enter anything else to continue without saving.", "l")
        printMenuLine(menuStr02)
        Console.WriteLine()

        Dim input As String = Console.ReadLine()
        If input = "1" Then
            Try
                Dim file As New StreamWriter(Environment.CurrentDirectory & "\diff.txt", False)
                file.WriteLine(outputToFile)
                file.Close()
            Catch ex As Exception
                exc(ex)
            End Try
        End If
        Console.WriteLine()
        printMenuLine(bmenu("Press any key to return to the winapp2ool menu.", "l"))
        Console.ReadKey()
    End Sub

    Public Function compareTo(oldFile As iniFile, newFile As iniFile) As List(Of String)

        Dim outList As New List(Of String)
        Dim comparedList As New List(Of String)

        For i As Integer = 0 To oldFile.sections.Count - 1
            Dim curSection As iniSection = oldFile.sections.Values(i)
            Dim curName As String = curSection.name
            Try
                If newFile.sections.Keys.Contains(curName) And Not comparedList.Contains(curName) Then
                    Dim sSection As iniSection = newFile.sections(curName)
                    If Not curSection.compareTo(sSection) Then outList.Add(getDiff(sSection, "modified."))
                ElseIf Not newfile.sections.Keys.Contains(curName) And Not comparedList.Contains(curName) Then
                    outList.Add(getDiff(curSection, "removed."))
                End If
                comparedList.Add(curName)

            Catch ex As Exception
                exc(ex)
            End Try
        Next

        For i As Integer = 0 To secondFile.sections.Count - 1
            Dim curSection As iniSection = secondFile.sections.Values(i)
            Dim curName As String = curSection.name

            If Not firstFile.sections.Keys.Contains(curName) Then outList.Add(getDiff(curSection, "added."))
        Next

        Return outList
    End Function

    Private Function getDiff(section As iniSection, changeType As String) As String
        Dim out As String = ""
        out += mkMenuLine(section.name & " has been " & changeType, "c") & Environment.NewLine
        out += mkMenuLine(menuStr02, "") & Environment.NewLine & Environment.NewLine
        out += section.ToString & Environment.NewLine
        out += menuStr00
        Return out
    End Function

End Module