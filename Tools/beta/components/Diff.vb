Option Strict On
Imports System.IO

Module Diff

    Dim oldFile As iniFile
    Dim newFile As iniFile
    Dim oldFileDir As String = Environment.CurrentDirectory
    Dim newFileDir As String = Environment.CurrentDirectory
    Dim oldFileName As String = "\winapp2.ini"
    Dim newFileName As String = "\winapp2.ini"
    Dim outputToFile As String
    Dim exitCode As Boolean = False
    Dim menuHasTopper As Boolean = False
    Dim downloadFile As Boolean
    Dim downloadNCC As Boolean
    Dim saveLog As Boolean

    Private Sub printMenu()
        If Not menuHasTopper Then
            menuHasTopper = True
            printMenuLine(tmenu("Diff"))
        End If
        printMenuLine(menuStr03)
        printMenuLine("This tool will output the diff between two winapp2 files", "c")
        printMenuLine("Log Saving: " & If(saveLog, "On", "Off"), "c")
        printMenuLine(menuStr01)
        printMenuLine("Menu: Enter a number to select", "c")
        printMenuLine(menuStr01)
        printMenuLine("0. Exit                        - Return to the winapp2ool menu", "l")
        printMenuLine("1. Run (default)               - Run the diff tool", "l")
        printMenuLine("2. Run (online)                - Diff your local winapp2.ini against the latest", "l")
        printMenuLine("3. Run (online non ccleaner)   - Diff your local non-ccleaner winapp2.ini against the latest", "l")
        printMenuLine(menuStr01)
        printMenuLine("4. Toggle Save                        - Enable/Disable diff.txt saving", "l")
        printMenuLine(menuStr02)
    End Sub

    Public Sub main()
        menuHasTopper = False
        downloadFile = False
        exitCode = False
        saveLog = False
        outputToFile = ""
        Console.Clear()

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
                        initDiff()
                    Case "2"
                        downloadFile = True
                        downloadNCC = False
                        initDiff()
                    Case "3"
                        downloadFile = True
                        downloadNCC = True
                        initDiff()
                    Case "4"
                        saveLog = Not saveLog
                        Console.Clear()
                        printMenuLine(tmenu("Log saving toggled."))
                    Case Else
                        Console.Clear()
                        printMenuLine(tmenu("Invalid input. Please try again."))
                End Select
            Catch ex As Exception
                exc(ex)
            End Try
        Loop
    End Sub

    Private Sub initDiff()
        loadFiles()
        differ()
        If saveLog Then saveDiff()
        revertMenu(exitCode)
        Console.Clear()
    End Sub

    'Print out the menu for selecting the files to diff
    Private Function printFileLoader(ageType As String, ByRef path As String, ByRef name As String) As iniFile
        printMenuLine(tmenu("Diff file Loader"))
        printMenuLine(menuStr03)
        printMenuLine("Options", "c")
        printMenuLine("Enter the name of the " & ageType & " file", "l")
        printMenuLine("Enter '0' to return to the menu", "l")
        printMenuLine("Enter '1' to open the file/directory chooser", "l")
        printMenuLine("Leave blank to use the default (winapp2.ini)", "l")
        printMenuLine(menuStr02)
        Console.WriteLine()
        Dim input As String = Console.ReadLine()
        Select Case input
            Case ""
              name = "\winapp2.ini"
            Case "0"
                exitCode = True
                Return Nothing
            Case "1"
                fChooser(path, name, exitCode, "\winapp2.ini", "")
            Case Else
                name = "\" & input
        End Select

        Return validate(path, name, exitCode, "\winapp2.ini", "")
    End Function

    Public Sub remoteDiff(firstDir As String, firstName As String, secondDir As String, secondName As String, download As Boolean, ncc As Boolean, save As Boolean)
        'Handle commandline args for launching the differ
        oldFileDir = firstDir
        oldFileName = firstName
        newFileDir = secondDir
        newFileName = secondName
        downloadFile = download
        downloadNCC = ncc
        saveLog = save
        initDiff()
    End Sub

    'load up the files to diff
    Private Sub loadFiles()
        Console.Clear()
        Try
            'Always collect the older file
            oldFile = printFileLoader("older", oldFileDir, oldFileName)
            If exitCode Then Exit Sub

            'Collect the second file conditionally based on whether its a download or a local file
            newFile = CType(If(downloadFile, getRemoteWinapp(downloadNCC), printFileLoader("newer", newFileDir, newFileName)), iniFile)
            If exitCode Then Exit Sub
        Catch ex As Exception
            exc(ex)
        End Try
    End Sub

    Private Sub differ()
        If exitCode Then Exit Sub
        Console.Clear()
        Try
            'collect & verify version #s and print them out for the menu
            Dim fver As String = oldFile.comments(0).comment.ToString
            fver = IIf(fver.ToLower.Contains("version"), fver.TrimStart(CChar(";")).Replace("Version:", "version"), " version not given").ToString

            Dim sver As String = newFile.comments(0).comment.ToString
            sver = IIf(sver.ToLower.Contains("version"), sver.TrimStart(CChar(";")).Replace("Version:", "version"), " version not given").ToString

            outputToFile += tmenu("Changes made between" & fver & " and" & sver) & Environment.NewLine
            outputToFile += menu(menuStr02) & Environment.NewLine
            Console.WriteLine()
            outputToFile += menu(menuStr00) & Environment.NewLine

            'compare the files and then ennumerate their changes
            Dim outList As List(Of String) = compareTo()
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
            outputToFile += menu(menuStr02)
            If Not suppressOutput Then Console.Write(outputToFile)
        Catch ex As Exception
            If ex.Message = "The given key was not present in the dictionary." Then
                Console.WriteLine("Error encountered during diff: " & ex.Message)
                Console.WriteLine("This error is typically caused by invalid file names, please double check your input and try again.")
                Console.WriteLine()
            Else
                exc(ex)
            End If
        End Try

        Console.WriteLine()
        printMenuLine(bmenu("Press any key to return to the winapp2ool menu.", "l"))
        Console.ReadKey()
    End Sub

    Private Function compareTo() As List(Of String)

        Dim outList As New List(Of String)
        Dim comparedList As New List(Of String)

        For Each section In oldFile.sections.Values
            Try
                'If we're looking at an entry in the old file and the new file contains it, and we haven't yet processed this entry
                If newFile.sections.Keys.Contains(section.name) And Not comparedList.Contains(section.name) Then
                    Dim sSection As iniSection = newFile.sections(section.name)
                    'and if that entry in the new file does not compareTo the entry in the old file, we have a modified entry
                    If Not section.compareTo(sSection) Then outList.Add(getDiff(sSection, "modified."))

                ElseIf Not newFile.sections.Keys.Contains(section.name) And Not comparedList.Contains(section.name) Then
                    'If we do not have the entry in the new file, it has been removed between versions 
                    outList.Add(getDiff(section, "removed."))
                End If
                comparedList.Add(section.name)
            Catch ex As Exception
                exc(ex)
            End Try
        Next

        For Each section In newFile.sections.Values
            'Any sections from the new file which are not found in the old file have been added
            If Not oldFile.sections.Keys.Contains(section.name) Then outList.Add(getDiff(section, "added."))
        Next

        Return outList
    End Function

    Private Function getDiff(section As iniSection, changeType As String) As String
        'Return a string containing a box containing the change type and entry name, followed by the entry's tostring
        Dim out As String = ""
        out += mkMenuLine(section.name & " has been " & changeType, "c") & Environment.NewLine
        out += mkMenuLine(menuStr02, "") & Environment.NewLine & Environment.NewLine
        out += section.ToString & Environment.NewLine
        out += menuStr00
        Return out
    End Function

    Private Sub saveDiff()
        'Save diff.txt 
        Try
            Dim file As New StreamWriter(Environment.CurrentDirectory & "\diff.txt", False)
            file.Write(outputToFile)
            file.Close()
        Catch ex As Exception
            exc(ex)
        End Try
    End Sub
End Module