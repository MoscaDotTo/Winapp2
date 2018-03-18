Option Strict On
Imports System.IO

Module Merge

    Dim dir As String = Environment.CurrentDirectory
    Dim name As String = "\winapp2.ini"
    Dim sfDir As String = Environment.CurrentDirectory
    Dim sfName As String
    Dim firstFile As iniFile
    Dim secondFile As iniFile
    Dim menuHasTopper As Boolean = False
    Dim exitCode As Boolean = False
    Dim mergeMode As Boolean

    Public Sub remoteMerge(winappDir As String, winappName As String, sDir As String, sName As String, mm As Boolean)
        'Handle being called from the commandline
        dir = winappDir
        name = winappName
        sfDir = sDir
        sfName = sName
        mergeMode = mm
        initMerge()
    End Sub

    Private Sub printMenu()
        If menuHasTopper Then
            printMenuLine(mMenu("Merge"))
        Else
            printMenuLine(tmenu("Merge"))
            printMenuLine(menu(menuStr03))
        End If
        Dim mergeStatus As String = IIf(mergeMode, "Replace & Add", "Replace & Remove").ToString
        printMenuLine("This tool will merge winapp2.ini entries from an external file into winapp2.ini.", "c")
        printMenuLine(menuStr01)
        printMenuLine("Merge Mode: " & mergeStatus, "c")
        printMenuLine(menuStr04)
        printMenuLine("0. Exit", "l")
        printMenuLine("1. Run (default)          - Merge Removed Entries.ini with Winapp2.ini", "l")
        printMenuLine("2. Run (custom)           - Merge Custom.ini with Winapp2.ini", "l")
        printMenuLine(menuStr01)
        printMenuLine("3. Toggle Merge mode      - Switch between merge modes.", "l")
        printMenuLine(menuStr02)
    End Sub

    Public Sub main()
        Console.Clear()
        exitCode = False
        menuHasTopper = False
        mergeMode = True
        While Not exitCode
            printMenu()
            Console.WriteLine()
            Console.Write("Enter a number, or leave blank to run the default: ")
            Dim input As String = Console.ReadLine
            Select Case input
                Case "0"
                    exitCode = True
                Case "1", ""
                    sfName = "\Removed Entries.ini"
                    initMerge()
                Case "2"
                    sfName = "\Custom.ini"
                    initMerge()
                Case "3"
                    mergeMode = Not mergeMode
                    Console.Clear()
                Case Else
                    Console.Clear()
                    menuHasTopper = True
                    printMenuLine(tmenu("Invalid Input. Please try again."))
            End Select
        End While
    End Sub

    Public Sub initMerge()
        'Initialize our inifiles
        firstFile = validate(dir, name, exitCode, "\winapp2.ini", "")
        If exitCode Then Exit Sub
        secondFile = validate(sfDir, sfName, exitCode, sfName, "")
        If exitCode Then Exit Sub

        'Merge them
        merge()

        'Flip our menu boolean
        revertMenu(exitCode)
    End Sub

    Private Sub merge()
        Dim out As String = ""

        'Process the merge mode and update the inifiles accordingly
        processMergeMode(firstFile, secondFile)

        'Parse our two files
        Dim tmp As New winapp2file(firstFile)
        Dim tmp2 As New winapp2file(secondFile)

        printMenuLine(bmenu("Merging " & firstFile.name.TrimStart(CChar("\")) & " with " & secondFile.name.TrimStart(CChar("\")), "c"))

        Try

            'Add the entries from the second file to their respective sections in the first file
            tmp.cEntriesW.AddRange(tmp2.cEntriesW)
            tmp.fxEntriesW.AddRange(tmp2.fxEntriesW)
            tmp.tbEntriesW.AddRange(tmp2.tbEntriesW)
            tmp.mEntriesW.AddRange(tmp2.mEntriesW)

            'Rebuild the internal changes
            tmp.rebuildToIniFiles()

            'Sort the merged sections 
            sortIniFile(tmp.cEntries, replaceAndSort(tmp.cEntries.getSectionNamesAsList, "-", "  "))
            sortIniFile(tmp.fxEntries, replaceAndSort(tmp.fxEntries.getSectionNamesAsList, "-", "  "))
            sortIniFile(tmp.tbEntries, replaceAndSort(tmp.tbEntries.getSectionNamesAsList, "-", "  "))
            sortIniFile(tmp.mEntries, replaceAndSort(tmp.mEntries.getSectionNamesAsList, "-", "  "))

            Dim file As New StreamWriter(firstFile.dir & "\" & firstFile.name, False)

            'write the merged winapp2string to file
            out += tmp.winapp2string
            file.Write(out)
            file.Close()
        Catch ex As Exception
            exc(ex)
        End Try

        printMenuLine(bmenu("Finished merging files. Press any key to return to the menu", "c"))

    End Sub

    Private Sub processMergeMode(ByRef first As iniFile, ByRef second As iniFile)
        Dim removeList As New List(Of String)

        'If mergemode is true, replace any matches
        If mergeMode Then
            For Each section In second.sections.Keys
                If first.sections.Keys.Contains(section) Then
                    first.sections.Item(section) = second.sections.Item(section)
                    removeList.Add(section)
                End If
            Next
            For Each section In removeList
                second.sections.Remove(section)
            Next
        Else
            'if mergemode is false, remove any matches
            For Each section In second.sections.Keys
                If first.sections.Keys.Contains(section) Then
                    first.sections.Remove(section)
                    removeList.Add(section)
                End If
            Next
        End If

        'Remove any processed sections from the second file so that only entries to add remain 
        For Each section In removeList
            second.sections.Remove(section)
        Next
    End Sub
End Module