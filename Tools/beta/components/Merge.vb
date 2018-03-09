Option Strict On
Imports System.IO

Module Merge

    Dim dir As String = Environment.CurrentDirectory
    Dim name As String = "\winapp2.ini"
    Dim sfDir As String = Environment.CurrentDirectory
    Dim sfName As String
    Dim menuHasTopper As Boolean = False
    Dim exitCode As Boolean

    Private Sub printMenu()
        If menuHasTopper Then
            printMenuLine(mMenu("Merge"))
        Else
            printMenuLine(tmenu("Merge"))
            printMenuLine(menu(menuStr03))
        End If
        printMenuLine("This tool will merge winapp2.ini entries from an external file into winapp2.ini.", "c")
        printMenuLine(menuStr04)
        printMenuLine("0. Exit", "l")
        printMenuLine("1. Run (default)          - Merge Removed Entries.ini with Winapp2.ini", "l")
        printMenuLine("2. Run (custom)           - Merge Custom.ini with Winapp2.ini", "l")
        printMenuLine(menuStr02)

    End Sub

    Public Sub main()
        Console.Clear()
        exitCode = False
        menuHasTopper = False
        While Not exitCode
            printMenu()
            Console.WriteLine()
            Console.Write("Enter a number, or leave blank to run the default: ")
            Dim input As String = Console.ReadLine
            Select Case input
                Case "0"
                    exitCode = True
                Case "1", ""
                    initMerge("\Removed Entries.ini", sfDir, "\Removed Entries.ini")
                    revertMenu(exitCode)
                Case "2"
                    initMerge("\Custom.ini", sfDir, "\Custom.ini")
                    revertMenu(exitCode)
                Case Else
                    Console.Clear()
                    menuHasTopper = True
                    printMenuLine(tmenu("Invalid Input. Please try again."))
            End Select

        End While

    End Sub

    Public Sub initMerge(secondFileName As String, secondFileDir As String, defaultName As String)
        sfDir = secondFileDir
        sfName = secondFileName
        validate(dir, name, exitCode, "\winapp2.ini", "")
        If exitCode Then Exit Sub
        Dim mergeFile As New iniFile(dir, name)
        validate(sfDir, sfName, exitCode, defaultName, "")
        If exitCode Then Exit Sub
        Dim secondFile As New iniFile(sfDir, sfName)
        merge(mergeFile, secondFile)

    End Sub

    Private Sub merge(ByRef firstFile As iniFile, secondFile As iniFile)
        Dim out As String = ""

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
            file.WriteLine(out)

        Catch ex As Exception
            exc(ex)
        End Try

        printMenuLine(bmenu("Finished merging files. Press any key to return to the menu", "c"))

    End Sub

End Module