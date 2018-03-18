Option Strict On
Imports System.IO

Module CCiniDebug

    Dim winappDir As String = Environment.CurrentDirectory
    Dim winappName As String = "\winapp2.ini"
    Dim ccDir As String = Environment.CurrentDirectory
    Dim ccName As String = "\ccleaner.ini"
    Dim menuHasTopper As Boolean = False
    Dim exitCode As Boolean
    Dim winappini As iniFile
    Dim ccini As iniFile
    Dim pruneFile As Boolean

    Private Sub printMenu()
        If Not menuHasTopper Then
            menuHasTopper = True
            printMenuLine(tmenu("CCiniDebug"))
        End If
        printMenuLine(menuStr03)
        printMenuLine("This tool will sort alphabetically the contents of ccleaner.ini", "c")
        printMenuLine("and can also prune stale winapp2.ini entries from it", "c")
        printMenuLine(menuStr04)
        printMenuLine("0. Exit                          - Return to the winapp2ool menu", "l")
        printMenuLine("1. Run (default)                 - Prune and sort ccleaner.ini", "l")
        printMenuLine("2. Run (sort only)               - Only sort ccleaner.ini", "l")
        printMenuLine(menuStr02)

    End Sub

    Public Sub Main()
        Console.Clear()
        exitCode = False
        menuHasTopper = False
        pruneFile = True
        Do Until exitCode
            printMenu()
            Console.WriteLine()
            Console.Write("Enter a number, or leave blank to run the default: ")
            Dim input As String = Console.ReadLine

            Try
                Select Case input
                    Case "0"
                        Console.WriteLine("Returning to winapp2ool menu...")
                        exitCode = True
                    Case "1", ""
                        Console.Clear()
                        initDebug()
                        revertMenu(exitCode)
                    Case "2"
                        Console.Clear()
                        pruneFile = False
                        initDebug()
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

    Public Sub remoteCC(waPath As String, waName As String, ccPath As String, cciname As String, prunecc As Boolean)
        'for handling ccinidebug parameters given through commandline args
        winappDir = waPath
        winappName = waName
        ccDir = ccPath
        ccName = cciname
        pruneFile = prunecc
        initDebug()
    End Sub

    Private Sub initDebug()
        'Load our inifiles into memory and execute the debugging steps
        ccini = validate(ccDir, ccName, exitCode, "\ccleaner.ini", "")
        If exitCode Then Exit Sub

        'Load winapp2.ini and use it to prune ccleaner.ini of stale entries
        If pruneFile Then
            winappini = validate(winappDir, winappName, exitCode, "\winapp2.ini", "")
            If exitCode Then Exit Sub
            printMenuLine(tmenu("CCiniDebug Results"))
            printMenuLine(mMenu("Pruning..."))
            prune(ccini.sections("Options"))
        End If
        'Sort the keys
        sortCC()

        'Write ccleaner.ini back to file
        writeCCini()
    End Sub

    Private Sub prune(ByRef optionsSec As iniSection)

        'collect the keys we must remove
        Dim tbTrimmed As New List(Of Integer)
        printMenuLine("The following lines with be removed from ccleaner.ini:", "l")
        printMenuLine(menuStr01)
        For i As Integer = 0 To optionsSec.keys.Count - 1
            Dim optionStr As String = optionsSec.keys.Values(i).toString

            'only operate on (app) keys
            If optionStr.StartsWith("(App)") And optionStr.Contains("*") Then
                optionStr = optionStr.Replace("(App)", "")
                optionStr = optionStr.Replace("=True", "")
                optionStr = optionStr.Replace("=False", "")
                If Not winappini.sections.ContainsKey(optionStr) Then
                    printMenuLine(optionsSec.keys.Values(i).lineString, "l")
                    tbTrimmed.Add(optionsSec.keys.Keys(i))
                End If
            End If
        Next

        printMenuLine(menuStr01)
        printMenuLine(tbTrimmed.Count & " lines will be removed.", "l")
        printMenuLine(menuStr02)

        'Remove the keys
        For Each key In tbTrimmed
            optionsSec.keys.Remove(key)
        Next

    End Sub

    Private Sub sortCC()
        'Sort the options section of ccleaner.ini 
        Dim lineList As List(Of String) = ccini.sections("Options").getKeysAsList
        lineList.Sort()
        lineList.Insert(0, "[Options]")
        ccini.sections("Options") = New iniSection(lineList)
    End Sub

    Private Sub writeCCini()
        'Write ccleaner.ini back to file
        Dim file As StreamWriter
        Try
            file = New StreamWriter(Environment.CurrentDirectory & ccName, False)
            file.Write(ccini.toString)
            file.Close()
            printMenuLine(bmenu("File " & ccName.TrimStart(CChar("\")) & " saved. Press any key to return to the winapp2ool menu.", "c"))
            If Not suppressOutput Then Console.ReadKey()
        Catch ex As Exception
            exc(ex)
        End Try
    End Sub
End Module