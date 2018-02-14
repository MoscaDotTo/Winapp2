Option Strict On
Imports System.IO

Module CCiniDebug

    Dim dir As String = Environment.CurrentDirectory
    Dim name As String = "\winapp2.ini"
    Dim ccDir As String = Environment.CurrentDirectory
    Dim ccName As String = "\ccleaner.ini"
    Dim menuHasTopper As Boolean = False

    Private Sub printMenu()
        If Not menuHasTopper Then
            menuHasTopper = True
            tmenu("CCiniDebug")
        End If
        menu(menuStr03, "")
        menu("This tool will sort alphabetically the contents of ccleaner.ini", "c")
        menu("and can also prune stale winapp2.ini entries from it", "c")
        menu(menuStr04)
        menu("0. Exit                          - Return to the winapp2ool menu", "l")
        menu("1. Run (default)                 - Prune and sort ccleaner.ini", "l")
        menu("2. Run (sort only)               - Only sort ccleaner.ini", "l")
        menu(menuStr02, "")

    End Sub

    Public Sub Main()
        Console.Clear()
        Dim exitCode As Boolean = False
        Dim ccini As iniFile
        menuHasTopper = False
        Dim winappini As iniFile

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

                        validate(ccDir, ccName, exitCode, "\ccleaner.ini", "")
                        If exitCode Then
                            Exit Sub
                        End If
                        ccini = New iniFile(ccDir, ccName)

                        validate(dir, name, exitCode, "\winapp2.ini", "")
                        If exitCode Then
                            Exit Sub
                        End If
                        winappini = New iniFile(dir, name)

                        tmenu("CCiniDebug Results")
                        mMenu("Pruning...")
                        ccini.sections("Options") = prune(ccini.sections("Options"), winappini)
                        Dim iexitcode As Boolean = False
                        Console.WriteLine()
                        tmenu("Save File?")
                        Do Until iexitcode
                            menu(menuStr03)
                            menu("Options", "c")
                            menu("Enter '0' to return to the menu", "l")
                            menu("Enter '1' to change the save file name", "l")
                            menu("Leave blank to save as " & ccName.Split(CChar("\"))(1), "l")
                            menu(menuStr02)
                            Dim in2 As String = Console.ReadLine
                            Select Case in2
                                Case "0"
                                    iexitcode = True
                                Case ""
                                    iexitcode = True
                                Case "1"
                                    fChooser(ccDir, ccName, exitCode, "\ccleaner.ini", "")
                                    iexitcode = True
                                Case Else
                                    Console.Clear()
                                    tmenu("Invalid input. Please try again.")
                            End Select
                        Loop
                        writeCCini(ccini.sections("Options"))
                        revertMenu(exitCode)
                    Case "2"
                        Console.Clear()
                        validate(ccDir, ccName, exitCode, "\ccleaner.ini", "")
                        ccini = New iniFile(ccDir, ccName)
                        writeCCini(ccini.sections("Options"))
                        exitCode = True
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

    Private Sub writeCCini(ccini As iniSection)
        Dim file As StreamWriter
        Dim lineList As List(Of String) = ccini.getKeysAsList
        lineList.Sort()
        lineList.Insert(0, "[Options]")
        ccini = New iniSection(lineList)
        Try
            file = New StreamWriter(Environment.CurrentDirectory & ccName, False)
            file.Write(ccini.ToString)
            file.Close()
            bmenu("File " & ccName.Split(CChar("\"))(1) & " saved. Press any key to return to the winapp2ool menu.", "c")
            Console.ReadKey()
        Catch ex As Exception
            Console.WriteLine("Error: " & ex.ToString)
            Console.WriteLine("Please report this error on GitHub")
            Console.WriteLine()
        End Try
    End Sub

    Private Function prune(ccini As iniSection, winappini As iniFile) As iniSection

        'collect the keys we must remove
        Dim tbTrimmed As New List(Of Integer)
        menu("The following lines with be removed from ccleaner.ini:", "l")
        menu(menuStr01)
        For i As Integer = 0 To ccini.keys.Count - 1
            Dim optionStr As String = ccini.keys.Values(i).toString

            'only operate on (app) keys
            If optionStr.StartsWith("(App)") And optionStr.Contains("*") Then
                optionStr = optionStr.Replace("(App)", "")
                optionStr = optionStr.Replace("=True", "")
                optionStr = optionStr.Replace("=False", "")
                If Not winappini.sections.ContainsKey(optionStr) Then
                    menu(ccini.keys.Values(i).lineString, "l")
                    tbTrimmed.Add(ccini.keys.Keys(i))
                End If
            End If
        Next

        'reverse the keys we must remove to avoid any problems with modifying the dictionary as we do so
        tbTrimmed.Reverse()
        For i As Integer = 0 To tbTrimmed.Count - 1
            ccini.keys.Remove(tbTrimmed(i))
        Next
        menu(menuStr01)
        menu(tbTrimmed.Count & " lines will be removed.", "l")
        menu(menuStr02)
        Return ccini
    End Function
End Module