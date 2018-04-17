Option Strict On
Imports System.IO

Module CCiniDebug

    'File Handlers
    Dim winappFile As IFileHandlr = New IFileHandlr(Environment.CurrentDirectory, "winapp2.ini")
    Dim ccFile As IFileHandlr = New IFileHandlr(Environment.CurrentDirectory, "ccleaner.ini")
    Dim outputFile As IFileHandlr = New IFileHandlr(Environment.CurrentDirectory, "ccleaner.ini", "ccleaner-debugged.ini")
    Dim winappini As iniFile
    Dim ccini As iniFile

    'Menu vars
    Dim exitCode As Boolean
    Dim menuTopper As String = ""
    Dim menuItemLength As Integer = 35

    'Boolean module parameters
    Dim pruneFile As Boolean = True
    Dim saveFile As Boolean = True
    Dim sortFile As Boolean = True
    Dim settingsChanged As Boolean = False

    'Restore the default state of the module parameters
    Private Sub initDefaultSettings()
        ccFile.resetParams()
        outputFile.resetParams()
        winappFile.resetParams()
        pruneFile = True
        saveFile = True
        sortFile = True
        settingsChanged = False
    End Sub

    'Return the default module parameters to the commandline handler
    Public Sub initCCDebugParams(ByRef firstFile As IFileHandlr, ByRef secondFile As IFileHandlr, ByRef thirdFile As IFileHandlr, ByRef pf As Boolean, ByRef sa As Boolean, ByRef so As Boolean)
        initDefaultSettings()
        firstFile = winappFile
        secondFile = ccFile
        thirdFile = outputFile
        pf = pruneFile
        sa = saveFile
        so = sortFile
    End Sub

    'Handle commandline input and initalize a ccinidebug
    Public Sub remoteCC(firstfile As IFileHandlr, secondfile As IFileHandlr, thirdfile As IFileHandlr, pf As Boolean, sa As Boolean, so As Boolean)
        winappFile = firstfile
        ccFile = secondfile
        outputFile = thirdfile
        pruneFile = pf
        saveFile = sa
        sortFile = so
        initDebug()
    End Sub

    Private Sub printMenu()
        printMenuLine(tmenu(menuTopper))
        printMenuLine(menuStr03)
        printMenuLine("This tool will sort alphabetically the contents of ccleaner.ini", "c")
        printMenuLine("and can also prune stale winapp2.ini entries from it", "c")
        printMenuLine(menuStr04)
        printMenuLine("0. Exit", "Return to the winapp2ool menu", menuItemLength)
        printMenuLine("1. Run (default)", "Debug ccleaner.ini", menuItemLength)
        printMenuLine(menuStr01)
        printMenuLine("2. Toggle Pruning", If(pruneFile, "Disable", "Enable") & " removal of dead winapp2.ini settings", menuItemLength)
        printMenuLine("3. Toggle Saving", If(saveFile, "Disable", "Enable") & " automatic saving of changes made by CCiniDebug", menuItemLength)
        printMenuLine("4. Toggle Sorting", If(sortFile, "Disable", "Enable") & " alphabetical sorting of ccleaner.ini", menuItemLength)
        printMenuLine(menuStr01)
        printMenuLine("5. File Chooser (ccleaner.ini)", "Choose a new ccleaner.ini name or location", menuItemLength)

        If pruneFile Then printMenuLine("6. File Chooser (winapp2.ini)", "Choose a new winapp2.ini name or location", menuItemLength)

        If saveFile Then printMenuLine(getMenuNumber(New List(Of Boolean) From {pruneFile}, 6) & ". File Chooser (save)", "Change where CCiniDebug saves its changes", menuItemLength)

        printMenuLine(menuStr01)
        printMenuLine("Current ccleaner.ini:  " & replDir(ccFile.path), "l")
        If pruneFile Then printMenuLine("Current winapp2.ini:   " & replDir(winappFile.path), "l")
        If saveFile Then printMenuLine("Current save location: " & replDir(outputFile.path), "l")

        If settingsChanged Then
            printMenuLine(menuStr01)
            printMenuLine(getMenuNumber(New List(Of Boolean) From {pruneFile, saveFile}, 6) & ". Reset Settings", "Restore the default state of the CCiniDebug settings", menuItemLength)
        End If

        printMenuLine(menuStr02)
    End Sub

    Private Sub resetSettings()
        initDefaultSettings()
        menuTopper = "CCiniDebug settings have been reset to their defaults"
    End Sub

    Public Sub Main()
        exitCode = False
        menuTopper = "CCiniDebug"
        Do Until exitCode
            Console.Clear()
            printMenu()
            Console.WriteLine()
            Console.Write("Enter a number, or leave blank to run the default: ")
            Dim input As String = Console.ReadLine

            Try
                Dim menuNum As Integer = getMenuNumber(New List(Of Boolean) From {pruneFile, saveFile, sortFile, settingsChanged}, 4)

                Select Case True
                    Case input = "0"
                        Console.WriteLine("Returning to winapp2ool menu...")
                        exitCode = True
                    Case input = "1" Or input = ""
                        initDebug()
                    Case input = "2"
                        toggleSettingParam(pruneFile, "Pruning ", menuTopper, settingsChanged)
                    Case input = "3"
                        toggleSettingParam(saveFile, "Autosaving ", menuTopper, settingsChanged)
                    Case input = "4"
                        toggleSettingParam(sortFile, "Sorting ", menuTopper, settingsChanged)
                    Case input = "5"
                        changeFileParams(ccFile, menuTopper, settingsChanged, exitCode)

                    'When the input is 6, only one of the three possible settings is true, so select the first
                    Case input = "6" And menuNum >= 6

                        Select Case True
                            Case pruneFile
                                changeFileParams(winappFile, menuTopper, settingsChanged, exitCode)
                            Case saveFile
                                changeFileParams(outputFile, menuTopper, settingsChanged, exitCode)
                            Case settingsChanged
                                resetSettings()
                            Case Else
                                menuTopper = invInpStr
                        End Select

                    'When the input is 6, we either want to change the save parameters, or reset the settings
                    Case input = "7" And menuNum >= 7

                        Select Case True
                            Case pruneFile And saveFile
                                changeFileParams(outputFile, menuTopper, settingsChanged, exitCode)
                            Case settingsChanged And (Not pruneFile Or Not saveFile)
                                resetSettings()
                            Case Else
                                menuTopper = invInpStr
                        End Select

                    'If the input is 7, menuNum must also be 7 and the only thing we want to do is reset the settings
                    Case input = "8" And menuNum = 8
                        resetSettings()
                    Case Else
                        menuTopper = invInpStr
                End Select

            Catch ex As Exception
                exc(ex)
            End Try
        Loop
    End Sub

    Private Sub initDebug()
        Console.Clear()

        'Load our inifiles into memory and execute the debugging steps
        ccini = validate(ccFile, exitCode)
        If exitCode Then Exit Sub
        Console.Clear()

        'Load winapp2.ini and use it to prune ccleaner.ini of stale entries
        If pruneFile Then
            winappini = validate(winappFile, exitCode)
            If exitCode Then Exit Sub
            Console.Clear()

            printMenuLine(tmenu("CCiniDebug Results"))

            printMenuLine(mMenu("Pruning..."))
            prune(ccini.sections("Options"))
        Else
            printMenuLine(tmenu("CCiniDebug Results"))
        End If

        'Sort the keys
        If sortFile Then
            printMenuLine(mMenu("Sorting.."))
            sortCC()
            printMenuLine("Sorting complete.", "l")
        End If

        'Write ccleaner.ini back to file
        printMenuLine(mMenu("Finished running"))
        If saveFile Then
            writeCCini()
            printMenuLine(outputFile.name & " saved. " & anyKeyStr, "c")
        Else
            printMenuLine("Analysis complete. " & anyKeyStr, "c")
        End If

        printMenuLine(menuStr02)

        If Not suppressOutput Then Console.ReadKey()

        'Flip the exitCode boolean
        revertMenu(exitCode)
    End Sub

    Private Sub prune(ByRef optionsSec As iniSection)
        printMenuLine("Scanning " & ccFile.name & " for settings left over from removed winapp2.ini entries", "l")
        printMenuLine(menuStr01)

        'collect the keys we must remove
        Dim tbTrimmed As New List(Of Integer)

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
        printMenuLine(tbTrimmed.Count & " orphaned settings detected", "l")

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
        'Write our changes to ccleaner.ini back to file
        Dim file As StreamWriter
        Try
            file = New StreamWriter(outputFile.path, False)
            file.Write(ccini.toString)
            file.Close()
        Catch ex As Exception
            exc(ex)
        End Try
    End Sub
End Module