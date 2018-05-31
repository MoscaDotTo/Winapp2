Option Strict On
Imports System.IO

Module CCiniDebug

    'File Handlers
    Dim winappFile As iniFile = New iniFile(Environment.CurrentDirectory, "winapp2.ini")
    Dim ccFile As iniFile = New iniFile(Environment.CurrentDirectory, "ccleaner.ini")
    Dim outputFile As iniFile = New iniFile(Environment.CurrentDirectory, "ccleaner.ini", "ccleaner-debugged.ini")
    
    'Menu vars
    Dim settingsChanged As Boolean = False

    'Boolean module parameters
    Dim pruneFile As Boolean = True
    Dim saveFile As Boolean = True
    Dim sortFile As Boolean = True

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
    Public Sub initCCDebugParams(ByRef firstFile As iniFile, ByRef secondFile As iniFile, ByRef thirdFile As iniFile, ByRef pf As Boolean, ByRef sa As Boolean, ByRef so As Boolean)
        initDefaultSettings()
        firstFile = winappFile
        secondFile = ccFile
        thirdFile = outputFile
        pf = pruneFile
        sa = saveFile
        so = sortFile
    End Sub

    'Handle commandline input and initalize a ccinidebug
    Public Sub remoteCC(firstfile As iniFile, secondfile As iniFile, thirdfile As iniFile, pf As Boolean, sa As Boolean, so As Boolean)
        winappFile = firstfile
        ccFile = secondfile
        outputFile = thirdfile
        pruneFile = pf
        saveFile = sa
        sortFile = so
        initDebug()
    End Sub

    Private Sub printMenu()
        printMenuTop({"Sort alphabetically the contents of ccleaner.ini and prune stale winapp2.ini settings"}, True)
        printMenuOpt("Run (default)", "Debug ccleaner.ini")

        printBlankMenuLine()
        printMenuOpt("Toggle Pruning", enStr(pruneFile) & " removal of dead winapp2.ini settings")
        printMenuOpt("Toggle Saving", enStr(saveFile) & " automatic saving of changes made by CCiniDebug")
        printMenuOpt("Toggle Sorting", enStr(sortFile) & " alphabetical sorting of ccleaner.ini")

        printBlankMenuLine()
        printMenuOpt("File Chooser (ccleaner.ini)", "Choose a new ccleaner.ini name or location")
        printIf(pruneFile, "opt", "File Chooser (winapp2.ini)", "Choose a new winapp2.ini name or location")
        printIf(saveFile, "opt", "File Chooser (save)", "Change where CCiniDebug saves its changes")

        printBlankMenuLine()
        printMenuLine("Current ccleaner.ini:  " & replDir(ccFile.path), "l")
        printIf(pruneFile, "line", "Current winapp2.ini:   " & replDir(winappFile.path), "l")
        printIf(saveFile, "line", "Current save location: " & replDir(outputFile.path), "l")
        printIf(settingsChanged, "reset", "CCiniDebug", "")
        printMenuLine(menuStr02)
    End Sub

    Private Sub resetSettings()
        initDefaultSettings()
        menuTopper = "CCiniDebug settings have been reset to their defaults"
    End Sub

    Public Sub main()
        initMenu("CCiniDebug", 35)
        Do Until exitCode
            Console.Clear()
            printMenu()
            Console.WriteLine()
            Console.Write(promptStr)
            Dim input As String = Console.ReadLine

            Try
                Dim menuNum As Integer = getMenuNumber({pruneFile, saveFile, settingsChanged}, 5)

                Select Case True
                    Case input = "0"
                        Console.WriteLine("Returning to winapp2ool menu...")
                        exitCode = True
                    Case (input = "1" Or input = "") And (pruneFile Or saveFile Or sortFile)
                        initDebug()
                    Case input = "2"
                        toggleSettingParam(pruneFile, "Pruning ", settingsChanged)
                    Case input = "3"
                        toggleSettingParam(saveFile, "Autosaving ", settingsChanged)
                    Case input = "4"
                        toggleSettingParam(sortFile, "Sorting ", settingsChanged)
                    Case input = "5"
                        changeFileParams(ccFile, settingsChanged)

                    'When the input is 6, only one of the three possible settings is true, so select the first
                    Case input = "6" And menuNum >= 5

                        Select Case True
                            Case pruneFile
                                changeFileParams(winappFile, settingsChanged)
                            Case saveFile
                                changeFileParams(outputFile, settingsChanged)
                            Case settingsChanged
                                resetSettings()
                        End Select

                    'When the input is 6, we either want to change the save parameters, or reset the settings
                    Case input = "7" And menuNum >= 7

                        Select Case True
                            Case pruneFile And saveFile
                                changeFileParams(outputFile, settingsChanged)
                            Case settingsChanged And (Not pruneFile Or Not saveFile)
                                resetSettings()
                            Case Else
                                menuTopper = invInpStr
                        End Select

                    'If the input is 7, menuNum must also be 7 and the only thing we want to do is reset the settings
                    Case input = "8" And menuNum = 8
                        resetSettings()
                    Case Not (pruneFile Or saveFile Or sortFile)
                        menuTopper = "Please enable at least one option"
                    Case Else
                        menuTopper = invInpStr
                End Select

            Catch ex As Exception
                exc(ex)
            End Try
        Loop
        revertMenu()
    End Sub

    Private Sub initDebug()
        Console.Clear()

        'Load our inifiles into memory and execute the debugging steps
        ccFile.validate()
        If exitCode Then Exit Sub
        Console.Clear()
        'Load winapp2.ini and use it to prune ccleaner.ini of stale entries
        If pruneFile Then
            winappFile.validate()
            If exitCode Then Exit Sub
            Console.Clear()
            printMenuLine(tmenu("CCiniDebug Results"))
            printMenuLine(menuStr03)
            printMenuLine("Pruning...", "c")
            printBlankMenuLine()
            prune(ccFile.sections("Options"))
            printBlankMenuLine()
        Else
            printMenuLine(tmenu("CCiniDebug Results"))
            printMenuLine(menuStr03)
        End If

        'Sort the keys
        If sortFile Then
            printMenuLine("Sorting..", "c")
            sortCC()
            printMenuLine("Sorting complete.", "l")
            printBlankMenuLine()
        End If

        'Write ccleaner.ini back to file
        If saveFile Then
            writeCCini()
            printMenuLine(outputFile.name & " saved. " & anyKeyStr, "c")
        Else
            printMenuLine("Analysis complete. " & anyKeyStr, "c")
        End If

        printMenuLine(menuStr02)

        If Not suppressOutput Then Console.ReadKey()

        'Flip the exitCode boolean
        revertMenu()
    End Sub

    Private Sub prune(ByRef optionsSec As iniSection)
        printMenuLine("Scanning " & ccFile.name & " for settings left over from removed winapp2.ini entries", "l")
        printBlankMenuLine()

        'collect the keys we must remove
        Dim tbTrimmed As New List(Of Integer)

        For i As Integer = 0 To optionsSec.keys.Count - 1
            Dim optionStr As String = optionsSec.keys.Values(i).toString

            'only operate on (app) keys
            If optionStr.StartsWith("(App)") And optionStr.Contains("*") Then
                optionStr = optionStr.Replace("(App)", "")
                optionStr = optionStr.Replace("=True", "")
                optionStr = optionStr.Replace("=False", "")
                If Not winappFile.sections.ContainsKey(optionStr) Then
                    printMenuLine(optionsSec.keys.Values(i).lineString, "l")
                    tbTrimmed.Add(optionsSec.keys.Keys(i))
                End If
            End If
        Next

        printBlankMenuLine()
        printMenuLine(tbTrimmed.Count & " orphaned settings detected", "l")

        'Remove the keys
        For Each key In tbTrimmed
            optionsSec.keys.Remove(key)
        Next

    End Sub

    Private Sub sortCC()
        'Sort the options section of ccleaner.ini 
        Dim lineList As List(Of String) = ccFile.sections("Options").getKeysAsList
        lineList.Sort()
        lineList.Insert(0, "[Options]")
        ccFile.sections("Options") = New iniSection(lineList)
    End Sub

    Private Sub writeCCini()
        'Write our changes to ccleaner.ini back to file
        Dim file As StreamWriter
        Try
            file = New StreamWriter(outputFile.path, False)
            file.Write(ccFile.toString)
            file.Close()
        Catch ex As Exception
            exc(ex)
        End Try
    End Sub
End Module