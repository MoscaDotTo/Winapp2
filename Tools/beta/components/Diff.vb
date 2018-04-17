Option Strict On
Imports System.IO

Module Diff

    'File handlers
    Dim oFile As IFileHandlr = New IFileHandlr(Environment.CurrentDirectory, "winapp2.ini")
    Dim nFile As IFileHandlr = New IFileHandlr(Environment.CurrentDirectory, "")
    Dim logFile As IFileHandlr = New IFileHandlr(Environment.CurrentDirectory, "diff.txt")
    Dim oldFile As iniFile
    Dim newFile As iniFile
    Dim outputToFile As String

    'Menu settings
    Dim menuTopper As String = ""
    Dim menuItemLength As Integer = 35

    'Boolean module parameters
    Dim exitCode As Boolean = False
    Dim downloadFile As Boolean = False
    Dim downloadNCC As Boolean = False
    Dim saveLog As Boolean = False
    Dim settingsChanged As Boolean = False

    'Return the default parameters to the commandline handler
    Public Sub initDiffParams(ByRef firstFile As IFileHandlr, ByRef secondFile As IFileHandlr, ByRef thirdFile As IFileHandlr, ByRef d As Boolean, ByRef dncc As Boolean, ByRef sl As Boolean)
        initDefaultSettings()
        firstFile = oFile
        secondFile = nFile
        thirdFile = logFile
        d = downloadFile
        dncc = downloadNCC
        sl = saveLog
    End Sub

    'Handle calling Diff from the commandline with 
    Public Sub remoteDiff(ByRef firstFile As IFileHandlr, secondFile As IFileHandlr, thirdFile As IFileHandlr, d As Boolean, dncc As Boolean, sl As Boolean)
        oFile = firstFile
        nFile = secondFile
        logFile = thirdFile
        downloadFile = d
        downloadNCC = dncc
        saveLog = sl
        initDiff()
    End Sub

    'Restore all the module settings to their default state
    Private Sub initDefaultSettings()
        oFile.resetParams()
        nFile.resetParams()
        logFile.resetParams()
        downloadFile = False
        downloadNCC = False
        saveLog = False
        settingsChanged = False
    End Sub

    Private Sub printMenu()
        printMenuLine(tmenu(menuTopper))
        printMenuLine(menuStr03)
        printMenuLine("This tool will output the diff between two winapp2 files", "c")
        printMenuLine(menuStr01)
        printMenuLine("Menu: Enter a number to select", "c")
        printMenuLine(menuStr01)
        printMenuLine("0. Exit                        - Return to the winapp2ool menu", "l")
        printMenuLine("1. Run (default)               - Run the diff tool", "l")
        printMenuLine("2. Run (online)                - Diff your local winapp2.ini against the latest", "l")
        printMenuLine("3. Run (online non ccleaner)   - Diff your local non-ccleaner winapp2.ini against the latest", "l")
        printMenuLine(menuStr01)
        printMenuLine("4. Toggle Log Saving           - " & If(saveLog, "Disable", "Enable") & " automatic saving of the Diff output", "l")

        If saveLog Then
            printMenuLine(menuStr01)
            printMenuLine("5. File Chooser (log)          - Change where Diff saves its log", "l")
        End If

        printMenuLine(menuStr01)
        printMenuLine("Older file: " & replDir(oFile.path), "l")
        If nFile.name <> "" Then printMenuLine("Newer file: " & replDir(nFile.path), "l")
        If settingsChanged Then
            printMenuLine(menuStr01)
            printMenuLine(If(saveLog, "6.", "5.") & " Reset Settings              - Restore the default state of the Diff settings", "l")
        End If
        printMenuLine(menuStr02)
    End Sub

    Public Sub main()
        menuTopper = "Diff"
        exitCode = False
        outputToFile = ""

        Do Until exitCode
            Console.Clear()
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
                        toggleSettingParam(saveLog, "Logging ", menuTopper, settingsChanged)
                    Case Else
                        menuTopper = invInpStr
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
    Private Function printFileLoader(ageType As String, ByRef someFile As IFileHandlr) As iniFile
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
                someFile.name = "winapp2.ini"
            Case "0"
                exitCode = True
                Return Nothing
            Case "1"
                fChooser(someFile.dir, someFile.name, exitCode, "winapp2.ini", "")
            Case Else
                someFile.name = input
        End Select

        Return validate(someFile, exitCode)
    End Function

    'load up the files to diff
    Private Sub loadFiles()
        Console.Clear()
        Try
            'Always collect the older file
            oldFile = printFileLoader("older", oFile)
            If exitCode Then Exit Sub

            'Collect the second file conditionally based on whether its a download or a local file
            newFile = If(downloadFile, getRemoteWinapp(downloadNCC), printFileLoader("newer", nFile))
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