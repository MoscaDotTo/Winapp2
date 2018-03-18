Option Strict On
Imports System.IO

Public Module iniFileHandler

    'Strips just the key values from a list of inikeys and returns them as a list of strings
    Public Function getValuesFromKeyList(ByVal keyList As List(Of iniKey)) As List(Of String)
        Dim valList As New List(Of String)
        For Each key In keyList
            valList.Add(key.value)
        Next

        Return valList
    End Function

    'Strips just the line numbers from a list of inikeys and returns them as a list of integers
    Public Function getLineNumsFromKeyList(ByRef keyList As List(Of iniKey)) As List(Of Integer)
        Dim lineList As New List(Of Integer)
        For Each key In keyList
            lineList.Add(key.lineNumber)
        Next

        Return lineList
    End Function

    'Strips just the line numbers from the sections of an inifile and returns them as a list of integers
    Public Function getLineNumsFromSections(ByVal file As iniFile) As List(Of Integer)
        Dim outList As New List(Of Integer)
        For i As Integer = 0 To file.sections.Count - 1
            outList.Add(file.sections.Values(i).startingLineNumber)
        Next

        Return outList
    End Function

    'reorders the sections in an inifile to be in the same order as some sorted state provided to the function
    Public Sub sortIniFile(ByRef fileToBeSorted As iniFile, ByVal sortedKeys As List(Of String))
        Dim tempFile As New iniFile
        For Each entryName In sortedKeys
            tempFile.sections.Add(entryName, fileToBeSorted.sections.Item(entryName))
        Next

        fileToBeSorted = tempFile
    End Sub

    'Ensures that any call to an ini file on the system will be to a file that exists in a directory that exists
    Public Function validate(ByRef path As String, ByRef name As String, ByRef exitcode As Boolean, ByRef defName As String, ByVal defRen As String) As iniFile

        'if there's a pending exit, do that.
        If exitcode Then Return Nothing

        'Make sure both the file and the directory actually exist
        While Not File.Exists(path & name)
            If Not Directory.Exists(path) Then
                dChooser(path, exitcode)
            End If
            If exitcode Then Return Nothing
            If Not File.Exists(path & name) Then
                chkFileExist(path, name, exitcode, defName, defRen)
            End If
            If exitcode Then Return Nothing
        End While

        'Make sure that the file isn't empty
        Try
            Dim iniTester As New iniFile(path, name)
            While iniTester.sections.Count = 0
                Console.Clear()
                printMenuLine(bmenu("Empty ini file detected. Press any key to try again.", "c"))
                Console.ReadKey()
                fChooser(path, name, exitcode, defName, defRen)
                If exitcode Then Return Nothing
                iniTester = validate(path, name, exitcode, defName, defRen)
                If exitcode Then Return Nothing
            End While
            Console.Clear()
            Return iniTester
        Catch ex As Exception
            exc(ex)
            exitcode = True
            Return Nothing
        End Try
    End Function

    Public Sub chkFileExist(ByRef dir As String, ByRef name As String, ByRef exitcode As Boolean, ByVal defName As String, ByVal defRen As String)
        Dim iExitCode As Boolean = False

        While Not File.Exists(dir & name)
            If exitcode Then Exit Sub
            Console.Clear()
            printMenuLine(tmenu(name.TrimStart(CChar("\")) & " does not exist."))
            printMenuLine(menuStr03)
            While Not iExitCode
                printMenuLine("Options", "c")
                printMenuLine("Enter '0' to return to the menu", "l")
                printMenuLine("Enter '1' to specify a new directory", "l")
                printMenuLine("Leave blank to specify a new file", "l")
                printMenuLine(menuStr02)
                Dim input As String = Console.ReadLine
                Select Case input
                    Case "0"
                        iExitCode = True
                        exitcode = True
                        Exit Sub
                    Case "1"
                        dChooser(dir, exitcode)
                        If File.Exists(dir & name) Then
                            iExitCode = True
                        Else
                            Console.Clear()
                            printMenuLine(tmenu(name.TrimStart(CChar("\")) & " does not exist."))
                            printMenuLine(menuStr03)
                        End If
                    Case ""
                        fChooser(dir, name, exitcode, defName, defRen)
                        If File.Exists(dir & name) Then
                            iExitCode = True
                        Else
                            Console.Clear()
                            printMenuLine(tmenu(name.TrimStart(CChar("\")) & " does not exist."))
                            printMenuLine(menuStr03)
                        End If
                    Case Else
                        Console.Clear()
                        printMenuLine(tmenu("Invalid input. Please try again."))
                        printMenuLine(mMenu(name.TrimStart(CChar("\")) & " does not exist."))
                End Select
            End While
        End While
    End Sub

    Public Sub chkDirExist(ByRef dir As String, ByRef exitcode As Boolean)
        If exitcode Then Exit Sub
        While Not Directory.Exists(dir)
            If exitcode Then Exit Sub
            Dim iExitCode As Boolean = False

            printMenuLine(tmenu(dir & " does not exist."))
            printMenuLine("Options", "c")
            printMenuLine("Enter '0' to return to the menu", "l")
            printMenuLine("Enter '1' to create it", "l")
            printMenuLine("Leave blank to specify a new directory", "l")
            printMenuLine(menuStr02)
            While Not iExitCode
                Dim input As String = Console.ReadLine()
                Select Case input
                    Case "1"
                        Directory.CreateDirectory(dir)
                        iExitCode = True
                    Case "0"
                        dir = Environment.CurrentDirectory
                        exitcode = True
                        iExitCode = True
                        Exit Sub
                    Case ""
                        dir = Environment.CurrentDirectory
                        iExitCode = True
                        Exit Sub
                    Case Else
                        printMenuLine(tmenu("Invalid input. Please try again."))
                End Select
            End While
        End While
    End Sub

    Public Sub fChooser(ByRef dir As String, ByRef name As String, ByRef exitcode As Boolean, ByRef defaultName As String, ByRef defaultRename As String)
        If exitcode Then Exit Sub
        Console.Clear()
        printMenuLine(tmenu("File chooser"))
        printMenuLine(moMenu("Current Directory: "))
        printMenuLine(dir, "l")
        printMenuLine(menuStr01)
        printMenuLine("Current File: ", "c")
        printMenuLine(name.TrimStart(CChar("\")), "l")
        printMenuLine(menuStr03)
        printMenuLine("Options", "c")
        printMenuLine("Enter a new file name", "l")
        printMenuLine("Enter '0' to return to the menu", "l")
        Dim menuCounter As Integer = 1
        If defaultName <> "" Then
            printMenuLine("Enter '" & menuCounter & "' to use the default name (" & defaultName.TrimStart(CChar("\")) & ")", "l")
            menuCounter += 1
        End If
        If defaultRename <> "" Then
            printMenuLine("Enter '" & menuCounter & "' to use the default rename (" & defaultRename.TrimStart(CChar("\")) & ")", "l")
        End If
        printMenuLine("Enter '3' to select a new directory", "l")
        printMenuLine("Leave blank to continue using the current file name (" & name.TrimStart(CChar("\")) & ")", "l")
        printMenuLine(menuStr02)
        Dim input As String = Console.ReadLine
        Select Case True
            Case input = "0"
                exitcode = True
                Console.Clear()
                Exit Sub
            Case input = "1" And defaultName <> ""
                name = defaultName
            Case input = "1" And defaultName = ""
                name = defaultRename
            Case input = "2" And defaultRename <> ""
                name = defaultRename
            Case input = ""
                name = name
            Case input = "3"
                dChooser(dir, exitcode)
            Case Else
                name = "\" & input
        End Select
        Console.Clear()
        printMenuLine(tmenu("File Chooser"))
        Dim iexitcode As Boolean = False
        Do Until iexitcode
            printMenuLine(moMenu("Current Directory: "))
            printMenuLine(dir, "l")
            printMenuLine(menuStr01)
            printMenuLine("Current File: ", "c")
            printMenuLine(name.TrimStart(CChar("\")), "l")
            printMenuLine(menuStr03)
            printMenuLine("Options", "c")
            printMenuLine("Enter '0' to return to the menu", "l")
            printMenuLine("Enter '1' to change the file name", "l")
            printMenuLine("Enter '2' to change the directory", "l")
            printMenuLine("Leave blank to confirm file name", "l")
            printMenuLine(menuStr02)
            Console.WriteLine()
            If exitcode Then Exit Sub
            input = Console.ReadLine()
            Select Case input
                Case ""
                    iexitcode = True
                Case "0"
                    exitcode = True
                    Console.Clear()
                    iexitcode = True
                Case "1"
                    fChooser(dir, name, exitcode, defaultName, defaultRename)
                    iexitcode = True
                Case "2"
                    dChooser(dir, exitcode)
                    printMenuLine(tmenu("File Chooser"))
                Case Else
                    Console.Clear()
                    printMenuLine(tmenu("Invalid input. Please try again."))
            End Select

        Loop

    End Sub

    Public Sub dChooser(ByRef dir As String, ByRef exitCode As Boolean)
        If exitCode Then Exit Sub
        Console.Clear()
        printMenuLine(tmenu("Directory chooser"))
        printMenuLine(moMenu("Current Directory: "))
        printMenuLine(dir, "l")
        printMenuLine(menuStr03)
        printMenuLine("Options", "c")
        printMenuLine("Enter a new directory path", "l")
        printMenuLine("Enter 'parent' to move up a level", "l")
        printMenuLine("Enter '0' to return to the menu", "l")
        printMenuLine("Leave blank to use default (current folder)", "l")
        printMenuLine(menuStr02)
        Dim uPath As String = Console.ReadLine()
        Select Case uPath
            Case ""
                dir = Environment.CurrentDirectory
            Case "parent"
                dir = Directory.GetParent(dir).ToString
            Case "0"
                exitCode = True
                Exit Sub
            Case Else
                dir = uPath
                Console.Clear()
                chkDirExist(dir, exitCode)
        End Select
        Console.Clear()
        printMenuLine(tmenu("Directory chooser"))
        printMenuLine(moMenu("Current Directory: "))
        printMenuLine(dir, "l")
        printMenuLine(menuStr03)
        printMenuLine("Options", "c")
        printMenuLine("Enter '1' to change directory", "l")
        printMenuLine("Enter anything else to confirm directory change", "l")
        printMenuLine(menuStr02)
        uPath = Console.ReadLine()
        If uPath.Trim = "1" Then
            dChooser(dir, exitCode)
        End If
    End Sub

    Public Class iniFile
        Dim lineCount As Integer = 1
        Public name As String
        Public dir As String
        Public sections As New Dictionary(Of String, iniSection)
        Public comments As New Dictionary(Of Integer, iniComment)

        Public Overrides Function toString() As String
            Dim out As String = ""
            Dim i As Integer = 0
            For Each section In Me.sections.Values
                out += section.ToString
                If Not i = sections.Count - 1 Then
                    out += Environment.NewLine
                    i += 1
                End If

            Next
            Return out
        End Function

        Public Sub New()
            dir = ""
            name = ""
        End Sub

        Public Sub New(name As String)
            dir = Environment.CurrentDirectory
            Me.name = name
            createFile()
        End Sub

        Public Sub New(path As String, name As String)
            Me.name = name
            dir = path
            createFile()
        End Sub

        'This is used for constructing ini files sourced from the internet
        Public Sub New(lines As String(), name As String)
            Dim sectionToBeBuilt As New List(Of String)
            Dim lineTrackingList As New List(Of Integer)
            Dim lastLineWasEmpty As Boolean = False
            lineCount = 1
            For Each line In lines
                processiniLine(line, sectionToBeBuilt, lineTrackingList, lastLineWasEmpty, lineCount)
            Next
            If sectionToBeBuilt.Count <> 0 Then mkSection(sectionToBeBuilt, lineTrackingList)
        End Sub

        Private Sub processiniLine(ByRef currentLine As String, ByRef sectionToBeBuilt As List(Of String), ByRef lineTrackingList As List(Of Integer), ByRef lastLineWasEmpty As Boolean, ByRef lineCount As Integer)
            If currentLine.StartsWith(";") Then
                Dim newCom As New iniComment(currentLine, lineCount)

                If comments.Count > 0 Then
                    comments.Add(comments.Count, newCom)
                Else
                    comments.Add(0, newCom)
                End If
            Else
                If Not currentLine.StartsWith("[") And currentLine.Trim <> "" Then
                    'Most situations where this will arise, the file will be corrected in some manner. Currently disabled this output until an issue with the non-ccleaner ini is fixed
                    'If lastLineWasEmpty Then
                    '    Console.WriteLine("Error: Blank line detected within a section.")
                    '    Console.WriteLine("Line: " & lineCount)
                    '    Console.WriteLine()
                    'End If
                    sectionToBeBuilt.Add(currentLine)
                    lineTrackingList.Add(lineCount)
                    lastLineWasEmpty = False
                ElseIf currentLine.Trim <> "" Then

                    If Not sectionToBeBuilt.Count = 0 Then
                        mkSection(sectionToBeBuilt, lineTrackingList)
                        sectionToBeBuilt.Add(currentLine)
                        lineTrackingList.Add(lineCount)
                        lastLineWasEmpty = False
                    Else
                        sectionToBeBuilt.Add(currentLine)
                        lineTrackingList.Add(lineCount)
                        lastLineWasEmpty = False
                    End If
                Else
                    lastLineWasEmpty = True
                End If
            End If
            lineCount += 1
        End Sub

        Public Sub createFile()
            Try
                Dim r As StreamReader
                Dim sectionToBeBuilt As New List(Of String)
                Dim lineTrackingList As New List(Of Integer)
                Dim lastLineWasEmpty As Boolean = False
                r = New StreamReader(dir & "\" & name)

                Do While (r.Peek() > -1)
                    Dim currentLine As String = r.ReadLine.ToString
                    processiniLine(currentLine, sectionToBeBuilt, lineTrackingList, lastLineWasEmpty, lineCount)
                Loop
                If sectionToBeBuilt.Count <> 0 Then mkSection(sectionToBeBuilt, lineTrackingList)
                r.Close()
            Catch ex As Exception
                Console.WriteLine(ex.Message & Environment.NewLine & "Failure occurred during iniFile construction at line: " & lineCount & " in " & name)
            End Try

        End Sub

        Public Function findCommentLine(com As String) As Integer

            'find the line number of a particular comment by its string, return -1 if DNE
            For i As Integer = 0 To Me.comments.Count - 1
                If comments(i).comment.Equals(com) Then Return comments(i).lineNumber
            Next

            Return -1
        End Function

        Public Function getSectionNamesAsList() As List(Of String)
            Dim out As New List(Of String)
            For Each section In sections.Values
                out.Add(section.name)
            Next
            Return out
        End Function

        Private Sub mkSection(sectionToBeBuilt As List(Of String), lineTrackingList As List(Of Integer))
            Dim sectionHolder As New iniSection(sectionToBeBuilt, lineTrackingList)
            Try
                sections.Add(sectionHolder.name, sectionHolder)
            Catch ex As Exception

                'This will catch entries whose names are identical (case sensitive), but will not catch wholly duplicate FileKeys (etc) 
                If ex.Message = "An item with the same key has already been added." Then

                    Dim dupeInd As Integer
                    For i As Integer = 0 To Me.sections.Count - 1
                        If sections.Values(i).name = sectionToBeBuilt(0) Then
                            dupeInd = i
                            Exit For
                        End If
                    Next

                    Console.WriteLine("Error: Duplicate section name detected: " & sectionToBeBuilt(0))
                    Console.WriteLine("Line: " & lineCount)
                    Console.WriteLine("Duplicates the entry on line: " & sections.Values(dupeInd).startingLineNumber)
                    Console.WriteLine("This section will be ignored until it is given a unique name.")
                    Console.WriteLine()
                End If
            Finally
                sectionToBeBuilt.Clear()
                lineTrackingList.Clear()
            End Try
        End Sub

    End Class

    Public Class iniSection
        Public startingLineNumber As Integer
        Public endingLineNumber As Integer
        Public name As String
        Public keys As New Dictionary(Of Integer, iniKey)

        'expects that the number of parameters in keyTypeList is equal to the number of lists in the list of keylists, with the final list being an "error" list 
        'that holds keys of a type not defined in the keytypelist
        Public Sub constructKeyLists(ByVal keyTypeList As List(Of String), ByRef listOfKeyLists As List(Of List(Of iniKey)))

            For Each key In Me.keys.Values
                If keyTypeList.Contains(key.keyType.ToLower) Then
                    listOfKeyLists(keyTypeList.IndexOf(key.keyType.ToLower)).Add(key)
                Else
                    listOfKeyLists.Last.Add(key)
                End If
            Next
        End Sub

        Public Function getFullName() As String
            Return "[" & Me.name & "]"
        End Function

        Public Sub New()
            startingLineNumber = 0
            endingLineNumber = 0
            name = ""
        End Sub

        Public Sub New(ByVal listOfLines As List(Of String), listOfLineCounts As List(Of Integer))

            Dim tmp1 As String() = listOfLines(0).Split(Convert.ToChar("["))
            Dim tmp2 As String() = tmp1(1).Split(Convert.ToChar("]"))

            name = tmp2(0)

            startingLineNumber = listOfLineCounts(0)
            endingLineNumber = listOfLineCounts(listOfLineCounts.Count - 1)

            If listOfLines.Count > 1 Then
                For i As Integer = 1 To listOfLines.Count - 1
                    Dim curKey As New iniKey(listOfLines(i), listOfLineCounts(i))
                    keys.Add(i - 1, curKey)
                Next
            End If
        End Sub

        Public Sub New(ByVal listOfLines As List(Of String))

            Dim tmp1 As String() = listOfLines(0).Split(Convert.ToChar("["))
            Dim tmp2 As String() = tmp1(1).Split(Convert.ToChar("]"))

            name = tmp2(0)

            startingLineNumber = 1
            endingLineNumber = 1 + listOfLines.Count

            If listOfLines.Count > 1 Then
                For i As Integer = 1 To listOfLines.Count - 1
                    keys.Add(i - 1, New iniKey(listOfLines(i)))
                Next
            End If
        End Sub

        Public Function getKeysAsList() As List(Of String)

            Dim out As New List(Of String)

            For Each key In Me.keys.Values
               out.Add(key.toString)
            Next

            Return out
        End Function

        'returns true if the sections are the same, else returns false
        Public Function compareTo(secondSection As iniSection) As Boolean

            If keys.Count <> secondSection.keys.Count Then
                Return False
            Else
                For i As Integer = 0 To keys.Count - 1
                    If Not keys(i).compareTo(secondSection.keys(i)) Then Return False
                Next
            End If

            Return True

        End Function

        Public Overrides Function ToString() As String

            Dim out As String = Me.getFullName

            For i As Integer = 1 To Me.keys.Count
                out += Environment.NewLine & Me.keys(i - 1).toString
            Next

            out += Environment.NewLine
            Return out

        End Function
    End Class

    Public Class iniKey

        Public name As String
        Public value As String
        Public lineNumber As Integer
        Public keyType As String

        'Create an empty key
        Public Sub New()
            name = ""
            value = ""
            lineNumber = 0
            keyType = ""
        End Sub

        'Strip any numbers from the name value in a key (so numbered keys can be identified by "type")
        Private Function stripNums(keyName As String) As String
            For i As Integer = 0 To 9
                keyName = keyName.Replace(i.ToString, "")
            Next

            Return keyName
        End Function

        'Create a key with a line string from a file and line number counter 
        Public Sub New(ByVal line As String, ByVal count As Integer)

            Try
                'valid keys have the format name=value
                Dim splitLine As String() = line.Split(CChar("="))
                name = splitLine(0)
                value = splitLine(1)
                keyType = stripNums(name)
                lineNumber = count
            Catch ex As Exception
                exc(ex)
            End Try
        End Sub

        'for when trackling line numbers doesn't matter
        Public Sub New(ByVal line As String)

            Dim splitLine As String() = line.Split(CChar("="))
            name = splitLine(0)
            value = splitLine(1)
            keyType = stripNums(name)
        End Sub

        Public Overrides Function toString() As String
            'Output the key in name=value format
            Return Me.name & "=" & Me.value
        End Function

        'Output the key in Line: <line number> - name=value format
        Public Function lineString() As String
            Return "Line: " & lineNumber & " - " & name & "=" & value
        End Function

        'compare the name=value format of two different keys, return their equivalency as boolean
        Public Function compareTo(secondKey As iniKey) As Boolean
            Return Me.toString.Equals(secondKey.toString)
        End Function
    End Class

    'small wrapper class for capturing ini comment data
    Class iniComment

        Public comment As String
        Public lineNumber As Integer

        Public Sub New(c As String, l As Integer)
            comment = c
            lineNumber = l
        End Sub
    End Class
End Module