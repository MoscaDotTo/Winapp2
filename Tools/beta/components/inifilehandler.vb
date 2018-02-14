Option Strict On
Imports System.IO

Public Module iniFileHandler

    Public Sub validate(ByRef path As String, ByRef name As String, ByRef exitcode As Boolean, ByRef defName As String, ByVal defRen As String)
        If exitcode Then
            Exit Sub
        End If
        While Not File.Exists(path & name)
            If Not Directory.Exists(path) Then
                dChooser(path, exitcode)
            End If
            If exitcode Then
                Exit Sub
            End If
            If Not File.Exists(path & name) Then
                chkFileExist(path, name, exitcode, defName, defRen)
            End If
            If exitcode Then
                Exit Sub
            End If
        End While
        Try
            Dim iniTester As New iniFile(path, name)
            While iniTester.sections.Count = 0
                Console.Clear()
                bmenu("Empty ini file detected. Press any key to try again.", "c")
                Console.ReadKey()
                fChooser(path, name, exitcode, defName, defRen)
                If exitcode Then
                    Exit Sub
                End If
                validate(path, name, exitcode, defName, defRen)
                If exitcode Then
                    Exit Sub
                End If
                iniTester = New iniFile(path, name)

            End While
        Catch ex As Exception

        End Try
        Console.Clear()
    End Sub

    Public Sub chkFileExist(ByRef dir As String, ByRef name As String, ByRef exitcode As Boolean, ByVal defName As String, ByVal defRen As String)
        Dim iExitCode As Boolean = False

        While Not File.Exists(dir & name)
            If exitcode Then
                Exit Sub
            End If
            Console.Clear()
            tmenu(name.Split(CChar("\"))(1) & " does not exist.")
            menu(menuStr03)
            While Not iExitCode
                menu("Options", "c")
                menu("Enter '0' to return to the menu", "l")
                menu("Enter '1' to specify a new directory", "l")
                menu("Leave blank to specify a new file", "l")
                menu(menuStr02, "")
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
                            tmenu(name.Split(CChar("\"))(1) & " does not exist.")
                            menu(menuStr03)
                        End If
                    Case ""
                        fChooser(dir, name, exitcode, defName, defRen)
                        If File.Exists(dir & name) Then
                            iExitCode = True
                        Else
                            Console.Clear()
                            tmenu(name.Split(CChar("\"))(1) & " does not exist.")
                            menu(menuStr03)
                        End If
                    Case Else
                        Console.Clear()
                        tmenu("Invalid input. Please try again.")
                        mMenu(name.Split(CChar("\"))(1) & " does not exist.")
                End Select
            End While
        End While
    End Sub

    Public Sub chkDirExist(ByRef dir As String, ByRef exitcode As Boolean)
        While Not Directory.Exists(dir)
            Dim iExitCode As Boolean = False

            tmenu(dir & " does not exist.")
            menu("Options", "c")
            menu("Enter '0' to return to the menu", "l")
            menu("Enter '1' to create it", "l")
            menu("Leave blank to specify a new directory", "l")
            menu(menuStr02, "")
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
                        tmenu("Invalid input. Please try again.")
                End Select
            End While
        End While
    End Sub

    Public Sub fChooser(ByRef dir As String, ByRef name As String, ByRef exitcode As Boolean, ByRef defaultName As String, ByRef defaultRename As String)
        Console.Clear()
        tmenu("File chooser")
        moMenu("Current Directory: ")
        menu(dir, "l")
        menu(menuStr01)
        menu("Current File: ", "c")
        menu(name.Split(CChar("\"))(1), "l")
        menu(menuStr03)
        menu("Options", "c")
        menu("Enter a new file name", "l")
        menu("Enter '0' to return to the menu", "l")
        Dim menuCounter As Integer = 1
        If defaultName <> "" Then
            menu("Enter '" & menuCounter & "' to use the default name (" & defaultName.Split(CChar("\"))(1) & ")", "l")
            menuCounter += 1
        End If
        If defaultRename <> "" Then
            menu("Enter '" & menuCounter & "' to use the default rename (" & defaultRename.Split(CChar("\"))(1) & ")", "l")
        End If
        menu("Enter '3' to select a new directory", "l")
        menu("Leave blank to continue using the current file name (" & name.Split(CChar("\"))(1) & ")", "l")
        menu(menuStr02)
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
        tmenu("File Chooser")
        Dim iexitcode As Boolean = False
        Do Until iexitcode
            moMenu("Current Directory: ")
            menu(dir, "l")
            menu(menuStr01)
            menu("Current File: ", "c")
            menu(name.Split(CChar("\"))(1), "l")
            menu(menuStr03)
            menu("Options", "c")
            menu("Enter '0' to return to the menu", "l")
            menu("Enter '1' to change the file name", "l")
            menu("Enter '2' to change the directory", "l")
            menu("Leave blank to confirm file name", "l")
            menu(menuStr02)
            Console.WriteLine()
            If exitcode Then
                Exit Sub
            End If
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
                    tmenu("File Chooser")
                Case Else
                    Console.Clear()
                    tmenu("Invalid input. Please try again.")
            End Select

        Loop

    End Sub

    Public Sub dChooser(ByRef dir As String, ByRef exitCode As Boolean)
        Console.Clear()
        tmenu("Directory chooser")
        moMenu("Current Directory: ")
        menu(dir, "l")
        menu(menuStr03)
        menu("Options", "c")
        menu("Enter a new directory path", "l")
        menu("Enter 'parent' to move up a level", "l")
        menu("Enter '0' to return to the menu", "l")
        menu("Leave blank to use default (current folder)", "l")
        menu(menuStr02, "")
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
        tmenu("Directory chooser")
        moMenu("Current Directory: ")
        menu(dir, "l")
        menu(menuStr03)
        menu("Options", "c")
        menu("Enter '1' to change directory", "l")
        menu("Enter anything else to confirm directory change", "l")
        menu(menuStr02)
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
            For Each section As iniSection In Me.sections.Values
                out += section.ToString
                If Not i = Me.sections.Count - 1 Then
                    out += Environment.NewLine
                    i += 1
                End If

            Next
            Return out
        End Function

        Public Sub New()
            Me.dir = ""
            Me.name = ""
        End Sub

        Public Sub New(name As String)
            Me.dir = Environment.CurrentDirectory
            Me.name = name
            createFile()
        End Sub

        Public Sub New(path As String, name As String)
            Me.name = name
            Me.dir = path
            createFile()
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

                    If currentLine.StartsWith(";") Then
                        Dim newCom As New iniComment(currentLine, lineCount)
                        If comments.Count > 0 Then
                            comments.Add(comments.Count, newCom)
                        Else
                            comments.Add(0, newCom)
                        End If
                    Else
                        If Not currentLine.StartsWith("[") And currentLine.Trim <> "" Then
                            If lastLineWasEmpty Then
                                Console.WriteLine("Error: Blank line detected within a section.")
                                Console.WriteLine("Line: " & lineCount)
                                Console.WriteLine()
                            End If
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
                Loop
                If sectionToBeBuilt.Count <> 0 Then
                    mkSection(sectionToBeBuilt, lineTrackingList)
                End If
                r.Close()
            Catch ex As Exception
                Console.WriteLine(ex.Message & Environment.NewLine & "Failure occurred during iniFile construction at line: " & lineCount & " in " & name)
            End Try

        End Sub

        Public Function findCommentLine(com As String) As Integer

            'find the line number of a particular comment by its string, return -1 if DNE
            For i As Integer = 0 To Me.comments.Count - 1
                If Me.comments(i).comment.Equals(com) Then
                    Return Me.comments(i).lineNumber
                End If
            Next

            Return -1
        End Function

        Public Function getSectionNamesAsList() As List(Of String)
            Dim out As New List(Of String)
            For Each section As iniSection In Me.sections.Values
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
                        If Me.sections.Values(i).name = sectionToBeBuilt(0) Then
                            dupeInd = i
                            Exit For
                        End If
                    Next

                    Console.WriteLine("Error: Duplicate section name detected: " & sectionToBeBuilt(0))
                    Console.WriteLine("Line: " & lineCount)
                    Console.WriteLine("Duplicates the entry on line: " & Me.sections.Values(dupeInd).startingLineNumber)
                    Console.WriteLine("This section will be ignored until it is given a unique name.")
                    Console.WriteLine()
                End If
            Finally
                sectionToBeBuilt.Clear()
                lineTrackingList.Clear()
            End Try
        End Sub

        Public Function getStreamOfComments(startNum As Integer, endNum As Integer) As String
            Dim out As String = ""
            Try
                While Not startNum = endNum And startNum <= endNum
                    out += Me.comments(startNum).comment & Environment.NewLine
                    startNum += 1
                End While
                out += Me.comments(startNum).comment
            Catch ex As Exception
                Console.WriteLine(ex.ToString)
                Console.ReadKey()
            End Try
            Return out
        End Function

        Private Function getDiff(section As iniSection, changeType As String) As String
            Dim out As String = ""
            out += mkMenuLine(section.name & " has been " & changeType, "c") & Environment.NewLine
            out += mkMenuLine(menuStr02, "") & Environment.NewLine & Environment.NewLine
            out += section.ToString & Environment.NewLine
            out += menuStr00
            Return out
        End Function

        Public Function compareTo(secondFile As iniFile) As List(Of String)

            Dim outList As New List(Of String)
            Dim comparedList As New List(Of String)

            For i As Integer = 0 To Me.sections.Count - 1
                Dim curSection As iniSection = Me.sections.Values(i)
                Dim curName As String = curSection.name
                Try
                    If secondFile.sections.Keys.Contains(curName) And Not comparedList.Contains(curName) Then
                        Dim sSection As iniSection = secondFile.sections(curName)
                        If Not curSection.compareTo(sSection) Then
                            outList.Add(getDiff(sSection, "modified."))
                        End If
                        comparedList.Add(curName)
                    ElseIf Not secondFile.sections.Keys.Contains(curName) Then
                        outList.Add(getDiff(curSection, "removed."))
                        comparedList.Add(curName)
                    End If
                Catch ex As Exception
                    Console.WriteLine("Error: " & ex.ToString)
                    Console.WriteLine("Please report this error on GitHub")
                    Console.WriteLine()
                End Try
            Next

            For i As Integer = 0 To secondFile.sections.Count - 1
                Dim curSection As iniSection = secondFile.sections.Values(i)
                Dim curName As String = curSection.name

                If Not Me.sections.Keys.Contains(curName) Then
                    outList.Add(getDiff(curSection, "added."))
                End If
            Next

            Return outList
        End Function
    End Class

    Public Class iniSection
        Public startingLineNumber As Integer
        Public endingLineNumber As Integer
        Public name As String
        Public keys As New Dictionary(Of Integer, iniKey)

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
            Me.name = tmp2(0)
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
            Me.name = tmp2(0)
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
            For Each key As iniKey In Me.keys.Values
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
                    If Not keys(i).compareTo(secondSection.keys(i)) Then
                        Return False
                    End If
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

        Public Sub New()

            name = ""
            value = ""
            lineNumber = 0
            keyType = ""
        End Sub

        Private Function stripNums(keyName As String) As String

            For i As Integer = 0 To 9
                keyName = keyName.Replace(i.ToString, "")
            Next
            Return keyName
        End Function

        Public Sub New(ByVal line As String, ByVal count As Integer)

            Try
                Dim splitLine As String() = line.Split(Convert.ToChar("="))
                name = splitLine(0)
                value = splitLine(1)
                keyType = stripNums(name)
                lineNumber = count
            Catch ex As Exception
                Console.WriteLine(ex)
            End Try
        End Sub

        'for when trackling line numbers doesn't matter
        Public Sub New(ByVal line As String)
            Dim splitLine As String() = line.Split(Convert.ToChar("="))
            name = splitLine(0)
            value = splitLine(1)
            keyType = stripNums(name)
        End Sub

        Public Overrides Function toString() As String
            Return Me.name & "=" & Me.value
        End Function

        Public Function lineString() As String
            Return "Line: " & lineNumber & " - " & name & "=" & value
        End Function

        Public Function compareTo(secondKey As iniKey) As Boolean
            Return Me.toString.Equals(secondKey.toString)
        End Function
    End Class

    Class iniComment

        Public comment As String
        Public lineNumber As Integer

        Public Sub New(c As String, l As Integer)
            comment = c
            lineNumber = l
        End Sub
    End Class
End Module