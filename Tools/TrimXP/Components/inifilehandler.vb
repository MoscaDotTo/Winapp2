Option Strict On
Imports System.IO
Imports System.Text.RegularExpressions

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

    'reorders the sections in an inifile to be in the same order as some sorted state provided to the function
    Public Sub sortIniFile(ByRef fileToBeSorted As iniFile, ByVal sortedKeys As List(Of String))
        Dim tempFile As New iniFile
        For Each entryName In sortedKeys
            tempFile.sections.Add(entryName, fileToBeSorted.sections.Item(entryName))
        Next
        fileToBeSorted = tempFile
    End Sub

    Public Sub chkFileExist(someFile As iniFile)
        Dim iExitCode As Boolean = False

        While Not File.Exists(someFile.path)
            If exitCode Then Exit Sub
            menuTopper = "Error"
            While Not iExitCode
                Console.Clear()
                printMenuTop({someFile.name & " does not exist."}, True)
                printMenuOpt("File Chooser (default)", "Change the file name")
                printMenuOpt("Directory Chooser", "Change the directory")
                printMenuLine(menuStr02)
                Dim input As String = Console.ReadLine
                Select Case input
                    Case "0"
                        iExitCode = True
                        exitCode = True
                        Exit Sub
                    Case "1", ""
                        fileChooser(someFile)
                    Case "2"
                        dirChooser(someFile.dir)
                    Case Else
                        menuTopper = invInpStr
                End Select
                If Not File.Exists(someFile.path) And Not menuTopper = invInpStr Then menuTopper = "Error"
            End While
        End While
    End Sub

    Public Sub chkDirExist(ByRef dir As String)
        If exitCode Then Exit Sub
        While Not Directory.Exists(dir)
            If exitCode Then Exit Sub
            Dim iExitCode As Boolean = False
            menuTopper = "Error"

            While Not iExitCode
                printMenuTop({dir & "does not exist."}, True)
                printMenuOpt("Create Directory", "Create this directory")
                printMenuOpt("Directory Chooser (default)", "Specify a new directory")
                printMenuLine(menuStr02)
                Dim input As String = Console.ReadLine()
                Select Case input
                    Case "0"
                        dir = Environment.CurrentDirectory
                        exitCode = True
                        iExitCode = True
                        Exit Sub
                    Case "1"
                        Directory.CreateDirectory(dir)
                    Case "2", ""
                        dirChooser(dir)
                    Case Else
                        menuTopper = invInpStr
                End Select
                If Not Directory.Exists(dir) And Not menuTopper = invInpStr Then menuTopper = "Error"
            End While
        End While
    End Sub

    Public Sub fileChooser(ByRef someFile As iniFile)
        If exitCode Then Exit Sub
        Console.Clear()
        menuTopper = "File Chooser"
        printMenuTop({"Choose a file name, or open the directory chooser to choose a directory"}, True)
        printIf(someFile.initName <> "", "opt", someFile.initName, "Use the default name")
        printIf(someFile.secondName <> "", "opt", someFile.secondName, "Use the default rename")
        printMenuOpt("Directory Chooser", "Choose a new directory")
        printBlankMenuLine()
        printMenuLine("Current Directory: " & replDir(someFile.dir), "l")
        printMenuLine("Current File:      " & someFile.name, "l")
        printMenuLine(menuStr02)
        Console.Write("Enter a number, a new file name, or leave blank to continue using '" & someFile.name & "': ")
        Dim input As String = Console.ReadLine
        Select Case True
            Case input = "0"
                exitCode = True
                Console.Clear()
                Exit Sub
            Case input = ""
            Case input = "1" And someFile.initName <> ""
                someFile.name = someFile.initName
            Case input = "1" And someFile.initName = ""
                someFile.name = someFile.secondName
            Case input = "2" And someFile.secondName <> ""
                someFile.name = someFile.secondName
            Case input = "2" And someFile.secondName = ""
                dirChooser(someFile.dir)
            Case input = "3" And someFile.initName <> "" And someFile.secondName <> ""
                dirChooser(Dir)
            Case Else
                someFile.name = input
        End Select
        If exitCode Then Exit Sub
        Dim iExitCode As Boolean = False
        menuTopper = "File Chooser"
        Do Until iExitCode
            Console.Clear()
            printMenuTop({"Confirm your settings or return to the options to change them."}, True)
            printMenuOpt("File Chooser", "Change the file name")
            printMenuOpt("Directory Chooser", "Change the directory")
            printMenuOpt("Confirm (default)", "Save changes")
            printBlankMenuLine()
            printMenuLine("Current Directory: " & replDir(someFile.dir), "l")
            printMenuLine("Current File:      " & someFile.name, "l")
            printMenuLine(menuStr02)
            Console.WriteLine()
            Console.Write("Enter a number, or leave blank to run the default: ")
            input = Console.ReadLine()
            Select Case input
                Case "0"
                    exitCode = True
                    iExitCode = True
                Case "1"
                    fileChooser(someFile)
                    iExitCode = True
                Case "2"
                    dirChooser(someFile.dir)
                    iExitCode = True
                Case "3", ""
                    iExitCode = True
                Case Else
                    menuTopper = invInpStr
            End Select
        Loop
    End Sub

    Public Sub dirChooser(ByRef dir As String)
        If exitCode Then Exit Sub
        Console.Clear()
        menuTopper = "Directory Chooser"
        printMenuTop({"Choose a directory"}, True)
        printMenuOpt("Use default (default)", "Use the same folder as winapp2ool.exe")
        printMenuOpt("Parent Folder", "Go up a level")
        printMenuOpt("Current folder", "Continue using the same folder as below")
        printBlankMenuLine()
        printMenuLine("Current Directory: " & dir, "l")
        printMenuLine(menuStr02)
        Console.WriteLine()
        Console.Write("Choose a number from above, enter a new directory, or leave blank to run the default: ")
        Dim input As String = Console.ReadLine()
        Select Case input
            Case "0"
                exitCode = True
            Case "1", ""
                dir = Environment.CurrentDirectory
            Case "2"
                dir = Directory.GetParent(dir).ToString
            Case "3"
            Case Else
                dir = input
                Console.Clear()
                chkDirExist(dir)
        End Select
        If exitCode Then Exit Sub
        Console.Clear()
        menuTopper = "Directory Chooser"
        Dim iExitCode As Boolean = False
        Do Until iExitCode
            If exitCode Then Exit Sub
            printMenuTop({"Choose a directory"}, True)
            printMenuOpt("Directory Chooser", "Change the directory")
            printMenuOpt("Confirm (default)", "Use this directory")
            printBlankMenuLine()
            printMenuLine("Current Directory: " & dir, "l")
            printMenuLine(menuStr02)
            Console.Write("Choose a number from above, or leave blank to run the default: ")

            input = Console.ReadLine()
            Select Case input
                Case "0"
                    iExitCode = True
                    exitCode = True
                Case "1"
                    dirChooser(dir)
                    iExitCode = True
                Case "2", ""
                    iExitCode = True
                Case Else
                    menuTopper = invInpStr
            End Select
        Loop
    End Sub

    Public Class iniFile
        Dim lineCount As Integer = 1

        'The current state of the directory & name of the file
        Public dir As String
        Public name As String

        'The inital state of the direcotry & name of the file (for restoration purposes) 
        Public initDir As String
        Public initName As String

        'Suggested rename for output files
        Public secondName As String

        'Sections will be initally stored in the order they're read
        Public sections As New Dictionary(Of String, iniSection)

        'Any line comments will be saved in the order they're read 
        Public comments As New Dictionary(Of Integer, iniComment)

        Public Overrides Function toString() As String
            Dim out As String = ""

            For i As Integer = 0 To sections.Count - 2
                out += sections.Values(i).ToString & Environment.NewLine
            Next
            out += sections.Values.Last.ToString

            Return out
        End Function

        Public Sub New()
            dir = ""
            name = ""
            initDir = ""
            initName = ""
            secondName = ""
        End Sub

        Public Sub New(directory As String, filename As String)
            dir = directory
            name = filename
            initDir = directory
            initName = filename
            secondName = ""
        End Sub

        Public Sub New(directory As String, filename As String, rename As String)
            dir = directory
            name = filename
            initDir = directory
            initName = filename
            secondName = rename
        End Sub

        'Restore the parameters used to initalize the inifile
        Public Sub resetParams()
            dir = initDir
            name = initName
        End Sub

        'Output the filepath as a string
        Public Function path() As String
            Return dir & "\" & name
        End Function

        'Strips just the line numbers from the sections of an inifile and returns them as a list of integers
        Public Function getLineNumsFromSections() As List(Of Integer)
            Dim outList As New List(Of Integer)
            For Each section In sections.Values
                outList.Add(section.startingLineNumber)
            Next

            Return outList
        End Function

        'This is used for constructing ini files sourced from the internet
        Public Sub New(lines As String(), name As String)
            Dim sectionToBeBuilt As New List(Of String)
            Dim lineTrackingList As New List(Of Integer)
            Dim lastLineWasEmpty As Boolean = False
            lineCount = 1
            For Each line In lines
                processiniLine(line, sectionToBeBuilt, lineTrackingList, lastLineWasEmpty)
            Next
            If sectionToBeBuilt.Count <> 0 Then mkSection(sectionToBeBuilt, lineTrackingList)
        End Sub

        Private Sub processiniLine(ByRef currentLine As String, ByRef sectionToBeBuilt As List(Of String), ByRef lineTrackingList As List(Of Integer), ByRef lastLineWasEmpty As Boolean)
            Select Case True
                Case currentLine.StartsWith(";")
                    Dim newCom As New iniComment(currentLine, lineCount)
                    comments.Add(comments.Count, newCom)
                Case Not currentLine.StartsWith("[") And Not currentLine.Trim = ""
                    sectionToBeBuilt.Add(currentLine)
                    lineTrackingList.Add(lineCount)
                    lastLineWasEmpty = False
                Case currentLine.Trim <> ""
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
                Case Else
                    lastLineWasEmpty = True
            End Select
            lineCount += 1
        End Sub

        Public Sub init()
            Try
                Dim sectionToBeBuilt As New List(Of String)
                Dim lineTrackingList As New List(Of Integer)
                Dim lastLineWasEmpty As Boolean = False
                Dim r As New StreamReader(Me.path())
                Do While (r.Peek() > -1)
                    processiniLine(r.ReadLine.ToString, sectionToBeBuilt, lineTrackingList, lastLineWasEmpty)
                Loop
                If sectionToBeBuilt.Count <> 0 Then mkSection(sectionToBeBuilt, lineTrackingList)
                r.Close()
            Catch ex As Exception
                Console.WriteLine(ex.Message & Environment.NewLine & "Failure occurred during iniFile construction at line: " & lineCount & " in " & name)
            End Try
        End Sub

        'Ensures that any call to an ini file on the system will be to a file that exists in a directory that exists
        Public Sub validate()

            'if there's a pending exit, do that.
            If exitCode Then Exit Sub

            'Make sure both the file and the directory actually exist
            While Not File.Exists(path())
                chkDirExist(dir)
                If exitCode Then Exit Sub
                chkFileExist(Me)
                If exitCode Then Exit Sub
            End While

            'Make sure that the file isn't empty
            Try
                Dim iniTester As New iniFile(dir, name)
                iniTester.init()
                Dim clearAtEnd As Boolean = False
                While iniTester.sections.Count = 0
                    clearAtEnd = True
                    Console.Clear()
                    printMenuLine(bmenu("Empty ini file detected. Press any key to try again.", "c"))
                    Console.ReadKey()
                    fileChooser(iniTester)
                    If exitCode Then Exit Sub
                    iniTester.validate()
                    If exitCode Then Exit Sub
                End While
                sections = iniTester.sections
                comments = iniTester.comments
                If clearAtEnd Then Console.Clear()
            Catch ex As Exception
                exc(ex)
                exitCode = True
            End Try
        End Sub

        'find the line number of a particular comment by its string, return -1 if DNE
        Public Function findCommentLine(com As String) As Integer
            For Each comment In comments.Values
                If comment.comment = com Then Return comment.lineNumber
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

                    Dim lineErr As Integer
                    For Each section In sections.Values
                        If section.name = sectionToBeBuilt(0) Then
                            lineErr = section.startingLineNumber
                            Exit For
                        End If
                    Next

                    Console.WriteLine("Error: Duplicate section name detected: " & sectionToBeBuilt(0))
                    Console.WriteLine("Line: " & lineCount)
                    Console.WriteLine("Duplicates the entry on line: " & lineErr)
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

            Dim tmp1 As String() = listOfLines(0).Split(CChar("["))
            Dim tmp2 As String() = tmp1(1).Split(CChar("]"))

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

            Dim tmp1 As String() = listOfLines(0).Split(CChar("["))
            Dim tmp2 As String() = tmp1(1).Split(CChar("]"))

            name = tmp2(0)

            startingLineNumber = 1
            endingLineNumber = 1 + listOfLines.Count

            If listOfLines.Count > 1 Then
                For i As Integer = 1 To listOfLines.Count - 1
                    keys.Add(i - 1, New iniKey(listOfLines(i), 0))
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
        Public Function compareTo(ss As iniSection, ByRef removedKeys As List(Of iniKey), ByRef addedKeys As List(Of iniKey), ByRef updatedKeys As List(Of KeyValuePair(Of iniKey, iniKey))) As Boolean

            'Create a copy of the section so we can modify it
            Dim secondSection As New iniSection
            secondSection.name = ss.name
            secondSection.startingLineNumber = ss.startingLineNumber
            For i As Integer = 0 To ss.keys.Count - 1
                secondSection.keys.Add(i, ss.keys.Values(i))
            Next

            Dim noMatch As Boolean
            Dim tmpList As New List(Of Integer)

            For Each key In keys.Values
                noMatch = True
                For i As Integer = 0 To secondSection.keys.Values.Count - 1
                    If key.keyType.ToLower = secondSection.keys.Values(i).keyType.ToLower And key.value.ToLower = secondSection.keys.Values(i).value.ToLower Then
                        noMatch = False
                        tmpList.Add(i)
                        Exit For
                    End If
                Next
                If noMatch Then
                    'If the key isn't found in the second (newer) section, consider it removed for now
                    removedKeys.Add(key)
                End If
            Next

            'Remove all matched keys
            tmpList.Reverse()
            For Each ind In tmpList
                secondSection.keys.Remove(ind)
            Next

            'Assume any remaining keys have been added
            For Each key In secondSection.keys.Values
                addedKeys.Add(key)
            Next

            'Check for keys whose names match
            Dim rkTemp, akTemp As New List(Of iniKey)
            rkTemp = removedKeys.ToList
            akTemp = addedKeys.ToList
            For Each key In removedKeys
                For Each skey In addedKeys
                    If key.name.ToLower = skey.name.ToLower Then

                        Dim oldKey As New winapp2KeyParameters(key)
                        Dim newKey As New winapp2KeyParameters(skey)
                        oldKey.argsList.Sort()
                        newKey.argsList.Sort()

                        If oldKey.argsList.Count = newKey.argsList.Count Then

                            For i As Integer = 0 To oldKey.argsList.Count - 1
                                If Not oldKey.argsList(i).ToLower = newKey.argsList(i).ToLower Then
                                    updatedKeys.Add(New KeyValuePair(Of iniKey, iniKey)(key, skey))
                                    rkTemp.Remove(key)
                                    akTemp.Remove(skey)
                                    Exit For
                                End If
                            Next
                            rkTemp.Remove(key)
                            akTemp.Remove(skey)
                        Else
                            updatedKeys.Add(New KeyValuePair(Of iniKey, iniKey)(key, skey))
                            rkTemp.Remove(key)
                            akTemp.Remove(skey)
                        End If
                    End If
                Next
            Next

            'Update the lists
            addedKeys = akTemp
            removedKeys = rkTemp

            Return removedKeys.Count + addedKeys.Count + updatedKeys.Count = 0

        End Function

        Public Overrides Function ToString() As String

            Dim out As String = Me.getFullName

            For Each key In keys.Values
                out += Environment.NewLine & key.toString
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

        Public Function nameIs(n As String) As Boolean
            Return name = n
        End Function

        Public Function typeIs(t As String) As Boolean
            Return keyType = t
        End Function

        'Return whether or not a key starts with or ends with a value
        Public Function vStartsOrEndsWith(txt As String) As Boolean
            Return value.StartsWith(txt) Or value.EndsWith(txt)
        End Function

        Public Function nStartsOrEndsWith(txt As String) As Boolean
            Return name.StartsWith(txt) Or name.EndsWith(txt)
        End Function

        'Return whether or not the key's value contains a given string
        Public Function vHas(txt As String, toLower As Boolean) As Boolean
            Return If(toLower, value.ToLower.Contains(txt), value.Contains(txt))
        End Function

        Public Function vHas(txt As String) As Boolean
            Return vHas(txt, False)
        End Function

        Public Function vHasAny(txts As String()) As Boolean
            Return vHasAny(txts, False)
        End Function

        Public Function vHasAny(txts As String(), toLower As Boolean) As Boolean
            For Each txt In txts
                If vHas(txt, toLower) Then Return True
            Next
            Return False
        End Function

        'Return whether or not the key's value is equal to a given string
        Public Function vIs(txt As String) As Boolean
            Return vIs(txt, False)
        End Function

        Public Function vIs(txt As String, toLower As Boolean) As Boolean
            Return If(toLower, value.ToLower.Equals(txt), value.Equals(txt))
        End Function

        'Returns true if a key's value doesn't contain any in a given array of strings
        Public Function vHasnt(strs As String(), toLower As Boolean) As Boolean
            For Each txt In strs
                If vHas(txt, toLower) Then Return False
            Next
            Return True
        End Function

        'Strip any numbers from the name value in a key (so numbered keys can be identified by "type")
        Private Function stripNums(keyName As String) As String
            Return New Regex("[\d]").Replace(keyName, "")
        End Function

        'Create a key with a line string from a file and line number counter 
        Public Sub New(ByVal line As String, ByVal count As Integer)

            'valid keys have the format name=value
            Try
                Dim splitLine As String() = line.Split(CChar("="))
                lineNumber = count
                Select Case True
                    Case splitLine(0) <> "" And splitLine(1) <> ""
                        name = splitLine(0)
                        value = splitLine(1)
                        keyType = stripNums(name)
                    Case splitLine(0) = "" And splitLine(1) <> ""
                        name = "KeyTypeNotGiven"
                        value = splitLine(1)
                        keyType = "Error"
                    Case splitLine(0) <> "" And splitLine(1) = ""
                        name = "DeleteMe"
                        value = "This key was not provided with a value and will be deleted. The user should never see this, if you do, please report it as a bug on GitHub"
                        keyType = "DeleteMe"
                End Select
            Catch ex As Exception
                exc(ex)
            End Try
        End Sub

        'Output the key in name=value format
        Public Overrides Function toString() As String
            Return name & "=" & value
        End Function

        'Output the key in Line: <line number> - name=value format
        Public Function lineString() As String
            Return "Line: " & lineNumber & " - " & name & "=" & value
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