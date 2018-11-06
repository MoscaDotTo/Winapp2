'    Copyright (C) 2018 Robbie Ward
' 
'    This file is a part of Winapp2ool
' 
'    Winapp2ool is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    Winap2ool is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with Winapp2ool.  If not, see <http://www.gnu.org/licenses/>.
Option Strict On
Imports System.IO
Imports System.Text.RegularExpressions

''' <summary>
''' A collection of methods and classes for interacting with .ini files 
''' </summary>
Public Module iniFileHandler

    ''' <summary>
    ''' Returns the values from a list of iniKey objects as a list of strings
    ''' </summary>
    ''' <param name="keyList">A list of iniKey objects to parse values from</param>
    ''' <returns></returns>
    Public Function getValuesFromKeyList(ByVal keyList As List(Of iniKey)) As List(Of String)
        Dim valList As New List(Of String)
        For Each key In keyList
            valList.Add(key.value)
        Next
        Return valList
    End Function

    ''' <summary>
    ''' Returns the line numbers from a list of iniKey objects as a list of integers
    ''' </summary>
    ''' <param name="keyList">A list of iniKey objects to parse line numbers from</param>
    ''' <returns></returns>
    Public Function getLineNumsFromKeyList(ByRef keyList As List(Of iniKey)) As List(Of Integer)
        Dim lineList As New List(Of Integer)
        For Each key In keyList
            lineList.Add(key.lineNumber)
        Next
        Return lineList
    End Function

    ''' <summary>
    ''' Adds an iniKey to a list of iniKeys if the given condition is met
    ''' </summary>
    ''' <param name="key">The key to be added</param>
    ''' <param name="keyList">The list to be added to</param>
    ''' <param name="cond">The condition under which the key should be added to the list</param>
    Public Sub addKeyToListIf(ByRef key As iniKey, ByRef keyList As List(Of iniKey), cond As Boolean)
        If cond Then keyList.Add(key)
    End Sub

    ''' <summary>
    ''' Enforces that a user selected file exists
    ''' </summary>
    ''' <param name="someFile">An iniFile object with user defined path and name parameters</param>
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

    ''' <summary>
    ''' Enforces that a user defined directory exists, either by selecting a new one or creating one.
    ''' </summary>
    ''' <param name="dir">A user defined windows directory</param>
    Public Sub chkDirExist(ByRef dir As String)
        If exitCode Then Exit Sub
        If exitCode Then Exit Sub
        Dim iExitCode As Boolean = False
        menuTopper = "Error"

        While Not iExitCode
            Console.Clear()
            printMenuTop({dir & " does not exist."}, True)
            printMenuOpt("Create Directory", "Create this directory")
            printMenuOpt("Directory Chooser (default)", "Specify a new directory")
            printMenuLine(menuStr02)
            Dim input As String = Console.ReadLine()
            Select Case input
                Case "0"
                    dir = Environment.CurrentDirectory
                    exitCode = iExitCode = True
                    Exit Sub
                Case "1"
                    Directory.CreateDirectory(dir)
                Case "2", ""
                    dirChooser(dir)
                Case Else
                    menuTopper = invInpStr
            End Select
            'If the directory still doesn't exist and the last input wasn't valid, 
            If Not Directory.Exists(dir) And Not menuTopper = invInpStr Then menuTopper = "Error"
            If Directory.Exists(dir) Then iExitCode = True
        End While
    End Sub

    ''' <summary>
    ''' Presents a menu to the user allowing them to perform some file and directory operations
    ''' </summary>
    ''' <param name="someFile">An iniFile object with user definable parameters</param>
    Public Sub fileChooser(ByRef someFile As iniFile)
        If exitCode Then Exit Sub
        Console.Clear()
        handleFileChooserChoice(someFile)
        If exitCode Then Exit Sub
        handleFileChooserConfirm(someFile)
    End Sub

    ''' <summary>
    ''' Allows the user change the file name or directory of a given iniFile object
    ''' </summary>
    ''' <param name="someFile">The iniFile object whose parameters are being modified by the user</param>
    Private Sub handleFileChooserChoice(ByRef someFile As iniFile)
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
            Case (input = "1" And someFile.initName = "") Or (input = "2" And someFile.secondName <> "")
                someFile.name = someFile.secondName
            Case (input = "2" And someFile.secondName = "") Or (input = "3" And someFile.initName <> "" And someFile.secondName <> "")
                dirChooser(someFile.dir)
            Case Else
                someFile.name = input
        End Select
    End Sub

    ''' <summary>
    ''' Confirms the user's choice of a file's parameters in the File Chooser and allows them to make changes before saving
    ''' </summary>
    ''' <param name="someFile">The iniFile object whose parameters are being modified by the user</param>
    Public Sub handleFileChooserConfirm(ByRef someFile As iniFile)
        menuTopper = "File Chooser"
        Dim iExitCode As Boolean = False
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
            Dim input As String = Console.ReadLine()
            Select Case input
                Case "0"
                    exitCode = iExitCode = True
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

    ''' <summary>
    ''' Presents an interface to the user allowing them to operate on windows directory parameters
    ''' </summary>
    ''' <param name="dir">A user definable windows directory path</param>
    Public Sub dirChooser(ByRef dir As String)
        If exitCode Then Exit Sub
        Console.Clear()
        handleDirChooserChoice(dir)
        If exitCode Then Exit Sub
        handleDirChooserConfirm(dir)
    End Sub

    ''' <summary>
    ''' Allows the user to select a directory using a similar interface to the File Chooser
    ''' </summary>
    ''' <param name="dir">The String containing the directory the user is parameterizing</param>
    Private Sub handleDirChooserChoice(ByRef dir As String)
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
                Console.Clear()
                chkDirExist(dir)
            Case Else
                dir = input
                Console.Clear()
                chkDirExist(dir)
        End Select
    End Sub

    ''' <summary>
    ''' Confirms the user's choice of directory and allows them to change it 
    ''' </summary>
    ''' <param name="dir">The String containing the directory the user is parameterizing</param>
    Private Sub handleDirChooserConfirm(ByRef dir As String)
        menuTopper = "Directory Chooser"
        Dim iExitCode As Boolean = False
        Do Until iExitCode
            If exitCode Then Exit Sub
            Console.Clear()
            printMenuTop({"Choose a directory"}, True)
            printMenuOpt("Directory Chooser", "Change the directory")
            printMenuOpt("Confirm (default)", "Use this directory")
            printBlankMenuLine()
            printMenuLine("Current Directory: " & dir, "l")
            printMenuLine(menuStr02)
            Console.Write("Choose a number from above, or leave blank to run the default: ")
            Dim input As String = Console.ReadLine()
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

    ''' <summary>
    ''' An object representing a .ini configuration file
    ''' </summary>
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

        ''' <summary>
        ''' Returns an iniFile as it would appear on disk as a String
        ''' </summary>
        ''' <returns></returns>
        Public Overrides Function toString() As String
            Dim out As String = ""
            For i As Integer = 0 To sections.Count - 2
                out += sections.Values(i).ToString & Environment.NewLine
            Next
            out += sections.Values.Last.ToString
            Return out
        End Function

        ''' <summary>
        ''' Creates an empty iniFile.
        ''' </summary>
        Public Sub New()
            dir = ""
            name = ""
            initDir = ""
            initName = ""
            secondName = ""
        End Sub

        ''' <summary>
        ''' Creates an uninitalized iniFile with a directory and a filename.
        ''' </summary>
        ''' <param name="directory">A windows directory containing a .ini file</param>
        ''' <param name="filename">The name of the .ini file contained in the given directory </param>
        Public Sub New(directory As String, filename As String)
            dir = directory
            name = filename
            initDir = directory
            initName = filename
            secondName = ""
        End Sub

        ''' <summary>
        ''' Creates an uninitialized iniFile with a suggested alternative name
        ''' </summary>
        ''' <param name="directory">A windows directory containing a .ini file</param>
        ''' <param name="filename">The name of the .ini file contained in the given directory</param>
        ''' <param name="rename">The alternative/suggested name for the given .ini file</param>
        Public Sub New(directory As String, filename As String, rename As String)
            dir = directory
            name = filename
            initDir = directory
            initName = filename
            secondName = rename
        End Sub

        ''' <summary>
        ''' Writes the contents of a provided String to our iniFile's path, overwriting any existing contents
        ''' </summary>
        ''' <param name="tostr">The string to be written to disk</param>
        Public Sub overwriteToFile(tostr As String)
            Dim file As StreamWriter
            Try
                file = New StreamWriter(Me.path)
                file.Write(tostr)
                file.Close()
            Catch ex As Exception
                exc(ex)
            End Try
        End Sub

        ''' <summary>
        ''' Restores the initial directory and name parameters in the iniFile 
        ''' </summary>
        Public Sub resetParams()
            dir = initDir
            name = initName
        End Sub

        ''' <summary>
        ''' Returns the full windows file path of the iniFile as a String
        ''' </summary>
        ''' <returns></returns>
        Public Function path() As String
            Return dir & "\" & name
        End Function

        ''' <summary>
        ''' Returns the starting line number of each section in the iniFile as a list of integers
        ''' </summary>
        ''' <returns></returns>
        Public Function getLineNumsFromSections() As List(Of Integer)
            Dim outList As New List(Of Integer)
            For Each section In sections.Values
                outList.Add(section.startingLineNumber)
            Next
            Return outList
        End Function

        ''' <summary>
        ''' Constructs an iniFile object using an internet source
        ''' </summary>
        ''' <param name="lines">The array of Strings representing a remote .ini file</param>
        Public Sub New(lines As String())
            Dim sectionToBeBuilt As New List(Of String)
            Dim lineTrackingList As New List(Of Integer)
            Dim lastLineWasEmpty As Boolean = False
            lineCount = 1
            For Each line In lines
                processiniLine(line, sectionToBeBuilt, lineTrackingList, lastLineWasEmpty)
            Next
            If sectionToBeBuilt.Count <> 0 Then mkSection(sectionToBeBuilt, lineTrackingList)
        End Sub

        ''' <summary>
        ''' Processes a line in a .ini file and updates the iniFile object metadata accordingly
        ''' </summary>
        ''' <param name="currentLine">The current string being read</param>
        ''' <param name="sectionToBeBuilt">The pending list of strings to be built into an iniSection</param>
        ''' <param name="lineTrackingList">The associated list of line number integers for the section strings</param>
        ''' <param name="lastLineWasEmpty">The boolean representing whether or not the previous line was empty</param>
        Private Sub processiniLine(ByRef currentLine As String, ByRef sectionToBeBuilt As List(Of String), ByRef lineTrackingList As List(Of Integer), ByRef lastLineWasEmpty As Boolean)
            Select Case True
                Case currentLine.StartsWith(";")
                    Dim newCom As New iniComment(currentLine, lineCount)
                    comments.Add(comments.Count, newCom)
                Case (Not currentLine.StartsWith("[") And Not currentLine.Trim = "") Or (currentLine.Trim <> "" And sectionToBeBuilt.Count = 0)
                    lineNotEmpty(currentLine, sectionToBeBuilt, lineTrackingList, lastLineWasEmpty)
                Case currentLine.Trim <> "" And Not sectionToBeBuilt.Count = 0
                    mkSection(sectionToBeBuilt, lineTrackingList)
                    lineNotEmpty(currentLine, sectionToBeBuilt, lineTrackingList, lastLineWasEmpty)
                Case Else
                    lastLineWasEmpty = True
            End Select
            lineCount += 1
        End Sub

        ''' <summary>
        ''' Updates the current section being built with the current line, marks the lastlinewasempty (this line) as false
        ''' </summary>
        ''' <param name="currentLine">The current line to be added to the inisection</param>
        ''' <param name="sectionToBeBuilt">The list of Strings that will become an iniSection</param>
        ''' <param name="lineTrackingList">The line numbers of each line in the section to be built</param>
        ''' <param name="lastLineWasEmpty">The boolean indicating that this line was not empty</param>
        Private Sub lineNotEmpty(ByRef currentLine As String, ByRef sectionToBeBuilt As List(Of String), ByRef lineTrackingList As List(Of Integer), ByRef lastLineWasEmpty As Boolean)
            updSec(sectionToBeBuilt, lineTrackingList, currentLine)
            lastLineWasEmpty = False
        End Sub

        ''' <summary>
        ''' Manages line & number tracking for iniSections whose construction is pending
        ''' </summary>
        ''' <param name="secList">The list of strings for the iniSection</param>
        ''' <param name="lineList">The list of line numbers for the iniSection</param>
        ''' <param name="curLine">The current line to be added to the section</param>
        Private Sub updSec(ByRef secList As List(Of String), ByRef lineList As List(Of Integer), curLine As String)
            secList.Add(curLine)
            lineList.Add(lineCount)
        End Sub

        ''' <summary>
        ''' Attempts to read a .ini file from disk and initalize the iniFile object
        ''' </summary>
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

        ''' <summary>
        ''' Ensures that any call to an ini file on the system will be to a file that exists in a directory that exists.
        ''' </summary>
        Public Sub validate()
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

        ''' <summary>
        ''' Reorders the iniSections in an iniFile object to be in the same sorted state as a provided list of Strings
        ''' </summary>
        ''' <param name="sortedSections">The sorted state of the sections by name</param>
        Public Sub sortSections(ByVal sortedSections As List(Of String))
            Dim tempFile As New iniFile
            For Each sectionName In sortedSections
                tempFile.sections.Add(sectionName, sections.Item(sectionName))
            Next
            Me.sections = tempFile.sections
        End Sub

        ''' <summary>
        ''' Find the line number of a comment by its string. Returns -1 if not found
        ''' </summary>
        ''' <param name="com">The string containing the comment text to be searched for</param>
        ''' <returns></returns>
        Public Function findCommentLine(com As String) As Integer
            For Each comment In comments.Values
                If comment.comment = com Then Return comment.lineNumber
            Next
            Return -1
        End Function

        ''' <summary>
        ''' Returns the section names from the iniFile object as a list of Strings
        ''' </summary>
        ''' <returns></returns>
        Public Function getSectionNamesAsList() As List(Of String)
            Dim out As New List(Of String)
            For Each section In sections.Values
                out.Add(section.name)
            Next
            Return out
        End Function

        ''' <summary>
        ''' Attempts to create a new iniSection object and add it to the iniFile
        ''' </summary>
        ''' <param name="sectionToBeBuilt">The list of strings in the iniSection</param>
        ''' <param name="lineTrackingList">The list of line numbers associated with the given strings</param>
        Private Sub mkSection(sectionToBeBuilt As List(Of String), lineTrackingList As List(Of Integer))
            Dim sectionHolder As New iniSection(sectionToBeBuilt, lineTrackingList)
            Try
                sections.Add(sectionHolder.name, sectionHolder)
            Catch ex As Exception
                'This will catch entries whose names are identical (case sensitive), and ignore them 
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

    ''' <summary>
    ''' An object representing a section of a .ini file
    ''' </summary>
    Public Class iniSection
        Public startingLineNumber As Integer
        Public endingLineNumber As Integer
        Public name As String
        Public keys As New Dictionary(Of Integer, iniKey)

        ''' <summary>
        ''' Sorts a section's keys into lists based on keyType
        ''' </summary>
        ''' Expects that the number of parameters in keyTypeList is equal to the number of values in listOfKeyLists, with the final list being an "error" list
        ''' holding keys of types not defined in the keyTypeList
        ''' <param name="keyTypeList"></param>
        ''' <param name="listOfKeyLists"></param>
        Public Sub constructKeyLists(ByVal keyTypeList As List(Of String), ByRef listOfKeyLists As List(Of List(Of iniKey)))
            For Each key In Me.keys.Values
                If keyTypeList.Contains(key.keyType.ToLower) Then
                    listOfKeyLists(keyTypeList.IndexOf(key.keyType.ToLower)).Add(key)
                Else
                    listOfKeyLists.Last.Add(key)
                End If
            Next
        End Sub

        ''' <summary>
        ''' Returns the iniSection name as it would appear on disk.
        ''' </summary>
        ''' <returns></returns>
        Public Function getFullName() As String
            Return "[" & Me.name & "]"
        End Function

        ''' <summary>
        ''' Creates a new (empty) iniSection object.
        ''' </summary>
        Public Sub New()
            startingLineNumber = 0
            endingLineNumber = 0
            name = ""
        End Sub

        ''' <summary>
        ''' Creates a new iniSection object 
        ''' </summary>
        ''' <param name="listOfLines">The list of strings comprising the iniSection</param>
        ''' <param name="listOfLineCounts">The line numbers of the strings comprising the iniSection</param>
        Public Sub New(ByVal listOfLines As List(Of String), listOfLineCounts As List(Of Integer))
            name = listOfLines(0).Trim(CChar("[")).Trim(CChar("]"))
            startingLineNumber = listOfLineCounts(0)
            endingLineNumber = listOfLineCounts(listOfLineCounts.Count - 1)

            If listOfLines.Count > 1 Then
                For i As Integer = 1 To listOfLines.Count - 1
                    Dim curKey As New iniKey(listOfLines(i), listOfLineCounts(i))
                    keys.Add(i - 1, curKey)
                Next
            End If
        End Sub

        ''' <summary>
        ''' Creates a new iniSection object without tracking the line numbers
        ''' </summary>
        ''' <param name="listOfLines">The list of strings comprising the iniSection</param>
        Public Sub New(ByVal listOfLines As List(Of String))
            name = listOfLines(0).Trim(CChar("[")).Trim(CChar("]"))
            startingLineNumber = 1
            endingLineNumber = 1 + listOfLines.Count

            If listOfLines.Count > 1 Then
                For i As Integer = 1 To listOfLines.Count - 1
                    keys.Add(i - 1, New iniKey(listOfLines(i), 0))
                Next
            End If
        End Sub

        ''' <summary>
        ''' Returns the keys in the iniSection as a list of Strings
        ''' </summary>
        ''' <returns></returns>
        Public Function getKeysAsList() As List(Of String)
            Dim out As New List(Of String)
            For Each key In Me.keys.Values
                out.Add(key.toString)
            Next
            Return out
        End Function

        ''' <summary>
        ''' Compares two iniSections, returns false if they are not the same.
        ''' </summary>
        ''' <param name="ss">The section to be compared against</param>
        ''' <param name="removedKeys">A return list on iniKey objects that appear in the iniFile object but not the given</param>
        ''' <param name="addedKeys">A return list of iniKey objects that appear in the given iniFile object but not this one</param>
        ''' <returns></returns>
        Public Function compareTo(ss As iniSection, ByRef removedKeys As List(Of iniKey), ByRef addedKeys As List(Of iniKey)) As Boolean

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

            Return removedKeys.Count + addedKeys.Count = 0
        End Function

        ''' <summary>
        ''' Returns an iniSection as it would appear on disk as a String
        ''' </summary>
        ''' <returns></returns>
        Public Overrides Function ToString() As String
            Dim out As String = Me.getFullName
            For Each key In keys.Values
                out += Environment.NewLine & key.toString
            Next
            out += Environment.NewLine
            Return out
        End Function
    End Class

    ''' <summary>
    ''' An object representing the name value pairs that make up iniSections
    ''' </summary>
    Public Class iniKey

        Public name As String
        Public value As String
        Public lineNumber As Integer
        Public keyType As String

        ''' <summary>
        ''' Returns whether or not an iniKey's name is equal to a given value
        ''' </summary>
        ''' <param name="n">The string to check equality for</param>
        ''' <returns></returns>
        Public Function nameIs(n As String) As Boolean
            Return name = n
        End Function

        ''' <summary>
        ''' Returns whether or not an iniKey's type is equal to a given value
        ''' </summary>
        ''' <param name="t">The string to check equality for</param>
        ''' <returns></returns>
        Public Function typeIs(t As String) As Boolean
            Return keyType = t
        End Function

        ''' <summary>
        ''' Returns whether or not an iniKey object's value begins or ends with a given value
        ''' </summary>
        ''' <param name="txt">The given string to search for</param>
        ''' <returns></returns>
        Public Function vStartsOrEndsWith(txt As String) As Boolean
            Return value.StartsWith(txt) Or value.EndsWith(txt)
        End Function

        ''' <summary>
        ''' Returns whether or not an iniKey object's name begin or ends with a given value
        ''' </summary>
        ''' <param name="txt">The given string to search for</param>
        ''' <returns></returns>
        Public Function nStartsOrEndsWith(txt As String) As Boolean
            Return name.StartsWith(txt) Or name.EndsWith(txt)
        End Function

        ''' <summary>
        ''' Returns whether or not an iniKey object's value contains a given string with conditional case casting
        ''' </summary>
        ''' <param name="txt">The string to search for</param>
        ''' <param name="toLower">A boolean specifying whether or not the strings should be cast to lowercase</param>
        ''' <returns></returns>
        Public Function vHas(txt As String, toLower As Boolean) As Boolean
            Return If(toLower, value.ToLower.Contains(txt.ToLower), value.Contains(txt))
        End Function

        ''' <summary>
        ''' Returns whether or not an iniKey object's value contains a given string
        ''' </summary>
        ''' <param name="txt">The string to search for</param>
        ''' <returns></returns>
        Public Function vHas(txt As String) As Boolean
            Return vHas(txt, False)
        End Function

        ''' <summary>
        ''' Returns whether or not an iniKey object's value contains any of a given array of strings
        ''' </summary>
        ''' <param name="txts">The array of search strings</param>
        ''' <returns></returns>
        Public Function vHasAny(txts As String()) As Boolean
            Return vHasAny(txts, False)
        End Function

        ''' <summary>
        ''' Returns whether or not an iniKey object's value contains any of a given array of stringd with conditional case casting
        ''' </summary>
        ''' <param name="txts">The array of search strings</param>
        ''' <param name="toLower">A boolean specifying whether or not the strings should be cast to lowercase</param>
        ''' <returns></returns>
        Public Function vHasAny(txts As String(), toLower As Boolean) As Boolean
            For Each txt In txts
                If vHas(txt, toLower) Then Return True
            Next
            Return False
        End Function

        ''' <summary>
        ''' Returns whether or not an iniKey object's value is equal to a given string
        ''' </summary>
        ''' <param name="txt">The string to be searched for</param>
        ''' <returns></returns>
        Public Function vIs(txt As String) As Boolean
            Return vIs(txt, False)
        End Function

        ''' <summary>
        ''' Returns whether or not an iniKey object's value is equal to a given string with conditional case casting
        ''' </summary>
        ''' <param name="txt">The string to be searched for</param>
        ''' <param name="toLower">A boolean specifying whether or not the strings should be cast to lowercase</param>
        ''' <returns></returns>
        Public Function vIs(txt As String, toLower As Boolean) As Boolean
            Return If(toLower, value.ToLower.Equals(txt), value.Equals(txt))
        End Function

        ''' <summary>
        ''' Returns an iniKey object's keyName field with numbers removed
        ''' </summary>
        ''' <param name="keyName">The string containing the iniKey's keyname</param>
        ''' <returns></returns>
        Private Function stripNums(keyName As String) As String
            Return New Regex("[\d]").Replace(keyName, "")
        End Function

        ''' <summary>
        ''' Create an iniKey object from a string containing a name value pair
        ''' </summary>
        ''' <param name="line">A string in the format name=value</param>
        ''' <param name="count">The line number for the string</param>
        Public Sub New(ByVal line As String, ByVal count As Integer)
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

        ''' <summary>
        ''' Returns the key in name=value format as a String
        ''' </summary>
        ''' <returns></returns>
        Public Overrides Function toString() As String
            Return name & "=" & value
        End Function

        ''' <summary>
        ''' Outputs the key in Line: #### - name=value format
        ''' </summary>
        ''' <returns></returns>
        Public Function lineString() As String
            Return "Line: " & lineNumber & " - " & Me.toString
        End Function

    End Class

    ''' <summary>
    ''' An object representing a comment in a .ini file
    ''' </summary>
    Class iniComment

        Public comment As String
        Public lineNumber As Integer

        ''' <summary>
        ''' Creates a new iniComment object
        ''' </summary>
        ''' <param name="c">The comment text</param>
        ''' <param name="l">The line number</param>
        Public Sub New(c As String, l As Integer)
            comment = c
            lineNumber = l
        End Sub
    End Class
End Module