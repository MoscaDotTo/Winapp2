Option Strict On
Imports System.IO
''' <summary>
''' An object representing a .ini configuration file
''' </summary>
Public Class iniFile
    Dim lineCount As Integer = 1
    ' The current state of the directory & name of the file
    Public dir As String
    Public name As String
    ' The initial state of the directory & name of the file (for restoration purposes) 
    Public initDir As String
    Public initName As String
    ' Suggested rename for output files
    Public secondName As String
    ' Sections will be initially stored in the order they're read
    Public sections As New Dictionary(Of String, iniSection)
    ' Any line comments will be saved in the order they're read 
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
    ''' Creates an uninitialized iniFile with a directory and a filename.
    ''' </summary>
    ''' <param name="directory">A windows directory containing a .ini file</param>
    ''' <param name="filename">The name of the .ini file contained in the given directory </param>
    ''' <param name="rename">A provided suggestion for a rename should the user open the File Chooser on this file</param>
    Public Sub New(Optional directory As String = "", Optional filename As String = "", Optional rename As String = "")
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
        Return $"{dir}\{name}"
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
    ''' Processes a line in a .ini file and updates the iniFile object meta data accordingly
    ''' </summary>
    ''' <param name="currentLine">The current string being read</param>
    ''' <param name="sectionToBeBuilt">The pending list of strings to be built into an iniSection</param>
    ''' <param name="lineTrackingList">The associated list of line number integers for the section strings</param>
    ''' <param name="lastLineWasEmpty">The boolean representing whether or not the previous line was empty</param>
    Private Sub processiniLine(ByRef currentLine As String, ByRef sectionToBeBuilt As List(Of String), ByRef lineTrackingList As List(Of Integer), ByRef lastLineWasEmpty As Boolean)
        Select Case True
            Case currentLine.StartsWith(";")
                comments.Add(comments.Count, New iniComment(currentLine, lineCount))
            Case (Not currentLine.StartsWith("[") And Not currentLine.Trim = "") Or (currentLine.Trim <> "" And sectionToBeBuilt.Count = 0)
                updSec(sectionToBeBuilt, lineTrackingList, currentLine, lastLineWasEmpty)
            Case currentLine.Trim <> "" And Not sectionToBeBuilt.Count = 0
                mkSection(sectionToBeBuilt, lineTrackingList)
                updSec(sectionToBeBuilt, lineTrackingList, currentLine, lastLineWasEmpty)
            Case Else
                lastLineWasEmpty = True
        End Select
        lineCount += 1
    End Sub

    ''' <summary>
    ''' Manages line and number tracking for iniSections whose construction is pending
    ''' </summary>
    ''' <param name="secList">The list of strings for the iniSection</param>
    ''' <param name="lineList">The list of line numbers for the iniSection</param>
    ''' <param name="curLine">The current line to be added to the section</param>
    Private Sub updSec(ByRef secList As List(Of String), ByRef lineList As List(Of Integer), curLine As String, ByRef lastLineWasEmpty As Boolean)
        secList.Add(curLine)
        lineList.Add(lineCount)
        lastLineWasEmpty = False
    End Sub

    ''' <summary>
    ''' Attempts to read a .ini file from disk and initialize the iniFile object
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
            Console.WriteLine(ex.Message & Environment.NewLine & $"Failure occurred during iniFile construction at line: {lineCount} in {name}")
        End Try
    End Sub

    ''' <summary>
    ''' Ensures that any call to an ini file on the system will be to a file that exists in a directory that exists.
    ''' If an iniFile's sections already exist, skip this.
    ''' </summary>
    Public Sub validate()
        clrConsole()
        If pendingExit() Or name = "" Then Exit Sub
        ' Make sure both the file and the directory actually exist
        While Not File.Exists(path())
            chkDirExist(dir)
            If pendingExit() Then Exit Sub
            chkFileExist(Me)
            If pendingExit() Then Exit Sub
        End While
        ' Make sure that the file isn't empty
        Try
            Dim iniTester As New iniFile(dir, name)
            iniTester.init()
            Dim clearAtEnd As Boolean = False
            While iniTester.sections.Count = 0
                clearAtEnd = True
                clrConsole()
                printMenuLine(bmenu("Empty ini file detected. Press any key to try again."))
                Console.ReadKey()
                fileChooser(iniTester)
                If pendingExit() Then Exit Sub
                iniTester.validate()
                If pendingExit() Then Exit Sub
            End While
            sections = iniTester.sections
            comments = iniTester.comments
            clrConsole(clearAtEnd)
        Catch ex As Exception
            exc(ex)
            exitCode = True
        End Try
    End Sub

    ''' <summary>
    ''' Reorders the iniSections in an iniFile object to be in the same sorted state as a provided list of Strings
    ''' </summary>
    ''' <param name="sortedSections">The sorted state of the sections by name</param>
    Public Sub sortSections(sortedSections As strList)
        Dim tempFile As New iniFile
        sortedSections.items.ForEach(Sub(sectionName) tempFile.sections.Add(sectionName, sections.Item(sectionName)))
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
    Public Function namesToListOfStr() As List(Of String)
        Dim out As New List(Of String)
        For Each section In sections.Values
            out.Add(section.name)
        Next
        Return out
    End Function

    Public Function namesToStrList() As strList
        Dim out As New strList
        For Each section In sections.Values
            out.add(section.name)
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
                Console.WriteLine($"Error: Duplicate section name detected: {sectionToBeBuilt(0)}")
                Console.WriteLine($"Line: {lineCount}")
                Console.WriteLine($"Duplicates the entry on line: {lineErr}")
                Console.WriteLine("This section will be ignored until it is given a unique name.")
                Console.WriteLine()
                Console.WriteLine("Press enter to continue.")
                Console.ReadLine()
            End If
        Finally
            sectionToBeBuilt.Clear()
            lineTrackingList.Clear()
        End Try
    End Sub
End Class