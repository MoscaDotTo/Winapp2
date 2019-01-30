'    Copyright (C) 2018-2019 Robbie Ward
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
    ''' Enforces that a user selected file exists
    ''' </summary>
    ''' <param name="someFile">An iniFile object with user defined path and name parameters</param>
    Public Sub chkFileExist(someFile As iniFile)
        If pendingExit() Then Exit Sub
        Dim iExitCode As Boolean = False
        While Not File.Exists(someFile.path)
            If pendingExit() Then Exit Sub
            menuHeaderText = "Error"
            While Not iExitCode
                Console.Clear()
                printMenuTop({$"{someFile.name} does not exist."})
                print(1, "File Chooser (default)", "Change the file name")
                print(1, "Directory Chooser", "Change the directory", closeMenu:=True)
                Dim input As String = Console.ReadLine
                Select Case input
                    Case "0"
                        exitCode = True
                        iExitCode = True
                        Exit Sub
                    Case "1", ""
                        fileChooser(someFile)
                    Case "2"
                        dirChooser(someFile.dir)
                    Case Else
                        menuHeaderText = invInpStr
                End Select
                If Not File.Exists(someFile.path) And Not menuHeaderText = invInpStr Then menuHeaderText = "Error"
            End While
        End While
    End Sub

    ''' <summary>
    ''' Enforces that a user defined directory exists, either by selecting a new one or creating one.
    ''' </summary>
    ''' <param name="dir">A user defined windows directory</param>
    Public Sub chkDirExist(ByRef dir As String)
        If pendingExit() Or Directory.Exists(dir) Then Exit Sub
        Dim iExitCode As Boolean = False
        menuHeaderText = "Error"
        While Not iExitCode
            Console.Clear()
            printMenuTop({$"{dir} does not exist."})
            print(1, "Create Directory", "Create this directory")
            print(1, "Directory Chooser (default)", "Specify a new directory", closeMenu:=True)
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
                    menuHeaderText = invInpStr
            End Select
            If Not Directory.Exists(dir) And Not menuHeaderText = invInpStr Then menuHeaderText = "Error"
            If Directory.Exists(dir) Then iExitCode = True
        End While
    End Sub

    ''' <summary>
    ''' Presents a menu to the user allowing them to perform some file and directory operations
    ''' </summary>
    ''' <param name="someFile">An iniFile object with user definable parameters</param>
    Public Sub fileChooser(ByRef someFile As iniFile)
        If pendingExit() Then Exit Sub
        Console.Clear()
        handleFileChooserChoice(someFile)
        If pendingExit() Then Exit Sub
        handleFileChooserConfirm(someFile)
    End Sub

    ''' <summary>
    ''' Allows the user change the file name or directory of a given iniFile object
    ''' </summary>
    ''' <param name="someFile">The iniFile object whose parameters are being modified by the user</param>
    Private Sub handleFileChooserChoice(ByRef someFile As iniFile)
        menuHeaderText = "File Chooser"
        printMenuTop({"Choose a file name, or open the directory chooser to choose a directory"})
        print(1, someFile.initName, "Use the default name", someFile.initName <> "")
        print(1, someFile.secondName, "Use the default rename", someFile.secondName <> "")
        print(1, "Directory Chooser", "Choose a new directory", trailingBlank:=True)
        print(0, $"Current Directory: {replDir(someFile.dir)}")
        print(0, $"Current File:      {someFile.name}", closeMenu:=True)
        Console.Write($"{Environment.NewLine} Enter a number, {If(Not someFile.name = "", "", "or ")}a new file name{If(someFile.name = "", "", $", or leave blank to continue using '{someFile.name}'")}: ")
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
        menuHeaderText = "File Chooser"
        Dim iExitCode As Boolean = False
        Do Until iExitCode
            Console.Clear()
            printMenuTop({"Confirm your settings or return to the options to change them."})
            print(1, "File Chooser", "Change the file name")
            print(1, "Directory Chooser", "Change the directory")
            print(1, "Confirm (default)", "Save changes")
            print(0, $"Current Directory: {replDir(someFile.dir)}", leadingBlank:=True)
            print(0, $"Current File     : {someFile.name}", closeMenu:=True)
            Console.Write(Environment.NewLine & "Enter a number, or leave blank to run the default: ")
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
                    menuHeaderText = invInpStr
            End Select
        Loop
    End Sub

    ''' <summary>
    ''' Presents an interface to the user allowing them to operate on windows directory parameters
    ''' </summary>
    ''' <param name="dir">A user definable windows directory path</param>
    Public Sub dirChooser(ByRef dir As String)
        If pendingExit() Then Exit Sub
        Console.Clear()
        handleDirChooserChoice(dir)
        If pendingExit() Then Exit Sub
        handleDirChooserConfirm(dir)
    End Sub

    ''' <summary>
    ''' Allows the user to select a directory using a similar interface to the File Chooser
    ''' </summary>
    ''' <param name="dir">The String containing the directory the user is parameterizing</param>
    Private Sub handleDirChooserChoice(ByRef dir As String)
        Console.Clear()
        menuHeaderText = "Directory Chooser"
        printMenuTop({"Choose a directory"})
        print(1, "Use default (default)", "Use the same folder as winapp2ool.exe")
        print(1, "Parent Folder", "Go up a level")
        print(1, "Current folder", "Continue using the same folder as below")
        print(0, $"Current Directory: {dir}", leadingBlank:=True, closeMenu:=True)
        Console.Write(Environment.NewLine & "Choose a number from above, enter a new directory, or leave blank to run the default: ")
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
        menuHeaderText = "Directory Chooser"
        Dim iExitCode As Boolean = False
        Do Until iExitCode
            If pendingExit() Then Exit Sub
            Console.Clear()
            printMenuTop({"Choose a directory"})
            print(1, "Directory Chooser", "Change the directory")
            print(1, "Confirm (default)", "Use this directory")
            print(0, $"Current Directory: {dir}", leadingBlank:=True, closeMenu:=True)
            Console.Write(Environment.NewLine & "Choose a number from above, or leave blank to run the default: ")
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
                    menuHeaderText = invInpStr
            End Select
        Loop
    End Sub

    ''' <summary>
    ''' An object representing a .ini configuration file
    ''' </summary>
    Public Class iniFile
        Dim lineCount As Integer = 1
        ' The current state of the directory & name of the file
        Public dir As String
        Public name As String
        ' The inital state of the direcotry & name of the file (for restoration purposes) 
        Public initDir As String
        Public initName As String
        ' Suggested rename for output files
        Public secondName As String
        ' Sections will be initally stored in the order they're read
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
        ''' Creates an uninitalized iniFile with a directory and a filename.
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
        ''' Processes a line in a .ini file and updates the iniFile object metadata accordingly
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
                Console.WriteLine(ex.Message & Environment.NewLine & $"Failure occurred during iniFile construction at line: {lineCount} in {name}")
            End Try
        End Sub

        ''' <summary>
        ''' Ensures that any call to an ini file on the system will be to a file that exists in a directory that exists.
        ''' If an iniFile's sections already exist, skip this.
        ''' </summary>
        Public Sub validate()
            Console.Clear()
            If pendingExit() Or Me.name = "" Then Exit Sub
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
                    Console.Clear()
                    printMenuLine(bmenu("Empty ini file detected. Press any key to try again."))
                    Console.ReadKey()
                    fileChooser(iniTester)
                    If pendingExit() Then Exit Sub
                    iniTester.validate()
                    If pendingExit() Then Exit Sub
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
            sortedSections.ForEach(Sub(sectionName) tempFile.sections.Add(sectionName, sections.Item(sectionName)))
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

    ''' <summary>
    ''' An object representing a section of a .ini file
    ''' </summary>
    Public Class iniSection
        Public startingLineNumber As Integer
        Public endingLineNumber As Integer
        Public name As String
        Public keys As New Dictionary(Of Integer, iniKey)

        ''' <summary>
        ''' Sorts a section's keys into keylists based on their KeyType
        ''' </summary>
        ''' <param name="listOfKeyLists">The list of keyLists to be sorted into</param>
        ''' The last list in the keylist list holds the error keys
        Public Sub constKeyLists(ByRef listOfKeyLists As List(Of keyList))
            Dim keyTypeList As New List(Of String)
            listOfKeyLists.ForEach(Sub(kl) keyTypeList.Add(kl.keyType.ToLower))
            For Each key In Me.keys.Values
                Dim type = key.keyType.ToLower
                If keyTypeList.Contains(type) Then listOfKeyLists(keyTypeList.IndexOf(type)).add(key) Else listOfKeyLists.Last.add(key)
            Next
        End Sub

        ''' <summary>
        ''' Removes a series of keys from the section
        ''' </summary>
        ''' <param name="indicies"></param>
        Public Sub removeKeys(indicies As List(Of Integer))
            indicies.ForEach(Sub(ind) Me.keys.Remove(ind))
        End Sub

        ''' <summary>
        ''' Returns the iniSection name as it would appear on disk.
        ''' </summary>
        ''' <returns></returns>
        Public Function getFullName() As String
            Return $"[{Me.name}]"
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
        ''' Creates a new iniSection object without tracking the line numbers
        ''' </summary>
        ''' <param name="listOfLines">The list of Strings comprising the iniSection</param>
        ''' <param name="listOfLineCounts">The list of line numbers associated with the lines</param>
        Public Sub New(ByVal listOfLines As List(Of String), Optional listOfLineCounts As List(Of Integer) = Nothing)
            name = listOfLines(0).Trim(CChar("["), CChar("]"))
            startingLineNumber = If(listOfLineCounts IsNot Nothing, listOfLineCounts(0), 1)
            endingLineNumber = startingLineNumber + listOfLines.Count
            If listOfLines.Count > 1 Then
                For i As Integer = 1 To listOfLines.Count - 1
                    keys.Add(i - 1, New iniKey(listOfLines(i), If(listOfLineCounts Is Nothing, 0, listOfLineCounts(i))))
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
        Public Function compareTo(ss As iniSection, ByRef removedKeys As keyList, ByRef addedKeys As keyList) As Boolean
            ' Create a copy of the section so we can modify it
            Dim secondSection As New iniSection With {.name = ss.name, .startingLineNumber = ss.startingLineNumber}
            For i As Integer = 0 To ss.keys.Count - 1
                secondSection.keys.Add(i, ss.keys.Values(i))
            Next
            Dim noMatch As Boolean
            Dim tmpList As New List(Of Integer)
            For Each key In keys.Values
                noMatch = True
                For i As Integer = 0 To secondSection.keys.Values.Count - 1
                    Select Case True
                        Case key.keyType.ToLower = secondSection.keys.Values(i).keyType.ToLower And key.value.ToLower = secondSection.keys.Values(i).value.ToLower
                            noMatch = False
                            tmpList.Add(i)
                            Exit For
                    End Select
                Next
                ' If the key isn't found in the second (newer) section, consider it removed for now
                If noMatch Then removedKeys.add(key)
            Next
            ' Remove all matched keys
            tmpList.Reverse()
            secondSection.removeKeys(tmpList)
            ' Assume any remaining keys have been added
            For Each key In secondSection.keys.Values
                addedKeys.add(key)
            Next
            Return removedKeys.keyCount + addedKeys.keyCount = 0
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
    ''' A handy wrapper object for lists of iniKeys
    ''' </summary>
    Public Class keyList
        Public keys As List(Of iniKey)
        Public keyType As String

        ''' <summary>
        ''' Creates a new (empty) keylist
        ''' </summary>
        ''' <param name="kt">Optional String containing the expected KeyType of the keys in the list</param>
        Public Sub New(Optional kt As String = "")
            keys = New List(Of iniKey)
            keyType = kt
        End Sub

        ''' <summary>
        ''' Creates a new keylist using an existing list of iniKeys
        ''' </summary>
        ''' <param name="kl">A list of iniKeys to be inserted into the keylist</param>
        Public Sub New(kl As List(Of iniKey))
            keys = kl
            keyType = If(keys.Count > 0, keys(0).keyType, "")
        End Sub

        ''' <summary>
        ''' Conditionally adds a key to the keylist
        ''' </summary>
        ''' <param name="key">The key to be added</param>
        ''' <param name="cond">The condition under which to add the key</param>
        Public Sub add(key As iniKey, Optional cond As Boolean = True)
            If cond Then keys.Add(key)
        End Sub

        ''' <summary>
        ''' Adds a list of iniKeys to the keylist
        ''' </summary>
        ''' <param name="kl">The list to be added</param>
        Public Sub add(kl As List(Of iniKey))
            kl.ForEach(Sub(key) keys.Add(key))
        End Sub

        ''' <summary>
        ''' Removes a key from the keylist
        ''' </summary>
        ''' <param name="key">The key to be removed</param>
        Public Sub remove(key As iniKey)
            Me.keys.Remove(key)
        End Sub

        ''' <summary>
        ''' Removes a list of keys from the keylist
        ''' </summary>
        ''' <param name="kl">The list of keys to be removed</param>
        Public Sub remove(kl As List(Of iniKey))
            kl.ForEach(Sub(key) remove(key))
        End Sub

        ''' <summary>
        ''' Returns the number of keys in the keylist
        ''' </summary>
        ''' <returns></returns>
        Public Function keyCount() As Integer
            Return keys.Count
        End Function

        ''' <summary>
        ''' Returns whether or not the keyType of the list matches the input String
        ''' </summary>
        ''' <param name="type">The String against which to match the keylist's type</param>
        ''' <returns></returns>
        Public Function typeIs(type As String) As Boolean
            Return If(keyType = "", keys(0).keyType, keyType) = type
        End Function

        ''' <summary>
        ''' Returns the keylist in the form of a list of Strings
        ''' </summary>
        ''' <returns></returns>
        Public Function toListOfStr(Optional getVals As Boolean = False) As List(Of String)
            Dim out As New List(Of String)
            keys.ForEach(Sub(key) out.Add(If(getVals, key.value, key.toString)))
            Return out
        End Function

        ''' <summary>
        ''' Removes the last element in the key list if it exists
        ''' </summary>
        Public Sub removeLast()
            If keys.Count > 0 Then keys.Remove(keys.Last)
        End Sub

        ''' <summary>
        ''' Renumber keys according to the sorted state of the values
        ''' </summary>
        ''' <param name="sortedKeyValues"></param>
        Public Sub renumberKeys(sortedKeyValues As List(Of String))
            For i As Integer = 0 To Me.keyCount - 1
                keys(i).name = keys(i).keyType & i + 1
                keys(i).value = sortedKeyValues(i)
            Next
        End Sub

        ''' <summary>
        ''' Returns a list of integers containing the line numbers from the keylist
        ''' </summary>
        ''' <returns></returns>
        Public Function lineNums() As List(Of Integer)
            Dim out As New List(Of Integer)
            keys.ForEach(Sub(key) out.Add(key.lineNumber))
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
        Public Function nameIs(n As String, Optional tolower As Boolean = False) As Boolean
            Return If(tolower, name.ToLower = n.ToLower, name = n)
            name = n
        End Function

        ''' <summary>
        ''' Returns whether or not an iniKey's type is equal to a given value
        ''' </summary>
        ''' <param name="t">The string to check equality for</param>
        ''' <returns></returns>
        Public Function typeIs(t As String, Optional tolower As Boolean = False) As Boolean
            Return If(tolower, keyType.ToLower = t.ToLower, keyType = t)
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
        Public Function vHas(txt As String, Optional toLower As Boolean = False) As Boolean
            Return If(toLower, value.ToLower.Contains(txt.ToLower), value.Contains(txt))
        End Function

        ''' <summary>
        ''' Returns whether or not an iniKey object's value contains any of a given array of stringd with conditional case casting
        ''' </summary>
        ''' <param name="txts">The array of search strings</param>
        ''' <param name="toLower">A boolean specifying whether or not the strings should be cast to lowercase</param>
        ''' <returns></returns>
        Public Function vHasAny(txts As String(), Optional toLower As Boolean = False) As Boolean
            For Each txt In txts
                If vHas(txt, toLower) Then Return True
            Next
            Return False
        End Function

        ''' <summary>
        ''' Returns whether or not an iniKey object's value is equal to a given string with conditional case casting
        ''' </summary>
        ''' <param name="txt">The string to be searched for</param>
        ''' <param name="toLower">A boolean specifying whether or not the strings should be cast to lowercase</param>
        ''' <returns></returns>
        Public Function vIs(txt As String, Optional toLower As Boolean = False) As Boolean
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
        ''' Compares the names of two iniKeys and returns whether or not they match
        ''' </summary>
        ''' <param name="key"></param>
        ''' <returns></returns>
        Public Function compareNames(key As iniKey) As Boolean
            Return nameIs(key.name, True)
        End Function

        ''' <summary>
        ''' Compares the values of two iniKeys and returns whether or not they match
        ''' </summary>
        ''' <param name="key"></param>
        ''' <returns></returns>
        Public Function compareValues(key As iniKey) As Boolean
            Return vIs(key.value, True)
        End Function

        ''' <summary>
        ''' Create an iniKey object from a string containing a name value pair
        ''' </summary>
        ''' <param name="line">A string in the format name=value</param>
        ''' <param name="count">The line number for the string</param>
        Public Sub New(ByVal line As String, Optional ByVal count As Integer = 0)
            If line.Contains("=") Then
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
            Else
                name = line
                value = ""
                keyType = "Error"
            End If
        End Sub

        ''' <summary>
        ''' Returns the key in name=value format as a String
        ''' </summary>
        ''' <returns></returns>
        Public Overrides Function toString() As String
            Return $"{name}={value}"
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