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
''' <summary>
''' An object representing a .ini configuration file
''' </summary>
Public Class iniFile
    ''' <summary> The directory on the system in which the iniFile can be found </summary>
    Public Property Dir As String
    ''' <summary> The name of the file on disk </summary>
    Public Property Name As String
    ''' <summary> The <c> Dir </c> with which the iniFile was instantiated </summary>
    Public Property InitDir As String
    ''' <summary> The <c> Name </c> with which the iniFile was instantiated </summary>
    Public Property InitName As String

    ''' <summary> The individual sections found in the file <br/> <br/> Dictionary keys are the contained section's <c> Name </c> </summary>
    Public Property Sections As New Dictionary(Of String, iniSection)

    ''' <summary> Any comments found in the file, in the order they were found. Positions are not remembered. </summary>
    Public Property Comments As Dictionary(Of Integer, iniComment)

    ''' <summary> The suggested name for a file, shown during File Chooser prompts </summary>
    Public Property SecondName As String

    ''' <summary> The current line number of the file during reading </summary>
    Public Property LineCount As Integer
    ''' <summary> Holds the file name during attempted renames so changes can be reverted </summary>
    Private Property tmpRename As String
    ''' <summary> Indicates that this file must have a file name that exists </summary>
    Public Property mustExist As Boolean

    ''' <summary> Returns the full windows file path of the <c> iniFile </c> as a <c> String </c> </summary>
    Public Function Path() As String
        Return $"{Dir}\{Name}"
    End Function

    ''' <summary> Returns an <c> iniFile </c> as a single <c> String </c> </summary>
    Public Overrides Function toString() As String
        Dim out As String = ""
        For i As Integer = 0 To Sections.Count - 2
            out += Sections.Values(i).ToString & Environment.NewLine
        Next
        out += Sections.Values.Last.ToString
        Return out
    End Function

    ''' <summary> Returns a <c> Boolean </c> indicating the existence of an <c> iniSection </c> by its <c> Name </c> </summary>
    ''' <param name="sectionName"> The <c> Name </c> of the section to search for </param>
    ''' <returns> <c> True </c> if <paramref name="sectionName"/> matches the <c> Name </c> of an <c> iniSection </c> in the <c> iniFile </c>, <c> False </c> otherwise </returns>
    Public Function hasSection(sectionName As String) As Boolean
        Return Sections.ContainsKey(sectionName)
    End Function

    ''' <summary> Creates an <c> uninitialized iniFile </c> with a directory and a filename </summary>
    ''' <param name="directory"> A filesystem directory containing the ini file </param>
    ''' <param name="filename"> The name of the file on disk contained in <c> <paramref name="directory"/> </c> </param>
    ''' <param name="rename"> A provided suggestion for a rename should the user open the File Chooser on this file </param>
    ''' <param name="mExist"> Indicates that this file must exist for its owner to perform its work <br/> Optional, Default: <c> False </c> </param>
    Public Sub New(Optional directory As String = "", Optional filename As String = "", Optional rename As String = "", Optional mExist As Boolean = False)
        Dir = directory
        Name = filename
        InitDir = directory
        InitName = filename
        SecondName = rename
        Sections = New Dictionary(Of String, iniSection)
        Comments = New Dictionary(Of Integer, iniComment)
        LineCount = 1
        mustExist = mExist
    End Sub

    ''' <summary> Writes a given <c> String </c> (generally an ini file) to disk if the given 
    ''' <c> <paramref name="cond"/> </c> is met, overwriting any existing file contents </summary>
    ''' <param name="tostr"> The text to be written to disk </param>
    ''' <param name="cond"> Indicates that the file should be written to disk <br/> Optional, Default: <c> True </c> </param>
    Public Sub overwriteToFile(tostr As String, Optional cond As Boolean = True)
        If Not cond Then Exit Sub
        gLog($"Saving {Name}")
        Dim file As StreamWriter
        Try
            file = New StreamWriter(Me.Path)
            file.Write(tostr)
            file.Close()
            gLog("Save complete", indent:=True)
        Catch ex As Exception
            gLog("Save failed", indent:=True)
            exc(ex)
        End Try
    End Sub

    ''' <summary> Restores the <c> Directory </c> and <c> Name </c> properties used to instantiate the <c> iniFile </c> object </summary>
    Public Sub resetParams()
        Dir = InitDir
        Name = InitName
    End Sub

    ''' <summary> Returns the starting line number of each <c> iniSection </c> in the <c> iniFile </c> as a list of integers </summary>
    Public Function getLineNumsFromSections() As List(Of Integer)
        Dim outList As New List(Of Integer)
        For Each section In Sections.Values
            outList.Add(section.StartingLineNumber)
        Next
        Return outList
    End Function

    ''' <summary> Constructs an <c> iniFile </c> from a text stream </summary>
    ''' <param name="r"> A StreamReader object containing a Stream of an iniFile, remote or local </param>
    Public Sub New(r As StreamReader)
        Sections = New Dictionary(Of String, iniSection)
        Comments = New Dictionary(Of Integer, iniComment)
        LineCount = 1
        buildIniFromStream(r)
    End Sub

    ''' <summary> Processes a line in a ini file and updates the <c> iniFile </c> object meta data accordingly </summary>
    ''' <param name="currentLine"> The current line being read </param>
    ''' <param name="sectionToBeBuilt"> The lines comprising an <c> iniSection </c> whose construction is pending </param>
    ''' <param name="lineTrackingList"> The corresponding <c> Integers </c> representing line numbers associated with the <c> iniSection's </c> lines </param>
    Private Sub processiniLine(ByRef currentLine As String, ByRef sectionToBeBuilt As List(Of String), ByRef lineTrackingList As List(Of Integer))
        Select Case True
            Case currentLine.StartsWith(";")
                Comments.Add(Comments.Count, New iniComment(currentLine, LineCount))
            Case (Not currentLine.StartsWith("[") And Not currentLine.Trim = "") Or (currentLine.Trim <> "" And sectionToBeBuilt.Count = 0)
                updSec(sectionToBeBuilt, lineTrackingList, currentLine)
            Case currentLine.Trim <> "" And Not sectionToBeBuilt.Count = 0
                mkSection(sectionToBeBuilt, lineTrackingList)
                updSec(sectionToBeBuilt, lineTrackingList, currentLine)
        End Select
        LineCount += 1
    End Sub

    ''' <summary> Manages line and number tracking for <c> iniSections </c> whose construction is pending </summary>
    ''' <param name="secList"> The list of strings to be built into an <c> iniSection </c> </param>
    ''' <param name="lineList"> The corresponding list of integers representing line numbers associated with the <c> iniSection's </c> strings </param>
    ''' <param name="curLine"> The current line to be added to the section </param>
    Private Sub updSec(ByRef secList As List(Of String), ByRef lineList As List(Of Integer), curLine As String)
        secList.Add(curLine)
        lineList.Add(LineCount)
    End Sub

    ''' <summary> Populates the Sections and Comments of an iniFile using a StreamReader from either disk or the internet </summary>
    ''' <param name="r"> A StreamReader object containing a Stream of an iniFile, remote or local </param>
    Public Sub buildIniFromStream(ByRef r As StreamReader)
        Dim sectionToBeBuilt As New List(Of String)
        Dim lineTrackingList As New List(Of Integer)
        Do While r.Peek() > -1
            processiniLine(r.ReadLine, sectionToBeBuilt, lineTrackingList)
        Loop
        If sectionToBeBuilt.Count <> 0 Then mkSection(sectionToBeBuilt, lineTrackingList)
        r.Close()
    End Sub

    ''' <summary> Attempts to read a .ini file from disk and initialize the iniFile object </summary>
    Public Sub init()
        Try
            buildIniFromStream(New StreamReader(Me.Path()))
        Catch ex As Exception
            Console.WriteLine(ex.Message & Environment.NewLine & $"Failure occurred during iniFile construction at line: {LineCount} in {Name}")
            Console.ReadLine()
        End Try
    End Sub

    ''' <summary> Ensures that any call to an ini file on the system will be to a file that exists in a directory that exists </summary>
    Public Sub validate()
        gLog($"Validating {Name}", ascend:=True)
        Sections = New Dictionary(Of String, iniSection)
        Comments = New Dictionary(Of Integer, iniComment)
        ' Make sure both the file and the directory actually exist
        While Not exists()
            initModule("File Chooser", AddressOf printFileChooserMenu, AddressOf handleFileChooserInput)
            If Not exists() Then Exit Sub
        End While
        init()
        gLog($"ini created with {Sections.Count} sections", indent:=True, descend:=True)
    End Sub

    ''' <summary> Reorders the iniSections in an iniFile object to be in the same sorted state as a provided list of Strings </summary>
    ''' <param name="sortedSections"> The sorted state of the sections by name </param>
    Public Sub sortSections(sortedSections As strList)
        Dim tempFile As New iniFile
        sortedSections.Items.ForEach(Sub(sectionName) tempFile.Sections.Add(sectionName, Sections.Item(sectionName)))
        Me.Sections = tempFile.Sections
    End Sub

    ''' <summary> Find the line number of a comment by its string. Returns -1 if not found </summary>
    ''' <param name="com"> The string containing the comment text to be searched for </param>
    Public Function findCommentLine(com As String) As Integer
        For Each comment In Comments.Values
            If comment.Comment = com Then Return comment.LineNumber
        Next
        Return -1
    End Function

    ''' <summary> Returns the section names from the iniFile object as a list of Strings </summary>
    Public Function namesToStrList() As strList
        Dim out As New strList
        For Each section In Sections.Values
            out.add(section.Name)
        Next
        Return out
    End Function

    ''' <summary> Attempts to create a new iniSection object and add it to the iniFile </summary>
    ''' <param name="sectionToBeBuilt"> The list of strings in the iniSection </param>
    ''' <param name="lineTrackingList"> The list of line numbers associated with the given strings </param>
    Private Sub mkSection(sectionToBeBuilt As List(Of String), lineTrackingList As List(Of Integer))
        Try
            Dim sectionHolder As New iniSection(sectionToBeBuilt, lineTrackingList)
            Sections.Add(sectionHolder.Name, sectionHolder)
        Catch ex As Exception
            'This will catch entries whose names are identical (case sensitive), and ignore them 
            If ex.GetType.FullName = "System.ArgumentException" Then
                Dim lineErr = -1
                For Each section In Sections.Values
                    If section.getFullName = sectionToBeBuilt(0) Then
                        lineErr = section.StartingLineNumber
                        Exit For
                    End If
                Next
                Console.WriteLine($"Error: Duplicate section name detected: {sectionToBeBuilt(0)}")
                Console.WriteLine($"Line: {lineTrackingList(0)}")
                Console.WriteLine($"Duplicates the entry on line: {lineErr}")
                Console.WriteLine("This section will be ignored until it is given a unique name.")
                Console.WriteLine()
                Console.WriteLine("Press enter to continue.")
                Console.ReadLine()
            Else
                exc(ex)
            End If
        Finally
            sectionToBeBuilt.Clear()
            lineTrackingList.Clear()
        End Try
    End Sub

    ''' <summary> Prints the menu for the File Chooser submodule </summary>
    Public Sub printFileChooserMenu()
        printMenuTop({"Choose a file name, or open the directory chooser to choose a directory"})
        print(1, InitName, "Use the default name", InitName <> "")
        print(1, SecondName, "Use the default rename", SecondName <> "")
        print(1, "Directory Chooser", "Choose a new directory", trailingBlank:=True)
        print(0, $"Current Directory: {replDir(Dir)}")
        print(0, $"Current File:      {Name}", closeMenu:=True)
    End Sub

    ''' <summary> Handles the input for the File Chooser submodule </summary>
    ''' <param name="input"> The user's input </param>
    Public Sub handleFileChooserInput(input As String)
        Select Case True
            Case input = "0"
                exitModule()
            Case input = ""
                exitIfExists()
            Case input = "1" And InitName <> ""
                reName(InitName)
            Case (input = "1" And InitName = "") Or (input = "2" And SecondName <> "")
                reName(SecondName)
            Case (input = "2" And SecondName = "") Or (input = "3" And InitName <> "" And SecondName <> "")
                initModule("Directory Chooser", AddressOf printDirChooserMenu, AddressOf handleDirChooserInput)
            Case Else
                reName(input)
        End Select
    End Sub

    ''' <summary> Assigns the Name of the iniFile to the given name and stores the previous name in a temporary container </summary>
    ''' <param name="nname"> The new filename </param>
    Private Sub reName(nname As String)
        tmpRename = Name
        Name = nname
        exitIfExists(True)
    End Sub

    ''' <summary> Leaves the submodule as long as the inifile exists on disk in the case that mustExist is true </summary>
    ''' <param name="undoPendingRename"> Indicates that there's a pending rename that should be reverted if the renamed file doesn't exist <br/>Optional, Default: False </param>
    Private Sub exitIfExists(Optional undoPendingRename As Boolean = False)
        If Not exists() And mustExist Then
            setHeaderText($"{Name} does not exist", True)
            If undoPendingRename Then Name = tmpRename
        Else
            exitModule()
        End If
    End Sub

    ''' <summary> Prints the menu for the Directory Chooser submodule </summary>
    Public Sub printDirChooserMenu()
        printMenuTop({"Choose a directory", "Enter a new directory below or choose an option"})
        print(1, "Use default (default)", "Use the same folder as winapp2ool.exe")
        print(1, "Parent Folder", "Go up a level")
        print(1, "Current folder", "Continue using the same folder as below")
        print(0, $"Current Directory: {Dir}", leadingBlank:=True, closeMenu:=True)
    End Sub

    ''' <summary> Handles the user input for the Directory Chooser submodule </summary>
    ''' <param name="input"> The user's input </param>
    Public Sub handleDirChooserInput(input As String)
        Select Case True
            Case input = "0"
                exitModule()
            Case input = "1" Or input = ""
                Dir = Environment.CurrentDirectory
                exitModule()
            Case input = "2"
                Dir = Directory.GetParent(Dir).ToString
                exitModule()
            Case input = "3"
                exitModule()
            Case Else
                Dim tmpDir = Dir
                Dir = input
                If Not exists(False) Then
                    setHeaderText($"{Dir} does not exist", cHeader:=True)
                Else
                    exitModule()
                End If
        End Select
    End Sub

    ''' <summary> Returns true if the path (or optionally just the directory) of the inifile exists on the file system </summary>>
    ''' <param name="checkPath"> Indicates that the full path should be checked for existence. When <paramref name="checkPath"/> is False, only the directory will be checked
    ''' <br/> Optional, Default: True </param>
    Public Function exists(Optional checkPath As Boolean = True) As Boolean
        If checkPath Then Return File.Exists(Path)
        Return Directory.Exists(Dir)
    End Function
End Class