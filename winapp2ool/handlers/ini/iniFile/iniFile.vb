'    Copyright (C) 2018-2026 Hazel Ward
'
'    This file is a part of Winapp2ool
'
'    Winapp2ool is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    Winapp2ool is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with Winapp2ool.  If not, see <http://www.gnu.org/licenses/>.

Option Strict On
Imports System.IO
Imports System.Text

''' <summary>
''' Controls the order in which sections are emitted when serializing an <c> iniFile2 </c>
''' </summary>
Public Enum IniFileWriteFormat

    ''' <summary>
    ''' Sections are written in the order they were added (default behavior)
    ''' </summary>
    Insertion = 0

    ''' <summary>
    ''' Sections are written in case-insensitive alphabetical order by section name
    ''' </summary>
    Alphabetical = 1

End Enum

''' <summary>
''' An object representing a parsed .ini file with O(1) section lookup. Section names are unique
''' (case-insensitive): when a file repeats a section, we keep the first and drop the later one
''' along with its keys.
''' </summary>
Public Class iniFile2

    Implements IEnumerable(Of iniSection2)

    ''' <summary>
    ''' The directory on the filesystem in which the file can be found
    ''' </summary>
    Public ReadOnly Property Dir As String

    ''' <summary>
    ''' The name of the file on disk
    ''' </summary>
    Public ReadOnly Property Name As String

    ''' <summary>
    ''' Returns the full filesystem path of the file
    ''' </summary>
    Public Function Path() As String

        Return $"{Dir}\{Name}"

    End Function

    Private ReadOnly _ordered As New List(Of iniSection2)
    Private ReadOnly _byName As New Dictionary(Of String, iniSection2)(StringComparer.OrdinalIgnoreCase)
    Private ReadOnly _comments As New List(Of iniComment2)

    ''' <summary>
    ''' All comment lines encountered during parsing, in the order they appeared in the file.
    ''' Comment text includes the leading semicolon.
    ''' We capture comments for reading only. <see cref="ToString()"/> doesn't write them back.
    ''' </summary>
    Public ReadOnly Property Comments As List(Of iniComment2)
        Get
            Return _comments
        End Get
    End Property

    ''' <summary>
    ''' The number of sections in the file
    ''' </summary>
    Public ReadOnly Property Count As Integer

        Get

            Return _ordered.Count

        End Get

    End Property

    ''' <summary>
    ''' Returns whether a section with the given name exists in the file
    ''' </summary>
    ''' 
    ''' <param name="name">
    ''' The section name to search for (case-insensitive)
    ''' </param>
    Public Function Contains(name As String) As Boolean

        If name Is Nothing Then argIsNull(NameOf(name)) : Return False

        Return _byName.ContainsKey(name)

    End Function

    ''' <summary>
    ''' Returns the section with the given name, or <c> Nothing </c> if the file doesn't have one.
    ''' Use <see cref="GetOrCreateSection"/> when a missing section should be created.
    ''' </summary>
    ''' 
    ''' <param name="name">
    ''' The section name to look up (case-insensitive)
    ''' </param>
    Public Function GetSection(name As String) As iniSection2

        If name Is Nothing Then argIsNull(NameOf(name)) : Return Nothing

        Dim result As iniSection2 = Nothing
        _byName.TryGetValue(name, result)

        Return result

    End Function

    ''' <summary>Returns the section with the given name, creating and adding it if the file doesn't have one</summary>
    ''' 
    ''' <param name="name">
    ''' The section name to look up or create (case-insensitive)
    ''' </param>
    Public Function GetOrCreateSection(name As String) As iniSection2

        If name Is Nothing Then argIsNull(NameOf(name)) : Return Nothing

        Dim existing = GetSection(name)
        If existing IsNot Nothing Then Return existing

        Dim s As New iniSection2(name)
        AddSection(s)

        Return s

    End Function

    ''' <summary>
    ''' Removes the section with the given name from the file.
    ''' Does nothing if the section is not present.
    ''' </summary>
    '''
    ''' <param name="name">
    ''' The section name to remove (case-insensitive)
    ''' </param>
    Public Sub RemoveSection(name As String)

        If name Is Nothing Then argIsNull(NameOf(name)) : Return

        Dim s = GetSection(name)
        If s Is Nothing Then Return

        _ordered.Remove(s)
        _byName.Remove(name)

    End Sub

    ''' <summary>
    ''' Adds a section to the file. Duplicate names (case-insensitive) are silently ignored.
    ''' </summary>
    '''
    ''' <param name="section">
    ''' The section to add
    ''' </param>
    Public Sub AddSection(section As iniSection2)

        If section Is Nothing Then argIsNull(NameOf(section)) : Return

        If _byName.ContainsKey(section.Name) Then Return

        _ordered.Add(section)
        _byName.Add(section.Name, section)

    End Sub

    ''' <summary>
    ''' Creates a new <c> iniFile2 </c> with no sections. Outside this class, use
    ''' <see cref="FromFile"/>, <see cref="FromStream"/> or <see cref="Empty"/>.
    ''' </summary>
    Private Sub New(dir As String, name As String)

        Me.Dir = dir
        Me.Name = name

    End Sub

    ''' <summary>
    ''' Parses an ini file from a filesystem path. A missing file is reported through
    ''' <c> handleFileNotFoundException </c>. Any other read error, including a missing
    ''' directory, goes uncaught to the caller.
    ''' </summary>
    '''
    ''' <param name="path">
    ''' The path to an ini file. Everything after the last backslash becomes <c> Name </c>
    ''' and everything before it becomes <c> Dir </c>.
    ''' </param>
    '''
    ''' <returns>The parsed file, or one with no sections if the file doesn't exist</returns>
    Public Shared Function FromFile(path As String) As iniFile2

        If path Is Nothing Then argIsNull(NameOf(path)) : Return New iniFile2("", "")

        Dim slashPos = path.LastIndexOf("\"c)
        Dim dir = If(slashPos >= 0, path.Substring(0, slashPos), "")
        Dim name = If(slashPos >= 0, path.Substring(slashPos + 1), path)
        Dim f As New iniFile2(dir, name)

        Try

            Using reader As New StreamReader(path)
                f.ParseStream(reader)
            End Using

        Catch ex As FileNotFoundException

            handleFileNotFoundException(ex)

        End Try

        Return f

    End Function

    ''' <summary>
    ''' Returns a new <c> iniFile2 </c> with no sections and the given path components,
    ''' for building an output file programmatically
    ''' </summary>
    ''' <param name="dir">The directory component of the file path</param>
    ''' <param name="name">The filename component</param>
    Public Shared Function Empty(dir As String, name As String) As iniFile2
        Return New iniFile2(If(dir Is Nothing, "", dir), If(name Is Nothing, "", name))
    End Function

    ''' <summary>
    ''' Returns an ini file parsed from an already-open <c> StreamReader </c>.
    ''' We read to the end of the stream but leave closing it to the caller.
    ''' </summary>
    '''
    ''' <param name="r">
    ''' A <c> StreamReader </c> containing ini file content
    ''' </param>
    '''
    ''' <param name="dir">
    ''' The directory from which the stream originates <br /><br />
    ''' Optional, Default: <c> "" </c>
    ''' </param>
    '''
    ''' <param name="name">
    ''' The filename from which the stream originates <br /><br />
    ''' Optional, Default: <c> "" </c>
    ''' </param>
    Public Shared Function FromStream(r As StreamReader,
                                      Optional dir As String = "",
                                      Optional name As String = "") As iniFile2

        If r Is Nothing Then argIsNull(NameOf(r)) : Return New iniFile2(dir, name)

        Dim f As New iniFile2(dir, name)
        f.ParseStream(r)

        Return f

    End Function

    ''' <summary>
    ''' Reads <paramref name="r"/> line by line into sections, keys and comments. We skip blank
    ''' lines and drop any key that appears before the first section header. A repeated section
    ''' header starts a section that <see cref="AddSection"/> refuses, so its keys are lost too.
    ''' </summary>
    Private Sub ParseStream(r As StreamReader)

        Dim currentSection As iniSection2 = Nothing
        Dim lineNumber As Integer = 1

        Do While r.Peek() > -1

            Dim line As String = r.ReadLine()

            If line.Length = 0 OrElse line.TrimStart().Length = 0 Then
                ' skip blank lines
            ElseIf line.StartsWith(";", StringComparison.InvariantCulture) Then
                _comments.Add(New iniComment2(line, lineNumber))
            ElseIf line.StartsWith("[", StringComparison.InvariantCulture) Then
                Dim sectionName = line.TrimStart(CChar("[")).TrimEnd(CChar("]"))
                currentSection = New iniSection2(sectionName, lineNumber)
                AddSection(currentSection)
            ElseIf currentSection IsNot Nothing Then
                currentSection.AddKey(New iniKey2(line, lineNumber))
            End If
            lineNumber += 1
        Loop
    End Sub

    ''' <summary>
    ''' Writes <paramref name="text"/> to this file's <c> Path </c>, creating the directory and
    ''' file if they don't exist. If access is denied, we offer to restart winapp2ool elevated.
    ''' </summary>
    '''
    ''' <param name="text">The text to write</param>
    '''
    ''' <param name="condition">
    ''' Indicates whether to write at all <br /><br />
    ''' Optional, Default: <c> True </c>
    ''' </param>
    '''
    ''' <returns>
    ''' <c> True </c> if the file was written, <br />
    ''' <c> False </c> if the write was skipped or failed
    ''' </returns>
    Public Function OverwriteToFile(text As String, Optional condition As Boolean = True) As Boolean

        If Not condition Then Return False

        gLog($"Saving {Name}")

        Try

            If Not IO.File.Exists(Path()) Then
                gLog($"Target file doesn't exist")
                gLog($"Creating {Dir}")
                IO.Directory.CreateDirectory(Dir)
                gLog($"Creating {Path()}")
                Using IO.File.CreateText(Path()) : End Using
            End If

            Using file As New IO.StreamWriter(Path())
                file.Write(text)
            End Using
            gLog("  Save complete")
            Return True

        Catch ex As IO.IOException

            gLog("  Save failed")
            handleIOException(ex)

        Catch ex As UnauthorizedAccessException

            gLog("  Save failed")
            handleUnauthorizedAccessException(ex)
            offerElevatedRestart(Dir)

        End Try

        Return False

    End Function

    ''' <summary>Returns the file as it would appear on disk</summary>
    Public Overrides Function ToString() As String
        Return Serialize(_ordered)
    End Function

    ''' <summary>
    ''' Returns the file as it would appear on disk, with sections ordered
    ''' according to <paramref name="format"/>
    ''' </summary>
    '''
    ''' <param name="format">
    ''' Controls the section emission order. <see cref="IniFileWriteFormat.Insertion"/>
    ''' delegates to <see cref="ToString()"/>; <see cref="IniFileWriteFormat.Alphabetical"/>
    ''' sorts sections by name (case-insensitive, ordinal) before serializing
    ''' </param>
    Public Overloads Function ToString(format As IniFileWriteFormat) As String

        If format = IniFileWriteFormat.Insertion Then Return ToString()

        Return Serialize(_ordered.OrderBy(Function(s) s.Name, StringComparer.InvariantCultureIgnoreCase))

    End Function

    ''' <summary>
    ''' Serializes a sequence of <c> iniSection2 </c> objects into ini file text, with a blank
    ''' line between sections. The text ends with the newline that closes the last section.
    ''' </summary>
    '''
    ''' <param name="sections">
    ''' The sections to serialize, in the order they should appear
    ''' </param>
    Private Shared Function Serialize(sections As IEnumerable(Of iniSection2)) As String

        Dim sb As New StringBuilder()
        Dim first = True

        For Each section In sections

            If Not first Then sb.Append(Environment.NewLine)
            sb.Append(section.ToString())
            first = False

        Next

        Return sb.ToString()

    End Function

    ''' <summary>Returns an enumerator over the sections in the order they were added</summary>
    Public Function GetEnumerator() As IEnumerator(Of iniSection2) Implements IEnumerable(Of iniSection2).GetEnumerator
        Return _ordered.GetEnumerator()
    End Function

    Private Function GetEnumeratorNonGeneric() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
        Return _ordered.GetEnumerator()
    End Function

End Class
