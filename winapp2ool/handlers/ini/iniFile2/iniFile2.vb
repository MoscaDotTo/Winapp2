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

''' <summary>
''' An object representing a parsed .ini file with O(1) section lookup
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
    ''' Comments are captured for reading only — they are not written back by <c>ToString</c>.
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
    ''' Returns the section with the given name, or <c>Nothing</c> if not found. <br /> <br />
    ''' Unlike <c> iniFile.getSection </c>, this never returns a "phantom" empty section. <br /> <br />
    ''' 
    ''' Use <c> iniFile2.GetOfCreateSection </c> when get-or-create semantics are needed 
    ''' while replacing <c> iniFile.getSection </c>
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

    ''' <summary>
    ''' Returns the section with the given name, creating and adding it if absent. <br />
    ''' 
    ''' Use this instead of <c> iniFile2.GetSection </c> when get-or-create semantics are needed 
    ''' while replacing <c> iniFile.getSection </c> 
    ''' </summary>
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
    ''' 
    ''' </summary>
    ''' 
    ''' <param name="dir">
    ''' 
    ''' </param>
    ''' 
    ''' <param name="name">
    ''' 
    ''' </param>
    Private Sub New(dir As String, name As String)

        Me.Dir = dir
        Me.Name = name

    End Sub

    ''' <summary>
    ''' Parses an ini file from a filesystem path
    ''' </summary>
    ''' 
    ''' <param name="path">
    ''' The absolute path to an ini file
    ''' </param>
    Public Shared Function FromFile(path As String) As iniFile2

        If path Is Nothing Then argIsNull(NameOf(path)) : Return New iniFile2("", "")

        Dim slashPos = path.LastIndexOf("\"c)
        Dim dir = If(slashPos >= 0, path.Substring(0, slashPos), "")
        Dim name = If(slashPos >= 0, path.Substring(slashPos + 1), path)
        Dim f As New iniFile2(dir, name)

        Try

            Dim reader As New StreamReader(path)
            f.ParseStream(reader)
            reader.Close()

        Catch ex As FileNotFoundException

            handleFileNotFoundException(ex)

        End Try

        Return f

    End Function

    ''' <summary>
    ''' Creates an empty <c>iniFile2</c> with the given path components,
    ''' for use when building an output file programmatically
    ''' </summary>
    ''' <param name="dir">The directory component of the file path</param>
    ''' <param name="name">The filename component</param>
    Public Shared Function Empty(dir As String, name As String) As iniFile2
        Return New iniFile2(If(dir Is Nothing, "", dir), If(name Is Nothing, "", name))
    End Function

    ''' <summary>
    ''' Parses an ini file from an already-open <c> StreamReader </c>
    ''' </summary>
    ''' 
    ''' <param name="r">
    ''' A <c> StreamReader </c> containing ini file content
    ''' </param>
    ''' 
    ''' <param name="dir">
    ''' The directory from which the stream originates 
    ''' </param>
    ''' 
    ''' <param name="name">
    ''' The filename from which the stream originates 
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
    ''' 
    ''' </summary>
    ''' <param name="r">
    ''' 
    ''' </param>
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

    ''' <summary>Returns the file as it would appear on disk</summary>
    Public Overrides Function ToString() As String
        If _ordered.Count = 0 Then Return ""
        If _ordered.Count = 1 Then Return _ordered(0).ToString()
        Dim out As String = ""
        For i = 0 To _ordered.Count - 2
            out += _ordered(i).ToString() & Environment.NewLine
        Next
        out += _ordered.Last.ToString()
        Return out
    End Function

    Public Function GetEnumerator() As IEnumerator(Of iniSection2) Implements IEnumerable(Of iniSection2).GetEnumerator
        Return _ordered.GetEnumerator()
    End Function

    Private Function GetEnumeratorNonGeneric() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
        Return _ordered.GetEnumerator()
    End Function

End Class
