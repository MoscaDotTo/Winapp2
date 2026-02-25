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
Imports System.Globalization
Imports System.Text

''' <summary>
''' Represents a winapp2.ini file with a single flat entry list and no parallel structure.
''' Built from an <c>iniFile2</c>.
''' </summary>
Public Class winapp2file2

    ''' <summary>
    ''' The 12 section header labels in winapp2.ini order.
    ''' Empty string is the unlabelled main section (last).
    ''' </summary>
    Public Shared ReadOnly Property FileSectionHeaders As New List(Of String) From {
        "Chrome/Chromium based browsers",
        "Edge Chromium",
        "Vivaldi", "Brave",
        "Opera",
        "Firefox/Mozilla based browsers",
        "Thunderbird", "Internet Explorer",
        "Language entries",
        "Potentially very long scan time (and also dangerous) entries",
        "Dangerous entries",
        ""
    }

    ' Category routing values, in the same order as FileSectionHeaders indices 0–10
    Private Shared ReadOnly LeadingCategories As New List(Of String) From {
        "3029", "3006", "3033", "3034", "3027", "3026", "3030", "3001",
        "Language Files", "Dangerous Long", "Dangerous"
    }

    ' 12 category lists (index matches FileSectionHeaders)
    Private ReadOnly _categories As New List(Of List(Of winapp2entry2))

    ' Dirty-cache for flat Entries view
    Private _entriesCache As List(Of winapp2entry2) = Nothing
    Private _entriesDirty As Boolean = True

    ''' <summary>All entries in winapp2.ini order (categories in sequence, main section last)</summary>
    Public ReadOnly Property Entries As IReadOnlyList(Of winapp2entry2)
        Get
            If _entriesDirty OrElse _entriesCache Is Nothing Then
                _entriesCache = _categories.SelectMany(Function(c) c).ToList()
                _entriesDirty = False
            End If
            Return _entriesCache
        End Get
    End Property

    ''' <summary>Whether this represents the Non-CCleaner variant of winapp2.ini</summary>
    Public ReadOnly Property IsNCC As Boolean

    ''' <summary>The version comment string, e.g. "; Version: 260219"</summary>
    Public ReadOnly Property Version As String

    ''' <summary>The directory from which the source file was loaded</summary>
    Public ReadOnly Property Dir As String

    ''' <summary>The filename of the source file</summary>
    Public ReadOnly Property Name As String

    ''' <summary>Total number of entries across all categories</summary>
    Public ReadOnly Property Count As Integer
        Get
            Return _categories.Sum(Function(c) c.Count)
        End Get
    End Property

    ''' <summary>
    ''' Creates a <c>winapp2file2</c> from an <c>iniFile2</c>
    ''' </summary>
    ''' <param name="file">A winapp2.ini format <c>iniFile2</c></param>
    ''' <param name="useTodaysDate">When True, Version is set to today's date in YYMMDD format</param>
    Public Sub New(file As iniFile2, Optional useTodaysDate As Boolean = False)

        If file Is Nothing Then argIsNull(NameOf(file)) : Return

        Dir  = file.Dir
        Name = file.Name

        For i = 0 To FileSectionHeaders.Count - 1
            _categories.Add(New List(Of winapp2entry2))
        Next

        IsNCC = file.Comments.Any(Function(c) c.Text.Contains(
            "This is the non-CCleaner version of Winapp2"))

        Version = DetermineVersionString(file, useTodaysDate)

        For Each section In file
            Dim entry As New winapp2entry2(section)
            Dim ind = GetCategoryIndex(entry, LeadingCategories)
            _categories(ind).Add(entry)
        Next

    End Sub

    Private Shared Function DetermineVersionString(file As iniFile2,
                                                   useTodaysDate As Boolean) As String

        If useTodaysDate Then
            Return $"; Version: {DateTime.Now.ToString("yyMMdd", CultureInfo.InvariantCulture)}"
        End If

        If file.Comments.Count > 0 AndAlso
           file.Comments(0).Text.ToUpperInvariant().Contains("VERSION") Then
            Return file.Comments(0).Text
        End If

        Return "; Version: 000000"

    End Function

    Private Shared Function GetCategoryIndex(entry As winapp2entry2,
                                             categoryValues As List(Of String)) As Integer

        If entry.LangSecRef.Count = 0 AndAlso entry.SectionKey.Count = 0 Then
            Return categoryValues.Count
        End If

        Dim categoryKey As String = If(entry.LangSecRef.Count > 0,
                                       entry.LangSecRef(0).Value,
                                       entry.SectionKey(0).Value)

        Dim ind = categoryValues.IndexOf(categoryKey)

        If ind = -1 AndAlso categoryKey.ToUpperInvariant().StartsWith(
                "DANGEROUS", StringComparison.InvariantCulture) Then
            ind = categoryValues.Count - 1
        End If

        If ind = -1 Then ind = categoryValues.Count

        Return ind

    End Function

    ''' <summary>
    ''' Returns all entries as a single <c>iniFile2</c> in winapp2.ini order
    ''' </summary>
    Public Function ToIni() As iniFile2

        Dim out = iniFile2.Empty(Dir, Name)

        For Each entry In Entries
            out.AddSection(entry.ToIniSection())
        Next

        Return out

    End Function

    ''' <summary>
    ''' Adds an entry to the appropriate category based on its LangSecRef or Section key
    ''' </summary>
    ''' <param name="entry">The entry to add</param>
    Public Sub AddEntry(entry As winapp2entry2)

        If entry Is Nothing Then argIsNull(NameOf(entry)) : Return

        Dim ind = GetCategoryIndex(entry, LeadingCategories)
        _categories(ind).Add(entry)
        _entriesDirty = True

    End Sub

    ''' <summary>
    ''' Removes an entry from the file. Searches all categories.
    ''' </summary>
    ''' <param name="entry">The entry to remove</param>
    Public Sub RemoveEntry(entry As winapp2entry2)

        If entry Is Nothing Then argIsNull(NameOf(entry)) : Return

        For Each cat In _categories
            If cat.Remove(entry) Then
                _entriesDirty = True
                Return
            End If
        Next

    End Sub

    ''' <summary>
    ''' Builds and returns the complete winapp2.ini text including preamble comments,
    ''' replicating the output of the legacy <c>winapp2file.winapp2string()</c>
    ''' </summary>
    Public Function ToWinapp2String() As String

        Dim builder As New StringBuilder()
        Dim fileName = If(IsNCC, "Winapp2 (Non-CCleaner version)", "Winapp2")
        Dim licLink  = If(IsNCC,
            "https://github.com/MoscaDotTo/Winapp2/blob/master/Non-CCleaner/License.md",
            "https://github.com/MoscaDotTo/Winapp2/blob/master/License.md")

        builder.AppendLine(Version)
        builder.AppendLine($"; # of entries: {Count:#,###}")
        builder.AppendLine(";")
        builder.AppendLine($"; {fileName}.ini is fully licensed under the CC-BY-SA-4.0 license agreement. Please refer to our license agreement before using Winapp2.ini: {licLink}")
        builder.AppendLine($"; You may copy, modify, remix, share, show, and transmit {fileName} but you must redistribute under the same license and you must attribute the original work to the winapp2 project")
        builder.AppendLine(";")

        If IsNCC Then
            builder.AppendLine("; This is the non-CCleaner version of Winapp2 that contains extra entries that were removed due to them being added to CCleaner.")
            builder.AppendLine("; Do not use this file for CCleaner as the extra cleaners may cause conflicts with CCleaner.")
        End If

        builder.AppendLine($"; You can get the latest Winapp2.ini here: https://github.com/MoscaDotTo/Winapp2")
        builder.AppendLine($"; Any contributions are appreciated. Please refer to our ReadMe to learn to make your own entries here: https://github.com/MoscaDotTo/Winapp2/blob/master/README.md")
        builder.AppendLine(";")
        builder.AppendLine($"; Try out Winapp2ool for many useful additional features including updating and trimming Winapp2.ini: https://github.com/MoscaDotTo/Winapp2/raw/master/winapp2ool/bin/Release/winapp2ool.exe")
        builder.AppendLine($"; You can find the Winapp2ool ReadMe here: https://github.com/MoscaDotTo/Winapp2/blob/master/winapp2ool/Readme.md")

        For i = 0 To _categories.Count - 2

            If _categories(i).Count = 0 Then Continue For

            builder.AppendLine(";")
            builder.AppendLine($"; {FileSectionHeaders(i)} ({_categories(i).Count})")
            builder.AppendLine()
            builder.AppendLine(CategoriesToString(_categories(i)))
            builder.AppendLine($"; End of {FileSectionHeaders(i)}")

        Next

        builder.AppendLine()

        If _categories.Last.Count > 0 Then
            builder.Append(CategoriesToString(_categories.Last))
        End If

        Return builder.ToString()

    End Function

    ''' <summary>
    ''' Serialises a list of entries to text, matching the legacy <c>iniFile.toString</c> output format:
    ''' each section ends with a trailing newline, sections separated by a blank line
    ''' </summary>
    Private Shared Function CategoriesToString(entries As List(Of winapp2entry2)) As String

        If entries.Count = 0 Then Return ""
        If entries.Count = 1 Then Return entries(0).ToIniSection().ToString()

        Dim out As String = ""

        For i = 0 To entries.Count - 2
            out += entries(i).ToIniSection().ToString() & Environment.NewLine
        Next

        out += entries.Last.ToIniSection().ToString()

        Return out

    End Function

End Class
