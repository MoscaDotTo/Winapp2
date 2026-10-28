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
''' Represents a winapp2.ini file, with its entries grouped into the file's sections.
''' Built from an <c> iniFile </c>.
''' </summary>
Public Class winapp2file

    ''' <summary>
    ''' The section header labels in winapp2.ini order.
    ''' Empty string is the unlabeled main section (last).
    ''' </summary>
    Public Shared ReadOnly Property FileSectionHeaders As New List(Of String) From {
        "Language entries",
        "Potentially very long scan time (and also dangerous) entries",
        "Dangerous entries",
        ""
    }

    ' Category routing values, in the same order as the labeled FileSectionHeaders
    Private Shared ReadOnly LeadingCategories As New List(Of String) From {
        "Language Files", "Dangerous Long", "Dangerous"
    }

    Private Const RepoUrl As String = "https://github.com/MoscaDotTo/Winapp2"
    Private Const LatestDownloadUrl As String = RepoUrl & "/releases/latest/download/"

    ' One category list per FileSectionHeaders entry, at the same index
    Private ReadOnly _categories As New List(Of List(Of winapp2entry))

    ' Dirty-cache for flat Entries view
    Private _entriesCache As List(Of winapp2entry) = Nothing
    Private _entriesDirty As Boolean = True

    ''' <summary>All entries in winapp2.ini order (categories in sequence, main section last)</summary>
    Public ReadOnly Property Entries As IReadOnlyList(Of winapp2entry)
        Get
            If _entriesDirty OrElse _entriesCache Is Nothing Then
                _entriesCache = _categories.SelectMany(Function(c) c).ToList()
                _entriesDirty = False
            End If
            Return _entriesCache
        End Get
    End Property

    ''' <summary>
    ''' The flavor the file's <c> ; Flavor: </c> preamble line names, or <c> Nothing </c> when
    ''' it has no such line or names an unknown flavor. A file without a flavor gets the header
    ''' that names none, since guessing would mislabel an older copy of any flavor. Only a run
    ''' told the flavor on the command line sets it.
    ''' </summary>
    Public Property Flavor As WinappFlavor?

    ''' <summary>
    ''' The version comment string, e.g. "; Version: 260219". It's <c> ; Version: 000000 </c>
    ''' when the source file's first comment isn't a version line.
    ''' </summary>
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
    ''' The per-category entry lists in <c> FileSectionHeaders </c> order.
    ''' Each inner list corresponds to one winapp2.ini section (the last is the unlabeled main section).
    ''' </summary>
    Public ReadOnly Property Categories As IReadOnlyList(Of IReadOnlyList(Of winapp2entry))
        Get
            Return _categories.Select(Function(c) DirectCast(c, IReadOnlyList(Of winapp2entry))).ToList()
        End Get
    End Property

    ''' <summary>
    ''' Creates a new <c> winapp2file </c> from <paramref name="file"/>. Each entry goes to the
    ''' section its LangSecRef, or failing that its Section key, names. Entries with neither, or
    ''' with a category no section claims, go to the unlabeled main section. The exception is a
    ''' category starting with "Dangerous", which goes to the Dangerous section.
    ''' </summary>
    '''
    ''' <param name="file">A winapp2.ini format <c> iniFile </c></param>
    '''
    ''' <param name="useTodaysDate">
    ''' Indicates whether to stamp <c> Version </c> with today's date (yyMMdd) instead of
    ''' reading it from the file <br /><br />
    ''' Optional, Default: <c> False </c>
    ''' </param>
    Public Sub New(file As iniFile, Optional useTodaysDate As Boolean = False)

        If file Is Nothing Then argIsNull(NameOf(file)) : Return

        Dir  = file.Dir
        Name = file.Name

        For i = 0 To FileSectionHeaders.Count - 1
            _categories.Add(New List(Of winapp2entry))
        Next

        Flavor = ReadFlavor(file)

        Version = DetermineVersionString(file, useTodaysDate)

        For Each section In file
            Dim entry As New winapp2entry(section)
            Dim ind = GetCategoryIndex(entry, LeadingCategories)
            _categories(ind).Add(entry)
        Next

    End Sub

    Private Shared Function DetermineVersionString(file As iniFile,
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

    ''' <summary>
    ''' Returns the flavor named by the first known <c> ; Flavor: </c> line in the preamble: the
    ''' comments on consecutive lines from line 1, which is the block the header and
    ''' <see cref="ToIni"/> write. Section line numbers can't mark where it ends, since sections
    ''' Transmute adds keep the line numbers they had in their own file.
    ''' </summary>
    '''
    ''' <param name="file">A winapp2.ini format <c> iniFile </c></param>
    '''
    ''' <returns>
    ''' The flavor, <br />
    ''' <c> Nothing </c> if no preamble line names a known one
    ''' </returns>
    Private Shared Function ReadFlavor(file As iniFile) As WinappFlavor?

        Dim expectedLine = 1

        For Each comment In file.Comments

            If comment.LineNumber <> expectedLine Then Exit For
            expectedLine += 1

            Dim flavor = WinappFlavorInfo.ParseFlavorLine(comment.Text)
            If flavor.HasValue Then Return flavor

        Next

        Return Nothing

    End Function

    Private Shared Function GetCategoryIndex(entry As winapp2entry,
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
    ''' Returns all entries as a single <c> iniFile </c> in winapp2.ini order. <br /> <br />
    '''
    ''' We re-emit the two pieces of preamble state this class carries, the version string and
    ''' the flavor, as comments on the returned file. Without them a round trip through the
    ''' winapp2 layer silently resets the version to <c> 000000 </c> and forgets the flavor,
    ''' since this class' constructor reads both back off <c> iniFile.Comments </c>, and Diff
    ''' reads the version there too
    ''' </summary>
    Public Function ToIni() As iniFile

        Dim out = iniFile.Empty(Dir, Name)

        out.Comments.Add(New iniComment(Version, 1))

        If Flavor.HasValue Then
            out.Comments.Add(New iniComment($"; Flavor: {WinappFlavorInfo.Lookup(Flavor.Value).DisplayName}", 2))
        End If

        For Each entry In Entries
            out.AddSection(entry.ToIniSection())
        Next

        Return out

    End Function

    ''' <summary>
    ''' Adds an entry to the appropriate category based on its LangSecRef or Section key
    ''' </summary>
    ''' <param name="entry">The entry to add</param>
    Public Sub AddEntry(entry As winapp2entry)

        If entry Is Nothing Then argIsNull(NameOf(entry)) : Return

        Dim ind = GetCategoryIndex(entry, LeadingCategories)
        _categories(ind).Add(entry)
        _entriesDirty = True

    End Sub

    ''' <summary>
    ''' Removes an entry from the file. Searches all categories.
    ''' </summary>
    ''' <param name="entry">The entry to remove</param>
    Public Sub RemoveEntry(entry As winapp2entry)

        If entry Is Nothing Then argIsNull(NameOf(entry)) : Return

        For Each cat In _categories
            If cat.Remove(entry) Then
                _entriesDirty = True
                Return
            End If
        Next

    End Sub

    ''' <summary>
    ''' Sorts entries within each category using winapp2 ordering:
    ''' <c> - </c> is treated as whitespace and embedded numbers are padded so that
    ''' <c> Item2 </c> sorts before <c> Item10 </c>. 
    ''' </summary>
    Public Sub SortEntries()

        For Each cat In _categories

            If cat.Count < 2 Then Continue For

            Dim names As New List(Of String)
            For Each entry In cat : names.Add(entry.Name) : Next

            Dim sortedNames = replaceAndSort(names, "-", "  ")

            Dim nameToEntries As New Dictionary(Of String, Queue(Of winapp2entry))(StringComparer.OrdinalIgnoreCase)
            For Each entry In cat

                If Not nameToEntries.ContainsKey(entry.Name) Then nameToEntries(entry.Name) = New Queue(Of winapp2entry)
                nameToEntries(entry.Name).Enqueue(entry)

            Next

            cat.Clear()
            For Each entryName In sortedNames : cat.Add(nameToEntries(entryName).Dequeue()) : Next

        Next

        _entriesDirty = True

    End Sub

    ''' <summary>
    ''' Returns each entry name that appears more than once (case-insensitive), once per extra
    ''' copy. <c> iniFile </c> drops repeated sections while parsing, so only entries added
    ''' through <see cref="AddEntry"/> can repeat.
    ''' </summary>
    Public Function FindDuplicateEntryNames() As IReadOnlyList(Of String)

        Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Dim dupes As New List(Of String)

        For Each entry In Entries
            If Not seen.Add(entry.Name) Then dupes.Add(entry.Name)
        Next

        Return dupes

    End Function

    ''' <summary>
    ''' Returns the complete winapp2.ini text: the header, then each labeled section that has
    ''' entries, wrapped in comments carrying its name and entry count, then the main section.
    ''' <br /><br />
    '''
    ''' The header's first two lines, the version and the entry count, are read by the
    ''' updater, Diff and the build, so their format is fixed. A file with a
    ''' <see cref="Flavor"/> gets a <c> ; Flavor: </c> third line, a description and links to
    ''' its own release assets. One without gets the same header minus those, linking to the
    ''' latest release.
    ''' </summary>
    Public Function ToWinapp2String() As String

        Dim builder As New StringBuilder()

        builder.AppendLine(Version)
        builder.AppendLine($"; # of entries: {Count:#,###}")

        If Flavor.HasValue Then

            Dim info = WinappFlavorInfo.Lookup(Flavor.Value)

            builder.AppendLine($"; Flavor: {info.DisplayName}")
            builder.AppendLine(";")
            For Each line In info.Description : builder.AppendLine($"; {line}") : Next
            builder.AppendLine(";")
            builder.AppendLine("; Latest copy of this file:")
            builder.AppendLine($";   {LatestDownloadUrl}{info.AssetName}")
            builder.AppendLine("; Changelog for the latest release:")
            builder.AppendLine($";   {LatestDownloadUrl}{info.ChangelogAssetName}")
            builder.AppendLine("; Each changelog compares one release with the one before it. To see what changed since your")
            builder.AppendLine("; own copy, compare it with the latest using winapp2ool's Diff.")

        Else

            builder.AppendLine(";")
            builder.AppendLine("; Latest copy of winapp2.ini, in every flavor:")
            builder.AppendLine($";   {RepoUrl}/releases/latest")

        End If

        builder.AppendLine(";")
        builder.AppendLine("; About winapp2.ini, and how to install it for your cleaner:")
        builder.AppendLine($";   {RepoUrl}")
        builder.AppendLine(";")
        builder.AppendLine("; Winapp2ool installs, updates and trims this file:")
        builder.AppendLine($";   {RepoUrl}/tree/master/winapp2ool")
        builder.AppendLine($";   {LatestDownloadUrl}winapp2ool.exe")
        builder.AppendLine(";")
        builder.AppendLine("; winapp2.ini is licensed under CC BY-SA 4.0. You may share and adapt it if you credit the")
        builder.AppendLine("; winapp2 project and release your changes under the same license:")
        builder.AppendLine($";   {RepoUrl}/blob/master/License.md")
        builder.AppendLine(";")
        builder.AppendLine("; This file is built automatically from the winapp2 repository. To add or fix an entry, see:")
        builder.AppendLine($";   {RepoUrl}/blob/master/CONTRIBUTING.md")

        For i = 0 To _categories.Count - 2

            If _categories(i).Count = 0 Then Continue For

            builder.AppendLine(";")
            builder.AppendLine($"; {FileSectionHeaders(i)} ({_categories(i).Count})")
            builder.AppendLine()
            builder.AppendLine(CategoriesToString(_categories(i)))
            builder.AppendLine($"; End of {FileSectionHeaders(i)}")

        Next

        builder.AppendLine()

        If _categories.Last.Count > 0 Then builder.Append(CategoriesToString(_categories.Last))

        Return builder.ToString()

    End Function

    ''' <summary>
    ''' Serializes a list of entries with a blank line between each. Each entry's text ends
    ''' with a newline, so the result does too.
    ''' </summary>
    Private Shared Function CategoriesToString(entries As List(Of winapp2entry)) As String

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
