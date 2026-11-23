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
Imports System.IO
Imports System.Text

''' <summary>
''' Writes the do-not-edit header that every generator puts at the top of its build artifacts
''' (EntryBuilder's and BrowserBuilder's output under <c> Assembler\Generated </c>),
''' and holds the repository URL and the repository-relative source folders those headers link.
''' This is the only place either is spelled. <see cref="RepoSourceFolder"/> reads EntryBuilder's
''' link off the folder a run read instead, falling back to <see cref="EntryBuilderSources"/>.
''' <br /><br />
'''
''' A header names the generator, links the sources its entries come from on <c> master </c>,
''' and points to the latest release (or archived.ini, for an archive) and the contributor
''' guide. It carries no license: the artifacts are repository files, and the published files
''' carry the license.
''' </summary>
Public Module ArtifactHeader

    ''' <summary>
    ''' The winapp2 repository on GitHub
    ''' </summary>
    Public Const RepoUrl As String = "https://github.com/MoscaDotTo/Winapp2"

    ''' <summary>
    ''' The folder holding every generator's sources, relative to the repository root
    ''' </summary>
    Public Const GeneratorSources As String = "Assembler/Sources"

    ''' <summary>
    ''' EntryBuilder's source folder, relative to the repository root
    ''' </summary>
    Public Const EntryBuilderSources As String = GeneratorSources & "/EntryBuilder"

    ''' <summary>
    ''' BrowserBuilder's source folder, relative to the repository root
    ''' </summary>
    Public Const BrowserBuilderSources As String = GeneratorSources & "/BrowserBuilder"

    ''' <summary>
    ''' winapp3.ini's source folder, relative to the repository root, as the published
    ''' winapp3.ini links it
    ''' </summary>
    Public Const Winapp3Sources As String = GeneratorSources & "/Winapp3"

    ''' <summary>
    ''' The one label of the link to the latest release, in artifact headers and in the header
    ''' of a published file with no flavor or marker
    ''' </summary>
    Public Const LatestReleaseLabel As String = "Latest copy of every published file:"

    ''' <summary>
    ''' The name of the repository folder that holds the generator sources, matched in any case
    ''' </summary>
    Private Const RepoAnchorFolder As String = "Assembler"

    ''' <summary>
    ''' Returns <paramref name="sourceDir"/>'s path from its nearest ancestor named
    ''' <c> Assembler </c>, that folder included, with forward slashes, so
    ''' <c> C:\repo\Assembler\Sources\Winapp3 </c> gives <c> Assembler/Sources/Winapp3 </c>
    ''' whatever the checkout root or the working directory. Each folder is spelled as it is on
    ''' disk, since GitHub's <c> tree/ </c> links are case-sensitive and a hand-typed
    ''' <c> -1d \sources\winapp3 </c> would otherwise link nowhere.
    ''' </summary>
    '''
    ''' <param name="sourceDir">
    ''' The folder a generator read its sources from
    ''' </param>
    '''
    ''' <param name="fallback">
    ''' The repository-relative folder to return when <paramref name="sourceDir"/> has no
    ''' <c> Assembler </c> ancestor, as a folder outside the repository doesn't
    ''' </param>
    Public Function RepoSourceFolder(sourceDir As String,
                                     fallback As String) As String

        If String.IsNullOrEmpty(sourceDir) Then

            Return fallback

        End If

        Dim fullPath = Path.GetFullPath(sourceDir)

        If fullPath.Length > Path.GetPathRoot(fullPath).Length Then

            fullPath = fullPath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)

        End If

        ' Find the anchor by name before touching the disk, so a folder outside the repository
        ' never lists anything
        Dim chain As New List(Of DirectoryInfo)
        Dim current As New DirectoryInfo(fullPath)

        Do While current.Parent IsNot Nothing

            chain.Insert(0, current)
            current = current.Parent

            If current.Name.Equals(RepoAnchorFolder, StringComparison.OrdinalIgnoreCase) Then

                chain.Insert(0, current)
                Return String.Join("/", chain.Select(Function(folder) OnDiskName(folder)))

            End If

        Loop

        Return fallback

    End Function

    ''' <summary>
    ''' Returns <paramref name="folder"/>'s name as its parent lists it, or as given when it
    ''' doesn't exist or its parent refuses to be listed
    ''' </summary>
    '''
    ''' <param name="folder">
    ''' A folder that isn't a drive root
    ''' </param>
    Private Function OnDiskName(folder As DirectoryInfo) As String

        Try

            If Not folder.Exists Then

                Return folder.Name

            End If

            Dim listed = folder.Parent.GetDirectories(folder.Name).FirstOrDefault(Function(d) d.Name.Equals(folder.Name, StringComparison.OrdinalIgnoreCase))

            Return If(listed Is Nothing, folder.Name, listed.Name)

        Catch ex As UnauthorizedAccessException

            Return folder.Name

        Catch ex As IOException

            Return folder.Name

        Catch ex As Security.SecurityException

            Return folder.Name

        End Try

    End Function

    ''' <summary>
    ''' The format of every entry count in a header, the published winapp2.ini's included.
    ''' It prints <c> 0 </c> for an empty file, where <c> #,### </c> printed nothing.
    ''' </summary>
    Private Const EntryCountFormat As String = "#,##0"

    ''' <summary>
    ''' Returns the <c> ; # of entries: </c> line that opens every header
    ''' </summary>
    '''
    ''' <param name="count">
    ''' The number of entries in the file
    ''' </param>
    Public Function EntryCountLine(count As Integer) As String

        Return $"; # of entries: {count.ToString(EntryCountFormat, CultureInfo.CurrentCulture)}"

    End Function

    ''' <summary>
    ''' Returns the link to a source file (a <c> blob/ </c> link) or folder (a <c> tree/ </c> link)
    ''' on <c> master </c>. A <c> # </c> is written as <c> %23 </c>, since a bare one would start
    ''' the URL's fragment and the link would open the folder instead of <c> #.ini </c>.
    ''' </summary>
    '''
    ''' <param name="sourcePath">
    ''' The source's path relative to the repository root, with forward slashes
    ''' </param>
    '''
    ''' <param name="sourceIsFolder">
    ''' Indicates whether <paramref name="sourcePath"/> names a folder rather than a file
    ''' </param>
    Public Function SourceLink(sourcePath As String,
                               sourceIsFolder As Boolean) As String

        Dim kind = If(sourceIsFolder, "tree", "blob")

        Return $"{RepoUrl}/{kind}/master/{sourcePath.Replace("#", "%23")}"

    End Function

    ''' <summary>
    ''' Returns an artifact's header, ending with the blank line that separates it from the
    ''' first section, so a caller writes the header followed by the file's sections. No line
    ''' ends in whitespace. <br /><br />
    '''
    ''' Header text follows the voice of the hand-written source headers: one sentence per line,
    ''' no closing periods, no contractions, and a capital only where a sentence starts. A line
    ''' never opens with a lowercase file name, so the header says "This file" rather than naming
    ''' it.
    ''' </summary>
    '''
    ''' <param name="generatorName">
    ''' The module that writes the artifact, such as <c> EntryBuilder </c>
    ''' </param>
    '''
    ''' <param name="entryCount">
    ''' The number of entries in the artifact
    ''' </param>
    '''
    ''' <param name="sourcePath">
    ''' The path of the artifact's sources relative to the repository root: the one source file
    ''' of a per-letter artifact, or the source folder of a single-file one
    ''' </param>
    '''
    ''' <param name="sourceIsFolder">
    ''' Indicates whether <paramref name="sourcePath"/> names a folder rather than a file
    ''' </param>
    '''
    ''' <param name="description">
    ''' A line saying what the artifact holds, written above the source link. <br /><br />
    ''' Optional, Default: <c> "" </c> (no line)
    ''' </param>
    '''
    ''' <param name="isArchive">
    ''' Indicates whether the artifact is a piece of <c> archived.ini </c>, which links that
    ''' release asset instead of the latest release <br /><br />
    ''' Optional, Default: <c> False </c>
    ''' </param>
    Public Function Text(generatorName As String,
                         entryCount As Integer,
                         sourcePath As String,
                         sourceIsFolder As Boolean,
                         Optional description As String = "",
                         Optional isArchive As Boolean = False) As String

        Dim sb As New StringBuilder()
        sb.AppendLine(EntryCountLine(entryCount))
        sb.AppendLine($"; This file is generated by winapp2ool's {generatorName}")
        sb.AppendLine("; Do not edit it directly")
        sb.AppendLine("; The next build overwrites any changes made to it")
        sb.AppendLine(";")

        If description.Length > 0 Then

            sb.AppendLine($"; {description}")

        End If

        sb.AppendLine(If(sourceIsFolder, "; Source folder:", "; Source file:"))
        sb.AppendLine($";   {SourceLink(sourcePath, sourceIsFolder)}")
        sb.AppendLine(";")
        sb.AppendLine("; This file is not intended to be used standalone or with any cleaning software")

        If isArchive Then

            sb.AppendLine($"; Latest copy of {winapp2file.ArchiveFileName}:")
            sb.AppendLine($";   {winapp2file.LatestDownloadUrl}{winapp2file.ArchiveFileName}")

        Else

            sb.AppendLine($"; {LatestReleaseLabel}")
            sb.AppendLine($";   {RepoUrl}/releases/latest")

        End If

        sb.AppendLine(";")
        sb.AppendLine("; To add or fix an entry, see:")
        sb.AppendLine($";   {RepoUrl}/blob/master/CONTRIBUTING.md")
        sb.AppendLine()

        Return sb.ToString()

    End Function

End Module
