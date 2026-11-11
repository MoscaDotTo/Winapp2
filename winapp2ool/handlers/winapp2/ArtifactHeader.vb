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
''' Writes the do-not-edit header that every generator puts at the top of its build artifacts
''' (EntryBuilder's, BrowserBuilder's and UWPBuilder's output under <c> Assembler\Generated </c>),
''' and holds the repository URL and the repository-relative source folders those headers link.
''' This is the only place either is spelled. <br /><br />
'''
''' A header names the file and the generator, links the sources its entries come from on
''' <c> master </c>, and points to the finished winapp2.ini and the contributor guide. It carries
''' no license: the artifacts are repository files, and the published winapp2.ini carries the
''' license.
''' </summary>
Public Module ArtifactHeader

    ''' <summary>
    ''' The winapp2 repository on GitHub
    ''' </summary>
    Public Const RepoUrl As String = "https://github.com/MoscaDotTo/Winapp2"

    ''' <summary>
    ''' EntryBuilder's source folder, relative to the repository root
    ''' </summary>
    Public Const EntryBuilderSources As String = "Assembler/Sources/EntryBuilder"

    ''' <summary>
    ''' BrowserBuilder's source folder, relative to the repository root
    ''' </summary>
    Public Const BrowserBuilderSources As String = "Assembler/Sources/BrowserBuilder"

    ''' <summary>
    ''' UWPBuilder's source folder, relative to the repository root. It holds <c> UWP.ini </c>
    ''' and the <c> AppInfo </c> folder.
    ''' </summary>
    Public Const UWPBuilderSources As String = "Assembler/Sources/UWPBuilder"

    ''' <summary>
    ''' UWPBuilder's per-letter AppInfo folder, relative to the repository root
    ''' </summary>
    Public Const UWPBuilderAppInfoSources As String = UWPBuilderSources & "/AppInfo"

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
    ''' ends in whitespace.
    ''' </summary>
    '''
    ''' <param name="fileName">
    ''' The artifact's file name, such as <c> A.ini </c>
    ''' </param>
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
    ''' A line saying what the artifact holds, written below the source link. Only per-letter
    ''' artifacts have one. <br /><br />
    ''' Optional, Default: <c> "" </c> (no line)
    ''' </param>
    Public Function Text(fileName As String,
                         generatorName As String,
                         entryCount As Integer,
                         sourcePath As String,
                         sourceIsFolder As Boolean,
                         Optional description As String = "") As String

        Dim comeFrom = If(sourceIsFolder, "its entries come from the files in:", "its entries come from:")

        Dim sb As New StringBuilder()
        sb.AppendLine(EntryCountLine(entryCount))
        sb.AppendLine($"; {fileName} is a build artifact generated by winapp2ool's {generatorName}. Don't edit it directly:")
        sb.AppendLine($"; the next build overwrites it. In the winapp2 repository, {comeFrom}")
        sb.AppendLine($";   {SourceLink(sourcePath, sourceIsFolder)}")

        If description.Length > 0 Then

            sb.AppendLine($"; {description}")

        End If

        sb.AppendLine(";")
        sb.AppendLine("; This is one piece of winapp2.ini, not a usable winapp2.ini. The finished file, in every flavor:")
        sb.AppendLine($";   {RepoUrl}/releases/latest")
        sb.AppendLine("; To add or fix an entry, see:")
        sb.AppendLine($";   {RepoUrl}/blob/master/CONTRIBUTING.md")
        sb.AppendLine()

        Return sb.ToString()

    End Function

End Module
