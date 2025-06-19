'    Copyright (C) 2018-2025 Hazel Ward
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
''' Represents a winapp2.ini format iniFile, and enables easy access to format specific iniFile information
''' </summary>
''' 
''' Docs last updated: 2024-05-23 | Code last updated: 2024-05-23
Public Class winapp2file

    ''' <summary> 
    ''' The names of the sections of entries as they appear in winapp2.ini <br/>
    ''' "" = main section, bottommost (on disk, last index in memory) section, has label otherwise
    ''' </summary>
    ''' 
    ''' Docs last updated: 2024-05-23 | Code last updated: 2024-05-23
    Public ReadOnly Property FileSectionHeaders As New List(Of String) From {
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

    ''' <summary> 
    ''' A list of iniFiles each containing one of the headers contents 
    ''' <br/> Onto with the indicies from <c> FileSectionHeaders </c> <br  />
    ''' ie. index 0 = Chrome, 1 = Edge Chromium, 2 = Opera, ... 9 = ""
    ''' </summary>
    ''' 
    ''' Docs last updated: 2024-05-23 | Code last updated: 2024-05-23
    Public Property EntrySections As New List(Of iniFile)

    ''' <summary> 
    ''' The list of winapp2entry objects for each header section 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2024-05-23 | Code last updated: 2024-05-23
    Public Property Winapp2entries As New List(Of List(Of winapp2entry))

    ''' <summary> 
    ''' Indicates whether or not this object represents a Non-CCleaner variant of winapp2.ini 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2024-05-23 | Code last updated: 2024-05-23
    Public Property IsNCC As Boolean = False

    '''<summary> 
    '''The directory of the iniFile object used to instantiate this object 
    '''</summary>
    '''
    ''' Docs last updated: 2024-05-23 | Code last updated: 2024-05-23
    Public Property Dir As String = ""

    '''<summary> 
    '''The file name of the iniFile object used to instantiate this object 
    '''</summary>
    '''
    ''' Docs last updated: 2024-05-23 | Code last updated: 2024-05-23
    Public Property Name As String = ""

    '''<summary> 
    '''The version in YYMMDD format of the winapp2.ini file 
    '''<br /> Default value: <c> "; version 000000" </c>
    '''<br /> This will be updated to the current date if the file is saved with the <c> useTodaysDate </c> parameter set to <c> True </c>
    '''<br /> Otherwise, this will be set to the version string found in the file, if one exists 
    '''</summary>
    '''
    ''' Docs last updated: 2024-05-19 | Code last updated: 2024-05-19
    Public Property Version As String = "; Version: 000000"

    ''' <summary> 
    ''' Creates a new <c> winapp2file </c> from an <c> iniFile </c> 
    ''' </summary>
    ''' 
    ''' <param name="file"> 
    ''' A winapp2.ini format <c> iniFile </c> object 
    ''' </param>
    ''' 
    ''' <param name="useTodaysDate"> 
    ''' 
    ''' Indicates that the version should be updated to reflect the current date in YYMMDD format 
    ''' <br /> Optional, Default: <c> False </c>
    ''' 
    ''' </param>
    ''' 
    ''' Docs last updated: 2024-05-19 | Code last updated: 2024-05-19
    Public Sub New(ByVal file As iniFile, Optional useTodaysDate As Boolean = False)

        If file Is Nothing Then argIsNull(NameOf(file)) : Return

        Dir = file.Dir
        Name = file.Name

        EntrySections = FileSectionHeaders.Select(Function(header) New iniFile With {.Name = header}).ToList()
        Winapp2entries = FileSectionHeaders.Select(Function(f) New List(Of winapp2entry)).ToList()

        ' Determine if we're the Non-CCleaner variant of the ini
        IsNCC = Not file.findCommentLine("; This is the non-CCleaner version of Winapp2 that contains extra entries that were removed due to them being added to CCleaner.") = -1

        ' Determine the version string
        Version = DetermineVersionString(file, useTodaysDate)

        ' Build the header sections for browsers/Thunderbird/winapp3
        ' Order: Chrome/Chromium, Edge Chromium, Opera, Brave, Vivaldi, Firefox, Thunderbird, Internet Explorer, Language Files, Dangerous Long, Dangerous
        Dim LeadingCategories As New List(Of String) From {"3029", "3006", "3033", "3034", "3027", "3026", "3030", "3001", "Language Files", "Dangerous Long", "Dangerous"}

        For Each section In file.Sections.Values

            Dim tmpwa2entry As New winapp2entry(section)
            Dim ind = GetCategoryIndex(tmpwa2entry, LeadingCategories)

            EntrySections(ind).Sections.Add(section.Name, section)
            Winapp2entries(ind).Add(tmpwa2entry)

        Next

    End Sub

    ''' <summary>
    ''' Determines and returns the index of the inner ini category to which given <c> <paramref name="entry"/> </c> belongs
    ''' </summary>
    ''' 
    ''' <param name="entry">
    ''' A <c> winapp2entry </c> object to be placed into winapp2.ini 
    ''' </param>
    ''' 
    ''' <param name="categoryValues">
    ''' The list of categories by section and/or langsecref in winapp2.ini order into which each <c> winapp2entry </c> is placed 
    ''' </param>
    ''' 
    ''' <returns> The index of the inner ini category to which given <c> <paramref name="entry"/> </c> belongs </returns>
    ''' 
    ''' Docs last updated: 2024-05-20 | Code last updated: 2024-05-19
    Private Function GetCategoryIndex(entry As winapp2entry, categoryValues As List(Of String)) As Integer

        ' If there's no langSecRef or sectionKey, we're in the main section for sure
        If entry.LangSecRef.KeyCount = 0 AndAlso entry.SectionKey.KeyCount = 0 Then Return -1

        Dim categoryKey As String = If(entry.LangSecRef.KeyCount > 0, entry.LangSecRef.Keys.First.Value, entry.SectionKey.Keys.First.Value)
        Dim ind = categoryValues.IndexOf(categoryKey)

        ' Winapp3 entries are in the second to last section
        If ind = -1 AndAlso categoryKey.ToUpperInvariant.StartsWith("DANGEROUS", StringComparison.InvariantCulture) Then ind = categoryValues.Count - 1

        ' If ind is still -1, we're in the main section
        If ind = -1 Then ind = categoryValues.Count

        Return ind

    End Function

    ''' <summary>
    ''' Returns either the version string found in the file or the current date in YYMMDD format
    ''' </summary>
    ''' 
    ''' <param name="file">
    ''' A winapp2.ini format <c> iniFile </c>
    ''' </param>
    ''' 
    ''' <param name="useTodaysDate"> 
    ''' Indicates that the version should be updated to reflect the current date in YYMMDD format 
    ''' <br /> Optional, Default: <c> False </c>
    ''' </param>
    ''' 
    ''' <returns> The version string found in the file or an updated version string with the current date </returns>
    ''' 
    ''' Docs last updated: 2024-05-19 | Code last updated: 2024-05-19
    Private Function DetermineVersionString(file As iniFile, Optional useTodaysDate As Boolean = False) As String

        If useTodaysDate Then Return $"; Version: {DateTime.Now.ToString("yyMMdd", CultureInfo.InvariantCulture)}"

        Dim hasVersionComment = file.Comments.Count > 0 AndAlso file.Comments.Values(0).Comment.ToUpperInvariant.Contains("VERSION")

        If hasVersionComment Then Return file.Comments.Values(0).Comment

    End Function

    ''' <summary>
    ''' Returns the total number of entries stored in the internal iniFile objects
    ''' </summary>
    ''' 
    ''' Docs last updated: 2024-05-18 | Code last updated: 2024-05-18
    Public Function count() As Integer

        Return EntrySections.Sum(Function(section) section.Sections.Count)

    End Function


    ''' <summary>
    ''' Sorts the internal iniFile objects in winapp2.ini format order
    ''' </summary>
    ''' 
    ''' Docs last updated: 2024-05-23 | Code last updated: 2024-05-23
    Public Sub sortInneriniFiles()

        EntrySections.ForEach(Sub(innerIni) innerIni.sortSections(sortEntryNames(innerIni)))

    End Sub

    ''' <summary> 
    ''' Rebuilds a <c> list of winapp2entry </c> objects back into iniSection objects and returns the collection of them as a single <c> iniFile </c>
    ''' </summary>
    ''' 
    ''' <param name="entryList">
    ''' The list of <c> winapp2entry </c> objects to be rebuilt into a single <c> iniFile </c>
    ''' </param>
    ''' 
    ''' Docs last updated: 2024-05-23 | Code last updated: 2024-05-23
    Private Function rebuildInnerIni(ByRef entryList As List(Of winapp2entry)) As iniFile

        Dim tmpini As New iniFile

        entryList.ForEach(Sub(entry) tmpini.Sections.Add(entry.Name, New iniSection(entry.dumpToListOfStrings) With {.StartingLineNumber = entry.LineNum}))

        Return tmpini

    End Function

    ''' <summary> 
    ''' Updates the internal <c> iniFile </c> objects to reflect any changes which have been made to the <c> winapp2entry </c> objects 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2024-05-23 | Code last updated: 2024-05-23
    Public Sub rebuildWinapp2ChangesToIniFiles()

        For i = 0 To EntrySections.Count - 1

            EntrySections(i) = rebuildInnerIni(Winapp2entries(i))
            EntrySections(i).Name = FileSectionHeaders(i)

        Next

    End Sub

    '''<summary>
    ''' Returns the <c> winapp2file's </c> inner <c> iniFiles </c> as a single <c> inifile </c> object 
    ''' with its sections ordered as they would be in winapp2.ini on disk 
    '''</summary>
    '''
    ''' <returns> A single <c> iniFile </c> containing all the ordered sections from all of the inner <c> iniFiles </c> </returns>
    ''' 
    ''' Docs last updated: 2024-05-23 | Code last updated: 2024-05-23
    Public Function toIni() As iniFile

        Dim out As New iniFile

        For Each section In EntrySections.SelectMany(Function(entry) entry.Sections.Values)

            out.Sections.Add(section.Name, section)

        Next

        Return out

    End Function

    ''' <summary> 
    ''' Builds and returns the winapp2.ini text including the preamble comments 
    ''' </summary>
    ''' 
    ''' <returns> The complete text of winapp2.ini as a single String </returns>
    ''' 
    ''' Docs last updated: 2024-05-18 | Code last updated: 2024-05-18
    Public Function winapp2string() As String

        Dim builder As New StringBuilder()
        Dim fileName = If(IsNCC, "Winapp2 (Non-CCleaner version)", "Winapp2")
        Dim licLink = If(IsNCC, "https://github.com/MoscaDotTo/Winapp2/blob/master/Non-CCleaner/License.md", "https://github.com/MoscaDotTo/Winapp2/blob/master/License.md")

        builder.AppendLine(Version)
        builder.AppendLine($"; # of entries: {count():#,###}")
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

        For i = 0 To EntrySections.Count - 2

            If EntrySections(i).Sections.Count <= 0 Then Continue For

            builder.AppendLine(";")
            builder.AppendLine($"; {EntrySections(i).Name} ({EntrySections(i).Sections.Count})")
            builder.AppendLine()
            builder.AppendLine(EntrySections(i).toString)
            builder.AppendLine($"; End of {EntrySections(i).Name}")

        Next

        builder.AppendLine()
        If EntrySections.Last.Sections.Count > 0 Then builder.Append($"{EntrySections.Last.toString}")

        Return builder.ToString()

    End Function

End Class