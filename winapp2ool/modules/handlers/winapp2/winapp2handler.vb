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
Imports System.Text.RegularExpressions

''' <summary>
''' Provides an object model and some helpful functions for working with winapp2.ini format .ini files
''' </summary>
Public Module winapp2handler
    ''' <summary>
    ''' Sorts a list of strings after performing some mutations on the data (if necessary). Returns the sorted list of strings.
    ''' </summary>
    ''' <param name="ListToBeSorted">A list of strings for sorting </param>
    ''' <param name="characterToReplace">A character (string) to replace</param>
    ''' <param name="replacementText">The chosen replacement text</param>
    ''' <returns> The sorted state of the listToBeSorted</returns>
    Public Function replaceAndSort(ListToBeSorted As strList, characterToReplace As String, replacementText As String) As strList
        Dim changes As New changeDict
        ' Replace our target characters if they exist
        For i As Integer = 0 To ListToBeSorted.items.Count - 1
            Dim item As String = ListToBeSorted.items(i)
            If item.Contains(characterToReplace) Then
                Dim renamedItem As String = item.Replace(characterToReplace, replacementText)
                changes.trackChanges(item, renamedItem)
                ListToBeSorted.items(i) = renamedItem
            End If
        Next
        ' Pad numbers if necessary 
        findAndReplaceNumbers(ListToBeSorted, changes)
        ' Copy the modified list to be sorted and sort it
        Dim sortedEntryList As New strList
        sortedEntryList.items.AddRange(ListToBeSorted.items)
        sortedEntryList.items.Sort()
        ' Restore the original state of our data
        changes.undoChanges({ListToBeSorted, sortedEntryList})
        Return sortedEntryList
    End Function

    ''' <summary>
    ''' Searches the input list for numbers and returns the length of the longest number.
    ''' </summary>
    ''' <param name="lst">A list of strings to be searched</param>
    ''' <returns>The length of the longest number in lst</returns>
    Private Function findLongestNumLength(ByRef lst As strList) As Integer
        Dim out As Integer = 0
        For Each item In lst.items
            For Each mtch As Match In Regex.Matches(item, "[\d]+")
                If mtch.Length > out Then out = mtch.Length
            Next
        Next
        Return out
    End Function

    ''' <summary>
    ''' Detects the longest length number in a given list of strings and prepends all shorter numbers with zeros such that all numbers are the same length
    ''' </summary>
    ''' This is to maintain numerical precedence in string sorting, ie. larger numbers come alphabetically "after" smaller numbers. 
    ''' <param name="listToBeSorted">The list to be modified prior to sorting</param>
    ''' <param name="changes">The dictionary of changes made to the strings in listToBeSorted</param>
    Private Sub findAndReplaceNumbers(ByRef listToBeSorted As strList, ByRef changes As changeDict)
        Dim longestNumLen As Integer = findLongestNumLength(listToBeSorted)
        If longestNumLen < 2 Then Exit Sub
        For i As Integer = 0 To listToBeSorted.count - 1
            Dim baseString As String = listToBeSorted.items(i)
            Dim paddedString As String = baseString
            Dim numberAndDecimals As New Regex("[\d]+(\.?[\d]+|\b)*")
            For Each m As Match In numberAndDecimals.Matches(baseString)
                ' Special procedure for numbers with any amount of decimal points in them
                Dim currentMatch As String = m.ToString
                If currentMatch.Contains(".") Then
                    Dim out = ""
                    Dim tStr As String() = currentMatch.Split(CChar("."))
                    For p As Integer = 0 To tStr.Length - 1
                        out += padNumberStr(longestNumLen, tStr(p))
                        If p < tStr.Length - 1 Then out += "."
                    Next
                    paddedString = paddedString.Replace(currentMatch, out)
                Else
                    ' Grab characters from both sides so that we don't have to worry about duplicate m matches 
                    Dim numsPlusReplBits As New Regex($"([^\d]|\b){currentMatch}([^\d]|\b)")
                    For Each mm As Match In numsPlusReplBits.Matches(paddedString)
                        Dim replacementText As String = mm.ToString.Replace(currentMatch, padNumberStr(longestNumLen, currentMatch))
                        paddedString = paddedString.Replace(mm.ToString, replacementText)
                    Next
                End If
            Next
            ' Don't rename if we didn't change anything
            If baseString.Equals(paddedString) Then Continue For
            ' Rename and track changes appropriately
            changes.trackChanges(baseString, paddedString)
            listToBeSorted.replaceStrAtIndexOf(baseString, paddedString)
        Next
    End Sub



    ''' <summary>
    ''' Returns the value from an ExcludeKey with the Flag parameter removed as a String
    ''' </summary>
    ''' <param name="key">An ExcludeKey iniKey</param>
    ''' <returns></returns>
    Public Function pathFromExcludeKey(key As iniKey) As String
        Dim pathFromKey As String = key.Value.TrimStart(CType("FILE|", Char()))
        pathFromKey = pathFromKey.TrimStart(CType("PATH|", Char()))
        pathFromKey = pathFromKey.TrimStart(CType("REG|", Char()))
        Return pathFromKey
    End Function

    ''' <summary>
    ''' Pads a number to a given length by preceding it with zeros (0's) and returns the padded number
    ''' </summary>
    ''' <param name="longestNumLen">The desired maximum length of a number</param>
    ''' <param name="num">a given number</param>
    ''' <returns></returns>
    Private Function padNumberStr(longestNumLen As Integer, num As String) As String
        Dim replMatch As String = ""
        While replMatch.Length < longestNumLen - num.Length
            replMatch += "0"
        End While
        Return replMatch & num
    End Function

    ''' <summary>
    ''' Removes winapp2entry objects from a given winapp2file sectionList
    ''' </summary>
    ''' <param name="sectionList">The list of winapp2sections </param>
    ''' <param name="removalList"></param>
    Public Sub removeEntries(ByRef sectionList As List(Of winapp2entry), ByRef removalList As List(Of winapp2entry))
        For Each item In removalList
            sectionList.Remove(item)
        Next
        removalList.Clear()
    End Sub

    ''' <summary>
    ''' Sort the sections in an inifile by name
    ''' </summary>
    ''' <param name="file">The iniFile to be sorted</param>
    ''' <returns></returns>
    Public Function sortEntryNames(ByVal file As iniFile) As strList
        Return replaceAndSort(file.namesToStrList, "-", "  ")
    End Function

    ''' <summary>
    ''' Represents a winapp2.ini format iniFile, and enables easy access to format specific iniFile information
    ''' </summary>
    Public Class winapp2file
        Public entryList As List(Of String)
        ' "" = main section, bottom most in all circumstances and appearing without a label 
        ReadOnly sectionHeaderFooter As String() = {"Chrome/Chromium based browsers", "Firefox/Mozilla based browsers", "Thunderbird",
            "Language entries", "Potentially very long scan time (and also dangerous) entries", "Dangerous entries", ""}
        ' As above, index 0 = Chrome, 1 = Firefox, 2 = Thunderbird.... 6 = ""
        Public entrySections(6) As iniFile
        Public entryLines(6) As List(Of Integer)
        Public winapp2entries(6) As List(Of winapp2entry)
        Public isNCC As Boolean
        Public dir As String
        Public name As String
        Dim version As String

        ''' <summary>
        ''' Create a new meta winapp2 object from an iniFile object
        ''' </summary>
        ''' <param name="file">A winapp2.ini format iniFile object</param>
        Public Sub New(ByVal file As iniFile)
            entryList = New List(Of String)
            For i As Integer = 0 To 6
                entrySections(i) = New iniFile With {.name = sectionHeaderFooter(i)}
                entryLines(i) = New List(Of Integer)
                winapp2entries(i) = New List(Of winapp2entry)
            Next
            ' Determine if we're the Non-CCleaner variant of the ini
            isNCC = Not file.findCommentLine("; This is the non-CCleaner version of Winapp2 that contains extra entries that were removed due to them being added to CCleaner.") = -1
            ' Determine the version string
            If file.comments.Count = 0 Then version = "; version 000000"
            If file.comments.Count > 0 Then version = If(Not file.comments.Values(0).comment.ToLower.Contains("version"), "; version 000000", file.comments.Values(0).comment)
            ' Build the header sections for browsers/Thunderbird/winapp3
            Dim langSecRefs As New List(Of String) From {"3029", "3026", "3030", "Language Files", "Dangerous Long", "Dangerous"}
            For Each section In file.sections.Values
                Dim tmpwa2entry As New winapp2entry(section)
                Dim ind = -1
                If tmpwa2entry.langSecRef.keyCount > 0 Then
                    ind = langSecRefs.IndexOf(tmpwa2entry.langSecRef.keys.First.Value)
                ElseIf tmpwa2entry.sectionKey.keyCount > 0 Then
                    ind = langSecRefs.IndexOf(tmpwa2entry.sectionKey.keys.First.Value)
                End If
                If ind = -1 Then ind = 6
                addToInnerFile(ind, tmpwa2entry, section)
            Next
        End Sub

        ''' <summary>
        ''' Inserts an iniSection into its respective tracking file and records the winapp2entry object form accordingly. 
        ''' </summary>
        ''' <param name="ind">The index of the tracking file</param>
        ''' <param name="entry">The section in winapp2entry format</param>
        ''' <param name="section">A section to be tracked</param>
        Private Sub addToInnerFile(ind As Integer, entry As winapp2entry, section As iniSection)
            If Not entrySections(ind).sections.Keys.Contains(section.name) Then
                entrySections(ind).sections.Add(section.name, section)
                entryLines(ind).Add(section.startingLineNumber)
                winapp2entries(ind).Add(entry)
            End If
        End Sub

        ''' <summary>
        ''' Returns the total number of entries stored in the internal iniFile objects
        ''' </summary>
        ''' <returns></returns>
        Public Function count() As Integer
            Dim out As Integer = 0
            For Each section In entrySections
                out += section.sections.Count
            Next
            Return out
        End Function

        ''' <summary>
        ''' Sorts the internal iniFile objects in winapp2.ini format order
        ''' </summary>
        Public Sub sortInneriniFiles()
            For Each innerIni In entrySections
                innerIni.sortSections(sortEntryNames(innerIni))
            Next
        End Sub

        ''' <summary>
        ''' Rebuilds a list of winapp2entry objects back into iniSection objects and returns the collection of them as an iniFile
        ''' </summary>
        ''' <param name="entryList"></param>
        ''' <returns></returns>
        Private Function rebuildInnerIni(ByRef entryList As List(Of winapp2entry)) As iniFile
            Dim tmpini As New iniFile
            For Each entry In entryList
                tmpini.sections.Add(entry.name, New iniSection(entry.dumpToListOfStrings))
                tmpini.sections.Values.Last.startingLineNumber = entry.lineNum
            Next
            Return tmpini
        End Function

        ''' <summary>
        ''' Updates the internal iniFile objects
        ''' </summary>
        Public Sub rebuildToIniFiles()
            For i As Integer = 0 To entrySections.Count - 1
                entrySections(i) = rebuildInnerIni(winapp2entries(i))
                entrySections(i).name = sectionHeaderFooter(i)
            Next
        End Sub

        ''' <summary>
        ''' Builds and returns the winapp2.ini text including header comments for writing back to a file
        ''' </summary>
        ''' <returns></returns>
        Public Function winapp2string() As String
            Dim fileName As String = If(isNCC, "Winapp2 (Non-CCleaner version)", "Winapp2")
            Dim licLink As String = appendNewLine(If(isNCC, "https://github.com/MoscaDotTo/Winapp2/blob/master/Non-CCleaner/License.md", "https://github.com/MoscaDotTo/Winapp2/blob/master/License.md"))
            ' Version string (YYMMDD format) & entry count 
            Dim out As String = appendNewLine(version)
            out += appendNewLine($"; # of entries: {count.ToString("#,###")}")
            out += appendNewLine(";")
            out += $"; {fileName} is fully licensed under the CC-BY-SA-4.0 license agreement. Please refer to our license agreement before using Winapp2: {licLink}"
            out += appendNewLine($"; If you plan on modifying, distributing, and/or hosting {fileName} for your own program or website, please ask first.")
            out += appendNewLine(";")
            If isNCC Then
                out += appendNewLine("; This is the non-CCleaner version of Winapp2 that contains extra entries that were removed due to them being added to CCleaner.")
                out += appendNewLine("; Do not use this file for CCleaner as the extra cleaners may cause conflicts with CCleaner.")
            End If
            out += appendNewLine("; You can get the latest Winapp2 here: https://github.com/MoscaDotTo/Winapp2")
            out += appendNewLine("; Any contributions are appreciated. Please refer to our ReadMe to learn to make your own entries here: https://github.com/MoscaDotTo/Winapp2/blob/master/README.md")
            out += appendNewLine("; Try out Winapp2ool for many useful additional features including updating and trimming winapp2.ini: https://github.com/MoscaDotTo/Winapp2/raw/master/winapp2ool/bin/Release/winapp2ool.exe")
            out += appendNewLine("; You can find the Winapp2ool ReadMe here: https://github.com/MoscaDotTo/Winapp2/blob/master/winapp2ool/Readme.md")
            ' Adds each section's toString if it exists with a proper header and footer, followed by the main section (if it exists)
            For i As Integer = 0 To 5
                If entrySections(i).sections.Count > 0 Then
                    out += appendNewLine("; ")
                    out += appendNewLine(appendNewLine("; " & entrySections(i).name))
                    out += entrySections(i).toString
                    out += appendNewLine($"{prependNewLines(False)}; End of {entrySections(i).name}")
                End If
            Next
            If entrySections.Last.sections.Count > 0 Then out += prependNewLines(False) & entrySections.Last.toString
            Return out
        End Function
    End Class
End Module