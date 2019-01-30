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
Module winapp2handler
    ''' <summary>
    ''' Sorts a list of strings after performing some mutations on the data (if necessary). Returns the sorted list of strings.
    ''' </summary>
    ''' <param name="ListToBeSorted">A list of strings for sorting </param>
    ''' <param name="characterToReplace">A character (string) to replace</param>
    ''' <param name="replacementText">The chosen replacement text</param>
    ''' <returns> The sorted state of the listToBeSorted</returns>
    Public Function replaceAndSort(ListToBeSorted As List(Of String), characterToReplace As String, replacementText As String) As List(Of String)
        Dim changes As New Dictionary(Of String, String)
        ' Replace our target characters if they exist
        For i As Integer = 0 To ListToBeSorted.Count - 1
            Dim item As String = ListToBeSorted(i)
            If item.Contains(characterToReplace) Then
                Dim renamedItem As String = item.Replace(characterToReplace, replacementText)
                trackChanges(changes, item, renamedItem)
                ListToBeSorted(i) = renamedItem
            End If
        Next
        ' Pad numbers if necessary 
        findAndReplaceNumbers(ListToBeSorted, changes)
        ' Copy the modified list to be sorted and sort it
        Dim sortedEntryList As New List(Of String)
        sortedEntryList.AddRange(ListToBeSorted)
        sortedEntryList.Sort()
        ' Restore the original state of our data
        undoChanges(changes, {ListToBeSorted, sortedEntryList})
        Return sortedEntryList
    End Function

    ''' <summary>
    ''' Tracks renames made while mutating data for string sorting.
    ''' </summary>
    ''' <param name="changeDict">The dictionary containing changes</param>
    ''' <param name="currentValue">The current value of a string</param>
    ''' <param name="newValue">The new value of a piece of a string</param>
    Public Sub trackChanges(ByRef changeDict As Dictionary(Of String, String), currentValue As String, newValue As String)
        If changeDict.Keys.Contains(currentValue) Then
            changeDict.Add(newValue, changeDict.Item(currentValue))
            changeDict.Remove(currentValue)
        Else
            changeDict.Add(newValue, currentValue)
        End If
    End Sub

    ''' <summary>
    ''' Restores the original state of data mutated for the purposes of string sorting.
    ''' </summary>
    ''' <param name="changeDict">The dictionary containing changes made to data in the lists</param>
    ''' <param name="lstArray">An array containing lists of strings whose data has been modified</param>
    Public Sub undoChanges(ByRef changeDict As Dictionary(Of String, String), ByRef lstArray As List(Of String)())
        For Each lst In lstArray
            For Each key In changeDict.Keys
                replaceStrAtIndexOf(lst, key, changeDict.Item(key))
            Next
        Next
    End Sub

    ''' <summary>
    ''' Searches the input list for numbers and returns the length of the longest number.
    ''' </summary>
    ''' <param name="lst">A list of strings to be searched</param>
    ''' <returns>The length of the longest number in lst</returns>
    Private Function findLongestNumLength(ByRef lst As List(Of String)) As Integer
        Dim out As Integer = 0
        For Each item In lst
            For Each mtch As Match In Regex.Matches(item, "[\d]+")
                If mtch.Length > out Then out = mtch.Length
            Next
        Next
        Return out
    End Function

    ''' <summary>
    ''' Detects the longest length number in a given list of strings and prepends all shorter numbers with zeroes such that all numbers are the same length
    ''' </summary>
    ''' This is to maintain numerical precedence in string sorting, ie. larger numbers come alphabetically "after" smaller numbers. 
    ''' <param name="listToBeSorted">The list to be modified prior to sorting</param>
    ''' <param name="changes">The dictionary of changes made to the strings in listToBeSorted</param>
    Private Sub findAndReplaceNumbers(ByRef listToBeSorted As List(Of String), ByRef changes As Dictionary(Of String, String))
        Dim longestNumLen As Integer = findLongestNumLength(listToBeSorted)
        If longestNumLen < 2 Then Exit Sub
        For i As Integer = 0 To listToBeSorted.Count - 1
            Dim baseString As String = listToBeSorted(i)
            Dim paddedString As String = baseString
            Dim numberAndDecimals As New Regex("[\d]+(\.?[\d]+|\b)*")
            For Each m As Match In numberAndDecimals.Matches(baseString)
                ' Special procedure for numbers with any amount of decimal points in them
                Dim currentMatch As String = m.ToString
                If currentMatch.Contains(".") Then
                    Dim out As String = ""
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
            If baseString = paddedString Then Continue For
            ' Rename and track changes appropriately
            trackChanges(changes, baseString, paddedString)
            replaceStrAtIndexOf(listToBeSorted, baseString, paddedString)
        Next
    End Sub

    ''' <summary>
    ''' Replaces an item in a list of strings at the index of another given item
    ''' </summary>
    ''' <param name="list">The list containing all the strings</param>
    ''' <param name="indexOfText">The text to be replaced</param>
    ''' <param name="newText">The replacement text</param>
    Public Sub replaceStrAtIndexOf(ByRef list As List(Of String), indexOfText As String, newText As String)
        list(list.IndexOf(indexOfText)) = newText
    End Sub

    ''' <summary>
    ''' Pads a number to a given length by preceeding it with zeroes (0's) and returns the padded number
    ''' </summary>
    ''' <param name="longestNumLen">The desired maxmimum length of a number</param>
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
    Public Function sortEntryNames(ByVal file As iniFile) As List(Of String)
        Return replaceAndSort(file.getSectionNamesAsList, "-", "  ")
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
            ' Build the header sections for browsers/thunderbird/winapp3
            Dim langSecRefs As New List(Of String) From {"3029", "3026", "3030", "Language Files", "Dangerous Long", "Dangerous"}
            For Each section In file.sections.Values
                Dim tmpwa2entry As New winapp2entry(section)
                Dim ind = -1
                If tmpwa2entry.langSecRef.keyCount > 0 Then
                    ind = langSecRefs.IndexOf(tmpwa2entry.langSecRef.keys.First.value)
                ElseIf tmpwa2entry.sectionKey.keyCount > 0 Then
                    ind = langSecRefs.IndexOf(tmpwa2entry.sectionKey.keys.First.value)
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
        ''' Rebuilds a list of winapp2entry objects back into iniSection objcts and returns the collection of them as an iniFile
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
            out += appendNewLine("; Any contributions are appreciated. Please refer to our readme to learn to make your own entries here: https://github.com/MoscaDotTo/Winapp2/blob/master/README.md")
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

    ''' <summary>
    ''' Represents a winapp2.ini format iniKey and provides direct access to the keys by their type
    ''' </summary>
    Public Class winapp2entry
        Public name As String
        Public fullName As String
        Public detectOS As New keyList("DetectOS")
        Public langSecRef As New keyList("LangSecRef")
        Public sectionKey As New keyList("Section")
        Public specialDetect As New keyList("SpecialDetect")
        Public detects As New keyList("Detect")
        Public detectFiles As New keyList("DetectFile")
        Public warningKey As New keyList("Warning")
        Public defaultKey As New keyList("Default")
        Public fileKeys As New keyList("FileKey")
        Public regKeys As New keyList("RegKey")
        Public excludeKeys As New keyList("ExcludeKey")
        Public errorKeys As New keyList("Error")
        Public keyListList As New List(Of keyList) From {detectOS, langSecRef, sectionKey, specialDetect, detects, detectFiles,
                                                            warningKey, defaultKey, fileKeys, regKeys, excludeKeys, errorKeys}
        Public lineNum As New Integer

        ''' <summary>
        ''' Construct a new winapp2entry object from an iniSection
        ''' </summary>
        ''' <param name="section">A winapp2.ini format iniSection object</param>
        Public Sub New(ByVal section As iniSection)
            fullName = section.getFullName
            name = section.name
            updKeyListList()
            lineNum = section.startingLineNumber
            section.constKeyLists(keyListList)
        End Sub

        ''' <summary>
        ''' Clears and updates the keyListList with the current state of the keys
        ''' </summary>
        Private Sub updKeyListList()
            keyListList = New List(Of keyList) From {detectOS, langSecRef, sectionKey, specialDetect, detects, detectFiles,
                                                      warningKey, defaultKey, fileKeys, regKeys, excludeKeys, errorKeys}
        End Sub

        ''' <summary>
        ''' Returns the keys in each keyList back as a list of Strings in winapp2.ini (style) order
        ''' </summary>
        ''' <returns></returns>
        Public Function dumpToListOfStrings() As List(Of String)
            Dim outList As New List(Of String) From {fullName}
            updKeyListList()
            keyListList.ForEach(Sub(lst) lst.keys.ForEach(Sub(key) outList.Add(key.toString)))
            Return outList
        End Function
    End Class

    ''' <summary>
    ''' Provides a few helpful methods for dissecting winapp2key objects
    ''' </summary>
    Public Class winapp2KeyParameters
        Public pathString As String = ""
        Public argsList As New List(Of String)
        Public flagString As String = ""
        Public keyType As String = ""
        Public keyNum As String = ""

        Public Sub New(key As iniKey)
            keyType = key.keyType
            Dim splitKey As String() = key.value.Split(CChar("|"))
            Select Case key.keyType
                Case "FileKey"
                    keyNum = key.keyType.Replace("FileKey", "")
                    If splitKey.Count > 1 Then
                        pathString = splitKey(0)
                        argsList.AddRange(splitKey(1).Split(CChar(";")))
                        flagString = If(splitKey.Count >= 3, splitKey.Last, "None")
                    Else
                        pathString = key.value
                    End If
                Case "ExcludeKey"
                    Select Case splitKey.Count
                        Case 2
                            pathString = splitKey(1)
                            flagString = splitKey(0)
                        Case 3
                            pathString = splitKey(1)
                            argsList.AddRange(splitKey(2).Split(CChar(";")))
                            flagString = splitKey(0)
                    End Select
                Case "RegKey"
                    pathString = splitKey(0)
                    If splitKey.Count > 1 Then argsList.Add(splitKey(1))
            End Select
        End Sub

        ''' <summary>
        ''' Reconstructs a FileKey to hold the format of FileKeyX=PATH|FILE;FILE;FILE....|FLAG
        ''' </summary>
        ''' <param name="key">An iniKey to be reconstructed</param>
        ''' Also trims empty comments 
        Public Sub reconstructKey(ByRef key As iniKey)
            Dim out As String = ""
            out += $"{pathString}{If(argsList.Count > 0, "|", "")}"
            If argsList.Count > 1 Then
                For i As Integer = 0 To argsList.Count - 2
                    If Not argsList(i) = "" Then out += argsList(i) & ";"
                Next
            End If
            If argsList.Count > 0 Then out += argsList.Last
            If Not flagString = "None" Then out += "|" & flagString
            key.value = out
        End Sub
    End Class
End Module