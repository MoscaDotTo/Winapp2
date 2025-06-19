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
Imports System.Text.RegularExpressions

''' <summary>
''' Provides an object model and some helpful functions for working with winapp2.ini format .ini files
''' </summary>
Public Module winapp2handler
    ''' <summary> Sorts a list of <c> Strings </c> after performing some mutations on the data (if necessary). Returns the sorted list of strings. </summary>
    ''' <param name="ListToBeSorted"> A <c> list (of String)s </c> to be sorted </param>
    ''' <param name="textToBeReplaced"> The <c> String </c> data that will be replaced during mutations </param>
    ''' <param name="replacementText">The data with which <c> <paramref name="textToBeReplaced"/> </c> will be replaced </param>
    Public Function replaceAndSort(ListToBeSorted As strList, textToBeReplaced As String, replacementText As String) As strList
        If ListToBeSorted Is Nothing Then argIsNull(NameOf(ListToBeSorted)) : Return Nothing
        Dim changes As New changeDict
        ' Replace our target characters if they exist
        For i = 0 To ListToBeSorted.Items.Count - 1
            Dim item = ListToBeSorted.Items(i)
            If item.Contains(textToBeReplaced) Then
                Dim renamedItem = item.Replace(textToBeReplaced, replacementText)
                changes.trackChanges(item, renamedItem)
                ListToBeSorted.Items(i) = renamedItem
            End If
        Next
        ' Pad numbers if necessary 
        findAndReplaceNumbers(ListToBeSorted, changes)
        ' Copy the modified list to be sorted and sort it
        Dim sortedEntryList As New strList
        sortedEntryList.Items.AddRange(ListToBeSorted.Items)
        sortedEntryList.Items.Sort()
        ' Restore the original state of our data
        changes.undoChanges({ListToBeSorted, sortedEntryList})
        Return sortedEntryList
    End Function

    ''' <summary> Searches the <c> <paramref name="lst"/> </c> for integers and returns the length of the longest integer found </summary>
    ''' <param name="lst"> A list of strings to be searched </param>
    Private Function findLongestNumLength(ByRef lst As strList) As Integer
        Dim out = 0
        For Each item In lst.Items
            For Each mtch As Match In Regex.Matches(item, "[\d]+")
                If mtch.Length > out Then out = mtch.Length
            Next
        Next
        Return out
    End Function

    ''' <summary> Detects the length (number of digits) in the "longest" integer in a given <c> list (of String)s </c> and prepends all shorter integers with zeros such that all the integers in all Strings are the same length </summary>
    ''' This is to maintain numerical precedence in string sorting, ie. larger numbers come alphabetically "after" smaller numbers. 
    ''' <param name="listToBeSorted"> The list to be modified and sorted </param>
    ''' <param name="changes"> Dictonary of the changes made to the Strings in <c> <paramref name="listToBeSorted"/> </c></param>
    Private Sub findAndReplaceNumbers(ByRef listToBeSorted As strList, ByRef changes As changeDict)
        Dim longestNumLen = findLongestNumLength(listToBeSorted)
        If longestNumLen < 2 Then Exit Sub
        For i = 0 To listToBeSorted.Count - 1
            Dim baseString = listToBeSorted.Items(i)
            Dim paddedString = baseString
            Dim numberAndDecimals As New Regex("[\d]+(\.?[\d]+|\b)*")
            For Each m As Match In numberAndDecimals.Matches(baseString)
                ' Special procedure for numbers with any amount of decimal points in them
                Dim currentMatch = m.ToString
                If currentMatch.Contains(".") Then
                    Dim out = ""
                    Dim tStr = currentMatch.Split(CChar("."))
                    For p = 0 To tStr.Length - 1
                        out += padNumberStr(longestNumLen, tStr(p))
                        If p < tStr.Length - 1 Then out += "."
                    Next
                    paddedString = paddedString.Replace(currentMatch, out)
                Else
                    ' Grab characters from both sides so that we don't have to worry about duplicate m matches 
                    Dim numsPlusReplBits As New Regex($"([^\d]|\b){currentMatch}([^\d]|\b)")
                    For Each mm As Match In numsPlusReplBits.Matches(paddedString)
                        Dim replacementText = mm.ToString.Replace(currentMatch, padNumberStr(longestNumLen, currentMatch))
                        paddedString = paddedString.Replace(mm.ToString, replacementText)
                    Next
                End If
            Next
            ' Don't rename if we didn't change anything
            If baseString.Equals(paddedString, StringComparison.InvariantCulture) Then Continue For
            ' Rename and track changes appropriately
            changes.trackChanges(baseString, paddedString)
            listToBeSorted.replaceStrAtIndexOf(baseString, paddedString)
        Next
    End Sub

    ''' <summary> Returns the path from an ExcludeKey with the <c> Flag </c> parameter removed, as a <c> String </c></summary>
    ''' <param name="key"> An ExcludeKey <c> iniKey </c></param>
    Public Function pathFromExcludeKey(key As iniKey) As String
        If key Is Nothing Then argIsNull(NameOf(key)) : Return Nothing
        Dim pathFromKey = key.Value.TrimStart(CType("FILE|", Char()))
        pathFromKey = pathFromKey.TrimStart(CType("PATH|", Char()))
        pathFromKey = pathFromKey.TrimStart(CType("REG|", Char()))
        Return pathFromKey
    End Function

    ''' <summary> Pads a given number to a given length by prepending it with zeros (0's), returns the padded number </summary>
    ''' <param name="longestNumLen"> The desired maximum length of a number </param>
    ''' <param name="num"> The given number </param>
    Private Function padNumberStr(longestNumLen As Integer, num As String) As String
        Dim replMatch = ""
        While replMatch.Length < longestNumLen - num.Length
            replMatch += "0"
        End While
        Return replMatch & num
    End Function

    ''' <summary> Removes <c> winapp2entry </c> objects from a given <c> winapp2file's </c> <c> sectionList </c></summary>
    ''' <param name="sectionList">The list of <c> winapp2entrys </c> representing a given "section" in the file </param>
    ''' <param name="removalList">The list of <c> winapp2entrys </c> to be removed from a section </param>
    Public Sub removeEntries(ByRef sectionList As List(Of winapp2entry), ByRef removalList As List(Of winapp2entry))
        If removalList Is Nothing Then argIsNull(NameOf(removalList)) : Return
        If sectionList Is Nothing Then argIsNull(NameOf(sectionList)) : Return
        For Each item In removalList
            sectionList.Remove(item)
        Next
        removalList.Clear()
    End Sub

    ''' <summary> Returns the <c> Names </c> of the <c> iniSections </c> in an <c> iniFile </c> sorted in winapp2.ini order as a <c> strList </c></summary>
    ''' <param name="file"> The <c> iniFile </c> whose sections will be sorted </param>
    Public Function sortEntryNames(ByVal file As iniFile) As strList
        If file Is Nothing Then argIsNull(NameOf(file)) : Return Nothing
        Return replaceAndSort(file.namesToStrList, "-", "  ")
    End Function
End Module