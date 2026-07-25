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
Imports System.Text.RegularExpressions

''' <summary>
''' Provides an object model and some helpful functions for working with winapp2.ini format .ini files
''' </summary>
Public Module winapp2handler

    ''' <summary>
    ''' Matches a run of digits, optionally followed by ".digits" segments
    ''' (e.g. "12", "1.2.3"). Compiled once and shared across all callers.
    ''' </summary>
    Private ReadOnly numberAndDecimals As New Regex("[\d]+(\.?[\d]+|\b)*",
                                                    RegexOptions.Compiled Or RegexOptions.CultureInvariant)

    ''' <summary>
    ''' Matches a run of digits. Used by <c> findLongestNumLength </c>.
    ''' </summary>
    Private ReadOnly digitRun As New Regex("[\d]+",
                                           RegexOptions.Compiled Or RegexOptions.CultureInvariant)

    ''' <summary> Sorts a list of <c> Strings </c> against mutated sort keys built from the data, without modifying the
    ''' data itself. Returns the original strings in their sorted order. </summary>
    ''' <param name="ListToBeSorted"> A <c> list (of String)s </c> to be sorted. Left unmodified by this function </param>
    ''' <param name="textToBeReplaced"> The <c> String </c> data that will be replaced when building sort keys </param>
    ''' <param name="replacementText">The data with which <c> <paramref name="textToBeReplaced"/> </c> will be replaced </param>
    Public Function replaceAndSort(ListToBeSorted As strList, textToBeReplaced As String, replacementText As String) As strList
        If ListToBeSorted Is Nothing Then argIsNull(NameOf(ListToBeSorted)) : Return Nothing
        ' Sort keys are built alongside the originals rather than over them. Duplicate values are legal in both of the
        ' things sorted here (entry names within a category, key values within an entry), so two items may well produce
        ' the same sort key. Keeping each key beside its own index is what stops a collision from losing an original
        Dim sortKeys As New strList
        For Each item In ListToBeSorted.Items
            sortKeys.add(If(item.Contains(textToBeReplaced), item.Replace(textToBeReplaced, replacementText), item))
        Next
        ' Pad numbers if necessary
        padNumbers(sortKeys)
        ' Sort a permutation of the indices so that every sorted position still knows which original it came from.
        ' Equal sort keys break the tie on the original index, which keeps the sort stable and the output deterministic
        Dim order As New List(Of Integer)
        For i = 0 To ListToBeSorted.Count - 1
            order.Add(i)
        Next
        order.Sort(Function(a, b)
                       Dim cmp = sortKeys.Items(a).CompareTo(sortKeys.Items(b))
                       If cmp <> 0 Then Return cmp
                       Return a.CompareTo(b)
                   End Function)
        Dim sortedEntryList As New strList
        For Each idx In order
            sortedEntryList.add(ListToBeSorted.Items(idx))
        Next
        Return sortedEntryList
    End Function

    ''' <summary> Searches the <c> <paramref name="lst"/> </c> for integers and returns the length of the longest integer found </summary>
    ''' <param name="lst"> A list of strings to be searched </param>
    Private Function findLongestNumLength(ByRef lst As strList) As Integer
        Dim out = 0
        For Each item In lst.Items
            For Each mtch As Match In digitRun.Matches(item)
                If mtch.Length > out Then out = mtch.Length
            Next
        Next
        Return out
    End Function

    ''' <summary> Detects the length (number of digits) in the "longest" integer in a given list of sort keys and prepends
    ''' all shorter integers with zeros such that all the integers in all Strings are the same length.
    ''' This is to maintain numerical precedence in string sorting, ie. larger numbers come alphabetically "after" smaller numbers. </summary>
    ''' <param name="sortKeys"> The list of sort keys to be padded in place </param>
    Private Sub padNumbers(sortKeys As strList)
        Dim longestNumLen = findLongestNumLength(sortKeys)
        If longestNumLen < 2 Then Exit Sub
        Dim padTo = longestNumLen
        Dim evaluator As MatchEvaluator =
            Function(m As Match) As String
                Dim s = m.Value
                If s.IndexOf("."c) < 0 Then Return padNumberStr(padTo, s)
                Dim parts = s.Split("."c)
                For p = 0 To parts.Length - 1
                    parts(p) = padNumberStr(padTo, parts(p))
                Next
                Return String.Join(".", parts)
            End Function
        ' Assigned by index rather than by value lookup: duplicate sort keys are expected here, and a
        ' lookup would repeatedly find the first of them instead of the one being padded
        For i = 0 To sortKeys.Count - 1
            sortKeys.Items(i) = numberAndDecimals.Replace(sortKeys.Items(i), evaluator)
        Next
    End Sub

    ''' <summary> Pads a given number to a given length by prepending it with zeros (0's), returns the padded number </summary>
    ''' <param name="longestNumLen"> The desired maximum length of a number </param>
    ''' <param name="num"> The given number </param>
    Private Function padNumberStr(longestNumLen As Integer, num As String) As String
        Dim deficit = longestNumLen - num.Length
        If deficit <= 0 Then Return num
        Return New String("0"c, deficit) & num
    End Function

End Module