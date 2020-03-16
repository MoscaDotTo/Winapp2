'    Copyright (C) 2018-2020 Robbie Ward
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
''' <summary>
''' An object representing a section of a .ini file
''' </summary>
Public Class iniSection
    ''' <summary> The line number from which the <c> Name </c> of the <c> iniSection </c> was originally read </summary>
    Public Property StartingLineNumber As Integer
    ''' <summary> The line number from which the last <c> iniKey </c> in the <c> iniSection </c> was originally read </summary>
    Public Property EndingLineNumber As Integer
    ''' <summary> The name of the <c> iniSection </c> without [Braces] </summary>
    Public Property Name As String
    ''' <summary>The list of <c> iniKeys </c> contained in the <c> iniSection </c></summary>
    Public Property Keys As New keyList


    ''' <summary> Sorts a section's <c>iniKeys </c> into <c> keyLists </c> based on their <c> KeyType </c> </summary>
    ''' <param name="listOfKeyLists">The list of <c>KeyLists </c> into which <c> iniKeys </c> will be sorted <br />
    ''' The last list in the keylistlist holds the keys determined to contain errors </param> 
    Public Sub constKeyLists(ByRef listOfKeyLists As List(Of keyList))
        If listOfKeyLists Is Nothing Then argIsNull(NameOf(listOfKeyLists)) : Return
        Dim keyTypeList As New List(Of String)
        listOfKeyLists.ForEach(Sub(kl) keyTypeList.Add(kl.KeyType.ToUpperInvariant))
        For Each key In Keys.Keys
            Dim type = key.KeyType.ToUpperInvariant
            If keyTypeList.Contains(type) Then listOfKeyLists(keyTypeList.IndexOf(type)).add(key) Else listOfKeyLists.Last.add(key)
        Next
    End Sub

    ''' <summary>Removes keys from the <c> iniSection </c> by their index </summary>
    ''' <param name="indicies"> A given list of indicies pointing to keys to remove from the <c> iniSection </c> </param>
    Public Sub removeKeys(indicies As List(Of Integer))
        If indicies Is Nothing Then argIsNull(NameOf(indicies)) : Return
        ' Sort and reverse the indicies so that we remove from the end of the list towards the beginning
        indicies.Sort()
        indicies.Reverse()
        indicies.ForEach(Sub(ind) Keys.remove(ind))
    End Sub

    ''' <summary> Returns the <c> iniSection's </c> <c> Name </c> as it would appear on disk </summary>
    Public Function getFullName() As String
        Return $"[{Name}]"
    End Function

    ''' <summary> Creates a new (empty) <c> iniSection </c> object </summary>
    Public Sub New()
        StartingLineNumber = 0
        EndingLineNumber = 0
        Name = ""
        Keys = New keyList
    End Sub

    ''' <summary> Creates a new <c> iniSection </c> object without tracking the line numbers </summary>
    ''' <param name="listOfLines">The list of Strings comprising the <c> iniSection </c></param>
    ''' <param name="listOfLineCounts">The list of line numbers associated with the lines </param>
    Public Sub New(ByVal listOfLines As List(Of String), Optional listOfLineCounts As List(Of Integer) = Nothing)
        If listOfLines Is Nothing Then argIsNull(NameOf(listOfLines)) : Return
        Name = listOfLines(0).Trim(CChar("["), CChar("]"))
        StartingLineNumber = If(listOfLineCounts IsNot Nothing, listOfLineCounts(0), 1)
        EndingLineNumber = StartingLineNumber + listOfLines.Count
        Keys = New keyList
        If listOfLines.Count > 1 Then
            For i As Integer = 1 To listOfLines.Count - 1
                Keys.add(New iniKey(listOfLines(i), If(listOfLineCounts Is Nothing, 0, listOfLineCounts(i))))
            Next
        End If
    End Sub

    ''' <summary> Returns the keys in the <c> iniSection </c> as a <c> strList </c> </summary>
    Public Function getKeysAsStrList() As strList
        Dim out As New strList
        For Each key In Keys.Keys
            out.add(key.toString)
        Next
        Return out
    End Function

    ''' <summary> Compares two <c> iniSections </c>, returns <c> False </c> if they are not the same </summary>
    ''' <param name="ss"> The <c> iniSection </c> against which this object will be compared </param>
    ''' <param name="removedKeys"> A return <c> keyList </c> that appear in this <c> iniSection </c> object but not the given <c> iniSection </c> </param>
    ''' <param name="addedKeys"> A return <c> keyList </c> that appear in the given <c> iniSection </c> object but not this object </param>
    Public Function compareTo(ss As iniSection, ByRef removedKeys As keyList, ByRef addedKeys As keyList) As Boolean
        If ss Is Nothing Then argIsNull(NameOf(ss)) : Return False
        If removedKeys Is Nothing Then argIsNull(NameOf(removedKeys)) : Return False
        If addedKeys Is Nothing Then argIsNull(NameOf(addedKeys)) : Return False
        ' Create a copy of the section so we can modify it
        Dim secondSection As New iniSection With {.Name = ss.Name, .StartingLineNumber = ss.StartingLineNumber}
        For i = 0 To ss.Keys.KeyCount - 1
            secondSection.Keys.add(ss.Keys.Keys(i))
        Next
        Dim noMatch As Boolean
        Dim tmpList As New List(Of Integer)
        For Each key In Keys.Keys
            noMatch = True
            For i = 0 To secondSection.Keys.KeyCount - 1
                Dim sKey = secondSection.Keys.Keys(i)
                Select Case True
                    Case key.compareTypes(sKey) And key.compareValues(sKey)
                        noMatch = False
                        ' In the case where a change includes that a line has been duplicated, nasty things can happen :(
                        If Not tmpList.Contains(i) Then tmpList.Add(i)
                        Exit For
                End Select
            Next
            ' If the key isn't found in the second (newer) section, consider it removed for now
            removedKeys.add(key, noMatch)
        Next
        ' Remove all matched keys
        secondSection.removeKeys(tmpList)
        ' Assume any remaining keys have been added
        For Each key In secondSection.Keys.Keys
            addedKeys.add(key)
        Next
        Return removedKeys.KeyCount + addedKeys.KeyCount = 0
    End Function

    ''' <summary> Returns the <c> iniSection </c> as it would appear on disk as a <c> String </c> </summary>
    Public Overrides Function ToString() As String
        Dim out = Me.getFullName
        For Each key In Keys.Keys
            out += Environment.NewLine & key.toString
        Next
        out += Environment.NewLine
        Return out
    End Function
End Class