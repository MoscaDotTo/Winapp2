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

''' <summary>
''' Holds scans and repairs for <c> WinappDebug </c> that are disabled by default —
''' either incomplete or out of scope for the normal lint process.
''' </summary>
Module experimentalScans

    ''' <summary>Holds the full text of every RegKey seen during cross-entry duplicate checks</summary>
    Private Property regKeyTracker As New HashSet(Of String)

    ''' <summary>Holds the path of every FileKey seen during cross-entry duplicate checks</summary>
    Private Property fileKeyTracker As New HashSet(Of String)

    ''' <summary>Holds the path of every Detect seen during cross-entry duplicate checks</summary>
    Private Property detectTracker As New HashSet(Of String)

    ''' <summary>Holds the path of every DetectFile seen during cross-entry duplicate checks</summary>
    Private Property detectFileTracker As New HashSet(Of String)

    ''' <summary>Clears all cross-entry key trackers between lint runs</summary>
    Public Sub resetKeyTrackers()
        regKeyTracker.Clear()
        fileKeyTracker.Clear()
        detectFileTracker.Clear()
        detectTracker.Clear()
    End Sub

    ''' <summary>
    ''' Attempts to merge FileKeys with identical paths and flags into a single key.
    ''' When two FileKeys share the same path and deletion flag, their patterns are
    ''' combined into the earlier key and the later key is removed.
    ''' </summary>
    '''
    ''' <param name="entry">
    ''' The <c> winapp2entry2 </c> whose FileKeys will be assessed for merge opportunities
    ''' </param>
    Public Sub cOptimization(entry As winapp2entry2)

        If entry.FileKeys.Count < 2 Then Return

        ' seenPaths(i)/seenFlags(i)/mergedValues(i) track the state of each FileKey
        ' as it is processed; mergedValues(i) is updated in place when another key merges into it.
        Dim seenPaths As New List(Of String)
        Dim seenFlags As New List(Of fileKeyFlag)
        Dim mergedValues As New List(Of String)
        Dim removedIndices As New HashSet(Of Integer)

        For i = 0 To entry.FileKeys.Count - 1

            Dim current = New fileKeyParams2(entry.FileKeys(i).Value)
            seenPaths.Add(current.Path)
            seenFlags.Add(current.Flag)
            mergedValues.Add(entry.FileKeys(i).Value)

            If i = 0 Then Continue For

            ' Find the earliest surviving key with the same path and flag
            Dim matchIdx = -1
            For j = 0 To i - 1
                If Not removedIndices.Contains(j) AndAlso
                   seenPaths(j).Equals(current.Path, StringComparison.OrdinalIgnoreCase) AndAlso
                   seenFlags(j) = current.Flag Then
                    matchIdx = j
                    Exit For
                End If
            Next

            If matchIdx < 0 Then Continue For

            gLog($"{entry.FileKeys(i)} has a path that matches another key")
            gLog($"Matching key has index {matchIdx} in the unique path list")

            ' Append current key's patterns to the surviving key's value
            Dim mergeTarget = New fileKeyParams2(mergedValues(matchIdx))
            Dim combined = New List(Of String)(mergeTarget.Patterns)
            combined.AddRange(current.Patterns)

            Dim newValue = mergeTarget.Path & "|" & String.Join(";", combined)
            Select Case current.Flag
                Case fileKeyFlag.Recurse : newValue &= "|RECURSE"
                Case fileKeyFlag.RemoveSelf : newValue &= "|REMOVESELF"
                Case fileKeyFlag.Unknown : newValue &= "|" & current.RawFlag
            End Select

            mergedValues(matchIdx) = newValue
            removedIndices.Add(i)
            gLog($"Key will be merged and have the new value: {newValue}")

        Next

        If removedIndices.Count = 0 Then Return

        ' Build the merged result and the list of keys being removed, renumbering from 1
        Dim resultKeys As New List(Of iniKey2)
        Dim removedKeys As New List(Of iniKey2)
        Dim keyNum = 1

        For i = 0 To entry.FileKeys.Count - 1
            If removedIndices.Contains(i) Then
                removedKeys.Add(entry.FileKeys(i))
            Else
                resultKeys.Add(New iniKey2($"FileKey{keyNum}={mergedValues(i)}"))
                keyNum += 1
            End If
        Next

        printOptiSect("Optimization opportunity detected", entry.FileKeys)
        printOptiSect("The following keys can be merged into other keys:", removedKeys)
        printOptiSect("The resulting key list will be reduced to:", resultKeys)

        If lintOpti.ShouldRepair Then entry.ReplaceFileKeys(resultKeys)

    End Sub

    ''' <summary>Prints a labeled section of keys for optimization output</summary>
    Private Sub printOptiSect(boxStr As String, keys As IEnumerable(Of iniKey2))
        print(3, boxStr, buffr:=True, trailr:=True)
        For Each key In keys
            cwl(key.ToString())
        Next
        cwl()
    End Sub

End Module
