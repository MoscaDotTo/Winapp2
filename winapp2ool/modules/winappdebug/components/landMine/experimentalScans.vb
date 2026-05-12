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
''' One survivor in the FileKey merge pass. Holds the parsed components of the
''' first-seen key for a given (path, flag) combination and the running list of
''' merged patterns from any subsequent keys that get folded into it.
''' </summary>
Friend Class fileKeySurvivor

    ''' <summary>The parsed components of the first-seen key in this group</summary>
    Public ReadOnly Property Params As fileKeyParams2

    ''' <summary>Accumulated patterns; seeded from <c>Params.Patterns</c> and appended to on each merge</summary>
    Public ReadOnly Property Patterns As List(Of String)

    Public Sub New(parsed As fileKeyParams2)
        Me.Params = parsed
        Me.Patterns = New List(Of String)(parsed.Patterns)
    End Sub

End Class

''' <summary>
''' Holds scans and repairs for <c> WinappDebug </c> that are disabled by default —
''' either incomplete or out of scope for the normal lint process.
''' </summary>
Module experimentalScans

    ''' <summary>
    ''' Attempts to merge FileKeys with identical paths and flags into a single key.
    ''' When two FileKeys share the same path and deletion flag, their patterns are
    ''' combined into the earlier key and the later key is removed.
    ''' </summary>
    '''
    ''' <remarks>
    ''' Each FileKey value is parsed exactly once via <c>fileKeyParams2</c>. Match
    ''' lookup is O(1) by (path, flag) — chained merges into the same survivor do
    ''' not rebuild or re-parse the running value. The merged value string is
    ''' constructed once per survivor at the end, when output keys are emitted.
    ''' </remarks>
    '''
    ''' <param name="entry">
    ''' The <c> winapp2entry2 </c> whose FileKeys will be assessed for merge opportunities
    ''' </param>
    Public Sub cOptimization(result As EntryLintResult, entry As winapp2entry2)

        If entry.FileKeys.Count < 2 Then Return

        Dim survivors As New List(Of fileKeySurvivor)
        Dim indexByKey As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
        Dim removedKeys As New List(Of iniKey2)

        For Each key In entry.FileKeys

            Dim parsed As New fileKeyParams2(key.Value)
            Dim hashKey = parsed.Path & ChrW(0) & CInt(parsed.Flag).ToString()

            Dim idx As Integer

            If indexByKey.TryGetValue(hashKey, idx) Then

                ' Path+flag match: fold this key's patterns into the survivor and
                ' record the original key for removal.
                survivors(idx).Patterns.AddRange(parsed.Patterns)
                removedKeys.Add(key)
                gLog($"{key} has a path that matches another key")

            Else

                indexByKey(hashKey) = survivors.Count
                survivors.Add(New fileKeySurvivor(parsed))

            End If

        Next

        If removedKeys.Count = 0 Then Return

        ' Build output keys once per survivor, renumbering from 1.
        Dim resultKeys As New List(Of iniKey2)

        For i = 0 To survivors.Count - 1

            Dim s = survivors(i)
            Dim newValue = s.Params.Path

            If s.Patterns.Count > 0 Then newValue &= "|" & String.Join(";", s.Patterns)

            Select Case s.Params.Flag
                Case fileKeyFlag.Recurse : newValue &= "|RECURSE"
                Case fileKeyFlag.RemoveSelf : newValue &= "|REMOVESELF"
                Case fileKeyFlag.Unknown : newValue &= "|" & s.Params.RawFlag
            End Select

            resultKeys.Add(New iniKey2($"FileKey{i + 1}={newValue}"))

        Next

        result.DeferSection(New MenuSection().AddColoredLine(result.EntryName & " has keys which can be merged", ConsoleColor.Magenta, centered:=True))
        result.DeferSection(buildOptiSect(removedKeys, resultKeys))

        If lintOpti.ShouldRepair Then entry.ReplaceFileKeys(resultKeys)

    End Sub

    ''' <summary>Builds a labeled section of keys for optimization output</summary>
    Private Function buildOptiSect(Remkeys As IEnumerable(Of iniKey2),
                                   resultKeys As IEnumerable(Of iniKey2)) As MenuSection

        Dim out As New MenuSection
        out.AddDivider(solid:=False).AddLine("The following keys can be merged into other keys:", True).AddDivider(solid:=False)
        For Each key In Remkeys : out.AddLine(key.ToString) : Next
        out.AddDivider(solid:=False)
        out.AddLine("The resulting key list will be reduced to:", True).AddDivider(solid:=False)
        For Each key In resultKeys : out.AddLine(key.ToString) : Next
        out.AddBlank()
        Return out

    End Function

End Module
