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
''' One survivor of the FileKey merge pass. Holds the parsed pieces of the first key seen
''' for a given (path, flag) pair, with the path compared ignoring case, plus the running
''' pattern list from anything folded into it afterwards.
''' </summary>
Friend Class fileKeySurvivor

    ''' <summary>The parsed components of the first-seen key in this group</summary>
    Public ReadOnly Property Params As fileKeyParams

    ''' <summary>
    ''' Accumulated patterns, seeded from <c> Params.Patterns </c> and appended to on each merge.
    ''' Duplicates are kept.
    ''' </summary>
    Public ReadOnly Property Patterns As List(Of String)

    ''' <summary>Creates a new <c> fileKeySurvivor </c> holding a copy of <paramref name="parsed"/>'s patterns</summary>
    ''' <param name="parsed">The first key seen for this path and flag</param>
    Public Sub New(parsed As fileKeyParams)
        Me.Params = parsed
        Me.Patterns = New List(Of String)(parsed.Patterns)
    End Sub

End Class

''' <summary>
''' Holds the FileKey merge behind the Optimizations rule, the one <c> WinappDebug </c>
''' rule that is off by default.
''' </summary>
Module experimentalScans

    ''' <summary>
    ''' Merges FileKeys that share a path (ignoring case) and flag. When any merge is possible,
    ''' we queue a report block on <paramref name="result"/> listing the keys folded away and the
    ''' resulting key list. <see cref="ErrorsFound"/> doesn't count it.
    ''' <br /><br />
    '''
    ''' If the Optimizations rule's <c> ShouldRepair </c> is on (checked directly, not through
    ''' <see cref="lintRule.fixFormat"/>), we replace the whole FileKey bucket with one new key per
    ''' path and flag, in first-seen order and numbered from 1. Each holds the patterns of every
    ''' key folded into it, in order and without removing duplicates. Every unrecognized flag
    ''' counts as the same flag, and the merged key keeps the first one's text.
    ''' </summary>
    '''
    ''' <remarks>
    ''' Every FileKey value is parsed once through <c> fileKeyParams </c> and matched by
    ''' (path, flag) in O(1). Merging several keys into the same survivor never rebuilds or
    ''' re-parses the running value. Each survivor's merged value string gets built once at
    ''' the end, when the output keys are emitted.
    ''' </remarks>
    '''
    ''' <param name="result">The lint result that receives the merge report</param>
    '''
    ''' <param name="entry">
    ''' The <c> winapp2entry </c> whose FileKeys will be assessed for merge opportunities
    ''' </param>
    Public Sub cOptimization(result As EntryLintResult, entry As winapp2entry)

        If entry.FileKeys.Count < 2 Then Return

        Dim survivors As New List(Of fileKeySurvivor)
        Dim indexByKey As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
        Dim removedKeys As New List(Of iniKey)

        For Each key In entry.FileKeys

            Dim parsed As New fileKeyParams(key.Value)
            Dim hashKey = parsed.Path & ChrW(0) & CInt(parsed.Flag).ToString()

            Dim idx As Integer

            If indexByKey.TryGetValue(hashKey, idx) Then

                ' The path and flag both match, so fold this key's patterns into the
                ' survivor and mark the original for removal
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
        Dim resultKeys As New List(Of iniKey)

        For i = 0 To survivors.Count - 1

            Dim s = survivors(i)
            Dim newValue = s.Params.Path

            If s.Patterns.Count > 0 Then newValue &= "|" & String.Join(";", s.Patterns)

            Select Case s.Params.Flag
                Case fileKeyFlag.Recurse : newValue &= "|RECURSE"
                Case fileKeyFlag.RemoveSelf : newValue &= "|REMOVESELF"
                Case fileKeyFlag.Unknown : newValue &= "|" & s.Params.RawFlag
            End Select

            resultKeys.Add(New iniKey($"FileKey{i + 1}={newValue}"))

        Next

        result.DeferSection(New MenuSection().AddColoredLine(result.EntryName & " has keys which can be merged", ConsoleColor.Magenta, centered:=True))
        result.DeferSection(buildOptiSect(removedKeys, resultKeys))

        If lintOpti.ShouldRepair Then entry.ReplaceFileKeys(resultKeys)

    End Sub

    ''' <summary>Builds a labeled section of keys for optimization output</summary>
    Private Function buildOptiSect(Remkeys As IEnumerable(Of iniKey),
                                   resultKeys As IEnumerable(Of iniKey)) As MenuSection

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
