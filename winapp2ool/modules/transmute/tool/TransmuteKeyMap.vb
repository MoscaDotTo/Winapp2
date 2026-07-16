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
''' Implements Transmute's <c> [*Map: label] </c> key mapping rules: source-file sections which
''' match keys anywhere in the base file by KeyType and Value, and replace the whole key —
''' including its Name — with one or more replacement key lines. This is the only Transmute operation able
''' to change a key's Name, which makes data-driven category conversions possible
''' (eg. <c> Section=Google Chrome Web Browser </c> → <c> LangSecRef=3029 </c>) <br /> <br />
'''
''' Rule sections contain: <br />
''' <c> Match=&lt;Name&gt;=&lt;Value&gt; </c> — repeatable (as <c> Match1= </c>, <c> Match2= </c>, ...).
''' A base key matches when its KeyType (Name with numbers stripped) and Value both equal the
''' match key's, case-insensitively — consistent with Remove ByValue matching <br />
''' <c> Replace=&lt;Name&gt;=&lt;Value&gt; </c> — repeatable (as <c> Replace1= </c>, <c> Replace2= </c>, ...).
''' The full replacement key line(s). The first replacement takes the matched key's ordinal
''' position and any remaining replacements are inserted immediately after it in file order,
''' so a single rule can de-abstract one key into several (eg. a wildcard
''' <c> DetectFile </c> into hardcoded per-variant paths). Replacement key numbering is the
''' rule author's responsibility — keys are written as given, consistent with Add mode <br /> <br />
'''
''' Rules are applied in a single pass in file order with first-match-wins semantics: each base
''' key is evaluated against its original value only, so a key replaced by an earlier rule is
''' never re-matched by a later rule in the same run, and inserted replacement keys are never
''' evaluated at all. Replacement happens in place, preserving the section's key order
''' </summary>
Public Module TransmuteKeyMap

    ''' <summary>
    ''' The section name prefix identifying a key mapping rule, eg. <c> [*Map: Chrome] </c>
    ''' </summary>
    Public Const MapSectionPrefix As String = "*Map:"

    ''' <summary>
    ''' A parsed <c> [*Map: label] </c> rule: the match criteria, the replacement key,
    ''' and hit counters populated during application
    ''' </summary>
    Private Class TransmuteMapRule

        ''' <summary> The human-readable label from the section name, used in reporting </summary>
        Public Property Label As String

        ''' <summary> The match criteria, each parsed into an <c> iniKey2 </c> for KeyType+Value comparison </summary>
        Public Property Matches As New List(Of iniKey2)

        ''' <summary> The replacement key lines in file order, cloned for each hit </summary>
        Public Property Replacements As New List(Of iniKey2)

        ''' <summary> The total number of keys replaced by this rule </summary>
        Public Property Hits As Integer

        ''' <summary> The number of sections in which this rule replaced at least one key </summary>
        Public Property SectionsHit As Integer

    End Class

    ''' <summary>
    ''' Parses and applies a set of <c> [*Map:] </c> rule sections against every section of the
    ''' <c> <paramref name="baseFile"/> </c>. <br /> <br />
    '''
    ''' Rules are only applied when the current transmute mode is Replace ByKey (the Flavorizer
    ''' key-replacement stage, or a standalone Replace ByKey run); under any other mode the rules
    ''' are skipped with a warning. Malformed rules (no <c> Replace= </c>, no <c> Match= </c>, or
    ''' an unsplittable value) are skipped individually with a warning. <br /> <br />
    '''
    ''' The menu receives one summary line per rule; per-hit detail is written to the log.
    ''' A rule which matches nothing across the whole file emits a warning, signalling that the
    ''' rule may be stale
    ''' </summary>
    '''
    ''' <param name="baseFile">
    ''' The <c> iniFile2 </c> whose keys will be evaluated against the mapping rules
    ''' </param>
    '''
    ''' <param name="mapSections">
    ''' The source-file sections whose names begin with <c> *Map: </c>
    ''' </param>
    '''
    ''' <param name="menuOutput">
    ''' The <c> MenuSection </c> containing output to be displayed to the user
    ''' </param>
    '''
    Friend Sub applyKeyMapRules(ByRef baseFile As iniFile2,
                                      mapSections As List(Of iniSection2),
                                ByRef menuOutput As MenuSection)

        If mapSections.Count = 0 Then Return

        If Not (Transmutator = TransmuteMode.Replace AndAlso TransmuteReplaceMode = ReplaceMode.ByKey) Then

            Dim wrongModeMsg = $"[*Map:] rules are only applied in Replace ByKey mode - skipping {mapSections.Count} rule section(s)"
            menuOutput.AddWarning(wrongModeMsg)
            gLog(wrongModeMsg)
            Return

        End If

        Dim rules As New List(Of TransmuteMapRule)

        For Each mapSection In mapSections

            Dim rule = parseMapRule(mapSection, menuOutput)
            If rule IsNot Nothing Then rules.Add(rule)

        Next

        If rules.Count = 0 Then Return

        Using gLogScope($"Applying {rules.Count} key mapping rule(s) to {baseFile.Count} sections")

            For Each baseSection In baseFile

                Dim rulesHitInSection As New HashSet(Of TransmuteMapRule)

                ' Snapshot the key list so replacement keys inserted during this pass are
                ' never re-evaluated against later rules (first-match-wins on original values)
                For Each baseKey In baseSection.Keys.ToList()

                    For Each rule In rules

                        If Not ruleMatches(rule, baseKey) Then Continue For

                        Dim newKeys As New List(Of iniKey2)
                        For Each replacement In rule.Replacements : newKeys.Add(New iniKey2(replacement.ToString())) : Next
                        baseSection.Keys.Replace(baseKey, newKeys)

                        rule.Hits += 1
                        rulesHitInSection.Add(rule)
                        gLog($"{baseSection.Name}: {baseKey} -> {String.Join(" | ", newKeys)}")

                        Exit For

                    Next

                Next

                For Each rule In rulesHitInSection : rule.SectionsHit += 1 : Next

            Next

        End Using

        For Each rule In rules

            If rule.Hits = 0 Then

                Dim staleMsg = $"[Map: {rule.Label}] matched nothing in {baseFile.Name} - the rule may be stale"
                menuOutput.AddWarning(staleMsg)
                gLog(staleMsg)
                Continue For

            End If

            Dim matchDesc = String.Join(" | ", rule.Matches)
            Dim replaceDesc = String.Join(" | ", rule.Replacements)
            Dim summary = $"*[Map: {rule.Label}] {matchDesc} -> {replaceDesc} ({rule.SectionsHit} sections)"
            menuOutput.AddColoredLine(summary, ConsoleColor.Yellow)
            gLog(summary)

        Next

    End Sub

    ''' <summary>
    ''' Determines whether <c> <paramref name="baseKey"/> </c> satisfies any of the match criteria
    ''' of <c> <paramref name="rule"/> </c>: KeyType and Value must both match, case-insensitively.
    ''' Comparing KeyTypes strips numbers from both sides, so <c> Match=FileKey=... </c> matches
    ''' any <c> FileKeyN </c> with that value
    ''' </summary>
    '''
    ''' <param name="rule">
    ''' The rule providing the match criteria
    ''' </param>
    '''
    ''' <param name="baseKey">
    ''' The base file key being evaluated
    ''' </param>
    '''
    Private Function ruleMatches(rule As TransmuteMapRule, baseKey As iniKey2) As Boolean

        For Each matchKey In rule.Matches

            If baseKey.KeyType.Equals(matchKey.KeyType, StringComparison.OrdinalIgnoreCase) AndAlso
               baseKey.Value.Equals(matchKey.Value, StringComparison.OrdinalIgnoreCase) Then Return True

        Next

        Return False

    End Function

    ''' <summary>
    ''' Parses a <c> [*Map: label] </c> section into a <c> TransmuteMapRule </c>, emitting a
    ''' warning and returning <c> Nothing </c> when the section is malformed: no <c> Match= </c>
    ''' keys, no <c> Replace= </c> keys, an unrecognized key, or a match/replace
    ''' value which does not itself contain a <c> Name=Value </c> pair
    ''' </summary>
    '''
    ''' <param name="mapSection">
    ''' The source-file section to parse, whose name begins with <c> *Map: </c>
    ''' </param>
    '''
    ''' <param name="menuOutput">
    ''' The <c> MenuSection </c> containing output to be displayed to the user
    ''' </param>
    '''
    Private Function parseMapRule(mapSection As iniSection2,
                                  ByRef menuOutput As MenuSection) As TransmuteMapRule

        Dim label = mapSection.Name.Substring(MapSectionPrefix.Length).Trim()
        If label.Length = 0 Then label = mapSection.Name

        Dim rule As New TransmuteMapRule With {.Label = label}

        For Each key In mapSection.Keys

            Select Case True

                Case key.KeyType.Equals("Match", StringComparison.OrdinalIgnoreCase)

                    Dim matchKey = parseKeyLine(key.Value)

                    If matchKey Is Nothing Then

                        warnMalformedRule(label, $"{key.Name} value '{key.Value}' is not a Name=Value pair", menuOutput)
                        Return Nothing

                    End If

                    rule.Matches.Add(matchKey)

                Case key.KeyType.Equals("Replace", StringComparison.OrdinalIgnoreCase)

                    Dim replacementKey = parseKeyLine(key.Value)

                    If replacementKey Is Nothing Then

                        warnMalformedRule(label, $"{key.Name} value '{key.Value}' is not a Name=Value pair", menuOutput)
                        Return Nothing

                    End If

                    rule.Replacements.Add(replacementKey)

                Case Else

                    warnMalformedRule(label, $"unrecognized key '{key.Name}' (expected Match or Replace)", menuOutput)
                    Return Nothing

            End Select

        Next

        If rule.Matches.Count = 0 Then

            warnMalformedRule(label, "no Match= keys were provided", menuOutput)
            Return Nothing

        End If

        If rule.Replacements.Count = 0 Then

            warnMalformedRule(label, "no Replace= key was provided", menuOutput)
            Return Nothing

        End If

        Return rule

    End Function

    ''' <summary>
    ''' Parses a <c> Name=Value </c> line into an <c> iniKey2 </c>, returning <c> Nothing </c>
    ''' when the line has no <c> = </c> or an empty Name
    ''' </summary>
    '''
    ''' <param name="line">
    ''' The raw key line, ie. the value of a <c> Match= </c> or <c> Replace= </c> key
    ''' </param>
    '''
    Private Function parseKeyLine(line As String) As iniKey2

        Dim eqPos = line.IndexOf("="c)
        If eqPos < 1 Then Return Nothing

        Return New iniKey2(line)

    End Function

    ''' <summary>
    ''' Emits the warning for a malformed <c> [*Map:] </c> rule to both the menu and the log
    ''' </summary>
    '''
    ''' <param name="label">
    ''' The rule's label, for identification in the warning
    ''' </param>
    '''
    ''' <param name="reason">
    ''' A short description of what makes the rule malformed
    ''' </param>
    '''
    ''' <param name="menuOutput">
    ''' The <c> MenuSection </c> containing output to be displayed to the user
    ''' </param>
    '''
    Private Sub warnMalformedRule(label As String,
                                  reason As String,
                                  ByRef menuOutput As MenuSection)

        Dim malformedMsg = $"[Map: {label}] is malformed and will be skipped: {reason}"
        menuOutput.AddWarning(malformedMsg)
        gLog(malformedMsg)

    End Sub

End Module
