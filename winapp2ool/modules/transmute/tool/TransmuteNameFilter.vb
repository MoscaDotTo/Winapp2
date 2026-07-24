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

Imports System.Linq
Imports System.Text.RegularExpressions

''' <summary>
''' Implements Transmute's <c> [*Name: scaffold] </c> sentinel: a name-filtered global operation.
''' Where <c> [*] </c> applies its keys to every section, a <c> [*Name:] </c> section applies its
''' payload keys only to base sections selected by two orthogonal filters — an anchored name suffix
''' and an optional set of content predicates. <br /> <br />
'''
''' The section header carries a <c> scaffold </c> label; a base section is selected when: <br />
''' 1. its name ends with <c> " scaffold *" </c> (case-insensitive, anchored on the winapp2 entry
'''    terminator so <c> "… Web Browsing Session *" </c> does not also catch
'''    <c> "… Web Browsing Session Backups *" </c>), AND <br />
''' 2. it satisfies the rule's <c> Match= </c> predicates — no predicates means the name filter
'''    alone decides; otherwise the section must contain at least one key matching any one predicate
'''    (KeyType equal, value equal or glob-matched; OR-semantics across the set) <br /> <br />
'''
''' The two filters do different jobs: the name suffix selects <i> which </i> scaffold, the
''' predicate set confirms <i> what kind </i> of entry (eg. that it is a browser via
''' <c> Match=Section=* Web Browser </c>). The predicate set cannot separate one scaffold from
''' another, so the anchored suffix must carry that distinction. <br /> <br />
'''
''' Every key that is not a <c> Match= </c> key is payload, applied under the current transmute
''' mode. Payload values support the <c> %EntryName% </c> token like <c> [*] </c>. In Add mode a
''' payload key already present in a selected section is skipped, so the operation never duplicates
''' an existing key
''' </summary>
Friend Module TransmuteNameFilter

    ''' <summary>
    ''' The section name prefix identifying a name-filter rule, eg. <c> [*Name: Bookmark Backups] </c>
    ''' </summary>
    Public Const NameSectionPrefix As String = "*Name:"

    ''' <summary>
    ''' A parsed <c> [*Name: scaffold] </c> rule: the scaffold suffix label, the optional content
    ''' predicates, and the payload keys to apply to each selected section
    ''' </summary>
    Private Class TransmuteNameRule

        ''' <summary> The scaffold label from the section name, matched as an anchored name suffix </summary>
        Public Property Scaffold As String

        ''' <summary> The content predicates, each parsed into an <c> iniKey2 </c> for KeyType+Value matching </summary>
        Public Property Matches As New List(Of iniKey2)

        ''' <summary> The keys to apply to each selected section under the current transmute mode </summary>
        Public Property Payload As New List(Of iniKey2)

    End Class

    ''' <summary>
    ''' Parses and applies a set of <c> [*Name:] </c> rule sections against the sections of the
    ''' <c> <paramref name="baseFile"/> </c> under the current transmute mode. <br /> <br />
    '''
    ''' Section-level modes are refused (Remove BySection and Replace BySection ignore the payload
    ''' keys entirely, making a name-filtered whole-section operation incoherent here). The menu
    ''' receives one summary line per rule; per-hit detail is written to the log. A rule which
    ''' selects no sections emits a warning, signalling that the rule may be stale
    ''' </summary>
    '''
    ''' <param name="baseFile">
    ''' The <c> iniFile2 </c> whose sections will be filtered and modified
    ''' </param>
    '''
    ''' <param name="nameSections">
    ''' The source-file sections whose names begin with <c> *Name: </c>
    ''' </param>
    '''
    ''' <param name="menuOutput">
    ''' The <c> MenuSection </c> containing output to be displayed to the user
    ''' </param>
    '''
    Public Sub applyNameFilterSections(ByRef baseFile As iniFile2,
                                             nameSections As List(Of iniSection2),
                                       ByRef menuOutput As MenuSection)

        If nameSections.Count = 0 Then Return

        If Transmutator = TransmuteMode.Remove AndAlso TransmuteRemoveMode = RemoveMode.BySection Then

            Dim refuseRemMsg = "[*Name:] cannot be used with Remove BySection (a name-filtered section removal ignores the payload keys) - skipping"
            menuOutput.AddWarning(refuseRemMsg)
            gLog(refuseRemMsg)
            Return

        End If

        If Transmutator = TransmuteMode.Replace AndAlso TransmuteReplaceMode = ReplaceMode.BySection Then

            Dim refuseReplMsg = "[*Name:] cannot be used with Replace BySection (a name-filtered section replacement ignores the payload keys) - skipping"
            menuOutput.AddWarning(refuseReplMsg)
            gLog(refuseReplMsg)
            Return

        End If

        Dim rules As New List(Of TransmuteNameRule)

        For Each nameSection In nameSections

            Dim rule = parseNameRule(nameSection, menuOutput)
            If rule IsNot Nothing Then rules.Add(rule)

        Next

        If rules.Count = 0 Then Return

        Dim totalSections = baseFile.Count

        Using gLogScope($"Applying {rules.Count} name-filter section(s) ({Transmutator}) to {totalSections} sections")

            For Each rule In rules

                ' Adding a numbered key to every selected section creates instant duplicates, so
                ' refuse them the same way [*] does — drop the offenders and keep any unnumbered rest
                If Transmutator = TransmuteMode.Add Then

                    For Each numberedKey In rule.Payload.Where(Function(k) Not k.Name.Equals(k.KeyType, StringComparison.OrdinalIgnoreCase)).ToList()

                        Dim refuseKeyMsg = $"[*Name: {rule.Scaffold}] Refusing to add numbered key {numberedKey.Name} - name-filter adds must be unnumbered"
                        menuOutput.AddWarning(refuseKeyMsg)
                        gLog(refuseKeyMsg)
                        rule.Payload.Remove(numberedKey)

                    Next

                    If rule.Payload.Count = 0 Then Continue For

                End If

                Dim suffix = $" {rule.Scaffold} *"
                Dim matchedCount = 0
                Dim appliedCount = 0

                For Each baseSection In baseFile

                    If Not baseSection.Name.EndsWith(suffix, StringComparison.OrdinalIgnoreCase) Then Continue For
                    If Not sectionSatisfiesPredicates(baseSection, rule.Matches) Then Continue For

                    matchedCount += 1
                    If applyPayload(baseSection, rule.Payload, menuOutput) > 0 Then appliedCount += 1

                Next

                If matchedCount = 0 Then

                    Dim staleMsg = $"[*Name: {rule.Scaffold}] selected no sections in {baseFile.Name} - the rule may be stale"
                    menuOutput.AddWarning(staleMsg)
                    gLog(staleMsg)
                    Continue For

                End If

                Dim payloadDesc = String.Join(", ", rule.Payload)
                Dim summary As String
                Dim summaryColor As ConsoleColor

                Select Case Transmutator

                    Case TransmuteMode.Add

                        summary = $"+[*Name: {rule.Scaffold}] Added {payloadDesc} to {appliedCount} of {matchedCount} selected sections"
                        summaryColor = ConsoleColor.Green

                    Case TransmuteMode.Replace

                        summary = $"*[*Name: {rule.Scaffold}] Replaced {payloadDesc} in {appliedCount} of {matchedCount} selected sections"
                        summaryColor = ConsoleColor.Yellow

                    Case Else

                        summary = $"-[*Name: {rule.Scaffold}] Removed {payloadDesc} from {appliedCount} of {matchedCount} selected sections"
                        summaryColor = ConsoleColor.Red

                End Select

                menuOutput.AddColoredLine(summary, summaryColor)
                gLog(summary)

            Next

        End Using

    End Sub

    ''' <summary>
    ''' Applies a rule's payload keys to a single selected <c> <paramref name="baseSection"/> </c>
    ''' under the current transmute mode, expanding the <c> %EntryName% </c> token per section.
    ''' In Add mode, a payload key whose Name already exists in the section is skipped
    ''' </summary>
    '''
    ''' <param name="baseSection">
    ''' The selected base section receiving the payload
    ''' </param>
    '''
    ''' <param name="payload">
    ''' The rule's payload keys
    ''' </param>
    '''
    ''' <param name="menuOutput">
    ''' The <c> MenuSection </c> passed through to the reused Replace/Remove helpers (quiet)
    ''' </param>
    '''
    ''' <returns>
    ''' The number of payload keys applied to <c> <paramref name="baseSection"/> </c>
    ''' </returns>
    Public Function applyPayload(baseSection As iniSection2,
                                  payload As List(Of iniKey2),
                            ByRef menuOutput As MenuSection) As Integer

        Dim payloadSource As New iniSection2(baseSection.Name)

        For Each payloadKey In payload

            Dim value = payloadKey.Value
            If value.IndexOf(EntryNameToken, StringComparison.OrdinalIgnoreCase) >= 0 Then value = expandEntryNameToken(value, baseSection.Name)
            payloadSource.AddKey(New iniKey2($"{payloadKey.Name}={value}"))

        Next

        Select Case Transmutator

            Case TransmuteMode.Add

                Dim hits = 0

                For Each payloadKey In payloadSource.Keys

                    If baseSection.HasKey(payloadKey.Name) Then Continue For

                    baseSection.AddKey(New iniKey2(payloadKey.ToString()))
                    hits += 1
                    gLog($"{baseSection.Name}: += {payloadKey}")

                Next

                Return hits

            Case TransmuteMode.Replace

                Return replaceKeysInBase(baseSection, payloadSource, menuOutput, quiet:=True)

            Case Else

                Return remKeys(baseSection, payloadSource, menuOutput, quiet:=True)

        End Select

    End Function

    ''' <summary>
    ''' Determines whether <c> <paramref name="baseSection"/> </c> satisfies a rule's content
    ''' predicates: an empty predicate set is always satisfied; otherwise the section must contain
    ''' at least one key whose KeyType equals a predicate's and whose Value matches the predicate's
    ''' (via <c> valueGlobMatches </c>), case-insensitively
    ''' </summary>
    '''
    ''' <param name="baseSection">
    ''' The candidate section already selected by the name suffix
    ''' </param>
    '''
    ''' <param name="matches">
    ''' The rule's content predicates
    ''' </param>
    '''
    Public Function sectionSatisfiesPredicates(baseSection As iniSection2,
                                                matches As List(Of iniKey2)) As Boolean

        If matches.Count = 0 Then Return True

        For Each baseKey In baseSection.Keys

            For Each matchKey In matches

                If baseKey.KeyType.Equals(matchKey.KeyType, StringComparison.OrdinalIgnoreCase) AndAlso
                   valueGlobMatches(matchKey.Value, baseKey.Value) Then Return True

            Next

        Next

        Return False

    End Function

    ''' <summary>
    ''' Matches <c> <paramref name="value"/> </c> against <c> <paramref name="pattern"/> </c>,
    ''' where <c> * </c> in the pattern is a wildcard matching any run of characters
    ''' (including none). A bare <c> * </c> matches anything, consistent with <c> [*Map:] </c>;
    ''' a pattern with no <c> * </c> is a case-insensitive equality check. All matching is
    ''' case-insensitive and the pattern is anchored end to end
    ''' </summary>
    '''
    ''' <param name="pattern">
    ''' The match pattern, potentially containing one or more <c> * </c> wildcards
    ''' </param>
    '''
    ''' <param name="value">
    ''' The base key value being tested
    ''' </param>
    '''
    Public Function valueGlobMatches(pattern As String, value As String) As Boolean

        If pattern = "*" Then Return True
        If Not pattern.Contains("*") Then Return value.Equals(pattern, StringComparison.OrdinalIgnoreCase)

        Dim regexPattern = "^" & Regex.Escape(pattern).Replace("\*", ".*") & "$"
        Return Regex.IsMatch(value, regexPattern, RegexOptions.IgnoreCase)

    End Function

    ''' <summary>
    ''' Parses a <c> [*Name: scaffold] </c> section into a <c> TransmuteNameRule </c>, splitting
    ''' <c> Match= </c> keys (content predicates, repeatable as <c> Match1= </c>, <c> Match2= </c>,
    ''' …) from the payload keys. Emits a warning and returns <c> Nothing </c> when the scaffold
    ''' label is empty, a <c> Match= </c> value is not a <c> Name=Value </c> pair, or no payload
    ''' keys were provided
    ''' </summary>
    '''
    ''' <param name="nameSection">
    ''' The source-file section to parse, whose name begins with <c> *Name: </c>
    ''' </param>
    '''
    ''' <param name="menuOutput">
    ''' The <c> MenuSection </c> containing output to be displayed to the user
    ''' </param>
    '''
    Private Function parseNameRule(nameSection As iniSection2,
                                   ByRef menuOutput As MenuSection) As TransmuteNameRule

        Dim scaffold = nameSection.Name.Substring(NameSectionPrefix.Length).Trim()

        If scaffold.Length = 0 Then

            warnMalformedRule(nameSection.Name, "no scaffold label was provided after [*Name:]", menuOutput)
            Return Nothing

        End If

        Dim rule As New TransmuteNameRule With {.Scaffold = scaffold}

        For Each key In nameSection.Keys

            If key.KeyType.Equals("Match", StringComparison.OrdinalIgnoreCase) Then

                Dim matchKey = splitNameValue(key.Value)

                If matchKey Is Nothing Then

                    warnMalformedRule(scaffold, $"{key.Name} value '{key.Value}' is not a Name=Value pair", menuOutput)
                    Return Nothing

                End If

                rule.Matches.Add(matchKey)

            Else

                rule.Payload.Add(key)

            End If

        Next

        If rule.Payload.Count = 0 Then

            warnMalformedRule(scaffold, "no payload keys were provided (only Match= predicates)", menuOutput)
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
    ''' The raw key line, ie. the value of a <c> Match= </c> key
    ''' </param>
    '''
    Public Function splitNameValue(line As String) As iniKey2

        Dim eqPos = line.IndexOf("="c)
        If eqPos < 1 Then Return Nothing

        Return New iniKey2(line)

    End Function

    ''' <summary>
    ''' Emits the warning for a malformed <c> [*Name:] </c> rule to both the menu and the log
    ''' </summary>
    '''
    ''' <param name="label">
    ''' The rule's scaffold label (or section name), for identification in the warning
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
    Public Sub warnMalformedRule(label As String,
                                  reason As String,
                                  ByRef menuOutput As MenuSection)

        Dim malformedMsg = $"[*Name: {label}] is malformed and will be skipped: {reason}"
        menuOutput.AddWarning(malformedMsg)
        gLog(malformedMsg)

    End Sub

End Module
