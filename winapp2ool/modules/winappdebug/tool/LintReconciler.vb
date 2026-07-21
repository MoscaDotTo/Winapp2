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

Imports System.IO

''' <summary>
''' Guards the generative modules' <see cref="WinappDebug.remotedebug"/> normalization pass
''' with a semantic reconciliation: the optimization pass is only permitted to make
''' formatting-level changes (key reordering and renumbering, FileKey pattern-list
''' alphabetization, and merging FileKeys that share a path and flag). Any change that
''' loses or rewrites cleaning content — a dropped entry, a discarded malformed key, a
''' rewritten value — indicates the generator emitted invalid data, which the linter then
''' silently destroyed. That is a correctness bug in the generator or its sources, not a
''' cleanup, so the gate reports every such difference, dumps the pre-optimization file
''' next to the output for forensics, and fails the build (nonzero exit code) in scripted
''' use. <br /><br />
'''
''' The comparison decomposes each entry into a set of <em>semantic units</em>:
''' each FileKey contributes one <c> path|pattern|flag </c> triple per semicolon-delimited
''' pattern (via <see cref="fileKeyParams2"/>), and every other key contributes its
''' number-stripped <c> KeyType=Value </c> pair. Units are compared case-insensitively as
''' sets, which makes the three sanctioned optimization classes invisible by construction:
''' reordering and renumbering do not change set membership, pattern sorting does not
''' change the triple set, and a merge of two same-path FileKeys produces exactly the
''' union of their triples. Everything else — in either direction — is reported.
''' Detection is default-deny: a future lint rule that changes semantics will fire this
''' gate until its behavior is explicitly accounted for. <br /><br />
'''
''' Reporting (but never detection) pairs a lost unit with a gained unit into a single
''' <c> old → new </c> rewrite finding when the pairing is unambiguous: FileKey triples
''' that agree on two of their three components, or any other key type with exactly one
''' lost and one gained unit. Anything that cannot be paired one-for-one stays reported
''' as separate lost/gained findings.
''' </summary>
Public Module LintReconciler

    ''' <summary>
    ''' Runs <see cref="WinappDebug.remotedebug"/> with full optimizations over
    ''' <paramref name="givenIni"/> and reconciles the result against the input.
    ''' When the pass made only formatting-level changes, behaves exactly like
    ''' <c> remotedebug(givenIni, True) </c>. When semantic differences are found,
    ''' writes the pre-optimization content to <c> &lt;name&gt;.prelint.ini </c>
    ''' alongside the output, reports every difference, and sets a nonzero process
    ''' exit code so scripted builds fail.
    ''' </summary>
    '''
    ''' <param name="givenIni">
    ''' The generated <c> iniFile2 </c> to normalize; its <c> Dir </c> and <c> Name </c>
    ''' determine where the forensic pre-lint dump is written on gate failure
    ''' </param>
    '''
    ''' <param name="callingModule">
    ''' The generative module's name, embedded in gate messages for source-of-warning
    ''' localisation
    ''' </param>
    '''
    ''' <param name="menuOutput">
    ''' The <c> MenuSection </c> receiving gate warnings for display
    ''' </param>
    '''
    ''' <returns>
    ''' The normalized <c> iniFile2 </c>, exactly as <see cref="WinappDebug.remotedebug"/>
    ''' returns it (the gate reports but does not roll back)
    ''' </returns>
    Public Function remotedebugGuarded(givenIni As iniFile2,
                                       callingModule As String,
                                       menuOutput As MenuSection) As iniFile2

        If givenIni Is Nothing Then argIsNull(NameOf(givenIni)) : Return Nothing

        ' Capture the pre-lint state up front: remotedebug wraps the given file's sections
        ' by reference, so both the forensic dump and the unit set must be taken before
        ' the pass runs
        Dim preText = givenIni.ToString()
        Dim preUnits = CollectSemanticUnits(givenIni)

        Dim linted = remotedebug(givenIni, True)

        Dim losses = FindSemanticLosses(preUnits, CollectSemanticUnits(linted))

        If losses.Count = 0 Then Return linted

        Using gLogScope($"{callingModule} lint reconciliation FAILED with {losses.Count} finding(s)")

            Dim headline = $"{callingModule}: the optimization pass semantically changed generated output"
            menuOutput.AddWarning(headline)

            For Each loss In losses

                gLog(loss)
                menuOutput.AddWarning(loss)

            Next

            Dim prelintName = Path.GetFileNameWithoutExtension(linted.Name) & ".prelint.ini"
            Dim prelintDump = iniFile2.Empty(linted.Dir, prelintName)

            Try

                prelintDump.OverwriteToFile(preText)
                Dim dumpMsg = $"Pre-optimization content preserved for review: {prelintDump.Path()}"
                gLog(dumpMsg)
                menuOutput.AddWarning(dumpMsg)

            Catch ex As IOException

                handleIOException(ex)

            End Try

            ' Nonzero exit fails the scripted build pipeline; interactive runs surface the
            ' warnings above and are otherwise unaffected
            Environment.ExitCode = 1

        End Using

        Return linted

    End Function

    ''' <summary>
    ''' Decomposes every entry of <paramref name="sourceFile"/> into its semantic units.
    ''' The outer dictionary is keyed by entry name; each inner dictionary maps a
    ''' case-normalized unit (the comparison key) to its display form (used in gate
    ''' messages).
    ''' </summary>
    '''
    ''' <param name="sourceFile">
    ''' The file whose entries will be decomposed
    ''' </param>
    '''
    ''' <returns>
    ''' The per-entry semantic unit map for <paramref name="sourceFile"/>
    ''' </returns>
    Public Function CollectSemanticUnits(sourceFile As iniFile2) As Dictionary(Of String, Dictionary(Of String, String))

        If sourceFile Is Nothing Then argIsNull(NameOf(sourceFile)) : Return Nothing

        Dim units As New Dictionary(Of String, Dictionary(Of String, String))(StringComparer.InvariantCultureIgnoreCase)

        For Each section In sourceFile

            Dim entryUnits As New Dictionary(Of String, String)(StringComparer.InvariantCultureIgnoreCase)

            For Each key In section.Keys

                For Each unit In UnitsForKey(key)

                    ' Exact duplicates collapse here by design: deduplication is a
                    ' sanctioned, semantics-preserving optimization
                    entryUnits(unit) = unit

                Next

            Next

            units(section.Name) = entryUnits

        Next

        Return units

    End Function

    ''' <summary>
    ''' Compares the pre- and post-optimization unit maps and describes every semantic
    ''' difference: entries removed outright, units present before the pass but not after
    ''' (lost content), and units present after but not before (invented or rewritten
    ''' content). A lost/gained pair that corresponds unambiguously is reported as a
    ''' single <c> old → new </c> rewrite finding; everything else stays reported as
    ''' separate lost and gained findings.
    ''' </summary>
    '''
    ''' <param name="preUnits">
    ''' The unit map captured before the optimization pass
    ''' </param>
    '''
    ''' <param name="postUnits">
    ''' The unit map captured after the optimization pass
    ''' </param>
    '''
    ''' <returns>
    ''' One human-readable finding per semantic difference; empty when the pass was
    ''' formatting-only
    ''' </returns>
    Public Function FindSemanticLosses(preUnits As Dictionary(Of String, Dictionary(Of String, String)),
                                       postUnits As Dictionary(Of String, Dictionary(Of String, String))) As List(Of String)

        If preUnits Is Nothing Then argIsNull(NameOf(preUnits)) : Return Nothing
        If postUnits Is Nothing Then argIsNull(NameOf(postUnits)) : Return Nothing

        Dim losses As New List(Of String)

        For Each entryName In preUnits.Keys

            If Not postUnits.ContainsKey(entryName) Then

                losses.Add($"[{entryName}] was removed entirely by the optimization pass")
                Continue For

            End If

            losses.AddRange(DescribeEntryDifferences(entryName, preUnits(entryName), postUnits(entryName)))

        Next

        For Each entryName In postUnits.Keys

            If Not preUnits.ContainsKey(entryName) Then losses.Add($"[{entryName}] was introduced by the optimization pass")

        Next

        Return losses

    End Function

    ''' <summary>
    ''' Describes the semantic differences within a single entry. The lost and gained
    ''' unit sets are computed exactly as before (detection is unchanged), then a
    ''' reporting-only pairing pass promotes unambiguous lost/gained pairs into single
    ''' rewrite findings; whatever remains unpaired is reported as lost or gained.
    ''' </summary>
    '''
    ''' <param name="entryName">
    ''' The entry's name, embedded in each finding
    ''' </param>
    '''
    ''' <param name="pre">
    ''' The entry's semantic units before the optimization pass
    ''' </param>
    '''
    ''' <param name="post">
    ''' The entry's semantic units after the optimization pass
    ''' </param>
    '''
    ''' <returns>
    ''' One finding per rewrite, loss, or gain; empty when the entry's units are identical
    ''' </returns>
    Private Function DescribeEntryDifferences(entryName As String,
                                              pre As Dictionary(Of String, String),
                                              post As Dictionary(Of String, String)) As List(Of String)

        Dim lost As New List(Of String)
        Dim gained As New List(Of String)

        For Each unit In pre.Keys

            If Not post.ContainsKey(unit) Then lost.Add(pre(unit))

        Next

        For Each unit In post.Keys

            If Not pre.ContainsKey(unit) Then gained.Add(post(unit))

        Next

        Dim findings As New List(Of String)

        For Each rewrite In PairRewrites(lost, gained)

            findings.Add($"[{entryName}] content rewritten during optimization: {rewrite.Key} → {rewrite.Value}")

        Next

        For Each unit In lost

            findings.Add($"[{entryName}] lost content during optimization: {unit}")

        Next

        For Each unit In gained

            findings.Add($"[{entryName}] gained unexpected content during optimization: {unit}")

        Next

        Return findings

    End Function

    ''' <summary>
    ''' Pairs lost units with gained units where the correspondence is unambiguous,
    ''' removing paired units from <paramref name="lost"/> and <paramref name="gained"/>
    ''' in place. FileKey triples pair when exactly one lost and one gained triple agree
    ''' on two of their three components — the disagreeing third is the rewrite — with
    ''' the most specific passes running first so a flag change is not misread as a
    ''' pattern change. Every other unit pairs when its KeyType carries exactly one lost
    ''' and one gained instance. Ambiguous groups (more than one candidate on either
    ''' side) are left untouched.
    ''' </summary>
    '''
    ''' <param name="lost">
    ''' Display-form units present before the pass but not after; paired units are removed
    ''' </param>
    '''
    ''' <param name="gained">
    ''' Display-form units present after the pass but not before; paired units are removed
    ''' </param>
    '''
    ''' <returns>
    ''' Pairs of (old unit, new unit) for each rewrite found; empty if none qualify
    ''' </returns>
    Private Function PairRewrites(lost As List(Of String),
                                  gained As List(Of String)) As List(Of KeyValuePair(Of String, String))

        Dim pairs As New List(Of KeyValuePair(Of String, String))

        PairBy(lost, gained, pairs, Function(u) TripleBucket(u, keepPath:=True, keepPattern:=True, keepFlag:=False))
        PairBy(lost, gained, pairs, Function(u) TripleBucket(u, keepPath:=True, keepPattern:=False, keepFlag:=True))
        PairBy(lost, gained, pairs, Function(u) TripleBucket(u, keepPath:=False, keepPattern:=True, keepFlag:=True))

        PairBy(lost, gained, pairs, AddressOf KeyTypeBucket)

        Return pairs

    End Function

    ''' <summary>
    ''' Groups <paramref name="lost"/> and <paramref name="gained"/> by
    ''' <paramref name="bucketFor"/> and records a (old, new) pair for every bucket
    ''' holding exactly one unit on each side, removing both from their lists.
    ''' Units for which <paramref name="bucketFor"/> returns <c> Nothing </c> do not
    ''' participate in this pass.
    ''' </summary>
    '''
    ''' <param name="lost">
    ''' The still-unpaired lost units; matched units are removed in place
    ''' </param>
    '''
    ''' <param name="gained">
    ''' The still-unpaired gained units; matched units are removed in place
    ''' </param>
    '''
    ''' <param name="pairs">
    ''' The accumulator matched (old, new) pairs are appended to
    ''' </param>
    '''
    ''' <param name="bucketFor">
    ''' Maps a unit to its comparison bucket, or <c> Nothing </c> to exclude it
    ''' </param>
    Private Sub PairBy(lost As List(Of String),
                       gained As List(Of String),
                       pairs As List(Of KeyValuePair(Of String, String)),
                       bucketFor As Func(Of String, String))

        Dim lostBuckets = GroupByBucket(lost, bucketFor)
        Dim gainedBuckets = GroupByBucket(gained, bucketFor)

        For Each bucket In lostBuckets.Keys

            If lostBuckets(bucket).Count <> 1 Then Continue For

            Dim gainedForBucket As List(Of String) = Nothing
            If Not gainedBuckets.TryGetValue(bucket, gainedForBucket) Then Continue For
            If gainedForBucket.Count <> 1 Then Continue For

            pairs.Add(New KeyValuePair(Of String, String)(lostBuckets(bucket)(0), gainedForBucket(0)))
            lost.Remove(lostBuckets(bucket)(0))
            gained.Remove(gainedForBucket(0))

        Next

    End Sub

    ''' <summary>
    ''' Groups units by their bucket key, skipping units whose bucket is <c> Nothing </c>
    ''' </summary>
    '''
    ''' <param name="units">
    ''' The units to group
    ''' </param>
    '''
    ''' <param name="bucketFor">
    ''' Maps a unit to its comparison bucket, or <c> Nothing </c> to exclude it
    ''' </param>
    '''
    ''' <returns>
    ''' A case-insensitive map of bucket key to the units sharing it
    ''' </returns>
    Private Function GroupByBucket(units As List(Of String),
                                   bucketFor As Func(Of String, String)) As Dictionary(Of String, List(Of String))

        Dim buckets As New Dictionary(Of String, List(Of String))(StringComparer.InvariantCultureIgnoreCase)

        For Each unit In units

            Dim bucket = bucketFor(unit)
            If bucket Is Nothing Then Continue For

            If Not buckets.ContainsKey(bucket) Then buckets.Add(bucket, New List(Of String))
            buckets(bucket).Add(unit)

        Next

        Return buckets

    End Function

    ''' <summary>
    ''' Builds a comparison bucket from a FileKey triple unit by keeping the selected
    ''' components and masking the rest, so units disagreeing only on masked components
    ''' land in the same bucket. Returns <c> Nothing </c> for non-triple units (including
    ''' the malformed-FileKey fallback form, which carries an <c> = </c> instead of a
    ''' space and is bucketed by KeyType like any other unit).
    ''' </summary>
    '''
    ''' <param name="unit">
    ''' The display-form unit to bucket
    ''' </param>
    '''
    ''' <param name="keepPath">
    ''' Whether the path component participates in the bucket
    ''' </param>
    '''
    ''' <param name="keepPattern">
    ''' Whether the pattern component participates in the bucket
    ''' </param>
    '''
    ''' <param name="keepFlag">
    ''' Whether the flag component participates in the bucket
    ''' </param>
    '''
    ''' <returns>
    ''' The bucket key, or <c> Nothing </c> if <paramref name="unit"/> is not a FileKey triple
    ''' </returns>
    Private Function TripleBucket(unit As String,
                                  keepPath As Boolean,
                                  keepPattern As Boolean,
                                  keepFlag As Boolean) As String

        If Not unit.StartsWith("FileKey ", StringComparison.InvariantCultureIgnoreCase) Then Return Nothing

        Dim parts = unit.Substring("FileKey ".Length).Split("|"c)
        If parts.Length <> 3 Then Return Nothing

        ' Pipes cannot appear inside path, pattern, or flag, and the mask position is
        ' constant within a pass, so this composite key cannot collide across units
        Return $"{If(keepPath, parts(0), "*")}|{If(keepPattern, parts(1), "*")}|{If(keepFlag, parts(2), "*")}"

    End Function

    ''' <summary>
    ''' Builds a comparison bucket from a <c> KeyType=Value </c> unit: its KeyType.
    ''' Returns <c> Nothing </c> for FileKey triple units, which are paired component-wise
    ''' by <see cref="TripleBucket"/> instead.
    ''' </summary>
    '''
    ''' <param name="unit">
    ''' The display-form unit to bucket
    ''' </param>
    '''
    ''' <returns>
    ''' The unit's KeyType, or <c> Nothing </c> if the unit has no <c> = </c> separator
    ''' or is a FileKey triple
    ''' </returns>
    Private Function KeyTypeBucket(unit As String) As String

        If unit.StartsWith("FileKey ", StringComparison.InvariantCultureIgnoreCase) Then Return Nothing

        Dim eqIdx = unit.IndexOf("="c)
        Return If(eqIdx > 0, unit.Substring(0, eqIdx), Nothing)

    End Function

    ''' <summary>
    ''' Produces the semantic units for one key. A FileKey yields one
    ''' <c> path|pattern|flag </c> triple per pattern so that pattern reordering and
    ''' same-path merges are invisible to set comparison; a FileKey whose value parses to
    ''' zero patterns falls back to the whole raw value so malformed keys cannot vanish
    ''' silently. Every other key type yields its number-stripped
    ''' <c> KeyType=Value </c> pair.
    ''' </summary>
    '''
    ''' <param name="key">
    ''' The key to decompose
    ''' </param>
    '''
    ''' <returns>
    ''' The display-form unit strings for <paramref name="key"/> (comparison uses a
    ''' case-insensitive dictionary, so no separate normalization is required)
    ''' </returns>
    Private Function UnitsForKey(key As iniKey2) As List(Of String)

        Dim units As New List(Of String)

        If Not key.KeyType.Equals("FileKey", StringComparison.InvariantCultureIgnoreCase) Then

            units.Add($"{key.KeyType}={key.Value}")
            Return units

        End If

        Dim params As New fileKeyParams2(key.Value)

        If params.Patterns.Count = 0 Then

            units.Add($"FileKey={key.Value}")
            Return units

        End If

        ' RawFlag holds only unrecognized flag text — RECURSE/REMOVESELF land in the Flag
        ' enum with an empty RawFlag, so the unit must encode the effective flag or the
        ' comparison goes blind to recursion-scope changes
        Dim flagText = EffectiveFlagText(params)

        For Each pattern In params.Patterns

            units.Add($"FileKey {params.Path}|{pattern}|{flagText}")

        Next

        Return units

    End Function

    ''' <summary>
    ''' Produces the flag component of a FileKey unit: the canonical flag name for
    ''' recognized flags, the verbatim raw text for unrecognized ones, and an empty
    ''' string when no flag is present
    ''' </summary>
    '''
    ''' <param name="params">
    ''' The parsed FileKey whose flag is rendered
    ''' </param>
    '''
    ''' <returns>
    ''' The flag text used in the unit's third component
    ''' </returns>
    Private Function EffectiveFlagText(params As fileKeyParams2) As String

        Select Case params.Flag

            Case fileKeyFlag.None : Return ""

            Case fileKeyFlag.Unknown : Return params.RawFlag

            Case Else : Return params.Flag.ToString().ToUpperInvariant()

        End Select

    End Function

End Module
