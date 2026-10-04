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
''' Guards the generative modules' <see cref="WinappDebug.remotedebug"/> normalization pass.
''' The optimization pass is only allowed to change formatting: reordering and renumbering
''' keys, alphabetizing a FileKey's pattern list, and merging FileKeys that share a path and
''' flag. Anything that loses or rewrites real cleaning content, whether that's a dropped
''' entry, a discarded malformed key, or a rewritten value, means the generator emitted
''' invalid data and the linter quietly ate it. That's a bug in the generator or its sources
''' rather than a cleanup, so the gate reports every difference it finds, dumps the
''' pre-optimization file beside the output so it can be picked apart later, and fails
''' scripted builds with a nonzero exit code. <br /><br />
'''
''' The comparison breaks each entry down into <em>semantic units</em>. Every FileKey gives
''' one <c> path|pattern|flag </c> triple per semicolon-delimited pattern (via
''' <see cref="fileKeyParams"/>), and every other key gives its number-stripped
''' <c> KeyType=Value </c> pair. Units are compared case-insensitively as sets, which is why
''' the three sanctioned optimizations don't show up here: reordering and renumbering leave
''' set membership alone, sorting patterns leaves the triple set alone, and merging two
''' same-path FileKeys gives exactly the union of their triples. Anything else gets reported,
''' in either direction. The gate assumes the worst, so a future lint rule that changes
''' semantics will keep tripping it until somebody sits down and works out what it's
''' actually doing to the output. <br /><br />
'''
''' When reporting, though never when detecting, a lost unit and a gained unit are folded
''' into one <c> old → new </c> rewrite finding if the pairing is unambiguous: FileKey triples
''' agreeing on two of their three components, or any other key type carrying exactly one
''' lost and one gained unit. Whatever can't be paired one for one stays reported separately.
''' </summary>
Public Module LintReconciler

    ''' <summary>
    ''' Runs <see cref="WinappDebug.remotedebug"/> with full optimizations over
    ''' <paramref name="givenIni"/> and reconciles the result against what went in.
    ''' If the pass only changed formatting, this does exactly what
    ''' <c> remotedebug(givenIni, True) </c> does. If it changed anything semantic, the
    ''' pre-optimization content goes to <c> &lt;name&gt;.prelint.ini </c> beside the output,
    ''' every difference gets reported, and the process exit code goes nonzero so a scripted
    ''' build fails.
    ''' </summary>
    '''
    ''' <param name="givenIni">
    ''' The generated <c> iniFile </c> to normalize. Its <c> Dir </c> and <c> Name </c> decide
    ''' where the pre-lint dump lands when the gate fails
    ''' </param>
    '''
    ''' <param name="callingModule">
    ''' The generative module's name, which goes into the gate messages so you can tell
    ''' where a warning came from
    ''' </param>
    '''
    ''' <param name="menuOutput">
    ''' The <c> MenuSection </c> receiving gate warnings for display
    ''' </param>
    '''
    ''' <returns>
    ''' The normalized <c> iniFile </c>, exactly as <see cref="WinappDebug.remotedebug"/>
    ''' returns it. The gate reports, it doesn't roll anything back
    ''' </returns>
    Public Function remotedebugGuarded(givenIni As iniFile,
                                       callingModule As String,
                                       menuOutput As MenuSection) As iniFile

        If givenIni Is Nothing Then argIsNull(NameOf(givenIni)) : Return Nothing

        ' Grab the pre-lint state first. remotedebug wraps the given file's sections by
        ' reference, so the dump and the unit set both have to come off it before the pass runs
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
            Dim prelintDump = iniFile.Empty(linted.Dir, prelintName)

            Try

                prelintDump.OverwriteToFile(preText)
                Dim dumpMsg = $"Pre-optimization content preserved for review: {prelintDump.Path()}"
                gLog(dumpMsg)
                menuOutput.AddWarning(dumpMsg)

            Catch ex As IOException

                handleIOException(ex)

            End Try

            ' A nonzero exit fails the scripted build. Interactive runs just show the
            ' warnings above and carry on as normal
            Environment.ExitCode = 1

        End Using

        Return linted

    End Function

    ''' <summary>
    ''' Breaks every entry of <paramref name="sourceFile"/> down into its semantic units.
    ''' The outer dictionary is keyed by entry name. Each inner one maps a case-normalized
    ''' unit to its display form. We compare on the former and print the latter in the gate
    ''' messages.
    ''' </summary>
    '''
    ''' <param name="sourceFile">
    ''' The file whose entries will be decomposed
    ''' </param>
    '''
    ''' <returns>
    ''' The per-entry semantic unit map for <paramref name="sourceFile"/>
    ''' </returns>
    Public Function CollectSemanticUnits(sourceFile As iniFile) As Dictionary(Of String, Dictionary(Of String, String))

        If sourceFile Is Nothing Then argIsNull(NameOf(sourceFile)) : Return Nothing

        Dim units As New Dictionary(Of String, Dictionary(Of String, String))(StringComparer.InvariantCultureIgnoreCase)

        For Each section In sourceFile

            Dim entryUnits As New Dictionary(Of String, String)(StringComparer.InvariantCultureIgnoreCase)

            For Each key In section.Keys

                For Each unit In UnitsForKey(key)

                    ' Exact duplicates collapse here on purpose, since throwing them out
                    ' is one of the optimizations we allow
                    entryUnits(unit) = unit

                Next

            Next

            units(section.Name) = entryUnits

        Next

        Return units

    End Function

    ''' <summary>
    ''' Compares the pre- and post-optimization unit maps and describes every semantic
    ''' difference it finds: entries that were removed outright, units that were there
    ''' before the pass but not after, and units that turned up afterwards without having
    ''' been there to start with. Where a lost and a gained unit clearly correspond, they're
    ''' reported as one <c> old → new </c> rewrite. Everything else is reported as a plain
    ''' loss or gain.
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
    ''' Describes the semantic differences inside a single entry. The lost and gained sets
    ''' get worked out first, then a pairing pass folds the obvious lost/gained pairs into
    ''' single rewrite findings. That pairing only affects how things are reported, never
    ''' what gets detected. Anything left unpaired is reported as a plain loss or gain.
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
    ''' Pairs up lost and gained units where the correspondence is obvious, dropping whatever
    ''' it pairs from <paramref name="lost"/> and <paramref name="gained"/> in place. Two
    ''' FileKey triples pair when exactly one lost and one gained triple agree on two of their
    ''' three components. The third one, the one they disagree on, is the rewrite. The most
    ''' specific passes run first so that a flag change doesn't get read as a pattern change.
    ''' Every other unit pairs when its KeyType has exactly one lost and one gained instance.
    ''' If there's more than one candidate on either side we leave the whole group alone.
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
    ''' <paramref name="bucketFor"/> and records an (old, new) pair for every bucket holding
    ''' exactly one unit on each side, pulling both out of their lists as it goes. A unit
    ''' whose <paramref name="bucketFor"/> comes back <c> Nothing </c> sits this pass out.
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
    ''' Builds a comparison bucket out of a FileKey triple, keeping the components asked for
    ''' and masking the rest, so two units that only disagree on a masked component land in
    ''' the same bucket. Anything that isn't a triple gets <c> Nothing </c> back. That includes
    ''' the malformed-FileKey fallback form, which carries an <c> = </c> rather than a space
    ''' and is bucketed by KeyType like everything else.
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

        ' A pipe can't appear inside a path, pattern, or flag, and the mask position doesn't
        ' move within a pass, so this composite key can't collide across units
        Return $"{If(keepPath, parts(0), "*")}|{If(keepPattern, parts(1), "*")}|{If(keepFlag, parts(2), "*")}"

    End Function

    ''' <summary>
    ''' The bucket for a <c> KeyType=Value </c> unit is just its KeyType. FileKey triples get
    ''' <c> Nothing </c> back, since <see cref="TripleBucket"/> pairs those component by
    ''' component instead.
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
    ''' Produces the semantic units for one key. A FileKey gives one
    ''' <c> path|pattern|flag </c> triple per pattern, so that reordering patterns or merging
    ''' two same-path keys never shows up in the comparison. If a FileKey's value parses to no
    ''' patterns at all we fall back to the whole raw value, so a malformed key can't quietly
    ''' disappear on us. Every other key type gives its number-stripped
    ''' <c> KeyType=Value </c> pair.
    ''' </summary>
    '''
    ''' <param name="key">
    ''' The key to decompose
    ''' </param>
    '''
    ''' <returns>
    ''' The display-form unit strings for <paramref name="key"/>. The comparison runs through
    ''' a case-insensitive dictionary, so we don't normalize them separately
    ''' </returns>
    Private Function UnitsForKey(key As iniKey) As List(Of String)

        Dim units As New List(Of String)

        If Not key.KeyType.Equals("FileKey", StringComparison.InvariantCultureIgnoreCase) Then

            units.Add($"{key.KeyType}={key.Value}")
            Return units

        End If

        Dim params As New fileKeyParams(key.Value)

        If params.Patterns.Count = 0 Then

            units.Add($"FileKey={key.Value}")
            Return units

        End If

        ' RawFlag only holds flag text we didn't recognize. RECURSE and REMOVESELF land in
        ' the Flag enum and leave RawFlag empty, so the unit has to carry the effective flag
        ' or the comparison goes blind to changes in recursion scope
        Dim flagText = EffectiveFlagText(params)

        For Each pattern In params.Patterns

            units.Add($"FileKey {params.Path}|{pattern}|{flagText}")

        Next

        Return units

    End Function

    ''' <summary>
    ''' Produces the flag component of a FileKey unit. A recognized flag gets its canonical
    ''' name, an unrecognized one gets its raw text verbatim, and a key with no flag at all
    ''' gets an empty string
    ''' </summary>
    '''
    ''' <param name="params">
    ''' The parsed FileKey whose flag is rendered
    ''' </param>
    '''
    ''' <returns>
    ''' The flag text used in the unit's third component
    ''' </returns>
    Private Function EffectiveFlagText(params As fileKeyParams) As String

        Select Case params.Flag

            Case fileKeyFlag.None : Return ""

            Case fileKeyFlag.Unknown : Return params.RawFlag

            Case Else : Return params.Flag.ToString().ToUpperInvariant()

        End Select

    End Function

End Module
