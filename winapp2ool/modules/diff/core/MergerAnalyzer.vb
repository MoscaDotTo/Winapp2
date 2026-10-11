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
''' Re-diffs every entry that received merged content against the combined keys of its sources,
''' and works out which source keys each one needn't list because another entry kept them or a
''' split source's own section lists them. Its results replace the merge targets' tracker
''' contents and are read by <see cref="DiffOutputRenderer"/> when it builds the merger blocks.
''' </summary>
Public Class MergerAnalyzer

    Private ReadOnly _state As DiffState
    Private ReadOnly _file1 As iniFile
    Private ReadOnly _file2 As iniFile
    Private ReadOnly _keyAnalyzer As KeyModificationAnalyzer

    ''' <summary>
    ''' Maps merge target name → (key value → its source old entries, as
    ''' <see cref="DiffOutputRenderer.CreditSources"/> formats them). Built by <c> Analyze </c>
    ''' for every modified and added merge target.
    ''' </summary>
    Private ReadOnly _mergerSourceMaps As New Dictionary(Of String, Dictionary(Of String, String))(StringComparer.OrdinalIgnoreCase)

    ''' <summary>
    ''' Maps merged source entry name → (key type and value → the new entries it went into that
    ''' have the key, unchanged or captured). Built at the start of <c> Analyze </c>.
    ''' </summary>
    Private _keptBy As Dictionary(Of String, Dictionary(Of String, HashSet(Of String)))

    ''' <summary>
    ''' Maps merge target name → the source keys it doesn't have that another of their targets
    ''' kept, and those targets. Set by <c> StripKeysKeptElsewhere </c>.
    ''' </summary>
    Private ReadOnly _keptElsewhere As New Dictionary(Of String, KeptElsewhereInfo)(StringComparer.OrdinalIgnoreCase)

    ''' <summary>
    ''' Maps removed entry name → its keys that no entry it was split into kept. Filled for every
    ''' merged source at the start of <c> Analyze </c>.
    ''' </summary>
    Private ReadOnly _unkeptKeys As New Dictionary(Of String, List(Of iniKey))(StringComparer.OrdinalIgnoreCase)

    Private _modifiedTargets As New List(Of String)
    Private _addedTargets As New List(Of String)

    ''' <summary>
    ''' Creates a new <c> MergerAnalyzer </c>
    ''' </summary>
    '''
    ''' <param name="state">
    ''' Shared diff state tracking all entry changes
    ''' </param>
    '''
    ''' <param name="file1">
    ''' The old version of winapp2.ini as an <c> iniFile </c>
    ''' </param>
    '''
    ''' <param name="file2">
    ''' The new version of winapp2.ini as an <c> iniFile </c>
    ''' </param>
    '''
    ''' <param name="keyAnalyzer">
    ''' Used to re-diff merge targets against their combined source keys
    ''' </param>
    Public Sub New(state As DiffState,
                   file1 As iniFile,
                   file2 As iniFile,
                   keyAnalyzer As KeyModificationAnalyzer)

        _state = state
        _file1 = file1
        _file2 = file2
        _keyAnalyzer = keyAnalyzer

    End Sub

    ''' <summary>
    ''' The entries present in both files that received merged content, excluding renames,
    ''' sorted by name ignoring case. Set by <see cref="Analyze"/>.
    ''' </summary>
    Public ReadOnly Property ModifiedTargets As IReadOnlyList(Of String)
        Get
            Return _modifiedTargets
        End Get
    End Property

    ''' <summary>
    ''' The added entries that contain merged content, excluding renames, sorted by name
    ''' ignoring case. Set by <see cref="Analyze"/>.
    ''' </summary>
    Public ReadOnly Property AddedTargets As IReadOnlyList(Of String)
        Get
            Return _addedTargets
        End Get
    End Property

    ''' <summary>
    ''' Re-diffs each modified merge target against the combined keys of its own old version and
    ''' its sources, then each added merge target against its sources' keys, and strips from
    ''' each the source keys it needn't list. A modified target's re-diff lists the keys it gave
    ''' away as removed again, so we take those back out. For an added target, we then add the
    ''' keys carried over unchanged from a source to its added-key tracker list, creating it if
    ''' needed, which <see cref="DiffStatisticsCalculator.CalculateAddedWithMergersStatistics"/>
    ''' relies on. Run it after <see cref="DiffStatisticsCalculator.DetectCrossEntryMovements"/> and
    ''' before <see cref="ModifiedEntryBlock.BuildAll"/>, which reads the re-diffs' results.
    ''' </summary>
    Public Sub Analyze()

        _keptBy = BuildKeptBy()

        For Each source In _state.MergedEntries.OldToNewMergeDict.Keys
            _unkeptKeys(source) = FindUnkeptKeys(source)
        Next

        _modifiedTargets = _state.MergedEntries.MergeDict.Keys _
            .OrderBy(Function(k) k, StringComparer.OrdinalIgnoreCase) _
            .Where(Function(targetEntry) Not _state.ModifiedEntries.AddedEntryNames.Contains(targetEntry) AndAlso
                                          Not _state.MergedEntries.RenamedEntryNames.Contains(targetEntry)) _
            .ToList()

        For Each targetEntry In _modifiedTargets

            Dim combined = BuildCombinedOldKeys(_state.MergedEntries.MergeDict(targetEntry), targetEntry)
            _mergerSourceMaps(targetEntry) = combined.SourceEntryMap
            _keyAnalyzer.FindModificationsFromCombinedKeys(combined.Keys, _file2.GetSection(targetEntry))
            StripKeysKeptElsewhere(targetEntry)

            ' The re-diff starts over, so it lists the keys this entry gave away as removed again
            Dim movedOut = _state.KeyMovements.MovedOutOf(targetEntry)
            Dim removed As List(Of iniKey) = Nothing
            If movedOut.Count = 0 OrElse Not _state.ModifiedEntries.RemovedKeyTracker.TryGetValue(targetEntry, removed) Then Continue For

            removed.RemoveAll(Function(key) movedOut.Any(Function(m) String.Equals(m.MovedKey.KeyType, key.KeyType, StringComparison.OrdinalIgnoreCase) AndAlso
                                                                     String.Equals(m.MovedKey.Value, key.Value, StringComparison.OrdinalIgnoreCase)))

        Next

        _addedTargets = _state.ModifiedEntries.AddedEntryNames _
            .OrderBy(Function(s) s, StringComparer.OrdinalIgnoreCase) _
            .Where(Function(entry) Not _state.MergedEntries.RenamedEntryNames.Contains(entry) AndAlso
                                    _state.MergedEntries.MergeDict.ContainsKey(entry)) _
            .ToList()

        For Each entry In _addedTargets

            Dim combined = BuildCombinedOldKeys(_state.MergedEntries.MergeDict(entry), entry)
            _mergerSourceMaps(entry) = combined.SourceEntryMap

            _keyAnalyzer.FindModificationsForAddedEntryFromKeys(combined.Keys, _file2.GetSection(entry))
            StripKeysKeptElsewhere(entry)

            RecordCarriedOverKeys(entry, combined.SourceEntryMap)

        Next

    End Sub

    ''' <summary>
    ''' Appends to an added merge target's added keys each of its keys that came unchanged from
    ''' a source: one its re-diff neither added, removed nor updated, whose value a source has.
    ''' Stores the list as the entry's added-key tracker list when it isn't empty.
    ''' </summary>
    '''
    ''' <param name="entry">
    ''' The name of the added entry, after its re-diff
    ''' </param>
    '''
    ''' <param name="sourceEntryMap">
    ''' The entry's map of key value → its source entries
    ''' </param>
    Private Sub RecordCarriedOverKeys(entry As String, sourceEntryMap As Dictionary(Of String, String))

        Dim addedKeys = If(_state.ModifiedEntries.AddedKeyTracker.ContainsKey(entry),
                           _state.ModifiedEntries.AddedKeyTracker(entry), New List(Of iniKey))

        Dim removedKeys = If(_state.ModifiedEntries.RemovedKeyTracker.ContainsKey(entry),
                             _state.ModifiedEntries.RemovedKeyTracker(entry), New List(Of iniKey))

        Dim updatedKeysDict = If(_state.ModifiedEntries.ModifiedKeyTracker.ContainsKey(entry),
                                 _state.ModifiedEntries.ModifiedKeyTracker(entry),
                                 New Dictionary(Of iniKey, List(Of iniKey)))

        Dim addedKeyIds As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        For Each k In addedKeys : addedKeyIds.Add($"{k.KeyType}|{k.Value}") : Next

        Dim removedKeyIds As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        For Each k In removedKeys : removedKeyIds.Add($"{k.KeyType}|{k.Value}") : Next

        Dim carriedOverKeys As New List(Of iniKey)
        For Each key In _file2.GetSection(entry).Keys

            Dim keyId = $"{key.KeyType}|{key.Value}"
            If addedKeyIds.Contains(keyId) OrElse removedKeyIds.Contains(keyId) Then Continue For

            Dim isUpdated = False
            For Each kvp In updatedKeysDict

                If kvp.Key.Value = key.Value Then isUpdated = True : Exit For

            Next

            If isUpdated Then Continue For

            If sourceEntryMap.ContainsKey(key.Value) Then carriedOverKeys.Add(key)

        Next

        addedKeys.AddRange(carriedOverKeys)
        If addedKeys.Count > 0 Then _state.ModifiedEntries.AddedKeyTracker(entry) = addedKeys

    End Sub

    ''' <summary>
    ''' Returns <paramref name="target"/>'s map of key value → its source entries, as
    ''' <see cref="DiffOutputRenderer.CreditSources"/> formats them
    ''' </summary>
    '''
    ''' <param name="target">
    ''' The name of a merge target
    ''' </param>
    '''
    ''' <returns>
    ''' The map, or <c> Nothing </c> if <see cref="Analyze"/> didn't re-diff
    ''' <paramref name="target"/>
    ''' </returns>
    Public Function SourceMap(target As String) As Dictionary(Of String, String)

        Dim map As Dictionary(Of String, String) = Nothing
        _mergerSourceMaps.TryGetValue(target, map)
        Return map

    End Function

    ''' <summary>
    ''' Returns the source keys <paramref name="target"/> doesn't have and needn't list, as
    ''' <c> StripKeysKeptElsewhere </c> found them
    ''' </summary>
    '''
    ''' <param name="target">
    ''' The name of a merge target
    ''' </param>
    '''
    ''' <returns>
    ''' The record, or <c> Nothing </c> if no such key was stripped from <paramref name="target"/>
    ''' </returns>
    Friend Function KeptElsewhere(target As String) As KeptElsewhereInfo

        Dim info As KeptElsewhereInfo = Nothing
        _keptElsewhere.TryGetValue(target, info)
        Return info

    End Function

    ''' <summary>
    ''' Returns the keys of a removed entry split across several entries that none of those
    ''' entries has, unchanged or captured
    ''' </summary>
    '''
    ''' <param name="source">
    ''' The name of a removed entry that was merged into other entries
    ''' </param>
    '''
    ''' <returns>
    ''' The keys no entry kept, in the old entry's order, or an empty list when
    ''' <paramref name="source"/> went into only one entry, whose merger block lists them instead
    ''' </returns>
    Public Function UnkeptKeys(source As String) As List(Of iniKey)

        Dim unkept As List(Of iniKey) = Nothing
        If _unkeptKeys.TryGetValue(source, unkept) Then Return unkept

        Return New List(Of iniKey)

    End Function

    ''' <summary>
    ''' Computes what <see cref="UnkeptKeys"/> returns for <paramref name="source"/>
    ''' </summary>
    '''
    ''' <param name="source">
    ''' The name of a removed entry that was merged into other entries
    ''' </param>
    '''
    ''' <returns>
    ''' The keys no entry kept, in the old entry's order, or an empty list when
    ''' <paramref name="source"/> went into only one entry
    ''' </returns>
    Private Function FindUnkeptKeys(source As String) As List(Of iniKey)

        Dim unkept As New List(Of iniKey)
        Dim targets As List(Of String) = Nothing
        Dim section = _file1.GetSection(source)

        If section IsNot Nothing AndAlso _state.MergedEntries.OldToNewMergeDict.TryGetValue(source, targets) AndAlso targets.Count > 1 Then

            Dim kept As Dictionary(Of String, HashSet(Of String)) = Nothing
            _keptBy.TryGetValue(source, kept)
            unkept.AddRange(section.Keys.Where(Function(k) kept Is Nothing OrElse Not kept.ContainsKey(KeyId(k))))

        End If

        Return unkept

    End Function

    ''' <summary>
    ''' Removes from <paramref name="target"/>'s removed keys each merged source key that it
    ''' needn't list, and records what it removed for <see cref="KeptElsewhere"/>: keys another
    ''' entry the source went into kept, and keys of a split source that no entry kept, which the
    ''' removed entry's own section lists instead. Without this, each piece of a split entry
    ''' listed every one of its sibling's keys, and every key the split lost, as dropped. The
    ''' target's own old keys are left alone.
    ''' </summary>
    '''
    ''' <param name="target">
    ''' The name of an entry that received merged content, after its re-diff
    ''' </param>
    Private Sub StripKeysKeptElsewhere(target As String)

        Dim removed As List(Of iniKey) = Nothing
        If Not _state.ModifiedEntries.RemovedKeyTracker.TryGetValue(target, removed) Then Return

        Dim ownOldKeys = If(_file1.Contains(target), _file1.GetSection(target).Keys.Select(Function(own) KeyId(own)).ToList(), New List(Of String))
        Dim keptBy As New SortedSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Dim listedWith As New SortedSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Dim keptCount = 0
        Dim unkeptCount = 0

        removed.RemoveAll(Function(key)

                              If ownOldKeys.Contains(KeyId(key)) Then Return False

                              Dim others = TargetsKeeping(target, key)
                              If others.Count > 0 Then

                                  keptBy.UnionWith(others)
                                  keptCount += 1
                                  Return True

                              End If

                              Dim splits = SplitSourcesListing(target, key)
                              If splits.Count = 0 Then Return False

                              listedWith.UnionWith(splits)
                              unkeptCount += 1
                              Return True

                          End Function)

        If keptCount + unkeptCount = 0 Then Return

        _keptElsewhere(target) = New KeptElsewhereInfo(keptCount, keptBy.ToList(), unkeptCount, listedWith.ToList())

    End Sub

    ''' <summary>
    ''' Finds the split sources of <paramref name="target"/> whose own section lists
    ''' <paramref name="key"/> as kept by no entry
    ''' </summary>
    '''
    ''' <param name="target">
    ''' The name of the entry the sources were merged into
    ''' </param>
    '''
    ''' <param name="key">
    ''' A source key <paramref name="target"/> doesn't have
    ''' </param>
    '''
    ''' <returns>
    ''' The sources that have the key, or an empty list unless every one of them was split and
    ''' lists it, since a source merged into <paramref name="target"/> alone has no other place
    ''' to show it
    ''' </returns>
    Private Function SplitSourcesListing(target As String, key As iniKey) As List(Of String)

        Dim id = KeyId(key)
        Dim listing As New List(Of String)

        For Each source In _state.MergedEntries.MergeDict(target)

            Dim section = _file1.GetSection(source)
            If section Is Nothing OrElse Not section.Keys.Any(Function(k) KeyId(k) = id) Then Continue For
            If Not UnkeptKeys(source).Any(Function(k) KeyId(k) = id) Then Return New List(Of String)

            listing.Add(source)

        Next

        Return listing

    End Function

    ''' <summary>
    ''' Finds the entries other than <paramref name="target"/> that have a key from one of
    ''' <paramref name="target"/>'s merged sources, unchanged or captured
    ''' </summary>
    '''
    ''' <param name="target">
    ''' The name of the entry the sources were merged into
    ''' </param>
    '''
    ''' <param name="key">
    ''' The source key to look for
    ''' </param>
    '''
    ''' <returns>
    ''' The names of the other entries that kept <paramref name="key"/>, empty if none did or if
    ''' <paramref name="target"/> kept it too. A wildcard can cover a key that the re-diff still
    ''' lists as removed, because only an added key captures there, and stripping it then would
    ''' leave every target pointing at another and none listing it.
    ''' </returns>
    Private Function TargetsKeeping(target As String, key As iniKey) As HashSet(Of String)

        Dim id = KeyId(key)
        Dim others As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        For Each source In _state.MergedEntries.MergeDict(target)

            Dim keys As Dictionary(Of String, HashSet(Of String)) = Nothing
            Dim keepers As HashSet(Of String) = Nothing
            If _keptBy.TryGetValue(source, keys) AndAlso keys.TryGetValue(id, keepers) Then others.UnionWith(keepers)

        Next

        If others.Remove(target) Then others.Clear()
        Return others

    End Function

    ''' <summary>
    ''' Records, for every merged source entry, which of the new entries it went into have each of
    ''' its keys, either unchanged or captured by one of their keys
    ''' </summary>
    '''
    ''' <returns>
    ''' The map <c> _keptBy </c> holds
    ''' </returns>
    Private Function BuildKeptBy() As Dictionary(Of String, Dictionary(Of String, HashSet(Of String)))

        Dim keptBy As New Dictionary(Of String, Dictionary(Of String, HashSet(Of String)))(StringComparer.OrdinalIgnoreCase)

        For Each merge In _state.MergedEntries.MergeDict

            Dim newSection = _file2.GetSection(merge.Key)
            If newSection Is Nothing Then Continue For

            Dim newKeys = newSection.Keys.ToList()
            Dim newIds = New HashSet(Of String)(newKeys.Select(Function(newKey) KeyId(newKey)))

            For Each source In merge.Value

                Dim oldSection = _file1.GetSection(source)
                If oldSection Is Nothing Then Continue For

                Dim keys As Dictionary(Of String, HashSet(Of String)) = Nothing
                If Not keptBy.TryGetValue(source, keys) Then

                    keys = New Dictionary(Of String, HashSet(Of String))
                    keptBy(source) = keys

                End If

                For Each oldKey In oldSection.Keys

                    Dim id = KeyId(oldKey)
                    If Not newIds.Contains(id) AndAlso Not newKeys.Any(Function(newKey) KeyComparisonStrategyFactory.CompareKeys(newKey, oldKey)) Then Continue For

                    If Not keys.ContainsKey(id) Then keys(id) = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                    keys(id).Add(merge.Key)

                Next

            Next

        Next

        Return keptBy

    End Function

    ''' <summary>
    ''' Identifies a key by its type and value, ignoring case, so keys that differ only in their
    ''' number compare equal
    ''' </summary>
    '''
    ''' <param name="key">
    ''' The key to identify
    ''' </param>
    '''
    ''' <returns>
    ''' The key's type and value, upper-cased
    ''' </returns>
    Private Shared Function KeyId(key As iniKey) As String

        Return $"{key.KeyType}|{key.Value}".ToUpperInvariant()

    End Function

    ''' <summary>
    ''' Builds a list of <c> iniKey </c> objects from all old entries
    ''' named in <paramref name="mergeSourceNames"/>, plus optionally from
    ''' <paramref name="targetEntry"/> itself if it existed in file1 and is not
    ''' already a named merge source. We drop a key whose value (case-insensitive) an earlier
    ''' key already has, and credit it to every source that has it, in source order.
    ''' <paramref name="targetEntry"/> is credited only for values no source has. The entries that
    ''' moved keys into <paramref name="targetEntry"/> are credited after the merge sources, but
    ''' their keys aren't combined, so a moved key still compares as added.
    ''' </summary>
    '''
    ''' <param name="mergeSourceNames">
    ''' Names of the old entries whose keys are to be combined
    ''' </param>
    '''
    ''' <param name="targetEntry">
    ''' The entry receiving the merged content; when non-empty,
    ''' its own file1 keys are appended if it existed in file1
    ''' and is not listed as a source, and the keys moved into it are credited.
    ''' Pass an empty string to skip both.
    ''' </param>
    '''
    ''' <returns>
    ''' A <c> CombinedOldKeyResult </c> with the deduplicated key list
    ''' and the value → source-entry attribution map
    ''' </returns>
    Private Function BuildCombinedOldKeys(mergeSourceNames As IEnumerable(Of String),
                                          targetEntry As String) As CombinedOldKeyResult

        Dim combinedKeys As New List(Of iniKey)
        Dim sources As New Dictionary(Of String, List(Of String))(StringComparer.OrdinalIgnoreCase)

        For Each oldEntryName In mergeSourceNames

            Dim oldEnt = _file1.GetSection(oldEntryName)
            If oldEnt Is Nothing Then Continue For

            For Each key In oldEnt.Keys

                If Not sources.ContainsKey(key.Value) Then

                    combinedKeys.Add(key)
                    sources.Add(key.Value, New List(Of String))

                End If

                If Not sources(key.Value).Contains(oldEntryName) Then sources(key.Value).Add(oldEntryName)

            Next

        Next

        If targetEntry <> "" AndAlso
           _file1.Contains(targetEntry) AndAlso
           Not mergeSourceNames.Contains(targetEntry) Then

            For Each key In _file1.GetSection(targetEntry).Keys

                If sources.ContainsKey(key.Value) Then Continue For

                combinedKeys.Add(key)
                sources.Add(key.Value, New List(Of String) From {targetEntry})

            Next

        End If

        If targetEntry <> "" Then

            For Each moved In _state.KeyMovements.SourcesMovedInto(targetEntry)

                If Not sources.ContainsKey(moved.Key) Then sources.Add(moved.Key, New List(Of String))
                sources(moved.Key).AddRange(moved.Value.Except(sources(moved.Key)).ToList())

            Next

        End If

        Dim sourceMap = sources.ToDictionary(Function(kvp) kvp.Key, Function(kvp) DiffOutputRenderer.CreditSources(kvp.Value), StringComparer.OrdinalIgnoreCase)
        Return New CombinedOldKeyResult(combinedKeys, sourceMap)

    End Function

    ''' <summary>
    ''' The source keys a merge target doesn't have and needn't list, as
    ''' <c> StripKeysKeptElsewhere </c> found them
    ''' </summary>
    Friend Class KeptElsewhereInfo

        Public ReadOnly KeyCount As Integer
        Public ReadOnly Entries As List(Of String)
        Public ReadOnly UnkeptCount As Integer
        Public ReadOnly UnkeptSources As List(Of String)

        ''' <summary>
        ''' Creates a new <c> KeptElsewhereInfo </c>
        ''' </summary>
        '''
        ''' <param name="count">The number of source keys other entries kept</param>
        '''
        ''' <param name="keptBy">The entries that kept them, sorted by name</param>
        '''
        ''' <param name="unkept">The number of split sources' keys no entry kept</param>
        '''
        ''' <param name="listedWith">The split sources whose sections list those keys, sorted by name</param>
        Public Sub New(count As Integer, keptBy As List(Of String), unkept As Integer, listedWith As List(Of String))

            KeyCount = count
            Entries = keptBy
            UnkeptCount = unkept
            UnkeptSources = listedWith

        End Sub

    End Class

    ''' <summary>
    ''' The result of <see cref="BuildCombinedOldKeys"/>
    ''' </summary>
    Private Class CombinedOldKeyResult

        Public ReadOnly Keys As List(Of iniKey)
        Public ReadOnly SourceEntryMap As Dictionary(Of String, String)

        ''' <summary>
        ''' Creates a new <c> CombinedOldKeyResult </c>
        ''' </summary>
        '''
        ''' <param name="keys">The combined, deduplicated old keys</param>
        '''
        ''' <param name="sourceMap">Map of key value → the old entries it came from, formatted by <see cref="DiffOutputRenderer.CreditSources"/></param>
        Public Sub New(keys As List(Of iniKey), sourceMap As Dictionary(Of String, String))

            Me.Keys = keys
            SourceEntryMap = sourceMap

        End Sub

    End Class

End Class
