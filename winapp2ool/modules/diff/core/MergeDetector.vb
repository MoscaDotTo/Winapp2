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
''' Detects when a removed entry has been renamed to or merged into one or more new entries.
''' Matches candidates by comparing their <c> FileKey </c> and <c> RegKey </c> values only,
''' records renames and mergers in <see cref="DiffState"/>, and invokes a callback for
''' key-level change tracking when it records a new rename. <br /><br />
'''
''' Finding a match (<see cref="FindBestMatch"/>) reads no tracker state and is safe to run in
''' parallel. Recording one (<see cref="RecordMatch"/>) is not, and the caller records matches
''' one at a time in entry name order so contested renames resolve the same way on every run.
''' </summary>
Public Class MergeDetector

    Private ReadOnly _state As DiffState
    Private ReadOnly _diffFile As iniFile

    ''' <summary>Creates a new <c> MergeDetector </c></summary>
    '''
    ''' <param name="diffState">
    ''' Shared diff state tracking all entry changes
    ''' </param>
    '''
    ''' <param name="newFile">
    ''' The new version of winapp2.ini
    ''' </param>
    Public Sub New(diffState As DiffState,
                   newFile As iniFile)

        _state = diffState
        _diffFile = newFile

    End Sub

    ''' <summary>
    ''' Records the rename or mergers in <paramref name="bestMatch"/> in the <see cref="DiffState"/>
    ''' tracking collections. A rename whose target is already the rename of a different old entry
    ''' is recorded as a merger into that target. Not thread-safe.
    ''' </summary>
    '''
    ''' <param name="bestMatch">
    ''' The result of <see cref="FindBestMatch"/> for <paramref name="oldSection"/>
    ''' </param>
    '''
    ''' <param name="oldSection">
    ''' The removed entry being recorded
    ''' </param>
    Public Sub RecordMatch(bestMatch As MatchResult, oldSection As iniSection)

        If Not bestMatch.IsRename AndAlso Not bestMatch.IsMerge Then Return
        If bestMatch.IsRename AndAlso ConfirmRename(bestMatch.TargetName, oldSection) Then Return

        For Each targetName In bestMatch.AllTargetNames

            Dim mergeTarget = _diffFile.GetSection(targetName)
            If mergeTarget IsNot Nothing Then TrackMerger(oldSection, mergeTarget)

        Next

    End Sub

    ''' <summary>
    ''' Returns the cached <c> iniSection </c> for the given old entry, inserting it on first access
    ''' </summary>
    ''' 
    ''' <param name="section">
    ''' The old entry to cache
    ''' </param>
    ''' 
    ''' <returns>
    ''' The cached instance (always the same object for a given entry name)
    ''' </returns>
    Private Function GetOrCreateCachedSection(section As iniSection) As iniSection

        SyncLock _state.Caches.CachedOldEntries

            Dim cached As iniSection = Nothing
            If Not _state.Caches.CachedOldEntries.TryGetValue(section.Name, cached) Then

                cached = section
                _state.Caches.CachedOldEntries(section.Name) = cached

            End If

            Return cached

        End SyncLock

    End Function

    ''' <summary>
    ''' Returns the cached <c> iniSection </c> for the given new entry, inserting it on first access
    ''' </summary>
    ''' 
    ''' <param name="section">
    ''' The new entry to cache
    ''' </param>
    ''' 
    ''' <returns>
    ''' The cached instance (always the same object for a given entry name)
    ''' </returns>
    Private Function GetOrCreateNewCachedSection(section As iniSection) As iniSection

        SyncLock _state.Caches.CachedNewEntries

            Dim cached As iniSection = Nothing
            If Not _state.Caches.CachedNewEntries.TryGetValue(section.Name, cached) Then

                cached = section
                _state.Caches.CachedNewEntries(section.Name) = cached

            End If

            Return cached

        End SyncLock

    End Function

    ''' <summary>
    ''' Scores each candidate against the old entry's FileKeys and RegKeys and returns a
    ''' <see cref="MatchResult"/>. The first candidate in <paramref name="candidates"/> that
    ''' qualifies as a rename wins at once: it must be an added entry that matches every old
    ''' FileKey and RegKey, has the same number of each, and raised no wildcard-reduction flag.
    ''' Gaining patterns doesn't stop a rename. Otherwise every candidate matching at least one key
    ''' becomes a merger target. Reads no tracker state, so it is safe to call in parallel.
    ''' </summary>
    '''
    ''' <param name="candidates">
    ''' New entries (added or modified) to score against <paramref name="oldSection"/>
    ''' </param>
    '''
    ''' <param name="oldSection">
    ''' The removed entry whose keys are used as the match baseline
    ''' </param>
    '''
    ''' <returns>
    ''' A <c> MatchResult </c> describing the best outcome found; all flags <c> False </c> if no match qualifies
    ''' </returns>
    Public Function FindBestMatch(candidates As List(Of iniSection),
                                  oldSection As iniSection) As MatchResult

        Dim result As New MatchResult()
        Dim highestMatchCount = 0
        Dim qualifyingMergeTargets As New List(Of String)

        oldSection = GetOrCreateCachedSection(oldSection)
        Dim oldFileKeys = oldSection.Keys.GetByType("FileKey")
        Dim oldRegKeys = oldSection.Keys.GetByType("RegKey")

        Dim oldHasFileKeys = oldFileKeys.Count > 0
        Dim oldHasRegKeys = oldRegKeys.Count > 0

        result.TotalOldKeys = oldFileKeys.Count + oldRegKeys.Count

        For Each candidateSection In candidates

            Dim newSection = GetOrCreateNewCachedSection(candidateSection)
            Dim matchInfo = GetOrComputeMatchInfo(oldSection.Name, candidateSection.Name, newSection, oldFileKeys, oldRegKeys, oldHasFileKeys, oldHasRegKeys)

            If matchInfo.TotalMatches > highestMatchCount Then highestMatchCount = matchInfo.TotalMatches

            If matchInfo.FileKeyMatches = 0 AndAlso matchInfo.RegKeyMatches = 0 Then Continue For

            ' A rename target must be a name that didn't exist in the old file; matching onto a
            ' pre-existing (modified) entry means the old entry was absorbed into it - a merger
            Dim candidateIsAdded = _state.ModifiedEntries.AddedEntryNames.Contains(candidateSection.Name)

            Dim isRename = candidateIsAdded AndAlso
                      matchInfo.AllKeysMatched AndAlso
                      matchInfo.CountsMatch AndAlso
                      Not matchInfo.PossibleWildCardReduction

            If isRename Then

                result.IsRename = True
                result.TargetName = candidateSection.Name
                result.AllTargetNames.Add(candidateSection.Name)
                result.TotalMatchedKeys = matchInfo.TotalMatches
                Return result

            End If

            If Not qualifyingMergeTargets.Contains(candidateSection.Name) Then qualifyingMergeTargets.Add(candidateSection.Name)

        Next

        If qualifyingMergeTargets.Count = 0 Then Return result

        result.IsMerge = True
        result.AllTargetNames.AddRange(qualifyingMergeTargets)
        result.TargetName = qualifyingMergeTargets(0)
        result.TotalMatchedKeys = highestMatchCount
        Return result

    End Function

    ''' <summary>
    ''' Returns a cached <see cref="KeyMatchInfo"/> for the old/new entry pair, computing and caching it on first access.
    ''' The cache key is <c> "{oldName}|{newName}" </c>, compared case-sensitively.
    ''' </summary>
    ''' 
    ''' <param name="oldName">
    ''' Name of the old (removed) entry; forms the cache key prefix
    ''' </param>
    ''' 
    ''' <param name="newName">
    ''' Name of the new (candidate) entry; forms the cache key suffix
    ''' </param>
    ''' 
    ''' <param name="newSection">
    ''' The candidate section whose keys are matched against the old entry's keys
    ''' </param>
    ''' 
    ''' <param name="oldFileKeys">
    ''' Pre-computed FileKey list from the old entry
    ''' </param>
    ''' 
    ''' <param name="oldRegKeys">
    ''' Pre-computed RegKey list from the old entry
    ''' </param>
    ''' 
    ''' <param name="oldHasFileKeys">
    ''' Indicates whether the old entry has any FileKeys
    ''' </param>
    ''' 
    ''' <param name="oldHasRegKeys">
    ''' Indicates whether the old entry has any RegKeys
    ''' </param>
    ''' 
    ''' <returns>
    ''' A <c> KeyMatchInfo </c> with match counts and flags for the old/new pair
    ''' </returns>
    Private Function GetOrComputeMatchInfo(oldName As String,
                                           newName As String,
                                           newSection As iniSection,
                                           oldFileKeys As IReadOnlyList(Of iniKey),
                                           oldRegKeys As IReadOnlyList(Of iniKey),
                                           oldHasFileKeys As Boolean,
                                           oldHasRegKeys As Boolean) As KeyMatchInfo

        Dim cacheKey = $"{oldName}|{newName}"
        Dim cachedResult As KeyMatchInfo = Nothing
        If _state.Caches.MatchInfoCache.TryGetValue(cacheKey, cachedResult) Then Return cachedResult

        Dim matchInfo = AssessKeyMatches(newSection, oldFileKeys, oldRegKeys, oldHasFileKeys, oldHasRegKeys)
        _state.Caches.MatchInfoCache.TryAdd(cacheKey, matchInfo)
        Return matchInfo

    End Function

    ''' <summary>
    ''' Compares the old entry's FileKeys and RegKeys against the corresponding lists in
    ''' <paramref name="newSection"/> and returns a fully populated <see cref="KeyMatchInfo"/>.
    ''' Key types absent from the old entry are treated as fully matched. The counts match only
    ''' when every old key of that type matched and the new entry has as many keys of that type,
    ''' which doesn't require each old key to have matched a different new key.
    ''' </summary>
    ''' 
    ''' <param name="newSection">
    ''' The candidate new entry to match against
    ''' </param>
    ''' 
    ''' <param name="oldFileKeys">
    ''' FileKeys from the old entry
    ''' </param>
    ''' 
    ''' <param name="oldRegKeys">
    ''' RegKeys from the old entry
    ''' </param>
    ''' 
    ''' <param name="oldHasFileKeys">
    ''' Indicates whether the old entry has any FileKeys
    ''' </param>
    ''' 
    ''' <param name="oldHasRegKeys">
    ''' Indicates whether the old entry has any RegKeys
    ''' </param>
    ''' 
    ''' <returns>
    ''' A <c> KeyMatchInfo </c> populated with per-type match counts, flags, and matched key sets
    ''' </returns>
    Private Function AssessKeyMatches(newSection As iniSection,
                                      oldFileKeys As IReadOnlyList(Of iniKey),
                                      oldRegKeys As IReadOnlyList(Of iniKey),
                                      oldHasFileKeys As Boolean,
                                      oldHasRegKeys As Boolean) As KeyMatchInfo

        Dim info As New KeyMatchInfo()

        Dim newFileKeys = newSection.Keys.GetByType("FileKey")
        Dim newRegKeys = newSection.Keys.GetByType("RegKey")

        If oldHasFileKeys Then

            info.FileKeyMatches = CountMatches(oldFileKeys, newFileKeys, DisallowedPaths,
                                               info.PossibleWildCardReduction, info.MatchedOldFileKeys)

            info.AllFileKeysMatched = info.FileKeyMatches = oldFileKeys.Count
            info.FileKeyCountsMatch = info.AllFileKeysMatched AndAlso newFileKeys.Count = oldFileKeys.Count

        Else

            info.AllFileKeysMatched = True
            info.FileKeyCountsMatch = True

        End If

        If oldHasRegKeys Then

            info.RegKeyMatches = CountMatches(oldRegKeys, newRegKeys, DisallowedPaths,
                                              info.PossibleWildCardReduction, info.MatchedOldRegKeys)

            info.AllRegKeysMatched = info.RegKeyMatches = oldRegKeys.Count
            info.RegKeyCountsMatch = info.AllRegKeysMatched AndAlso newRegKeys.Count = oldRegKeys.Count

        Else

            info.AllRegKeysMatched = True
            info.RegKeyCountsMatch = True

        End If

        info.TotalMatches = info.FileKeyMatches + info.RegKeyMatches
        info.AllKeysMatched = info.AllFileKeysMatched AndAlso info.AllRegKeysMatched
        info.CountsMatch = info.FileKeyCountsMatch AndAlso info.RegKeyCountsMatch

        Return info

    End Function

    ''' <summary>
    ''' Counts how many keys in <paramref name="oldKeys"/> are matched by at least one key in
    ''' <paramref name="newKeys"/>. Several old keys may match the same new key. We try an exact
    ''' value match (ignoring case) first, then <see cref="KeyComparisonStrategyFactory.CompareKeys"/>
    ''' against each new key in turn. An old key whose whole value is in
    ''' <paramref name="disallowedValues"/> never counts, and a <see cref="KeyComparisonStrategyFactory.CompareKeys"/>
    ''' match is rejected when the new key's path (before its first <c> | </c>) is in it. Both
    ''' lookups use the set's own comparer, which for <c> DisallowedPaths </c> is case-sensitive.
    ''' </summary>
    ''' 
    ''' <param name="oldKeys">
    ''' Keys from the old entry to match
    ''' </param>
    ''' 
    ''' <param name="newKeys">
    ''' Keys from the new entry to match against
    ''' </param>
    ''' 
    ''' <param name="disallowedValues">
    ''' Values too broad to count as meaningful matches; may be <c> Nothing </c>
    ''' </param>
    ''' 
    ''' <param name="possibleWildCardReduction">
    ''' Set to <c> True </c> when some old key is covered only by new keys whose match raised the
    ''' wildcard-reduction verdict from <see cref="KeyComparisonStrategyFactory.CompareKeys"/>.
    ''' Never set back to <c> False </c>, and independent of key order.
    ''' </param>
    ''' 
    ''' <param name="matchedKeys">
    ''' Populated with each old key that was successfully matched
    ''' </param>
    ''' 
    ''' <returns>
    ''' The number of old keys that were matched by at least one new key
    ''' </returns>
    Private Function CountMatches(oldKeys As IEnumerable(Of iniKey),
                                  newKeys As IEnumerable(Of iniKey),
                                  disallowedValues As HashSet(Of String),
                            ByRef possibleWildCardReduction As Boolean,
                                  matchedKeys As HashSet(Of iniKey)) As Integer

        Dim matchCount = 0
        Dim newKeysList = newKeys.ToList()
        Dim newKeyValues As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        For Each newKey In newKeysList
            If Not newKeyValues.Contains(newKey.Value) Then newKeyValues.Add(newKey.Value)
        Next

        For Each oldKey In oldKeys

            If disallowedValues IsNot Nothing AndAlso disallowedValues.Contains(oldKey.Value) Then Continue For

            Dim matched = False
            Dim keyWildCardReduction = False

            If newKeyValues.Contains(oldKey.Value) Then

                matched = True

            Else

                ' A key that one new key covers cleanly raises no flag, whichever new key comes first
                For Each newKey In newKeysList

                    Dim wildCardReduction = False
                    Dim keyMatched = KeyComparisonStrategyFactory.CompareKeys(newKey, oldKey, wildCardReduction)
                    If keyMatched AndAlso disallowedValues IsNot Nothing Then

                        Dim newKeyPath = GetPathWithoutFlags(newKey.Value)
                        If disallowedValues.Contains(newKeyPath) Then keyMatched = False

                    End If

                    If Not keyMatched Then Continue For

                    matched = True
                    keyWildCardReduction = wildCardReduction
                    If Not wildCardReduction Then Exit For

                Next

            End If

            If Not matched Then Continue For

            matchCount += 1
            matchedKeys.Add(oldKey)
            possibleWildCardReduction = possibleWildCardReduction OrElse keyWildCardReduction

        Next

        Return matchCount

    End Function

    ''' <summary>
    ''' Returns the path portion of a key value, stripping any pipe-delimited flags
    ''' </summary>
    ''' 
    ''' <param name="value">
    ''' The raw key value string, optionally containing a <c> | </c> separator
    ''' </param>
    ''' 
    ''' <returns>
    ''' The substring before the first <c> | </c>, or the full string if no pipe is present
    ''' </returns>
    Private Function GetPathWithoutFlags(value As String) As String

        Return If(value.Contains("|"), value.Substring(0, value.IndexOf("|", StringComparison.InvariantCultureIgnoreCase)), value)

    End Function

    ''' <summary>
    ''' Attempts to record a rename from <paramref name="oldSection"/> to <paramref name="newName"/>.
    ''' If <paramref name="newName"/> is already registered as a rename target, the registration is
    ''' rejected and the caller falls back to merger tracking. Records no key-level changes, since
    ''' a later merger can still demote the rename.
    ''' </summary>
    '''
    ''' <param name="newName">
    ''' The candidate new entry name
    ''' </param>
    '''
    ''' <param name="oldSection">
    ''' The removed entry being renamed
    ''' </param>
    '''
    ''' <returns>
    ''' <c> True </c> if the rename was recorded <br />
    ''' <c> False </c> if <paramref name="newName"/> is already a rename target
    ''' </returns>
    Private Function ConfirmRename(newName As String, oldSection As iniSection) As Boolean

        If _state.MergedEntries.RenamedEntryPairs.ContainsKey(newName) Then Return False

        _state.MergedEntries.RenamedEntryNames.Add(newName)
        _state.MergedEntries.RenamedEntryPairs.Add(newName, oldSection.Name)

        Return True

    End Function

    ''' <summary>
    ''' Records a merger relationship between <paramref name="oldSection"/> and <paramref name="newSection"/>
    ''' in <c> MergedEntryNames </c>, <c> MergeDict </c> and <c> OldToNewMergeDict </c>. If
    ''' <paramref name="newSection"/> was previously recorded as a rename target, the rename is demoted
    ''' to a merger and its source is folded in.
    ''' </summary>
    '''
    ''' <param name="oldSection">
    ''' The removed entry that was merged
    ''' </param>
    ''' 
    ''' <param name="newSection">
    ''' The new entry that received content from <paramref name="oldSection"/>
    ''' </param>
    Private Sub TrackMerger(oldSection As iniSection, newSection As iniSection)

        Dim mergeName = newSection.Name
        Dim oldName = oldSection.Name

        _state.MergedEntries.MergedEntryNames.Add(mergeName)

        If Not _state.MergedEntries.MergeDict.ContainsKey(mergeName) Then _state.MergedEntries.MergeDict.Add(mergeName, New List(Of String))

        If Not _state.MergedEntries.MergeDict(mergeName).Contains(oldName) Then _state.MergedEntries.MergeDict(mergeName).Add(oldName)

        If Not _state.MergedEntries.OldToNewMergeDict.ContainsKey(oldName) Then _state.MergedEntries.OldToNewMergeDict.Add(oldName, New List(Of String))

        If Not _state.MergedEntries.OldToNewMergeDict(oldName).Contains(mergeName) Then _state.MergedEntries.OldToNewMergeDict(oldName).Add(mergeName)

        If Not _state.MergedEntries.RenamedEntryNames.Contains(mergeName) Then Return

        Dim renameHolder As String = Nothing
        If Not _state.MergedEntries.RenamedEntryPairs.TryGetValue(mergeName, renameHolder) Then Return

        If Not _state.MergedEntries.MergeDict(mergeName).Contains(renameHolder) Then _state.MergedEntries.MergeDict(mergeName).Add(renameHolder)

        If Not _state.MergedEntries.OldToNewMergeDict.ContainsKey(renameHolder) Then _state.MergedEntries.OldToNewMergeDict.Add(renameHolder, New List(Of String))

        If Not _state.MergedEntries.OldToNewMergeDict(renameHolder).Contains(mergeName) Then _state.MergedEntries.OldToNewMergeDict(renameHolder).Add(mergeName)

        _state.MergedEntries.RenamedEntryPairs.Remove(mergeName)
        _state.MergedEntries.RenamedEntryNames.Remove(mergeName)

    End Sub

End Class

''' <summary>
''' Result of matching a removed entry against candidates
''' </summary>
Public Class MatchResult

    ''' <summary>
    ''' Indicates whether the match is a rename: the target is an added entry that matched every
    ''' old FileKey and RegKey with equal counts and raised no wildcard-reduction flag
    ''' </summary>
    Public Property IsRename As Boolean

    ''' <summary>
    ''' Indicates whether the old entry was merged into one or more new entries
    ''' </summary>
    Public Property IsMerge As Boolean

    ''' <summary>
    ''' The primary target entry name: the rename target or the first qualifying merge target
    ''' </summary>
    Public Property TargetName As String

    ''' <summary>
    ''' All target entry names; may contain multiple entries in the case of a split merger
    ''' </summary>
    Public Property AllTargetNames As New List(Of String)

    ''' <summary>
    ''' Number of old keys matched in the rename target, or the highest match count among all
    ''' candidates for a merger
    ''' </summary>
    Public Property TotalMatchedKeys As Integer

    ''' <summary>
    ''' Total number of old FileKeys and RegKeys in the entry being assessed
    ''' </summary>
    Public Property TotalOldKeys As Integer

End Class
