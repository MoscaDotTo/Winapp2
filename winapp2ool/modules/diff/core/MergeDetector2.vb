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
''' Matches candidates by comparing <c>iniKey2</c> values, records confirmed renames and
''' mergers in <c>DiffState</c>, and invokes a callback for key-level change tracking
''' when a match is confirmed.
''' </summary>
Public Class MergeDetector2

    Private ReadOnly _state As DiffState
    Private ReadOnly _diffFile2 As iniFile2
    Private ReadOnly _findModificationsCallback As Action(Of iniSection2, iniSection2)

    ''' <summary>
    ''' Initializes a new instance of <c>MergeDetector2</c>
    ''' </summary>
    ''' 
    ''' <param name="diffState">
    ''' Shared diff state tracking all entry changes
    ''' </param>
    ''' 
    ''' <param name="newFile">
    ''' The new version of winapp2.ini
    ''' </param>
    ''' 
    ''' <param name="findModsCallback">
    ''' Callback invoked to track key-level changes when a rename or merger is confirmed
    ''' </param>
    Public Sub New(diffState As DiffState,
                   newFile As iniFile2,
                   findModsCallback As Action(Of iniSection2, iniSection2))

        _state = diffState
        _diffFile2 = newFile
        _findModificationsCallback = findModsCallback

    End Sub

    ''' <summary>
    ''' Determines whether a removed entry has been renamed or merged into one or more new entries.
    ''' Updates <c>DiffState</c> tracking collections accordingly.
    ''' </summary>
    '''
    ''' <param name="candidates">
    ''' New entries (added or modified) that are potential rename or merger targets for <paramref name="oldSection2"/>
    ''' </param>
    '''
    ''' <param name="oldSection2">
    ''' The removed entry being assessed
    ''' </param>
    '''
    ''' <returns>
    ''' <c>True</c> if a rename or merger was recorded; <c>False</c> if no match was found
    ''' </returns>
    Public Function AssessRenamesAndMergers(candidates As List(Of iniSection2),
                                            oldSection2 As iniSection2) As Boolean

        If candidates.Count = 0 Then Return False

        Dim cachedOld = GetOrCreateCachedSection2(oldSection2)
        Dim bestMatch = FindBestMatch(candidates, cachedOld)

        If bestMatch.IsRename Then

            If ConfirmRename(bestMatch.TargetName, oldSection2) Then Return True

            ' Rename rejected — target already renamed from another entry.
            ' Treat this as a merger instead so the entry isn't silently dropped.
            TrackBestMatches(False, bestMatch, oldSection2)
            Return True

        End If

        If bestMatch.IsMerge OrElse bestMatch.HasPartialMatch Then

            TrackBestMatches(bestMatch.IsMerge, bestMatch, oldSection2)
            Return True

        End If

        Return False

    End Function

    ''' <summary>
    ''' Dispatches merger tracking for all qualifying targets in <paramref name="bestMatch"/>.
    ''' When <paramref name="isMerge"/> is <c>False</c>, only the primary target is tracked
    ''' (used when a rename was rejected and the entry is reclassified as a merger).
    ''' </summary>
    ''' 
    ''' <param name="isMerge">
    ''' <c>True</c> to track all targets in <c>AllTargetNames</c> <br />
    ''' <c>False</c> to track only <c>TargetName</c>
    ''' </param>
    ''' 
    ''' <param name="bestMatch">
    ''' The match result from <c>FindBestMatch</c>
    ''' </param>
    ''' 
    ''' <param name="oldSection2">
    ''' The removed entry being tracked
    ''' </param>
    Private Sub TrackBestMatches(isMerge As Boolean,
                                 bestMatch As MatchResult,
                                 oldSection2 As iniSection2)

        If Not isMerge Then

            Dim singleTarget = _diffFile2.GetSection(bestMatch.TargetName)
            If singleTarget IsNot Nothing Then TrackMerger(oldSection2, singleTarget)

            Return

        End If

        For Each targetName In bestMatch.AllTargetNames

            Dim mergeTarget = _diffFile2.GetSection(targetName)
            If mergeTarget IsNot Nothing Then TrackMerger(oldSection2, mergeTarget)

        Next

    End Sub

    ''' <summary>
    ''' Returns the cached <c>iniSection2</c> for the given old entry, inserting it on first access
    ''' </summary>
    ''' 
    ''' <param name="section2">
    ''' The old entry to cache
    ''' </param>
    ''' 
    ''' <returns>
    ''' The cached instance (always the same object for a given entry name)
    ''' </returns>
    Private Function GetOrCreateCachedSection2(section2 As iniSection2) As iniSection2

        SyncLock _state.Caches.CachedOldEntries2

            Dim cached As iniSection2 = Nothing
            If Not _state.Caches.CachedOldEntries2.TryGetValue(section2.Name, cached) Then

                cached = section2
                _state.Caches.CachedOldEntries2(section2.Name) = cached

            End If

            Return cached

        End SyncLock

    End Function

    ''' <summary>
    ''' Returns the cached <c>iniSection2</c> for the given new entry, inserting it on first access
    ''' </summary>
    ''' 
    ''' <param name="section2">
    ''' The new entry to cache
    ''' </param>
    ''' 
    ''' <returns>
    ''' The cached instance (always the same object for a given entry name)
    ''' </returns>
    Private Function GetOrCreateNewCachedSection2(section2 As iniSection2) As iniSection2

        SyncLock _state.Caches.CachedNewEntries2

            Dim cached As iniSection2 = Nothing
            If Not _state.Caches.CachedNewEntries2.TryGetValue(section2.Name, cached) Then

                cached = section2
                _state.Caches.CachedNewEntries2(section2.Name) = cached

            End If

            Return cached

        End SyncLock

    End Function

    ''' <summary>
    ''' Scores each candidate against the old entry's FileKeys and RegKeys and returns
    ''' the best-fitting <c>MatchResult</c>. Evaluates rename (all keys matched, counts equal,
    ''' no structural changes), merger (one or more keys matched), and partial-match outcomes.
    ''' </summary>
    ''' 
    ''' <param name="candidates">
    ''' New entries to score against <paramref name="oldSection2"/>
    ''' </param>
    ''' 
    ''' <param name="oldSection2">
    ''' The removed entry whose keys are used as the match baseline
    ''' </param>
    ''' 
    ''' <returns>
    ''' A <c>MatchResult</c> describing the best outcome found; all flags <c>False</c> if no match qualifies
    ''' </returns>
    Private Function FindBestMatch(candidates As List(Of iniSection2),
                                   oldSection2 As iniSection2) As MatchResult

        Dim result As New MatchResult()
        Dim highestMatchCount = 0
        Dim bestCandidateName = ""
        Dim foundMerger = False
        Dim qualifyingMergeTargets As New List(Of String)

        Dim oldFileKeys = oldSection2.Keys.GetByType("FileKey")
        Dim oldRegKeys = oldSection2.Keys.GetByType("RegKey")

        Dim oldHasFileKeys = oldFileKeys.Count > 0
        Dim oldHasRegKeys = oldRegKeys.Count > 0

        result.TotalOldKeys = oldFileKeys.Count + oldRegKeys.Count

        For Each candidateSection In candidates

            Dim newSection2 = GetOrCreateNewCachedSection2(candidateSection)
            Dim matchInfo = GetOrComputeMatchInfo(oldSection2.Name, candidateSection.Name, newSection2, oldFileKeys, oldRegKeys, oldHasFileKeys, oldHasRegKeys)

            If matchInfo.TotalMatches > highestMatchCount Then
                highestMatchCount = matchInfo.TotalMatches
                bestCandidateName = candidateSection.Name
            End If

            If matchInfo.FileKeyMatches = 0 AndAlso matchInfo.RegKeyMatches = 0 Then Continue For

            Dim thisSpecificPairIsRename As Boolean
            SyncLock _state.MergedEntries

                thisSpecificPairIsRename = _state.MergedEntries.RenamedEntryNames.Contains(candidateSection.Name) AndAlso
                                           IsRenamedFrom(candidateSection.Name, oldSection2.Name)

            End SyncLock
            If thisSpecificPairIsRename Then Continue For

            Dim isRename = matchInfo.AllKeysMatched AndAlso
                      matchInfo.CountsMatch AndAlso
                      Not matchInfo.MatchHadMoreParams AndAlso
                      Not matchInfo.PossibleWildCardReduction

            If isRename Then

                result.IsRename = True
                result.TargetName = candidateSection.Name
                result.AllTargetNames.Add(candidateSection.Name)
                result.TotalMatchedKeys = matchInfo.TotalMatches
                Return result

            End If

            Dim meetsThreshold = matchInfo.TotalMatches >= 1
            Dim isCompleteMerger = matchInfo.AllKeysMatched AndAlso Not isRename

            If meetsThreshold OrElse isCompleteMerger Then

                foundMerger = True
                If Not qualifyingMergeTargets.Contains(candidateSection.Name) Then qualifyingMergeTargets.Add(candidateSection.Name)

            End If

        Next

        If foundMerger AndAlso qualifyingMergeTargets.Count > 0 Then

            result.IsMerge = True
            result.AllTargetNames.AddRange(qualifyingMergeTargets)
            result.TargetName = qualifyingMergeTargets(0)
            result.TotalMatchedKeys = highestMatchCount
            Return result

        End If

        If Not String.IsNullOrEmpty(bestCandidateName) AndAlso highestMatchCount > 0 Then

            result.HasPartialMatch = True
            result.TargetName = bestCandidateName
            result.AllTargetNames.Add(bestCandidateName)
            result.TotalMatchedKeys = highestMatchCount

        End If

        Return result

    End Function

    ''' <summary>
    ''' Returns <c>True</c> if <paramref name="newName"/> is already recorded as a rename of <paramref name="oldName"/>
    ''' </summary>
    ''' 
    ''' <param name="newName">
    ''' The new entry name to look up in <c>RenamedEntryPairs</c>
    ''' </param>
    ''' 
    ''' <param name="oldName">
    ''' The expected old entry name to match against the stored value</param>
    ''' 
    ''' <returns>
    ''' <c>True</c> if the pair is an exact match; <c>False</c> otherwise
    ''' </returns>
    Private Function IsRenamedFrom(newName As String, oldName As String) As Boolean

        Dim storedOldName As String = Nothing
        If Not _state.MergedEntries.RenamedEntryPairs.TryGetValue(newName, storedOldName) Then Return False
        Return storedOldName.Equals(oldName, StringComparison.InvariantCultureIgnoreCase)

    End Function

    ''' <summary>
    ''' Returns a cached <c>KeyMatchInfo2</c> for the old/new entry pair, computing and caching it on first access.
    ''' The cache key is <c>"{oldName}|{newName}"</c>.
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
    ''' <param name="newSection2">
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
    ''' Whether the old entry has any FileKeys
    ''' </param>
    ''' 
    ''' <param name="oldHasRegKeys">
    ''' Whether the old entry has any RegKeys
    ''' </param>
    ''' 
    ''' <returns>
    ''' A <c>KeyMatchInfo2</c> with match counts and flags for the old/new pair
    ''' </returns>
    Private Function GetOrComputeMatchInfo(oldName As String,
                                           newName As String,
                                           newSection2 As iniSection2,
                                           oldFileKeys As IReadOnlyList(Of iniKey2),
                                           oldRegKeys As IReadOnlyList(Of iniKey2),
                                           oldHasFileKeys As Boolean,
                                           oldHasRegKeys As Boolean) As KeyMatchInfo2

        Dim cacheKey = $"{oldName}|{newName}"
        Dim cachedResult As KeyMatchInfo2 = Nothing
        If _state.Caches.MatchInfoCache2.TryGetValue(cacheKey, cachedResult) Then Return cachedResult

        Dim matchInfo = AssessKeyMatches(newSection2, oldFileKeys, oldRegKeys, oldHasFileKeys, oldHasRegKeys)
        _state.Caches.MatchInfoCache2.TryAdd(cacheKey, matchInfo)
        Return matchInfo

    End Function

    ''' <summary>
    ''' Compares the old entry's FileKeys and RegKeys against the corresponding lists in
    ''' <paramref name="newSection2"/> and returns a fully populated <c>KeyMatchInfo2</c>.
    ''' Key types absent from the old entry are treated as fully matched.
    ''' </summary>
    ''' 
    ''' <param name="newSection2">
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
    ''' Whether the old entry has any FileKeys
    ''' </param>
    ''' 
    ''' <param name="oldHasRegKeys">
    ''' Whether the old entry has any RegKeys
    ''' </param>
    ''' 
    ''' <returns>
    ''' A <c>KeyMatchInfo2</c> populated with per-type match counts, flags, and matched key sets
    ''' </returns>
    Private Function AssessKeyMatches(newSection2 As iniSection2,
                                      oldFileKeys As IReadOnlyList(Of iniKey2),
                                      oldRegKeys As IReadOnlyList(Of iniKey2),
                                      oldHasFileKeys As Boolean,
                                      oldHasRegKeys As Boolean) As KeyMatchInfo2

        Dim info As New KeyMatchInfo2()

        Dim newFileKeys = newSection2.Keys.GetByType("FileKey")
        Dim newRegKeys = newSection2.Keys.GetByType("RegKey")

        If oldHasFileKeys Then

            info.FileKeyMatches = CountMatches(oldFileKeys, newFileKeys, DisallowedPaths,
                                               info.MatchHadMoreParams, info.PossibleWildCardReduction,
                                               info.MatchedOldFileKeys)

            info.AllFileKeysMatched = info.FileKeyMatches = oldFileKeys.Count
            info.FileKeyCountsMatch = info.AllFileKeysMatched AndAlso newFileKeys.Count = oldFileKeys.Count

        Else

            info.AllFileKeysMatched = True
            info.FileKeyCountsMatch = True

        End If

        If oldHasRegKeys Then

            info.RegKeyMatches = CountMatches(oldRegKeys, newRegKeys, DisallowedPaths, info.MatchHadMoreParams,
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
    ''' <paramref name="newKeys"/>, using exact-value fast-path and wildcard/regex fallback.
    ''' Keys whose values appear in <paramref name="disallowedValues"/> are skipped.
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
    ''' Path values too broad to count as meaningful matches; may be <c>Nothing</c>
    ''' </param>
    ''' 
    ''' <param name="matchHadMoreParams">
    ''' Set to <c>True</c> if any matched new key has more pipe-delimited parameters than its old 
    ''' </param>
    ''' 
    ''' <param name="possibleWildCardReduction">
    ''' Set to <c>True</c> if any match appears to reduce wildcard coverage
    ''' </param>
    ''' 
    ''' <param name="matchedKeys">
    ''' Populated with each old key that was successfully matched
    ''' </param>
    ''' 
    ''' <returns>
    ''' The number of old keys that were matched by at least one new key
    ''' </returns>
    Private Function CountMatches(oldKeys As IEnumerable(Of iniKey2),
                                  newKeys As IEnumerable(Of iniKey2),
                                  disallowedValues As HashSet(Of String),
                            ByRef matchHadMoreParams As Boolean,
                            ByRef possibleWildCardReduction As Boolean,
                                  matchedKeys As HashSet(Of iniKey2)) As Integer

        Dim matchCount = 0
        Dim newKeysList = newKeys.ToList()
        Dim newKeyValues As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        For Each newKey In newKeysList
            If Not newKeyValues.Contains(newKey.Value) Then newKeyValues.Add(newKey.Value)
        Next

        For Each oldKey In oldKeys

            If disallowedValues IsNot Nothing AndAlso disallowedValues.Contains(oldKey.Value) Then Continue For

            Dim matched = False

            If newKeyValues.Contains(oldKey.Value) Then

                matched = True

            Else

                For Each newKey In newKeysList

                    Dim keyMatched = KeyComparisonStrategyFactory.CompareKeys(newKey, oldKey, matchHadMoreParams, possibleWildCardReduction)
                    If keyMatched AndAlso disallowedValues IsNot Nothing Then

                        Dim newKeyPath = GetPathWithoutFlags(newKey.Value)
                        If disallowedValues.Contains(newKeyPath) Then keyMatched = False

                    End If

                    If keyMatched Then matched = True : Exit For

                Next

            End If

            If matched Then matchCount += 1 : matchedKeys.Add(oldKey)

        Next

        Return matchCount

    End Function

    ''' <summary>
    ''' Returns the path portion of a key value, stripping any pipe-delimited flags
    ''' </summary>
    ''' 
    ''' <param name="value">
    ''' The raw key value string, optionally containing a <c>|</c> separator
    ''' </param>
    ''' 
    ''' <returns>
    ''' The substring before the first <c>|</c>, or the full string if no pipe is present
    ''' </returns>
    Private Function GetPathWithoutFlags(value As String) As String

        Return If(value.Contains("|"), value.Substring(0, value.IndexOf("|", StringComparison.InvariantCultureIgnoreCase)), value)

    End Function

    ''' <summary>
    ''' Attempts to record a rename from <paramref name="oldSection2"/> to <paramref name="newName"/>.
    ''' If <paramref name="newName"/> is already registered as a rename target from a different entry,
    ''' the registration is rejected and the caller should fall back to merger tracking.
    ''' On success, invokes the modifications callback to record key-level changes.
    ''' </summary>
    ''' 
    ''' <param name="newName">
    ''' The candidate new entry name
    ''' </param>
    ''' 
    ''' <param name="oldSection2">
    ''' The removed entry being renamed
    ''' </param>
    ''' 
    ''' <returns>
    ''' <c>True</c> if the rename was accepted or was already registered for this exact pair <br/>
    ''' <c>False</c> if <paramref name="newName"/> is already a rename target from a different old entry
    ''' </returns>
    Private Function ConfirmRename(newName As String, oldSection2 As iniSection2) As Boolean

        Dim newSection2 As iniSection2 = Nothing

        SyncLock _state.MergedEntries

            Dim storedOldName As String = Nothing

            If _state.MergedEntries.RenamedEntryPairs.TryGetValue(newName, storedOldName) Then

                Return storedOldName.Equals(oldSection2.Name, StringComparison.InvariantCultureIgnoreCase)

            End If

            _state.MergedEntries.RenamedEntryNames.Add(newName)
            _state.MergedEntries.RenamedEntryPairs.Add(newName, oldSection2.Name)
            newSection2 = _diffFile2.GetSection(newName)

        End SyncLock

        If _findModificationsCallback IsNot Nothing AndAlso newSection2 IsNot Nothing Then _findModificationsCallback(oldSection2, newSection2)

        Return True

    End Function

    ''' <summary>
    ''' Records a merger relationship between <paramref name="oldSection2"/> and <paramref name="newSection2"/>
    ''' in <c>MergeDict</c> and <c>OldToNewMergeDict</c>. If <paramref name="newSection2"/> was previously
    ''' recorded as a rename target, the rename is demoted to a merger and its source is folded in.
    ''' 
    ''' </summary>
    ''' <param name="oldSection2">
    ''' The removed entry that was merged
    ''' </param>
    ''' 
    ''' <param name="newSection2">
    ''' The new entry that received content from <paramref name="oldSection2"/>
    ''' </param>
    Private Sub TrackMerger(oldSection2 As iniSection2, newSection2 As iniSection2)

        Dim mergeName = newSection2.Name
        Dim oldName = oldSection2.Name

        SyncLock _state.MergedEntries

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

        End SyncLock

    End Sub

End Class

''' <summary>
''' Result of matching a removed entry against candidates
''' </summary>
Public Class MatchResult

    ''' <summary>
    ''' Whether the match is a rename
    ''' (all keys matched, counts equal, no structural changes)
    ''' </summary>
    Public Property IsRename As Boolean

    ''' <summary>
    ''' Whether the old entry was merged into one or more new entries
    ''' </summary>
    Public Property IsMerge As Boolean

    ''' <summary>
    ''' Whether a best candidate was found but no full merge threshold was met
    ''' </summary>
    Public Property HasPartialMatch As Boolean

    ''' <summary>
    ''' The primary target entry name (rename target or best merge target)
    ''' </summary>
    Public Property TargetName As String

    ''' <summary>
    ''' All target entry names; may contain multiple entries in the case of a split merger
    ''' </summary>
    Public Property AllTargetNames As New List(Of String)

    ''' <summary>
    ''' Number of old keys matched in the best candidate entry
    ''' </summary>
    Public Property TotalMatchedKeys As Integer

    ''' <summary>
    ''' Total number of old FileKeys and RegKeys in the entry being assessed
    ''' </summary>
    Public Property TotalOldKeys As Integer

End Class
