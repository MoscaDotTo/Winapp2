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
''' Adapts <c>MergeDetector</c> for use with <c>iniFile2</c>/<c>iniSection2</c>.
''' Public API takes <c>iniSection2</c>; internal matching logic uses legacy <c>iniSection</c>/<c>winapp2entry</c>
''' after conversion at the boundary via <c>DiffFileBridge</c>.
''' </summary>
Public Class MergeDetector2

    Private ReadOnly _state As DiffState
    Private ReadOnly _diffFile2 As iniFile2
    Private ReadOnly _findModificationsCallback As Action(Of iniSection2, iniSection2)
    Private ReadOnly _newSectionCache As New Dictionary(Of String, iniSection)(StringComparer.OrdinalIgnoreCase)

    ''' <summary>
    ''' Initializes a new instance of <c>MergeDetector2</c>
    ''' </summary>
    ''' <param name="diffState">Shared diff state tracking all entry changes</param>
    ''' <param name="newFile">The new version of winapp2.ini as an <c>iniFile2</c></param>
    ''' <param name="findModsCallback">Callback invoked to track key-level changes when a rename or merger is confirmed</param>
    Public Sub New(diffState As DiffState,
                   newFile As iniFile2,
                   findModsCallback As Action(Of iniSection2, iniSection2))

        _state = diffState
        _diffFile2 = newFile
        _findModificationsCallback = findModsCallback

        For Each section In _diffFile2
            _newSectionCache(section.Name) = DiffFileBridge.ToIniSection(section)
        Next

    End Sub

    ''' <summary>
    ''' Determines if a removed entry was renamed into or merged with one or more added or modified entries.
    ''' Converts <c>iniSection2</c> inputs to <c>iniSection</c> for internal matching logic.
    ''' </summary>
    Public Function AssessRenamesAndMergers(candidates As List(Of iniSection2),
                                            oldSection2 As iniSection2) As Boolean

        If candidates.Count = 0 Then Return False

        Dim oldSection = DiffFileBridge.ToIniSection(oldSection2)
        Dim convertedCandidates As New List(Of iniSection)

        For Each c In candidates
            Dim cached As iniSection = Nothing
            If _newSectionCache.TryGetValue(c.Name, cached) Then convertedCandidates.Add(cached)
        Next

        Dim oldEntry = GetOrCreateCachedEntry(oldSection)
        Dim bestMatch = FindBestMatch(convertedCandidates, oldEntry, oldSection)

        If bestMatch.IsRename Then Return ConfirmRename(bestMatch.TargetName, oldSection2)

        If bestMatch.IsMerge OrElse bestMatch.HasPartialMatch Then

            TrackBestMatches(bestMatch.IsMerge, bestMatch, oldSection)
            Return True

        End If

        Return False

    End Function

    Private Sub TrackBestMatches(isMerge As Boolean, bestMatch As MatchResult, oldSection As iniSection)

        Dim isIntoAddedEntry = _state.ModifiedEntries.AddedEntryNames.Contains(bestMatch.TargetName)

        If Not isMerge Then

            Dim singleTarget As iniSection = Nothing
            If _newSectionCache.TryGetValue(bestMatch.TargetName, singleTarget) Then TrackMerger(oldSection, singleTarget)
            Return

        End If

        For Each targetName In bestMatch.AllTargetNames
            Dim mergeTarget As iniSection = Nothing
            If _newSectionCache.TryGetValue(targetName, mergeTarget) Then TrackMerger(oldSection, mergeTarget)
        Next

    End Sub

    Private Function GetOrCreateCachedEntry(section As iniSection) As winapp2entry

        SyncLock _state.Caches.CachedOldEntries

            If Not _state.Caches.CachedOldEntries.ContainsKey(section.Name) Then _state.Caches.CachedOldEntries.Add(section.Name, New winapp2entry(section))

            Return _state.Caches.CachedOldEntries(section.Name)

        End SyncLock

    End Function

    Private Function FindBestMatch(candidates As List(Of iniSection),
                                   oldEntry As winapp2entry,
                                   oldSection As iniSection) As MatchResult

        Dim result As New MatchResult()
        Dim highestMatchCount = 0
        Dim bestCandidateName = ""
        Dim foundMerger = False
        Dim qualifyingMergeTargets As New List(Of String)

        Dim oldHasFileKeys = oldEntry.FileKeys.KeyCount > 0
        Dim oldHasRegKeys = oldEntry.RegKeys.KeyCount > 0

        Dim totalOldKeys = oldEntry.FileKeys.KeyCount + oldEntry.RegKeys.KeyCount
        result.TotalOldKeys = totalOldKeys

        For Each candidateSection In candidates

            Dim newEntry = GetOrCreateNewCachedEntry(candidateSection)
            Dim matchInfo = GetOrComputeMatchInfo(oldSection.Name, candidateSection.Name, oldEntry, newEntry, oldHasFileKeys, oldHasRegKeys)

            If matchInfo.TotalMatches > highestMatchCount Then
                highestMatchCount = matchInfo.TotalMatches
                bestCandidateName = candidateSection.Name
            End If

            If matchInfo.FileKeyMatches = 0 AndAlso matchInfo.RegKeyMatches = 0 Then Continue For

            Dim thisSpecificPairIsRename As Boolean
            SyncLock _state.MergedEntries

                thisSpecificPairIsRename = _state.MergedEntries.RenamedEntryNames.Contains(candidateSection.Name) AndAlso
                                           IsRenamedFrom(candidateSection.Name, oldSection.Name)

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

    Private Function IsRenamedFrom(newName As String, oldName As String) As Boolean

        Dim idx = _state.MergedEntries.RenamedEntryPairs.IndexOf(newName)
        If idx < 0 OrElse idx + 1 >= _state.MergedEntries.RenamedEntryPairs.Count Then Return False
        Return _state.MergedEntries.RenamedEntryPairs(idx + 1).Equals(oldName, StringComparison.InvariantCultureIgnoreCase)

    End Function

    Private Function GetOrComputeMatchInfo(oldName As String,
                                           newName As String,
                                           oldEntry As winapp2entry,
                                           newEntry As winapp2entry,
                                           oldHasFileKeys As Boolean,
                                           oldHasRegKeys As Boolean) As KeyMatchInfo

        Dim cacheKey = $"{oldName}|{newName}"
        Dim cachedResult As KeyMatchInfo = Nothing
        If _state.Caches.MatchInfoCache.TryGetValue(cacheKey, cachedResult) Then Return cachedResult

        Dim matchInfo = AssessKeyMatches(oldEntry, newEntry, oldHasFileKeys, oldHasRegKeys)
        _state.Caches.MatchInfoCache.TryAdd(cacheKey, matchInfo)
        Return matchInfo

    End Function

    Private Function GetOrCreateNewCachedEntry(section As iniSection) As winapp2entry

        SyncLock _state.Caches.CachedNewEntries

            If Not _state.Caches.CachedNewEntries.ContainsKey(section.Name) Then _state.Caches.CachedNewEntries.Add(section.Name, New winapp2entry(section))

            Return _state.Caches.CachedNewEntries(section.Name)

        End SyncLock

    End Function

    Private Function AssessKeyMatches(oldEntry As winapp2entry,
                                      newEntry As winapp2entry,
                                      oldHasFileKeys As Boolean,
                                      oldHasRegKeys As Boolean) As KeyMatchInfo

        Dim info As New KeyMatchInfo()

        If oldHasFileKeys Then

            info.FileKeyMatches = CountMatches(oldEntry.FileKeys, newEntry.FileKeys, DisallowedPaths,
                                               info.MatchHadMoreParams, info.PossibleWildCardReduction,
                                               info.MatchedOldFileKeys)

            info.AllFileKeysMatched = info.FileKeyMatches = oldEntry.FileKeys.KeyCount
            info.FileKeyCountsMatch = info.AllFileKeysMatched AndAlso newEntry.FileKeys.KeyCount = oldEntry.FileKeys.KeyCount

        Else

            info.AllFileKeysMatched = True
            info.FileKeyCountsMatch = True

        End If

        If oldHasRegKeys Then

            info.RegKeyMatches = CountMatches(oldEntry.RegKeys, newEntry.RegKeys, DisallowedPaths, info.MatchHadMoreParams,
                                              info.PossibleWildCardReduction, info.MatchedOldRegKeys)

            info.AllRegKeysMatched = info.RegKeyMatches = oldEntry.RegKeys.KeyCount
            info.RegKeyCountsMatch = info.AllRegKeysMatched AndAlso newEntry.RegKeys.KeyCount = oldEntry.RegKeys.KeyCount

        Else

            info.AllRegKeysMatched = True
            info.RegKeyCountsMatch = True

        End If

        info.TotalMatches = info.FileKeyMatches + info.RegKeyMatches
        info.AllKeysMatched = info.AllFileKeysMatched AndAlso info.AllRegKeysMatched
        info.CountsMatch = info.FileKeyCountsMatch AndAlso info.RegKeyCountsMatch

        Return info

    End Function

    Private Function CountMatches(oldKeys As keyList,
                                  newKeys As keyList,
                                  disallowedValues As HashSet(Of String),
                            ByRef matchHadMoreParams As Boolean,
                            ByRef possibleWildCardReduction As Boolean,
                                  matchedKeys As HashSet(Of iniKey)) As Integer

        Dim matchCount = 0
        Dim newKeyValues As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        For Each newKey In newKeys.Keys
            If Not newKeyValues.Contains(newKey.Value) Then newKeyValues.Add(newKey.Value)
        Next

        For Each oldKey In oldKeys.Keys

            If disallowedValues IsNot Nothing AndAlso disallowedValues.Contains(oldKey.Value) Then Continue For

            Dim matched = False

            If newKeyValues.Contains(oldKey.Value) Then

                matched = True

            Else

                For Each newKey In newKeys.Keys

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

    Private Function GetPathWithoutFlags(value As String) As String
        Return If(value.Contains("|"), value.Substring(0, value.IndexOf("|", StringComparison.InvariantCultureIgnoreCase)), value)
    End Function

    Private Function ConfirmRename(newName As String, oldSection2 As iniSection2) As Boolean

        Dim newSection2 As iniSection2 = Nothing

        SyncLock _state.MergedEntries

            Dim ind = _state.MergedEntries.RenamedEntryPairs.IndexOf(newName)

            If ind >= 0 Then

                Return ind + 1 < _state.MergedEntries.RenamedEntryPairs.Count AndAlso
                       _state.MergedEntries.RenamedEntryPairs(ind + 1).Equals(oldSection2.Name, StringComparison.InvariantCultureIgnoreCase)

            End If

            _state.MergedEntries.RenamedEntryNames.Add(newName)
            _state.MergedEntries.RenamedEntryPairs.Add(newName)
            _state.MergedEntries.RenamedEntryPairs.Add(oldSection2.Name)
            newSection2 = _diffFile2.GetSection(newName)

        End SyncLock

        If _findModificationsCallback IsNot Nothing AndAlso newSection2 IsNot Nothing Then _findModificationsCallback(oldSection2, newSection2)

        Return True

    End Function

    Private Sub TrackMerger(oldSection As iniSection, newSection As iniSection)

        Dim mergeName = newSection.Name
        Dim oldName = oldSection.Name

        SyncLock _state.MergedEntries

            _state.MergedEntries.MergedEntryNames.Add(mergeName)

            If Not _state.MergedEntries.MergeDict.ContainsKey(mergeName) Then _state.MergedEntries.MergeDict.Add(mergeName, New List(Of String))

            If Not _state.MergedEntries.MergeDict(mergeName).Contains(oldName) Then _state.MergedEntries.MergeDict(mergeName).Add(oldName)

            If Not _state.MergedEntries.OldToNewMergeDict.ContainsKey(oldName) Then _state.MergedEntries.OldToNewMergeDict.Add(oldName, New List(Of String))

            If Not _state.MergedEntries.OldToNewMergeDict(oldName).Contains(mergeName) Then _state.MergedEntries.OldToNewMergeDict(oldName).Add(mergeName)

            If Not _state.MergedEntries.RenamedEntryNames.Contains(mergeName) Then Return

            Dim ind = _state.MergedEntries.RenamedEntryPairs.IndexOf(mergeName)
            If ind < 0 OrElse ind + 1 >= _state.MergedEntries.RenamedEntryPairs.Count Then Return

            Dim renameHolder = _state.MergedEntries.RenamedEntryPairs(ind + 1)

            If Not _state.MergedEntries.MergeDict(mergeName).Contains(renameHolder) Then _state.MergedEntries.MergeDict(mergeName).Add(renameHolder)

            If Not _state.MergedEntries.OldToNewMergeDict.ContainsKey(renameHolder) Then _state.MergedEntries.OldToNewMergeDict.Add(renameHolder, New List(Of String))

            If Not _state.MergedEntries.OldToNewMergeDict(renameHolder).Contains(mergeName) Then _state.MergedEntries.OldToNewMergeDict(renameHolder).Add(mergeName)

            _state.MergedEntries.RenamedEntryPairs.RemoveAt(ind + 1)
            _state.MergedEntries.RenamedEntryPairs.RemoveAt(ind)
            _state.MergedEntries.RenamedEntryNames.Remove(mergeName)

        End SyncLock

    End Sub

End Class
