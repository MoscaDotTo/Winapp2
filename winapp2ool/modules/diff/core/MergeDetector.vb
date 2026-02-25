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
''' Handles detection of merged and renamed entries in diff operations
''' </summary>
Public Class MergeDetector

    ''' <summary>
    ''' Shared diff state tracking all entry changes
    ''' </summary>
    Private ReadOnly state As DiffState

    ''' <summary>
    ''' The new version of winapp2.ini
    ''' </summary>
    Private ReadOnly diffFile2 As iniFile

    ''' <summary>
    ''' Callback invoked to track key-level changes when a rename or merger is confirmed
    ''' </summary>
    Private ReadOnly findModificationsCallback As Action(Of iniSection, iniSection)

    ''' <summary>
    ''' Initializes a new instance of <c>MergeDetector</c>
    ''' </summary>
    '''
    ''' <param name="diffState">Shared diff state tracking all entry changes</param>
    ''' <param name="newFile">The new version of winapp2.ini</param>
    ''' <param name="findModsCallback">Callback invoked to track key-level changes when a rename or merger is confirmed</param>
    Public Sub New(diffState As DiffState,
                   newFile As iniFile,
                   findModsCallback As Action(Of iniSection, iniSection))

        state = diffState
        diffFile2 = newFile
        findModificationsCallback = findModsCallback

    End Sub

    ''' <summary>
    ''' Determines if a removed entry was renamed into or merged with one or more added or modified entries
    ''' </summary>
    '''
    ''' <param name="entriesAddedOrModified">
    ''' Candidate entries from the new file to compare against the removed entry
    ''' </param>
    '''
    ''' <param name="oldSectionVersion">
    ''' The removed entry from the old file
    ''' </param>
    '''
    ''' <returns>
    ''' <c>True</c> if the removed entry was identified as a rename or merger; <c>False</c> otherwise
    ''' </returns>
    Public Function AssessRenamesAndMergers(entriesAddedOrModified As List(Of iniSection),
                                            oldSectionVersion As iniSection) As Boolean

        If entriesAddedOrModified.Count = 0 Then Return False

        Dim oldEntry = GetOrCreateCachedEntry(oldSectionVersion)

        Dim bestMatch = FindBestMatch(entriesAddedOrModified, oldEntry, oldSectionVersion)

        If bestMatch.IsRename Then Return ConfirmRename(bestMatch.TargetName, oldSectionVersion)

        If bestMatch.IsMerge OrElse bestMatch.HasPartialMatch Then

            TrackBestMatches(bestMatch.IsMerge, bestMatch, oldSectionVersion)

            Return True

        End If

        Return False

    End Function

    ''' <summary>
    ''' Records the best matching new entry or entries as a merger target for the given old section
    ''' </summary>
    Private Sub TrackBestMatches(isMerge As Boolean,
                                 bestMatch As MatchResult,
                                 oldSectionVersion As iniSection)

        Dim isIntoAddedEntry = state.ModifiedEntries.AddedEntryNames.Contains(bestMatch.TargetName)

        If Not isMerge Then TrackMerger(oldSectionVersion, diffFile2.Sections(bestMatch.TargetName)) : Return

        For Each targetName In bestMatch.AllTargetNames

            TrackMerger(oldSectionVersion, diffFile2.Sections(targetName))

        Next

    End Sub

    ''' <summary>
    ''' Returns a cached <c>winapp2entry</c> for the given old-file section, creating and caching it on first access
    ''' </summary>
    Private Function GetOrCreateCachedEntry(section As iniSection) As winapp2entry

        SyncLock state.Caches.CachedOldEntries

            If Not state.Caches.CachedOldEntries.ContainsKey(section.Name) Then state.Caches.CachedOldEntries.Add(section.Name, New winapp2entry(section))

            Return state.Caches.CachedOldEntries(section.Name)

        End SyncLock

    End Function

    ''' <summary>
    ''' Finds the best matching candidate entry for a removed entry
    ''' </summary>
    ''' 
    ''' <param name="candidates">
    ''' The set of candidate entries to compare against the old entry
    ''' </param>
    ''' 
    ''' <param name="oldEntry">
    ''' The old entry to find a match for
    ''' </param>
    ''' 
    ''' <param name="oldSection">
    ''' The old section corresponding to the old entry
    ''' </param>
    ''' 
    ''' <returns>
    ''' The best matching candidate entry for the removed entry
    ''' </returns>
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
        result.TotalOldKeys = totalOldKeys  ' ← Set early, used for all result types


        For Each candidateSection In candidates

            Dim newEntry = GetOrCreateNewCachedEntry(candidateSection)
            Dim matchInfo = GetOrComputeMatchInfo(oldSection.Name, candidateSection.Name, oldEntry, newEntry, oldHasFileKeys, oldHasRegKeys)

            If matchInfo.TotalMatches > highestMatchCount Then

                highestMatchCount = matchInfo.TotalMatches
                bestCandidateName = candidateSection.Name

            End If

            If matchInfo.FileKeyMatches = 0 AndAlso matchInfo.RegKeyMatches = 0 Then Continue For

            ' Only skip if this SPECIFIC old->new pair was already processed as a rename
            Dim thisSpecificPairIsRename As Boolean
            SyncLock state.MergedEntries

                thisSpecificPairIsRename = state.MergedEntries.RenamedEntryNames.Contains(candidateSection.Name) AndAlso
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


    ''' <summary>
    ''' Determines if newName was renamed from oldName
    ''' </summary>
    ''' 
    ''' <param name="newName">
    ''' The new entry name to look up in the renamed pairs list
    ''' </param>
    '''
    ''' <param name="oldName">
    ''' The old entry name to check as the source of the rename
    ''' </param>
    '''
    ''' <returns>
    ''' <c>True</c> if <paramref name="newName"/> was renamed from <paramref name="oldName"/>
    ''' </returns>
    '''
    ''' <remarks>
    ''' Must be called within SyncLock state.MergedEntries
    ''' </remarks>
    Private Function IsRenamedFrom(newName As String, oldName As String) As Boolean

        ' 
        Dim idx = state.MergedEntries.RenamedEntryPairs.IndexOf(newName)
        If idx < 0 OrElse idx + 1 >= state.MergedEntries.RenamedEntryPairs.Count Then Return False
        Return state.MergedEntries.RenamedEntryPairs(idx + 1).Equals(oldName, StringComparison.InvariantCultureIgnoreCase)

    End Function

    ''' <summary>
    ''' Gets or computes key match information between two entries
    ''' </summary>
    '''
    ''' <param name="oldName">Name of the old entry, used as part of the match info cache key</param>
    ''' <param name="newName">Name of the new entry, used as part of the match info cache key</param>
    ''' <param name="oldEntry">The old entry being compared</param>
    ''' <param name="newEntry">The new entry being compared against</param>
    ''' <param name="oldHasFileKeys">Whether the old entry has any FileKeys</param>
    ''' <param name="oldHasRegKeys">Whether the old entry has any RegKeys</param>
    '''
    ''' <returns>
    ''' A <c>KeyMatchInfo</c> describing how many and which keys matched between the two entries
    ''' </returns>
    Private Function GetOrComputeMatchInfo(oldName As String,
                                          newName As String,
                                          oldEntry As winapp2entry,
                                          newEntry As winapp2entry,
                                          oldHasFileKeys As Boolean,
                                          oldHasRegKeys As Boolean) As KeyMatchInfo

        Dim cacheKey = $"{oldName}|{newName}"
        Dim cachedResult As KeyMatchInfo = Nothing
        If state.Caches.MatchInfoCache.TryGetValue(cacheKey, cachedResult) Then Return cachedResult


        Dim matchInfo = AssessKeyMatches(oldEntry, newEntry, oldHasFileKeys, oldHasRegKeys)
        state.Caches.MatchInfoCache.TryAdd(cacheKey, matchInfo)
        Return matchInfo

    End Function

    ''' <summary>
    ''' Returns a cached <c>winapp2entry</c> for the given new-file section, creating and caching it on first access
    ''' </summary>
    Private Function GetOrCreateNewCachedEntry(section As iniSection) As winapp2entry

        SyncLock state.Caches.CachedNewEntries

            If Not state.Caches.CachedNewEntries.ContainsKey(section.Name) Then state.Caches.CachedNewEntries.Add(section.Name, New winapp2entry(section))

            Return state.Caches.CachedNewEntries(section.Name)

        End SyncLock

    End Function

    ''' <summary>
    ''' Computes the full <c>KeyMatchInfo</c> for two entries by counting FileKey and RegKey matches
    ''' </summary>
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

    ''' <summary>
    ''' Counts how many old keys are matched by keys in the new key list, skipping disallowed path values
    ''' </summary>
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

    ''' <summary>
    ''' Returns the path portion of a key value before the first pipe character, or the full value if no pipe is present
    ''' </summary>
    Private Function GetPathWithoutFlags(value As String) As String

        Return If(value.Contains("|"), value.Substring(0, value.IndexOf("|", StringComparison.InvariantCultureIgnoreCase)), value)

    End Function

    ''' <summary>
    ''' Attempts to confirm and record a rename from <paramref name="oldSection"/> to <paramref name="newName"/>
    ''' </summary>
    Private Function ConfirmRename(newName As String,
                                   oldSection As iniSection) As Boolean

        Dim newSection As iniSection = Nothing

        SyncLock state.MergedEntries

            Dim ind = state.MergedEntries.RenamedEntryPairs.IndexOf(newName)
            If ind >= 0 Then

                Return ind + 1 < state.MergedEntries.RenamedEntryPairs.Count AndAlso
                       state.MergedEntries.RenamedEntryPairs(ind + 1).Equals(oldSection.Name, StringComparison.InvariantCultureIgnoreCase)

            End If

            state.MergedEntries.RenamedEntryNames.Add(newName)
            state.MergedEntries.RenamedEntryPairs.Add(newName)
            state.MergedEntries.RenamedEntryPairs.Add(oldSection.Name)
            newSection = diffFile2.Sections(newName)

        End SyncLock

        If findModificationsCallback IsNot Nothing Then findModificationsCallback(oldSection, newSection)

        Return True

    End Function

    ''' <summary>
    ''' Records a merger between <paramref name="oldSection"/> and <paramref name="newSection"/> in the diff state
    ''' </summary>
    Private Sub TrackMerger(oldSection As iniSection, newSection As iniSection)
        Dim mergeName = newSection.Name
        Dim oldName = oldSection.Name

        SyncLock state.MergedEntries

            state.MergedEntries.MergedEntryNames.Add(mergeName)

            If Not state.MergedEntries.MergeDict.ContainsKey(mergeName) Then state.MergedEntries.MergeDict.Add(mergeName, New List(Of String))

            If Not state.MergedEntries.MergeDict(mergeName).Contains(oldName) Then state.MergedEntries.MergeDict(mergeName).Add(oldName)

            If Not state.MergedEntries.OldToNewMergeDict.ContainsKey(oldName) Then state.MergedEntries.OldToNewMergeDict.Add(oldName, New List(Of String))

            If Not state.MergedEntries.OldToNewMergeDict(oldName).Contains(mergeName) Then state.MergedEntries.OldToNewMergeDict(oldName).Add(mergeName)

            If Not state.MergedEntries.RenamedEntryNames.Contains(mergeName) Then Return

            Dim ind = state.MergedEntries.RenamedEntryPairs.IndexOf(mergeName)
            If ind < 0 OrElse ind + 1 >= state.MergedEntries.RenamedEntryPairs.Count Then Return

            Dim renameHolder = state.MergedEntries.RenamedEntryPairs(ind + 1)

            If Not state.MergedEntries.MergeDict(mergeName).Contains(renameHolder) Then state.MergedEntries.MergeDict(mergeName).Add(renameHolder)


            If Not state.MergedEntries.OldToNewMergeDict.ContainsKey(renameHolder) Then state.MergedEntries.OldToNewMergeDict.Add(renameHolder, New List(Of String))

            If Not state.MergedEntries.OldToNewMergeDict(renameHolder).Contains(mergeName) Then state.MergedEntries.OldToNewMergeDict(renameHolder).Add(mergeName)

            state.MergedEntries.RenamedEntryPairs.RemoveAt(ind + 1)
            state.MergedEntries.RenamedEntryPairs.RemoveAt(ind)
            state.MergedEntries.RenamedEntryNames.Remove(mergeName)

        End SyncLock

    End Sub

End Class

''' <summary>
''' Result of matching a removed entry against candidates
''' </summary>
Public Class MatchResult

    ''' <summary>Whether the match is a rename (all keys matched, counts equal, no structural changes)</summary>
    Public Property IsRename As Boolean

    ''' <summary>Whether the old entry was merged into one or more new entries</summary>
    Public Property IsMerge As Boolean

    ''' <summary>Whether a best candidate was found but no full merge threshold was met</summary>
    Public Property HasPartialMatch As Boolean

    ''' <summary>The primary target entry name (rename target or best merge target)</summary>
    Public Property TargetName As String

    ''' <summary>All target entry names; may contain multiple entries in the case of a split merger</summary>
    Public Property AllTargetNames As New List(Of String)

    ''' <summary>Number of old keys matched in the best candidate entry</summary>
    Public Property TotalMatchedKeys As Integer

    ''' <summary>Total number of old FileKeys and RegKeys in the entry being assessed</summary>
    Public Property TotalOldKeys As Integer

End Class

''' <summary>
''' Information about key matches between two entries
''' </summary>
Public Class KeyMatchInfo

    ''' <summary>
    ''' Number of FileKey values from the old entry matched in the new entry
    ''' </summary>
    Public Property FileKeyMatches As Integer

    ''' <summary>
    ''' Number of RegKey values from the old entry matched in the new entry
    ''' </summary>
    Public Property RegKeyMatches As Integer

    ''' <summary>
    ''' Number of Detect values from the old entry matched in the new entry
    ''' </summary>
    Public Property DetectMatches As Integer

    ''' <summary>
    ''' Number of DetectFile values from the old entry matched in the new entry
    ''' </summary>
    Public Property DetectFileMatches As Integer

    ''' <summary>
    ''' Number of other key values (Section, LangSecRef, DetectOS, Warning, etc.) 
    ''' from the old entry matched in the new entry
    ''' </summary>
    Public Property OtherKeyMatches As Integer

    ''' <summary>
    ''' Sum of all key type match counts
    ''' </summary>
    Public Property TotalMatches As Integer

    ''' <summary>
    ''' Whether all FileKeys from the old entry were matched in the new entry
    ''' </summary>
    Public Property AllFileKeysMatched As Boolean

    ''' <summary>
    ''' Whether all RegKeys from the old entry were matched in the new entry
    ''' </summary>
    Public Property AllRegKeysMatched As Boolean

    ''' <summary>
    ''' Whether all Detect keys from the old entry were matched in the new entry
    ''' </summary>
    Public Property AllDetectsMatched As Boolean

    ''' <summary>
    ''' Whether all DetectFile keys from the old entry were matched in the new entry
    ''' </summary>
    Public Property AllDetectFilesMatched As Boolean

    ''' <summary>
    ''' Whether all other keys from the old entry were matched in the new entry
    ''' </summary>
    Public Property AllOtherKeysMatched As Boolean

    ''' <summary>
    ''' Whether every key from the old entry was matched in the new entry across all key types
    ''' </summary>
    Public Property AllKeysMatched As Boolean

    ''' <summary>
    ''' Whether the count of FileKeys is the same in both old and new entries
    ''' </summary>
    Public Property FileKeyCountsMatch As Boolean

    ''' <summary>
    ''' Whether the count of RegKeys is the same in both old and new entries
    ''' </summary>
    Public Property RegKeyCountsMatch As Boolean

    ''' <summary>
    ''' Whether FileKey and RegKey counts both match between old and new entries
    ''' </summary>
    Public Property CountsMatch As Boolean

    ''' <summary>
    ''' Whether any matched new key has more pipe-delimited parameters than its old counterpart
    ''' </summary>
    Public Property MatchHadMoreParams As Boolean

    ''' <summary>
    ''' Whether any matched key appears to have reduced wildcard specificity
    ''' </summary>
    Public Property PossibleWildCardReduction As Boolean

    ''' <summary>
    ''' Set of old FileKey objects that were matched in the new entry
    ''' </summary>
    Public Property MatchedOldFileKeys As New HashSet(Of iniKey)

    ''' <summary>
    ''' Set of old RegKey objects that were matched in the new entry
    ''' </summary>
    Public Property MatchedOldRegKeys As New HashSet(Of iniKey)

    ''' <summary>
    ''' Set of old Detect key objects that were matched in the new entry
    ''' </summary>
    Public Property MatchedOldDetects As New HashSet(Of iniKey)

    ''' <summary>
    ''' Set of old DetectFile key objects that were matched in the new entry
    ''' </summary>
    Public Property MatchedOldDetectFiles As New HashSet(Of iniKey)

    ''' <summary>
    ''' Set of old other key objects that were matched in the new entry
    ''' </summary>
    Public Property MatchedOldOtherKeys As New HashSet(Of iniKey)

    ''' <summary>
    ''' Set of new FileKey objects that captured at least one old key
    ''' </summary>
    Public Property MatchedNewFileKeys As New HashSet(Of iniKey)

    ''' <summary>
    ''' Set of new RegKey objects that captured at least one old key
    ''' </summary>
    Public Property MatchedNewRegKeys As New HashSet(Of iniKey)

    ''' <summary>
    ''' Set of new Detect key objects that captured at least one old key
    ''' </summary>
    Public Property MatchedNewDetects As New HashSet(Of iniKey)

    ''' <summary>
    ''' Set of new DetectFile key objects that captured at least one old key
    ''' </summary>
    Public Property MatchedNewDetectFiles As New HashSet(Of iniKey)

    ''' <summary>
    ''' Set of new other key objects that captured at least one old key
    ''' </summary>
    Public Property MatchedNewOtherKeys As New HashSet(Of iniKey)

    ''' <summary>
    ''' Total OLD keys that were matched across all key types
    ''' </summary>
    Public ReadOnly Property TotalOldKeysMatched As Integer

        Get

            Return MatchedOldFileKeys.Count + MatchedOldRegKeys.Count + MatchedOldDetects.Count +
                   MatchedOldDetectFiles.Count + MatchedOldOtherKeys.Count

        End Get

    End Property

    ''' <summary>
    ''' Total NEW keys that captured at least one old key
    ''' </summary>
    Public ReadOnly Property TotalNewKeysCapturing As Integer

        Get

            Return MatchedNewFileKeys.Count + MatchedNewRegKeys.Count + MatchedNewDetects.Count +
                   MatchedNewDetectFiles.Count + MatchedNewOtherKeys.Count

        End Get

    End Property

End Class
