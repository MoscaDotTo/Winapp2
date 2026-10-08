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

Imports System.Collections.Concurrent

''' <summary>
''' Holds the shared, mutable state of one Diff run: the entry and key trackers, counters,
''' caches and key movements that the core classes write to and read from
''' </summary>
Public Class DiffState

    ''' <summary>
    ''' Tracks entries that have been merged or renamed
    ''' </summary>
    Public Property MergedEntries As New MergedEntryTracker()

    ''' <summary>
    ''' Tracks entries that have been modified and the specific key changes
    ''' </summary>
    Public Property ModifiedEntries As New ModifiedEntryTracker()

    ''' <summary>
    ''' Tracks statistical counters for the diff operation
    ''' </summary>
    Public Property Statistics As New DiffStatistics()

    ''' <summary>
    ''' Manages caches for performance optimization during diff operations
    ''' </summary>
    Public Property Caches As New DiffCaches()

    ''' <summary>
    ''' Tracks keys that moved between entries
    ''' </summary>
    Public Property KeyMovements As New KeyMovementTracker()

    ''' <summary>
    ''' Clears every tracker and cache and resets the counters. <see cref="DiffStatistics.Reset"/>
    ''' leaves some counters alone.
    ''' </summary>
    Public Sub Clear()

        MergedEntries.Clear()
        ModifiedEntries.Clear()
        Statistics.Reset()
        Caches.Clear()
        KeyMovements.Clear()

    End Sub

End Class

''' <summary>
''' Tracks entries that have been merged or renamed
''' </summary>
Public Class MergedEntryTracker

    ''' <summary>
    ''' Names of the new-file entries that received content from one or more removed entries
    ''' </summary>
    '''
    ''' <remarks>
    ''' Not thread-safe. The removal pass writes this tracker from one thread, after its parallel
    ''' matching finishes.
    ''' </remarks>
    Public Property MergedEntryNames As New HashSet(Of String)

    ''' <summary>
    ''' Tracks merge mappings: NewEntryName -> List(Of OldEntryNames)
    ''' </summary>
    Public Property MergeDict As New Dictionary(Of String, List(Of String))

    ''' <summary>
    ''' Tracks merge mappings: OldEntryName -> List(Of NewEntryNames)
    ''' </summary>
    Public Property OldToNewMergeDict As New Dictionary(Of String, List(Of String))

    ''' <summary>
    ''' New names of the entries recognized as renames of a removed entry. A rename whose target
    ''' later also takes in merged content is moved out of here and recorded as a merger.
    ''' </summary>
    Public Property RenamedEntryNames As New HashSet(Of String)

    ''' <summary>
    ''' Maps each renamed entry's new name to its old name.
    ''' </summary>
    Public Property RenamedEntryPairs As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

    ''' <summary>
    ''' Sorts every <see cref="MergeDict"/> and <see cref="OldToNewMergeDict"/> list by name,
    ''' ignoring case. The output lists merge sources in this order, and a key that several
    ''' sources share credits them in this order.
    ''' </summary>
    Public Sub SortMergeLists()

        For Each names In MergeDict.Values : names.Sort(StringComparer.OrdinalIgnoreCase) : Next
        For Each names In OldToNewMergeDict.Values : names.Sort(StringComparer.OrdinalIgnoreCase) : Next

    End Sub

    ''' <summary>
    ''' Clears all tracking data
    ''' </summary>
    Public Sub Clear()

        MergedEntryNames.Clear()
        MergeDict.Clear()
        OldToNewMergeDict.Clear()
        RenamedEntryNames.Clear()
        RenamedEntryPairs.Clear()

    End Sub

End Class

''' <summary>
''' Tracks which entries were added, removed and modified, and the key changes found for each.
''' The three key trackers are keyed by the entry's new name and hold results for every entry
''' Diff compared, including rename and merge targets, not only modified entries.
''' </summary>
'''
''' <remarks>
''' Not thread-safe. Every write to it happens on one thread.
''' </remarks>
Public Class ModifiedEntryTracker

    ''' <summary>
    ''' Names of entries present in both files whose keys changed
    ''' </summary>
    Public Property ModifiedEntryNames As New HashSet(Of String)

    ''' <summary>
    ''' Names of entries present only in the new file, including rename and merge targets
    ''' </summary>
    Public Property AddedEntryNames As New HashSet(Of String)

    ''' <summary>
    ''' Names of entries present only in the old file, including those later found to be
    ''' renamed or merged
    ''' </summary>
    Public Property RemovedEntryNames As New HashSet(Of String)

    ''' <summary>
    ''' Updated keys per entry: EntryName -> (NewKey -> List(Of OldKeys)). A rename adds a
    ''' synthetic <c> Name </c> key pair holding the new and old entry names.
    ''' </summary>
    Public Property ModifiedKeyTracker As New Dictionary(Of String, Dictionary(Of iniKey, List(Of iniKey)))

    ''' <summary>
    ''' Tracks removed keys per entry: EntryName -> List(Of Keys)
    ''' </summary>
    Public Property RemovedKeyTracker As New Dictionary(Of String, List(Of iniKey))

    ''' <summary>
    ''' Tracks added keys per entry: EntryName -> List(Of Keys)
    ''' </summary>
    Public Property AddedKeyTracker As New Dictionary(Of String, List(Of iniKey))

    ''' <summary>
    ''' The added and modified new-file sections that removed entries are matched against when
    ''' looking for renames and mergers
    ''' </summary>
    Public Property PotentialMatches As New List(Of iniSection)

    ''' <summary>
    ''' Clears all tracking data
    ''' </summary>
    Public Sub Clear()

        ModifiedEntryNames.Clear()
        AddedEntryNames.Clear()
        RemovedEntryNames.Clear()
        ModifiedKeyTracker.Clear()
        RemovedKeyTracker.Clear()
        AddedKeyTracker.Clear()
        PotentialMatches.Clear()

    End Sub

End Class

''' <summary>
''' Tracks statistical counters for the diff operation
''' </summary>
Public Class DiffStatistics

    ''' <summary>
    ''' Counts total keys added in modified entries, minus every key move that
    ''' <see cref="DiffStatisticsCalculator.DetectCrossEntryMovements"/> finds, including moves
    ''' that involve an entry which isn't modified, so it can go negative
    ''' </summary>
    Public Property ModEntriesAddedKeyTotal As Integer = 0

    ''' <summary>
    ''' Counts modified entries that have at least one added key
    ''' </summary>
    Public Property ModEntriesAddedKeyEntryCount As Integer = 0

    ''' <summary>
    ''' Counts modified entries that have at least one removed key
    ''' </summary>
    Public Property ModEntriesRemovedKeyEntryCount As Integer = 0

    ''' <summary>
    ''' Counts total keys updated in modified entries
    ''' </summary>
    Public Property ModEntriesUpdatedKeyTotal As Integer = 0

    ''' <summary>
    ''' Counts total old keys in modified entries that an updated key replaced
    ''' </summary>
    Public Property ModEntriesReplacedByUpdateTotal As Integer = 0

    ''' <summary>
    ''' Counts total keys removed without replacement from modified entries, minus every key
    ''' move that <see cref="DiffStatisticsCalculator.DetectCrossEntryMovements"/> finds,
    ''' including moves that involve an entry which isn't modified, so it can go negative
    ''' </summary>
    Public Property ModEntriesRemovedKeysWithoutReplacementTotal As Integer = 0

    ''' <summary>
    ''' Counts total keys that moved between entries
    ''' </summary>
    Public Property ModEntriesMovedKeysTotal As Integer = 0

    ''' <summary>
    ''' Counts the distinct entries that lost a key to another entry
    ''' </summary>
    Public Property ModEntriesMovedKeysSourceCount As Integer = 0

    ''' <summary>
    ''' Counts the distinct entries that gained a key from another entry
    ''' </summary>
    Public Property ModEntriesMovedKeysTargetCount As Integer = 0

    ''' <summary>
    ''' Counts modified entries that have at least one updated key
    ''' </summary>
    Public Property ModEntriesUpdatedKeyEntryCount As Integer = 0

    ''' <summary>
    ''' Counts added entries, other than renames, that had one or more removed entries merged into them
    ''' </summary>
    Public Property AddedWithMergersEntryCount As Integer = 0

    ''' <summary>
    ''' Counts the distinct removed entries merged into those added entries
    ''' </summary>
    Public Property AddedWithMergersSourceEntryCount As Integer = 0

    ''' <summary>
    ''' Counts total novel keys in added-with-merger entries: added keys whose value matches no key
    ''' value in the entry's merged sources
    ''' </summary>
    Public Property AddedWithMergersNovelKeysTotal As Integer = 0

    ''' <summary>
    ''' Counts added-with-merger entries that contain at least one novel key (a key not from any merged source)
    ''' </summary>
    Public Property AddedWithMergersNovelKeysEntryCount As Integer = 0

    ''' <summary>
    ''' Counts total keys in added-with-merger entries that updated, rather than copied, one or more
    ''' keys from their merged sources. Exact copies count as carried over instead.
    ''' </summary>
    Public Property AddedWithMergersCapturingKeysTotal As Integer = 0

    ''' <summary>
    ''' Counts the distinct FileKey and RegKey values from merged source entries that reappear,
    ''' exactly or as captured by another key, in any entry they were merged into. Only sources
    ''' merged into at least one added entry are counted.
    ''' </summary>
    Public Property AddedWithMergersCapturedKeysTotal As Integer = 0

    ''' <summary>
    ''' Counts added-with-merger entries that contain at least one capturing key
    ''' </summary>
    Public Property AddedWithMergersCapturingEntryCount As Integer = 0

    ''' <summary>
    ''' Counts the distinct FileKey and RegKey values from the same merged source entries that no
    ''' merge target captured
    ''' </summary>
    Public Property AddedWithMergersDroppedKeysTotal As Integer = 0

    ''' <summary>
    ''' Counts added-with-merger entries that dropped at least one key from a merged source entry
    ''' </summary>
    Public Property AddedWithMergersDroppedEntryCount As Integer = 0

    ''' <summary>
    ''' Counts total keys from merged source entries that were carried over unchanged into the added-with-merger entry
    ''' </summary>
    Public Property AddedWithMergersCarriedOverKeysTotal As Integer = 0

    ''' <summary>
    ''' Counts added-with-merger entries that contain at least one key carried over unchanged from a merged source entry
    ''' </summary>
    Public Property AddedWithMergersCarriedOverKeysEntryCount As Integer = 0

    ''' <summary>
    ''' Counts total keys added in renamed entries
    ''' </summary>
    Public Property RenamedEntriesAddedKeyTotal As Integer = 0

    ''' <summary>
    ''' Counts renamed entries that have at least one added key
    ''' </summary>
    Public Property RenamedEntriesAddedKeyEntryCount As Integer = 0

    ''' <summary>
    ''' Counts total keys removed in renamed entries
    ''' </summary>
    Public Property RenamedEntriesRemovedKeyTotal As Integer = 0

    ''' <summary>
    ''' Counts renamed entries that have at least one removed key
    ''' </summary>
    Public Property RenamedEntriesRemovedKeyEntryCount As Integer = 0

    ''' <summary>
    ''' Counts total keys updated in renamed entries
    ''' </summary>
    Public Property RenamedEntriesUpdatedKeyTotal As Integer = 0

    ''' <summary>
    ''' Counts total old keys replaced by updates in renamed entries
    ''' </summary>
    Public Property RenamedEntriesReplacedByUpdateTotal As Integer = 0

    ''' <summary>
    ''' Counts renamed entries that have at least one updated key
    ''' </summary>
    Public Property RenamedEntriesUpdatedKeyEntryCount As Integer = 0

    ''' <summary>
    ''' Counts renamed entries that are name-only changes (no key differences)
    ''' </summary>
    Public Property RenamedEntriesNameOnlyCount As Integer = 0

    ''' <summary>
    ''' Section key values (e.g. <c> "Brave Web Browser" </c>) that appear in the new file
    ''' but not in the old file, indicating newly added browser support.
    ''' Populated by <see cref="DiffStatisticsCalculator.DetectNewBrowserSupport"/>.
    ''' </summary>
    Public Property NewBrowserSectionValues As New List(Of String)

    ''' <summary>
    ''' Section key values (e.g. <c> "Internet Explorer" </c>) that appear in the old file
    ''' but not in the new file, indicating removed browser support.
    ''' Populated by <see cref="DiffStatisticsCalculator.DetectNewBrowserSupport"/>.
    ''' </summary>
    Public Property RemovedBrowserSectionValues As New List(Of String)

    ''' <summary>
    ''' Resets the modified-entry, moved-key, added-with-mergers, and renamed-entry counters to zero and clears both
    ''' browser lists.
    ''' </summary>
    Public Sub Reset()

        AddedWithMergersCapturedKeysTotal = 0
        AddedWithMergersCapturingEntryCount = 0
        AddedWithMergersCapturingKeysTotal = 0
        AddedWithMergersCarriedOverKeysEntryCount = 0
        AddedWithMergersCarriedOverKeysTotal = 0
        AddedWithMergersDroppedEntryCount = 0
        AddedWithMergersDroppedKeysTotal = 0
        ModEntriesAddedKeyTotal = 0
        ModEntriesAddedKeyEntryCount = 0
        ModEntriesRemovedKeyEntryCount = 0
        ModEntriesUpdatedKeyTotal = 0
        ModEntriesReplacedByUpdateTotal = 0
        ModEntriesRemovedKeysWithoutReplacementTotal = 0
        ModEntriesMovedKeysTotal = 0
        ModEntriesMovedKeysSourceCount = 0
        ModEntriesMovedKeysTargetCount = 0
        ModEntriesUpdatedKeyEntryCount = 0
        RenamedEntriesAddedKeyTotal = 0
        RenamedEntriesAddedKeyEntryCount = 0
        RenamedEntriesRemovedKeyTotal = 0
        RenamedEntriesRemovedKeyEntryCount = 0
        RenamedEntriesUpdatedKeyTotal = 0
        RenamedEntriesReplacedByUpdateTotal = 0
        RenamedEntriesUpdatedKeyEntryCount = 0
        RenamedEntriesNameOnlyCount = 0
        NewBrowserSectionValues.Clear()
        RemovedBrowserSectionValues.Clear()

    End Sub

End Class

''' <summary>
''' Manages caches for performance optimization during diff operations
''' </summary>
Public Class DiffCaches

    ''' <summary>
    ''' Caches old entries by name for quick lookup
    ''' </summary>
    Public Property CachedOldEntries As New Dictionary(Of String, iniSection)

    ''' <summary>
    ''' Caches new entries by name for quick lookup
    ''' </summary>
    Public Property CachedNewEntries As New Dictionary(Of String, iniSection)

    ''' <summary>
    ''' Caches the key match assessment for each old/new entry pair, keyed <c> oldName|newName </c>
    ''' </summary>
    Public Property MatchInfoCache As New ConcurrentDictionary(Of String, KeyMatchInfo)

    ''' <summary>
    ''' Clears all caches
    ''' </summary>
    Public Sub Clear()

        CachedOldEntries.Clear()
        CachedNewEntries.Clear()
        MatchInfoCache.Clear()

    End Sub

End Class

''' <summary>
''' Tracks keys that moved between entries
''' </summary>
Public Class KeyMovementTracker


    ''' <summary>
    ''' Maps each moved key's signature to where it moved from and to
    ''' </summary>
    ''' <remarks>
    ''' Signature format: <c> {KeyName}{MovementKeySeparator}{KeyValue}{MovementKeySeparator}{SourceEntry} </c>,
    ''' compared case-insensitively
    ''' </remarks>
    Public Property MovedKeys As New Dictionary(Of String, KeyMovementInfo)(StringComparer.OrdinalIgnoreCase)

    ''' <summary>
    ''' Clears all tracking data
    ''' </summary>
    Public Sub Clear()

        MovedKeys.Clear()

    End Sub

End Class

''' <summary>
''' Information about a key that moved between entries
''' </summary>
Public Class KeyMovementInfo

    ''' <summary>
    ''' Source entry name where the key was originally located
    ''' </summary>
    Public Property SourceEntry As String

    ''' <summary>
    ''' Target entry name where the key was moved to
    ''' </summary>
    Public Property TargetEntry As String

    ''' <summary>
    ''' Creates a new <c> KeyMovementInfo </c>
    ''' </summary>
    ''' 
    ''' <param name="source">
    ''' Source entry name
    ''' </param>
    ''' 
    ''' <param name="target">
    ''' Target entry name
    ''' </param>
    Public Sub New(source As String,
                   target As String)

        SourceEntry = source
        TargetEntry = target

    End Sub

End Class