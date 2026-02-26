'    Copyright (C) 2018-2025 Hazel Ward
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
''' Encapsulates all state management for the Diff module
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
    ''' Resets all state to initial values
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
    ''' Tracks names of merged entries
    ''' </summary>
    ''' 
    ''' <remarks> 
    ''' Use regular HashSet but always access under SyncLock 
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
    ''' Tracks names of renamed entries
    ''' </summary>
    Public Property RenamedEntryNames As New HashSet(Of String)

    ''' <summary>
    ''' Tracks renamed entry pairs: OldName, NewName, OldName, NewName, ...
    ''' </summary>
    Public Property RenamedEntryPairs As New List(Of String)

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
''' Tracks entries that have been modified and the specific key changes
''' </summary>
Public Class ModifiedEntryTracker

    ''' <summary>
    ''' Tracks names of modified entries
    ''' </summary>
    Public Property ModifiedEntryNames As New HashSet(Of String)

    ''' <summary>
    ''' Tracks names of added entries
    ''' </summary>
    Public Property AddedEntryNames As New HashSet(Of String)

    ''' <summary>
    ''' Tracks names of removed entries
    ''' </summary>
    Public Property RemovedEntryNames As New HashSet(Of String)

    ''' <summary>
    ''' Tracks modified keys per entry: EntryName -> (KeyType -> List(Of Keys))
    ''' </summary>
    Public Property ModifiedKeyTracker As New Dictionary(Of String, Dictionary(Of iniKey, keyList))

    ''' <summary>
    ''' Tracks removed keys per entry: EntryName -> List(Of Keys)
    ''' </summary>
    Public Property RemovedKeyTracker As New Dictionary(Of String, keyList)

    ''' <summary>
    ''' Tracks added keys per entry: EntryName -> List(Of Keys)
    ''' </summary>
    Public Property AddedKeyTracker As New Dictionary(Of String, keyList)

    ''' <summary>
    ''' Tracks potential matching sections for modified entries
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
    ''' Counts the number of entries that were merged into newly added entries <br />
    ''' ie. entries that were removed but which had all of their contents merged into entries that were added
    ''' </summary>
    Public Property RemovedByAdditionCount As Integer = 0

    ''' <summary>
    ''' Counts the total number of entries that were merged <br /> ie. entries that were 
    ''' removed but which had some or all of their contents merged into other entries 
    ''' </summary>
    Public Property MergedEntryCount As Integer = 0

    ''' <summary>
    ''' Counts entries that were added with mergers
    ''' </summary>
    Public Property AddedEntryWithMergerCount As Integer = 0

    ''' <summary>
    ''' Counts total keys added in modified entries
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
    ''' Counts total entries where keys were replaced by updates
    ''' </summary>
    Public Property ModEntriesReplacedByUpdateTotal As Integer = 0

    ''' <summary>
    ''' Counts total entries where keys were removed without replacement
    ''' </summary>
    Public Property ModEntriesRemovedKeysWithoutReplacementTotal As Integer = 0

    ''' <summary>
    ''' Counts total keys that moved between entries
    ''' </summary>
    Public Property ModEntriesMovedKeysTotal As Integer = 0

    ''' <summary>
    ''' Counts total source entries providing keys that moved between entries
    ''' </summary>
    Public Property ModEntriesMovedKeysSourceCount As Integer = 0

    ''' <summary>
    ''' Counts total target entries receiving keys that moved between entries
    ''' </summary>
    Public Property ModEntriesMovedKeysTargetCount As Integer = 0

    ''' <summary>
    ''' Counts total keys added in modified entries
    ''' </summary>
    Public Property ModEntriesUpdatedKeyEntryCount As Integer = 0

    ''' <summary>
    ''' Counts entries that were added with mergers (i.e., entries that were
    ''' added and also had one or more old entries merged into them)
    ''' </summary>
    Public Property AddedWithMergersEntryCount As Integer = 0

    ''' <summary>
    ''' Counts entries that were merged into added entries (ie. old entries
    ''' that were merged into entries which were added)
    ''' </summary>
    Public Property AddedWithMergersSourceEntryCount As Integer = 0

    ' Novel keys (new keys not from merged sources)
    ''' <summary>
    ''' Counts total novel keys added in entries that were added with mergers <br /> ie. keys 
    ''' that were added in entries that had mergers, but were not part of the merged old entries)
    ''' </summary>
    Public Property AddedWithMergersNovelKeysTotal As Integer = 0
    Public Property AddedWithMergersNovelKeysEntryCount As Integer = 0

    ' Capturing keys (new keys that capture old keys)
    Public Property AddedWithMergersCapturingKeysTotal As Integer = 0
    Public Property AddedWithMergersCapturedKeysTotal As Integer = 0
    Public Property AddedWithMergersCapturingEntryCount As Integer = 0

    ' Dropped keys (keys from merged entries not carried over)
    Public Property AddedWithMergersDroppedKeysTotal As Integer = 0
    Public Property AddedWithMergersDroppedEntryCount As Integer = 0

    ' Carried over keys (keys from merged entries that were carried over to the added entry)
    Public Property AddedWithMergersCarriedOverKeysTotal As Integer = 0
    Public Property AddedWithMergersCarriedOverKeysEntryCount As Integer = 0

    ''' <summary>
    ''' Resets all counters to zero
    ''' </summary>
    Public Sub Reset()

        RemovedByAdditionCount = 0
        MergedEntryCount = 0
        AddedEntryWithMergerCount = 0
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

    End Sub

End Class

''' <summary>
''' Manages caches for performance optimization during diff operations
''' </summary>
Public Class DiffCaches

    ''' <summary>
    ''' Caches old entries by name for quick lookup
    ''' </summary>
    Public Property CachedOldEntries As New Dictionary(Of String, winapp2entry)

    ''' <summary>
    ''' Caches new entries by name for quick lookup
    ''' </summary>
    Public Property CachedNewEntries As New Dictionary(Of String, winapp2entry)

    ''' <summary>
    ''' Caches key match information to avoid redundant computations
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
    ''' Tracks keys that moved between entries
    ''' </summary>
    ''' <remarks>
    ''' Dictionary: Key signature -> Movement info <br />
    ''' Key format: "{KeyName}{MovementKeySeparator}{KeyValue}{MovementKeySeparator}{SourceEntry}"
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
    ''' Creates a new KeyMovementInfo instance
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