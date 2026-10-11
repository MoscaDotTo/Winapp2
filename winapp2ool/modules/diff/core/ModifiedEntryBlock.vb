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
''' The key lists one modified entry's block prints. We build every block once the analysis is
''' complete; the renderer prints these lists, and the summary's Modified totals are sums over
''' them, so the two can't disagree.
''' </summary>
Public Class ModifiedEntryBlock

    ''' <summary>
    ''' The entry's name
    ''' </summary>
    Public ReadOnly Property Name As String

    ''' <summary>
    ''' Indicates whether the entry took in content from removed entries, which gives it a merger
    ''' block measured against its own old keys and its sources' keys together
    ''' </summary>
    Public ReadOnly Property IsMerger As Boolean

    ''' <summary>
    ''' The keys the block lists as added. A merger block keeps keys that moved in from another
    ''' entry here, with their credit.
    ''' </summary>
    Public ReadOnly Property Added As List(Of iniKey)

    ''' <summary>
    ''' The added keys that didn't move in from another entry. The summary counts moved-in keys
    ''' on their own line, so this is what it counts as added.
    ''' </summary>
    Public ReadOnly Property Novel As List(Of iniKey)

    ''' <summary>
    ''' The removed keys the entry's own old version had (same type and value)
    ''' </summary>
    Public ReadOnly Property Removed As List(Of iniKey)

    ''' <summary>
    ''' In a merger block, the removed keys only a merged source had
    ''' </summary>
    Public ReadOnly Property Dropped As List(Of iniKey)

    ''' <summary>
    ''' Each updated key with the old keys it replaced
    ''' </summary>
    Public ReadOnly Property Updated As Dictionary(Of iniKey, List(Of iniKey))

    ''' <summary>
    ''' Creates the block for <paramref name="name"/> from the trackers as the analysis left them
    ''' </summary>
    '''
    ''' <param name="state">
    ''' The diff's state, after every comparison and the movement detection have run
    ''' </param>
    '''
    ''' <param name="oldFile">
    ''' The old version of winapp2.ini, whose version of the entry decides which removed keys
    ''' were its own
    ''' </param>
    '''
    ''' <param name="name">
    ''' The name of an entry in <c> ModifiedEntryNames </c>
    ''' </param>
    Friend Sub New(state As DiffState,
                   oldFile As iniFile,
                   name As String)

        Me.Name = name
        IsMerger = state.MergedEntries.MergedEntryNames.Contains(name)

        Dim tracked As List(Of iniKey) = Nothing
        Added = If(state.ModifiedEntries.AddedKeyTracker.TryGetValue(name, tracked), New List(Of iniKey)(tracked), New List(Of iniKey))
        Dim allRemoved = If(state.ModifiedEntries.RemovedKeyTracker.TryGetValue(name, tracked), New List(Of iniKey)(tracked), New List(Of iniKey))

        Dim trackedUpdates As Dictionary(Of iniKey, List(Of iniKey)) = Nothing
        Updated = If(state.ModifiedEntries.ModifiedKeyTracker.TryGetValue(name, trackedUpdates),
                     New Dictionary(Of iniKey, List(Of iniKey))(trackedUpdates),
                     New Dictionary(Of iniKey, List(Of iniKey)))

        If Not IsMerger Then

            Novel = Added
            Removed = allRemoved
            Dropped = New List(Of iniKey)
            Return

        End If

        ' A merger re-diff starts over from the entry's and its sources' old keys, so the keys that
        ' moved in are back among the added ones. Match on name as well as value, since a Detect
        ' and a RegKey often share a value
        Dim movedIn = state.KeyMovements.MovedInto(name)
        Novel = Added.Where(Function(key) Not movedIn.Any(Function(m) String.Equals(m.TargetKey.Name, key.Name, StringComparison.OrdinalIgnoreCase) AndAlso
                                                                      String.Equals(m.TargetKey.Value, key.Value, StringComparison.OrdinalIgnoreCase))).ToList()

        Dim ownOldKeys = oldFile.GetSection(name).Keys
        Dim wasOwnKey = Function(key As iniKey) ownOldKeys.Any(Function(own) String.Equals(own.KeyType, key.KeyType, StringComparison.OrdinalIgnoreCase) AndAlso
                                                                             String.Equals(own.Value, key.Value, StringComparison.OrdinalIgnoreCase))

        Removed = allRemoved.Where(wasOwnKey).ToList()
        Dropped = allRemoved.Where(Function(key) Not wasOwnKey(key)).ToList()

    End Sub

    ''' <summary>
    ''' Builds the block of every entry in <c> ModifiedEntryNames </c> into
    ''' <c> ModifiedBlocks </c>. Run it once the merger re-diffs and the movement detection are
    ''' done, since both change the trackers it reads.
    ''' </summary>
    '''
    ''' <param name="state">
    ''' The diff's state
    ''' </param>
    '''
    ''' <param name="oldFile">
    ''' The old version of winapp2.ini
    ''' </param>
    Friend Shared Sub BuildAll(state As DiffState,
                               oldFile As iniFile)

        For Each entryName In state.ModifiedEntries.ModifiedEntryNames
            state.ModifiedBlocks(entryName) = New ModifiedEntryBlock(state, oldFile, entryName)
        Next

    End Sub

End Class
