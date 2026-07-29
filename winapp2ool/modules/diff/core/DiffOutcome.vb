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

Imports System.Text

''' <summary>
''' The bottom line of a completed Diff, reduced to counts. Where the diff log narrates every
''' change at length, this answers only "did anything change, and how much" — the question a
''' scripted build pipeline asks when deciding whether the rebuild is worth publishing
''' </summary>
'''
''' <remarks>
''' This is a semantic verdict, not a textual one. Diff normalizes deprecated paths on both
''' sides before comparing (see <c> PathReplacements </c>), does not see comments outside the
''' preamble, and ignores ordering, so <c> HasChanges </c> being <c> False </c> does NOT prove
''' the two files are textually identical. Callers gating on "should this build be published"
''' should compare the files themselves and use this for reporting
''' </remarks>
Public Class DiffOutcome

    ''' <summary>
    ''' The number of entries present in the new file but not the old one
    ''' </summary>
    Public ReadOnly Property AddedEntries As Integer

    ''' <summary>
    ''' The number of entries present in the old file but not the new one, including
    ''' those accounted for by a rename or a merge
    ''' </summary>
    Public ReadOnly Property RemovedEntries As Integer

    ''' <summary>
    ''' The number of entries present in both files whose keys differ
    ''' </summary>
    Public ReadOnly Property ModifiedEntries As Integer

    ''' <summary>
    ''' The number of removed entries whose content reappeared under a new name
    ''' </summary>
    Public ReadOnly Property RenamedEntries As Integer

    ''' <summary>
    ''' The number of removed entries whose content was folded into some other entry
    ''' </summary>
    Public ReadOnly Property MergedEntries As Integer

    ''' <summary>
    ''' The number of keys added across all modified entries
    ''' </summary>
    Public ReadOnly Property AddedKeys As Integer

    ''' <summary>
    ''' The number of keys removed without replacement across all modified entries
    ''' </summary>
    Public ReadOnly Property RemovedKeys As Integer

    ''' <summary>
    ''' The number of keys updated in place across all modified entries
    ''' </summary>
    Public ReadOnly Property UpdatedKeys As Integer

    ''' <summary>
    ''' The number of keys that relocated from one entry to another
    ''' </summary>
    Public ReadOnly Property MovedKeys As Integer

    ''' <summary>
    ''' The number of entries in the old file
    ''' </summary>
    Public ReadOnly Property OldEntryCount As Integer

    ''' <summary>
    ''' The number of entries in the new file
    ''' </summary>
    Public ReadOnly Property NewEntryCount As Integer

    ''' <summary>
    ''' Indicates that the Diff observed at least one change of any kind
    ''' </summary>
    Public ReadOnly Property HasChanges As Boolean
        Get

            Return AddedEntries > 0 OrElse
                   RemovedEntries > 0 OrElse
                   ModifiedEntries > 0 OrElse
                   AddedKeys > 0 OrElse
                   RemovedKeys > 0 OrElse
                   UpdatedKeys > 0 OrElse
                   MovedKeys > 0 OrElse
                   OldEntryCount <> NewEntryCount

        End Get
    End Property

    ''' <summary>
    ''' Collects the counts a completed Diff accumulated in its <c> DiffState </c>
    ''' </summary>
    '''
    ''' <param name="state">
    ''' The <c> DiffState </c> of a Diff run whose pipeline has completed
    ''' </param>
    '''
    ''' <param name="oldEntries">
    ''' The number of entries in the old file
    ''' </param>
    '''
    ''' <param name="newEntries">
    ''' The number of entries in the new file
    ''' </param>
    Public Sub New(state As DiffState,
                   oldEntries As Integer,
                   newEntries As Integer)

        If state Is Nothing Then argIsNull(NameOf(state)) : Return

        AddedEntries = state.ModifiedEntries.AddedEntryNames.Count
        RemovedEntries = state.ModifiedEntries.RemovedEntryNames.Count
        ModifiedEntries = state.ModifiedEntries.ModifiedEntryNames.Count
        RenamedEntries = state.MergedEntries.RenamedEntryNames.Count
        MergedEntries = state.MergedEntries.OldToNewMergeDict.Count

        AddedKeys = state.Statistics.ModEntriesAddedKeyTotal
        RemovedKeys = state.Statistics.ModEntriesRemovedKeysWithoutReplacementTotal
        UpdatedKeys = state.Statistics.ModEntriesUpdatedKeyTotal
        MovedKeys = state.Statistics.ModEntriesMovedKeysTotal

        OldEntryCount = oldEntries
        NewEntryCount = newEntries

    End Sub

    ''' <summary>
    ''' Renders the outcome as <c> key=value </c> lines for a calling script to parse.
    ''' Every value is an integer or <c> true </c>/<c> false </c>, so the text is safe to
    ''' read with PowerShell's <c> ConvertFrom-StringData </c>
    ''' </summary>
    '''
    ''' <returns>
    ''' The outcome as newline-separated <c> key=value </c> pairs
    ''' </returns>
    Public Function ToSummaryText() As String

        Dim out As New StringBuilder

        out.AppendLine($"haschanges={HasChanges.ToString().ToLowerInvariant()}")
        out.AppendLine($"oldentrycount={OldEntryCount}")
        out.AppendLine($"newentrycount={NewEntryCount}")
        out.AppendLine($"addedentries={AddedEntries}")
        out.AppendLine($"removedentries={RemovedEntries}")
        out.AppendLine($"modifiedentries={ModifiedEntries}")
        out.AppendLine($"renamedentries={RenamedEntries}")
        out.AppendLine($"mergedentries={MergedEntries}")
        out.AppendLine($"addedkeys={AddedKeys}")
        out.AppendLine($"removedkeys={RemovedKeys}")
        out.AppendLine($"updatedkeys={UpdatedKeys}")
        out.AppendLine($"movedkeys={MovedKeys}")

        Return out.ToString()

    End Function

    ''' <summary>
    ''' Renders the outcome as a single human-readable line for console and log output
    ''' </summary>
    '''
    ''' <returns>
    ''' A one-line description of everything the Diff found
    ''' </returns>
    Public Overrides Function ToString() As String

        If Not HasChanges Then Return "No changes detected"

        Dim parts As New List(Of String)

        If AddedEntries > 0 Then parts.Add($"{AddedEntries} added")
        If RemovedEntries > 0 Then parts.Add($"{RemovedEntries} removed")
        If ModifiedEntries > 0 Then parts.Add($"{ModifiedEntries} modified")
        If RenamedEntries > 0 Then parts.Add($"{RenamedEntries} renamed")
        If MergedEntries > 0 Then parts.Add($"{MergedEntries} merged")

        Dim entryPart = If(parts.Count = 0, "no entry changes", $"entries: {String.Join(", ", parts)}")

        parts.Clear()

        If AddedKeys > 0 Then parts.Add($"{AddedKeys} added")
        If RemovedKeys > 0 Then parts.Add($"{RemovedKeys} removed")
        If UpdatedKeys > 0 Then parts.Add($"{UpdatedKeys} updated")
        If MovedKeys > 0 Then parts.Add($"{MovedKeys} moved")

        Dim keyPart = If(parts.Count = 0, "no key changes", $"keys: {String.Join(", ", parts)}")

        Return $"{entryPart}; {keyPart}"

    End Function

End Class
