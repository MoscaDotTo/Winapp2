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

Imports System.Text.RegularExpressions

''' <summary>
''' Base class for key comparison strategies. Each one decides whether a new key matches or
''' covers an old key. The test is one-directional, so callers that want either direction call
''' it twice with the keys swapped.
''' </summary>
Public MustInherit Class KeyComparisonStrategy

    ''' <summary>
    ''' Characters rewritten before regex matching: <c> * </c> becomes a regex wildcard and
    ''' the rest are escaped
    ''' </summary>
    Protected ReadOnly regexCharsIn As String() = {"*", "+", "{", "}", "[", "]", "$", "(", ")"}

    ''' <summary>
    ''' Regex replacements corresponding to each entry in <see cref="regexCharsIn"/>
    ''' </summary>
    Protected ReadOnly regexCharsOut As String() = {".*", "\+", "\{", "\}", "\[", "\]", "\$", "\(", "\)"}

    Private Shared ReadOnly _regexCache As New Concurrent.ConcurrentDictionary(Of String, Regex)(StringComparer.Ordinal)

    ''' <summary>
    ''' Returns whether <paramref name="newKey"/> is equivalent to <paramref name="oldKey"/> or
    ''' covers it, for example through a wildcard or a parent path
    ''' </summary>
    '''
    ''' <param name="newKey">
    ''' The key from the new version
    ''' </param>
    '''
    ''' <param name="oldKey">
    ''' The key from the old version
    ''' </param>
    '''
    ''' <param name="matchedFileKeyHasMoreParams">
    ''' Assigned by strategies that compare FileKey patterns one by one: <c> True </c> when the
    ''' new key has more patterns than the old. Other strategies leave it unchanged. <br /><br />
    ''' Optional, Default: <c> False </c>
    ''' </param>
    '''
    ''' <param name="possibleWildCardReduction">
    ''' Assigned alongside <paramref name="matchedFileKeyHasMoreParams"/>: <c> True </c> when the
    ''' match appears to narrow wildcard coverage <br /><br />
    ''' Optional, Default: <c> False </c>
    ''' </param>
    Public MustOverride Function Compare(newKey As iniKey,
                                         oldKey As iniKey,
                          Optional ByRef matchedFileKeyHasMoreParams As Boolean = False,
                          Optional ByRef possibleWildCardReduction As Boolean = False) As Boolean


    ''' <summary>
    ''' Returns whether <paramref name="newVal"/> equals or captures <paramref name="oldVal"/>.
    ''' A <paramref name="newVal"/> of <c> * </c> or <c> .* </c> captures anything, and equal values
    ''' match ignoring case. Otherwise a <paramref name="newVal"/> without a <c> * </c> doesn't match.
    ''' One starting with <c> .*. </c> matches any <paramref name="oldVal"/> ending in the text after
    ''' the <c> .* </c>, ignoring case. Any other <paramref name="newVal"/> is used as a case-insensitive
    ''' regex, and matches when <paramref name="oldVal"/> starts with the text of the regex's first
    ''' match in it, so <c> History.* </c> doesn't match <c> Media History </c>. The match needn't
    ''' reach the end of <paramref name="oldVal"/>.
    ''' </summary>
    '''
    ''' <param name="newVal">
    ''' The new value, expected to be regex text already (see <see cref="regexCharsIn"/>)
    ''' when it holds a wildcard
    ''' </param>
    '''
    ''' <param name="oldVal">
    ''' The old value, tested as plain text
    ''' </param>
    Protected Shared Function CompareValues(newVal As String,
                                            oldVal As String) As Boolean

        Dim newValHasWildcard = newVal.Contains("*")
        Dim newValIsOnlyWildcard = newVal.Equals(".*", StringComparison.InvariantCultureIgnoreCase) OrElse
                                   newVal.Equals("*", StringComparison.InvariantCultureIgnoreCase)

        Dim oldValHasWildcard = oldVal.Contains("*")
        Dim oldValIsOnlyWildcard = oldVal.Equals(".*", StringComparison.InvariantCultureIgnoreCase) OrElse
                                   oldVal.Equals("*", StringComparison.InvariantCultureIgnoreCase)

        ' Catch-all wildcards always match
        If newValIsOnlyWildcard Then Return True

        ' Exact match check
        Dim matched = String.Equals(newVal, oldVal, StringComparison.InvariantCultureIgnoreCase)
        If matched Then Return matched

        ' No wildcard in newVal means no match possible (already checked exact match)
        If Not newValHasWildcard Then Return False

        ' Handle file extension patterns (*.ext or .*.ext after sanitization)
        ' These should only match files that actually END with that extension
        ' .*.log should match "error.log" but NOT "LOG" or "LOG.old"
        If newVal.StartsWith(".*.", StringComparison.InvariantCultureIgnoreCase) Then

            ' Extract the extension (e.g., ".log" from ".*.log")
            Dim extension = newVal.Substring(2) ' Remove ".*" prefix
            Return oldVal.EndsWith(extension, StringComparison.InvariantCultureIgnoreCase)

        End If

        ' For other wildcard patterns, use compiled regex matching (cached per pattern)
        Dim compiled = _regexCache.GetOrAdd(newVal, Function(p) New Regex(p, RegexOptions.IgnoreCase))
        Dim firstMatch = compiled.Match(oldVal)
        If Not firstMatch.Success Then Return False

        ' Ensure we captured from the beginning to avoid false positives
        ' e.g., "Media History" shouldn't match "History*"
        Return oldVal.StartsWith(firstMatch.Value, StringComparison.InvariantCultureIgnoreCase)

    End Function

End Class

''' <summary>
''' Strategy for every key type except FileKey, DetectFile, Detect and RegKey: values must be equal
''' </summary>
Public Class SimpleKeyComparisonStrategy

    Inherits KeyComparisonStrategy

    ''' <summary>
    ''' Returns whether <paramref name="newKey"/> and <paramref name="oldKey"/> share a type
    ''' and have identical values, both compared ignoring case
    ''' </summary>
    ''' 
    ''' <param name="newKey">
    ''' The key from the new version
    ''' </param>
    ''' 
    ''' <param name="oldKey">
    ''' The key from the old version
    ''' </param>
    ''' 
    ''' <param name="matchedFileKeyHasMoreParams">
    ''' Unused by this strategy; present for interface compatibility
    ''' </param>
    ''' 
    ''' <param name="possibleWildCardReduction">
    ''' Unused by this strategy; present for interface compatibility
    ''' </param>
    ''' 
    Public Overrides Function Compare(newKey As iniKey,
                                      oldKey As iniKey,
                       Optional ByRef matchedFileKeyHasMoreParams As Boolean = False,
                       Optional ByRef possibleWildCardReduction As Boolean = False) As Boolean

        If Not newKey.compareTypes(oldKey) Then Return False

        Return String.Equals(newKey.Value, oldKey.Value, StringComparison.InvariantCultureIgnoreCase)

    End Function

End Class

''' <summary>
''' Strategy for comparing FileKey and DetectFile keys as paths, with wildcards
''' </summary>
Public Class PathKeyComparisonStrategy

    Inherits KeyComparisonStrategy

    ''' <summary>
    ''' Returns whether <paramref name="newKey"/> matches or covers <paramref name="oldKey"/>,
    ''' comparing their backslash-delimited components in order with <see cref="CompareValues"/>.
    ''' Types must match and equal values match at once, both ignoring case. <br /><br />
    '''
    ''' A new key with more components than the old never matches. A new FileKey needs the same
    ''' number of components unless its value contains <c> RECURSE </c> or <c> REMOVESELF </c>
    ''' (case-sensitive), when it may be shorter and so cover the old key's subfolders. A new
    ''' DetectFile may always be shorter. The first component is compared as written (unless it's
    ''' the only one), and the rest after <see cref="SanitizePath"/> has rewritten them. A component right after <c> Packages </c>
    ''' also matches when the old key's wildcard there covers the new component, unless the old
    ''' component is a bare <c> * </c>. <br /><br />
    '''
    ''' For a FileKey the new key's last component goes to <see cref="FinalizeFileKeyEquivalence"/>,
    ''' which ignores the flag after the second pipe.
    ''' </summary>
    ''' <param name="newKey">The key from the new version</param>
    ''' <param name="oldKey">The key from the old version</param>
    ''' <param name="matchedFileKeyHasMoreParams">
    ''' Assigned when a FileKey match is decided pattern by pattern: <c> True </c> if the new key
    ''' has more semicolon-delimited patterns than the old key. Left unchanged otherwise. <br /><br />
    ''' Optional, Default: <c> False </c>
    ''' </param>
    ''' <param name="possibleWildCardReduction">
    ''' Assigned with <paramref name="matchedFileKeyHasMoreParams"/>: <c> True </c> if the match
    ''' appears to narrow wildcard coverage. Left unchanged otherwise. <br /><br />
    ''' Optional, Default: <c> False </c>
    ''' </param>
    Public Overrides Function Compare(newKey As iniKey,
                                      oldKey As iniKey,
                       Optional ByRef matchedFileKeyHasMoreParams As Boolean = False,
                       Optional ByRef possibleWildCardReduction As Boolean = False) As Boolean

        If Not newKey.compareTypes(oldKey) Then Return False

        If String.Equals(newKey.Value, oldKey.Value, StringComparison.InvariantCultureIgnoreCase) Then Return True

        Dim newKeySplit = newKey.BackslashSplit
        Dim oldKeySplit = oldKey.BackslashSplit

        If newKeySplit.Length > oldKeySplit.Length Then Return False

        Dim isFileKey = newKey.typeIs("FileKey")
        Dim isRecurse = isFileKey AndAlso (newKey.Value.Contains("RECURSE") OrElse newKey.Value.Contains("REMOVESELF"))

        If isFileKey AndAlso Not isRecurse AndAlso newKeySplit.Length < oldKeySplit.Length Then Return False

        Dim isSanitized = False

        For i = 0 To newKeySplit.Length - 1

            If Not isSanitized AndAlso (i >= 1 OrElse newKeySplit.Length - 1 = 0) Then SanitizePath(newKey.Value, newKeySplit, oldKeySplit) : isSanitized = True

            Dim newVal = newKeySplit(i)
            Dim oldVal = oldKeySplit(i)
            Dim isLastPiece = i = newKeySplit.Length - 1

            If isLastPiece AndAlso isFileKey Then Return FinalizeFileKeyEquivalence(oldVal, newVal, matchedFileKeyHasMoreParams, possibleWildCardReduction)

            If CompareValues(newVal, oldVal) Then Continue For

            If IsPackageMonikerComponent(oldKeySplit, i) AndAlso oldVal.Contains("*") Then

                Dim pristineOldSplit = oldKey.BackslashSplit
                Dim monikerPattern = {pristineOldSplit(i)}
                SanitizeRegex(monikerPattern)

                If monikerPattern(0).Equals(".*") Then Return False

                If CompareValues(monikerPattern(0), newVal) Then Continue For

            End If

            Return False

        Next

        Return True

    End Function

    ''' <summary>
    ''' Rewrites path components as regex text with <see cref="SanitizeRegex"/>, deciding from
    ''' <paramref name="keyValue"/>. If it has no <c> * </c>, we do nothing. If its first <c> * </c>
    ''' comes before its first pipe, or it has no pipe, we rewrite every component of both splits.
    ''' Otherwise we rewrite only the last component of each.
    ''' </summary>
    '''
    ''' <param name="keyValue">
    ''' The new key's full value, used to locate the first wildcard and pipe separator
    ''' </param>
    '''
    ''' <param name="newKeySplit">
    ''' Backslash-split components of the new key value; sanitized in place when wildcards precede any pipe
    ''' </param>
    '''
    ''' <param name="oldKeySplit">
    ''' Backslash-split components of the old key value; sanitized in place when wildcards precede any pipe
    ''' </param>
    Private Sub SanitizePath(keyValue As String,
                       ByRef newKeySplit As String(),
                       ByRef oldKeySplit As String())

        Dim firstWildcardIndex = keyValue.IndexOf("*", StringComparison.InvariantCultureIgnoreCase)
        If firstWildcardIndex = -1 Then Return

        Dim pipeIndex = keyValue.IndexOf("|", StringComparison.InvariantCultureIgnoreCase)
        Dim wildcardIsBeforePipe = firstWildcardIndex < pipeIndex

        If wildcardIsBeforePipe OrElse pipeIndex = -1 Then

            SanitizeRegex(newKeySplit)
            SanitizeRegex(oldKeySplit)

            Return

        End If

        ' Sanitize flags separately if wildcard is after pipe
        Dim flags = {newKeySplit.Last, oldKeySplit.Last}
        SanitizeRegex(flags)
        newKeySplit(newKeySplit.Length - 1) = flags(0)
        oldKeySplit(oldKeySplit.Length - 1) = flags(1)

    End Sub

    ''' <summary>
    ''' Rewrites each string as regex text: <c> * </c> becomes <c> .* </c> and <c> + { } [ ] $ ( ) </c>
    ''' are escaped. Other regex characters, including <c> . </c>, pass through unchanged.
    ''' </summary>
    ''' <param name="splitPath">The array of path components to sanitize in place</param>
    Private Sub SanitizeRegex(ByRef splitPath As String())

        For k = 0 To splitPath.Length - 1

            For j = 0 To regexCharsIn.Length - 1

                If splitPath(k).Contains(regexCharsIn(j)) Then splitPath(k) = splitPath(k).Replace(regexCharsIn(j), regexCharsOut(j))

            Next

        Next

    End Sub

    ''' <summary>
    ''' Returns whether the component at <paramref name="index"/> directly follows a
    ''' <c> Packages </c> component (ignoring case), which makes it a UWP package-family folder
    ''' that may take a reverse-wildcard match
    ''' </summary>
    '''
    ''' <param name="keySplit">
    ''' Backslash-split components of a key value
    ''' </param>
    '''
    ''' <param name="index">
    ''' Index of the component under consideration
    ''' </param>
    Private Shared Function IsPackageMonikerComponent(keySplit As String(),
                                                      index As Integer) As Boolean

        Return index >= 1 AndAlso keySplit(index - 1).Equals("Packages", StringComparison.InvariantCultureIgnoreCase)

    End Function

    ''' <summary>
    ''' Returns whether the new FileKey's last component matches the old key's component at the
    ''' same position. The text before the first pipe must match. Then the pattern lists (between
    ''' the first and second pipe) match if <see cref="CompareValues"/> accepts them whole, or, when
    ''' either list has a <c> ; </c>, if any new pattern matches any old pattern
    ''' (<see cref="MatchParameters"/>). Nothing after the second pipe is compared, so adding or
    ''' dropping <c> RECURSE </c> alone doesn't stop a match. When the old component is a mid-path
    ''' folder (a shorter recursive new key), its pattern list is empty, so in practice only a
    ''' bare <c> * </c> pattern, alone or in the list, covers it.
    ''' </summary>
    '''
    ''' <param name="oldVal">
    ''' The final path component of the old key value, including pipe-delimited pattern and flags
    ''' </param>
    '''
    ''' <param name="newVal">
    ''' The final path component of the new key value, including pipe-delimited pattern and flags
    ''' </param>
    '''
    ''' <param name="matchedFileKeyHasMoreParams">
    ''' Assigned by <see cref="MatchParameters"/> when it decides the match; left unchanged otherwise
    ''' </param>
    '''
    ''' <param name="possibleWildCardReduction">
    ''' Assigned by <see cref="MatchParameters"/> when it decides the match; left unchanged otherwise
    ''' </param>
    Private Function FinalizeFileKeyEquivalence(oldVal As String,
                                               newVal As String,
                                               ByRef matchedFileKeyHasMoreParams As Boolean,
                                               ByRef possibleWildCardReduction As Boolean) As Boolean

        Dim pipe = CChar("|")
        Dim semi = CChar(";")
        Dim oldSplit = oldVal.Split(pipe)
        Dim newSplit = newVal.Split(pipe)
        Dim oldValFinal = oldSplit(0)
        Dim newValFinal = newSplit(0)
        Dim oldFlags = If(oldSplit.Length > 1, oldSplit(1), "")
        Dim flags = If(newSplit.Length > 1, newSplit(1), "")

        If Not CompareValues(newValFinal, oldValFinal) Then Return False

        If CompareValues(flags, oldFlags) Then Return True

        If Not (flags.Contains(semi) OrElse oldFlags.Contains(semi)) Then Return False

        Return MatchParameters(flags, oldFlags, matchedFileKeyHasMoreParams, possibleWildCardReduction)

    End Function

    ''' <summary>
    ''' Returns whether at least one new pattern matches at least one old pattern under
    ''' <see cref="CompareValues"/>. On the first matching pair we assign both outputs, overwriting
    ''' whatever they held, and stop. When nothing matches we leave them unchanged.
    ''' </summary>
    '''
    ''' <param name="flags">
    ''' The semicolon-delimited pattern list from the new key's pipe section
    ''' </param>
    '''
    ''' <param name="oldFlags">
    ''' The semicolon-delimited pattern list from the old key's pipe section
    ''' </param>
    '''
    ''' <param name="matchedFileKeyHasMoreParams">
    ''' Set to whether the new list has more patterns than the old
    ''' </param>
    '''
    ''' <param name="possibleWildCardReduction">
    ''' Set to whether the matched old pattern has a <c> * </c> that the matched new pattern lacks,
    ''' or the new list is shorter than the old and has a <c> * </c> anywhere
    ''' </param>
    Private Function MatchParameters(flags As String,
                                oldFlags As String,
                                ByRef matchedFileKeyHasMoreParams As Boolean,
                                ByRef possibleWildCardReduction As Boolean) As Boolean

        Dim delimiter = CChar(";")
        Dim splitParams = flags.Split(delimiter)
        Dim oldSplitParams = oldFlags.Split(delimiter)

        For Each param In splitParams

            For Each oldParam In oldSplitParams

                If Not CompareValues(param, oldParam) Then Continue For

                matchedFileKeyHasMoreParams = splitParams.Length > oldSplitParams.Length

                ' Check for wildcard reduction in individual parameters OR overall parameter count
                ' Example: *bookmarks.bak → bookmarks.bak (individual param loses wildcard)
                ' Example: *.bak;*.tmp → *.bak (parameter count reduction with wildcard)
                possibleWildCardReduction = (oldParam.Contains("*") AndAlso Not param.Contains("*")) OrElse
                                            (splitParams.Length < oldSplitParams.Length AndAlso flags.Contains("*"))

                Return True

            Next

        Next

        Return False

    End Function

End Class

''' <summary>
''' Strategy for comparing Detect and RegKey keys as registry paths, where a parent path covers its children
''' </summary>
Public Class DetectKeyComparisonStrategy

    Inherits KeyComparisonStrategy

    ''' <summary>
    ''' Returns whether <paramref name="newKey"/> matches or covers <paramref name="oldKey"/>.
    ''' Types must match and equal values match at once, both ignoring case. A RegKey's value
    ''' name (after its first pipe) is split off its path, and a Detect key's whole value is its
    ''' path. On the same path (ignoring case), a new RegKey without a value name covers the old
    ''' key, and one with a value name matches when <see cref="CompareValues"/> accepts the old
    ''' value name. A new key whose path is a parent of the old key's path covers it, except a
    ''' RegKey with a value name.
    ''' </summary>
    '''
    ''' <param name="newKey">
    ''' The key from the new version
    ''' </param>
    '''
    ''' <param name="oldKey">
    ''' The key from the old version
    ''' </param>
    '''
    ''' <param name="matchedFileKeyHasMoreParams">
    ''' Unused by this strategy; present for interface compatibility
    ''' </param>
    '''
    ''' <param name="possibleWildCardReduction">
    ''' Unused by this strategy; present for interface compatibility
    ''' </param>
    '''
    ''' <returns>
    ''' <c> True </c> if the keys are equivalent, or <paramref name="newKey"/> is a parent path of
    ''' <paramref name="oldKey"/>. A value-targeted RegKey requires an exact path match, since a
    ''' value deletion beneath one key is not subsumed by a parent path.
    ''' </returns>
    Public Overrides Function Compare(newKey As iniKey,
                                      oldKey As iniKey,
                       Optional ByRef matchedFileKeyHasMoreParams As Boolean = False,
                       Optional ByRef possibleWildCardReduction As Boolean = False) As Boolean

        If Not newKey.compareTypes(oldKey) Then Return False

        If String.Equals(newKey.Value, oldKey.Value, StringComparison.InvariantCultureIgnoreCase) Then Return True

        Dim isRegKey = newKey.typeIs("RegKey")

        Dim newPath As String
        Dim oldPath As String
        Dim newFlags As String = ""
        Dim oldFlags As String = ""

        If isRegKey Then

            Dim newSplit = newKey.PipeSplit
            Dim oldSplit = oldKey.PipeSplit
            newPath = newSplit(0)
            oldPath = oldSplit(0)
            If newSplit.Length > 1 Then newFlags = newSplit(1)
            If oldSplit.Length > 1 Then oldFlags = oldSplit(1)

        Else

            newPath = newKey.Value
            oldPath = oldKey.Value

        End If

        If String.Equals(newPath, oldPath, StringComparison.InvariantCultureIgnoreCase) Then

            If isRegKey AndAlso newFlags.Length > 0 Then Return CompareValues(newFlags, oldFlags)

            Return True

        End If

        If oldPath.StartsWith(newPath & "\", StringComparison.InvariantCultureIgnoreCase) Then

            ' A value-targeted RegKey deletes a value beneath one specific key, so a parent path does
            ' not subsume it: A|V and A\B|V delete the same value from different keys, not the same key.
            ' Parent capture remains valid only for whole-key deletions (newFlags empty) and Detect keys.
            If isRegKey AndAlso newFlags.Length > 0 Then Return False

            Return True

        End If

        Return False

    End Function

End Class

''' <summary>
''' Factory for creating appropriate key comparison strategies
''' </summary>
Public Class KeyComparisonStrategyFactory

    ''' <summary>
    ''' Singleton instance of <c> SimpleKeyComparisonStrategy </c> 
    ''' for non-path key types
    ''' </summary>
    Private Shared ReadOnly simpleStrategy As New SimpleKeyComparisonStrategy()

    ''' <summary>
    ''' Singleton instance of <c> PathKeyComparisonStrategy </c>
    ''' for FileKey and DetectFile key types
    ''' </summary>
    Private Shared ReadOnly pathStrategy As New PathKeyComparisonStrategy()

    ''' <summary>
    ''' Singleton instance of <c> DetectKeyComparisonStrategy </c>
    ''' for Detect and RegKey key types
    ''' </summary>
    Private Shared ReadOnly detectStrategy As New DetectKeyComparisonStrategy()

    ''' <summary>
    ''' Returns the strategy for <paramref name="key"/>'s type: paths for <c> FileKey </c> and
    ''' <c> DetectFile </c>, registry paths for <c> Detect </c> and <c> RegKey </c>, and exact values
    ''' for everything else. The type name match is case-sensitive.
    ''' </summary>
    '''
    ''' <param name="key">
    ''' The key whose type determines which strategy to return
    ''' </param>
    '''
    Public Shared Function GetStrategy(key As iniKey) As KeyComparisonStrategy

        Select Case key.KeyType

            Case "FileKey", "DetectFile" : Return pathStrategy

            Case "Detect", "RegKey" : Return detectStrategy

            Case Else : Return simpleStrategy

        End Select

    End Function

    ''' <summary>
    ''' Returns whether <paramref name="newKey"/> matches or covers <paramref name="oldKey"/>, using
    ''' the strategy <see cref="GetStrategy"/> picks for <paramref name="newKey"/>'s type. The test
    ''' is one-directional.
    ''' </summary>
    '''
    ''' <param name="newKey">
    ''' The key from the new version
    ''' </param>
    '''
    ''' <param name="oldKey">
    ''' The key from the old version
    ''' </param>
    '''
    ''' <param name="matchedFileKeyHasMoreParams">
    ''' Assigned when a FileKey match is decided pattern by pattern: <c> True </c> if the new key
    ''' has more semicolon-delimited patterns than the old. Left unchanged otherwise. <br /><br />
    ''' Optional, Default: <c> False </c>
    ''' </param>
    '''
    ''' <param name="possibleWildCardReduction">
    ''' Assigned with <paramref name="matchedFileKeyHasMoreParams"/>: <c> True </c> if the match
    ''' appears to narrow wildcard coverage. Left unchanged otherwise. <br /><br />
    ''' Optional, Default: <c> False </c>
    ''' </param>
    Public Shared Function CompareKeys(newKey As iniKey,
                                       oldKey As iniKey,
                        Optional ByRef matchedFileKeyHasMoreParams As Boolean = False,
                        Optional ByRef possibleWildCardReduction As Boolean = False) As Boolean

        Dim strategy = GetStrategy(newKey)
        Return strategy.Compare(newKey, oldKey, matchedFileKeyHasMoreParams, possibleWildCardReduction)

    End Function

End Class

''' <summary>
''' Information about how an old entry's FileKeys and RegKeys match a new entry's.
''' Other key types aren't counted.
''' </summary>
Public Class KeyMatchInfo

    ''' <summary>
    ''' Number of FileKey values from the old entry matched in the new entry, not counting
    ''' values in <c> DisallowedPaths </c>
    ''' </summary>
    Public Property FileKeyMatches As Integer

    ''' <summary>
    ''' Number of RegKey values from the old entry matched in the new entry, not counting
    ''' values in <c> DisallowedPaths </c>
    ''' </summary>
    Public Property RegKeyMatches As Integer

    ''' <summary>
    ''' Sum of FileKey and RegKey match counts
    ''' </summary>
    Public Property TotalMatches As Integer

    ''' <summary>
    ''' Indicates whether all FileKeys from the old entry were matched in the new entry
    ''' </summary>
    Public Property AllFileKeysMatched As Boolean = True

    ''' <summary>
    ''' Indicates whether all RegKeys from the old entry were matched in the new entry
    ''' </summary>
    Public Property AllRegKeysMatched As Boolean = True

    ''' <summary>
    ''' Indicates whether every FileKey and RegKey from the old entry was matched
    ''' </summary>
    Public Property AllKeysMatched As Boolean

    ''' <summary>
    ''' Indicates whether every old FileKey matched and both entries have the same number of FileKeys.
    ''' Stays <c> True </c> when the old entry has no FileKeys.
    ''' </summary>
    Public Property FileKeyCountsMatch As Boolean = True

    ''' <summary>
    ''' Indicates whether every old RegKey matched and both entries have the same number of RegKeys.
    ''' Stays <c> True </c> when the old entry has no RegKeys.
    ''' </summary>
    Public Property RegKeyCountsMatch As Boolean = True

    ''' <summary>
    ''' Indicates whether both <see cref="FileKeyCountsMatch"/> and <see cref="RegKeyCountsMatch"/> hold
    ''' </summary>
    Public Property CountsMatch As Boolean

    ''' <summary>
    ''' Indicates whether a matched new FileKey has more semicolon-delimited patterns than its old
    ''' counterpart. Each FileKey match decided pattern by pattern overwrites it, so it reflects
    ''' the last such match rather than any of them.
    ''' </summary>
    Public Property MatchHadMoreParams As Boolean

    ''' <summary>
    ''' Indicates whether a matched FileKey appears to have narrowed its wildcard coverage,
    ''' overwritten the same way as <see cref="MatchHadMoreParams"/>
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

End Class