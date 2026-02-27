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
''' Base class for key comparison strategies
''' </summary>
Public MustInherit Class KeyComparisonStrategy

    ''' <summary>
    ''' Regex special characters to escape
    ''' </summary>
    Protected ReadOnly regexCharsIn As String() = {"*", "+", "{", "}", "[", "]", "$", "(", ")"}
    Protected ReadOnly regexCharsOut As String() = {".*", "\+", "\{", "\}", "\[", "\]", "\$", "\(", "\)"}


    ''' <summary>
    ''' Compares two iniKey objects for equivalence
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
    ''' Indicates the new key has more parameters
    ''' </param>
    ''' 
    ''' <param name="possibleWildCardReduction">
    ''' Indicates possible wildcard consolidation
    ''' </param>
    ''' 
    ''' <returns>
    ''' True if keys are equivalent
    ''' </returns>
    Public MustOverride Function Compare(newKey As iniKey,
                                         oldKey As iniKey,
                          Optional ByRef matchedFileKeyHasMoreParams As Boolean = False,
                          Optional ByRef possibleWildCardReduction As Boolean = False) As Boolean

    ''' <summary>
    ''' Compares two iniKey2 objects for equivalence
    ''' </summary>
    Public MustOverride Function Compare(newKey As iniKey2,
                                         oldKey As iniKey2,
                          Optional ByRef matchedFileKeyHasMoreParams As Boolean = False,
                          Optional ByRef possibleWildCardReduction As Boolean = False) As Boolean


    ''' <summary>
    ''' Checks the equivalence of two strings using regex. Regex matches must capture the 0th 
    ''' character in the <c> <paramref name="oldVal"/> </c>to be considered equivalent. 
    ''' </summary>
    ''' 
    ''' <param name="newVal">
    ''' The new value, assessed to see if it is identical to or 
    ''' captures via regex <c> <paramref name="oldVal"/> </c>
    ''' </param>
    ''' 
    ''' <param name="oldVal">
    ''' The old value, assessed to see if it is itentical to or 
    ''' captured via regex by <c> <paramref name="newVal"/> </c>
    ''' </param>
    ''' 
    ''' <returns>
    ''' <c> True </c> if values are equivalent <br />
    ''' <c> False </c> otherwise
    ''' </returns>
    ''' 
    ''' <remarks>
    ''' Supports wildcards (*) 
    ''' Attempts to prevent overly permissive captures (e.g., "History*" matching "Media History")
    ''' </remarks>
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

        ' For other wildcard patterns, use regex matching
        ' Single Match call: .Success replaces the redundant IsMatch, .Value replaces .ToString
        Dim firstMatch = Regex.Match(oldVal, newVal, RegexOptions.IgnoreCase)
        If Not firstMatch.Success Then Return False

        ' Ensure we captured from the beginning to avoid false positives
        ' e.g., "Media History" shouldn't match "History*"
        Return oldVal.StartsWith(firstMatch.Value, StringComparison.InvariantCultureIgnoreCase)

    End Function

End Class

''' <summary>
''' Strategy for comparing simple keys (most key types)
''' </summary>
Public Class SimpleKeyComparisonStrategy

    Inherits KeyComparisonStrategy

    Public Overrides Function Compare(newKey As iniKey,
                                      oldKey As iniKey,
                       Optional ByRef matchedFileKeyHasMoreParams As Boolean = False,
                       Optional ByRef possibleWildCardReduction As Boolean = False) As Boolean

        If Not newKey.compareTypes(oldKey) Then Return False

        Return String.Equals(newKey.Value, oldKey.Value, StringComparison.InvariantCultureIgnoreCase)

    End Function

    Public Overrides Function Compare(newKey As iniKey2,
                                      oldKey As iniKey2,
                       Optional ByRef matchedFileKeyHasMoreParams As Boolean = False,
                       Optional ByRef possibleWildCardReduction As Boolean = False) As Boolean

        If Not newKey.compareTypes(oldKey) Then Return False

        Return String.Equals(newKey.Value, oldKey.Value, StringComparison.InvariantCultureIgnoreCase)

    End Function

End Class

''' <summary>
''' Strategy for comparing FileKey and DetectFile keys (supports wildcards and paths)
''' </summary>
Public Class PathKeyComparisonStrategy

    Inherits KeyComparisonStrategy

    Private ReadOnly regexCharsIn As String() = {"*", "+", "{", "}", "[", "]", "$", "(", ")"}
    Private ReadOnly regexCharsOut As String() = {".*", "\+", "\{", "\}", "\[", "\]", "\$", "\(", "\)"}

    Public Overrides Function Compare(newKey As iniKey,
                                      oldKey As iniKey,
                       Optional ByRef matchedFileKeyHasMoreParams As Boolean = False,
                       Optional ByRef possibleWildCardReduction As Boolean = False) As Boolean

        If Not newKey.compareTypes(oldKey) Then Return False

        If String.Equals(newKey.Value, oldKey.Value, StringComparison.InvariantCultureIgnoreCase) Then Return True

        Dim newKeySplit = newKey.Value.Split(CChar("\"))
        Dim oldKeySplit = oldKey.Value.Split(CChar("\"))

        If newKeySplit.Length > oldKeySplit.Length Then Return False

        Dim isFileKey = newKey.typeIs("FileKey")
        Dim isRecurse = isFileKey AndAlso (newKey.Value.Contains("RECURSE") OrElse newKey.Value.Contains("REMOVESELF"))

        If isFileKey AndAlso Not isRecurse AndAlso newKeySplit.Length < oldKeySplit.Length Then Return False

        Dim isSanitized = False

        For i = 0 To newKeySplit.Length - 1

            ' Sanitize regex characters after matching first path component
            If Not isSanitized AndAlso (i >= 1 OrElse newKeySplit.Length - 1 = 0) Then SanitizePath(newKey.Value, newKeySplit, oldKeySplit) : isSanitized = True

            Dim newVal = newKeySplit(i)
            Dim oldVal = oldKeySplit(i)
            Dim isLastPiece = i = newKeySplit.Length - 1

            If isLastPiece AndAlso isFileKey Then Return FinalizeFileKeyEquivalence(oldVal, newVal, oldKeySplit, newKeySplit,
                                                                                    matchedFileKeyHasMoreParams, possibleWildCardReduction)
            If Not CompareValues(newVal, oldVal) Then Return False

        Next

        Return True

    End Function

    Public Overrides Function Compare(newKey As iniKey2,
                                      oldKey As iniKey2,
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

            If isLastPiece AndAlso isFileKey Then Return FinalizeFileKeyEquivalence(oldVal, newVal, oldKeySplit, newKeySplit,
                                                                                    matchedFileKeyHasMoreParams, possibleWildCardReduction)
            If Not CompareValues(newVal, oldVal) Then Return False

        Next

        Return True

    End Function

    ''' <summary>
    ''' Escapes regex special characters in path components
    ''' </summary>
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
    ''' Escapes regex special characters in string array
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
    ''' Compares the final parameter components for FileKeys
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
    ''' <param name="oldKeySplit">
    ''' All backslash-split path components of the old key value
    ''' </param>
    '''
    ''' <param name="newKeySplit">
    ''' All backslash-split path components of the new key value
    ''' </param>
    '''
    ''' <param name="matchedFileKeyHasMoreParams">
    ''' Set to <c>True</c> if the new key's parameter list is longer than the old key's
    ''' </param>
    '''
    ''' <param name="possibleWildCardReduction">
    ''' Set to <c>True</c> if the new key appears to have reduced wildcard specificity
    ''' </param>
    '''
    ''' <returns>
    ''' <c>True</c> if the final FileKey components are considered equivalent
    ''' </returns>
    Private Function FinalizeFileKeyEquivalence(oldVal As String,
                                               newVal As String,
                                               oldKeySplit As String(),
                                               newKeySplit As String(),
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
    ''' Confirms that any two parameters for a pair of FileKeys match
    ''' </summary>
    ''' 
    ''' <param name="flags">
    ''' The semicolon-delimited parameter string from the new key's pipe section
    ''' </param>
    '''
    ''' <param name="oldFlags">
    ''' The semicolon-delimited parameter string from the old key's pipe section
    ''' </param>
    '''
    ''' <param name="matchedFileKeyHasMoreParams">
    ''' Set to <c>True</c> if the new parameter list is longer than the old
    ''' </param>
    '''
    ''' <param name="possibleWildCardReduction">
    ''' Set to <c>True</c> if the match appears to reduce wildcard coverage
    ''' </param>
    '''
    ''' <returns>
    ''' <c>True</c> if at least one new parameter matches at least one old parameter
    ''' </returns>
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
''' Strategy for comparing Detect keys
''' </summary>
Public Class DetectKeyComparisonStrategy

    Inherits KeyComparisonStrategy

    ''' <summary>
    ''' Compares two Detect or RegKey keys, treating parent registry paths as capturing their children
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
    ''' <c>True</c> if the keys are equivalent or <paramref name="newKey"/> is a parent path of <paramref name="oldKey"/>
    ''' </returns>
    Public Overrides Function Compare(newKey As iniKey,
                                      oldKey As iniKey,
                       Optional ByRef matchedFileKeyHasMoreParams As Boolean = False,
                       Optional ByRef possibleWildCardReduction As Boolean = False) As Boolean

        If Not newKey.compareTypes(oldKey) Then Return False

        ' Exact match
        If String.Equals(newKey.Value, oldKey.Value, StringComparison.InvariantCultureIgnoreCase) Then Return True

        Dim isRegKey = newKey.typeIs("RegKey")

        ' Extract registry path (before pipe if RegKey, entire value if Detect)
        Dim newPath As String
        Dim oldPath As String
        Dim newFlags As String = ""
        Dim oldFlags As String = ""

        If isRegKey Then

            ' RegKey format: HKCU\Path|*.* 
            Dim newSplit = newKey.Value.Split(CChar("|"))
            Dim oldSplit = oldKey.Value.Split(CChar("|"))
            newPath = newSplit(0)
            oldPath = oldSplit(0)
            If newSplit.Length > 1 Then newFlags = newSplit(1)
            If oldSplit.Length > 1 Then oldFlags = oldSplit(1)

        Else

            ' Detect keys are pure registry paths, no pipe
            newPath = newKey.Value
            oldPath = oldKey.Value

        End If

        ' Check if paths match exactly
        If String.Equals(newPath, oldPath, StringComparison.InvariantCultureIgnoreCase) Then

            ' Paths match - for RegKey, also check file pattern flags
            ' File patterns after pipe can have wildcards
            If isRegKey AndAlso newFlags.Length > 0 Then Return CompareValues(newFlags, oldFlags)

            Return True

        End If

        ' ✅ Registry path hierarchy: parent captures children
        ' HKCU\Software\App captures HKCU\Software\App\1.0
        ' HKCU\Software\App captures HKCU\Software\App\1.0\SubKey
        If oldPath.StartsWith(newPath & "\", StringComparison.InvariantCultureIgnoreCase) Then

            ' newPath is a parent of oldPath (shorter path captures longer path)
            ' For RegKey, file pattern flags must also match if present
            If isRegKey AndAlso newFlags.Length > 0 Then Return CompareValues(newFlags, oldFlags)

            Return True

        End If

        ' No match
        Return False

    End Function

    Public Overrides Function Compare(newKey As iniKey2,
                                      oldKey As iniKey2,
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

            If isRegKey AndAlso newFlags.Length > 0 Then Return CompareValues(newFlags, oldFlags)

            Return True

        End If

        Return False

    End Function

End Class

''' <summary>
''' Factory for creating appropriate key comparison strategies
''' </summary>
Public Class KeyComparisonStrategyFactory

    ''' <summary>Singleton instance of <c>SimpleKeyComparisonStrategy</c> for non-path key types</summary>
    Private Shared ReadOnly simpleStrategy As New SimpleKeyComparisonStrategy()

    ''' <summary>Singleton instance of <c>PathKeyComparisonStrategy</c> for FileKey and DetectFile key types</summary>
    Private Shared ReadOnly pathStrategy As New PathKeyComparisonStrategy()

    ''' <summary>Singleton instance of <c>DetectKeyComparisonStrategy</c> for Detect and RegKey key types</summary>
    Private Shared ReadOnly detectStrategy As New DetectKeyComparisonStrategy()

    ''' <summary>
    ''' Gets the appropriate strategy for the given key type
    ''' </summary>
    ''' 
    ''' <param name="key">
    ''' The key whose type determines which strategy to return
    ''' </param>
    '''
    ''' <returns>
    ''' The <c>KeyComparisonStrategy</c> appropriate for <paramref name="key"/>'s type
    ''' </returns>
    Public Shared Function GetStrategy(key As iniKey) As KeyComparisonStrategy

        Select Case key.KeyType

            Case "FileKey", "DetectFile" : Return pathStrategy

            Case "Detect", "RegKey" : Return detectStrategy

            Case Else : Return simpleStrategy

        End Select

    End Function

    ''' <summary>
    ''' Compares two keys using the appropriate strategy
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
    ''' Set to <c>True</c> if the new key has more pipe-delimited parameters than the old
    ''' </param>
    '''
    ''' <param name="possibleWildCardReduction">
    ''' Set to <c>True</c> if the match appears to reduce wildcard coverage
    ''' </param>
    '''
    ''' <returns>
    ''' <c>True</c> if the two keys are considered equivalent under the appropriate strategy
    ''' </returns>
    Public Shared Function CompareKeys(newKey As iniKey,
                                       oldKey As iniKey,
                        Optional ByRef matchedFileKeyHasMoreParams As Boolean = False,
                        Optional ByRef possibleWildCardReduction As Boolean = False) As Boolean

        Dim strategy = GetStrategy(newKey)
        Return strategy.Compare(newKey, oldKey, matchedFileKeyHasMoreParams, possibleWildCardReduction)

    End Function

    ''' <summary>
    ''' Gets the appropriate strategy for the given <c>iniKey2</c> key type
    ''' </summary>
    Public Shared Function GetStrategy(key As iniKey2) As KeyComparisonStrategy

        Select Case key.KeyType

            Case "FileKey", "DetectFile" : Return pathStrategy

            Case "Detect", "RegKey" : Return detectStrategy

            Case Else : Return simpleStrategy

        End Select

    End Function

    ''' <summary>
    ''' Compares two <c>iniKey2</c> keys using the appropriate strategy.
    ''' </summary>
    Public Shared Function CompareKeys(newKey As iniKey2,
                                       oldKey As iniKey2,
                        Optional ByRef matchedFileKeyHasMoreParams As Boolean = False,
                        Optional ByRef possibleWildCardReduction As Boolean = False) As Boolean

        Dim strategy = GetStrategy(newKey)
        Return strategy.Compare(newKey, oldKey, matchedFileKeyHasMoreParams, possibleWildCardReduction)

    End Function

    ''' <summary>
    ''' Utility class for matching all keys in one <c>keyList</c> against another
    ''' </summary>
    Public Class KeyListMatcher

        ''' <summary>
        ''' Assesses how many keys in <paramref name="oldKeyList"/> are matched by keys in <paramref name="newKeyList"/>
        ''' </summary>
        '''
        ''' <param name="oldKeyList">The key list from the old version of an entry</param>
        ''' <param name="newKeyList">The key list from the new version of an entry</param>
        ''' <param name="disallowedValues">Optional set of path values too broad to count as matches</param>
        '''
        ''' <returns>A <c>KeyMatchResult</c> with match counts and flags</returns>
        Public Shared Function AssessMatches(oldKeyList As keyList,
                                             newKeyList As keyList,
                                    Optional disallowedValues As HashSet(Of String) = Nothing) As KeyMatchResult

            Dim result As New KeyMatchResult()
            Dim newKeyValues As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            ' Build lookup
            For Each newKey In newKeyList.Keys

                If Not newKeyValues.Contains(newKey.Value) Then newKeyValues.Add(newKey.Value)

            Next

            ' Compare keys
            For Each oldKey In oldKeyList.Keys

                If disallowedValues IsNot Nothing AndAlso disallowedValues.Contains(oldKey.Value) Then Continue For

                If newKeyValues.Contains(oldKey.Value) Then result.MatchCount += 1 : Continue For

                ' Try wildcard/regex matching
                For Each newKey In newKeyList.Keys

                    Dim matched = KeyComparisonStrategyFactory.CompareKeys(newKey, oldKey, result.MatchHadMoreParams, result.PossibleWildCardReduction)

                    If matched AndAlso disallowedValues IsNot Nothing Then

                        Dim newKeyPath = GetPathWithoutFlags(newKey.Value)

                        If disallowedValues.Contains(newKeyPath) Then matched = False

                    End If

                    If matched Then result.MatchCount += 1 : Exit For

                Next

            Next

            result.AllKeysMatched = result.MatchCount = oldKeyList.KeyCount
            result.CountsMatch = result.AllKeysMatched AndAlso newKeyList.KeyCount = oldKeyList.KeyCount

            Return result

        End Function

        Private Shared Function GetPathWithoutFlags(value As String) As String

            Return If(value.Contains("|"), value.Substring(0, value.IndexOf("|", StringComparison.InvariantCultureIgnoreCase)), value)

        End Function

    End Class

    ''' <summary>
    ''' Aggregate result of comparing all keys in one <c>keyList</c> against another
    ''' </summary>
    Public Class KeyMatchResult

        ''' <summary>Number of old keys that were matched by at least one new key</summary>
        Public Property MatchCount As Integer

        ''' <summary>Whether every old key was matched by at least one new key</summary>
        Public Property AllKeysMatched As Boolean

        ''' <summary>Whether all old keys were matched and the key counts are equal</summary>
        Public Property CountsMatch As Boolean

        ''' <summary>Whether any matched new key has more pipe-delimited parameters than its old counterpart</summary>
        Public Property MatchHadMoreParams As Boolean

        ''' <summary>Whether any matched key appears to have reduced wildcard specificity</summary>
        Public Property PossibleWildCardReduction As Boolean

    End Class

End Class

''' <summary>
''' Information about key matches between two <c>iniSection2</c> entries.
''' Mirrors <c>KeyMatchInfo</c> using <c>iniKey2</c> matched key sets.
''' </summary>
Public Class KeyMatchInfo2

    ''' <summary>Number of FileKey values from the old entry matched in the new entry</summary>
    Public Property FileKeyMatches As Integer

    ''' <summary>Number of RegKey values from the old entry matched in the new entry</summary>
    Public Property RegKeyMatches As Integer

    ''' <summary>Sum of FileKey and RegKey match counts</summary>
    Public Property TotalMatches As Integer

    ''' <summary>Whether all FileKeys from the old entry were matched in the new entry</summary>
    Public Property AllFileKeysMatched As Boolean = True

    ''' <summary>Whether all RegKeys from the old entry were matched in the new entry</summary>
    Public Property AllRegKeysMatched As Boolean = True

    ''' <summary>Whether every FileKey and RegKey from the old entry was matched</summary>
    Public Property AllKeysMatched As Boolean

    ''' <summary>Whether the count of FileKeys is the same in both old and new entries</summary>
    Public Property FileKeyCountsMatch As Boolean = True

    ''' <summary>Whether the count of RegKeys is the same in both old and new entries</summary>
    Public Property RegKeyCountsMatch As Boolean = True

    ''' <summary>Whether FileKey and RegKey counts both match between old and new entries</summary>
    Public Property CountsMatch As Boolean

    ''' <summary>Whether any matched new key has more pipe-delimited parameters than its old counterpart</summary>
    Public Property MatchHadMoreParams As Boolean

    ''' <summary>Whether any matched key appears to have reduced wildcard specificity</summary>
    Public Property PossibleWildCardReduction As Boolean

    ''' <summary>Set of old FileKey objects that were matched in the new entry</summary>
    Public Property MatchedOldFileKeys As New HashSet(Of iniKey2)

    ''' <summary>Set of old RegKey objects that were matched in the new entry</summary>
    Public Property MatchedOldRegKeys As New HashSet(Of iniKey2)

End Class