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
''' Represents a winapp2.ini entry with typed read-only key collections,
''' built from an <c> iniSection </c>
''' </summary>
Public Class winapp2entry

    ''' <summary>The name of the entry, without brackets</summary>
    Public Property Name As String

    ''' <summary>The full entry name with brackets</summary>
    Public ReadOnly Property FullName As String
        Get
            Return $"[{Name}]"
        End Get
    End Property

    ''' <summary>The starting line number of the source section</summary>
    Public ReadOnly Property LineNum As Integer

    Private ReadOnly _detectOS      As New List(Of iniKey)
    Private ReadOnly _langSecRef    As New List(Of iniKey)
    Private ReadOnly _sectionKey    As New List(Of iniKey)
    Private ReadOnly _specialDetect As New List(Of iniKey)
    Private ReadOnly _detects       As New List(Of iniKey)
    Private ReadOnly _detectFiles   As New List(Of iniKey)
    Private ReadOnly _defaultKey    As New List(Of iniKey)
    Private ReadOnly _warningKey    As New List(Of iniKey)
    Private ReadOnly _fileKeys      As New List(Of iniKey)
    Private ReadOnly _regKeys       As New List(Of iniKey)
    Private ReadOnly _excludeKeys   As New List(Of iniKey)
    Private ReadOnly _errorKeys     As New List(Of iniKey)

    ''' <summary>Keys with KeyType "DetectOS". Valid syntax: one key only.</summary>
    Public ReadOnly Property DetectOS As IReadOnlyList(Of iniKey)
        Get
            Return _detectOS
        End Get
    End Property

    ''' <summary>Keys with KeyType "LangSecRef". Valid syntax: one key only.</summary>
    Public ReadOnly Property LangSecRef As IReadOnlyList(Of iniKey)
        Get
            Return _langSecRef
        End Get
    End Property

    ''' <summary>Keys with KeyType "Section". Valid syntax: one key only.</summary>
    Public ReadOnly Property SectionKey As IReadOnlyList(Of iniKey)
        Get
            Return _sectionKey
        End Get
    End Property

    ''' <summary>Keys with KeyType "SpecialDetect" (deprecated)</summary>
    Public ReadOnly Property SpecialDetect As IReadOnlyList(Of iniKey)
        Get
            Return _specialDetect
        End Get
    End Property

    ''' <summary>Keys with KeyType "Detect"</summary>
    Public ReadOnly Property Detects As IReadOnlyList(Of iniKey)
        Get
            Return _detects
        End Get
    End Property

    ''' <summary>Keys with KeyType "DetectFile"</summary>
    Public ReadOnly Property DetectFiles As IReadOnlyList(Of iniKey)
        Get
            Return _detectFiles
        End Get
    End Property

    ''' <summary>Keys with KeyType "Default". Valid syntax: one key only.</summary>
    Public ReadOnly Property DefaultKey As IReadOnlyList(Of iniKey)
        Get
            Return _defaultKey
        End Get
    End Property

    ''' <summary>Keys with KeyType "Warning". Valid syntax: one key only.</summary>
    Public ReadOnly Property WarningKey As IReadOnlyList(Of iniKey)
        Get
            Return _warningKey
        End Get
    End Property

    ''' <summary>Keys with KeyType "FileKey"</summary>
    Public ReadOnly Property FileKeys As IReadOnlyList(Of iniKey)
        Get
            Return _fileKeys
        End Get
    End Property

    ''' <summary>
    ''' Replaces the FileKey bucket with <paramref name="newKeys"/> as given. We don't renumber or
    ''' sort them; <see cref="RenumberKeys"/> does that.
    ''' </summary>
    '''
    ''' <param name="newKeys">
    ''' The replacement FileKey sequence
    ''' </param>
    Public Sub ReplaceFileKeys(newKeys As IEnumerable(Of iniKey))
        _fileKeys.Clear()
        _fileKeys.AddRange(newKeys)
    End Sub

    ''' <summary>Keys with KeyType "RegKey"</summary>
    Public ReadOnly Property RegKeys As IReadOnlyList(Of iniKey)
        Get
            Return _regKeys
        End Get
    End Property

    ''' <summary>Keys with KeyType "ExcludeKey"</summary>
    Public ReadOnly Property ExcludeKeys As IReadOnlyList(Of iniKey)
        Get
            Return _excludeKeys
        End Get
    End Property

    ''' <summary>Keys with unrecognized KeyTypes</summary>
    Public ReadOnly Property ErrorKeys As IReadOnlyList(Of iniKey)
        Get
            Return _errorKeys
        End Get
    End Property

    ''' <summary>Indicates whether this entry has any detection key (DetectOS, Detect, DetectFile, or SpecialDetect)</summary>
    Public ReadOnly Property HasDetectionKey As Boolean
        Get
            Return _detectOS.Count > 0 OrElse _detects.Count > 0 OrElse
                   _detectFiles.Count > 0 OrElse _specialDetect.Count > 0
        End Get
    End Property

    ''' <summary>Indicates whether DetectOS is the only detection key type present</summary>
    Public ReadOnly Property HasOnlyDetectOS As Boolean
        Get
            Return _detectOS.Count > 0 AndAlso
                   _detects.Count = 0 AndAlso _detectFiles.Count = 0 AndAlso _specialDetect.Count = 0
        End Get
    End Property

    ''' <summary>Indicates whether the entry name ends with the required " *" suffix</summary>
    Public ReadOnly Property HasValidNameSuffix As Boolean
        Get
            Return Name.EndsWith(" *", StringComparison.InvariantCulture)
        End Get
    End Property

    ''' <summary>
    ''' Indicates whether the entry has one kind of categorization key, LangSecRef or Section,
    ''' but not both and not neither. Several keys of the same kind still pass here;
    ''' <see cref="SingletonViolations"/> reports those.
    ''' </summary>
    Public ReadOnly Property HasValidCategorization As Boolean
        Get
            Return (_langSecRef.Count > 0) Xor (_sectionKey.Count > 0)
        End Get
    End Property

    ''' <summary>Indicates whether the entry has at least one FileKey or RegKey</summary>
    Public ReadOnly Property HasDeletionKey As Boolean
        Get
            Return _fileKeys.Count > 0 OrElse _regKeys.Count > 0
        End Get
    End Property

    ''' <summary>
    ''' All extra keys in singleton buckets: every key after index 0 in DetectOS,
    ''' LangSecRef, Section, SpecialDetect, Default, and Warning.
    ''' An empty list means no singleton violations exist.
    ''' </summary>
    Public ReadOnly Property SingletonViolations As IReadOnlyList(Of iniKey)
        Get
            Dim result As New List(Of iniKey)
            For Each lst In {_detectOS, _langSecRef, _sectionKey, _specialDetect, _defaultKey, _warningKey}
                If lst.Count > 1 Then result.AddRange(lst.Skip(1))
            Next
            Return result
        End Get
    End Property

    ''' <summary>
    ''' Indicates whether ExcludeKeys are consistent with the deletion keys present.
    ''' False when FILE or PATH ExcludeKeys exist without FileKeys,
    ''' or when REG ExcludeKeys exist without RegKeys.
    ''' </summary>
    Public ReadOnly Property HasConsistentExcludeKeys As Boolean
        Get
            Dim hasFileExcludes = _excludeKeys.Any(Function(k)
                Dim p As New excludeKeyParams(k.Value)
                Return p.Flag = excludeKeyFlag.File OrElse p.Flag = excludeKeyFlag.Path
            End Function)
            Dim hasRegExcludes = _excludeKeys.Any(Function(k)
                Dim p As New excludeKeyParams(k.Value)
                Return p.Flag = excludeKeyFlag.Reg
            End Function)
            Return (Not hasFileExcludes OrElse _fileKeys.Count > 0) AndAlso
                   (Not hasRegExcludes OrElse _regKeys.Count > 0)
        End Get
    End Property

    ''' <summary>
    ''' Every key bucket in winapp2.ini key order, with <c> ErrorKeys </c> last.
    ''' <see cref="GetBucketIndex"/> gives each key type's position.
    ''' </summary>
    Public ReadOnly Property KeyLists As IReadOnlyList(Of IReadOnlyList(Of iniKey))

    ''' <summary>
    ''' Creates a new <c> winapp2entry </c> from <paramref name="section"/>, sorting each key into
    ''' its bucket by KeyType (case-insensitive). The entry shares the section's key objects, so
    ''' changing a key here changes it in the section too.
    ''' </summary>
    '''
    ''' <param name="section">A winapp2.ini format <c> iniSection </c></param>
    Public Sub New(section As iniSection)

        If section Is Nothing Then argIsNull(NameOf(section)) : Return

        Name    = section.Name
        LineNum = section.StartingLineNumber

        For Each key In section.Keys

            Select Case key.KeyType.ToUpperInvariant()
                Case "DETECTOS"      : _detectOS.Add(key)
                Case "LANGSECREF"    : _langSecRef.Add(key)
                Case "SECTION"       : _sectionKey.Add(key)
                Case "SPECIALDETECT" : _specialDetect.Add(key)
                Case "DETECT"        : _detects.Add(key)
                Case "DETECTFILE"    : _detectFiles.Add(key)
                Case "DEFAULT"       : _defaultKey.Add(key)
                Case "WARNING"       : _warningKey.Add(key)
                Case "FILEKEY"       : _fileKeys.Add(key)
                Case "REGKEY"        : _regKeys.Add(key)
                Case "EXCLUDEKEY"    : _excludeKeys.Add(key)
                Case Else            : _errorKeys.Add(key)
            End Select

        Next

        KeyLists = New List(Of IReadOnlyList(Of iniKey)) From {
            _detectOS, _langSecRef, _sectionKey, _specialDetect, _detects, _detectFiles,
            _defaultKey, _warningKey, _fileKeys, _regKeys, _excludeKeys, _errorKeys
        }

    End Sub

    ''' <summary>
    ''' Adds a key to the appropriate typed bucket based on its KeyType
    ''' </summary>
    ''' <param name="key">The key to add</param>
    Public Sub AddKey(key As iniKey)

        If key Is Nothing Then argIsNull(NameOf(key)) : Return

        Select Case key.KeyType.ToUpperInvariant()
            Case "DETECTOS"      : _detectOS.Add(key)
            Case "LANGSECREF"    : _langSecRef.Add(key)
            Case "SECTION"       : _sectionKey.Add(key)
            Case "SPECIALDETECT" : _specialDetect.Add(key)
            Case "DETECT"        : _detects.Add(key)
            Case "DETECTFILE"    : _detectFiles.Add(key)
            Case "DEFAULT"       : _defaultKey.Add(key)
            Case "WARNING"       : _warningKey.Add(key)
            Case "FILEKEY"       : _fileKeys.Add(key)
            Case "REGKEY"        : _regKeys.Add(key)
            Case "EXCLUDEKEY"    : _excludeKeys.Add(key)
            Case Else            : _errorKeys.Add(key)
        End Select

    End Sub

    ''' <summary>
    ''' Returns the 0-based index into <c> KeyLists </c> for the given key type name, or -1 if unrecognized.
    ''' </summary>
    ''' <param name="keyType">The key type name, e.g. "FileKey"</param>
    Public Shared Function GetBucketIndex(keyType As String) As Integer
        Select Case keyType.ToUpperInvariant()
            Case "DETECTOS"      : Return 0
            Case "LANGSECREF"    : Return 1
            Case "SECTION"       : Return 2
            Case "SPECIALDETECT" : Return 3
            Case "DETECT"        : Return 4
            Case "DETECTFILE"    : Return 5
            Case "DEFAULT"       : Return 6
            Case "WARNING"       : Return 7
            Case "FILEKEY"       : Return 8
            Case "REGKEY"        : Return 9
            Case "EXCLUDEKEY"    : Return 10
            Case "ERROR"         : Return 11
            Case Else            : Return -1
        End Select
    End Function

    ''' <summary>
    ''' Removes a key from the error bucket directly, bypassing <c> KeyType </c> routing.
    ''' Required when <c> cValidity </c> has partially repaired a key's Name before deciding
    ''' it cannot be salvaged, leaving the key's <c> KeyType </c> in an inconsistent state.
    ''' </summary>
    ''' <param name="key">The key to remove from the error bucket</param>
    Public Sub ForceRemoveErrorKey(key As iniKey)
        _errorKeys.Remove(key)
    End Sub

    ''' <summary>
    ''' Moves any error key whose <c> KeyType </c> is now a recognized winapp2.ini type
    ''' into the appropriate typed bucket. Called after <c> cValidity </c> has had a chance
    ''' to repair broken keys (e.g. fixing a missing "=" restores a valid KeyType).
    ''' </summary>
    Public Sub ReclassifyErrorKeys()
        Dim toMove = _errorKeys.Where(Function(k)
            Return GetBucketIndex(k.KeyType) >= 0 AndAlso GetBucketIndex(k.KeyType) <= 10
        End Function).ToList()
        For Each k In toMove
            _errorKeys.Remove(k)
            AddKey(k)
        Next
    End Sub

    ''' <summary>
    ''' Removes a key from its typed bucket. We find the bucket from the key's current KeyType, so
    ''' a key whose name changed since it was added won't be found. <see cref="ForceRemoveErrorKey"/>
    ''' covers that case for error keys.
    ''' </summary>
    ''' <param name="key">The key to remove</param>
    Public Sub RemoveKey(key As iniKey)

        If key Is Nothing Then argIsNull(NameOf(key)) : Return

        Select Case key.KeyType.ToUpperInvariant()
            Case "DETECTOS"      : _detectOS.Remove(key)
            Case "LANGSECREF"    : _langSecRef.Remove(key)
            Case "SECTION"       : _sectionKey.Remove(key)
            Case "SPECIALDETECT" : _specialDetect.Remove(key)
            Case "DETECT"        : _detects.Remove(key)
            Case "DETECTFILE"    : _detectFiles.Remove(key)
            Case "DEFAULT"       : _defaultKey.Remove(key)
            Case "WARNING"       : _warningKey.Remove(key)
            Case "FILEKEY"       : _fileKeys.Remove(key)
            Case "REGKEY"        : _regKeys.Remove(key)
            Case "EXCLUDEKEY"    : _excludeKeys.Remove(key)
            Case Else            : _errorKeys.Remove(key)
        End Select

    End Sub

    ''' <summary>Returns a new <c> iniSection </c> holding this entry's keys in winapp2.ini order</summary>
    Public Function ToIniSection() As iniSection

        Dim s As New iniSection(Name, LineNum)

        For Each lst In KeyLists
            For Each key In lst
                s.AddKey(key)
            Next
        Next

        Return s

    End Function

    ''' <summary>
    ''' Renumbers the deletion key buckets (FileKey, RegKey, ExcludeKey) sequentially from 1,
    ''' after sorting each bucket by value (case-insensitive).
    ''' </summary>
    Public Sub RenumberKeys()

        RenumberBucket(_fileKeys)
        RenumberBucket(_regKeys)
        RenumberBucket(_excludeKeys)

    End Sub

    Private Shared Sub RenumberBucket(bucket As List(Of iniKey))

        If bucket.Count = 0 Then Return

        Dim keyType = bucket(0).KeyType

        bucket.Sort(Function(a, b) String.Compare(a.Value, b.Value, StringComparison.OrdinalIgnoreCase))

        For i = 0 To bucket.Count - 1
            bucket(i).Name = keyType & CStr(i + 1)
        Next

    End Sub

End Class
