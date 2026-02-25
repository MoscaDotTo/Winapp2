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
''' built from an <c>iniSection2</c>
''' </summary>
Public Class winapp2entry2

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

    Private ReadOnly _detectOS      As New List(Of iniKey2)
    Private ReadOnly _langSecRef    As New List(Of iniKey2)
    Private ReadOnly _sectionKey    As New List(Of iniKey2)
    Private ReadOnly _specialDetect As New List(Of iniKey2)
    Private ReadOnly _detects       As New List(Of iniKey2)
    Private ReadOnly _detectFiles   As New List(Of iniKey2)
    Private ReadOnly _defaultKey    As New List(Of iniKey2)
    Private ReadOnly _warningKey    As New List(Of iniKey2)
    Private ReadOnly _fileKeys      As New List(Of iniKey2)
    Private ReadOnly _regKeys       As New List(Of iniKey2)
    Private ReadOnly _excludeKeys   As New List(Of iniKey2)
    Private ReadOnly _errorKeys     As New List(Of iniKey2)

    ''' <summary>Keys with KeyType "DetectOS". Valid syntax: one key only.</summary>
    Public ReadOnly Property DetectOS As IReadOnlyList(Of iniKey2)
        Get
            Return _detectOS
        End Get
    End Property

    ''' <summary>Keys with KeyType "LangSecRef". Valid syntax: one key only.</summary>
    Public ReadOnly Property LangSecRef As IReadOnlyList(Of iniKey2)
        Get
            Return _langSecRef
        End Get
    End Property

    ''' <summary>Keys with KeyType "Section". Valid syntax: one key only.</summary>
    Public ReadOnly Property SectionKey As IReadOnlyList(Of iniKey2)
        Get
            Return _sectionKey
        End Get
    End Property

    ''' <summary>Keys with KeyType "SpecialDetect" (deprecated)</summary>
    Public ReadOnly Property SpecialDetect As IReadOnlyList(Of iniKey2)
        Get
            Return _specialDetect
        End Get
    End Property

    ''' <summary>Keys with KeyType "Detect"</summary>
    Public ReadOnly Property Detects As IReadOnlyList(Of iniKey2)
        Get
            Return _detects
        End Get
    End Property

    ''' <summary>Keys with KeyType "DetectFile"</summary>
    Public ReadOnly Property DetectFiles As IReadOnlyList(Of iniKey2)
        Get
            Return _detectFiles
        End Get
    End Property

    ''' <summary>Keys with KeyType "Default". Valid syntax: one key only.</summary>
    Public ReadOnly Property DefaultKey As IReadOnlyList(Of iniKey2)
        Get
            Return _defaultKey
        End Get
    End Property

    ''' <summary>Keys with KeyType "Warning". Valid syntax: one key only.</summary>
    Public ReadOnly Property WarningKey As IReadOnlyList(Of iniKey2)
        Get
            Return _warningKey
        End Get
    End Property

    ''' <summary>Keys with KeyType "FileKey"</summary>
    Public ReadOnly Property FileKeys As IReadOnlyList(Of iniKey2)
        Get
            Return _fileKeys
        End Get
    End Property

    ''' <summary>Keys with KeyType "RegKey"</summary>
    Public ReadOnly Property RegKeys As IReadOnlyList(Of iniKey2)
        Get
            Return _regKeys
        End Get
    End Property

    ''' <summary>Keys with KeyType "ExcludeKey"</summary>
    Public ReadOnly Property ExcludeKeys As IReadOnlyList(Of iniKey2)
        Get
            Return _excludeKeys
        End Get
    End Property

    ''' <summary>Keys with unrecognized KeyTypes</summary>
    Public ReadOnly Property ErrorKeys As IReadOnlyList(Of iniKey2)
        Get
            Return _errorKeys
        End Get
    End Property

    ''' <summary>
    ''' All key lists in winapp2.ini declaration order, mirroring <c>KeyListList</c>
    ''' on the legacy <c>winapp2entry</c>
    ''' </summary>
    Public ReadOnly Property KeyLists As IReadOnlyList(Of IReadOnlyList(Of iniKey2))

    ''' <summary>
    ''' Creates a <c>winapp2entry2</c> from an <c>iniSection2</c>
    ''' </summary>
    ''' <param name="section">A winapp2.ini format <c>iniSection2</c></param>
    Public Sub New(section As iniSection2)

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

        KeyLists = New List(Of IReadOnlyList(Of iniKey2)) From {
            _detectOS, _langSecRef, _sectionKey, _specialDetect, _detects, _detectFiles,
            _defaultKey, _warningKey, _fileKeys, _regKeys, _excludeKeys, _errorKeys
        }

    End Sub

    ''' <summary>
    ''' Adds a key to the appropriate typed bucket based on its KeyType
    ''' </summary>
    ''' <param name="key">The key to add</param>
    Public Sub AddKey(key As iniKey2)

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
    ''' Removes a key from its typed bucket
    ''' </summary>
    ''' <param name="key">The key to remove</param>
    Public Sub RemoveKey(key As iniKey2)

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

    ''' <summary>
    ''' Reconstructs an <c>iniSection2</c> from the typed key buckets in winapp2.ini order
    ''' </summary>
    Public Function ToIniSection() As iniSection2

        Dim s As New iniSection2(Name, LineNum)

        For Each lst In KeyLists
            For Each key In lst
                s.AddKey(key)
            Next
        Next

        Return s

    End Function

End Class
