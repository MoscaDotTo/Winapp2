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
''' How a key type is numbered within an entry
''' </summary>
Public Enum KeyNumbering

    ''' <summary>One key per entry, never numbered</summary>
    Never

    ''' <summary>Unnumbered when the entry has one key of the type, numbered from 1 when it has more</summary>
    WhenSeveral

    ''' <summary>Always numbered from 1</summary>
    Always

End Enum

''' <summary>
''' What a key type does for its entry
''' </summary>
Public Enum KeyRole

    ''' <summary>Places the entry in a category (LangSecRef, Section)</summary>
    Category

    ''' <summary>Decides whether the entry applies to the system</summary>
    Detection

    ''' <summary>Sets the entry's initial state (Default)</summary>
    Setting

    ''' <summary>Text shown when the entry is enabled (Warning)</summary>
    Warning

    ''' <summary>Names something to delete</summary>
    Deletion

    ''' <summary>Names something to keep</summary>
    Exclusion

End Enum

''' <summary>
''' What kind of path a key type's value holds
''' </summary>
Public Enum KeyDomain

    ''' <summary>No path</summary>
    None

    ''' <summary>A filesystem path</summary>
    Filesystem

    ''' <summary>A registry path</summary>
    Registry

    ''' <summary>Either, chosen by the value's flag (ExcludeKey's FILE, PATH or REG)</summary>
    ByFlag

End Enum

''' <summary>
''' One winapp2.ini key type and the facts about it
''' </summary>
Public NotInheritable Class KeyTypeInfo

    ''' <summary>The type's position in winapp2.ini key order, from 0</summary>
    Public ReadOnly Property Order As Integer

    ''' <summary>The type's name, cased as winapp2.ini writes it</summary>
    Public ReadOnly Property Name As String

    ''' <summary>How keys of this type are numbered</summary>
    Public ReadOnly Property Numbering As KeyNumbering

    ''' <summary>What keys of this type do for their entry</summary>
    Public ReadOnly Property Role As KeyRole

    ''' <summary>What kind of path the value holds</summary>
    Public ReadOnly Property Domain As KeyDomain

    ''' <summary>Indicates whether the type is deprecated and should not be written</summary>
    Public ReadOnly Property Deprecated As Boolean

    Friend Sub New(order As Integer, name As String, numbering As KeyNumbering, role As KeyRole,
                   domain As KeyDomain, Optional deprecated As Boolean = False)

        Me.Order = order
        Me.Name = name
        Me.Numbering = numbering
        Me.Role = role
        Me.Domain = domain
        Me.Deprecated = deprecated

    End Sub

End Class

''' <summary>
''' The winapp2.ini key types, in the order an entry lists them. This is the one place their
''' names and attributes are written down; <c> KeyTypesTests </c> pins every fact.
''' </summary>
Public NotInheritable Class KeyTypes

    Private Sub New()
    End Sub

    Private Shared ReadOnly _rows As KeyTypeInfo() = {
        New KeyTypeInfo(0, "DetectOS", KeyNumbering.Never, KeyRole.Detection, KeyDomain.None),
        New KeyTypeInfo(1, "LangSecRef", KeyNumbering.Never, KeyRole.Category, KeyDomain.None),
        New KeyTypeInfo(2, "Section", KeyNumbering.Never, KeyRole.Category, KeyDomain.None),
        New KeyTypeInfo(3, "SpecialDetect", KeyNumbering.Never, KeyRole.Detection, KeyDomain.None, deprecated:=True),
        New KeyTypeInfo(4, "Detect", KeyNumbering.WhenSeveral, KeyRole.Detection, KeyDomain.Registry),
        New KeyTypeInfo(5, "DetectFile", KeyNumbering.WhenSeveral, KeyRole.Detection, KeyDomain.Filesystem),
        New KeyTypeInfo(6, "Default", KeyNumbering.Never, KeyRole.Setting, KeyDomain.None),
        New KeyTypeInfo(7, "Warning", KeyNumbering.Never, KeyRole.Warning, KeyDomain.None),
        New KeyTypeInfo(8, "FileKey", KeyNumbering.Always, KeyRole.Deletion, KeyDomain.Filesystem),
        New KeyTypeInfo(9, "RegKey", KeyNumbering.Always, KeyRole.Deletion, KeyDomain.Registry),
        New KeyTypeInfo(10, "ExcludeKey", KeyNumbering.Always, KeyRole.Exclusion, KeyDomain.ByFlag)
    }

    ' Upper-cased keys, so Find ignores case exactly as ToUpperInvariant does
    Private Shared ReadOnly _byUpperName As Dictionary(Of String, KeyTypeInfo) =
        _rows.ToDictionary(Function(t) t.Name.ToUpperInvariant(), StringComparer.Ordinal)

    ''' <summary>Every key type, in winapp2.ini key order</summary>
    Public Shared ReadOnly Property All As IReadOnlyList(Of KeyTypeInfo) = Array.AsReadOnly(_rows)

    ''' <summary>
    ''' Returns the key type called <paramref name="name"/>, ignoring case, or <c> Nothing </c> when
    ''' there is none. A name with a number on the end (<c> FileKey1 </c>) is not a key type; pass
    ''' a key's <c> KeyType </c>.
    ''' </summary>
    '''
    ''' <param name="name">A key type name such as <c> FileKey </c></param>
    Public Shared Function Find(name As String) As KeyTypeInfo

        If name Is Nothing Then Return Nothing
        Dim info As KeyTypeInfo = Nothing
        Return If(_byUpperName.TryGetValue(name.ToUpperInvariant(), info), info, Nothing)

    End Function

    ''' <summary>
    ''' Returns the key type with the longest name that <paramref name="text"/> starts with,
    ''' ignoring case, or <c> Nothing </c> when it starts with none. The longest match keeps
    ''' <c> DetectFile </c> and <c> DetectOS </c> from being read as <c> Detect </c>. The match
    ''' uses the current culture, as WinappDebug's missing <c> = </c> repair always has, so
    ''' characters the culture ignores, such as a soft hyphen, are skipped.
    ''' </summary>
    '''
    ''' <param name="text">The text to match, such as a key line missing its <c> = </c></param>
    Public Shared Function FindPrefix(text As String) As KeyTypeInfo

        If text Is Nothing Then Return Nothing
        Dim upperText = text.ToUpperInvariant()
        Dim best As KeyTypeInfo = Nothing
        For Each t In _rows
            If Not upperText.StartsWith(t.Name.ToUpperInvariant(), StringComparison.CurrentCulture) Then Continue For
            If best Is Nothing OrElse t.Name.Length > best.Name.Length Then best = t
        Next
        Return best

    End Function

End Class
