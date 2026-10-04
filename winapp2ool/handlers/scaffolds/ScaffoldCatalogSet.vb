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
''' One family's scaffold catalog: scaffold name to its <c> FileKeyBase= </c> templates, plus
''' the names of the scaffolds that declared <c> Tier=Legacy </c>. A legacy scaffold is kept in
''' the catalog so it can be selected by name, but the <c> All </c> sentinel withholds it. That
''' lets a pattern that only older engines write stay on record without shipping to every entry.
''' <br /><br />
'''
''' It is still a <c> Dictionary </c>, so a consumer that only reads templates needn't know about
''' tiers. A plain dictionary has no tiers: <see cref="ScaffoldCatalogs.ResolveScaffolds"/> treats
''' it as having no legacy scaffolds, and <see cref="ScaffoldCatalogs.ParseSection"/> warns on a
''' <c> Tier=Legacy </c> it can't record.
''' </summary>
Public Class ScaffoldCatalog
    Inherits Dictionary(Of String, List(Of String))

    ''' <summary> Names of the scaffolds declaring <c> Tier=Legacy </c> </summary>
    Public ReadOnly Property Legacy As HashSet(Of String) = New HashSet(Of String)(StringComparer.InvariantCultureIgnoreCase)

    ''' <summary> Creates an empty catalog keyed case-insensitively by scaffold name </summary>
    Public Sub New()

        MyBase.New(StringComparer.InvariantCultureIgnoreCase)

    End Sub

End Class

''' <summary>
''' The scaffold catalogs loaded from one scaffold directory, partitioned by engine family.
''' Produced by <see cref="ScaffoldCatalogs.LoadCatalogDirectory"/> and consumed by the entry
''' generators.
''' <br /><br />
'''
''' Family lookup never fails. <see cref="ForFamily"/> creates an empty catalog on first
''' request for a family the directory did not supply, so a builder that binds a family whose
''' catalog is missing emits zero keys for it rather than throwing. Loading doesn't warn about
''' that case, but every entry that declares the family's root then warns once for each
''' scaffold it requests, since none of them is in the empty catalog.
''' </summary>
Public Class ScaffoldCatalogSet

    ''' <summary>
    ''' The per-family scaffold catalogs, keyed case-insensitively by family token
    ''' (<c> WebView </c>, <c> QtWebEngine </c>, <c> Electron </c>). Each value is that family's
    ''' catalog, keyed by scaffold name.
    ''' </summary>
    Private ReadOnly _families As New Dictionary(Of String, ScaffoldCatalog)(StringComparer.InvariantCultureIgnoreCase)

    ''' <summary>
    ''' The family tokens present in this set, in first-seen order. This includes any family
    ''' that <see cref="ForFamily"/> created empty because a caller asked for it.
    ''' </summary>
    Public ReadOnly Property Families As IEnumerable(Of String)
        Get
            Return _families.Keys
        End Get
    End Property

    ''' <summary>
    ''' Returns the catalog for <paramref name="familyLabel"/>, creating and registering an
    ''' empty one if the family is not yet present. Every caller gets the same catalog object,
    ''' not a copy. <see cref="ScaffoldCatalogs.LoadCatalogDirectory"/> fills it in place this
    ''' way, so a consumer that only wants to read must not change it.
    ''' </summary>
    '''
    ''' <param name="familyLabel">
    ''' The family token, e.g. <c> Electron </c>. Matched case-insensitively.
    ''' </param>
    '''
    ''' <returns>
    ''' That family's catalog, keyed by scaffold name; empty when the family supplied none
    ''' </returns>
    Public Function ForFamily(familyLabel As String) As ScaffoldCatalog

        Dim catalog As ScaffoldCatalog = Nothing

        If Not _families.TryGetValue(familyLabel, catalog) Then

            catalog = New ScaffoldCatalog
            _families.Add(familyLabel, catalog)

        End If

        Return catalog

    End Function

End Class
