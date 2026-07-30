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
''' Tests for the shared scaffold substrate in <c> WebViewScaffolds </c> — the selection grammar
''' every scaffold family resolves through (<c> ResolveScaffolds </c>) and the catalog parser
''' (<c> ParseSection </c>). This layer previously had no coverage at all despite being consumed by
''' two modules and three engine families.
''' <br /><br />
'''
''' The selection grammar is the part worth pinning: explicit list, the <c> All </c> sentinel,
''' default fallback when the key is absent, "key present but empty" meaning emit nothing,
''' exclusions applied to either source, and unknown-name filtering. A regression in any of these
''' silently changes what gets deleted on users' machines rather than failing loudly.
''' </summary>
<TestClass()> Public Class WebViewScaffoldsTests

    ''' <summary>
    ''' Helper: build a catalog with the given scaffold names, each holding one dummy template
    ''' </summary>
    '''
    ''' <param name="names">
    ''' The scaffold names to register
    ''' </param>
    '''
    ''' <returns>
    ''' A case-insensitive catalog dictionary shaped like <c> LoadCatalog </c>'s result
    ''' </returns>
    Private Shared Function BuildCatalog(ParamArray names As String()) As Dictionary(Of String, List(Of String))

        Dim catalog As New Dictionary(Of String, List(Of String))(StringComparer.InvariantCultureIgnoreCase)

        For Each n In names

            catalog(n) = New List(Of String) From {$"%Root%\{n}|*"}

        Next

        Return catalog

    End Function

    ''' <summary>
    ''' Helper: resolve a selection against a catalog, discarding the menu output
    ''' </summary>
    '''
    ''' <param name="selected">
    ''' The explicit selection list; ignored when <paramref name="keyPresent"/> is False
    ''' </param>
    '''
    ''' <param name="keyPresent">
    ''' Whether the entry declared the family's scaffolds key at all
    ''' </param>
    '''
    ''' <param name="excluded">
    ''' Names to subtract from the active selection
    ''' </param>
    '''
    ''' <param name="catalog">
    ''' The catalog of known scaffolds
    ''' </param>
    '''
    ''' <param name="defaults">
    ''' The family's default set, used when the key is absent
    ''' </param>
    '''
    ''' <returns>
    ''' The resolved scaffold names
    ''' </returns>
    Private Shared Function Resolve(selected As String(),
                                    keyPresent As Boolean,
                                    excluded As String(),
                                    catalog As Dictionary(Of String, List(Of String)),
                                    defaults As String()) As List(Of String)

        Return winapp2ool.WebViewScaffolds.ResolveScaffolds(
            New List(Of String)(selected),
            keyPresent,
            New List(Of String)(excluded),
            catalog,
            defaults,
            "TestFamily",
            "[Test Entry *]",
            New winapp2ool.MenuSection)

    End Function

    ''' <summary>
    ''' With no scaffolds key declared, the family's default set is used
    ''' </summary>
    <TestMethod()> Public Sub Resolve_KeyAbsent_UsesDefaults()

        Dim catalog = BuildCatalog("Caches", "Telemetry", "WebCookies")
        Dim result = Resolve({}, False, {}, catalog, {"Caches", "Telemetry"})

        Assert.AreEqual(2, result.Count)
        Assert.IsTrue(result.Contains("Caches"))
        Assert.IsTrue(result.Contains("Telemetry"))
        Assert.IsFalse(result.Contains("WebCookies"), "host-risk scaffold must not appear in the default set")

    End Sub

    ''' <summary>
    ''' A declared but empty scaffolds key means emit nothing — distinct from the key being absent,
    ''' which falls back to the defaults
    ''' </summary>
    <TestMethod()> Public Sub Resolve_KeyPresentButEmpty_EmitsNothing()

        Dim catalog = BuildCatalog("Caches", "Telemetry")
        Dim result = Resolve({}, True, {}, catalog, {"Caches", "Telemetry"})

        Assert.AreEqual(0, result.Count)

    End Sub

    ''' <summary>
    ''' An explicit list is honoured verbatim and overrides the defaults
    ''' </summary>
    <TestMethod()> Public Sub Resolve_ExplicitList_OverridesDefaults()

        Dim catalog = BuildCatalog("Caches", "Telemetry", "WebCookies")
        Dim result = Resolve({"WebCookies"}, True, {}, catalog, {"Caches", "Telemetry"})

        Assert.AreEqual(1, result.Count)
        Assert.AreEqual("WebCookies", result(0))

    End Sub

    ''' <summary>
    ''' The <c> All </c> sentinel expands to every scaffold in the catalog, case-insensitively
    ''' </summary>
    <TestMethod()> Public Sub Resolve_AllSentinel_ExpandsWholeCatalog()

        Dim catalog = BuildCatalog("Caches", "Telemetry", "WebCookies", "WebStorage")
        Dim result = Resolve({"aLL"}, True, {}, catalog, {"Caches"})

        Assert.AreEqual(4, result.Count)

    End Sub

    ''' <summary>
    ''' <c> All </c> combined with exclusions yields the catalog minus the excluded names — the
    ''' shape every WebView-declaring entry in the corpus actually uses
    ''' </summary>
    <TestMethod()> Public Sub Resolve_AllSentinelWithExclusions_SubtractsExcluded()

        Dim catalog = BuildCatalog("Caches", "Telemetry", "WebCookies", "WebStorage", "LoginData")
        Dim result = Resolve({"All"}, True, {"WebStorage", "LoginData"}, catalog, {"Caches"})

        Assert.AreEqual(3, result.Count)
        Assert.IsFalse(result.Contains("WebStorage"))
        Assert.IsFalse(result.Contains("LoginData"))

    End Sub

    ''' <summary>
    ''' Exclusions apply to the default set too, not only to an explicit selection
    ''' </summary>
    <TestMethod()> Public Sub Resolve_ExclusionsAgainstDefaults_Subtract()

        Dim catalog = BuildCatalog("Caches", "Telemetry")
        Dim result = Resolve({}, False, {"Telemetry"}, catalog, {"Caches", "Telemetry"})

        Assert.AreEqual(1, result.Count)
        Assert.AreEqual("Caches", result(0))

    End Sub

    ''' <summary>
    ''' Exclusion matching is case-insensitive
    ''' </summary>
    <TestMethod()> Public Sub Resolve_ExclusionCaseInsensitive_Subtracts()

        Dim catalog = BuildCatalog("Caches", "Telemetry")
        Dim result = Resolve({}, False, {"tELEMETRY"}, catalog, {"Caches", "Telemetry"})

        Assert.AreEqual(1, result.Count)
        Assert.AreEqual("Caches", result(0))

    End Sub

    ''' <summary>
    ''' Names absent from the catalog are dropped rather than emitted or throwing — a misspelled
    ''' scaffold name must not abort a build
    ''' </summary>
    <TestMethod()> Public Sub Resolve_UnknownName_IsDropped()

        Dim catalog = BuildCatalog("Caches", "Telemetry")
        Dim result = Resolve({"Caches", "Cahces"}, True, {}, catalog, {"Caches"})

        Assert.AreEqual(1, result.Count)
        Assert.AreEqual("Caches", result(0))

    End Sub

    ''' <summary>
    ''' Selecting a name that differs only by case still resolves, since the catalog is
    ''' case-insensitive
    ''' </summary>
    <TestMethod()> Public Sub Resolve_SelectionCaseInsensitive_Resolves()

        Dim catalog = BuildCatalog("StorageQuota")
        Dim result = Resolve({"storagequota"}, True, {}, catalog, {})

        Assert.AreEqual(1, result.Count)

    End Sub

    ''' <summary>
    ''' The three shipping families' default sets are asserted here so a change to any of them is a
    ''' deliberate test edit rather than a silent shift in what every entry cleans. Electron's set
    ''' deliberately omits <c> VisitedLinks </c> — the surface is absent from modern Electron — even
    ''' though QtWebEngine's includes it.
    ''' </summary>
    <TestMethod()> Public Sub DefaultSets_MatchDocumentedTiers()

        CollectionAssert.AreEqual({"Caches", "Telemetry"},
                                  winapp2ool.WebViewScaffolds.DefaultScaffolds)

        CollectionAssert.AreEqual({"Caches", "StorageQuota", "Telemetry", "VisitedLinks"},
                                  winapp2ool.WebViewScaffolds.QtWebEngineDefaultScaffolds)

        CollectionAssert.AreEqual({"AppLogs", "Caches", "StorageQuota", "Telemetry", "UpdaterCache"},
                                  winapp2ool.WebViewScaffolds.ElectronDefaultScaffolds)

    End Sub

    ''' <summary>
    ''' <c> ParseSection </c> collects <c> FileKeyBase= </c> values in order and strips the family
    ''' prefix from the section header to derive the scaffold name
    ''' </summary>
    <TestMethod()> Public Sub ParseSection_CollectsTemplatesInOrder()

        Dim section As New winapp2ool.iniSection2("ElectronScaffold: Caches")
        section.AddKey(New winapp2ool.iniKey2("FileKeyBase=%ElectronRoot%\blob_storage|*|REMOVESELF"))
        section.AddKey(New winapp2ool.iniKey2("FileKeyBase=%ElectronRoot%\*Cache*|*|REMOVESELF"))

        Dim catalog As New Dictionary(Of String, List(Of String))(StringComparer.InvariantCultureIgnoreCase)
        winapp2ool.WebViewScaffolds.ParseSection(section, catalog, New winapp2ool.MenuSection, "ElectronScaffold:")

        Assert.IsTrue(catalog.ContainsKey("Caches"))
        Assert.AreEqual(2, catalog("Caches").Count)
        Assert.AreEqual("%ElectronRoot%\blob_storage|*|REMOVESELF", catalog("Caches")(0))
        Assert.AreEqual("%ElectronRoot%\*Cache*|*|REMOVESELF", catalog("Caches")(1))

    End Sub

    ''' <summary>
    ''' A scaffold section with no keys registers with an empty template list, so selecting it is a
    ''' no-op rather than a missing-key error
    ''' </summary>
    <TestMethod()> Public Sub ParseSection_EmptySection_RegistersEmptyList()

        Dim section As New winapp2ool.iniSection2("ElectronScaffold: Placeholder")

        Dim catalog As New Dictionary(Of String, List(Of String))(StringComparer.InvariantCultureIgnoreCase)
        winapp2ool.WebViewScaffolds.ParseSection(section, catalog, New winapp2ool.MenuSection, "ElectronScaffold:")

        Assert.IsTrue(catalog.ContainsKey("Placeholder"))
        Assert.AreEqual(0, catalog("Placeholder").Count)

    End Sub

    ''' <summary>
    ''' A duplicate scaffold name lets the last definition win rather than throwing on the
    ''' dictionary insert
    ''' </summary>
    <TestMethod()> Public Sub ParseSection_DuplicateName_LastDefinitionWins()

        Dim catalog As New Dictionary(Of String, List(Of String))(StringComparer.InvariantCultureIgnoreCase)

        Dim first As New winapp2ool.iniSection2("ElectronScaffold: Caches")
        first.AddKey(New winapp2ool.iniKey2("FileKeyBase=%ElectronRoot%\one|*"))
        winapp2ool.WebViewScaffolds.ParseSection(first, catalog, New winapp2ool.MenuSection, "ElectronScaffold:")

        Dim second As New winapp2ool.iniSection2("ElectronScaffold: Caches")
        second.AddKey(New winapp2ool.iniKey2("FileKeyBase=%ElectronRoot%\two|*"))
        winapp2ool.WebViewScaffolds.ParseSection(second, catalog, New winapp2ool.MenuSection, "ElectronScaffold:")

        Assert.AreEqual(1, catalog.Count)
        Assert.AreEqual(1, catalog("Caches").Count)
        Assert.AreEqual("%ElectronRoot%\two|*", catalog("Caches")(0))

    End Sub

End Class
