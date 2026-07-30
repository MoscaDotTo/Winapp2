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
''' Tests for the shared scaffold substrate in <c> ScaffoldCatalogs </c> — the selection grammar
''' every scaffold family resolves through (<c> ResolveScaffolds </c>), the catalog parser
''' (<c> ParseSection </c>), the directory loader that discovers families from section headers
''' (<c> LoadCatalogDirectory </c>), and the placeholder substitution engine both builders share
''' (<c> FanOutPlaceholder </c> / <c> BindFamilyTemplates </c>).
''' <br /><br />
'''
''' The selection grammar is the part worth pinning: explicit list, the <c> All </c> sentinel,
''' default fallback when the key is absent, "key present but empty" meaning emit nothing,
''' exclusions applied to either source, and unknown-name filtering. A regression in any of these
''' silently changes what gets deleted on users' machines rather than failing loudly. The
''' substitution engine's drop-on-empty-roots rule matters for the same reason in reverse: lose it
''' and a literal <c> %ElectronUpdaterRoot% </c> ships inside a published FileKey.
''' </summary>
<TestClass()> Public Class ScaffoldCatalogsTests

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

        Return winapp2ool.ScaffoldCatalogs.ResolveScaffolds(
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
                                  winapp2ool.ScaffoldCatalogs.DefaultScaffolds)

        CollectionAssert.AreEqual({"Caches", "StorageQuota", "Telemetry", "VisitedLinks"},
                                  winapp2ool.ScaffoldCatalogs.QtWebEngineDefaultScaffolds)

        CollectionAssert.AreEqual({"AppLogs", "Caches", "StorageQuota", "Telemetry", "UpdaterCache"},
                                  winapp2ool.ScaffoldCatalogs.ElectronDefaultScaffolds)

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
        winapp2ool.ScaffoldCatalogs.ParseSection(section, catalog, New winapp2ool.MenuSection, "ElectronScaffold:")

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
        winapp2ool.ScaffoldCatalogs.ParseSection(section, catalog, New winapp2ool.MenuSection, "ElectronScaffold:")

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
        winapp2ool.ScaffoldCatalogs.ParseSection(first, catalog, New winapp2ool.MenuSection, "ElectronScaffold:")

        Dim second As New winapp2ool.iniSection2("ElectronScaffold: Caches")
        second.AddKey(New winapp2ool.iniKey2("FileKeyBase=%ElectronRoot%\two|*"))
        winapp2ool.ScaffoldCatalogs.ParseSection(second, catalog, New winapp2ool.MenuSection, "ElectronScaffold:")

        Assert.AreEqual(1, catalog.Count)
        Assert.AreEqual(1, catalog("Caches").Count)
        Assert.AreEqual("%ElectronRoot%\two|*", catalog("Caches")(0))

    End Sub

    ''' <summary>
    ''' Helper: write catalog files into a fresh temp directory and load them
    ''' </summary>
    '''
    ''' <param name="files">
    ''' Filename / file-content pairs to write into the scratch directory
    ''' </param>
    '''
    ''' <param name="warnings">
    ''' Receives the loader's captured diagnostic lines
    ''' </param>
    '''
    ''' <returns>
    ''' The loaded catalog set
    ''' </returns>
    Private Shared Function LoadDir(files As Dictionary(Of String, String),
                                    ByRef warnings As List(Of String)) As winapp2ool.ScaffoldCatalogSet

        Dim dir = IO.Path.Combine(IO.Path.GetTempPath(), "w2scaffolds_" & Guid.NewGuid().ToString("N"))
        IO.Directory.CreateDirectory(dir)

        Try

            For Each pair In files : IO.File.WriteAllText(IO.Path.Combine(dir, pair.Key), pair.Value) : Next

            Using cap = winapp2ool.gLogCapture()

                Dim loaded = winapp2ool.ScaffoldCatalogs.LoadCatalogDirectory(dir, New winapp2ool.MenuSection)
                warnings = cap.Lines.ToList()
                Return loaded

            End Using

        Finally

            IO.Directory.Delete(dir, True)

        End Try

    End Function

    ''' <summary>
    ''' The directory loader routes sections into families by the text before <c> Scaffold: </c> in
    ''' the header, not by filename. This is the property that lets a new family ship as a dropped-in
    ''' file with no settings, CLI slot, or build-script argument of its own — and it means a catalog
    ''' may be split or renamed freely.
    ''' </summary>
    <TestMethod()> Public Sub LoadCatalogDirectory_RoutesByHeaderNotFilename()

        Dim warnings As List(Of String) = Nothing

        Dim loaded = LoadDir(New Dictionary(Of String, String) From {
            {"anything.ini", "[WebViewScaffold: Caches]" & vbCrLf & "FileKeyBase=%WebViewRoot%\Cache|*" & vbCrLf & vbCrLf &
                             "[ElectronScaffold: AppLogs]" & vbCrLf & "FileKeyBase=%ElectronRoot%\logs|*" & vbCrLf}}, warnings)

        Assert.AreEqual(1, loaded.ForFamily("WebView").Count)
        Assert.AreEqual(1, loaded.ForFamily("Electron").Count)
        Assert.IsTrue(loaded.ForFamily("WebView").ContainsKey("Caches"))
        Assert.IsTrue(loaded.ForFamily("Electron").ContainsKey("AppLogs"))

    End Sub

    ''' <summary>
    ''' Several files may contribute to one family — the family's catalog is the union of every
    ''' matching section in the directory
    ''' </summary>
    <TestMethod()> Public Sub LoadCatalogDirectory_MergesFamilyAcrossFiles()

        Dim warnings As List(Of String) = Nothing

        Dim loaded = LoadDir(New Dictionary(Of String, String) From {
            {"a.ini", "[ElectronScaffold: Caches]" & vbCrLf & "FileKeyBase=%ElectronRoot%\Cache|*" & vbCrLf},
            {"b.ini", "[ElectronScaffold: AppLogs]" & vbCrLf & "FileKeyBase=%ElectronRoot%\logs|*" & vbCrLf}}, warnings)

        Assert.AreEqual(2, loaded.ForFamily("Electron").Count)

    End Sub

    ''' <summary>
    ''' A family no builder consumes warns rather than loading silently. This is the misspelling
    ''' backstop: <c> [ElctronScaffold: …] </c> parses fine as a family named <c> Elctron </c> and
    ''' would otherwise cost every one of that catalog's keys with no diagnostic at all.
    ''' </summary>
    <TestMethod()> Public Sub LoadCatalogDirectory_UnknownFamily_Warns()

        Dim warnings As List(Of String) = Nothing

        LoadDir(New Dictionary(Of String, String) From {
            {"typo.ini", "[ElctronScaffold: Caches]" & vbCrLf & "FileKeyBase=%ElectronRoot%\Cache|*" & vbCrLf}}, warnings)

        Assert.IsTrue(warnings.Any(Function(w) w.Contains("Elctron") AndAlso w.Contains("not consumed")),
                      "Expected an unconsumed-family warning, got: " & String.Join(" | ", warnings))

    End Sub

    ''' <summary>
    ''' A section with no <c> Scaffold: </c> marker cannot be assigned a family, so it is reported
    ''' rather than guessed at
    ''' </summary>
    <TestMethod()> Public Sub LoadCatalogDirectory_SectionWithoutMarker_Warns()

        Dim warnings As List(Of String) = Nothing

        Dim loaded = LoadDir(New Dictionary(Of String, String) From {
            {"stray.ini", "[Some Entry *]" & vbCrLf & "FileKey1=%AppData%\x|*" & vbCrLf}}, warnings)

        Assert.IsFalse(loaded.Families.Any())
        Assert.IsTrue(warnings.Any(Function(w) w.Contains("Unexpected section")),
                      "Expected an unexpected-section warning, got: " & String.Join(" | ", warnings))

    End Sub

    ''' <summary>
    ''' A missing scaffold directory warns and yields an empty set — a build continues with no
    ''' scaffold keys rather than throwing
    ''' </summary>
    <TestMethod()> Public Sub LoadCatalogDirectory_MissingDirectory_WarnsAndReturnsEmpty()

        Dim warnings As List(Of String)

        Using cap = winapp2ool.gLogCapture()

            Dim loaded = winapp2ool.ScaffoldCatalogs.LoadCatalogDirectory(
                IO.Path.Combine(IO.Path.GetTempPath(), "w2scaffolds_absent_" & Guid.NewGuid().ToString("N")),
                New winapp2ool.MenuSection)

            warnings = cap.Lines.ToList()
            Assert.IsFalse(loaded.Families.Any())
            Assert.AreEqual(0, loaded.ForFamily("WebView").Count)

        End Using

        Assert.IsTrue(warnings.Any(Function(w) w.Contains("Scaffold directory not found")),
                      "Expected a missing-directory warning, got: " & String.Join(" | ", warnings))

    End Sub

    ''' <summary>
    ''' A template referencing a placeholder whose binding has no roots is <em>dropped</em>. This is
    ''' the whole mechanism behind Electron's default-on <c> UpdaterCache </c> costing nothing for the
    ''' majority of entries that declare no <c> ElectronUpdaterRoot= </c>; without it a literal
    ''' <c> %ElectronUpdaterRoot% </c> would ship inside a published FileKey.
    ''' </summary>
    <TestMethod()> Public Sub FanOutPlaceholder_EmptyRoots_DropsTemplate()

        Dim result = winapp2ool.ScaffoldCatalogs.FanOutPlaceholder(
            New List(Of String) From {"%ElectronUpdaterRoot%\pending|*"},
            "%ElectronUpdaterRoot%", New List(Of String))

        Assert.AreEqual(0, result.Count)

    End Sub

    ''' <summary>
    ''' A template that does not contain the placeholder passes through exactly once, so the helper
    ''' can be chained per placeholder without multiplying placeholder-free strings
    ''' </summary>
    <TestMethod()> Public Sub FanOutPlaceholder_PlaceholderAbsent_PassesThroughOnce()

        Dim result = winapp2ool.ScaffoldCatalogs.FanOutPlaceholder(
            New List(Of String) From {"%ElectronRoot%\logs|*"},
            "%ElectronUpdaterRoot%", New List(Of String) From {"C:\a", "C:\b"})

        Assert.AreEqual(1, result.Count)
        Assert.AreEqual("%ElectronRoot%\logs|*", result(0))

    End Sub

    ''' <summary>
    ''' Chained bindings compose: a template referencing two placeholders multiplies across both
    ''' root lists rather than pairing them positionally
    ''' </summary>
    <TestMethod()> Public Sub FanOutPlaceholder_ChainedBindings_Multiply()

        Dim step1 = winapp2ool.ScaffoldCatalogs.FanOutPlaceholder(
            New List(Of String) From {"%A%-%B%"}, "%A%", New List(Of String) From {"1", "2"})

        Dim step2 = winapp2ool.ScaffoldCatalogs.FanOutPlaceholder(
            step1, "%B%", New List(Of String) From {"x", "y"})

        CollectionAssert.AreEquivalent({"1-x", "1-y", "2-x", "2-y"}, step2)

    End Sub

    ''' <summary>
    ''' <c> BindFamilyTemplates </c> reports the pre-substitution template count alongside the
    ''' substituted results, so a consumer can tell an empty scaffold (a legitimate no-op) from a
    ''' non-empty one whose every template was dropped for want of a root — only the latter is worth
    ''' warning about
    ''' </summary>
    <TestMethod()> Public Sub BindFamilyTemplates_ReportsTemplateCountSeparately()

        Dim catalog As New Dictionary(Of String, List(Of String))(StringComparer.InvariantCultureIgnoreCase) From {
            {"UpdaterCache", New List(Of String) From {"%ElectronUpdaterRoot%\pending|*"}},
            {"Empty", New List(Of String)}}

        Dim bindings As New List(Of winapp2ool.ScaffoldCatalogs.ScaffoldRootBinding) From {
            New winapp2ool.ScaffoldCatalogs.ScaffoldRootBinding("%ElectronRoot%", New List(Of String) From {"C:\ud"}),
            New winapp2ool.ScaffoldCatalogs.ScaffoldRootBinding("%ElectronUpdaterRoot%", New List(Of String))}

        Dim results = winapp2ool.ScaffoldCatalogs.BindFamilyTemplates(
            bindings, New List(Of String) From {"UpdaterCache", "Empty"}, catalog)

        Assert.AreEqual(1, results(0).TemplateCount)
        Assert.AreEqual(0, results(0).Templates.Count)
        Assert.AreEqual(0, results(1).TemplateCount)
        Assert.AreEqual(0, results(1).Templates.Count)

    End Sub

    ''' <summary>
    ''' Every family in <c> ScaffoldFamilies </c> is known to <b> both </b> builders' parsers.
    ''' <br /><br />
    '''
    ''' This is the parity guard. Electron originally shipped to EntryBuilder alone, on the theory
    ''' that MSIX packages rarely bundle Electron — which overlooked that UWPBuilder's hybrid
    ''' win32+UWP entries carry the desktop install's paths too, and a hybrid's win32 half is exactly
    ''' where Electron turns up. A family reaching one builder and not the other is not a visible
    ''' failure; it is silently missing FileKeys. This test is what makes it visible.
    ''' </summary>
    <TestMethod()> Public Sub ScaffoldFamilies_KeyVocabularyPresentInBothBuilders()

        For Each family In winapp2ool.ScaffoldCatalogs.ScaffoldFamilies

            For Each key In {$"{family}ROOT", $"{family}SCAFFOLDS", $"EXCLUDE{family}SCAFFOLDS"}

                Assert.IsTrue(winapp2ool.EntryBuilder.ReservedKeys.Any(
                                  Function(r) String.Equals(r, key, StringComparison.InvariantCultureIgnoreCase)),
                              $"EntryBuilder does not reserve '{key}=' for the {family} family")

                Assert.IsTrue(winapp2ool.UWPBuilder.UWPReservedKeys.Any(
                                  Function(r) String.Equals(r, key, StringComparison.InvariantCultureIgnoreCase)),
                              $"UWPBuilder does not reserve '{key}=' for the {family} family")

            Next

            Assert.IsTrue(winapp2ool.ScaffoldCatalogs.DefaultsForFamily(family).Any(),
                          $"The {family} family has no default scaffold set registered")

        Next

    End Sub

End Class
