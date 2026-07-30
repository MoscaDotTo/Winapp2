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

Imports System.IO

''' <summary>
''' Shared, use-case-agnostic substrate for embedded-Chromium FileKey emission. Consumed
''' by entry generators (UWPBuilder, EntryBuilder) that need a curated, consistent set of
''' cleaning targets for an application's embedded browser-engine data folder.
''' <br /><br />
'''
''' The module is engine-family-neutral: a family is identified by the text preceding
''' <c> Scaffold: </c> in a catalog section header, so one directory scan
''' (<see cref="LoadCatalogDirectory"/>) discovers every family present without any
''' per-family plumbing. <see cref="ScaffoldFamilies"/> is the roster of families the
''' builders actually bind:
''' <list type="bullet">
''' <item><c> WebView </c> — WebView2 / EBWebView layout (<c> Assembler\Scaffolds\webview.ini </c>),
''' templates use the <c> %WebViewRoot% </c> placeholder.</item>
''' <item><c> QtWebEngine </c> — QtWebEngine layout (<c> Assembler\Scaffolds\qtwebengine.ini </c>),
''' templates use the <c> %QtWebEngineRoot% </c> placeholder. QtWebEngine bundles an older
''' Chromium with a flatter on-disk layout (no <c> Default\Network\ </c> subfolder, no
''' privacy-sandbox surface), so it warrants a separate catalog rather than reusing the
''' WebView2 paths.</item>
''' <item><c> Electron </c> — Electron layout (<c> Assembler\Scaffolds\electron.ini </c>),
''' templates use <c> %ElectronRoot% </c> (the app's <c> userData </c> folder, which for
''' Electron <em>is</em> the Chromium profile — no <c> Default\ </c> segment) and
''' <c> %ElectronUpdaterRoot% </c> (the electron-updater download cache, whose directory name
''' is not derivable from the userData path). This is the only family with two placeholders,
''' so its consumers bind both — see <see cref="BindFamilyTemplates"/>. It is a separate
''' catalog because Electron collapses WebView2's user-data and profile levels into one
''' directory <em>and</em> straddles the Chromium 89 network-service migration, needing both
''' <c> Network\Cookies </c> and legacy root-level <c> Cookies </c>.</item>
''' </list>
''' <br />
'''
''' <b> Parity contract: every family in <see cref="ScaffoldFamilies"/> ships to every
''' consumer. </b> Electron originally shipped to EntryBuilder alone on the theory that MSIX
''' packages rarely bundle Electron, which overlooked that UWPBuilder's hybrid win32+UWP
''' entries carry the win32 install's paths too — and a hybrid's win32 half is precisely where
''' Electron turns up. A scaffold family lands in both builders or neither;
''' <c> ScaffoldParityTests </c> enforces it against both parsers' reserved-key sets.
''' <br /><br />
'''
''' A family may bind more than one placeholder. Consumers expand a template by chaining one
''' substitution per (placeholder, roots) pair; a template whose placeholder has an empty root
''' list is dropped rather than emitted with the placeholder left literal, which is what lets
''' the Electron family's updater templates stay inert for an entry that declared only
''' <c> ElectronRoot= </c>.
''' <br />
'''
''' Source files contribute <c> [{Family}Scaffold: Name] </c> sections whose
''' <c> FileKeyBase= </c> values use a family-specific root placeholder. Callers load a whole
''' scaffold directory via <see cref="LoadCatalogDirectory"/>, resolve the active selection
''' per entry via <see cref="ResolveScaffolds"/>, substitute their own host-specific root
''' paths via <see cref="BindFamilyTemplates"/>, and then apply whatever further expansion
''' their own DSL provides.
''' <br /><br />
'''
''' The placeholder is implementation-neutral by name: a UWP consumer feeds it the
''' <c> %LocalAppData%\Packages\...\EBWebView </c> form, while a win32 consumer would feed
''' the host-relative data folder under <c> %LocalAppData% </c>, <c> %ProgramData% </c>, or
''' wherever the host stores its data. The library is agnostic to root-path conventions;
''' callers handle their own variable expansion before passing the resolved root to the
''' emitter.
''' </summary>
Public Module ScaffoldCatalogs

    ''' <summary>
    ''' The literal marker separating a family name from a scaffold name in a catalog section
    ''' header — <c> [WebViewScaffold: Caches] </c> yields family <c> WebView </c>, scaffold
    ''' <c> Caches </c>. Family identity therefore comes from the section header rather than
    ''' the filename, so a catalog may be split or merged across files freely.
    ''' </summary>
    Public Const FamilyMarker As String = "Scaffold:"

    ''' <summary>
    ''' The scaffold families the builders bind. A section in a scaffold directory naming a
    ''' family outside this roster is loaded but warned about, since nothing will consume it
    ''' — which is what catches a misspelled family (<c> [ElctronScaffold: …] </c>) and a
    ''' catalog file added ahead of the code that reads it.
    ''' <br /><br />
    '''
    ''' Adding a family here is not sufficient to ship it: each builder must also reserve the
    ''' family's key vocabulary (<c> {Family}Root= </c>, <c> {Family}Scaffolds= </c>,
    ''' <c> Exclude{Family}Scaffolds= </c>) and bind its placeholders. That is deliberately
    ''' load-bearing — the parity test walks this array and asserts both parsers know every
    ''' name in it.
    ''' </summary>
    Public ReadOnly ScaffoldFamilies As String() = {"WebView", "QtWebEngine", "Electron"}

    ''' <summary>
    ''' Default scaffold names emitted when a caller requests WebView2 scaffolding without
    ''' an explicit selection list. Limited to low-risk categories that are safe for any
    ''' WebView-hosting application; host-risk categories (cookies, history, session,
    ''' web storage) require explicit opt-in by the caller.
    ''' </summary>
    Public ReadOnly DefaultScaffolds As String() = {"Caches", "Telemetry"}

    ''' <summary>
    ''' Default scaffold names emitted when a caller requests QtWebEngine scaffolding
    ''' without an explicit selection list. Broader than <see cref="DefaultScaffolds"/>:
    ''' QtWebEngine's catalog splits two low-risk targets into their own scaffolds that
    ''' the WebView2 catalog leaves bundled behind host-risk gates — <c> StorageQuota </c>
    ''' (a rebuildable quota accounting file) and <c> VisitedLinks </c> (a link-coloring
    ''' bloom filter, which QtWebEngine-hosting application shells have no UI for). Both
    ''' are default-on here because the hand-written entries this catalog replaces treated
    ''' them as routine cleaning. Host-risk categories (cookies, history, session, site
    ''' storage) still require explicit opt-in.
    ''' </summary>
    Public ReadOnly QtWebEngineDefaultScaffolds As String() = {"Caches", "StorageQuota", "Telemetry", "VisitedLinks"}

    ''' <summary>
    ''' Default scaffold names emitted when a caller requests Electron scaffolding without an
    ''' explicit selection list. Wider than the other two families because an Electron
    ''' <c> userData </c> folder holds application diagnostics alongside Chromium state:
    ''' <c> AppLogs </c> covers electron-log's output (both the modern <c> logs\ </c> and the
    ''' legacy <c> &lt;userData&gt;\&lt;ProductName&gt;\logs\ </c>), and <c> UpdaterCache </c>
    ''' covers the electron-updater download cache — the latter costing nothing for an entry
    ''' that declared no <c> ElectronUpdaterRoot= </c>, since its templates are then dropped.
    ''' <br /><br />
    '''
    ''' Note the deliberate asymmetry with <see cref="QtWebEngineDefaultScaffolds"/>: there is
    ''' no <c> VisitedLinks </c> scaffold here. A 21-installation disk survey found
    ''' <c> Visited Links </c> in none of them, so the pattern rides in <c> Telemetry </c>'s
    ''' legacy list rather than earning a scaffold of its own. Host-risk categories (cookies,
    ''' site storage) still require explicit opt-in.
    ''' </summary>
    Public ReadOnly ElectronDefaultScaffolds As String() = {"AppLogs", "Caches", "StorageQuota", "Telemetry", "UpdaterCache"}

    ''' <summary>
    ''' Returns the default scaffold set for <paramref name="familyLabel"/>, or an empty set
    ''' for a family with no defaults registered. Lets a consumer drive all of
    ''' <see cref="ScaffoldFamilies"/> from one loop rather than a per-family
    ''' <c> Select Case </c> that can silently omit a family.
    ''' </summary>
    '''
    ''' <param name="familyLabel">
    ''' The family token, e.g. <c> WebView </c> or <c> Electron </c>
    ''' </param>
    '''
    ''' <returns>
    ''' The family's default scaffold names, or an empty array when the family is unknown
    ''' </returns>
    Public Function DefaultsForFamily(familyLabel As String) As String()

        Select Case familyLabel.ToUpperInvariant()

            Case "WEBVIEW" : Return DefaultScaffolds

            Case "QTWEBENGINE" : Return QtWebEngineDefaultScaffolds

            Case "ELECTRON" : Return ElectronDefaultScaffolds

            Case Else : Return New String() {}

        End Select

    End Function

    ''' <summary>
    ''' One (placeholder, roots) pair for a scaffold family. WebView and QtWebEngine bind a
    ''' single pair each, Electron binds two (<c> %ElectronRoot% </c> and
    ''' <c> %ElectronUpdaterRoot% </c>), and <see cref="BindFamilyTemplates"/> chains a
    ''' substitution per binding so a template referencing several placeholders multiplies
    ''' across all of their root lists. Binding a placeholder to an empty root list is
    ''' meaningful: it is what keeps the Electron <c> UpdaterCache </c> scaffold inert for
    ''' entries that declared no <c> ElectronUpdaterRoot= </c> rather than emitting a literal
    ''' placeholder into a FileKey.
    ''' </summary>
    Public Structure ScaffoldRootBinding

        ''' <summary> The literal placeholder token, e.g. <c> %ElectronRoot% </c> </summary>
        Public Placeholder As String

        ''' <summary> The roots substituted for the placeholder, possibly empty </summary>
        Public Roots As List(Of String)

        ''' <summary>
        ''' Creates a binding pairing one placeholder with the roots that replace it
        ''' </summary>
        '''
        ''' <param name="placeholder">
        ''' The literal placeholder token appearing in the family's catalog templates
        ''' </param>
        '''
        ''' <param name="roots">
        ''' The declared roots for that placeholder, already fully resolved by the caller
        ''' </param>
        Public Sub New(placeholder As String, roots As List(Of String))

            Me.Placeholder = placeholder
            Me.Roots = roots

        End Sub

    End Structure

    ''' <summary>
    ''' One scaffold's root-substituted templates, as produced by
    ''' <see cref="BindFamilyTemplates"/>. Results are returned per scaffold rather than as one
    ''' flat list so a consumer can run its own further expansion over each group and still
    ''' report which named scaffold produced nothing.
    ''' </summary>
    Public Structure ScaffoldBindingResult

        ''' <summary> The scaffold's name as it appears in the catalog </summary>
        Public ScaffoldName As String

        ''' <summary>
        ''' How many <c> FileKeyBase= </c> templates the scaffold declared, before substitution.
        ''' Zero means an empty scaffold — a legitimate no-op that consumers must not confuse
        ''' with a non-empty scaffold whose every template was dropped for want of a root.
        ''' </summary>
        Public TemplateCount As Integer

        ''' <summary>
        ''' The root-substituted templates. Shorter than <see cref="TemplateCount"/> whenever a
        ''' template referenced a placeholder bound to no roots, and longer whenever a template
        ''' fanned out across several roots.
        ''' </summary>
        Public Templates As List(Of String)

        ''' <summary>
        ''' Creates a result pairing a scaffold with the templates its bindings produced
        ''' </summary>
        '''
        ''' <param name="scaffoldName">
        ''' The scaffold's catalog name
        ''' </param>
        '''
        ''' <param name="templateCount">
        ''' The count of templates the scaffold declared before substitution
        ''' </param>
        '''
        ''' <param name="templates">
        ''' The root-substituted templates
        ''' </param>
        Public Sub New(scaffoldName As String, templateCount As Integer, templates As List(Of String))

            Me.ScaffoldName = scaffoldName
            Me.TemplateCount = templateCount
            Me.Templates = templates

        End Sub

    End Structure

    ''' <summary>
    ''' Parses a <c> [{prefix} ...] </c> scaffold section, collecting its
    ''' <c> FileKeyBase= </c> values into <paramref name="scaffolds"/> keyed by the
    ''' scaffold's name (the portion of the section header after
    ''' <paramref name="sectionPrefix"/>, trimmed). Warns on unrecognised key types and on
    ''' duplicate scaffold names (last definition wins).
    ''' </summary>
    '''
    ''' <param name="scaffoldSection">
    ''' The scaffold section to parse. The caller is responsible for filtering sections by
    ''' name prefix before invoking this routine.
    ''' </param>
    '''
    ''' <param name="scaffolds">
    ''' The accumulator dictionary, keyed by scaffold name (case-insensitive). The value is
    ''' the ordered list of <c> FileKeyBase= </c> templates for that scaffold; templates
    ''' may contain the family's root placeholder for later substitution by the caller.
    ''' </param>
    '''
    ''' <param name="menuOutput">
    ''' The <c> MenuSection </c> receiving progress lines and warnings for display
    ''' </param>
    '''
    ''' <param name="sectionPrefix">
    ''' The section-header prefix that identifies this family's scaffolds (e.g.
    ''' <c> WebViewScaffold: </c> or <c> QtWebEngineScaffold: </c>). Stripped from the
    ''' header to derive the scaffold name and used as the human label in diagnostics.
    ''' </param>
    Public Sub ParseSection(scaffoldSection As iniSection2,
                            scaffolds As Dictionary(Of String, List(Of String)),
                            menuOutput As MenuSection,
                            Optional sectionPrefix As String = "WebViewScaffold:")

        Dim sectionLabel = sectionPrefix.TrimEnd(":"c)
        Dim scaffoldName = scaffoldSection.Name.Substring(sectionPrefix.Length).Trim()
        Dim scaffoldMsg = $"Processing {sectionLabel}: {scaffoldName}"
        menuOutput.AddColoredLine(scaffoldMsg, ConsoleColor.Magenta)
        gLog($"  {scaffoldMsg}")

        If scaffolds.ContainsKey(scaffoldName) Then

            Dim dupMsg = $"Duplicate {sectionLabel} name '{scaffoldName}'; last definition wins"
            gLog(dupMsg)
            menuOutput.AddWarning(dupMsg)

        End If

        Dim keys As New List(Of String)

        For Each key In scaffoldSection.Keys

            Select Case key.KeyType.ToUpperInvariant()

                Case "FILEKEYBASE" : keys.Add(key.Value)

                Case Else

                    Dim errMsg = $"Unexpected key in {sectionLabel} [{scaffoldSection.Name}]: {key.Name}"
                    gLog(errMsg)
                    menuOutput.AddWarning(errMsg)

            End Select

        Next

        scaffolds(scaffoldName) = keys

    End Sub

    ''' <summary>
    ''' Loads a complete scaffold catalog from a stand-alone ini file. Each section whose
    ''' header starts with <paramref name="sectionPrefix"/> in <paramref name="catalogPath"/>
    ''' is parsed via <see cref="ParseSection"/>; any other section type is reported as
    ''' unexpected. Missing or empty catalog files yield an empty dictionary and a warning
    ''' rather than an exception, so callers can continue with no scaffolds rather than
    ''' aborting the run.
    ''' <br /><br />
    '''
    ''' Single-family loading of one named file. The builders instead load a whole scaffold
    ''' directory via <see cref="LoadCatalogDirectory"/>, which discovers families rather than
    ''' being told one; this remains the primitive that does so and the entry point for a
    ''' caller that genuinely has one file and one family in hand.
    ''' </summary>
    '''
    ''' <param name="catalogPath">
    ''' Absolute path to the catalog file (e.g. <c> Assembler\Scaffolds\webview.ini </c> or
    ''' <c> Assembler\Scaffolds\qtwebengine.ini </c>)
    ''' </param>
    '''
    ''' <param name="menuOutput">
    ''' The <c> MenuSection </c> receiving progress lines and warnings for display
    ''' </param>
    '''
    ''' <param name="sectionPrefix">
    ''' The section-header prefix identifying this family's scaffolds. Defaults to
    ''' <c> WebViewScaffold: </c>; QtWebEngine callers pass <c> QtWebEngineScaffold: </c>.
    ''' </param>
    '''
    ''' <returns>
    ''' A case-insensitive dictionary keyed by scaffold name. Empty scaffold sections register
    ''' with an empty template list — selecting them is a no-op rather than an error.
    ''' </returns>
    Public Function LoadCatalog(catalogPath As String,
                                menuOutput As MenuSection,
                                Optional sectionPrefix As String = "WebViewScaffold:") As Dictionary(Of String, List(Of String))

        Dim catalog As New Dictionary(Of String, List(Of String))(StringComparer.InvariantCultureIgnoreCase)
        Dim sectionLabel = sectionPrefix.TrimEnd(":"c)

        Dim catalogIni = iniFile2.FromFile(catalogPath)

        If catalogIni.Count = 0 Then

            Dim emptyMsg = $"{sectionLabel} catalog at {catalogPath} is empty or missing"
            gLog(emptyMsg)
            menuOutput.AddWarning(emptyMsg)
            Return catalog

        End If

        For Each section In catalogIni

            If section.Name.StartsWith(sectionPrefix, StringComparison.InvariantCulture) Then

                ParseSection(section, catalog, menuOutput, sectionPrefix)

            Else

                Dim unexpectedMsg = $"Unexpected section in {sectionLabel} catalog: [{section.Name}]"
                gLog(unexpectedMsg)
                menuOutput.AddWarning(unexpectedMsg)

            End If

        Next

        Return catalog

    End Function

    ''' <summary>
    ''' Loads every scaffold catalog in <paramref name="scaffoldDir"/> into one
    ''' <see cref="ScaffoldCatalogSet"/>, routing each section to a family by the text
    ''' preceding <see cref="FamilyMarker"/> in its header. Families are therefore discovered
    ''' from the source data rather than declared by the caller, so a new catalog file needs no
    ''' new setting, CLI slot, or build-script argument — dropping <c> tauri.ini </c> into the
    ''' directory registers a <c> Tauri </c> family, and the only remaining work is the
    ''' consuming builder's key vocabulary.
    ''' <br /><br />
    '''
    ''' Files are read in sorted order so the merge is deterministic, and several files may
    ''' contribute to one family — a family's catalog is the union of every
    ''' <c> [{Family}Scaffold: …] </c> section in the directory, with
    ''' <see cref="ParseSection"/>'s last-definition-wins rule spanning files.
    ''' <br /><br />
    '''
    ''' Nothing here is fatal. A missing directory, an unreadable section, or a family nothing
    ''' consumes all warn and continue, matching <see cref="LoadCatalog"/>'s contract that a
    ''' catalog problem costs scaffold keys rather than aborting a build.
    ''' </summary>
    '''
    ''' <param name="scaffoldDir">
    ''' Absolute path to the scaffold directory (typically <c> Assembler\Scaffolds </c>)
    ''' </param>
    '''
    ''' <param name="menuOutput">
    ''' The <c> MenuSection </c> receiving progress lines and warnings for display
    ''' </param>
    '''
    ''' <returns>
    ''' The loaded families. Empty when the directory is missing or holds no catalogs.
    ''' </returns>
    Public Function LoadCatalogDirectory(scaffoldDir As String,
                                         menuOutput As MenuSection) As ScaffoldCatalogSet

        Dim catalogs As New ScaffoldCatalogSet

        If Not Directory.Exists(scaffoldDir) Then

            Dim missingMsg = $"Scaffold directory not found: {scaffoldDir}; no scaffold keys will be emitted"
            gLog(missingMsg)
            menuOutput.AddWarning(missingMsg)
            Return catalogs

        End If

        Dim files = Directory.GetFiles(scaffoldDir, "*.ini", SearchOption.TopDirectoryOnly).ToList()
        files.Sort()

        For Each filePath In files

            For Each section In iniFile2.FromFile(filePath)

                Dim markerAt = section.Name.IndexOf(FamilyMarker, StringComparison.InvariantCultureIgnoreCase)

                ' A header with no marker at all is the misspelling case ([ElectronScafold: …])
                ' or a stray section someone parked in the scaffold directory. Either way no
                ' family can be derived from it, so it is reported rather than guessed at.
                If markerAt <= 0 Then

                    Dim unexpectedMsg = $"Unexpected section in scaffold catalog {IO.Path.GetFileName(filePath)}: [{section.Name}]"
                    gLog(unexpectedMsg)
                    menuOutput.AddWarning(unexpectedMsg)
                    Continue For

                End If

                Dim family = section.Name.Substring(0, markerAt).Trim()
                ParseSection(section, catalogs.ForFamily(family), menuOutput, family & FamilyMarker)

            Next

        Next

        ' Reported after the scan rather than per-section so a whole unconsumed catalog file
        ' produces one warning instead of one per scaffold in it.
        For Each family In catalogs.Families

            If ScaffoldFamilies.Any(Function(f) String.Equals(f, family, StringComparison.InvariantCultureIgnoreCase)) Then Continue For

            Dim unknownMsg = $"Scaffold family '{family}' in {scaffoldDir} is not consumed by any builder; check the spelling of its section headers"
            gLog(unknownMsg)
            menuOutput.AddWarning(unknownMsg)

        Next

        Return catalogs

    End Function

    ''' <summary>
    ''' Resolves the set of scaffold names a caller should emit for one entry, applying the
    ''' shared selection grammar used by every scaffold family: explicit selection, the
    ''' <c> All </c> sentinel, default fallback, exclusions, and unknown-name filtering.
    ''' Family-agnostic — callers supply the family's default set and a label used to phrase
    ''' diagnostics (e.g. <c> WebView </c> → <c> WebViewScaffolds=All </c>).
    ''' </summary>
    '''
    ''' <param name="scaffoldNames">
    ''' The explicit selection list parsed from the entry's <c> {Family}Scaffolds= </c> key.
    ''' Ignored when <paramref name="keyPresent"/> is False.
    ''' </param>
    '''
    ''' <param name="keyPresent">
    ''' Whether the entry declared <c> {Family}Scaffolds= </c> at all. Distinguishes
    ''' "key absent → use <paramref name="defaultSet"/>" from "key present but empty → emit
    ''' nothing".
    ''' </param>
    '''
    ''' <param name="excluded">
    ''' Names parsed from <c> Exclude{Family}Scaffolds= </c>, subtracted from the active
    ''' selection after it is determined.
    ''' </param>
    '''
    ''' <param name="available">
    ''' The catalog of known scaffolds for this family, keyed by name.
    ''' </param>
    '''
    ''' <param name="defaultSet">
    ''' The default scaffold names used when the entry did not declare an explicit selection.
    ''' </param>
    '''
    ''' <param name="familyLabel">
    ''' Family token used to phrase diagnostics — <c> WebView </c> or <c> QtWebEngine </c>.
    ''' Produces messages like <c> {familyLabel}Scaffolds=All </c> and
    ''' <c> Unknown {familyLabel} scaffold </c>.
    ''' </param>
    '''
    ''' <param name="specName">
    ''' The entry name, embedded in diagnostics for source localisation.
    ''' </param>
    '''
    ''' <param name="menuOutput">
    ''' The <c> MenuSection </c> receiving warnings for display.
    ''' </param>
    '''
    ''' <returns>
    ''' The ordered list of scaffold names to emit for this entry, with unknown names removed.
    ''' </returns>
    Public Function ResolveScaffolds(scaffoldNames As List(Of String),
                                     keyPresent As Boolean,
                                     excluded As List(Of String),
                                     available As Dictionary(Of String, List(Of String)),
                                     defaultSet As String(),
                                     familyLabel As String,
                                     specName As String,
                                     menuOutput As MenuSection) As List(Of String)

        Dim selected As List(Of String)
        Dim usedAllSentinel As Boolean = False

        If keyPresent Then

            usedAllSentinel = scaffoldNames.Any(
                Function(s) String.Equals(s, "All", StringComparison.InvariantCultureIgnoreCase))

            If usedAllSentinel Then

                selected = New List(Of String)(available.Keys)

                Dim allMsg = $"{familyLabel}Scaffolds=All in [{specName}]; expanded to {selected.Count} scaffold(s) from catalog"
                gLog(allMsg)

                Dim extras = scaffoldNames.Where(
                    Function(s) Not String.Equals(s, "All", StringComparison.InvariantCultureIgnoreCase)).ToList()

                If extras.Count > 0 Then

                    Dim extraMsg = $"{familyLabel}Scaffolds=All in [{specName}] with redundant additional names ({String.Join(", ", extras)}); ignoring"
                    gLog(extraMsg)
                    menuOutput.AddWarning(extraMsg)

                End If

            Else

                selected = New List(Of String)(scaffoldNames)

            End If

            If excluded.Count > 0 AndAlso Not usedAllSentinel Then

                Dim mixMsg = $"Both {familyLabel}Scaffolds and Exclude{familyLabel}Scaffolds set in [{specName}]; applying exclusions to explicit list"
                gLog(mixMsg)
                menuOutput.AddWarning(mixMsg)

            End If

        Else

            selected = New List(Of String)(defaultSet)

        End If

        For Each ex In excluded

            selected.RemoveAll(Function(s) String.Equals(s, ex, StringComparison.InvariantCultureIgnoreCase))

        Next

        Dim valid As New List(Of String)

        For Each name In selected

            If available.ContainsKey(name) Then

                valid.Add(name)

            Else

                Dim unknownMsg = $"Unknown {familyLabel} scaffold '{name}' requested by [{specName}], skipping"
                gLog(unknownMsg)
                menuOutput.AddWarning(unknownMsg)

            End If

        Next

        Return valid

    End Function

    ''' <summary>
    ''' Substitutes a family's declared roots into the catalog templates of every selected
    ''' scaffold, chaining one <see cref="FanOutPlaceholder"/> pass per binding. A template
    ''' referencing several placeholders multiplies across all of their root lists; a template
    ''' referencing a placeholder whose binding has no roots is <em>dropped</em> rather than
    ''' emitted with the placeholder left literal. That drop rule is what makes the Electron
    ''' family's two-placeholder catalog work — an entry declaring only <c> ElectronRoot= </c>
    ''' silently contributes nothing from the <c> UpdaterCache </c> scaffold.
    ''' <br /><br />
    '''
    ''' Results are grouped per scaffold, carrying the pre-substitution template count, so a
    ''' consumer can run its own further expansion over each group and still tell an empty
    ''' scaffold (a legitimate no-op) from a non-empty one that produced nothing because its
    ''' templates all wanted a root the entry never declared. Consumers own that diagnostic
    ''' because only they know whether the selection was explicit — warning when a default-set
    ''' member yields nothing would fire on nearly every Electron entry.
    ''' </summary>
    '''
    ''' <param name="bindings">
    ''' The family's (placeholder, roots) pairs — one for WebView and QtWebEngine, two for Electron
    ''' </param>
    '''
    ''' <param name="selectedScaffolds">
    ''' The resolved scaffold names to emit, already filtered to catalog members by
    ''' <see cref="ResolveScaffolds"/>
    ''' </param>
    '''
    ''' <param name="catalog">
    ''' The scaffold catalog for this family, keyed by scaffold name
    ''' </param>
    '''
    ''' <returns>
    ''' One result per selected scaffold, in selection order
    ''' </returns>
    Public Function BindFamilyTemplates(bindings As List(Of ScaffoldRootBinding),
                                        selectedScaffolds As List(Of String),
                                        catalog As Dictionary(Of String, List(Of String))) As List(Of ScaffoldBindingResult)

        Dim results As New List(Of ScaffoldBindingResult)

        For Each scaffoldName In selectedScaffolds

            Dim templates = catalog(scaffoldName)
            Dim bound As New List(Of String)

            For Each template In templates

                Dim rootExpanded As New List(Of String) From {template}

                For Each binding In bindings

                    rootExpanded = FanOutPlaceholder(rootExpanded, binding.Placeholder, binding.Roots)

                Next

                bound.AddRange(rootExpanded)

            Next

            results.Add(New ScaffoldBindingResult(scaffoldName, templates.Count, bound))

        Next

        Return results

    End Function

    ''' <summary>
    ''' Replaces <paramref name="placeholder"/> in each input template with every root in
    ''' <paramref name="roots"/>, producing one output per (template, root) pair. Templates
    ''' that do not contain the placeholder pass through unchanged (one output each), so the
    ''' helper can be chained per placeholder without dropping placeholder-free strings.
    ''' </summary>
    '''
    ''' <param name="templates">
    ''' The current working set of (possibly already partially expanded) templates
    ''' </param>
    '''
    ''' <param name="placeholder">
    ''' The literal placeholder token to substitute (e.g. <c> %WebViewRoot% </c>)
    ''' </param>
    '''
    ''' <param name="roots">
    ''' The root paths to substitute for the placeholder
    ''' </param>
    '''
    ''' <returns>
    ''' The expanded set; a template containing the placeholder is dropped when
    ''' <paramref name="roots"/> is empty (matching the legacy behaviour of emitting nothing
    ''' for a placeholder with no declared root)
    ''' </returns>
    Public Function FanOutPlaceholder(templates As List(Of String),
                                      placeholder As String,
                                      roots As List(Of String)) As List(Of String)

        Dim result As New List(Of String)

        For Each t In templates

            If t.Contains(placeholder) Then

                For Each root In roots
                    result.Add(t.Replace(placeholder, root))
                Next

            Else

                result.Add(t)

            End If

        Next

        Return result

    End Function

End Module
