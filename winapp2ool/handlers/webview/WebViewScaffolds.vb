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
''' Shared, use-case-agnostic substrate for embedded-Chromium FileKey emission. Consumed
''' by entry generators (UWPBuilder, EntryBuilder) that need a curated, consistent set of
''' cleaning targets for an application's embedded WebView2 / EBWebView or QtWebEngine
''' data folder.
''' <br /><br />
'''
''' The module is engine-family-neutral: a single section-prefix parameter selects which
''' family of scaffold sections to read from a catalog, so multiple catalogs can be
''' processed by the same loader. The shipping families are:
''' <list type="bullet">
''' <item><c> WebViewScaffold </c> — WebView2 / EBWebView layout (<c> Assembler\Scaffolds\webview.ini </c>),
''' templates use the <c> %WebViewRoot% </c> placeholder.</item>
''' <item><c> QtWebEngineScaffold </c> — QtWebEngine layout (<c> Assembler\Scaffolds\qtwebengine.ini </c>),
''' templates use the <c> %QtWebEngineRoot% </c> placeholder. QtWebEngine bundles an older
''' Chromium with a flatter on-disk layout (no <c> Default\Network\ </c> subfolder, no
''' privacy-sandbox surface), so it warrants a separate catalog rather than reusing the
''' WebView2 paths.</item>
''' </list>
''' <br />
'''
''' Source files contribute <c> [{Family}: Name] </c> sections whose <c> FileKeyBase= </c>
''' values use a family-specific root placeholder. Callers parse their own source files,
''' delegate scaffold sections to <see cref="ParseSection"/> (or load a whole catalog via
''' <see cref="LoadCatalog"/>), resolve the active selection via <see cref="ResolveScaffolds"/>,
''' and later substitute their own host-specific root paths into the resulting templates.
''' <br /><br />
'''
''' The placeholder is implementation-neutral by name: a UWP consumer feeds it the
''' <c> %LocalAppData%\Packages\...\EBWebView </c> form, while a win32 consumer would feed
''' the host-relative data folder under <c> %LocalAppData% </c>, <c> %ProgramData% </c>, or
''' wherever the host stores its data. The library is agnostic to root-path conventions;
''' callers handle their own variable expansion before passing the resolved root to the
''' emitter.
''' </summary>
Public Module WebViewScaffolds

    ''' <summary>
    ''' Default scaffold names emitted when a caller requests WebView2 scaffolding without
    ''' an explicit selection list. Limited to low-risk categories that are safe for any
    ''' WebView-hosting application; host-risk categories (cookies, history, session,
    ''' web storage) require explicit opt-in by the caller.
    ''' </summary>
    Public ReadOnly DefaultScaffolds As String() = {"Caches", "Telemetry"}

    ''' <summary>
    ''' Default scaffold names emitted when a caller requests QtWebEngine scaffolding
    ''' without an explicit selection list. Mirrors <see cref="DefaultScaffolds"/> — the
    ''' same low-risk Caches + Telemetry baseline — but kept as a separate constant so the
    ''' two families can diverge as QtWebEngine's catalog evolves independently.
    ''' </summary>
    Public ReadOnly QtWebEngineDefaultScaffolds As String() = {"Caches", "Telemetry"}

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

End Module
