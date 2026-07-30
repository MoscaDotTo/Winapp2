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

Imports System.Text

''' <summary>
''' Tests for UWPBuilder's integration of the shared <c> VariableExpander </c> substrate:
''' the three-phase expansion order (<c> %Package% </c> → root placeholder →
''' <c> &lt;var&gt; </c>), per-key-class domain classification, the open-vocabulary parser
''' and its typo backstops, and the vocabulary shared with EntryBuilder. The expansion
''' engine itself is covered by <c> VariableExpanderTests </c>; these tests cover the wiring.
''' </summary>
<TestClass()> Public Class UWPBuilderTests

    ''' <summary>
    ''' Helper: parse literal ini text and return its first section
    ''' </summary>
    Private Shared Function FirstSection(text As String) As winapp2ool.iniSection2

        Dim bytes = Encoding.UTF8.GetBytes(text)
        Using ms As New IO.MemoryStream(bytes)
            Using reader As New IO.StreamReader(ms)
                Dim parsed = winapp2ool.iniFile2.FromStream(reader, "", "test.ini")
                For Each section In parsed
                    Return section
                Next
            End Using
        End Using

        Throw New InvalidOperationException("No section parsed from test input")

    End Function

    ''' <summary>
    ''' Helper: run an AppInfo section through the parser, capturing gLog diagnostics
    ''' </summary>
    Private Shared Function ParseApp(text As String,
                                     ByRef diagnostics As List(Of String)) As winapp2ool.UWPBuilder.UWPAppInfo

        Dim menu As New winapp2ool.MenuSection
        Using cap = winapp2ool.gLogCapture()

            Dim app = winapp2ool.UWPBuilder.parseAppInfo(FirstSection(text), menu)
            diagnostics = cap.Lines.ToList()
            Return app

        End Using

    End Function

    ''' <summary>
    ''' Helper: parse then generate, with optional scaffold inputs, returning the emitted section
    ''' </summary>
    Private Shared Function Build(text As String,
                         Optional scaffoldFileKeys As List(Of String) = Nothing,
                         Optional scaffoldDetectFiles As List(Of String) = Nothing,
                         Optional webViewCatalog As Dictionary(Of String, List(Of String)) = Nothing,
                         Optional qtCatalog As Dictionary(Of String, List(Of String)) = Nothing,
                         Optional electronCatalog As Dictionary(Of String, List(Of String)) = Nothing) As winapp2ool.iniSection2

        Dim menu As New winapp2ool.MenuSection
        Using cap = winapp2ool.gLogCapture()

            Dim app = winapp2ool.UWPBuilder.parseAppInfo(FirstSection(text), menu)

            Dim catalogs As New winapp2ool.ScaffoldCatalogSet
            CopyCatalog(webViewCatalog, catalogs.ForFamily("WebView"))
            CopyCatalog(qtCatalog, catalogs.ForFamily("QtWebEngine"))
            CopyCatalog(electronCatalog, catalogs.ForFamily("Electron"))

            Return winapp2ool.UWPBuilder.generateUWPEntry(app,
                If(scaffoldFileKeys, New List(Of String)),
                If(scaffoldDetectFiles, New List(Of String)),
                catalogs,
                menu)

        End Using

    End Function

    ''' <summary>
    ''' Helper: fill one family's catalog inside a <c> ScaffoldCatalogSet </c> from a plain
    ''' dictionary, so tests can keep expressing catalogs as literals
    ''' </summary>
    Private Shared Sub CopyCatalog(source As Dictionary(Of String, List(Of String)),
                                   target As Dictionary(Of String, List(Of String)))

        If source Is Nothing Then Return

        For Each pair In source : target(pair.Key) = pair.Value : Next

    End Sub

    ''' <summary>
    ''' Helper: collect the values of every key of a given type from a generated section
    ''' </summary>
    Private Shared Function ValuesOf(section As winapp2ool.iniSection2, keyType As String) As List(Of String)

        Dim result As New List(Of String)

        For Each key In section.Keys

            If String.Equals(key.KeyType, keyType, StringComparison.InvariantCultureIgnoreCase) Then result.Add(key.Value)

        Next

        Return result

    End Function

    ''' <summary>
    ''' Minimal valid AppInfo preamble — a package and a category, so the entry is not skipped
    ''' </summary>
    Private Const Preamble As String = "[Test App *]" & vbCrLf &
                                       "Package=Contoso.App_8wekyb3d8bbwe" & vbCrLf &
                                       "LangSecRef=3021" & vbCrLf

    ' ----- Open vocabulary -----

    ''' <summary>
    ''' An unrecognised key is a variable declaration, not an error. This is the change that
    ''' replaces the parser's former "Unexpected key type" warning
    ''' </summary>
    <TestMethod()> Public Sub UnknownKey_BecomesVariableDeclaration()

        Dim diags As List(Of String) = Nothing
        Dim app = ParseApp(Preamble & "Version=11.0,16.0" & vbCrLf, diags)

        Assert.IsTrue(app.Variables.IsDeclared("Version"), "Version= should be declared as a variable")
        Assert.AreEqual(2, app.Variables.Values("Version").Count)
        Assert.IsFalse(diags.Any(Function(d) d.Contains("Unexpected key type")),
                       "Open vocabulary must not warn on unknown keys: " & String.Join("; ", diags))

    End Sub

    ''' <summary>
    ''' A declared variable that nothing references is the typo backstop for a misspelled
    ''' reserved key (e.g. Pakcage=), replacing the lost unknown-key warning
    ''' </summary>
    <TestMethod()> Public Sub UnreferencedVariable_IsReportable()

        Dim diags As List(Of String) = Nothing
        Dim app = ParseApp(Preamble & "Pakcage=Contoso.Typo_8wekyb3d8bbwe" & vbCrLf, diags)

        Dim unused = app.Variables.UnreferencedNames()

        Assert.AreEqual(1, unused.Count)
        Assert.AreEqual("Pakcage", unused(0), ignoreCase:=True)

    End Sub

    ''' <summary>
    ''' Nested declarations are flattened by ResolveAll before generation, so a variable
    ''' referencing another variable expands to the full cross product
    ''' </summary>
    <TestMethod()> Public Sub NestedVariable_IsResolvedAtParseTime()

        Dim diags As List(Of String) = Nothing
        Dim app = ParseApp(Preamble &
                           "Root=Alpha,Beta" & vbCrLf &
                           "Leaf=<Root>Suffix" & vbCrLf, diags)

        Dim leafValues = app.Variables.Values("Leaf")

        Assert.AreEqual(2, leafValues.Count)
        CollectionAssert.AreEquivalent(New List(Of String) From {"AlphaSuffix", "BetaSuffix"}, leafValues)
        Assert.IsTrue(app.NestedVariableRefs > 0, "Nested reference count should be captured before flattening")

    End Sub

    ' ----- Expansion and domains -----

    ''' <summary>
    ''' The canonical payoff case: a version list fans one RegKey template into one key
    ''' per version, renumbered from 1
    ''' </summary>
    <TestMethod()> Public Sub RegKey_FansOutOverDeclaredVersions()

        Dim section = Build(Preamble &
                            "Version=11.0,16.0" & vbCrLf &
                            "RegKey=HKCU\Software\Contoso\<Version>\File MRU" & vbCrLf)

        Dim regKeys = ValuesOf(section, "RegKey")

        Assert.AreEqual(2, regKeys.Count)
        CollectionAssert.AreEquivalent(
            New List(Of String) From {"HKCU\Software\Contoso\11.0\File MRU",
                                      "HKCU\Software\Contoso\16.0\File MRU"}, regKeys)

    End Sub

    ''' <summary>
    ''' Filesystem-domain keys cannot carry a literal angle bracket, so an undeclared token
    ''' drops the key rather than emitting an invalid path
    ''' </summary>
    <TestMethod()> Public Sub UndeclaredToken_InFileKey_DropsKey()

        Dim section = Build(Preamble & "FileKey=%Package%\Cache\<Missing>|*" & vbCrLf)

        Assert.AreEqual(0, ValuesOf(section, "FileKey").Count)

    End Sub

    ''' <summary>
    ''' Registry-domain keys may contain angle brackets, so an undeclared token is kept
    ''' verbatim with an advisory rather than dropping the key
    ''' </summary>
    <TestMethod()> Public Sub UndeclaredToken_InRegKey_KeepsLiteral()

        Dim section = Build(Preamble & "RegKey=HKCU\Software\Contoso\<Missing>" & vbCrLf)

        Dim regKeys = ValuesOf(section, "RegKey")

        Assert.AreEqual(1, regKeys.Count)
        Assert.AreEqual("HKCU\Software\Contoso\<Missing>", regKeys(0))

    End Sub

    ''' <summary>
    ''' Detect is a registry-domain key and must follow the registry undeclared-token rule
    ''' </summary>
    <TestMethod()> Public Sub Detect_IsRegistryDomain()

        Dim section = Build(Preamble & "Detect=HKCU\Software\Contoso\<Missing>" & vbCrLf)

        Assert.AreEqual(1, ValuesOf(section, "Detect").Count)

    End Sub

    ''' <summary>
    ''' DetectFile is a filesystem-domain key and must follow the filesystem drop rule
    ''' </summary>
    <TestMethod()> Public Sub DetectFile_IsFilesystemDomain()

        Dim section = Build(Preamble & "DetectFile=%LocalAppData%\Contoso\<Missing>" & vbCrLf)

        Assert.AreEqual(0, ValuesOf(section, "DetectFile").Count)

    End Sub

    ''' <summary>
    ''' An ExcludeKey's domain comes from its own flag, not from a fixed per-class rule:
    ''' a REG exclude keeps an undeclared token literal
    ''' </summary>
    <TestMethod()> Public Sub ExcludeKey_RegFlag_KeepsUndeclaredLiteral()

        Dim section = Build(Preamble & "ExcludeKey=REG|HKCU\Software\Contoso\<Missing>" & vbCrLf)

        Assert.AreEqual(1, ValuesOf(section, "ExcludeKey").Count)

    End Sub

    ''' <summary>
    ''' The same construct with a PATH flag is filesystem-domain and drops instead
    ''' </summary>
    <TestMethod()> Public Sub ExcludeKey_PathFlag_DropsUndeclared()

        Dim section = Build(Preamble & "ExcludeKey=PATH|%LocalAppData%\Contoso\<Missing>|*" & vbCrLf)

        Assert.AreEqual(0, ValuesOf(section, "ExcludeKey").Count)

    End Sub

    ' ----- Phase ordering -----

    ''' <summary>
    ''' The three phases must run %Package% → root substitution → variables. A root
    ''' declaring both a package reference and a variable token exercises all three: the
    ''' package resolves first, the root text is substituted into the scaffold template,
    ''' and the token is expanded last against the entry's symbol table
    ''' </summary>
    <TestMethod()> Public Sub PhaseOrder_PackageThenRootThenVariables()

        Dim catalog As New Dictionary(Of String, List(Of String)) From {
            {"Caches", New List(Of String) From {"%WebViewRoot%\Default\Cache|*|RECURSE"}}
        }

        Dim section = Build(Preamble &
                            "Profile=A,B" & vbCrLf &
                            "WebViewRoot=%Package%\LocalState\<Profile>\EBWebView" & vbCrLf &
                            "WebViewScaffolds=Caches" & vbCrLf,
                            webViewCatalog:=catalog)

        Dim fileKeys = ValuesOf(section, "FileKey")

        Assert.AreEqual(2, fileKeys.Count)
        CollectionAssert.AreEquivalent(
            New List(Of String) From {
                "%LocalAppData%\Packages\Contoso.App_8wekyb3d8bbwe\LocalState\A\EBWebView\Default\Cache|*|RECURSE",
                "%LocalAppData%\Packages\Contoso.App_8wekyb3d8bbwe\LocalState\B\EBWebView\Default\Cache|*|RECURSE"},
            fileKeys)

    End Sub

    ''' <summary>
    ''' The QtWebEngine block follows the same phase order as the WebView block
    ''' </summary>
    <TestMethod()> Public Sub PhaseOrder_AppliesToQtWebEngineBlock()

        Dim catalog As New Dictionary(Of String, List(Of String)) From {
            {"Caches", New List(Of String) From {"%QtWebEngineRoot%\Cache|*|RECURSE"}}
        }

        Dim section = Build(Preamble &
                            "QtWebEngineRoot=%Package%\LocalState\QtWebEngine\<Default,OffTheRecord>" & vbCrLf &
                            "QtWebEngineScaffolds=Caches" & vbCrLf,
                            qtCatalog:=catalog)

        Assert.AreEqual(2, ValuesOf(section, "FileKey").Count)

    End Sub

    ' ----- Electron family -----

    ''' <summary>
    ''' The Electron family emits for a hybrid entry — the case that motivated shipping it here at
    ''' all. The root is a plain win32 path with no <c> %Package% </c> reference (Electron lives in
    ''' the desktop half of a hybrid, not inside the MSIX container), and both placeholders bind
    ''' independently.
    ''' </summary>
    <TestMethod()> Public Sub Electron_EmitsForHybridWin32Root()

        Dim catalog As New Dictionary(Of String, List(Of String)) From {
            {"Caches", New List(Of String) From {"%ElectronRoot%\Cache|*|RECURSE"}},
            {"UpdaterCache", New List(Of String) From {"%ElectronUpdaterRoot%\pending|*"}}
        }

        Dim section = Build(Preamble &
                            "ElectronRoot=%LocalAppData%\Amazon Music" & vbCrLf &
                            "ElectronUpdaterRoot=%LocalAppData%\amazon-music-updater" & vbCrLf &
                            "ElectronScaffolds=Caches,UpdaterCache" & vbCrLf,
                            electronCatalog:=catalog)

        CollectionAssert.AreEquivalent(
            New List(Of String) From {
                "%LocalAppData%\Amazon Music\Cache|*|RECURSE",
                "%LocalAppData%\amazon-music-updater\pending|*"},
            ValuesOf(section, "FileKey"))

    End Sub

    ''' <summary>
    ''' An entry declaring only <c> ElectronRoot= </c> drops the updater templates rather than
    ''' emitting a literal <c> %ElectronUpdaterRoot% </c> into a FileKey — the drop rule that lets
    ''' <c> UpdaterCache </c> stay in the default set at no cost
    ''' </summary>
    <TestMethod()> Public Sub Electron_NoUpdaterRoot_DropsUpdaterTemplates()

        Dim catalog As New Dictionary(Of String, List(Of String)) From {
            {"Caches", New List(Of String) From {"%ElectronRoot%\Cache|*|RECURSE"}},
            {"UpdaterCache", New List(Of String) From {"%ElectronUpdaterRoot%\pending|*"}}
        }

        Dim section = Build(Preamble & "ElectronRoot=%AppData%\Signal" & vbCrLf,
                            electronCatalog:=catalog)

        Dim fileKeys = ValuesOf(section, "FileKey")

        Assert.AreEqual(1, fileKeys.Count)
        Assert.AreEqual("%AppData%\Signal\Cache|*|RECURSE", fileKeys(0))

    End Sub

    ''' <summary>
    ''' Declaring only <c> ElectronUpdaterRoot= </c> still opts the entry into the family, since an
    ''' entry may legitimately want the updater cache alone
    ''' </summary>
    <TestMethod()> Public Sub Electron_UpdaterRootAlone_OptsInToFamily()

        Dim catalog As New Dictionary(Of String, List(Of String)) From {
            {"UpdaterCache", New List(Of String) From {"%ElectronUpdaterRoot%\pending|*"}}
        }

        Dim section = Build(Preamble &
                            "ElectronUpdaterRoot=%LocalAppData%\signal-updater" & vbCrLf &
                            "ElectronScaffolds=UpdaterCache" & vbCrLf,
                            electronCatalog:=catalog)

        Assert.AreEqual(1, ValuesOf(section, "FileKey").Count)

    End Sub

    ''' <summary>
    ''' Scaffold FileKey templates from UWP.ini are package-expanded then variable-expanded,
    ''' so a variable declared on the entry drives fan-out of the shared template
    ''' </summary>
    <TestMethod()> Public Sub ScaffoldFileKeys_AreVariableExpanded()

        Dim section = Build(Preamble & "Sub=One,Two" & vbCrLf,
                            scaffoldFileKeys:=New List(Of String) From {"%Package%\<Sub>|*"})

        Assert.AreEqual(2, ValuesOf(section, "FileKey").Count)

    End Sub

    ' ----- Vocabulary parity -----

    ''' <summary>
    ''' RegKeyBase is accepted as a synonym of RegKey, matching the existing
    ''' FileKeyBase/FileKey and ExcludeKey/ExcludeKeyBase synonym pairs
    ''' </summary>
    <TestMethod()> Public Sub RegKeyBase_IsSynonymForRegKey()

        Dim section = Build(Preamble & "RegKeyBase=HKCU\Software\Contoso\MRU" & vbCrLf)

        Dim regKeys = ValuesOf(section, "RegKey")

        Assert.AreEqual(1, regKeys.Count)
        Assert.AreEqual("HKCU\Software\Contoso\MRU", regKeys(0))

    End Sub

    ''' <summary>
    ''' Warning prose is emitted verbatim and never variable-expanded, so angle brackets
    ''' in the text survive intact
    ''' </summary>
    <TestMethod()> Public Sub Warning_IsPassedThroughVerbatim()

        Dim section = Build(Preamble & "Warning=Removes <all> saved data" & vbCrLf)

        Dim warnings = ValuesOf(section, "Warning")

        Assert.AreEqual(1, warnings.Count)
        Assert.AreEqual("Removes <all> saved data", warnings(0))

    End Sub

    ''' <summary>
    ''' DetectOS is valid winapp2 DSL and must be emitted, not dropped
    ''' </summary>
    <TestMethod()> Public Sub DetectOS_IsEmitted()

        Dim section = Build(Preamble & "DetectOS=6.1|10.0" & vbCrLf)

        Dim detectOS = ValuesOf(section, "DetectOS")

        Assert.AreEqual(1, detectOS.Count)
        Assert.AreEqual("6.1|10.0", detectOS(0))

    End Sub

    ' ----- Deprecation rejections survive the open vocabulary -----

    ''' <summary>
    ''' SpecialDetect is deprecated and must be refused rather than silently swallowed as a
    ''' variable declaration by the open vocabulary
    ''' </summary>
    <TestMethod()> Public Sub SpecialDetect_IsRefusedNotTreatedAsVariable()

        Dim diags As List(Of String) = Nothing
        Dim app = ParseApp(Preamble & "SpecialDetect=DET_CHROME" & vbCrLf, diags)

        Assert.IsFalse(app.Variables.IsDeclared("SpecialDetect"),
                       "SpecialDetect must not become a variable declaration")
        Assert.IsTrue(diags.Any(Function(d) d.Contains("SpecialDetect is deprecated")),
                      "Expected a deprecation warning, got: " & String.Join("; ", diags))

    End Sub

    ''' <summary>
    ''' Default is ignored with a warning rather than becoming a variable — UWPBuilder,
    ''' like EntryBuilder, never emits Default
    ''' </summary>
    <TestMethod()> Public Sub Default_IsIgnoredNotTreatedAsVariable()

        Dim diags As List(Of String) = Nothing
        Dim app = ParseApp(Preamble & "Default=False" & vbCrLf, diags)

        Assert.IsFalse(app.Variables.IsDeclared("Default"))
        Assert.IsTrue(diags.Any(Function(d) d.Contains("never emits Default")),
                      "Expected a Default warning, got: " & String.Join("; ", diags))

    End Sub

    ''' <summary>
    ''' Package values are never variable-expanded; a token there is a silent junk path
    ''' waiting to happen, so the parser warns explicitly
    ''' </summary>
    <TestMethod()> Public Sub PackageWithAngleBrackets_Warns()

        Dim diags As List(Of String) = Nothing
        Dim app = ParseApp("[Test App *]" & vbCrLf &
                           "Package=Contoso.App_<Arch>_8wekyb3d8bbwe" & vbCrLf &
                           "LangSecRef=3021" & vbCrLf, diags)

        Assert.IsTrue(diags.Any(Function(d) d.Contains("never variable-expanded")),
                      "Expected a Package token warning, got: " & String.Join("; ", diags))
        Assert.AreEqual("Contoso.App_<Arch>_8wekyb3d8bbwe", app.Packages(0))

    End Sub

End Class
