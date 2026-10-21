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
''' the two-phase expansion order (root placeholder → <c> &lt;var&gt; </c>), the reserved
''' <c> &lt;Package&gt; </c> variable and its indexes, the refusal of the retired
''' <c> %Package% </c> and <c> %WebViewRoot% </c> spellings, root substitution in AppInfo keys,
''' per-key-class domain classification, the open-vocabulary
''' parser and its typo backstops, and the vocabulary shared with EntryBuilder. Also covers the
''' Electron family's two roots and QtWebEngine's cache root. The expansion engine itself is
''' covered by <see cref="VariableExpanderTests"/>; these tests cover the wiring.
''' </summary>
<TestClass()> Public Class UWPBuilderTests

    ''' <summary>Returns the first section parsed from <paramref name="text"/>, and throws if there is none</summary>
    ''' <param name="text">Literal ini text</param>
    Private Shared Function FirstSection(text As String) As winapp2ool.iniSection

        Dim bytes = Encoding.UTF8.GetBytes(text)
        Using ms As New IO.MemoryStream(bytes)
            Using reader As New IO.StreamReader(ms)
                Dim parsed = winapp2ool.iniFile.FromStream(reader, "", "test.ini")
                For Each section In parsed
                    Return section
                Next
            End Using
        End Using

        Throw New InvalidOperationException("No section parsed from test input")

    End Function

    ''' <summary>Returns the first section of <paramref name="text"/> parsed as an AppInfo section</summary>
    '''
    ''' <param name="text">Literal ini text for one AppInfo section</param>
    '''
    ''' <param name="diagnostics">Set to every line the parser wrote to the log</param>
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
    ''' Parses the first section of <paramref name="text"/> as an AppInfo section and generates its
    ''' entry, discarding the menu output and log. We generate even when the parser marks the
    ''' entry skipped.
    ''' </summary>
    '''
    ''' <param name="text">Literal ini text for one AppInfo section</param>
    '''
    ''' <param name="scaffoldFileKeys">
    ''' The shared <c> FileKeyBase= </c> templates from UWP.ini's entry scaffold <br /><br />
    ''' Optional, Default: <c> Nothing </c> <br />
    ''' treated as an empty list
    ''' </param>
    '''
    ''' <param name="scaffoldDetectFiles">
    ''' The shared <c> DetectFileBase= </c> templates from UWP.ini's entry scaffold <br /><br />
    ''' Optional, Default: <c> Nothing </c> <br />
    ''' treated as an empty list
    ''' </param>
    '''
    ''' <param name="webViewCatalog">
    ''' The WebView scaffold catalog <br /><br />
    ''' Optional, Default: <c> Nothing </c> <br />
    ''' leaves the family's catalog empty
    ''' </param>
    '''
    ''' <param name="qtCatalog">
    ''' The QtWebEngine scaffold catalog <br /><br />
    ''' Optional, Default: <c> Nothing </c> <br />
    ''' leaves the family's catalog empty
    ''' </param>
    '''
    ''' <param name="electronCatalog">
    ''' The Electron scaffold catalog <br /><br />
    ''' Optional, Default: <c> Nothing </c> <br />
    ''' leaves the family's catalog empty
    ''' </param>
    '''
    ''' <returns>The generated entry section</returns>
    Private Shared Function Build(text As String,
                         Optional scaffoldFileKeys As List(Of String) = Nothing,
                         Optional scaffoldDetectFiles As List(Of String) = Nothing,
                         Optional webViewCatalog As Dictionary(Of String, List(Of String)) = Nothing,
                         Optional qtCatalog As Dictionary(Of String, List(Of String)) = Nothing,
                         Optional electronCatalog As Dictionary(Of String, List(Of String)) = Nothing) As winapp2ool.iniSection

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
    ''' Parses the first section of <paramref name="text"/> as an AppInfo section and generates its
    ''' entry with no scaffolds, keeping the log
    ''' </summary>
    '''
    ''' <param name="text">Literal ini text for one AppInfo section</param>
    '''
    ''' <param name="log">Set to every line parsing and generation wrote to the log</param>
    Private Shared Function BuildLogged(text As String, ByRef log As List(Of String)) As winapp2ool.iniSection

        Dim menu As New winapp2ool.MenuSection
        Using cap = winapp2ool.gLogCapture()

            Dim app = winapp2ool.UWPBuilder.parseAppInfo(FirstSection(text), menu)
            Dim section = winapp2ool.UWPBuilder.generateUWPEntry(app, New List(Of String), New List(Of String),
                                                                 New winapp2ool.ScaffoldCatalogSet, menu)
            log = cap.Lines.ToList()
            Return section

        End Using

    End Function

    ''' <summary>
    ''' Runs <paramref name="action"/> from a zero exit code and returns the exit code it leaves,
    ''' restoring the previous one afterwards
    ''' </summary>
    '''
    ''' <param name="action">The work to run</param>
    Private Shared Function ExitCodeAfter(action As Action) As Integer

        Dim savedExitCode = Environment.ExitCode

        Try

            Environment.ExitCode = 0
            action()
            Return Environment.ExitCode

        Finally

            Environment.ExitCode = savedExitCode

        End Try

    End Function

    ''' <summary>
    ''' Copies the entries of a plain dictionary into one family's catalog inside a
    ''' <c> ScaffoldCatalogSet </c>, so tests can keep expressing catalogs as literals. The
    ''' template lists are shared, not copied.
    ''' </summary>
    '''
    ''' <param name="source">The catalog to copy from, or <c> Nothing </c> to copy nothing</param>
    '''
    ''' <param name="target">The family's catalog to copy into</param>
    Private Shared Sub CopyCatalog(source As Dictionary(Of String, List(Of String)),
                                   target As Dictionary(Of String, List(Of String)))

        If source Is Nothing Then Return

        For Each pair In source : target(pair.Key) = pair.Value : Next

    End Sub

    ''' <summary>Returns the values of every key of a given type in a section, in key order</summary>
    '''
    ''' <param name="section">The generated section to read</param>
    '''
    ''' <param name="keyType">The KeyType to collect, matched case-insensitively</param>
    Private Shared Function ValuesOf(section As winapp2ool.iniSection, keyType As String) As List(Of String)

        Dim result As New List(Of String)

        For Each key In section.Keys

            If String.Equals(key.KeyType, keyType, StringComparison.InvariantCultureIgnoreCase) Then result.Add(key.Value)

        Next

        Return result

    End Function

    ''' <summary>
    ''' Minimal valid AppInfo preamble: a package and a category, so the entry is not skipped
    ''' </summary>
    Private Const Preamble As String = "[Test App *]" & vbCrLf &
                                       "Package=Hazel.App_abc123" & vbCrLf &
                                       "LangSecRef=3021" & vbCrLf

    ''' <summary>
    ''' AppInfo preamble declaring two packages, so tests can tell the packages apart
    ''' </summary>
    Private Const TwoPackages As String = "[Test App *]" & vbCrLf &
                                          "Package1=Hazel.One_abc123" & vbCrLf &
                                          "Package2=Hazel.Two_abc123" & vbCrLf &
                                          "LangSecRef=3021" & vbCrLf

    ' ----- The Package variable -----

    ''' <summary>
    ''' <c> &lt;Package&gt; </c> fans a key out once per package, in declaration order
    ''' </summary>
    <TestMethod()> Public Sub Package_FansOutOverAllPackages()

        Dim section = Build(TwoPackages & "FileKey=<Package>\LocalState|*.log" & vbCrLf)

        CollectionAssert.AreEqual(
            New List(Of String) From {"%LocalAppData%\Packages\Hazel.One_abc123\LocalState|*.log",
                                      "%LocalAppData%\Packages\Hazel.Two_abc123\LocalState|*.log"},
            ValuesOf(section, "FileKey"))

    End Sub

    ''' <summary>
    ''' <c> &lt;Package[2]&gt; </c> picks the second package declared
    ''' </summary>
    <TestMethod()> Public Sub PackageIndex_PicksThatPackage()

        Dim section = Build(TwoPackages & "FileKey=<Package[2]>\LocalState|*.log" & vbCrLf)

        CollectionAssert.AreEqual(
            New List(Of String) From {"%LocalAppData%\Packages\Hazel.Two_abc123\LocalState|*.log"},
            ValuesOf(section, "FileKey"))

    End Sub

    ''' <summary>
    ''' An index past the package count drops the key with a warning, where the old
    ''' <c> %Package3% </c> passed through as a junk path
    ''' </summary>
    <TestMethod()> Public Sub PackageIndex_PastTheEnd_WarnsAndDrops()

        Dim log As List(Of String) = Nothing
        Dim section = BuildLogged(TwoPackages & "FileKey=<Package[3]>\LocalState|*.log" & vbCrLf, log)

        Assert.AreEqual(0, ValuesOf(section, "FileKey").Count)
        Assert.IsTrue(log.Any(Function(l) l.Contains("<Package[3]> is out of range (2 values)")),
                      "Expected an out-of-range warning, got: " & String.Join("; ", log))

    End Sub

    ''' <summary>
    ''' A <c> Package= </c> value holding an inline list stays one package, so the indexes after
    ''' it don't shift
    ''' </summary>
    <TestMethod()> Public Sub PackageWithInlineList_IsNotFannedOut()

        Dim app = ParseApp("[Test App *]" & vbCrLf &
                           "Package1=Contoso.<a,b>_abc123" & vbCrLf &
                           "Package2=Hazel.Two_abc123" & vbCrLf &
                           "LangSecRef=3021" & vbCrLf, Nothing)

        CollectionAssert.AreEqual(
            New List(Of String) From {"%LocalAppData%\Packages\Contoso.<a,b>_abc123",
                                      "%LocalAppData%\Packages\Hazel.Two_abc123"},
            app.Variables.Values("Package"))

        Dim section = Build("[Test App *]" & vbCrLf &
                            "Package1=Contoso.<a,b>_abc123" & vbCrLf &
                            "Package2=Hazel.Two_abc123" & vbCrLf &
                            "LangSecRef=3021" & vbCrLf &
                            "FileKey=<Package[2]>\LocalState|*.log" & vbCrLf)

        CollectionAssert.AreEqual(
            New List(Of String) From {"%LocalAppData%\Packages\Hazel.Two_abc123\LocalState|*.log"},
            ValuesOf(section, "FileKey"))

    End Sub

    ''' <summary>
    ''' A root declared as <c> &lt;Package&gt;\... </c> yields one scaffold FileKey per package
    ''' </summary>
    <TestMethod()> Public Sub PackageInRoot_YieldsScaffoldFileKeysPerPackage()

        Dim catalog As New Dictionary(Of String, List(Of String)) From {
            {"Caches", New List(Of String) From {"<WebViewRoot>\*\Cache|*|RECURSE"}}
        }

        Dim section = Build(TwoPackages &
                            "WebViewRoot=<Package>\LocalState\EBWebView" & vbCrLf &
                            "WebViewScaffolds=Caches" & vbCrLf,
                            webViewCatalog:=catalog)

        CollectionAssert.AreEqual(
            New List(Of String) From {
                "%LocalAppData%\Packages\Hazel.One_abc123\LocalState\EBWebView\*\Cache|*|RECURSE",
                "%LocalAppData%\Packages\Hazel.Two_abc123\LocalState\EBWebView\*\Cache|*|RECURSE"},
            ValuesOf(section, "FileKey"))

    End Sub

    ''' <summary>
    ''' A variable declaration may reference <c> &lt;Package&gt; </c>, and the reference doesn't
    ''' count toward the nesting statistic
    ''' </summary>
    <TestMethod()> Public Sub PackageInVariable_Resolves()

        Dim app = ParseApp(Preamble & "Cache=<Package>\LocalCache" & vbCrLf, Nothing)

        CollectionAssert.AreEqual(New List(Of String) From {"%LocalAppData%\Packages\Hazel.App_abc123\LocalCache"},
                                  app.Variables.Values("Cache"))
        Assert.AreEqual(0, app.NestedVariableRefs)

    End Sub

    ''' <summary>
    ''' <c> Package </c> is never reported as unreferenced, since an entry with
    ''' <c> SkipUWPFileKeys= </c> may legitimately never use it
    ''' </summary>
    <TestMethod()> Public Sub Package_IsNotReportedUnreferenced()

        Dim app = ParseApp(Preamble & "SkipUWPFileKeys=True" & vbCrLf, Nothing)

        Assert.IsTrue(app.Variables.IsDeclared("Package"))
        Assert.AreEqual(0, app.Variables.UnreferencedNames().Count)

    End Sub

    ' ----- Retired %Package% spelling -----

    ''' <summary>
    ''' Parses <paramref name="text"/> as an AppInfo section, asserting that the run failed and
    ''' that the log names <paramref name="newSpelling"/>
    ''' </summary>
    '''
    ''' <param name="text">Literal ini text for one AppInfo section</param>
    '''
    ''' <param name="newSpelling">The spelling the refusal must suggest</param>
    Private Shared Function ParseRefused(text As String, newSpelling As String) As winapp2ool.UWPBuilder.UWPAppInfo

        Dim app As winapp2ool.UWPBuilder.UWPAppInfo = Nothing
        Dim diags As List(Of String) = Nothing

        Assert.AreEqual(1, ExitCodeAfter(Sub() app = ParseApp(text, diags)))
        Assert.IsTrue(diags.Any(Function(d) d.Contains($"`{newSpelling}`") AndAlso d.Contains("key dropped")),
                      "Expected a refusal naming " & newSpelling & ", got: " & String.Join("; ", diags))

        Return app

    End Function

    ''' <summary>
    ''' A FileKey still spelled <c> %Package% </c> is dropped and fails the run
    ''' </summary>
    <TestMethod()> Public Sub RetiredPackage_InFileKey_IsRefused()

        Dim app = ParseRefused(Preamble & "FileKey=%Package%\LocalState|*.log" & vbCrLf, "<Package>")

        Assert.AreEqual(0, app.AppKeys.Count)

    End Sub

    ''' <summary>
    ''' A <c> %Package </c> missing its closing <c> % </c> is refused too. That's the shape that
    ''' once published <c> %Package\LocalCache|*.log </c>.
    ''' </summary>
    <TestMethod()> Public Sub RetiredPackage_MissingClosingPercent_IsRefused()

        Dim app = ParseRefused(Preamble & "FileKeyBase=%Package\LocalCache|*.log" & vbCrLf, "<Package>")

        Assert.AreEqual(0, app.AppKeys.Count)

    End Sub

    ''' <summary>
    ''' A numbered <c> %Package2% </c> is refused with its indexed spelling
    ''' </summary>
    <TestMethod()> Public Sub RetiredNumberedPackage_IsRefused()

        Dim app = ParseRefused(TwoPackages & "ExcludeKey=FILE|%Package2%\LocalState\|settings.json" & vbCrLf, "<Package[2]>")

        Assert.AreEqual(0, app.ExcludeKeys.Count)

    End Sub

    ''' <summary>
    ''' Root declarations and variable declarations are checked like template keys
    ''' </summary>
    <TestMethod()> Public Sub RetiredPackage_InRootAndVariable_IsRefused()

        Dim app = ParseRefused(Preamble & "WebViewRoot=%Package%\LocalState\EBWebView" & vbCrLf, "<Package>")
        Assert.AreEqual(0, app.WebViewPaths.Count)

        app = ParseRefused(Preamble & "Cache=%Package%\LocalCache" & vbCrLf, "<Package>")
        Assert.IsFalse(app.Variables.IsDeclared("Cache"))

    End Sub

    ''' <summary>
    ''' A <c> UWP.ini </c> scaffold template still spelled <c> %Package% </c> is dropped and fails
    ''' the run
    ''' </summary>
    <TestMethod()> Public Sub RetiredPackage_InScaffold_IsRefused()

        Dim scaffold = FirstSection("[EntryScaffold: UWP App]" & vbCrLf &
                                    "DetectFileBase=<Package>" & vbCrLf &
                                    "FileKeyBase=%Package%\AC|*|RECURSE" & vbCrLf &
                                    "FileKeyBase=<Package>\TempState|*|REMOVESELF" & vbCrLf)
        Dim fileKeys As New List(Of String)
        Dim detectFiles As New List(Of String)

        Dim exitCode = ExitCodeAfter(Sub()
                                         Using winapp2ool.gLogCapture()
                                             winapp2ool.UWPBuilder.parseScaffold(scaffold, fileKeys, detectFiles, New winapp2ool.MenuSection)
                                         End Using
                                     End Sub)

        Assert.AreEqual(1, exitCode)
        CollectionAssert.AreEqual(New List(Of String) From {"<Package>\TempState|*|REMOVESELF"}, fileKeys)
        CollectionAssert.AreEqual(New List(Of String) From {"<Package>"}, detectFiles)

    End Sub

    ''' <summary>
    ''' The retired root spellings are refused in AppInfo keys too: a <c> %WebViewRoot% </c>
    ''' template and a root declaration holding an unclosed <c> %ElectronRoot </c>
    ''' </summary>
    <TestMethod()> Public Sub RetiredRoot_InAppInfoKey_IsRefused()

        Dim app = ParseRefused(Preamble & "FileKeyBase=%WebViewRoot%\EBWebView\*\Cache|*" & vbCrLf, "<WebViewRoot>")
        Assert.AreEqual(0, app.AppKeys.Count)

        app = ParseRefused(Preamble & "ElectronPath=%ElectronRoot\Signal" & vbCrLf, "<ElectronRoot>")
        Assert.AreEqual(0, app.ElectronPaths.Count)

    End Sub

    ''' <summary>
    ''' Package= values and Warning= prose are never checked, and a clean entry leaves the
    ''' exit code alone
    ''' </summary>
    <TestMethod()> Public Sub CleanEntry_LeavesExitCodeAlone()

        Dim exitCode = ExitCodeAfter(Sub() ParseApp(Preamble &
                                                    "Warning=Replaces the old %Package% spelling" & vbCrLf &
                                                    "FileKey=<Package>\LocalState|*.log" & vbCrLf, Nothing))

        Assert.AreEqual(0, exitCode)

    End Sub

    ' ----- Open vocabulary -----

    ''' <summary>
    ''' An unrecognized key is a variable declaration, not an error.
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
    ''' reserved key (e.g. Pakcage=)
    ''' </summary>
    <TestMethod()> Public Sub UnreferencedVariable_IsReportable()

        Dim diags As List(Of String) = Nothing
        Dim app = ParseApp(Preamble & "Pakcage=Hazel.Typo_abc123" & vbCrLf, diags)

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
    ''' per version
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

        Dim section = Build(Preamble & "FileKey=<Package>\Cache\<Missing>|*" & vbCrLf)

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
    ''' The two phases must run root substitution → variables. A root declaring both
    ''' <c> &lt;Package&gt; </c> and another variable exercises both: the root text is
    ''' substituted into the scaffold template, and both tokens are expanded last against the
    ''' entry's symbol table
    ''' </summary>
    <TestMethod()> Public Sub PhaseOrder_RootThenVariables()

        Dim catalog As New Dictionary(Of String, List(Of String)) From {
            {"Caches", New List(Of String) From {"<WebViewRoot>\Default\Cache|*|RECURSE"}}
        }

        Dim section = Build(Preamble &
                            "Profile=A,B" & vbCrLf &
                            "WebViewRoot=<Package>\LocalState\<Profile>\EBWebView" & vbCrLf &
                            "WebViewScaffolds=Caches" & vbCrLf,
                            webViewCatalog:=catalog)

        Dim fileKeys = ValuesOf(section, "FileKey")

        Assert.AreEqual(2, fileKeys.Count)
        CollectionAssert.AreEquivalent(
            New List(Of String) From {
                "%LocalAppData%\Packages\Hazel.App_abc123\LocalState\A\EBWebView\Default\Cache|*|RECURSE",
                "%LocalAppData%\Packages\Hazel.App_abc123\LocalState\B\EBWebView\Default\Cache|*|RECURSE"},
            fileKeys)

    End Sub

    ''' <summary>
    ''' The QtWebEngine block follows the same phase order as the WebView block
    ''' </summary>
    <TestMethod()> Public Sub PhaseOrder_AppliesToQtWebEngineBlock()

        Dim catalog As New Dictionary(Of String, List(Of String)) From {
            {"Caches", New List(Of String) From {"<QtWebEngineRoot>\Cache|*|RECURSE"}}
        }

        Dim section = Build(Preamble &
                            "QtWebEngineRoot=<Package>\LocalState\QtWebEngine\<Default,OffTheRecord>" & vbCrLf &
                            "QtWebEngineScaffolds=Caches" & vbCrLf,
                            qtCatalog:=catalog)

        Assert.AreEqual(2, ValuesOf(section, "FileKey").Count)

    End Sub

    ' ----- Electron family -----

    ''' <summary>
    ''' The Electron family emits for a hybrid entry. The root is a plain win32 path with no
    ''' <c> &lt;Package&gt; </c> reference (Electron lives in the desktop half of a hybrid, not inside
    ''' the MSIX container), and both placeholders bind independently.
    ''' </summary>
    <TestMethod()> Public Sub Electron_EmitsForHybridWin32Root()

        Dim catalog As New Dictionary(Of String, List(Of String)) From {
            {"Caches", New List(Of String) From {"<ElectronRoot>\Cache|*|RECURSE"}},
            {"UpdaterCache", New List(Of String) From {"<ElectronUpdaterRoot>\pending|*"}}
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
    ''' emitting a literal <c> &lt;ElectronUpdaterRoot&gt; </c> into a FileKey. That drop rule is what
    ''' lets <c> UpdaterCache </c> stay in the default set at no cost.
    ''' </summary>
    <TestMethod()> Public Sub Electron_NoUpdaterRoot_DropsUpdaterTemplates()

        Dim catalog As New Dictionary(Of String, List(Of String)) From {
            {"Caches", New List(Of String) From {"<ElectronRoot>\Cache|*|RECURSE"}},
            {"UpdaterCache", New List(Of String) From {"<ElectronUpdaterRoot>\pending|*"}}
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
            {"UpdaterCache", New List(Of String) From {"<ElectronUpdaterRoot>\pending|*"}}
        }

        Dim section = Build(Preamble &
                            "ElectronUpdaterRoot=%LocalAppData%\signal-updater" & vbCrLf &
                            "ElectronScaffolds=UpdaterCache" & vbCrLf,
                            electronCatalog:=catalog)

        Assert.AreEqual(1, ValuesOf(section, "FileKey").Count)

    End Sub

    ' ----- QtWebEngine cache root -----

    ''' <summary>
    ''' <c> QtWebEngineCachePath= </c>, UWPBuilder's alias of <c> QtWebEngineCacheRoot= </c>, binds
    ''' <c> &lt;QtWebEngineCacheRoot&gt; </c> independently of the profile root, and expands
    ''' <c> &lt;Package&gt; </c> like every other root
    ''' </summary>
    <TestMethod()> Public Sub QtWebEngine_CacheRoot_BindsAndSubstitutes()

        Dim catalog As New Dictionary(Of String, List(Of String)) From {
            {"Caches", New List(Of String) From {"<QtWebEngineRoot>\GPUCache|*", "<QtWebEngineCacheRoot>\Cache|*|RECURSE"}}
        }

        Dim section = Build(Preamble &
                            "QtWebEngineRoot=<Package>\LocalCache\QtWebEngine\Default" & vbCrLf &
                            "QtWebEngineCachePath=<Package>\LocalCache\cache\QtWebEngine\Default" & vbCrLf &
                            "QtWebEngineScaffolds=Caches" & vbCrLf,
                            qtCatalog:=catalog)

        CollectionAssert.AreEquivalent(
            New List(Of String) From {
                "%LocalAppData%\Packages\Hazel.App_abc123\LocalCache\QtWebEngine\Default\GPUCache|*",
                "%LocalAppData%\Packages\Hazel.App_abc123\LocalCache\cache\QtWebEngine\Default\Cache|*|RECURSE"},
            ValuesOf(section, "FileKey"))

    End Sub

    ''' <summary>
    ''' An entry declaring only <c> QtWebEngineRoot= </c> drops the cache templates rather than
    ''' emitting a literal <c> &lt;QtWebEngineCacheRoot&gt; </c> into a FileKey
    ''' </summary>
    <TestMethod()> Public Sub QtWebEngine_NoCacheRoot_DropsCacheTemplates()

        Dim catalog As New Dictionary(Of String, List(Of String)) From {
            {"Caches", New List(Of String) From {"<QtWebEngineRoot>\GPUCache|*", "<QtWebEngineCacheRoot>\Cache|*|RECURSE"}}
        }

        Dim section = Build(Preamble & "QtWebEngineRoot=%LocalAppData%\VideoKeeper\QtWebEngine\Default" & vbCrLf,
                            qtCatalog:=catalog)

        Dim fileKeys = ValuesOf(section, "FileKey")

        Assert.AreEqual(1, fileKeys.Count)
        Assert.AreEqual("%LocalAppData%\VideoKeeper\QtWebEngine\Default\GPUCache|*", fileKeys(0))

    End Sub

    ''' <summary>
    ''' Declaring only <c> QtWebEngineCacheRoot= </c> still opts the entry into the family
    ''' </summary>
    <TestMethod()> Public Sub QtWebEngine_CacheRootAlone_OptsInToFamily()

        Dim catalog As New Dictionary(Of String, List(Of String)) From {
            {"Caches", New List(Of String) From {"<QtWebEngineRoot>\GPUCache|*", "<QtWebEngineCacheRoot>\Cache|*|RECURSE"}}
        }

        Dim section = Build(Preamble &
                            "QtWebEngineCacheRoot=%LocalAppData%\VideoKeeper\cache\QtWebEngine\Default" & vbCrLf,
                            qtCatalog:=catalog)

        CollectionAssert.AreEqual(
            New List(Of String) From {"%LocalAppData%\VideoKeeper\cache\QtWebEngine\Default\Cache|*|RECURSE"},
            ValuesOf(section, "FileKey"))

    End Sub

    ' ----- Root substitution in AppInfo keys -----

    ''' <summary>
    ''' An AppInfo <c> FileKeyBase= </c> takes <c> &lt;WebViewRoot&gt; </c> once per declared
    ''' path. A path holding <c> &lt;Package&gt; </c> fans out per package along with the
    ''' template, giving one key per package rather than a product with the other path.
    ''' </summary>
    <TestMethod()> Public Sub AppInfoFileKeyBase_RootIsSubstitutedPerPath()

        Dim section = Build(TwoPackages &
                            "WebViewPath1=<Package>\LocalState" & vbCrLf &
                            "WebViewPath2=%LocalAppData%\Contoso" & vbCrLf &
                            "WebViewScaffolds=" & vbCrLf &
                            "FileKeyBase=<WebViewRoot>\EBWebView\*\Cache|*" & vbCrLf)

        CollectionAssert.AreEqual(
            New List(Of String) From {
                "%LocalAppData%\Packages\Hazel.One_abc123\LocalState\EBWebView\*\Cache|*",
                "%LocalAppData%\Packages\Hazel.Two_abc123\LocalState\EBWebView\*\Cache|*",
                "%LocalAppData%\Contoso\EBWebView\*\Cache|*"},
            ValuesOf(section, "FileKey"))

    End Sub

    ''' <summary>
    ''' A root token for a family the entry declared no path for drops the key silently, as
    ''' <c> FanOutPlaceholder </c> does for a catalog template, and doesn't fail the run
    ''' </summary>
    <TestMethod()> Public Sub AppInfoFileKeyBase_UndeclaredRoot_IsDropped()

        Dim section As winapp2ool.iniSection = Nothing
        Dim exitCode = ExitCodeAfter(Sub() section = Build(Preamble &
                                                           "FileKeyBase=<ElectronRoot>\logs|*" & vbCrLf &
                                                           "FileKey=<Package>\LocalState|*.log" & vbCrLf))

        Assert.AreEqual(0, exitCode)
        CollectionAssert.AreEqual(
            New List(Of String) From {"%LocalAppData%\Packages\Hazel.App_abc123\LocalState|*.log"},
            ValuesOf(section, "FileKey"))

    End Sub

    ''' <summary>
    ''' An AppInfo ExcludeKey takes root substitution, and a declared root is matched in any case
    ''' </summary>
    <TestMethod()> Public Sub AppInfoExcludeKey_RootIsSubstituted()

        Dim section = Build(Preamble &
                            "WebViewPath=<Package>\LocalState" & vbCrLf &
                            "WebViewScaffolds=" & vbCrLf &
                            "ExcludeKey=PATH|<webviewroot>\EBWebView\|*.json" & vbCrLf)

        CollectionAssert.AreEqual(
            New List(Of String) From {"PATH|%LocalAppData%\Packages\Hazel.App_abc123\LocalState\EBWebView\|*.json"},
            ValuesOf(section, "ExcludeKey"))

    End Sub

    ''' <summary>
    ''' An ExcludeKey's domain comes from its root-substituted value. Before substitution this
    ''' template has no flag and would read as a filesystem key, dropping the undeclared token;
    ''' the root supplies <c> REG| </c>, so the token stays literal.
    ''' </summary>
    <TestMethod()> Public Sub AppInfoExcludeKey_DomainComesFromRootExpandedFlag()

        Dim section = Build(Preamble &
                            "WebViewPath=REG|HKCU\Software\Contoso" & vbCrLf &
                            "WebViewScaffolds=" & vbCrLf &
                            "ExcludeKey=<WebViewRoot>\<Missing>" & vbCrLf)

        CollectionAssert.AreEqual(
            New List(Of String) From {"REG|HKCU\Software\Contoso\<Missing>"},
            ValuesOf(section, "ExcludeKey"))

    End Sub

    ''' <summary>
    ''' An AppInfo RegKey takes <c> &lt;WebViewRoot&gt; </c> once per declared path, and one
    ''' naming a root the entry didn't declare is dropped
    ''' </summary>
    <TestMethod()> Public Sub AppInfoRegKey_RootIsSubstitutedOrDropped()

        Dim section = Build(Preamble &
                            "WebViewPath1=HKCU\Software\Contoso" & vbCrLf &
                            "WebViewPath2=HKCU\Software\Fabrikam" & vbCrLf &
                            "WebViewScaffolds=" & vbCrLf &
                            "RegKey=<WebViewRoot>\WebView2" & vbCrLf &
                            "RegKeyBase=<QtWebEngineRoot>\Qt" & vbCrLf)

        CollectionAssert.AreEqual(
            New List(Of String) From {"HKCU\Software\Contoso\WebView2", "HKCU\Software\Fabrikam\WebView2"},
            ValuesOf(section, "RegKey"))

    End Sub

    ''' <summary>
    ''' Scaffold FileKey templates from UWP.ini are variable-expanded along with
    ''' <c> &lt;Package&gt; </c>, so a variable declared on the entry drives fan-out of the
    ''' shared template
    ''' </summary>
    <TestMethod()> Public Sub ScaffoldFileKeys_AreVariableExpanded()

        Dim section = Build(Preamble & "Sub=One,Two" & vbCrLf,
                            scaffoldFileKeys:=New List(Of String) From {"<Package>\<Sub>|*"})

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
    ''' Default is ignored with a warning rather than becoming a variable, since UWPBuilder,
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
