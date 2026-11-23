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
''' Tests for EntryBuilder's scaffold-family wiring of <c> QtWebEngineCacheRoot= </c>: that the key
''' is parsed, bound to <c> &lt;QtWebEngineCacheRoot&gt; </c> in catalog templates and <c> FileKeyBase= </c>
''' values, and opts the entry into the QtWebEngine family on its own. The substitution engine itself
''' is covered by <see cref="ScaffoldCatalogsTests"/>; these tests cover the parser and generator around it.
''' Also covers the refusal of the retired <c> %...Root% </c> and <c> %Package% </c> spellings, the
''' <c> Archive= </c> key and its <c> archived.ini </c> output, and Store apps: <c> Package= </c>,
''' the <c> &lt;Package&gt; </c> variable and its detection, the Package scaffold family, and the
''' refusal of UWPBuilder's leftover keys.
''' </summary>
<TestClass()> Public Class EntryBuilderTests

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

    ''' <summary>
    ''' Parses the first section of <paramref name="text"/> and generates its entry against
    ''' <paramref name="qtCatalog"/>, with empty WebView and Electron catalogs. We generate even
    ''' when the parser marks the entry skipped, so check <paramref name="skipped"/>.
    ''' </summary>
    '''
    ''' <param name="text">Literal source ini text for one EntryBuilder entry</param>
    '''
    ''' <param name="qtCatalog">The QtWebEngine scaffold catalog to generate against</param>
    '''
    ''' <param name="skipped">Set to the parsed entry's <c> ShouldSkip </c></param>
    '''
    ''' <returns>The values of the generated entry's FileKeys, in key order</returns>
    Private Shared Function BuildQt(text As String,
                                    qtCatalog As Dictionary(Of String, List(Of String)),
                                    ByRef skipped As Boolean) As List(Of String)

        Dim menu As New winapp2ool.MenuSection
        Using winapp2ool.gLogCapture()

            Dim spec = winapp2ool.EntryBuilder.parseEntrySpec(FirstSection(text), menu)
            skipped = spec.ShouldSkip

            Dim section = winapp2ool.EntryBuilder.generateEntry(spec,
                New Dictionary(Of String, List(Of String)),
                qtCatalog,
                New Dictionary(Of String, List(Of String)),
                New Dictionary(Of String, List(Of String)),
                New winapp2ool.EntryBuilder.EntryBuilderStats,
                menu)

            Return section.Keys.Where(Function(k) k.KeyType.Equals("FileKey", StringComparison.InvariantCultureIgnoreCase)) _
                               .Select(Function(k) k.Value).ToList()

        End Using

    End Function

    ''' <summary>
    ''' Minimal valid source preamble: a category, so the entry isn't skipped, and a detection, so
    ''' it isn't warned about as always-on
    ''' </summary>
    Private Const Preamble As String = "[Test App *]" & vbCrLf &
                                       "LangSecRef=3021" & vbCrLf &
                                       "DetectFile=%LocalAppData%\VideoKeeper" & vbCrLf

    ''' <summary>
    ''' Returns a catalog whose one scaffold, <c> Caches </c>, holds one profile-relative template
    ''' and one cache-relative template
    ''' </summary>
    Private Shared Function QtCatalog() As Dictionary(Of String, List(Of String))

        Return New Dictionary(Of String, List(Of String)) From {
            {"Caches", New List(Of String) From {"<QtWebEngineRoot>\GPUCache|*", "<QtWebEngineCacheRoot>\Cache|*|RECURSE"}}
        }

    End Function

    ''' <summary>
    ''' <c> QtWebEngineCacheRoot= </c> binds <c> &lt;QtWebEngineCacheRoot&gt; </c> independently of the
    ''' profile root
    ''' </summary>
    <TestMethod()> Public Sub QtWebEngine_CacheRoot_BindsAndSubstitutes()

        Dim skipped As Boolean
        Dim fileKeys = BuildQt(Preamble &
                               "QtWebEngineRoot=%LocalAppData%\VideoKeeper\QtWebEngine\Default" & vbCrLf &
                               "QtWebEngineCacheRoot=%LocalAppData%\VideoKeeper\cache\QtWebEngine\Default" & vbCrLf &
                               "QtWebEngineScaffolds=Caches" & vbCrLf,
                               QtCatalog(), skipped)

        Assert.IsFalse(skipped)
        CollectionAssert.AreEquivalent(
            New List(Of String) From {
                "%LocalAppData%\VideoKeeper\QtWebEngine\Default\GPUCache|*",
                "%LocalAppData%\VideoKeeper\cache\QtWebEngine\Default\Cache|*|RECURSE"},
            fileKeys)

    End Sub

    ''' <summary>
    ''' An entry declaring only <c> QtWebEngineRoot= </c> drops the cache templates rather than
    ''' emitting a literal <c> &lt;QtWebEngineCacheRoot&gt; </c> into a FileKey
    ''' </summary>
    <TestMethod()> Public Sub QtWebEngine_NoCacheRoot_DropsCacheTemplates()

        Dim skipped As Boolean
        Dim fileKeys = BuildQt(Preamble &
                               "QtWebEngineRoot=%LocalAppData%\VideoKeeper\QtWebEngine\Default" & vbCrLf,
                               QtCatalog(), skipped)

        CollectionAssert.AreEqual(
            New List(Of String) From {"%LocalAppData%\VideoKeeper\QtWebEngine\Default\GPUCache|*"},
            fileKeys)

    End Sub

    ''' <summary>
    ''' Declaring only <c> QtWebEngineCacheRoot= </c> still opts the entry into the family, and
    ''' counts as content, so the entry is not skipped as empty
    ''' </summary>
    <TestMethod()> Public Sub QtWebEngine_CacheRootAlone_OptsInToFamily()

        Dim skipped As Boolean
        Dim fileKeys = BuildQt(Preamble &
                               "QtWebEngineCacheRoot=%LocalAppData%\VideoKeeper\cache\QtWebEngine\Default" & vbCrLf,
                               QtCatalog(), skipped)

        Assert.IsFalse(skipped)
        CollectionAssert.AreEqual(
            New List(Of String) From {"%LocalAppData%\VideoKeeper\cache\QtWebEngine\Default\Cache|*|RECURSE"},
            fileKeys)

    End Sub

    ''' <summary>
    ''' <c> FileKeyBase= </c> substitutes <c> &lt;QtWebEngineCacheRoot&gt; </c> like every other root
    ''' placeholder
    ''' </summary>
    <TestMethod()> Public Sub QtWebEngine_CacheRoot_SubstitutesInFileKeyBase()

        Dim skipped As Boolean
        Dim fileKeys = BuildQt(Preamble &
                               "QtWebEngineCacheRoot=%LocalAppData%\VideoKeeper\cache\QtWebEngine\Default" & vbCrLf &
                               "ExcludeQtWebEngineScaffolds=Caches" & vbCrLf &
                               "FileKeyBase=<QtWebEngineCacheRoot>\Cache\Cache_Data|*" & vbCrLf,
                               QtCatalog(), skipped)

        CollectionAssert.AreEqual(
            New List(Of String) From {"%LocalAppData%\VideoKeeper\cache\QtWebEngine\Default\Cache\Cache_Data|*"},
            fileKeys)

    End Sub

    ' ----- Retired %...Root% spelling -----

    ''' <summary>
    ''' Parses the first section of <paramref name="text"/> from a zero exit code, restoring the
    ''' previous exit code afterwards
    ''' </summary>
    '''
    ''' <param name="text">Literal source ini text for one EntryBuilder entry</param>
    '''
    ''' <param name="log">Set to every line the parser wrote to the log</param>
    '''
    ''' <param name="exitCode">Set to the exit code parsing left</param>
    Private Shared Function ParseLogged(text As String,
                                        ByRef log As List(Of String),
                                        ByRef exitCode As Integer) As winapp2ool.EntryBuilder.EntrySpec

        Dim savedExitCode = Environment.ExitCode

        Try

            Environment.ExitCode = 0

            Using cap = winapp2ool.gLogCapture()

                Dim spec = winapp2ool.EntryBuilder.parseEntrySpec(FirstSection(text), New winapp2ool.MenuSection)
                log = cap.Lines.ToList()
                exitCode = Environment.ExitCode
                Return spec

            End Using

        Finally

            Environment.ExitCode = savedExitCode

        End Try

    End Function

    ''' <summary>
    ''' A <c> FileKeyBase= </c> still spelled <c> %QtWebEngineCacheRoot% </c> is dropped with a
    ''' warning naming the new spelling, and the run fails
    ''' </summary>
    <TestMethod()> Public Sub RetiredRootToken_InFileKeyBase_IsRefused()

        Dim log As List(Of String) = Nothing
        Dim exitCode As Integer
        Dim spec = ParseLogged(Preamble &
                               "QtWebEngineCacheRoot=%LocalAppData%\VideoKeeper\cache\QtWebEngine\Default" & vbCrLf &
                               "FileKeyBase=%QtWebEngineCacheRoot%\Cache|*" & vbCrLf, log, exitCode)

        Assert.AreEqual(1, exitCode)
        Assert.AreEqual(0, spec.FileKeyBases.Count)
        Assert.IsTrue(log.Any(Function(l) l.Contains("`%QtWebEngineCacheRoot%` in [Test App *].FileKeyBase is now spelled `<QtWebEngineCacheRoot>`; key dropped")),
                      "Expected a refusal, got: " & String.Join(" | ", log))

    End Sub

    ''' <summary>
    ''' A root token missing its closing <c> % </c> is refused too, in a literal
    ''' <c> FileKey= </c> and in a root declaration
    ''' </summary>
    <TestMethod()> Public Sub RetiredRootToken_UnclosedAndInRootDeclaration_IsRefused()

        Dim log As List(Of String) = Nothing
        Dim exitCode As Integer
        Dim spec = ParseLogged(Preamble &
                               "ElectronRoot=%AppData%\Signal" & vbCrLf &
                               "FileKey=%ElectronRoot\logs|*" & vbCrLf &
                               "QtWebEngineRoot=%ElectronRoot%\QtWebEngine\Default" & vbCrLf, log, exitCode)

        Assert.AreEqual(1, exitCode)
        Assert.AreEqual(0, spec.FileKeys.Count)
        Assert.AreEqual(0, spec.QtWebEngineRoots.Count)
        CollectionAssert.AreEqual({"%AppData%\Signal"}, spec.ElectronRoots)
        Assert.IsTrue(log.Any(Function(l) l.Contains("`%ElectronRoot` in [Test App *].FileKey looks like the retired `%ElectronRoot%` token, which is now spelled `<ElectronRoot>`; key dropped")),
                      "Expected an unclosed-token refusal, got: " & String.Join(" | ", log))

    End Sub

    ''' <summary>
    ''' A key still spelled <c> %Package% </c> or <c> %Package2% </c> is dropped with a warning
    ''' naming the new spelling, and the run fails. <c> Warning= </c> prose is never checked.
    ''' </summary>
    <TestMethod()> Public Sub RetiredPackageToken_IsRefused()

        Dim log As List(Of String) = Nothing
        Dim exitCode As Integer
        Dim spec = ParseLogged(TwoPackages &
                               "FileKey=%Package%\LocalState|*.log" & vbCrLf &
                               "ExcludeKey=FILE|%Package2%\LocalState\|settings.json" & vbCrLf &
                               "Warning=Replaces the old %Package% spelling" & vbCrLf, log, exitCode)

        Assert.AreEqual(1, exitCode)
        Assert.AreEqual(0, spec.FileKeys.Count)
        Assert.AreEqual(0, spec.ExcludeKeys.Count)
        Assert.AreEqual(1, spec.Warnings.Count)
        Assert.IsTrue(log.Any(Function(l) l.Contains("is now spelled `<Package>`; key dropped")), "Got: " & String.Join(" | ", log))
        Assert.IsTrue(log.Any(Function(l) l.Contains("is now spelled `<Package[2]>`; key dropped")), "Got: " & String.Join(" | ", log))

    End Sub

    ''' <summary>
    ''' A <c> Package= </c> value is used literally, so it isn't checked for retired tokens either
    ''' </summary>
    <TestMethod()> Public Sub PackageValue_IsExemptFromTheRetiredTokenCheck()

        Dim log As List(Of String) = Nothing
        Dim exitCode As Integer
        Dim spec = ParseLogged("[Store App *]" & vbCrLf & "Package=%Package%Odd_abc123" & vbCrLf & "LangSecRef=3021" & vbCrLf, log, exitCode)

        Assert.AreEqual(0, exitCode)
        CollectionAssert.AreEqual({"%Package%Odd_abc123"}, spec.Packages)

    End Sub

    ''' <summary>
    ''' A literal <c> FileKey= </c> takes a root placeholder in any case
    ''' </summary>
    <TestMethod()> Public Sub RootPlaceholder_InFileKey_MatchesCaseInsensitively()

        Dim skipped As Boolean
        Dim fileKeys = BuildQt(Preamble &
                               "QtWebEngineRoot=%LocalAppData%\VideoKeeper\QtWebEngine\Default" & vbCrLf &
                               "ExcludeQtWebEngineScaffolds=Caches" & vbCrLf &
                               "FileKey=<qtwebengineroot>\Visited Links|*" & vbCrLf,
                               QtCatalog(), skipped)

        CollectionAssert.AreEqual(
            New List(Of String) From {"%LocalAppData%\VideoKeeper\QtWebEngine\Default\Visited Links|*"},
            fileKeys)

    End Sub

    ' ----- Store apps -----

    ''' <summary> A Store app declaring two packages, so tests can tell the packages apart </summary>
    Private Const TwoPackages As String = "[Store App *]" & vbCrLf &
                                          "Package1=Hazel.One_abc123" & vbCrLf &
                                          "Package2=Hazel.Two_abc123" & vbCrLf &
                                          "LangSecRef=3021" & vbCrLf

    ''' <summary>
    ''' The Package family's catalog as it shipped in UWPBuilder's <c> UWP.ini </c>: one
    ''' <c> Baseline </c> scaffold of four templates
    ''' </summary>
    Private Shared Function PackageCatalog() As Dictionary(Of String, List(Of String))

        Return New Dictionary(Of String, List(Of String)) From {
            {"Baseline", New List(Of String) From {"<Package>\AC|*|RECURSE",
                                                   "<Package>\Settings|*.log*",
                                                   "<Package>\SystemAppData\Helium|*.log*",
                                                   "<Package>\TempState|*|REMOVESELF"}}
        }

    End Function

    ''' <summary>
    ''' Parses the first section of <paramref name="text"/> and generates its entry, with empty
    ''' QtWebEngine and Electron catalogs. We generate even when the parser marks the entry skipped.
    ''' </summary>
    '''
    ''' <param name="text">Literal source ini text for one EntryBuilder entry</param>
    '''
    ''' <param name="log">Set to every line parsing and generation wrote to the log</param>
    '''
    ''' <param name="webViewCatalog">
    ''' The WebView scaffold catalog <br /><br />
    ''' Optional, Default: <c> Nothing </c> <br />
    ''' an empty catalog
    ''' </param>
    '''
    ''' <param name="packageCatalog">
    ''' The Package scaffold catalog <br /><br />
    ''' Optional, Default: <c> Nothing </c> <br />
    ''' an empty catalog
    ''' </param>
    '''
    ''' <returns>The generated entry section</returns>
    Private Shared Function GenerateLogged(text As String,
                                           ByRef log As List(Of String),
                                           Optional webViewCatalog As Dictionary(Of String, List(Of String)) = Nothing,
                                           Optional packageCatalog As Dictionary(Of String, List(Of String)) = Nothing) As winapp2ool.iniSection

        Dim menu As New winapp2ool.MenuSection
        Using cap = winapp2ool.gLogCapture()

            Dim spec = winapp2ool.EntryBuilder.parseEntrySpec(FirstSection(text), menu)
            Dim section = winapp2ool.EntryBuilder.generateEntry(spec,
                If(webViewCatalog, New Dictionary(Of String, List(Of String))),
                New Dictionary(Of String, List(Of String)),
                New Dictionary(Of String, List(Of String)),
                If(packageCatalog, New Dictionary(Of String, List(Of String))),
                New winapp2ool.EntryBuilder.EntryBuilderStats,
                menu)
            log = cap.Lines.ToList()
            Return section

        End Using

    End Function

    ''' <summary>Returns the values of every key of a given type in a section, in key order</summary>
    '''
    ''' <param name="section">The generated section to read</param>
    '''
    ''' <param name="keyType">The KeyType to collect, matched case-insensitively</param>
    Private Shared Function ValuesOf(section As winapp2ool.iniSection, keyType As String) As List(Of String)

        Return section.Keys.Where(Function(k) k.KeyType.Equals(keyType, StringComparison.InvariantCultureIgnoreCase)) _
                           .Select(Function(k) k.Value).ToList()

    End Function

    ''' <summary>
    ''' <c> &lt;Package&gt; </c> fans a key out once per package, in declaration order, and
    ''' <c> &lt;Package[2]&gt; </c> picks the second package declared
    ''' </summary>
    <TestMethod()> Public Sub Package_FansOutAndIndexes()

        Dim log As List(Of String) = Nothing
        Dim section = GenerateLogged(TwoPackages &
                                     "PackageScaffolds=" & vbCrLf &
                                     "FileKey1=<Package>\LocalState|*.log" & vbCrLf &
                                     "FileKey2=<Package[2]>\LocalCache|*.tmp" & vbCrLf, log)

        CollectionAssert.AreEqual(
            New List(Of String) From {"%LocalAppData%\Packages\Hazel.One_abc123\LocalState|*.log",
                                      "%LocalAppData%\Packages\Hazel.Two_abc123\LocalState|*.log",
                                      "%LocalAppData%\Packages\Hazel.Two_abc123\LocalCache|*.tmp"},
            ValuesOf(section, "FileKey"))

    End Sub

    ''' <summary>
    ''' An index past the package count drops the key with a warning
    ''' </summary>
    <TestMethod()> Public Sub PackageIndex_PastTheEnd_WarnsAndDrops()

        Dim log As List(Of String) = Nothing
        Dim section = GenerateLogged(TwoPackages &
                                     "PackageScaffolds=" & vbCrLf &
                                     "FileKey1=<Package>\LocalState|*.log" & vbCrLf &
                                     "FileKey2=<Package[3]>\LocalState|*.log" & vbCrLf, log)

        Assert.AreEqual(2, ValuesOf(section, "FileKey").Count)
        Assert.IsTrue(log.Any(Function(l) l.Contains("<Package[3]> is out of range (2 values)")),
                      "Expected an out-of-range warning, got: " & String.Join(" | ", log))

    End Sub

    ''' <summary>
    ''' A <c> Package= </c> value holding a bracket warns and stays one package, so the indexes
    ''' after it don't shift
    ''' </summary>
    <TestMethod()> Public Sub PackageWithInlineList_WarnsAndIsNotFannedOut()

        Dim text = "[Store App *]" & vbCrLf &
                   "Package1=Contoso.<a,b>_abc123" & vbCrLf &
                   "Package2=Hazel.Two_abc123" & vbCrLf &
                   "LangSecRef=3021" & vbCrLf &
                   "PackageScaffolds=" & vbCrLf &
                   "FileKey=<Package[2]>\LocalState|*.log" & vbCrLf

        Dim log As List(Of String) = Nothing
        Dim exitCode As Integer
        Dim spec = ParseLogged(text, log, exitCode)

        CollectionAssert.AreEqual(
            New List(Of String) From {"%LocalAppData%\Packages\Contoso.<a,b>_abc123",
                                      "%LocalAppData%\Packages\Hazel.Two_abc123"},
            spec.Variables.Values("Package"))
        Assert.IsTrue(log.Any(Function(l) l.Contains("Package is never variable-expanded and the value is used literally")),
                      "Expected a Package bracket warning, got: " & String.Join(" | ", log))

        Dim section = GenerateLogged(text, log)

        CollectionAssert.AreEqual(
            New List(Of String) From {"%LocalAppData%\Packages\Hazel.Two_abc123\LocalState|*.log"},
            ValuesOf(section, "FileKey"))

    End Sub

    ''' <summary>
    ''' A root declared as <c> &lt;Package&gt;\... </c> resolves in phase 2, yielding one scaffold
    ''' FileKey per package
    ''' </summary>
    <TestMethod()> Public Sub PackageInRoot_YieldsScaffoldFileKeysPerPackage()

        Dim catalog As New Dictionary(Of String, List(Of String)) From {
            {"Caches", New List(Of String) From {"<WebViewRoot>\*\Cache|*|RECURSE"}}
        }

        Dim log As List(Of String) = Nothing
        Dim section = GenerateLogged(TwoPackages &
                                     "PackageScaffolds=" & vbCrLf &
                                     "WebViewRoot=<Package>\LocalState\EBWebView" & vbCrLf &
                                     "WebViewScaffolds=Caches" & vbCrLf, log, webViewCatalog:=catalog)

        CollectionAssert.AreEqual(
            New List(Of String) From {
                "%LocalAppData%\Packages\Hazel.One_abc123\LocalState\EBWebView\*\Cache|*|RECURSE",
                "%LocalAppData%\Packages\Hazel.Two_abc123\LocalState\EBWebView\*\Cache|*|RECURSE"},
            ValuesOf(section, "FileKey"))

    End Sub

    ''' <summary>
    ''' A variable declaration may reference <c> &lt;Package&gt; </c>, the reference doesn't count
    ''' toward the nesting statistic, and <c> Package </c> is never reported as unreferenced
    ''' </summary>
    <TestMethod()> Public Sub PackageInVariable_ResolvesAndIsNeverUnreferenced()

        Dim log As List(Of String) = Nothing
        Dim exitCode As Integer
        Dim spec = ParseLogged(TwoPackages & "PackageScaffolds=" & vbCrLf & "Cache=<Package[1]>\LocalCache" & vbCrLf & "FileKey=<Cache>|*" & vbCrLf, log, exitCode)

        CollectionAssert.AreEqual(New List(Of String) From {"%LocalAppData%\Packages\Hazel.One_abc123\LocalCache"},
                                  spec.Variables.Values("Cache"))
        Assert.AreEqual(0, spec.NestedVariableRefs)

        spec = ParseLogged(TwoPackages & "PackageScaffolds=" & vbCrLf, log, exitCode)
        Assert.IsTrue(spec.Variables.IsDeclared("Package"))
        Assert.AreEqual(0, spec.Variables.UnreferencedNames().Count)

    End Sub

    ''' <summary>
    ''' Each package adds a <c> DetectFile </c>, ahead of the one a <c> Root </c> variable infers
    ''' and the entry's own
    ''' </summary>
    <TestMethod()> Public Sub Package_AddsDetectFilesFirst()

        Dim log As List(Of String) = Nothing
        Dim section = GenerateLogged(TwoPackages &
                                     "Root=%LocalAppData%\StoreApp" & vbCrLf &
                                     "DetectFile=%ProgramFiles%\StoreApp" & vbCrLf &
                                     "PackageScaffolds=" & vbCrLf &
                                     "FileKey=<Root>|*.log" & vbCrLf, log)

        CollectionAssert.AreEqual(
            New List(Of String) From {"%LocalAppData%\Packages\Hazel.One_abc123",
                                      "%LocalAppData%\Packages\Hazel.Two_abc123",
                                      "%LocalAppData%\StoreApp",
                                      "%ProgramFiles%\StoreApp"},
            ValuesOf(section, "DetectFile"))
        Assert.IsFalse(log.Any(Function(l) l.Contains("declares no detection")), "Got: " & String.Join(" | ", log))

    End Sub

    ''' <summary>
    ''' An entry with only <c> Package= </c> and a category is built, with its package detection and
    ''' the default <c> Baseline </c> scaffold, rather than skipped as having nothing to emit
    ''' </summary>
    <TestMethod()> Public Sub PackageOnlyEntry_IsBuilt()

        Dim text = "[BreeZip *]" & vbCrLf & "Package=Hazel.BreeZip_abc123" & vbCrLf & "LangSecRef=3031" & vbCrLf

        Dim log As List(Of String) = Nothing
        Dim exitCode As Integer
        Dim spec = ParseLogged(text, log, exitCode)

        Assert.IsFalse(spec.ShouldSkip, "Got: " & String.Join(" | ", log))
        Assert.AreEqual(0, exitCode)

        Dim section = GenerateLogged(text, log, packageCatalog:=PackageCatalog())

        CollectionAssert.AreEqual({"%LocalAppData%\Packages\Hazel.BreeZip_abc123"}, ValuesOf(section, "DetectFile"))
        Assert.AreEqual(4, ValuesOf(section, "FileKey").Count)

    End Sub

    ''' <summary>
    ''' The Package catalog's <c> Baseline </c> yields, for a two-package entry, exactly the keys
    ''' UWPBuilder produced from <c> UWP.ini </c>'s templates, which it fanned out through the
    ''' <c> Package </c> variable. The second half runs the same templates through that variable
    ''' path here, as literal <c> FileKey= </c> values, and compares.
    ''' </summary>
    <TestMethod()> Public Sub PackageBaseline_MatchesUWPBuilderOutput()

        Dim log As List(Of String) = Nothing
        Dim fromCatalog = ValuesOf(GenerateLogged(TwoPackages, log, packageCatalog:=PackageCatalog()), "FileKey")

        CollectionAssert.AreEqual(
            New List(Of String) From {
                "%LocalAppData%\Packages\Hazel.One_abc123\AC|*|RECURSE",
                "%LocalAppData%\Packages\Hazel.Two_abc123\AC|*|RECURSE",
                "%LocalAppData%\Packages\Hazel.One_abc123\Settings|*.log*",
                "%LocalAppData%\Packages\Hazel.Two_abc123\Settings|*.log*",
                "%LocalAppData%\Packages\Hazel.One_abc123\SystemAppData\Helium|*.log*",
                "%LocalAppData%\Packages\Hazel.Two_abc123\SystemAppData\Helium|*.log*",
                "%LocalAppData%\Packages\Hazel.One_abc123\TempState|*|REMOVESELF",
                "%LocalAppData%\Packages\Hazel.Two_abc123\TempState|*|REMOVESELF"},
            fromCatalog)

        Dim templates = PackageCatalog()("Baseline")
        Dim asKeys = String.Concat(templates.Select(Function(t, i) $"FileKey{i + 1}={t}" & vbCrLf))
        Dim fromVariable = ValuesOf(GenerateLogged(TwoPackages & "PackageScaffolds=" & vbCrLf & asKeys, log), "FileKey")

        CollectionAssert.AreEqual(fromVariable, fromCatalog)

    End Sub

    ''' <summary>
    ''' An empty <c> PackageScaffolds= </c> selects no scaffolds, rather than the default set, and
    ''' leaves the package detection in place
    ''' </summary>
    <TestMethod()> Public Sub EmptyPackageScaffolds_SelectsNone()

        Dim log As List(Of String) = Nothing
        Dim section = GenerateLogged(TwoPackages &
                                     "PackageScaffolds=" & vbCrLf &
                                     "FileKey=<Package[1]>\LocalState|*.log" & vbCrLf, log, packageCatalog:=PackageCatalog())

        CollectionAssert.AreEqual(New List(Of String) From {"%LocalAppData%\Packages\Hazel.One_abc123\LocalState|*.log"},
                                  ValuesOf(section, "FileKey"))
        Assert.AreEqual(2, ValuesOf(section, "DetectFile").Count)

    End Sub

    ''' <summary>
    ''' <c> &lt;Package&gt; </c> expands in RegKeys and ExcludeKeys like any variable
    ''' </summary>
    <TestMethod()> Public Sub Package_ExpandsInRegKeysAndExcludeKeys()

        Dim log As List(Of String) = Nothing
        Dim section = GenerateLogged(TwoPackages &
                                     "PackageScaffolds=" & vbCrLf &
                                     "FileKey=<Package>\LocalState|*.log" & vbCrLf &
                                     "RegKey=HKCU\Software\Hazel|<Package[2]>" & vbCrLf &
                                     "ExcludeKey=FILE|<Package>\LocalState\|settings.json" & vbCrLf, log)

        CollectionAssert.AreEqual(New List(Of String) From {"HKCU\Software\Hazel|%LocalAppData%\Packages\Hazel.Two_abc123"},
                                  ValuesOf(section, "RegKey"))
        CollectionAssert.AreEqual(
            New List(Of String) From {"FILE|%LocalAppData%\Packages\Hazel.One_abc123\LocalState\|settings.json",
                                      "FILE|%LocalAppData%\Packages\Hazel.Two_abc123\LocalState\|settings.json"},
            ValuesOf(section, "ExcludeKey"))

    End Sub

    ''' <summary>
    ''' UWPBuilder's <c> SkipUWPFileKeys= </c> and its five <c> ...Path= </c> root spellings are
    ''' dropped with a warning naming the spelling to use, with a value or without, and fail the
    ''' run rather than becoming unused variables
    ''' </summary>
    <TestMethod()> Public Sub UWPBuilderKeys_AreRefused()

        Dim replacements As New Dictionary(Of String, String) From {
            {"SkipUWPFileKeys", "an empty PackageScaffolds="},
            {"WebViewPath", "WebViewRoot="},
            {"QtWebEnginePath", "QtWebEngineRoot="},
            {"QtWebEngineCachePath", "QtWebEngineCacheRoot="},
            {"ElectronPath", "ElectronRoot="},
            {"ElectronUpdaterPath", "ElectronUpdaterRoot="}
        }

        For Each pair In replacements

            For Each value In {"<Package>\LocalState", ""}

                Dim log As List(Of String) = Nothing
                Dim exitCode As Integer
                Dim spec = ParseLogged(TwoPackages & $"{pair.Key}={value}" & vbCrLf, log, exitCode)

                Assert.AreEqual(1, exitCode, $"{pair.Key}={value}")
                Assert.IsFalse(spec.Variables.IsDeclared(pair.Key), pair.Key)
                Assert.AreEqual(0, spec.WebViewRoots.Count + spec.QtWebEngineRoots.Count + spec.QtWebEngineCacheRoots.Count +
                                   spec.ElectronRoots.Count + spec.ElectronUpdaterRoots.Count, pair.Key)
                Assert.IsTrue(log.Any(Function(l) l.Contains($"{pair.Key}= in [Store App *] is UWPBuilder's spelling; key dropped. Use {pair.Value} instead")),
                              "Got: " & String.Join(" | ", log))

            Next

        Next

    End Sub

    ' ----- Keys with no value -----

    ''' <summary>
    ''' A bare <c> Skip= </c> skips the entry. A key with no value has the KeyType
    ''' <c> DeleteMe </c>, so the parser goes by its name.
    ''' </summary>
    <TestMethod()> Public Sub BareSkip_SkipsTheEntry()

        Dim log As List(Of String) = Nothing
        Dim exitCode As Integer
        Dim spec = ParseLogged(Preamble & "FileKey=%LocalAppData%\VideoKeeper|*.log" & vbCrLf & "Skip=" & vbCrLf, log, exitCode)

        Assert.IsTrue(spec.ShouldSkip)
        Assert.AreEqual(0, exitCode)
        Assert.IsFalse(log.Any(Function(l) l.Contains("DeleteMe")), "Got: " & String.Join(" | ", log))

    End Sub

    ''' <summary>
    ''' An empty <c> WebViewScaffolds= </c> selects no scaffolds, rather than the default set
    ''' </summary>
    <TestMethod()> Public Sub EmptyScaffoldSelection_SelectsNone()

        Dim webView As New Dictionary(Of String, List(Of String)) From {
            {"Caches", New List(Of String) From {"<WebViewRoot>\Cache|*"}}
        }

        Dim menu As New winapp2ool.MenuSection
        Using winapp2ool.gLogCapture()

            Dim spec = winapp2ool.EntryBuilder.parseEntrySpec(FirstSection(Preamble &
                "WebViewRoot=%LocalAppData%\VideoKeeper\EBWebView\Default" & vbCrLf &
                "WebViewScaffolds=" & vbCrLf &
                "FileKey=%LocalAppData%\VideoKeeper|*.log" & vbCrLf), menu)

            Dim section = winapp2ool.EntryBuilder.generateEntry(spec, webView,
                New Dictionary(Of String, List(Of String)),
                New Dictionary(Of String, List(Of String)),
                New Dictionary(Of String, List(Of String)),
                New winapp2ool.EntryBuilder.EntryBuilderStats,
                menu)

            Dim fileKeys = section.Keys.Where(Function(k) k.KeyType.Equals("FileKey", StringComparison.InvariantCultureIgnoreCase)) _
                                       .Select(Function(k) k.Value).ToList()

            CollectionAssert.AreEqual(New List(Of String) From {"%LocalAppData%\VideoKeeper|*.log"}, fileKeys)

        End Using

    End Sub

    ''' <summary>
    ''' Any other key with no value is dropped with a warning, without failing the run, rather
    ''' than giving an empty key or root
    ''' </summary>
    <TestMethod()> Public Sub EmptyValueKey_IsDroppedWithWarning()

        Dim log As List(Of String) = Nothing
        Dim exitCode As Integer
        Dim spec = ParseLogged(Preamble &
                               "FileKey=%LocalAppData%\VideoKeeper|*.log" & vbCrLf &
                               "FileKey2=" & vbCrLf &
                               "WebViewRoot=" & vbCrLf &
                               "Version=" & vbCrLf, log, exitCode)

        Assert.AreEqual(0, exitCode)
        Assert.AreEqual(1, spec.FileKeys.Count)
        Assert.AreEqual(0, spec.WebViewRoots.Count)
        Assert.IsFalse(spec.Variables.IsDeclared("Version"))
        Assert.IsTrue(log.Any(Function(l) l.Contains("[Test App *].FileKey2 has no value; key dropped")), "Got: " & String.Join(" | ", log))
        Assert.IsTrue(log.Any(Function(l) l.Contains("[Test App *].WebViewRoot has no value; key dropped")), "Got: " & String.Join(" | ", log))

    End Sub

    ' ----- Whole runs -----

    ''' <summary>
    ''' A run over a source directory with no sections writes nothing and fails the run
    ''' </summary>
    <TestMethod()> Public Sub NoSources_FailsTheRun()

        Dim dir = IO.Path.Combine(IO.Path.GetTempPath(), $"winapp2ool-entrybuilder-{Guid.NewGuid():N}")
        Dim savedSource = winapp2ool.EntryBuilderFile1
        Dim savedSuppress = winapp2ool.SuppressOutput
        Dim savedExitCode = Environment.ExitCode

        Try

            IO.Directory.CreateDirectory(dir)
            winapp2ool.EntryBuilderFile1 = New winapp2ool.iniFileChooser(dir, "", "")
            winapp2ool.SuppressOutput = True
            Environment.ExitCode = 0

            winapp2ool.EntryBuilder.initEntryBuilder()

            Assert.AreEqual(1, Environment.ExitCode)

        Finally

            winapp2ool.EntryBuilderFile1 = savedSource
            winapp2ool.SuppressOutput = savedSuppress
            Environment.ExitCode = savedExitCode
            IO.Directory.Delete(dir, recursive:=True)

        End Try

    End Sub

    ' ----- Archive= -----

    ''' <summary> A whole source entry with no Archive= key, live in every run below </summary>
    Private Const LiveEntry As String = "[Alpha App *]" & vbCrLf &
                                        "LangSecRef=3021" & vbCrLf &
                                        "DetectFile=%LocalAppData%\Alpha" & vbCrLf &
                                        "FileKey=%LocalAppData%\Alpha|*.log" & vbCrLf

    ''' <summary> Returns a whole source entry named <paramref name="name"/> carrying <paramref name="archiveLine"/> </summary>
    ''' <param name="name">The entry's name, without the trailing <c> * </c></param>
    ''' <param name="archiveLine">The entry's <c> Archive </c> line, without its line break</param>
    Private Shared Function ArchivedEntry(name As String, archiveLine As String) As String

        Return $"[{name} *]" & vbCrLf &
               "LangSecRef=3021" & vbCrLf &
               $"DetectFile=%LocalAppData%\{name}" & vbCrLf &
               $"FileKey=%LocalAppData%\{name}|*.log" & vbCrLf &
               archiveLine & vbCrLf

    End Function

    ''' <summary>
    ''' Runs EntryBuilder's command line over a scratch source directory holding
    ''' <paramref name="sources"/>, writing a single output file, and returns the exit code it
    ''' leaves. Every setting it touches is restored.
    ''' </summary>
    '''
    ''' <param name="sources">Each source file's name and text</param>
    '''
    ''' <param name="liveText">Set to the text of <c> entrybuilder.ini </c>, or <c> Nothing </c></param>
    '''
    ''' <param name="archiveText">Set to the text of <c> archived.ini </c> in the archive folder, or <c> Nothing </c></param>
    '''
    ''' <param name="outFiles">Set to the name and text of every file the run wrote in the output folder</param>
    '''
    ''' <param name="archiveArgs">
    ''' The archive folder arguments. <c> {archive} </c> in them stands for the scratch archive
    ''' folder.
    ''' </param>
    Private Shared Function RunArchiveBuild(sources As Dictionary(Of String, String),
                                            ByRef liveText As String,
                                            ByRef archiveText As String,
                                            ByRef outFiles As Dictionary(Of String, String),
                                            ParamArray archiveArgs As String()) As Integer

        Dim dir = IO.Path.Combine(IO.Path.GetTempPath(), $"winapp2ool-ebarchive-{Guid.NewGuid():N}")
        Dim sourceDir = IO.Path.Combine(dir, "Sources")
        Dim outDir = IO.Path.Combine(dir, "Out")
        Dim archiveDir = IO.Path.Combine(dir, "Archived")
        Dim saved = {winapp2ool.EntryBuilderFile1, winapp2ool.EntryBuilderFile2, winapp2ool.EntryBuilderFile3, winapp2ool.EntryBuilderArchiveFile}
        Dim savedSplit = winapp2ool.EntryBuilderSplitOutput
        Dim savedSuppress = winapp2ool.SuppressOutput
        Dim savedExitCode = Environment.ExitCode

        Try

            IO.Directory.CreateDirectory(sourceDir)
            IO.Directory.CreateDirectory(IO.Path.Combine(dir, "Scaffolds"))
            For Each source In sources
                IO.File.WriteAllText(IO.Path.Combine(sourceDir, source.Key), source.Value)
            Next

            winapp2ool.SuppressOutput = True
            Environment.ExitCode = 0

            Dim args = {"-1d", sourceDir, "-2d", outDir, "-3d", IO.Path.Combine(dir, "Scaffolds")}.Concat(archiveArgs.Select(Function(a) a.Replace("{archive}", archiveDir))).ToArray()
            setCmdLineArgs(AddressOf winapp2ool.EntryBuilder.handleCmdLine, args)

            Dim livePath = IO.Path.Combine(outDir, "entrybuilder.ini")
            Dim archivePath = IO.Path.Combine(archiveDir, "archived.ini")
            liveText = If(IO.File.Exists(livePath), IO.File.ReadAllText(livePath), Nothing)
            archiveText = If(IO.File.Exists(archivePath), IO.File.ReadAllText(archivePath), Nothing)
            outFiles = If(IO.Directory.Exists(outDir), IO.Directory.GetFiles(outDir).ToDictionary(Function(f) IO.Path.GetFileName(f), Function(f) IO.File.ReadAllText(f)), New Dictionary(Of String, String))

            Return Environment.ExitCode

        Finally

            winapp2ool.EntryBuilderFile1 = saved(0)
            winapp2ool.EntryBuilderFile2 = saved(1)
            winapp2ool.EntryBuilderFile3 = saved(2)
            winapp2ool.EntryBuilderArchiveFile = saved(3)
            winapp2ool.EntryBuilderSplitOutput = savedSplit
            winapp2ool.SuppressOutput = savedSuppress
            Environment.ExitCode = savedExitCode
            IO.Directory.Delete(dir, recursive:=True)

        End Try

    End Function

    ''' <summary>
    ''' An entry with <c> Archive= </c> goes to <c> archived.ini </c> in the folder slot 4 names,
    ''' built like a live entry and without the Archive key, and only there
    ''' </summary>
    <TestMethod()> Public Sub Archive_RoutesToTheArchiveFile()

        Dim live As String = Nothing
        Dim archive As String = Nothing
        Dim outFiles As Dictionary(Of String, String) = Nothing
        Dim sources As New Dictionary(Of String, String) From {{"A.ini", LiveEntry & ArchivedEntry("Retired App", "Archive=Discontinued (2012)")}}

        Assert.AreEqual(0, RunArchiveBuild(sources, live, archive, outFiles, "-4d", "{archive}"))

        StringAssert.Contains(live, "[Alpha App *]")
        Assert.IsFalse(live.Contains("Retired App"), live)
        Assert.IsTrue(archive.StartsWith("; # of entries: 1" & vbCrLf & "; This file is generated by winapp2ool's EntryBuilder" & vbCrLf, StringComparison.Ordinal), archive)
        StringAssert.Contains(archive, "[Retired App *]" & vbCrLf & "LangSecRef=3021" & vbCrLf & "DetectFile=%LocalAppData%\Retired App" & vbCrLf & "FileKey1=%LocalAppData%\Retired App|*.log")
        Assert.IsFalse(archive.Contains("Alpha App"), archive)
        Assert.IsFalse(archive.Contains(vbLf & "Archive"), archive)
        Assert.IsFalse(archive.Contains("Discontinued"), archive)

    End Sub

    ''' <summary> The <c> archive </c> alias sets slot 4 too, with a leading <c> \ </c> read against the working directory </summary>
    <TestMethod()> Public Sub Archive_AliasResolvesAgainstTheWorkingDirectory()

        Dim dir = IO.Path.Combine(IO.Path.GetTempPath(), $"winapp2ool-ebcwd-{Guid.NewGuid():N}")
        Dim savedCwd = Environment.CurrentDirectory
        Dim live As String = Nothing
        Dim archive As String = Nothing
        Dim outFiles As Dictionary(Of String, String) = Nothing
        Dim sources As New Dictionary(Of String, String) From {{"A.ini", LiveEntry & ArchivedEntry("Retired App", "Archive=Gone")}}

        Try

            IO.Directory.CreateDirectory(dir)
            Environment.CurrentDirectory = dir

            Assert.AreEqual(0, RunArchiveBuild(sources, live, archive, outFiles, "-archived", "\Arc\EntryBuilder"))
            StringAssert.Contains(IO.File.ReadAllText(IO.Path.Combine(dir, "Arc", "EntryBuilder", "archived.ini")), "[Retired App *]")

        Finally

            Environment.CurrentDirectory = savedCwd
            IO.Directory.Delete(dir, recursive:=True)

        End Try

    End Sub

    ''' <summary> A bare <c> Archive= </c> archives the entry without a warning </summary>
    <TestMethod()> Public Sub BareArchive_ArchivesTheEntry()

        Dim log As List(Of String) = Nothing
        Dim exitCode As Integer
        Dim spec = ParseLogged(ArchivedEntry("Retired App", "Archive="), log, exitCode)

        Assert.IsTrue(spec.ShouldArchive)
        Assert.IsFalse(spec.ShouldSkip)
        Assert.AreEqual(0, exitCode)
        Assert.IsFalse(log.Any(Function(l) l.Contains("no value")), "Got: " & String.Join(" | ", log))

    End Sub

    ''' <summary> <c> Archive=False </c>, in any case, builds nothing and fails the run, telling the author to delete the line </summary>
    <TestMethod()> Public Sub ArchiveFalse_IsRefused()

        For Each value In {"False", "false", " FALSE "}

            Dim log As List(Of String) = Nothing
            Dim exitCode As Integer
            Dim spec = ParseLogged(ArchivedEntry("Retired App", "Archive=" & value), log, exitCode)

            Assert.IsTrue(spec.ShouldSkip, value)
            Assert.IsFalse(spec.ShouldArchive, value)
            Assert.AreEqual(1, exitCode, value)
            Assert.IsTrue(log.Any(Function(l) l.Contains("[Retired App *] declares Archive=False; delete the Archive line")), "Got: " & String.Join(" | ", log))

        Next

    End Sub

    ''' <summary> <c> Skip= </c> and <c> Archive= </c> on one entry build neither and fail the run </summary>
    <TestMethod()> Public Sub SkipAndArchive_AreRefused()

        Dim log As List(Of String) = Nothing
        Dim exitCode As Integer
        Dim spec = ParseLogged(ArchivedEntry("Retired App", "Archive=Gone") & "Skip" & vbCrLf, log, exitCode)

        Assert.IsTrue(spec.ShouldSkip)
        Assert.IsFalse(spec.ShouldArchive)
        Assert.AreEqual(1, exitCode)
        Assert.IsTrue(log.Any(Function(l) l.Contains("[Retired App *] declares both Skip and Archive")), "Got: " & String.Join(" | ", log))

    End Sub

    ''' <summary> A structural skip wins over <c> Archive= </c>, with its own warning and without failing the run </summary>
    <TestMethod()> Public Sub StructuralSkip_BeatsArchive()

        Dim log As List(Of String) = Nothing
        Dim exitCode As Integer
        Dim spec = ParseLogged("[Retired App *]" & vbCrLf & "FileKey=%LocalAppData%\Retired|*.log" & vbCrLf & "Archive=Gone" & vbCrLf, log, exitCode)

        Assert.IsTrue(spec.ShouldSkip)
        Assert.IsFalse(spec.ShouldArchive)
        Assert.AreEqual(0, exitCode)
        Assert.IsTrue(log.Any(Function(l) l.Contains("No Section or LangSecRef in [Retired App *], skipping")), "Got: " & String.Join(" | ", log))

    End Sub

    ''' <summary>
    ''' A live and an archived section of one name, in two source files, meet in the source
    ''' merge: the first is kept live, the repeat is dropped and the run fails
    ''' </summary>
    <TestMethod()> Public Sub LiveAndArchivedOfOneName_FailTheRun()

        Dim live As String = Nothing
        Dim archive As String = Nothing
        Dim outFiles As Dictionary(Of String, String) = Nothing
        Dim sources As New Dictionary(Of String, String) From {
            {"A.ini", LiveEntry},
            {"B.ini", ArchivedEntry("Alpha App", "Archive=Gone")}
        }

        Assert.AreEqual(1, RunArchiveBuild(sources, live, archive, outFiles, "-4d", "{archive}"))
        StringAssert.Contains(live, "[Alpha App *]")
        Assert.IsFalse(archive.Contains("[Alpha App *]"), archive)

    End Sub

    ''' <summary>
    ''' With no archive folder, archived entries reach neither output and the run fails, while
    ''' the live output is still written
    ''' </summary>
    <TestMethod()> Public Sub NoArchiveFolder_DropsArchivedEntriesAndFails()

        Dim live As String = Nothing
        Dim archive As String = Nothing
        Dim outFiles As Dictionary(Of String, String) = Nothing
        Dim sources As New Dictionary(Of String, String) From {{"A.ini", LiveEntry & ArchivedEntry("Retired App", "Archive=Gone")}}

        Assert.AreEqual(1, RunArchiveBuild(sources, live, archive, outFiles))
        StringAssert.Contains(live, "[Alpha App *]")
        Assert.IsFalse(live.Contains("Retired App"), live)
        Assert.IsNull(archive)
        CollectionAssert.AreEqual({"entrybuilder.ini"}, outFiles.Keys.ToArray())

    End Sub

    ''' <summary> Given an archive folder, a run with nothing archived still writes a header-only <c> archived.ini </c> </summary>
    <TestMethod()> Public Sub NothingArchived_WritesHeaderOnlyArchive()

        Dim live As String = Nothing
        Dim archive As String = Nothing
        Dim outFiles As Dictionary(Of String, String) = Nothing
        Dim sources As New Dictionary(Of String, String) From {{"A.ini", LiveEntry}}

        Assert.AreEqual(0, RunArchiveBuild(sources, live, archive, outFiles, "-4d", "{archive}"))
        Assert.IsTrue(archive.StartsWith("; # of entries: 0" & vbCrLf, StringComparison.Ordinal), archive)
        Assert.IsFalse(archive.Contains("["), archive)

    End Sub

    ''' <summary> A Store app with <c> Archive= </c> goes to <c> archived.ini </c>, package detection and all </summary>
    <TestMethod()> Public Sub Archive_PackageEntryRoutesToTheArchiveFile()

        Dim live As String = Nothing
        Dim archive As String = Nothing
        Dim outFiles As Dictionary(Of String, String) = Nothing
        Dim sources As New Dictionary(Of String, String) From {
            {"A.ini", LiveEntry & "[Retired Store *]" & vbCrLf & "Package=Hazel.Retired_abc" & vbCrLf & "LangSecRef=3031" & vbCrLf &
                      "FileKey=<Package>\LocalState|*.log" & vbCrLf & "PackageScaffolds=" & vbCrLf & "Archive=Delisted" & vbCrLf}
        }

        Assert.AreEqual(0, RunArchiveBuild(sources, live, archive, outFiles, "-4d", "{archive}"))

        Assert.IsFalse(live.Contains("Retired Store"), live)
        StringAssert.Contains(archive, "[Retired Store *]" & vbCrLf & "LangSecRef=3031" & vbCrLf &
                                       "DetectFile=%LocalAppData%\Packages\Hazel.Retired_abc" & vbCrLf &
                                       "FileKey1=%LocalAppData%\Packages\Hazel.Retired_abc\LocalState|*.log")

    End Sub

    ''' <summary> With <c> -split </c>, a Store app goes to the letter file of its own name </summary>
    <TestMethod()> Public Sub Split_BucketsPackageEntriesByName()

        Dim live As String = Nothing
        Dim archive As String = Nothing
        Dim outFiles As Dictionary(Of String, String) = Nothing
        Dim sources As New Dictionary(Of String, String) From {
            {"A.ini", LiveEntry & "[Zeta Store *]" & vbCrLf & "Package=Hazel.Zeta_abc" & vbCrLf & "LangSecRef=3031" & vbCrLf &
                      "FileKey=<Package>\LocalState|*.log" & vbCrLf & "PackageScaffolds=" & vbCrLf}
        }

        Assert.AreEqual(0, RunArchiveBuild(sources, live, archive, outFiles, "-split"))

        StringAssert.Contains(outFiles("Z.ini"), "[Zeta Store *]" & vbCrLf & "LangSecRef=3031" & vbCrLf & "DetectFile=%LocalAppData%\Packages\Hazel.Zeta_abc")
        Assert.IsFalse(outFiles("A.ini").Contains("Zeta Store"), outFiles("A.ini"))

    End Sub

    ''' <summary> With <c> -split </c>, archived entries still go to the one unsplit <c> archived.ini </c> </summary>
    <TestMethod()> Public Sub Split_LeavesTheArchiveUnsplit()

        Dim live As String = Nothing
        Dim archive As String = Nothing
        Dim outFiles As Dictionary(Of String, String) = Nothing
        Dim sources As New Dictionary(Of String, String) From {{"A.ini", LiveEntry & ArchivedEntry("Retired App", "Archive=Gone")}}

        Assert.AreEqual(0, RunArchiveBuild(sources, live, archive, outFiles, "-split", "-4d", "{archive}"))
        Assert.AreEqual(27, outFiles.Count)
        StringAssert.Contains(archive, "[Retired App *]")

    End Sub

    ''' <summary>
    ''' Builds one live and one archived entry from a source folder at
    ''' <paramref name="sourceFolder"/> below a scratch folder, and returns the headers' text
    ''' </summary>
    '''
    ''' <param name="sourceFolder">The source folder's path below the scratch folder, created as given</param>
    '''
    ''' <param name="sourceArg">The <c> -1d </c> path below the scratch folder, which may differ from <paramref name="sourceFolder"/> in case</param>
    '''
    ''' <param name="split">Indicates whether to pass <c> -split </c></param>
    '''
    ''' <param name="liveText">Set to the text of <c> entrybuilder.ini </c>, or of <c> A.ini </c> with <paramref name="split"/></param>
    '''
    ''' <param name="archiveText">Set to the text of <c> archived.ini </c></param>
    Private Shared Sub RunBuildFrom(sourceFolder As String,
                                    sourceArg As String,
                                    split As Boolean,
                                    ByRef liveText As String,
                                    ByRef archiveText As String)

        Dim dir = IO.Path.Combine(IO.Path.GetTempPath(), $"winapp2ool-ebheader-{Guid.NewGuid():N}")
        Dim sourceDir = IO.Path.Combine(dir, sourceFolder)
        Dim outDir = IO.Path.Combine(dir, "Out")
        Dim archiveDir = IO.Path.Combine(dir, "Archived")
        Dim saved = {winapp2ool.EntryBuilderFile1, winapp2ool.EntryBuilderFile2, winapp2ool.EntryBuilderFile3, winapp2ool.EntryBuilderArchiveFile}
        Dim savedSplit = winapp2ool.EntryBuilderSplitOutput
        Dim savedSuppress = winapp2ool.SuppressOutput
        Dim savedExitCode = Environment.ExitCode

        Try

            IO.Directory.CreateDirectory(sourceDir)
            IO.Directory.CreateDirectory(IO.Path.Combine(dir, "Scaffolds"))
            IO.File.WriteAllText(IO.Path.Combine(sourceDir, "A.ini"), LiveEntry & ArchivedEntry("Retired App", "Archive=Gone"))

            winapp2ool.SuppressOutput = True

            Dim args = {"-1d", IO.Path.Combine(dir, sourceArg), "-2d", outDir, "-3d", IO.Path.Combine(dir, "Scaffolds"), "-4d", archiveDir}
            If split Then args = args.Concat({"-split"}).ToArray()
            setCmdLineArgs(AddressOf winapp2ool.EntryBuilder.handleCmdLine, args)

            liveText = IO.File.ReadAllText(IO.Path.Combine(outDir, If(split, "A.ini", "entrybuilder.ini")))
            archiveText = IO.File.ReadAllText(IO.Path.Combine(archiveDir, "archived.ini"))

        Finally

            winapp2ool.EntryBuilderFile1 = saved(0)
            winapp2ool.EntryBuilderFile2 = saved(1)
            winapp2ool.EntryBuilderFile3 = saved(2)
            winapp2ool.EntryBuilderArchiveFile = saved(3)
            winapp2ool.EntryBuilderSplitOutput = savedSplit
            winapp2ool.SuppressOutput = savedSuppress
            Environment.ExitCode = savedExitCode
            IO.Directory.Delete(dir, recursive:=True)

        End Try

    End Sub

    ''' <summary>
    ''' The live and archive headers link the source folder the run read, as its path from the
    ''' nearest <c> Assembler </c> folder spelled as it is on disk, whatever case <c> -1d </c>
    ''' used; a split letter file links its own source file there
    ''' </summary>
    <TestMethod()> Public Sub Header_LinksTheSourceFolderRead()

        ' The folder to create, its -1d spelling, and the link the headers carry
        Dim cases = {
            {"Assembler\Sources\EntryBuilder", "Assembler\Sources\EntryBuilder", "Assembler/Sources/EntryBuilder"},
            {"Assembler\Sources\Winapp3", "Assembler\Sources\Winapp3", "Assembler/Sources/Winapp3"},
            {"repo\Assembler\Sources\Winapp3", "repo\assembler\sources\winapp3", "Assembler/Sources/Winapp3"}
        }

        For i = 0 To cases.GetLength(0) - 1

            Dim live As String = Nothing
            Dim archive As String = Nothing
            Dim folderLink = $"; Source folder:{vbCrLf};   https://github.com/MoscaDotTo/Winapp2/tree/master/{cases(i, 2)}{vbCrLf}"

            RunBuildFrom(cases(i, 0), cases(i, 1), split:=False, live, archive)
            StringAssert.Contains(live, folderLink)
            StringAssert.Contains(archive, folderLink)

            RunBuildFrom(cases(i, 0), cases(i, 1), split:=True, live, archive)
            StringAssert.Contains(live, $"; Source file:{vbCrLf};   https://github.com/MoscaDotTo/Winapp2/blob/master/{cases(i, 2)}/A.ini{vbCrLf}")
            StringAssert.Contains(archive, folderLink)

        Next

    End Sub

    ''' <summary>
    ''' A source folder with no <c> Assembler </c> ancestor links EntryBuilder's own source folder,
    ''' in the live and the archive headers
    ''' </summary>
    <TestMethod()> Public Sub Header_FallsBackOutsideTheRepository()

        Dim live As String = Nothing
        Dim archive As String = Nothing
        Dim fallbackLink = $"; Source folder:{vbCrLf};   https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/Sources/EntryBuilder{vbCrLf}"

        RunBuildFrom("MySources\Winapp3", "MySources\Winapp3", split:=False, live, archive)

        StringAssert.Contains(live, fallbackLink)
        StringAssert.Contains(archive, fallbackLink)

    End Sub

End Class
