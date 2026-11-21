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
''' Also covers the refusal of the retired <c> %...Root% </c> spellings, which leaves
''' <c> %Package% </c> alone, and the <c> Archive= </c> key and its <c> archived.ini </c> output.
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
    ''' <c> %Package% </c> is UWPBuilder's retired token and never meant anything here, so
    ''' EntryBuilder leaves it alone
    ''' </summary>
    <TestMethod()> Public Sub RetiredPackageToken_IsNotRefused()

        Dim log As List(Of String) = Nothing
        Dim exitCode As Integer
        Dim spec = ParseLogged(Preamble & "FileKey=%Package%\LocalState|*.log" & vbCrLf, log, exitCode)

        Assert.AreEqual(0, exitCode)
        CollectionAssert.AreEqual({"%Package%\LocalState|*.log"}, spec.FileKeys)

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
    ''' <param name="outFiles">Set to the names of every file the run wrote in the output folder</param>
    '''
    ''' <param name="archiveArgs">
    ''' The archive folder arguments. <c> {archive} </c> in them stands for the scratch archive
    ''' folder.
    ''' </param>
    Private Shared Function RunArchiveBuild(sources As Dictionary(Of String, String),
                                            ByRef liveText As String,
                                            ByRef archiveText As String,
                                            ByRef outFiles As List(Of String),
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
            outFiles = If(IO.Directory.Exists(outDir), IO.Directory.GetFiles(outDir).Select(Function(f) IO.Path.GetFileName(f)).ToList(), New List(Of String))

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
        Dim outFiles As List(Of String) = Nothing
        Dim sources As New Dictionary(Of String, String) From {{"A.ini", LiveEntry & ArchivedEntry("Retired App", "Archive=Discontinued (2012)")}}

        Assert.AreEqual(0, RunArchiveBuild(sources, live, archive, outFiles, "-4d", "{archive}"))

        StringAssert.Contains(live, "[Alpha App *]")
        Assert.IsFalse(live.Contains("Retired App"), live)
        Assert.IsTrue(archive.StartsWith("; # of entries: 1" & vbCrLf & "; archived.ini is a build artifact generated by winapp2ool's EntryBuilder.", StringComparison.Ordinal), archive)
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
        Dim outFiles As List(Of String) = Nothing
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
        Dim outFiles As List(Of String) = Nothing
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
        Dim outFiles As List(Of String) = Nothing
        Dim sources As New Dictionary(Of String, String) From {{"A.ini", LiveEntry & ArchivedEntry("Retired App", "Archive=Gone")}}

        Assert.AreEqual(1, RunArchiveBuild(sources, live, archive, outFiles))
        StringAssert.Contains(live, "[Alpha App *]")
        Assert.IsFalse(live.Contains("Retired App"), live)
        Assert.IsNull(archive)
        CollectionAssert.AreEqual({"entrybuilder.ini"}, outFiles.ToArray())

    End Sub

    ''' <summary> Given an archive folder, a run with nothing archived still writes a header-only <c> archived.ini </c> </summary>
    <TestMethod()> Public Sub NothingArchived_WritesHeaderOnlyArchive()

        Dim live As String = Nothing
        Dim archive As String = Nothing
        Dim outFiles As List(Of String) = Nothing
        Dim sources As New Dictionary(Of String, String) From {{"A.ini", LiveEntry}}

        Assert.AreEqual(0, RunArchiveBuild(sources, live, archive, outFiles, "-4d", "{archive}"))
        Assert.IsTrue(archive.StartsWith("; # of entries: 0" & vbCrLf, StringComparison.Ordinal), archive)
        Assert.IsFalse(archive.Contains("["), archive)

    End Sub

    ''' <summary> With <c> -split </c>, archived entries still go to the one unsplit <c> archived.ini </c> </summary>
    <TestMethod()> Public Sub Split_LeavesTheArchiveUnsplit()

        Dim live As String = Nothing
        Dim archive As String = Nothing
        Dim outFiles As List(Of String) = Nothing
        Dim sources As New Dictionary(Of String, String) From {{"A.ini", LiveEntry & ArchivedEntry("Retired App", "Archive=Gone")}}

        Assert.AreEqual(0, RunArchiveBuild(sources, live, archive, outFiles, "-split", "-4d", "{archive}"))
        Assert.AreEqual(27, outFiles.Count)
        StringAssert.Contains(archive, "[Retired App *]")

    End Sub

End Class
