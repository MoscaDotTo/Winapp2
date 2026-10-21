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
''' <c> %Package% </c> alone.
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

End Class
