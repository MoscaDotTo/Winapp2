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
''' is parsed, bound to <c> %QtWebEngineCacheRoot% </c> in catalog templates and <c> FileKeyBase= </c>
''' values, and opts the entry into the QtWebEngine family on its own. The substitution engine itself
''' is covered by <see cref="ScaffoldCatalogsTests"/>; these tests cover the parser and generator around it.
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
            {"Caches", New List(Of String) From {"%QtWebEngineRoot%\GPUCache|*", "%QtWebEngineCacheRoot%\Cache|*|RECURSE"}}
        }

    End Function

    ''' <summary>
    ''' <c> QtWebEngineCacheRoot= </c> binds <c> %QtWebEngineCacheRoot% </c> independently of the
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
    ''' emitting a literal <c> %QtWebEngineCacheRoot% </c> into a FileKey
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
    ''' <c> FileKeyBase= </c> substitutes <c> %QtWebEngineCacheRoot% </c> like every other root
    ''' placeholder
    ''' </summary>
    <TestMethod()> Public Sub QtWebEngine_CacheRoot_SubstitutesInFileKeyBase()

        Dim skipped As Boolean
        Dim fileKeys = BuildQt(Preamble &
                               "QtWebEngineCacheRoot=%LocalAppData%\VideoKeeper\cache\QtWebEngine\Default" & vbCrLf &
                               "ExcludeQtWebEngineScaffolds=Caches" & vbCrLf &
                               "FileKeyBase=%QtWebEngineCacheRoot%\Cache\Cache_Data|*" & vbCrLf,
                               QtCatalog(), skipped)

        CollectionAssert.AreEqual(
            New List(Of String) From {"%LocalAppData%\VideoKeeper\cache\QtWebEngine\Default\Cache\Cache_Data|*"},
            fileKeys)

    End Sub

End Class
