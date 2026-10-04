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
''' Tests for EntryBuilder's scaffold-family wiring: that each root key is parsed, bound to its
''' placeholder, and opts the entry into its family. The substitution engine itself is covered
''' by <c> ScaffoldCatalogsTests </c>; these tests cover the parser and generator around it.
''' </summary>
<TestClass()> Public Class EntryBuilderTests

    ''' <summary>
    ''' Helper: parse literal ini text and return its first section
    ''' </summary>
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
    ''' Helper: parse then generate one source section against the given QtWebEngine catalog,
    ''' returning the parsed spec's skip flag and the emitted FileKey values
    ''' </summary>
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
    ''' Minimal valid source preamble: a category and a detection, so the entry is not skipped
    ''' for lacking either
    ''' </summary>
    Private Const Preamble As String = "[Test App *]" & vbCrLf &
                                       "LangSecRef=3021" & vbCrLf &
                                       "DetectFile=%LocalAppData%\VideoKeeper" & vbCrLf

    ''' <summary>
    ''' A catalog with one profile-relative template and one cache-relative template
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
