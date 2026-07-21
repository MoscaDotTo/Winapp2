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
Imports System.Text

''' <summary>
''' EntryBuilder is a winapp2ool module which generates winapp2.ini entries from a
''' shorthand DSL. It reads per-letter source files under
''' <c> Assembler\EntryBuilder\*.ini </c>, expands winapp2ool-private shorthand keys,
''' and emits a standard winapp2.ini-style output file consumed by the build
''' pipeline. <br /><br />
'''
''' Each section in a source file describes one application and corresponds to exactly
''' one output winapp2.ini entry: the section header is the entry name
''' (e.g. <c> [Discord *] </c>) and is round-tripped verbatim into the output. Any
''' valid winapp2 key passes through unchanged, so a raw entry can be pasted into a 
''' source file and incrementally enriched with shorthand. 
''' The deprecated <c> SpecialDetect= </c> is refused: the key is dropped with a
''' warning rather than passed through or misparsed as a variable declaration. <br /><br />
'''
''' Scaffold families are the first shorthand category. An entry opts into a family by
''' declaring one or more root paths for it: <c> WebViewRoot= </c> (one per WebView2 /
''' EBWebView data root) or <c> QtWebEngineRoot= </c> (one per QtWebEngine data folder).
''' The module then expands the selected scaffolds — <c> WebViewScaffolds= </c> /
''' <c> ExcludeWebViewScaffolds= </c> for WebView2, <c> QtWebEngineScaffolds= </c> /
''' <c> ExcludeQtWebEngineScaffolds= </c> for QtWebEngine — into <c> FileKey </c> lines,
''' substituting the family's placeholder (<c> %WebViewRoot% </c> or
''' <c> %QtWebEngineRoot% </c>) against each declared root. The two families are
''' independent: an entry may declare either, both, or neither. Non-scaffold targets are
''' written via <c> FileKeyBase= </c> / <c> RegKeyBase= </c> with the same substitution
''' rule (either placeholder). An entry without any root declaration emits no scaffold
''' output, which is the correct default for entries that just want pass-through plus
''' future shorthand. The DSL is open to additional scaffold catalogs in future (each
''' introduces its own opt-in trigger key). <br /><br />
'''
''' Recognised shorthand keys (stripped from output):
''' <c> WebViewRoot= </c>, <c> WebViewScaffolds= </c>, <c> ExcludeWebViewScaffolds= </c>,
''' <c> QtWebEngineRoot= </c>, <c> QtWebEngineScaffolds= </c>,
''' <c> ExcludeQtWebEngineScaffolds= </c>, <c> FileKeyBase= </c>, <c> RegKeyBase= </c>,
''' <c> ExcludeKeyBase= </c>, <c> Skip= </c>. <br /><br />
'''
''' Scaffolds are loaded from two shared catalogs via <see cref="WebViewScaffolds.LoadCatalog"/>:
''' <c> Assembler\Scaffolds\webview.ini </c> (<c> %WebViewRoot% </c> templates) and
''' <c> Assembler\Scaffolds\qtwebengine.ini </c> (<c> %QtWebEngineRoot% </c> templates),
''' each substituted per declared root at generation time. <br /><br />
'''
''' Unlike <c> UWPBuilder </c>, this module never emits <c> Default= </c>. Generated
''' entries are opt-in by design to keep the cleaning UI uncluttered as small Electron
''' apps proliferate.
''' </summary>
Public Module EntryBuilder

    ''' <summary>
    ''' Default scaffold names emitted when an entry does not declare an explicit
    ''' <c> WebViewScaffolds= </c> selection. Limited to low-risk categories; host-risk
    ''' scaffolds (cookies, history, session, web storage) require explicit opt-in
    ''' by the entry author.
    ''' </summary>
    Public ReadOnly DefaultScaffolds As String() = {"Caches", "Telemetry"}

    ''' <summary>
    ''' Nested-reference count at or above which an entry's variable block is reported as a
    ''' tangled abstraction in the run statistics. The trigger is <em>interdependency</em>,
    ''' not raw variable count: real run data shows variable count flags the DSL's best work
    ''' (e.g. an entry declaring 9 flat variables that fan out to 430 keys — a triumph, not a
    ''' concern), whereas the genuinely hard-to-read blocks are the ones whose variables
    ''' reference each other, forcing the reader to topologically sort a dependency graph.
    ''' Across the source the distribution is near-bimodal — a large cluster at one nested ref
    ''' (cheap one-hop convenience vars) and a small tail at 3-5 (the actual tangles) — so the
    ''' threshold sits in that gap. Set to 4 to surface only the severe cases (ABBYY-, Acrobat-
    ''' class); lower to 3 to also catch shallow multi-variable chains.
    ''' </summary>
    Private Const NestedRefWatchThreshold As Integer = 4

    ''' <summary>
    ''' The set of first-class key names the parser recognises (case-insensitive,
    ''' matched after <c> StripNums </c>). Anything not in this set is treated as a
    ''' variable declaration by <see cref="parseEntrySpec"/>. Centralised here so the
    ''' parser's <c> Select Case </c> and the reserved-name collision check cannot
    ''' drift apart.
    ''' </summary>
    Private ReadOnly ReservedKeys As String() = {
        "SECTION", "LANGSECREF",
        "WEBVIEWROOT", "QTWEBENGINEROOT",
        "DETECT", "DETECTFILE", "DETECTOS", "SPECIALDETECT",
        "FILEKEY", "FILEKEYBASE",
        "REGKEY", "REGKEYBASE",
        "EXCLUDEKEY", "EXCLUDEKEYBASE",
        "WEBVIEWSCAFFOLDS", "EXCLUDEWEBVIEWSCAFFOLDS",
        "QTWEBENGINESCAFFOLDS", "EXCLUDEQTWEBENGINESCAFFOLDS",
        "SKIP", "DEFAULT", "WARNING"
    }

    ''' <summary>
    ''' Parsed state for a single source-file section. One section corresponds to exactly
    ''' one winapp2 output entry; the section header is the entry name.
    ''' </summary>
    Private Structure EntrySpec

        ''' <summary>
        ''' The entry name used verbatim as the section header in the output, taken
        ''' from the source section header (e.g. <c> Discord * </c>)
        ''' </summary>
        Public Name As String

        ''' <summary>
        ''' The <c> Section= </c> value, or empty if <c> LangSecRef= </c> is used
        ''' </summary>
        Public SectionName As String

        ''' <summary>
        ''' The <c> LangSecRef= </c> value, or empty if <c> Section= </c> is used
        ''' </summary>
        Public LangSecRef As String

        ''' <summary>
        ''' Root paths of the application's Chromium-data folders. Each entry is a
        ''' literal path string used as the substitution value for <c> %WebViewRoot% </c>
        ''' in scaffold and FileKeyBase templates. Multiple roots produce one scaffold
        ''' expansion per root.
        ''' </summary>
        Public WebViewRoots As List(Of String)

        ''' <summary>
        ''' <c> DetectFile= </c> values in their original file order
        ''' </summary>
        Public DetectFiles As List(Of String)

        ''' <summary>
        ''' <c> Detect= </c> values in their original file order
        ''' </summary>
        Public Detects As List(Of String)

        ''' <summary>
        ''' The <c> DetectOS= </c> value, or empty if not declared
        ''' </summary>
        Public DetectOS As String

        ''' <summary>
        ''' Pass-through <c> Warning= </c> values in file order, emitted verbatim between
        ''' detection and deletion keys per canonical winapp2 key order. Prose text is
        ''' never subject to variable expansion
        ''' </summary>
        Public Warnings As List(Of String)

        ''' <summary>
        ''' <c> FileKeyBase= </c> templates (the generative form) in file order, appended
        ''' after the scaffold expansion. Templates containing <c> %WebViewRoot% </c> are
        ''' expanded once per declared root; templates without the variable are emitted
        ''' verbatim. Tracked separately from <see cref="FileKeys"/> so the statistics can
        ''' distinguish generated output from pass-through.
        ''' </summary>
        Public FileKeyBases As List(Of String)

        ''' <summary>
        ''' Pass-through <c> FileKey= </c> values in file order — keys the author wrote
        ''' literally rather than via a <c> Base </c> template. Still subject to
        ''' <c> %WebViewRoot% </c> and <c> &lt;var&gt; </c> expansion, but counted as
        ''' pass-through (not generated) in the run statistics.
        ''' </summary>
        Public FileKeys As List(Of String)

        ''' <summary>
        ''' <c> RegKeyBase= </c> templates (the generative form) in file order. Expanded
        ''' against <c> %WebViewRoot% </c> / variables and renumbered from 1. Tracked
        ''' separately from <see cref="RegKeys"/> for the generated-vs-pass-through split.
        ''' </summary>
        Public RegKeyBases As List(Of String)

        ''' <summary>
        ''' Pass-through <c> RegKey= </c> values in file order — keys the author wrote
        ''' literally rather than via a <c> Base </c> template. Counted as pass-through
        ''' (not generated) in the run statistics.
        ''' </summary>
        Public RegKeys As List(Of String)

        ''' <summary>
        ''' <c> ExcludeKeyBase= </c> templates (the generative form) in file order.
        ''' Templates containing <c> %WebViewRoot% </c> are expanded once per declared root;
        ''' others are emitted verbatim. Renumbered from 1. Tracked separately from
        ''' <see cref="ExcludeKeys"/> for the generated-vs-pass-through split.
        ''' </summary>
        Public ExcludeKeyBases As List(Of String)

        ''' <summary>
        ''' Pass-through <c> ExcludeKey= </c> values in file order — keys the author wrote
        ''' literally rather than via a <c> Base </c> template. Counted as pass-through
        ''' (not generated) in the run statistics.
        ''' </summary>
        Public ExcludeKeys As List(Of String)

        ''' <summary>
        ''' Scaffold names explicitly selected via <c> WebViewScaffolds= </c>. The sentinel
        ''' <c> All </c> (case-insensitive) expands to every scaffold in the catalog.
        ''' </summary>
        Public WebViewScaffoldNames As List(Of String)

        ''' <summary>
        ''' Tracks whether <c> WebViewScaffolds= </c> was declared at all, separately from
        ''' whether the resulting list is empty. Distinguishes "key absent → use
        ''' <see cref="DefaultScaffolds"/>" from "key present but empty → emit no
        ''' scaffold FileKeys."
        ''' </summary>
        Public WebViewScaffoldsKeyPresent As Boolean

        ''' <summary>
        ''' Scaffold names to subtract from the active selection (either
        ''' <see cref="WebViewScaffoldNames"/> when set, or <see cref="DefaultScaffolds"/>
        ''' otherwise). Unknown names are ignored.
        ''' </summary>
        Public ExcludedWebViewScaffolds As List(Of String)

        ''' <summary>
        ''' Root paths of the application's QtWebEngine data folders (the folder literally
        ''' named <c> QtWebEngine </c>). Each entry is a literal path string substituted for
        ''' <c> %QtWebEngineRoot% </c> in QtWebEngine scaffold and base templates. Declaring
        ''' any root opts the entry into QtWebEngine scaffold emission.
        ''' </summary>
        Public QtWebEngineRoots As List(Of String)

        ''' <summary>
        ''' QtWebEngine scaffold names explicitly selected via <c> QtWebEngineScaffolds= </c>.
        ''' The sentinel <c> All </c> (case-insensitive) expands to every scaffold in the
        ''' QtWebEngine catalog.
        ''' </summary>
        Public QtWebEngineScaffoldNames As List(Of String)

        ''' <summary>
        ''' Tracks whether <c> QtWebEngineScaffolds= </c> was declared at all, separately from
        ''' whether the resulting list is empty (mirrors <see cref="WebViewScaffoldsKeyPresent"/>).
        ''' </summary>
        Public QtWebEngineScaffoldsKeyPresent As Boolean

        ''' <summary>
        ''' QtWebEngine scaffold names to subtract from the active selection (either
        ''' <see cref="QtWebEngineScaffoldNames"/> when set, or
        ''' <see cref="WebViewScaffolds.QtWebEngineDefaultScaffolds"/> otherwise).
        ''' Unknown names are ignored.
        ''' </summary>
        Public ExcludedQtWebEngineScaffolds As List(Of String)

        ''' <summary>
        ''' When True, the entry is structurally invalid or marked <c> Skip= </c>
        ''' and is omitted from generation
        ''' </summary>
        Public ShouldSkip As Boolean

        ''' <summary>
        ''' When True, <see cref="inferRootDetection"/> synthesised a <c> Detect=&lt;Root&gt; </c>
        ''' or <c> DetectFile=&lt;Root&gt; </c> for this entry from a reserved <c> Root </c>
        ''' variable. Recorded so the run statistics can report how many entries received
        ''' inferred (rather than hand-authored) detection.
        ''' </summary>
        Public RootDetectionInferred As Boolean

        ''' <summary>
        ''' List variables declared as entry-level keys (<c> Name=v1,v2,v3 </c>),
        ''' referenced via <c> &lt;Name&gt; </c> placeholders in other key values
        ''' and fanned out at generation time by <see cref="VariableExpander"/>.
        ''' </summary>
        Public Variables As VariableSet

        ''' <summary>
        ''' Number of declared variables that reference another declared variable, captured
        ''' before <see cref="VariableExpander.ResolveAll"/> flattens the symbol table.
        ''' Feeds the abstraction-payoff statistic: a large, deeply-nested symbol table is a
        ''' sign the variable block has become noise rather than pattern-surfacing shorthand.
        ''' </summary>
        Public NestedVariableRefs As Integer

        ''' <summary>
        ''' Creates a new <c> EntrySpec </c> for an entry with the given name,
        ''' initialising all list fields to empty collections
        ''' </summary>
        '''
        ''' <param name="name">
        ''' The entry name, taken verbatim from the source section header
        ''' </param>
        Public Sub New(name As String)

            Me.Name = name
            SectionName = ""
            LangSecRef = ""
            WebViewRoots = New List(Of String)
            DetectFiles = New List(Of String)
            Detects = New List(Of String)
            DetectOS = ""
            Warnings = New List(Of String)
            FileKeyBases = New List(Of String)
            FileKeys = New List(Of String)
            RegKeyBases = New List(Of String)
            RegKeys = New List(Of String)
            ExcludeKeyBases = New List(Of String)
            ExcludeKeys = New List(Of String)
            WebViewScaffoldNames = New List(Of String)
            WebViewScaffoldsKeyPresent = False
            ExcludedWebViewScaffolds = New List(Of String)
            QtWebEngineRoots = New List(Of String)
            QtWebEngineScaffoldNames = New List(Of String)
            QtWebEngineScaffoldsKeyPresent = False
            ExcludedQtWebEngineScaffolds = New List(Of String)
            ShouldSkip = False
            RootDetectionInferred = False
            Variables = New VariableSet
            NestedVariableRefs = 0

        End Sub

    End Structure

    ''' <summary>
    ''' One entry whose variable block crossed <see cref="NestedRefWatchThreshold"/>, recorded
    ''' for the abstraction-payoff section of the run statistics. Carries the figures needed
    ''' to judge whether the symbol table earns its indirection: how many variables it declares,
    ''' how deeply they nest, and how many output keys they ultimately produce.
    ''' </summary>
    Private Structure TangledEntry

        ''' <summary> The entry name (source section header) </summary>
        Public Name As String

        ''' <summary> Number of list variables declared on the entry </summary>
        Public VarCount As Integer

        ''' <summary> Number of those variables that reference another declared variable </summary>
        Public NestedRefs As Integer

        ''' <summary> Number of output keys the entry generated </summary>
        Public KeyCount As Integer

        ''' <summary>
        ''' Creates a heavy-entry record
        ''' </summary>
        '''
        ''' <param name="name"> The entry name </param>
        '''
        ''' <param name="varCount"> Declared variable count </param>
        '''
        ''' <param name="nestedRefs"> Nested-reference count </param>
        '''
        ''' <param name="keyCount"> Generated key count </param>
        Public Sub New(name As String, varCount As Integer, nestedRefs As Integer, keyCount As Integer)

            Me.Name = name
            Me.VarCount = varCount
            Me.NestedRefs = nestedRefs
            Me.KeyCount = keyCount

        End Sub

    End Structure

    ''' <summary>
    ''' Mutable accumulator for a single EntryBuilder run's generation statistics. One
    ''' instance is threaded through <see cref="generateEntry"/> and the expansion helpers
    ''' so every emitted key is tallied at its source, then rendered as a summary box by
    ''' <see cref="renderStats"/>. A reference type by design — sharing one instance across
    ''' the whole run lets counts accumulate without <c> ByRef </c> plumbing.
    ''' </summary>
    Private NotInheritable Class EntryBuilderStats

        ''' <summary> Total sections read from the combined source directory </summary>
        Public SourceSections As Integer

        ''' <summary> Sections that produced an output entry (were not skipped) </summary>
        Public EntriesGenerated As Integer

        ''' <summary> Sections skipped as structurally invalid or explicitly marked <c> Skip= </c> </summary>
        Public EntriesSkipped As Integer

        ''' <summary> Generated entries that declared at least one <c> WebViewRoot= </c> </summary>
        Public EntriesWithScaffolds As Integer

        ''' <summary> Generated entries that declared at least one <c> QtWebEngineRoot= </c> </summary>
        Public EntriesWithQtScaffolds As Integer

        ''' <summary> Generated entries that received a synthesised <c> Root </c>-based detection </summary>
        Public InferredRootEntries As Integer

        ''' <summary> Total <c> WebViewRoot= </c> declarations across all generated entries </summary>
        Public WebViewRootDecls As Integer

        ''' <summary> Total <c> QtWebEngineRoot= </c> declarations across all generated entries </summary>
        Public QtWebEngineRootDecls As Integer

        ''' <summary> Total <c> FileKeyBase= </c> declarations (generative) across all generated entries </summary>
        Public FileKeyBaseDecls As Integer

        ''' <summary> Total <c> FileKey= </c> declarations (pass-through) across all generated entries </summary>
        Public FileKeyDecls As Integer

        ''' <summary> Total <c> RegKeyBase= </c> declarations (generative) across all generated entries </summary>
        Public RegKeyBaseDecls As Integer

        ''' <summary> Total <c> RegKey= </c> declarations (pass-through) across all generated entries </summary>
        Public RegKeyDecls As Integer

        ''' <summary> Total <c> ExcludeKeyBase= </c> declarations (generative) across all generated entries </summary>
        Public ExcludeKeyBaseDecls As Integer

        ''' <summary> Total <c> ExcludeKey= </c> declarations (pass-through) across all generated entries </summary>
        Public ExcludeKeyDecls As Integer

        ''' <summary> Category keys emitted (<c> LangSecRef </c> / <c> Section </c>) </summary>
        Public CategoryKeys As Integer

        ''' <summary> <c> Detect </c> keys emitted, after variable fan-out </summary>
        Public DetectKeys As Integer

        ''' <summary> <c> DetectFile </c> keys emitted, after variable fan-out </summary>
        Public DetectFileKeys As Integer

        ''' <summary> <c> DetectOS </c> keys emitted </summary>
        Public DetectOSKeys As Integer

        ''' <summary> <c> Warning </c> keys emitted (pass-through, verbatim) </summary>
        Public WarningKeys As Integer

        ''' <summary> <c> FileKey </c> keys produced by WebView scaffold catalog expansion (generated) </summary>
        Public ScaffoldFileKeys As Integer

        ''' <summary> <c> FileKey </c> keys produced by QtWebEngine scaffold catalog expansion (generated) </summary>
        Public QtScaffoldFileKeys As Integer

        ''' <summary> <c> FileKey </c> keys produced from <c> FileKeyBase= </c> templates (generated) </summary>
        Public FileKeyBaseKeys As Integer

        ''' <summary> <c> FileKey </c> keys produced from literal <c> FileKey= </c> declarations (pass-through) </summary>
        Public FileKeyPassThroughKeys As Integer

        ''' <summary> <c> RegKey </c> keys produced from <c> RegKeyBase= </c> templates (generated) </summary>
        Public RegKeyBaseKeys As Integer

        ''' <summary> <c> RegKey </c> keys produced from literal <c> RegKey= </c> declarations (pass-through) </summary>
        Public RegKeyPassThroughKeys As Integer

        ''' <summary> <c> ExcludeKey </c> keys produced from <c> ExcludeKeyBase= </c> templates (generated) </summary>
        Public ExcludeKeyBaseKeys As Integer

        ''' <summary> <c> ExcludeKey </c> keys produced from literal <c> ExcludeKey= </c> declarations (pass-through) </summary>
        Public ExcludeKeyPassThroughKeys As Integer

        ''' <summary>
        ''' Extra keys produced by <c> &lt;var&gt; </c> cartesian fan-out — the values
        ''' beyond the first that each template contributed. Counts only the phase-2
        ''' variable engine's multiplication, not the per-root <c> %WebViewRoot% </c> fan-out
        ''' (already reflected in the scaffold / base FileKey counts).
        ''' </summary>
        Public VariableFanoutKeys As Integer

        ''' <summary> Total list variables declared across all generated entries </summary>
        Public TotalVariablesDeclared As Integer

        ''' <summary>
        ''' Total nested variable references (declarations that reference another declaration)
        ''' across all generated entries — the run-wide topological-sort tax
        ''' </summary>
        Public NestedVariableRefs As Integer

        ''' <summary> Largest single-entry variable count seen this run </summary>
        Public MaxVariablesInEntry As Integer

        ''' <summary> Name of the entry holding the largest symbol table </summary>
        Public MaxVariablesEntryName As String = ""

        ''' <summary>
        ''' Entries whose nested-reference count reached <see cref="NestedRefWatchThreshold"/> —
        ''' the abstraction-payoff watch list, surfaced so the maintainer can judge whether those
        ''' interdependent declaration blocks still earn their indirection
        ''' </summary>
        Public ReadOnly TangledEntries As New List(Of TangledEntry)

        ''' <summary>
        ''' Content keys the builder generated from a template — scaffold expansion plus
        ''' every <c> FileKeyBase= </c> / <c> RegKeyBase= </c> / <c> ExcludeKeyBase= </c>
        ''' </summary>
        Public ReadOnly Property GeneratedContentKeys As Integer
            Get
                Return ScaffoldFileKeys + QtScaffoldFileKeys + FileKeyBaseKeys + RegKeyBaseKeys + ExcludeKeyBaseKeys
            End Get
        End Property

        ''' <summary>
        ''' Content keys passed through from literal <c> FileKey= </c> / <c> RegKey= </c> /
        ''' <c> ExcludeKey= </c> declarations the author wrote directly
        ''' </summary>
        Public ReadOnly Property PassThroughContentKeys As Integer
            Get
                Return FileKeyPassThroughKeys + RegKeyPassThroughKeys + ExcludeKeyPassThroughKeys
            End Get
        End Property

        ''' <summary> Total keys emitted across every generated entry (sum of the per-category counters) </summary>
        Public ReadOnly Property TotalKeys As Integer
            Get
                Return CategoryKeys + DetectKeys + DetectFileKeys + DetectOSKeys + WarningKeys +
                       GeneratedContentKeys + PassThroughContentKeys
            End Get
        End Property

    End Class

    ''' <summary>
    ''' Handles the command-line arguments for <c> EntryBuilder </c>
    ''' </summary>
    '''
    ''' <remarks>
    ''' Flags:
    ''' <list type="bullet">
    ''' <item><c> -split </c> — toggles per-letter artifact output (<c> #.ini </c> ...
    ''' <c> Z.ini </c> in the save target's directory) instead of a single output file</item>
    ''' </list>
    ''' </remarks>
    Public Sub handleCmdLine()

        Dim spec As New CliArgSpec("entrybuilder")
        spec.WithFile(1, EntryBuilderFile1) _
            .WithFile(2, EntryBuilderFile2) _
            .WithFile(3, EntryBuilderFile3) _
            .WithFile(4, EntryBuilderFile4) _
            .WithFlag("-split", Sub() EntryBuilderSplitOutput = Not EntryBuilderSplitOutput) _
            .Parse()

        initEntryBuilder()

    End Sub

    ''' <summary>
    ''' Reads the source directory, runs the generation pipeline, writes the output file,
    ''' and displays the results. Bails early with a header message if the source directory
    ''' contains no parseable sections.
    ''' </summary>
    Public Sub initEntryBuilder()

        clrConsole()

        Dim sourceDir = EntryBuilderFile1.Dir
        Dim sourceIni = combineSourceDir(sourceDir)

        If sourceIni.Count = 0 Then

            setNextMenuHeaderText($"No EntryBuilder source definitions found in: {sourceDir}", printColor:=ConsoleColor.Red)
            Return

        End If

        Dim catalogPath = EntryBuilderFile3.Path()
        Dim qtCatalogPath = EntryBuilderFile4.Path()

        Dim output As New MenuSection
        output.AddBoxWithText("Building shorthand-DSL entries")

        Using gLogScope("Building shorthand-DSL entries")

            processEntryBuilder(sourceIni, catalogPath, qtCatalogPath, output)

            output.AddBoxWithText("Entries built successfully")
            gLog("Entries built successfully")

        End Using

        output.AddAnyKeyPrompt()

        If SuppressOutput Then Return

        output.Print()
        crk()

    End Sub

    ''' <summary>
    ''' Orchestrates the EntryBuilder pipeline: loads the scaffold catalog, parses each
    ''' source section into an <c> EntrySpec </c>, generates one <c> iniSection2 </c>
    ''' per non-skipped entry, and writes the result to the output file with a header
    ''' comment block.
    ''' </summary>
    '''
    ''' <param name="sourceIni">
    ''' The combined contents of every <c> *.ini </c> in the source directory
    ''' </param>
    '''
    ''' <param name="catalogPath">
    ''' Absolute path to the shared WebView scaffold catalog (typically
    ''' <c> Assembler\Scaffolds\webview.ini </c>). Loaded once per run via
    ''' <see cref="WebViewScaffolds.LoadCatalog"/>.
    ''' </param>
    '''
    ''' <param name="qtCatalogPath">
    ''' Absolute path to the shared QtWebEngine scaffold catalog (typically
    ''' <c> Assembler\Scaffolds\qtwebengine.ini </c>). Loaded once per run via
    ''' <see cref="WebViewScaffolds.LoadCatalog"/> with the <c> QtWebEngineScaffold: </c> prefix.
    ''' </param>
    '''
    ''' <param name="menuOutput">
    ''' The <c> MenuSection </c> receiving progress lines and warnings for display
    ''' </param>
    Private Sub processEntryBuilder(sourceIni As iniFile2,
                                     catalogPath As String,
                                     qtCatalogPath As String,
                                     menuOutput As MenuSection)

        Using gLogScope("Processing EntryBuilder source files")

            Dim catalog = WebViewScaffolds.LoadCatalog(catalogPath, menuOutput)

            Dim catalogMsg = $"Loaded {catalog.Count} WebView scaffold(s) from {catalogPath}"
            menuOutput.AddColoredLine(catalogMsg, ConsoleColor.Yellow)
            gLog(catalogMsg)

            Dim qtCatalog = WebViewScaffolds.LoadCatalog(qtCatalogPath, menuOutput, "QtWebEngineScaffold:")

            Dim qtCatalogMsg = $"Loaded {qtCatalog.Count} QtWebEngine scaffold(s) from {qtCatalogPath}"
            menuOutput.AddColoredLine(qtCatalogMsg, ConsoleColor.Yellow)
            gLog(qtCatalogMsg)

            Dim entries As New List(Of EntrySpec)

            For Each section In sourceIni

                Dim spec = parseEntrySpec(section, menuOutput)
                If Not spec.ShouldSkip Then entries.Add(spec)

            Next

            Dim entriesMsg = $"Loaded {entries.Count} entry definition(s)"
            menuOutput.AddColoredLine(entriesMsg, ConsoleColor.Yellow)
            gLog(entriesMsg)

            Dim stats As New EntryBuilderStats With {
                .SourceSections = sourceIni.Count,
                .EntriesGenerated = entries.Count,
                .EntriesSkipped = sourceIni.Count - entries.Count
            }

            Dim outputFile = iniFile2.Empty(EntryBuilderFile2.Dir, EntryBuilderFile2.Name)

            For Each spec In entries

                stats.WebViewRootDecls += spec.WebViewRoots.Count
                stats.QtWebEngineRootDecls += spec.QtWebEngineRoots.Count
                stats.FileKeyBaseDecls += spec.FileKeyBases.Count
                stats.FileKeyDecls += spec.FileKeys.Count
                stats.RegKeyBaseDecls += spec.RegKeyBases.Count
                stats.RegKeyDecls += spec.RegKeys.Count
                stats.ExcludeKeyBaseDecls += spec.ExcludeKeyBases.Count
                stats.ExcludeKeyDecls += spec.ExcludeKeys.Count
                If spec.WebViewRoots.Count > 0 Then stats.EntriesWithScaffolds += 1
                If spec.QtWebEngineRoots.Count > 0 Then stats.EntriesWithQtScaffolds += 1
                If spec.RootDetectionInferred Then stats.InferredRootEntries += 1

                Dim entrySection = generateEntry(spec, catalog, qtCatalog, stats, menuOutput)
                outputFile.AddSection(entrySection)

                ' Abstraction payoff: an interdependent symbol table — variables referencing
                ' other variables — is the maintainer's signal that a declaration block has
                ' become noise rather than pattern-surfacing shorthand, since it forces the
                ' reader to topologically sort a dependency graph. Raw variable count is NOT
                ' the trigger: a large flat block of high-fan-out vars is the DSL succeeding.
                ' Tallied against the entry's realized key count so the leverage is visible.
                Dim varCount = spec.Variables.DeclaredNames().Count
                stats.TotalVariablesDeclared += varCount
                stats.NestedVariableRefs += spec.NestedVariableRefs

                If varCount > stats.MaxVariablesInEntry Then

                    stats.MaxVariablesInEntry = varCount
                    stats.MaxVariablesEntryName = spec.Name

                End If

                If spec.NestedVariableRefs >= NestedRefWatchThreshold Then

                    stats.TangledEntries.Add(New TangledEntry(spec.Name, varCount, spec.NestedVariableRefs, entrySection.Keys.Count))
                    gLog($"[{spec.Name}] tangled symbol table: {spec.NestedVariableRefs} nested refs across {varCount} variables -> {entrySection.Keys.Count} keys")

                End If

                ' Typo backstop: variables declared on the entry but never referenced
                ' by any <token> in any expanded key. Mirrors WebViewBuilder's old
                ' "Unexpected key type" warning but now scoped to actually-unused
                ' declarations rather than every unknown key.
                For Each unused In spec.Variables.UnreferencedNames()

                    Dim unusedMsg = $"Variable '{unused}=' declared in [{spec.Name}] but never referenced by any <{unused}> token; possible typo"
                    gLog(unusedMsg)
                    menuOutput.AddWarning(unusedMsg)

                Next

            Next

            Dim generatedMsg = $"Generated {outputFile.Count} entries"
            menuOutput.AddColoredLine(generatedMsg, ConsoleColor.Yellow)
            gLog($" {generatedMsg}")

            renderStats(stats, menuOutput)

            outputFile = remotedebugGuarded(outputFile, NameOf(EntryBuilder), menuOutput)

            If EntryBuilderSplitOutput Then

                writeSplitOutput(outputFile, menuOutput)

            Else

                Dim sb As New StringBuilder()
                sb.AppendLine($"; # of entries: {outputFile.Count:#,###}")
                sb.AppendLine($"; {EntryBuilderFile2.Name} is generated by the Winapp2ool Entry Builder")
                sb.AppendLine("; Entries in this file may be incomplete and are not intended to be used directly with any cleaning software")
                sb.AppendLine("; They are utilized by winapp2ool to create the final winapp2.ini file for distribution")
                sb.AppendLine("; If you are not maintaining winapp2.ini for distribution, you probably don't need this file!")
                sb.AppendLine("; Refer to the Winapp2ool documentation for more information: " & readMeUrl)
                sb.AppendLine("; You can find the complete winapp2.ini file here: " & baseFlavorLink)
                sb.AppendLine()
                sb.Append(outputFile.ToString())

                outputFile.OverwriteToFile(sb.ToString())

            End If

            gLog("EntryBuilder source files processed successfully")

        End Using

    End Sub

    ''' <summary>
    ''' Formats one statistic as an aligned <c> label … value </c> line for the summary box.
    ''' Values are right-aligned in a fixed column with thousands separators so the figures
    ''' line up regardless of label length.
    ''' </summary>
    '''
    ''' <param name="label">
    ''' The human-readable statistic name (printed left-aligned)
    ''' </param>
    '''
    ''' <param name="value">
    ''' The count to display (printed right-aligned, thousands-separated)
    ''' </param>
    '''
    ''' <returns>
    ''' A single padded line ready to hand to <c> AddColoredLine </c>
    ''' </returns>
    Private Function statLine(label As String, value As Integer) As String

        Return String.Format("  {0,-28}{1,8:#,##0}", label, value)

    End Function

    ''' <summary>
    ''' Renders the run statistics as a summary box on <paramref name="menuOutput"/> and
    ''' mirrors a compact form to <c> gLog </c>. Grouped into entry counts, the shorthand
    ''' input declarations (the "bases" keys are generated from), the generated-key breakdown
    ''' by provenance, and a generated-vs-pass-through summary. Pass-through tallies count
    ''' keys the author wrote as literal <c> FileKey= </c> / <c> RegKey= </c> /
    ''' <c> ExcludeKey= </c>; generated tallies count scaffold expansion plus the
    ''' <c> ...Base= </c> template forms.
    ''' </summary>
    '''
    ''' <param name="stats">
    ''' The populated statistics accumulator for the completed run
    ''' </param>
    '''
    ''' <param name="menuOutput">
    ''' The <c> MenuSection </c> receiving the summary lines for display
    ''' </param>
    Private Sub renderStats(stats As EntryBuilderStats, menuOutput As MenuSection)

        menuOutput.AddBoxWithText("EntryBuilder statistics")

        menuOutput.AddColoredLine("Entries", ConsoleColor.Cyan)
        menuOutput.AddColoredLine(statLine("Source sections read:", stats.SourceSections), ConsoleColor.White)
        menuOutput.AddColoredLine(statLine("Generated:", stats.EntriesGenerated), ConsoleColor.Green)
        menuOutput.AddColoredLine(statLine("Skipped:", stats.EntriesSkipped), If(stats.EntriesSkipped > 0, ConsoleColor.Yellow, ConsoleColor.DarkGray))
        menuOutput.AddColoredLine(statLine("With WebView scaffolds:", stats.EntriesWithScaffolds), ConsoleColor.White)
        menuOutput.AddColoredLine(statLine("With QtWebEngine scaffolds:", stats.EntriesWithQtScaffolds), ConsoleColor.White)
        menuOutput.AddColoredLine(statLine("With inferred Root:", stats.InferredRootEntries), ConsoleColor.White)

        menuOutput.AddBlank()
        menuOutput.AddColoredLine("Shorthand input declarations", ConsoleColor.Cyan)
        menuOutput.AddColoredLine(statLine("WebViewRoot:", stats.WebViewRootDecls), ConsoleColor.White)
        menuOutput.AddColoredLine(statLine("QtWebEngineRoot:", stats.QtWebEngineRootDecls), ConsoleColor.White)
        menuOutput.AddColoredLine(statLine("FileKeyBase (generated):", stats.FileKeyBaseDecls), ConsoleColor.White)
        menuOutput.AddColoredLine(statLine("FileKey (pass-through):", stats.FileKeyDecls), ConsoleColor.White)
        menuOutput.AddColoredLine(statLine("RegKeyBase (generated):", stats.RegKeyBaseDecls), ConsoleColor.White)
        menuOutput.AddColoredLine(statLine("RegKey (pass-through):", stats.RegKeyDecls), ConsoleColor.White)
        menuOutput.AddColoredLine(statLine("ExcludeKeyBase (generated):", stats.ExcludeKeyBaseDecls), ConsoleColor.White)
        menuOutput.AddColoredLine(statLine("ExcludeKey (pass-through):", stats.ExcludeKeyDecls), ConsoleColor.White)

        menuOutput.AddBlank()
        menuOutput.AddColoredLine(statLine("Keys generated:", stats.TotalKeys), ConsoleColor.Yellow)
        menuOutput.AddColoredLine(statLine("Category:", stats.CategoryKeys), ConsoleColor.White)
        menuOutput.AddColoredLine(statLine("Detect:", stats.DetectKeys), ConsoleColor.White)
        menuOutput.AddColoredLine(statLine("DetectFile:", stats.DetectFileKeys), ConsoleColor.White)
        menuOutput.AddColoredLine(statLine("DetectOS:", stats.DetectOSKeys), ConsoleColor.White)
        menuOutput.AddColoredLine(statLine("Warning:", stats.WarningKeys), ConsoleColor.White)
        menuOutput.AddColoredLine(statLine("FileKey (WebView scaffold):", stats.ScaffoldFileKeys), ConsoleColor.White)
        menuOutput.AddColoredLine(statLine("FileKey (Qt scaffold):", stats.QtScaffoldFileKeys), ConsoleColor.White)
        menuOutput.AddColoredLine(statLine("FileKey (base):", stats.FileKeyBaseKeys), ConsoleColor.White)
        menuOutput.AddColoredLine(statLine("FileKey (pass-through):", stats.FileKeyPassThroughKeys), ConsoleColor.White)
        menuOutput.AddColoredLine(statLine("RegKey (base):", stats.RegKeyBaseKeys), ConsoleColor.White)
        menuOutput.AddColoredLine(statLine("RegKey (pass-through):", stats.RegKeyPassThroughKeys), ConsoleColor.White)
        menuOutput.AddColoredLine(statLine("ExcludeKey (base):", stats.ExcludeKeyBaseKeys), ConsoleColor.White)
        menuOutput.AddColoredLine(statLine("ExcludeKey (pass-through):", stats.ExcludeKeyPassThroughKeys), ConsoleColor.White)

        menuOutput.AddBlank()
        menuOutput.AddColoredLine("Provenance (content keys)", ConsoleColor.Cyan)
        menuOutput.AddColoredLine(statLine("Generated (scaffold + base):", stats.GeneratedContentKeys), ConsoleColor.Green)
        menuOutput.AddColoredLine(statLine("Passed through (literal):", stats.PassThroughContentKeys), ConsoleColor.White)
        menuOutput.AddColoredLine(statLine("Variable fan-out keys:", stats.VariableFanoutKeys), ConsoleColor.White)

        renderAbstractionPayoff(stats, menuOutput)

        gLog($"Stats: {stats.EntriesGenerated} entries ({stats.EntriesSkipped} skipped) -> {stats.TotalKeys} keys; " &
             $"generated {stats.GeneratedContentKeys} [scaffold {stats.ScaffoldFileKeys}, Qt scaffold {stats.QtScaffoldFileKeys}, FileKeyBase {stats.FileKeyBaseKeys}, " &
             $"RegKeyBase {stats.RegKeyBaseKeys}, ExcludeKeyBase {stats.ExcludeKeyBaseKeys}], " &
             $"pass-through {stats.PassThroughContentKeys} [FileKey {stats.FileKeyPassThroughKeys}, " &
             $"RegKey {stats.RegKeyPassThroughKeys}, ExcludeKey {stats.ExcludeKeyPassThroughKeys}]; " &
             $"detect {stats.DetectKeys + stats.DetectFileKeys + stats.DetectOSKeys}; " &
             $"inferred Root in {stats.InferredRootEntries} entries; {stats.VariableFanoutKeys} variable fan-out keys; " &
             $"abstraction: {stats.TotalVariablesDeclared} vars, {stats.NestedVariableRefs} nested, {stats.TangledEntries.Count} tangled " &
             $"(largest table {stats.MaxVariablesInEntry}{If(stats.MaxVariablesEntryName.Length > 0, $" in [{stats.MaxVariablesEntryName}]", "")})")

    End Sub

    ''' <summary>
    ''' Renders the abstraction-payoff section: how much variable machinery the run declared,
    ''' how much of it is interdependent, the largest single symbol table (an informational
    ''' flat-block backstop), and the watch list of entries crossing
    ''' <see cref="NestedRefWatchThreshold"/>. Watch-list entries show variable count,
    ''' nested-reference count, and resulting key count side by side, ranked by nesting, so a
    ''' tangled block (the reader pays a topological-sort tax) stands out from a merely large
    ''' one. The list is bounded so a tangled run does not flood the menu.
    ''' </summary>
    '''
    ''' <param name="stats">
    ''' The populated statistics accumulator for the completed run
    ''' </param>
    '''
    ''' <param name="menuOutput">
    ''' The <c> MenuSection </c> receiving the abstraction lines for display
    ''' </param>
    Private Sub renderAbstractionPayoff(stats As EntryBuilderStats, menuOutput As MenuSection)

        Const watchListCap As Integer = 10

        menuOutput.AddBlank()
        menuOutput.AddColoredLine("Abstraction payoff", ConsoleColor.Cyan)
        menuOutput.AddColoredLine(statLine("Variables declared:", stats.TotalVariablesDeclared), ConsoleColor.White)
        menuOutput.AddColoredLine(statLine("Nested variable refs:", stats.NestedVariableRefs), ConsoleColor.White)

        If stats.MaxVariablesInEntry > 0 Then

            menuOutput.AddColoredLine(String.Format("  {0,-28}{1,8:#,##0}   [{2}]", "Largest symbol table:", stats.MaxVariablesInEntry, stats.MaxVariablesEntryName), ConsoleColor.White)

        End If

        menuOutput.AddColoredLine(statLine($"Tangled abstractions (>={NestedRefWatchThreshold} nested):", stats.TangledEntries.Count),
                                  If(stats.TangledEntries.Count > 0, ConsoleColor.DarkYellow, ConsoleColor.DarkGray))

        Dim ranked = stats.TangledEntries.OrderByDescending(Function(t) t.NestedRefs).ThenByDescending(Function(t) t.VarCount).ToList()

        For Each te In ranked.Take(watchListCap)

            menuOutput.AddColoredLine($"    {te.Name}  ({te.NestedRefs} nested, {te.VarCount} vars -> {te.KeyCount} keys)", ConsoleColor.DarkYellow)

        Next

        If ranked.Count > watchListCap Then menuOutput.AddColoredLine($"    (+{ranked.Count - watchListCap} more)", ConsoleColor.DarkGray)

    End Sub

    ''' <summary>
    ''' Parses one source section into a populated <c> EntrySpec </c>. Issues warnings
    ''' and sets <c> ShouldSkip </c> for entries that are structurally invalid (no
    ''' category, no content). Warns on stray <c> Default= </c> declarations (this
    ''' builder never emits <c> Default </c>) and on missing detection.
    ''' </summary>
    '''
    ''' <param name="entrySection">
    ''' The source section describing one entry
    ''' </param>
    '''
    ''' <param name="menuOutput">
    ''' The <c> MenuSection </c> receiving warnings for display
    ''' </param>
    '''
    ''' <returns>
    ''' An <c> EntrySpec </c> populated from <paramref name="entrySection"/>.
    ''' Check <c> ShouldSkip </c> before using the result.
    ''' </returns>
    Private Function parseEntrySpec(entrySection As iniSection2,
                                     menuOutput As MenuSection) As EntrySpec

        Dim spec As New EntrySpec(entrySection.Name)

        For Each key In entrySection.Keys

            Select Case key.KeyType.ToUpperInvariant()

                Case "SECTION" : spec.SectionName = key.Value

                Case "LANGSECREF" : spec.LangSecRef = key.Value

                Case "WEBVIEWROOT" : spec.WebViewRoots.Add(key.Value)

                Case "QTWEBENGINEROOT" : spec.QtWebEngineRoots.Add(key.Value)

                Case "DETECTFILE" : spec.DetectFiles.Add(key.Value)

                Case "DETECT" : spec.Detects.Add(key.Value)

                Case "DETECTOS", "SPECIALDETECT"

                    Dim sdMsg = $"DetectOS and SpecialDetect are deprecated; key dropped from [{entrySection.Name}]. Replace with Detect or DetectFile"
                    gLog(sdMsg)
                    menuOutput.AddWarning(sdMsg)

                Case "WARNING" : spec.Warnings.Add(key.Value)

                Case "FILEKEYBASE" : spec.FileKeyBases.Add(key.Value)

                Case "FILEKEY" : spec.FileKeys.Add(key.Value)

                Case "REGKEYBASE" : spec.RegKeyBases.Add(key.Value)

                Case "REGKEY" : spec.RegKeys.Add(key.Value)

                Case "EXCLUDEKEYBASE" : spec.ExcludeKeyBases.Add(key.Value)

                Case "EXCLUDEKEY" : spec.ExcludeKeys.Add(key.Value)

                Case "WEBVIEWSCAFFOLDS"

                    spec.WebViewScaffoldNames = splitCsv(key.Value)
                    spec.WebViewScaffoldsKeyPresent = True

                Case "EXCLUDEWEBVIEWSCAFFOLDS" : spec.ExcludedWebViewScaffolds = splitCsv(key.Value)

                Case "QTWEBENGINESCAFFOLDS"

                    spec.QtWebEngineScaffoldNames = splitCsv(key.Value)
                    spec.QtWebEngineScaffoldsKeyPresent = True

                Case "EXCLUDEQTWEBENGINESCAFFOLDS" : spec.ExcludedQtWebEngineScaffolds = splitCsv(key.Value)

                Case "SKIP" : spec.ShouldSkip = True

                Case "DEFAULT"

                    Dim defaultMsg = $"Default= declared in [{entrySection.Name}]; EntryBuilder never emits Default, ignoring"
                    gLog(defaultMsg)
                    menuOutput.AddWarning(defaultMsg)

                Case Else

                    ' Open-vocabulary: any unrecognised key is a list-variable declaration
                    ' (e.g. Version=9.0,10.0,11.0). Two typo backstops complement this:
                    ' (a) undeclared-token warnings during expansion catch <Verison> for
                    ' a Version= declaration, and (b) the UnreferencedNames check after
                    ' generation catches Versoin= that nothing references.
                    '
                    ' Defensive collision check: if the declared name (after StripNums)
                    ' matches a reserved first-class key, warn. With the current Select
                    ' Case structure this is dead code — every reserved name has a Case
                    ' branch above — but the check keeps the ReservedKeys constant and
                    ' the parser's Case set explicitly cross-referenced, so a future
                    ' parser refactor that drops a branch surfaces the omission.
                    If ReservedKeys.Any(Function(r) String.Equals(r, key.KeyType, StringComparison.InvariantCultureIgnoreCase)) Then

                        Dim shadowMsg = $"Variable declaration '{key.KeyType}=' in [{entrySection.Name}] shadows reserved key name; possible typo"
                        gLog(shadowMsg)
                        menuOutput.AddWarning(shadowMsg)

                    End If

                    spec.Variables.Add(key.KeyType, key.Value)

            End Select

        Next

        ' Capture the symbol table's nesting depth before ResolveAll flattens it — this
        ' feeds the abstraction-payoff statistic and would read as 0 afterwards.
        spec.NestedVariableRefs = spec.Variables.NestedReferenceCount()

        ' Resolve nested <token> references inside the variables themselves
        ' (e.g. RegRoots=HKCU\...\<HKCUVersions>) so that later template-time
        ' expansion sees fully-literal value lists.
        For Each diag In VariableExpander.ResolveAll(spec.Variables)

            Dim contextualMsg = $"[{entrySection.Name}]: {diag.Message}"
            gLog(contextualMsg)
            If diag.Severity = DiagnosticSeverity.Warning Then menuOutput.AddWarning(contextualMsg)

        Next

        ' Reserved Root variable: synthesise Detect/DetectFile=<Root> from the declaration.
        ' Runs after ResolveAll (so nested roots are literal) and before the detection
        ' checks below (so a Root-only entry is not falsely reported as always-on).
        inferRootDetection(spec, menuOutput)

        Dim bothCategories = spec.LangSecRef.Length > 0 AndAlso spec.SectionName.Length > 0
        Dim noCategory = spec.LangSecRef.Length = 0 AndAlso spec.SectionName.Length = 0
        Dim noContent = spec.WebViewRoots.Count = 0 AndAlso spec.QtWebEngineRoots.Count = 0 AndAlso
                        spec.FileKeyBases.Count = 0 AndAlso spec.FileKeys.Count = 0 AndAlso
                        spec.RegKeyBases.Count = 0 AndAlso spec.RegKeys.Count = 0
        Dim noDetection = spec.DetectFiles.Count = 0 AndAlso spec.Detects.Count = 0 AndAlso spec.DetectOS.Length = 0

        If bothCategories Then

            Dim catMsg = $"Both Section and LangSecRef present in [{entrySection.Name}], using LangSecRef"
            gLog(catMsg)
            menuOutput.AddWarning(catMsg)
            spec.SectionName = ""

        End If

        If noCategory Then

            Dim noCatMsg = $"No Section or LangSecRef in [{entrySection.Name}], skipping"
            gLog(noCatMsg)
            menuOutput.AddWarning(noCatMsg)
            spec.ShouldSkip = True

        End If

        If noContent Then

            Dim noContentMsg = $"[{entrySection.Name}] declares no WebViewRoot, QtWebEngineRoot, FileKeyBase, or RegKeyBase — nothing to emit, skipping"
            gLog(noContentMsg)
            menuOutput.AddWarning(noContentMsg)
            spec.ShouldSkip = True

        End If

        If noDetection AndAlso Not spec.ShouldSkip Then

            Dim noDetectMsg = $"[{entrySection.Name}] declares no detection (Detect / DetectFile / DetectOS); generated entry will be always-on"
            gLog(noDetectMsg)
            menuOutput.AddWarning(noDetectMsg)

        End If

        Return spec

    End Function

    ''' <summary>
    ''' Synthesises detection from a reserved <c> Root </c> variable. When an entry declares
    ''' <c> Root= </c>, EntryBuilder guarantees a <c> Detect=&lt;Root&gt; </c> (registry-domain
    ''' root) or <c> DetectFile=&lt;Root&gt; </c> (filesystem-domain root) is present, classifying
    ''' by the leading path segment of the resolved value against the known registry hives via
    ''' <see cref="regKeyParams2.HasValidRoot"/>. The synthesised key is the literal
    ''' <c> &lt;Root&gt; </c> token rather than the expanded value, so the normal phase-2 pass in
    ''' <see cref="generateEntry"/> fans it out over multi-valued roots and marks <c> Root </c>
    ''' referenced for the unreferenced-variable typo backstop. <br /><br />
    '''
    ''' Injection is idempotent and additive: if the target list already holds a
    ''' <c> &lt;Root&gt; </c> token (a hand-written <c> Detect=&lt;Root&gt; </c>) nothing is added,
    ''' and any other detection criteria are left untouched. The token is prepended so the root
    ''' anchor is the first detection key of its kind. <br /><br />
    '''
    ''' Authors who want a base-path variable WITHOUT automatic detection must name it something
    ''' other than <c> Root </c>.
    ''' </summary>
    '''
    ''' <param name="spec">
    ''' The entry being parsed; its <see cref="EntrySpec.Detects"/> /
    ''' <see cref="EntrySpec.DetectFiles"/> are mutated in place when a <c> Root </c> variable
    ''' is present
    ''' </param>
    '''
    ''' <param name="menuOutput">
    ''' The <c> MenuSection </c> receiving warnings for display
    ''' </param>
    Private Sub inferRootDetection(ByRef spec As EntrySpec, menuOutput As MenuSection)

        If Not spec.Variables.IsDeclared("Root") Then Return

        Dim rootValues = spec.Variables.Values("Root")
        If rootValues.Count = 0 Then Return

        ' A value still carrying a <token> means an unresolved nested reference (typically a
        ' cycle caught by ResolveAll); domain classification on the bracket text is meaningless,
        ' so skip inference and warn rather than emit a bogus detection key.
        If rootValues.Any(Function(v) v.Contains("<")) Then

            Dim unresolvedMsg = $"[{spec.Name}] Root has unresolved <token> references; skipping detection inference"
            gLog(unresolvedMsg)
            menuOutput.AddWarning(unresolvedMsg)
            Return

        End If

        ' Classify each resolved value's domain by its leading path segment. A homogeneous list
        ' routes to one Detect or DetectFile; a heterogeneous list is an authoring error — warn
        ' and classify by the first value (its <Root> token still fans out, surfacing the
        ' domain mismatch downstream in WinappDebug).
        Dim firstIsRegistry = New regKeyParams2(rootValues(0)).HasValidRoot

        If rootValues.Any(Function(v) New regKeyParams2(v).HasValidRoot <> firstIsRegistry) Then

            Dim mixedMsg = $"[{spec.Name}] Root mixes registry and filesystem values; inferring detection from the first value's domain"
            gLog(mixedMsg)
            menuOutput.AddWarning(mixedMsg)

        End If

        Dim target = If(firstIsRegistry, spec.Detects, spec.DetectFiles)

        ' Idempotent: never duplicate a hand-written Detect/DetectFile=<Root>.
        If target.Any(Function(d) String.Equals(d.Trim(), "<Root>", StringComparison.InvariantCultureIgnoreCase)) Then Return

        target.Insert(0, "<Root>")
        spec.RootDetectionInferred = True

        gLog($"[{spec.Name}] inferred {If(firstIsRegistry, "Detect", "DetectFile")}=<Root>")

    End Sub

    ''' <summary>
    ''' Splits a comma-separated value into a trimmed, non-empty list of tokens.
    ''' Used by <c> WebViewScaffolds= </c> and <c> ExcludeWebViewScaffolds= </c>. An empty input
    ''' yields an empty list.
    ''' </summary>
    '''
    ''' <param name="value">
    ''' The raw key value to split on commas
    ''' </param>
    '''
    ''' <returns>
    ''' Each non-empty, trimmed token from <paramref name="value"/>, in original order
    ''' </returns>
    Private Function splitCsv(value As String) As List(Of String)

        Return value.Split(","c) _
                    .Select(Function(s) s.Trim()) _
                    .Where(Function(s) s.Length > 0) _
                    .ToList()

    End Function

    ''' <summary>
    ''' Selects the set of WebView scaffold names that should be emitted for the given
    ''' entry. Resolution order:
    ''' <list type="bullet">
    ''' <item>
    ''' If the entry declared <c> WebViewScaffolds= </c> (regardless of value), use that
    ''' explicit list. A present-but-empty value yields no scaffolds.
    ''' </item>
    ''' <item>
    ''' If the explicit list contains the sentinel <c> All </c> (case-insensitive), it is
    ''' replaced with every scaffold name in <paramref name="available"/>; any other
    ''' names listed alongside <c> All </c> are redundant and emit a warning.
    ''' </item>
    ''' <item>
    ''' Otherwise, use <see cref="DefaultScaffolds"/>.
    ''' </item>
    ''' </list>
    ''' Any names in <c> ExcludedWebViewScaffolds </c> are then subtracted, and unknown names
    ''' (not present in <paramref name="available"/>) are dropped with a warning.
    ''' Combining <c> WebViewScaffolds= </c> and <c> ExcludeWebViewScaffolds= </c> warns when the
    ''' explicit list is anything other than <c> All </c>.
    ''' </summary>
    '''
    ''' <param name="spec">
    ''' The entry whose scaffold selection is being resolved
    ''' </param>
    '''
    ''' <param name="available">
    ''' The catalog of known scaffolds, keyed by name, loaded from
    ''' <c> Assembler\Scaffolds\webview.ini </c>
    ''' </param>
    '''
    ''' <param name="menuOutput">
    ''' The <c> MenuSection </c> receiving warnings for display
    ''' </param>
    '''
    ''' <returns>
    ''' The ordered list of scaffold names to emit for this entry, with unknowns removed
    ''' </returns>
    Private Function resolveWebViewScaffolds(spec As EntrySpec,
                                       available As Dictionary(Of String, List(Of String)),
                                       menuOutput As MenuSection) As List(Of String)

        Return WebViewScaffolds.ResolveScaffolds(spec.WebViewScaffoldNames, spec.WebViewScaffoldsKeyPresent,
                                                 spec.ExcludedWebViewScaffolds, available,
                                                 DefaultScaffolds, "WebView", spec.Name, menuOutput)

    End Function

    ''' <summary>
    ''' Selects the set of QtWebEngine scaffold names to emit for the given entry, following
    ''' the same resolution grammar as <see cref="resolveWebViewScaffolds"/> but driven by the
    ''' entry's <c> QtWebEngineScaffolds= </c> / <c> ExcludeQtWebEngineScaffolds= </c> keys and
    ''' the QtWebEngine default set. Delegates to <see cref="WebViewScaffolds.ResolveScaffolds"/>.
    ''' </summary>
    '''
    ''' <param name="spec">
    ''' The entry whose QtWebEngine scaffold selection is being resolved
    ''' </param>
    '''
    ''' <param name="available">
    ''' The catalog of known QtWebEngine scaffolds, keyed by name, loaded from
    ''' <c> Assembler\Scaffolds\qtwebengine.ini </c>
    ''' </param>
    '''
    ''' <param name="menuOutput">
    ''' The <c> MenuSection </c> receiving warnings for display
    ''' </param>
    '''
    ''' <returns>
    ''' The ordered list of QtWebEngine scaffold names to emit for this entry, with unknowns removed
    ''' </returns>
    Private Function resolveQtWebEngineScaffolds(spec As EntrySpec,
                                                 available As Dictionary(Of String, List(Of String)),
                                                 menuOutput As MenuSection) As List(Of String)

        Return WebViewScaffolds.ResolveScaffolds(spec.QtWebEngineScaffoldNames, spec.QtWebEngineScaffoldsKeyPresent,
                                                 spec.ExcludedQtWebEngineScaffolds, available,
                                                 WebViewScaffolds.QtWebEngineDefaultScaffolds, "QtWebEngine", spec.Name, menuOutput)

    End Function

    ''' <summary>
    ''' Generates one winapp2.ini section for the given entry. Emits keys in winapp2
    ''' canonical order: category, Detect, DetectFile, DetectOS, Warning, FileKey
    ''' (scaffold expansion followed by entry-level FileKeyBase), RegKey, ExcludeKey.
    ''' Never emits <c> Default= </c>.
    ''' </summary>
    '''
    ''' <param name="spec">
    ''' The parsed entry definition to generate output for
    ''' </param>
    '''
    ''' <param name="catalog">
    ''' The shared WebView scaffold catalog
    ''' </param>
    '''
    ''' <param name="stats">
    ''' The run-wide statistics accumulator; each key emitted here is tallied into the
    ''' relevant category counter
    ''' </param>
    '''
    ''' <param name="menuOutput">
    ''' The <c> MenuSection </c> receiving progress lines and warnings for display
    ''' </param>
    '''
    ''' <returns>
    ''' A fully populated <c> iniSection2 </c> ready to be added to the output file
    ''' </returns>
    Private Function generateEntry(spec As EntrySpec,
                                    catalog As Dictionary(Of String, List(Of String)),
                                    qtCatalog As Dictionary(Of String, List(Of String)),
                                    stats As EntryBuilderStats,
                                    menuOutput As MenuSection) As iniSection2

        Dim generatingMsg = $"Generating entry: {spec.Name}"
        menuOutput.AddColoredLine(generatingMsg, ConsoleColor.Magenta)
        gLog($"  {generatingMsg}")

        Dim section As New iniSection2(spec.Name)

        ' 1. Category
        Select Case True

            Case spec.LangSecRef.Length > 0 : section.AddKey(New iniKey2($"LangSecRef={spec.LangSecRef}")) : stats.CategoryKeys += 1

            Case spec.SectionName.Length > 0 : section.AddKey(New iniKey2($"Section={spec.SectionName}")) : stats.CategoryKeys += 1

            Case Else : gLog($"{spec.Name} has no category key")

        End Select

        ' 2. Detect keys (phase-2 expansion, Registry domain)
        Dim detectValues As New List(Of String)
        For Each d In spec.Detects : detectValues.AddRange(expandPhase2(d, spec, ExpansionDomain.Registry, "Detect", stats, menuOutput)) : Next

        stats.DetectKeys += detectValues.Count

        If detectValues.Count = 1 Then

            section.AddKey(New iniKey2($"Detect={detectValues(0)}"))

        Else

            For i = 0 To detectValues.Count - 1

                section.AddKey(New iniKey2($"Detect{i + 1}={detectValues(i)}"))

            Next

        End If

        ' 3. DetectFile keys (phase-2 expansion, Filesystem domain — no %WebViewRoot% by convention)
        Dim detectFileValues As New List(Of String)
        For Each df In spec.DetectFiles : detectFileValues.AddRange(expandPhase2(df, spec, ExpansionDomain.Filesystem, "DetectFile", stats, menuOutput)) : Next

        stats.DetectFileKeys += detectFileValues.Count

        If detectFileValues.Count = 1 Then

            section.AddKey(New iniKey2($"DetectFile={detectFileValues(0)}"))

        Else

            For i = 0 To detectFileValues.Count - 1

                section.AddKey(New iniKey2($"DetectFile{i + 1}={detectFileValues(i)}"))

            Next

        End If

        ' 4. DetectOS
        If spec.DetectOS.Length > 0 Then section.AddKey(New iniKey2($"DetectOS={spec.DetectOS}")) : stats.DetectOSKeys += 1

        ' 5. Warnings (pass-through, verbatim — prose is never variable-expanded)
        For Each w In spec.Warnings

            section.AddKey(New iniKey2($"Warning={w}"))
            stats.WarningKeys += 1

        Next

        ' 6. FileKeys, emitted in four provenance tiers: WebView scaffold expansion (generated),
        ' QtWebEngine scaffold expansion (generated), then FileKeyBase= templates (generated),
        ' then literal FileKey= (pass-through). Each tier is phase-2 expanded against the entry's
        ' variables and counted separately. Order does not survive remotedebug's per-bucket
        ' alphabetization, so the tiers are purely for accurate provenance accounting.
        Dim fileKeys As New List(Of String)

        If spec.WebViewRoots.Count > 0 Then

            Dim selectedScaffolds = resolveWebViewScaffolds(spec, catalog, menuOutput)
            fileKeys.AddRange(expandScaffoldFamily(spec.WebViewRoots, selectedScaffolds, catalog, "%WebViewRoot%", spec, stats, menuOutput))

        End If

        ' Tier boundaries: anything before each marker came from the prior provenance.
        stats.ScaffoldFileKeys += fileKeys.Count
        Dim afterScaffold = fileKeys.Count

        If spec.QtWebEngineRoots.Count > 0 Then

            Dim selectedQtScaffolds = resolveQtWebEngineScaffolds(spec, qtCatalog, menuOutput)
            fileKeys.AddRange(expandScaffoldFamily(spec.QtWebEngineRoots, selectedQtScaffolds, qtCatalog, "%QtWebEngineRoot%", spec, stats, menuOutput))

        End If

        stats.QtScaffoldFileKeys += fileKeys.Count - afterScaffold
        Dim afterQtScaffold = fileKeys.Count

        For Each fkb In spec.FileKeyBases : fileKeys.AddRange(expandFully(fkb, spec, ExpansionDomain.Filesystem, "FileKey", stats, menuOutput)) : Next

        stats.FileKeyBaseKeys += fileKeys.Count - afterQtScaffold
        Dim afterBase = fileKeys.Count

        For Each fk In spec.FileKeys : fileKeys.AddRange(expandFully(fk, spec, ExpansionDomain.Filesystem, "FileKey", stats, menuOutput)) : Next

        stats.FileKeyPassThroughKeys += fileKeys.Count - afterBase

        For i = 0 To fileKeys.Count - 1

            section.AddKey(New iniKey2($"FileKey{i + 1}={fileKeys(i)}"))

        Next

        ' 7. RegKeys with %WebViewRoot% / %QtWebEngineRoot% + variable expansion (Registry domain),
        ' renumbered from 1: RegKeyBase= templates (generated) first, then literal RegKey= (pass-through).
        Dim regKeyValues As New List(Of String)
        For Each rkb In spec.RegKeyBases : regKeyValues.AddRange(expandFully(rkb, spec, ExpansionDomain.Registry, "RegKey", stats, menuOutput)) : Next

        stats.RegKeyBaseKeys += regKeyValues.Count
        Dim afterRegBase = regKeyValues.Count

        For Each rk In spec.RegKeys : regKeyValues.AddRange(expandFully(rk, spec, ExpansionDomain.Registry, "RegKey", stats, menuOutput)) : Next

        stats.RegKeyPassThroughKeys += regKeyValues.Count - afterRegBase

        For i = 0 To regKeyValues.Count - 1

            section.AddKey(New iniKey2($"RegKey{i + 1}={regKeyValues(i)}"))

        Next

        ' 8. ExcludeKeys with %WebViewRoot% / %QtWebEngineRoot% expansion, renumbered from 1 after fan-out:
        ' ExcludeKeyBase= templates (generated) first, then literal ExcludeKey= (pass-through).
        Dim excludeKeyValues As New List(Of String)

        Dim exclBaseValues = expandExcludeKeys(spec.ExcludeKeyBases, spec, stats, menuOutput)
        stats.ExcludeKeyBaseKeys += exclBaseValues.Count
        excludeKeyValues.AddRange(exclBaseValues)

        Dim exclPassValues = expandExcludeKeys(spec.ExcludeKeys, spec, stats, menuOutput)
        stats.ExcludeKeyPassThroughKeys += exclPassValues.Count
        excludeKeyValues.AddRange(exclPassValues)

        For i = 0 To excludeKeyValues.Count - 1

            section.AddKey(New iniKey2($"ExcludeKey{i + 1}={excludeKeyValues(i)}"))

        Next

        ' Default= is intentionally never emitted

        Return section

    End Function

    ''' <summary>
    ''' Expands a list of ExcludeKey templates against the entry's roots and variables,
    ''' classifying each phase-1 result's variable-expansion domain by its parsed
    ''' <c> excludeKeyParams2.Flag </c> (REG → Registry; FILE, PATH, Unknown → Filesystem).
    ''' Shared by the <c> ExcludeKeyBase= </c> (generated) and <c> ExcludeKey= </c>
    ''' (pass-through) tiers so both apply identical domain classification while being
    ''' counted separately by the caller.
    ''' </summary>
    '''
    ''' <param name="templates">
    ''' The ExcludeKey template values to expand (one provenance tier at a time)
    ''' </param>
    '''
    ''' <param name="spec">
    ''' The entry whose roots and variables drive expansion
    ''' </param>
    '''
    ''' <param name="stats">
    ''' The run-wide statistics accumulator, forwarded to <see cref="expandPhase2"/> for
    ''' variable fan-out accounting
    ''' </param>
    '''
    ''' <param name="menuOutput">
    ''' The <c> MenuSection </c> receiving Warning-severity diagnostics
    ''' </param>
    '''
    ''' <returns>
    ''' Every expanded ExcludeKey value produced from <paramref name="templates"/>, in
    ''' fan-out order
    ''' </returns>
    Private Function expandExcludeKeys(templates As List(Of String),
                                        spec As EntrySpec,
                                        stats As EntryBuilderStats,
                                        menuOutput As MenuSection) As List(Of String)

        Dim values As New List(Of String)

        For Each exclKey In templates

            For Each rootExpanded In expandRoot(exclKey, spec)

                Dim parsed As New excludeKeyParams2(rootExpanded)
                Dim exclDomain As ExpansionDomain = If(parsed.Flag = excludeKeyFlag.Reg,
                                                       ExpansionDomain.Registry,
                                                       ExpansionDomain.Filesystem)

                values.AddRange(expandPhase2(rootExpanded, spec, exclDomain, "ExcludeKey", stats, menuOutput))

            Next

        Next

        Return values

    End Function

    ''' <summary>
    ''' Runs the phase-2 variable-expansion pass on <paramref name="template"/>,
    ''' routes any returned diagnostics onto <c> gLog </c> and
    ''' <paramref name="menuOutput"/>, and returns the expanded value set. The caller
    ''' is responsible for any phase-1 root-placeholder substitution beforehand.
    ''' </summary>
    '''
    ''' <param name="template">
    ''' One key value (already phase-1 expanded), possibly containing <c> &lt;var&gt; </c>
    ''' tokens
    ''' </param>
    '''
    ''' <param name="spec">
    ''' The entry whose <see cref="EntrySpec.Variables"/> set drives the expansion;
    ''' tokens observed update reference tracking for the post-generation typo backstop
    ''' </param>
    '''
    ''' <param name="domain">
    ''' The domain of the emitting key, determining undeclared-token behavior
    ''' </param>
    '''
    ''' <param name="keyLabel">
    ''' Short tag (e.g. <c> Detect </c>, <c> FileKey </c>) embedded into diagnostic
    ''' messages for source-of-warning localisation
    ''' </param>
    '''
    ''' <param name="stats">
    ''' The run-wide statistics accumulator; the values produced beyond the first are
    ''' tallied as <see cref="EntryBuilderStats.VariableFanoutKeys"/>
    ''' </param>
    '''
    ''' <param name="menuOutput">
    ''' The <c> MenuSection </c> receiving Warning-severity diagnostics for display
    ''' </param>
    '''
    ''' <returns>
    ''' The fan-out values; empty when the template was dropped
    ''' </returns>
    Private Function expandPhase2(template As String,
                                   spec As EntrySpec,
                                   domain As ExpansionDomain,
                                   keyLabel As String,
                                   stats As EntryBuilderStats,
                                   menuOutput As MenuSection) As List(Of String)

        Dim result = VariableExpander.Expand(template, spec.Variables, domain, $"[{spec.Name}].{keyLabel}")

        For Each diag In result.Diagnostics

            gLog(diag.Message)
            If diag.Severity = DiagnosticSeverity.Warning Then menuOutput.AddWarning(diag.Message)

        Next

        ' A clean single-value expansion (or a dropped key) adds nothing; only the
        ' second and subsequent values represent variable-driven fan-out.
        stats.VariableFanoutKeys += Math.Max(0, result.Values.Count - 1)

        Return result.Values

    End Function

    ''' <summary>
    ''' Runs both expansion phases on <paramref name="template"/>: phase-1 root-placeholder
    ''' substitution (<c> %WebViewRoot% </c> / <c> %QtWebEngineRoot% </c>) via
    ''' <see cref="expandRoot"/>, then phase-2 variable expansion via
    ''' <see cref="expandPhase2"/> on each phase-1 result. Diagnostics are flushed to
    ''' <c> gLog </c> / <paramref name="menuOutput"/> as they are produced.
    ''' </summary>
    '''
    ''' <param name="template">
    ''' One key value template, possibly containing <c> %WebViewRoot% </c> and
    ''' <c> &lt;var&gt; </c> tokens
    ''' </param>
    '''
    ''' <param name="spec">
    ''' The entry whose roots and variables drive expansion
    ''' </param>
    '''
    ''' <param name="domain">
    ''' Fixed domain for variable expansion (used for every phase-1 result). For
    ''' ExcludeKey, the domain is classified per phase-1 result by the caller and
    ''' this helper is not used.
    ''' </param>
    '''
    ''' <param name="keyLabel">
    ''' Short tag embedded into diagnostic messages
    ''' </param>
    '''
    ''' <param name="stats">
    ''' The run-wide statistics accumulator, forwarded to <see cref="expandPhase2"/>
    ''' </param>
    '''
    ''' <param name="menuOutput">
    ''' The <c> MenuSection </c> receiving Warning-severity diagnostics
    ''' </param>
    '''
    ''' <returns>
    ''' Every string produced by phase-1 × phase-2 expansion, in fan-out order
    ''' </returns>
    Private Function expandFully(template As String,
                                  spec As EntrySpec,
                                  domain As ExpansionDomain,
                                  keyLabel As String,
                                  stats As EntryBuilderStats,
                                  menuOutput As MenuSection) As List(Of String)

        Dim values As New List(Of String)

        For Each rootExpanded In expandRoot(template, spec)

            values.AddRange(expandPhase2(rootExpanded, spec, domain, keyLabel, stats, menuOutput))

        Next

        Return values

    End Function

    ''' <summary>
    ''' Expands a template's root placeholders against the entry's declared roots, composing
    ''' the two scaffold families: <c> %WebViewRoot% </c> fans out against
    ''' <see cref="EntrySpec.WebViewRoots"/> and <c> %QtWebEngineRoot% </c> against
    ''' <see cref="EntrySpec.QtWebEngineRoots"/>. A template containing neither placeholder is
    ''' returned verbatim as a single-element list (no duplication). The two fan-outs compose,
    ''' so a template referencing both placeholders multiplies across both root lists.
    ''' </summary>
    '''
    ''' <param name="template">
    ''' A key value string that may contain <c> %WebViewRoot% </c> and/or
    ''' <c> %QtWebEngineRoot% </c>
    ''' </param>
    '''
    ''' <param name="spec">
    ''' The entry whose declared root lists drive the placeholder substitution
    ''' </param>
    '''
    ''' <returns>
    ''' The fan-out of <paramref name="template"/> across whichever placeholders it contains,
    ''' or a single-element list containing it verbatim when it contains neither
    ''' </returns>
    Private Function expandRoot(template As String,
                                 spec As EntrySpec) As List(Of String)

        Dim result As New List(Of String) From {template}
        result = fanOutPlaceholder(result, "%WebViewRoot%", spec.WebViewRoots)
        result = fanOutPlaceholder(result, "%QtWebEngineRoot%", spec.QtWebEngineRoots)
        Return result

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
    Private Function fanOutPlaceholder(templates As List(Of String),
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

    ''' <summary>
    ''' Expands one scaffold family for an entry: for each declared root, each selected
    ''' scaffold, and each <c> FileKeyBase= </c> template in that scaffold, substitutes
    ''' <paramref name="placeholder"/> with the root and runs phase-2 variable expansion.
    ''' Shared by the WebView and QtWebEngine scaffold tiers, which differ only in their
    ''' root list, catalog, and placeholder.
    ''' </summary>
    '''
    ''' <param name="roots">
    ''' The entry's declared root paths for this family
    ''' </param>
    '''
    ''' <param name="selectedScaffolds">
    ''' The resolved scaffold names to emit (already filtered to catalog members)
    ''' </param>
    '''
    ''' <param name="catalog">
    ''' The scaffold catalog for this family, keyed by scaffold name
    ''' </param>
    '''
    ''' <param name="placeholder">
    ''' The family's root placeholder (<c> %WebViewRoot% </c> or <c> %QtWebEngineRoot% </c>)
    ''' </param>
    '''
    ''' <param name="spec">
    ''' The entry being generated, supplying the variable set for phase-2 expansion
    ''' </param>
    '''
    ''' <param name="stats">
    ''' The run-wide statistics accumulator, forwarded to <see cref="expandPhase2"/>
    ''' </param>
    '''
    ''' <param name="menuOutput">
    ''' The <c> MenuSection </c> receiving Warning-severity diagnostics
    ''' </param>
    '''
    ''' <returns>
    ''' Every expanded FileKey value produced by this family, in fan-out order
    ''' </returns>
    Private Function expandScaffoldFamily(roots As List(Of String),
                                           selectedScaffolds As List(Of String),
                                           catalog As Dictionary(Of String, List(Of String)),
                                           placeholder As String,
                                           spec As EntrySpec,
                                           stats As EntryBuilderStats,
                                           menuOutput As MenuSection) As List(Of String)

        Dim result As New List(Of String)

        For Each root In roots

            For Each scaffoldName In selectedScaffolds

                For Each template In catalog(scaffoldName)

                    Dim rootExpanded = template.Replace(placeholder, root)
                    result.AddRange(expandPhase2(rootExpanded, spec, ExpansionDomain.Filesystem, "FileKey", stats, menuOutput))

                Next

            Next

        Next

        Return result

    End Function

    ''' <summary>
    ''' Classifies an entry name into its per-letter artifact bucket: the uppercased
    ''' first character when it is an ASCII letter, otherwise <c> # </c> (digits,
    ''' punctuation, and anything non-alphabetic). Bucketing is by entry name rather
    ''' than by source file so an entry filed in the wrong source letter still lands
    ''' in its canonical artifact.
    ''' </summary>
    '''
    ''' <param name="entryName">
    ''' The entry name (section header text) to classify
    ''' </param>
    '''
    ''' <returns>
    ''' The single-character bucket name: <c> A </c>–<c> Z </c> or <c> # </c>
    ''' </returns>
    Public Function LetterBucketFor(entryName As String) As String

        If String.IsNullOrEmpty(entryName) Then Return "#"

        Dim c = Char.ToUpperInvariant(entryName(0))

        Return If(c >= "A"c AndAlso c <= "Z"c, c.ToString(), "#")

    End Function

    ''' <summary>
    ''' Writes the normalized output as 27 per-letter artifact files (<c> #.ini </c>,
    ''' <c> A.ini </c> ... <c> Z.ini </c>) in the save target's directory, each with a
    ''' deterministic do-not-edit header. All 27 files are always written — a letter
    ''' with no entries produces a header-only file — so the artifact set is stable
    ''' across runs and deletions of a letter's last entry cannot leave a stale file
    ''' behind.
    ''' </summary>
    '''
    ''' <param name="outputFile">
    ''' The fully generated, normalized output whose sections will be bucketed
    ''' </param>
    '''
    ''' <param name="menuOutput">
    ''' The <c> MenuSection </c> receiving the per-run summary line
    ''' </param>
    Private Sub writeSplitOutput(outputFile As iniFile2,
                                 menuOutput As MenuSection)

        Dim letters As New List(Of String) From {"#"}
        For c = AscW("A"c) To AscW("Z"c) : letters.Add(ChrW(c)) : Next

        Dim buckets = letters.ToDictionary(Function(l) l, Function(l) iniFile2.Empty(EntryBuilderFile2.Dir, l & ".ini"))

        For Each section In outputFile

            buckets(LetterBucketFor(section.Name)).AddSection(section)

        Next

        For Each letter In letters

            Dim bucket = buckets(letter)
            Dim letterDesc = If(letter = "#", "symbols or numbers", $"the letter {letter}")

            Dim sb As New StringBuilder()
            sb.AppendLine($"; # of entries: {bucket.Count:#,##0}")
            sb.AppendLine($"; {letter}.ini is generated by the Winapp2ool Entry Builder. DO NOT EDIT THIS FILE DIRECTLY")
            sb.AppendLine("; Changes made here will be overwritten by the next build; these entries' sources live in Assembler\EntryBuilder")
            sb.AppendLine("; Entries in this file may be incomplete and are not intended to be used directly with any cleaning software")
            sb.AppendLine("; They are utilized by winapp2ool to create the final winapp2.ini file for distribution")
            sb.AppendLine("; If you are not maintaining winapp2.ini for distribution, you probably don't need this file!")
            sb.AppendLine("; Refer to the Winapp2ool documentation for more information: " & readMeUrl)
            sb.AppendLine("; You can find the complete winapp2.ini file here: " & baseFlavorLink)
            sb.AppendLine(";")
            sb.AppendLine($"; This file contains winapp2.ini entries beginning with {letterDesc}")
            sb.AppendLine()
            sb.Append(bucket.ToString())

            bucket.OverwriteToFile(sb.ToString())

        Next

        Dim splitMsg = $"Wrote {letters.Count} per-letter artifact files to {EntryBuilderFile2.Dir}"
        menuOutput.AddColoredLine(splitMsg, ConsoleColor.Green)
        gLog(splitMsg)

    End Sub

    ''' <summary>
    ''' Combines all <c> *.ini </c> files in <paramref name="sourceDir"/> into a single
    ''' in-memory <c> iniFile2 </c>, processed in alphabetical order. Sections with
    ''' duplicate names across files are silently ignored (first-file-wins). Returns
    ''' an empty file if the directory does not exist or contains no parseable sections.
    ''' </summary>
    '''
    ''' <param name="sourceDir">
    ''' Path to the directory containing per-letter <c> *.ini </c> source files
    ''' </param>
    '''
    ''' <returns>
    ''' The merged <c> iniFile2 </c>, or an empty one if the directory is missing or empty
    ''' </returns>
    Private Function combineSourceDir(sourceDir As String) As iniFile2

        Dim combined = iniFile2.Empty("", "")

        If Not Directory.Exists(sourceDir) Then Return combined

        Dim files = Directory.GetFiles(sourceDir, "*.ini", SearchOption.TopDirectoryOnly).ToList()
        files.Sort()

        For Each filePath In files

            Dim f = iniFile2.FromFile(filePath)
            For Each section In f : combined.AddSection(section) : Next

        Next

        Return combined

    End Function

End Module
