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
''' UWPBuilder is a winapp2ool module which handles the generation of winapp2.ini
''' entries for Universal Windows Platform applications installed under
''' <c> %LocalAppData%\Packages </c>. <br /><br />
''' 
''' It also allows for hybrid win32+UWP entries for applications that have both a
''' UWP package and a win32 installation, by supporting importing arbitrary FileKey,
''' DetectFile, Detect, ExcludeKey, and RegKey values from the AppInfo definitions.  
''' 
''' <br /><br />
''' 
''' It reads from a single source directory containing:
''' <list type="bullet">
'''
''' <item>
''' <c> UWP.ini </c> <br /> The scaffold template with one or more
''' <c> [EntryScaffold: ...] </c> sections. Each scaffold's <c> DetectFileBase= </c>
''' lines are expanded per package to produce the entry's detection paths, and its
''' <c> FileKeyBase= </c> lines are applied to every application that doesn't set
''' <c> SkipUWPFileKeys= </c>, as a consistent baseline set of cleaning targets.
''' </item>
'''
''' <item>
''' An <c> AppInfo\ </c> subdirectory of per-letter <c> *.ini </c> files. <br />
''' Each section describes one UWP application with its <c> Package= </c>
''' identifier, category, optional <c> Detect= </c>, <c> DetectFile= </c>,
''' and <c> RegKey= </c> keys, and any app-specific <c> FileKeyBase= </c>,
''' <c> FileKey= </c>, and <c> ExcludeKey= </c> lines. These files are
''' combined in-memory at runtime; no intermediate combined file is required.
''' </item>
'''
''' </list>
'''
''' The <c> %Package% </c> DSL variable expands to
''' <c> %LocalAppData%\Packages\&lt;PACKAGE_FOLDER&gt; </c>, once per package for an app
''' with several. <c> %Package1% </c> / <c> %Package2% </c> in a key value select one
''' package by its position among the entry's <c> Package= </c> keys. The digit in a source
''' key name (<c> Package2= </c>) is ignored, so only declaration order counts.
'''
''' <br /><br />
'''
''' An entry may also draw scaffold FileKeys from the shared catalogs in
''' <c> Assembler\Scaffolds </c> by declaring a root for any family in
''' <see cref="ScaffoldCatalogs.ScaffoldFamilies"/>: <c> WebViewRoot= </c>,
''' <c> QtWebEngineRoot= </c> / <c> QtWebEngineCacheRoot= </c>, or <c> ElectronRoot= </c> /
''' <c> ElectronUpdaterRoot= </c>. Each <c> ...Root= </c> key also accepts a
''' <c> ...Path= </c> spelling.
''' Electron matters here despite MSIX packages rarely bundling it, because the
''' <b> hybrid win32+UWP </b> entries above carry the desktop install's paths too, and a
''' packaged app's desktop build is frequently Electron.
''' </summary>
Public Module UWPBuilder

    ''' <summary>
    ''' The AppInfo key types claimed by <see cref="parseAppInfo"/>. The parser's
    ''' <c> Select Case </c> doesn't read this list: it has its own branch for each name, and
    ''' any key type without a branch becomes a variable declaration. A name listed here but
    ''' missing a branch warns as a shadowed reserved key. For the reverse (a branch with no
    ''' list entry), only the scaffold family keys are guarded, by the test below.
    ''' Mirrors EntryBuilder's list of the same purpose, but carries UWPBuilder's own
    ''' vocabulary: <c> PACKAGE </c>, <c> SKIPUWPFILEKEYS </c> and the <c> ...PATH </c>
    ''' root spellings have no EntryBuilder counterpart.
    ''' <br /><br />
    '''
    ''' <c> Friend </c> rather than <c> Private </c> so the
    ''' <c> ScaffoldFamilies_KeyVocabularyPresentInBothBuilders </c> test can check that this
    ''' list holds each family's <c> {Family}Root= </c>, <c> {Family}Scaffolds= </c> and
    ''' <c> Exclude{Family}Scaffolds= </c> names. That test reads only this list, not the
    ''' parser's branches.
    ''' </summary>
    Friend ReadOnly UWPReservedKeys As String() = {
        "PACKAGE", "LANGSECREF", "SECTION",
        "DETECT", "DETECTFILE", "DETECTOS", "SPECIALDETECT",
        "FILEKEY", "FILEKEYBASE",
        "REGKEY", "REGKEYBASE",
        "EXCLUDEKEY", "EXCLUDEKEYBASE",
        "WEBVIEWPATH", "WEBVIEWROOT",
        "QTWEBENGINEPATH", "QTWEBENGINEROOT",
        "QTWEBENGINECACHEPATH", "QTWEBENGINECACHEROOT",
        "ELECTRONPATH", "ELECTRONROOT",
        "ELECTRONUPDATERPATH", "ELECTRONUPDATERROOT",
        "WEBVIEWSCAFFOLDS", "EXCLUDEWEBVIEWSCAFFOLDS",
        "QTWEBENGINESCAFFOLDS", "EXCLUDEQTWEBENGINESCAFFOLDS",
        "ELECTRONSCAFFOLDS", "EXCLUDEELECTRONSCAFFOLDS",
        "SKIP", "SKIPUWPFILEKEYS", "DEFAULT", "WARNING"
    }

    ''' <summary>
    ''' Stores the parsed information for a single UWP application entry
    ''' </summary>
    Friend Structure UWPAppInfo

        ''' <summary>
        ''' The entry name used verbatim as the section header in the output
        ''' </summary>
        Public Name As String

        ''' <summary>
        ''' The package folder names under <c> %LocalAppData%\Packages </c>, in declaration
        ''' order. Every <c> Package= </c> key adds one, and any digits in the key name
        ''' (<c> Package1= </c>, <c> Package2= </c>) are ignored, so <c> %Package1% </c> names
        ''' the first one declared.
        ''' </summary>
        Public Packages As List(Of String)

        ''' <summary>
        ''' The <c> LangSecRef= </c> value, or empty when the entry has none
        ''' </summary>
        Public LangSecRef As String

        ''' <summary>
        ''' The <c> Section= </c> value, or empty when the entry has none or also declares
        ''' <c> LangSecRef= </c>
        ''' </summary>
        Public SectionName As String

        ''' <summary>
        ''' <c> Detect= </c> registry key values in their original file order. Generation
        ''' variable-expands them in the registry domain and numbers the results from 1, or
        ''' leaves a lone result unnumbered. Supports hybrid win32+UWP entries that need multiple
        ''' detection conditions.
        ''' </summary>
        Public DetectKeys As List(Of String)

        ''' <summary>
        ''' Additional <c> DetectFile= </c> values in their original file order, appended after
        ''' the scaffold-generated DetectFile keys. Generation variable-expands them but doesn't
        ''' expand <c> %Package% </c> in them. Supports hybrid win32+UWP entries that need file
        ''' system detection for the win32 installation alongside package detection.
        ''' </summary>
        Public DetectFileKeys As List(Of String)

        ''' <summary>
        ''' App-specific FileKey and FileKeyBase values in their original file order.
        ''' Generation expands <c> %Package% </c> / <c> %PackageN% </c> in them and then
        ''' variable-expands every value.
        ''' </summary>
        Public AppKeys As List(Of String)

        ''' <summary>
        ''' <c> RegKey= </c> and <c> RegKeyBase= </c> values in their original file order.
        ''' Generation variable-expands them in the registry domain and renumbers them from 1.
        ''' </summary>
        Public RegKeys As List(Of String)

        ''' <summary>
        ''' <c> ExcludeKey= </c> and <c> ExcludeKeyBase= </c> values in their original file order.
        ''' Generation expands <c> %Package% </c> / <c> %PackageN% </c> in them and then
        ''' variable-expands every value, in the domain its flag names.
        ''' </summary>
        Public ExcludeKeys As List(Of String)

        ''' <summary>
        ''' <c> Warning= </c> prose values in their original file order. Emitted verbatim and
        ''' never variable-expanded, since the values are human-readable text, not path templates
        ''' </summary>
        Public Warnings As List(Of String)

        ''' <summary>
        ''' The <c> DetectOS= </c> value, or empty when the entry declares no OS filter.
        ''' Emitted verbatim; the value is a kernel version range, not a path template
        ''' </summary>
        Public DetectOS As String

        ''' <summary>
        ''' Indicates whether this entry is omitted from generation. Set by <c> Skip= </c>, or by
        ''' <see cref="parseAppInfo"/> when the entry has no package or no category.
        ''' </summary>
        Public ShouldSkip As Boolean

        ''' <summary>
        ''' Indicates whether to leave out the scaffold <c> FileKeyBase= </c> templates from
        ''' <c> UWP.ini </c> for this entry. Scaffold DetectFile keys and the shared catalog
        ''' families are unaffected. Use this when defining a secondary entry for an
        ''' application that shares packages with a primary entry, to avoid duplicating the
        ''' baseline cleaning targets.
        ''' </summary>
        Public SkipUWPFileKeys As Boolean

        ''' <summary>
        ''' Root paths of embedded WebView2 (EBWebView) data folders associated with this app.
        ''' Each entry is a path template that may contain <c> %Package% </c> or
        ''' <c> %PackageN% </c> references, expanded per package at generation time.
        ''' Declaring any path here opts the entry into <c> WebViewScaffold </c> emission;
        ''' an empty list disables WebView scaffolding entirely for this entry.
        ''' </summary>
        Public WebViewPaths As List(Of String)

        ''' <summary>
        ''' Explicit <c> WebViewScaffold </c> names selected for this entry. When
        ''' <see cref="WebViewScaffoldsKeyPresent"/> is set, this list replaces
        ''' <see cref="ScaffoldCatalogs.DefaultScaffolds"/>. Names are matched against the
        ''' <c> [WebViewScaffold: ...] </c> sections in the shared scaffold directory.
        ''' <br /><br />
        '''
        ''' The sentinel value <c> All </c> (case-insensitive) expands to every scaffold in the
        ''' catalog except the legacy tier, including host-risk categories. A legacy scaffold
        ''' named beside <c> All </c> is added, and any other known name beside it is redundant
        ''' and warns. <see cref="ScaffoldCatalogs.ResolveScaffolds"/> has the full rules.
        ''' </summary>
        Public WebViewScaffoldNames As List(Of String)

        ''' <summary>
        ''' <c> WebViewScaffold </c> names to subtract from the active selection (either
        ''' <see cref="WebViewScaffoldNames"/> when set, or
        ''' <see cref="ScaffoldCatalogs.DefaultScaffolds"/> otherwise). Unknown names are ignored.
        ''' </summary>
        Public ExcludedWebViewScaffolds As List(Of String)

        ''' <summary>
        ''' Indicates whether the AppInfo entry declared <c> WebViewScaffolds= </c> at all,
        ''' separately from whether the resulting list is empty. When the key is absent we use
        ''' <see cref="ScaffoldCatalogs.DefaultScaffolds"/>. When it is present but empty we emit
        ''' no WebView scaffold FileKeys, even with a <see cref="WebViewPaths"/> declaration.
        ''' </summary>
        Public WebViewScaffoldsKeyPresent As Boolean

        ''' <summary>
        ''' Paths of embedded QtWebEngine profile directories associated with this app: the
        ''' storage folder itself, profile segment included (e.g. <c> ...\QtWebEngine\Default </c>),
        ''' unlike <see cref="WebViewPaths"/> which names the parent of <c> Default\ </c>. The
        ''' QtWebEngine catalog does not bake the profile segment into its templates, so a host
        ''' running several profiles declares one path per profile. Each entry is a path
        ''' template that may contain <c> %Package% </c> / <c> %PackageN% </c> references,
        ''' expanded per package at generation time, then substituted for
        ''' <c> %QtWebEngineRoot% </c> in QtWebEngine scaffold templates. Declaring any path
        ''' here opts the entry into <c> QtWebEngineScaffold </c> emission.
        ''' </summary>
        Public QtWebEnginePaths As List(Of String)

        ''' <summary>
        ''' Paths of the application's QtWebEngine HTTP cache directories. Declared separately
        ''' from <see cref="QtWebEnginePaths"/> because Qt keeps the cache under its CacheLocation,
        ''' not inside the profile directory. Substituted for <c> %QtWebEngineCacheRoot% </c>;
        ''' when empty, templates referencing that placeholder are dropped rather than emitted
        ''' with the placeholder literal.
        ''' </summary>
        Public QtWebEngineCachePaths As List(Of String)

        ''' <summary>
        ''' Explicit <c> QtWebEngineScaffold </c> names selected for this entry. When
        ''' <see cref="QtWebEngineScaffoldsKeyPresent"/> is set, this list replaces
        ''' <see cref="ScaffoldCatalogs.QtWebEngineDefaultScaffolds"/>. The <c> All </c> sentinel
        ''' works as it does for <see cref="WebViewScaffoldNames"/>.
        ''' </summary>
        Public QtWebEngineScaffoldNames As List(Of String)

        ''' <summary>
        ''' <c> QtWebEngineScaffold </c> names to subtract from the active selection. Unknown
        ''' names are ignored.
        ''' </summary>
        Public ExcludedQtWebEngineScaffolds As List(Of String)

        ''' <summary>
        ''' Indicates whether the AppInfo entry declared <c> QtWebEngineScaffolds= </c> at all
        ''' (mirrors <see cref="WebViewScaffoldsKeyPresent"/> for the QtWebEngine family).
        ''' </summary>
        Public QtWebEngineScaffoldsKeyPresent As Boolean

        ''' <summary>
        ''' Paths of the application's Electron <c> userData </c> folders. For Electron this
        ''' folder <em>is</em> the Chromium profile, with no <c> Default\ </c> segment, unlike
        ''' <see cref="WebViewPaths"/>. Each entry is a path template that may contain
        ''' <c> %Package% </c> / <c> %PackageN% </c> references, expanded per package at
        ''' generation time, then substituted for <c> %ElectronRoot% </c>.
        ''' <br /><br />
        '''
        ''' Relevant to UWPBuilder because a <b> hybrid win32+UWP </b> entry carries the win32
        ''' install's paths alongside the package's, and the win32 half of a hybrid is where
        ''' Electron shows up. A packaged app whose desktop build is Electron needs this even
        ''' though nothing inside the MSIX container is.
        ''' </summary>
        Public ElectronPaths As List(Of String)

        ''' <summary>
        ''' Paths of the application's electron-updater download caches. Declared separately from
        ''' <see cref="ElectronPaths"/> because the directory name is electron-builder's appId
        ''' slug and is not derivable from the userData path. Substituted for
        ''' <c> %ElectronUpdaterRoot% </c>; when empty, templates referencing that placeholder
        ''' are dropped rather than emitted with the placeholder literal.
        ''' </summary>
        Public ElectronUpdaterPaths As List(Of String)

        ''' <summary>
        ''' Explicit <c> ElectronScaffold </c> names selected for this entry. When
        ''' <see cref="ElectronScaffoldsKeyPresent"/> is set, this list replaces
        ''' <see cref="ScaffoldCatalogs.ElectronDefaultScaffolds"/>. The <c> All </c> sentinel
        ''' works as it does for <see cref="WebViewScaffoldNames"/>.
        ''' </summary>
        Public ElectronScaffoldNames As List(Of String)

        ''' <summary>
        ''' <c> ElectronScaffold </c> names to subtract from the active selection. Unknown
        ''' names are ignored.
        ''' </summary>
        Public ExcludedElectronScaffolds As List(Of String)

        ''' <summary>
        ''' Indicates whether the AppInfo entry declared <c> ElectronScaffolds= </c> at all
        ''' (mirrors <see cref="WebViewScaffoldsKeyPresent"/> for the Electron family).
        ''' </summary>
        Public ElectronScaffoldsKeyPresent As Boolean

        ''' <summary>
        ''' The entry's open-vocabulary variable declarations: every key whose type (its name
        ''' without digits) the parser's <c> Select Case </c> has no branch for. Each declares a
        ''' comma-separated list of values fanned out at generation time by
        ''' <see cref="VariableExpander"/>, driving <c> &lt;token&gt; </c> expansion across this
        ''' entry's key templates
        ''' </summary>
        Public Variables As VariableSet

        ''' <summary>
        ''' How many of this entry's variable declarations reference another declared variable,
        ''' from <see cref="VariableSet.NestedReferenceCount"/>, captured before
        ''' <see cref="VariableExpander.ResolveAll"/> flattens the symbol table. UWPBuilder
        ''' records it but doesn't report it.
        ''' </summary>
        Public NestedVariableRefs As Integer

        ''' <summary>
        ''' Creates a new <c> UWPAppInfo </c> for an entry with the given name,
        ''' initializing all list fields to empty collections
        ''' </summary>
        '''
        ''' <param name="name">
        ''' The entry name, taken verbatim from the AppInfo section header
        ''' </param>
        Public Sub New(name As String)

            Me.Name = name
            Packages = New List(Of String)
            LangSecRef = ""
            SectionName = ""
            DetectKeys = New List(Of String)
            DetectFileKeys = New List(Of String)
            AppKeys = New List(Of String)
            RegKeys = New List(Of String)
            ExcludeKeys = New List(Of String)
            Warnings = New List(Of String)
            DetectOS = ""
            ShouldSkip = False
            SkipUWPFileKeys = False
            WebViewPaths = New List(Of String)
            WebViewScaffoldNames = New List(Of String)
            ExcludedWebViewScaffolds = New List(Of String)
            WebViewScaffoldsKeyPresent = False
            QtWebEnginePaths = New List(Of String)
            QtWebEngineCachePaths = New List(Of String)
            QtWebEngineScaffoldNames = New List(Of String)
            ExcludedQtWebEngineScaffolds = New List(Of String)
            QtWebEngineScaffoldsKeyPresent = False
            ElectronPaths = New List(Of String)
            ElectronUpdaterPaths = New List(Of String)
            ElectronScaffoldNames = New List(Of String)
            ExcludedElectronScaffolds = New List(Of String)
            ElectronScaffoldsKeyPresent = False
            Variables = New VariableSet()
            NestedVariableRefs = 0

        End Sub

    End Structure

    ''' <summary>
    ''' Binds the <c> -1f </c> / <c> -1d </c> through <c> -3f </c> / <c> -3d </c> arguments to <see cref="UWPFile1"/>,
    ''' <see cref="UWPFile2"/> and <see cref="UWPFile3"/>, starting from the default settings
    ''' rather than any saved ones, then runs the build
    ''' </summary>
    Public Sub handleCmdLine()

        InitDefaultUWPBuilderSettings()

        Dim spec As New CliArgSpec("uwpbuilder")
        spec.WithFile(1, UWPFile1).WithFile(2, UWPFile2).WithFile(3, UWPFile3).Parse()

        initUWPBuilder()

    End Sub

    ''' <summary>
    ''' Reads the source directory, runs the generation pipeline, writes the output file,
    ''' and displays the results unless <c> SuppressOutput </c> is set. When <c> UWP.ini </c>
    ''' is missing or empty, or <c> AppInfo\ </c> yields no sections, we set a header message
    ''' and return without writing anything. Of those cases only a missing <c> UWP.ini </c>
    ''' fails the run, and a missing source directory throws from <see cref="iniFile.FromFile"/>.
    ''' </summary>
    Public Sub initUWPBuilder()

        clrConsole()

        Dim sourceDir = UWPFile1.Dir
        Dim templateIni = iniFile.FromFile($"{sourceDir}\UWP.ini")

        If templateIni.Count = 0 Then

            setNextMenuHeaderText($"UWP.ini not found or empty in: {sourceDir}", printColor:=ConsoleColor.Red)
            Return

        End If

        Dim appInfoDir = $"{sourceDir}\AppInfo"
        Dim appsIni = combineAppInfoDir(appInfoDir)

        If appsIni.Count = 0 Then

            setNextMenuHeaderText($"No app definitions found in: {appInfoDir}", printColor:=ConsoleColor.Red)
            Return

        End If

        ' The shared scaffold directory is configured via UWPFile3 (typically
        ' Assembler\Scaffolds). Every catalog in it is loaded at once and families are derived
        ' from section headers, so a new family costs no new setting here. A missing directory
        ' or catalog warns and yields zero scaffold FileKeys rather than aborting the run.
        Dim scaffoldDir = UWPFile3.Dir

        Dim output As New MenuSection
        output.AddBoxWithText("Building UWP app entries")

        Using gLogScope("Building UWP app entries")

            processUWPBuilder(templateIni, appsIni, scaffoldDir, output)

            output.AddBoxWithText("UWP app entries built successfully")
            gLog("UWP app entries built successfully")

        End Using

        output.AddAnyKeyPrompt()

        If SuppressOutput Then Return

        output.Print()
        crk()

    End Sub

    ''' <summary>
    ''' Orchestrates the UWP builder process: parses scaffold templates and app definitions,
    ''' generates one <c> iniSection </c> per app that isn't skipped, warns about variables an
    ''' entry declared but never referenced, lints the result through
    ''' <see cref="remotedebugGuarded"/>, and writes it to <see cref="UWPFile2"/> under a header
    ''' comment block.
    ''' </summary>
    '''
    ''' <param name="templateIni">
    ''' The parsed <c> UWP.ini </c> scaffold template
    ''' </param>
    '''
    ''' <param name="appsIni">
    ''' The combined AppInfo sections from all per-letter source files
    ''' </param>
    '''
    ''' <param name="scaffoldDir">
    ''' Path to the shared scaffold directory (typically
    ''' <c> Assembler\Scaffolds </c>). Every catalog in it is loaded once per run via
    ''' <see cref="ScaffoldCatalogs.LoadCatalogDirectory"/>, which derives each family from its
    ''' section headers, and consumed by per-entry scaffolding.
    ''' </param>
    '''
    ''' <param name="menuOutput">
    ''' The <c> MenuSection </c> receiving progress lines and warnings for display
    ''' </param>
    Private Sub processUWPBuilder(templateIni As iniFile,
                                   appsIni As iniFile,
                                   scaffoldDir As String,
                                   menuOutput As MenuSection)

        Using gLogScope("Processing UWP builder files")

            Dim scaffoldFileKeys As New List(Of String)
            Dim scaffoldDetectFiles As New List(Of String)
            Dim scaffoldSet = ScaffoldCatalogs.LoadCatalogDirectory(scaffoldDir, menuOutput)
            Dim apps As New List(Of UWPAppInfo)

            For Each section In templateIni

                Select Case True

                    Case section.Name.StartsWith("EntryScaffold:", StringComparison.InvariantCulture)

                        parseScaffold(section, scaffoldFileKeys, scaffoldDetectFiles, menuOutput)

                    Case Else

                        Dim logMsg = $"Unexpected section in template file: [{section.Name}]"
                        gLog(logMsg)
                        menuOutput.AddWarning(logMsg)

                End Select

            Next

            Dim scaffoldMsg = $"Loaded {scaffoldDetectFiles.Count} scaffold DetectFile template(s), {scaffoldFileKeys.Count} scaffold FileKey template(s)"
            menuOutput.AddColoredLine(scaffoldMsg, ConsoleColor.Yellow)
            gLog(scaffoldMsg)

            For Each family In ScaffoldCatalogs.ScaffoldFamilies

                Dim familyMsg = $"Loaded {scaffoldSet.ForFamily(family).Count} {family} scaffold(s) from {scaffoldDir}"
                menuOutput.AddColoredLine(familyMsg, ConsoleColor.Yellow)
                gLog(familyMsg)

            Next

            For Each section In appsIni

                Dim app As UWPAppInfo = parseAppInfo(section, menuOutput)
                If Not app.ShouldSkip Then apps.Add(app)

            Next

            Dim appsMsg = $"Loaded {apps.Count} app definition(s)"
            menuOutput.AddColoredLine(appsMsg, ConsoleColor.Yellow)
            gLog(appsMsg)

            Dim outputFile = iniFile.Empty(UWPFile2.Dir, UWPFile2.Name)

            For Each app In apps

                Dim entrySection = generateUWPEntry(app, scaffoldFileKeys, scaffoldDetectFiles, scaffoldSet, menuOutput)
                outputFile.AddSection(entrySection)

                ' Typo backstop: variables declared on the entry but never referenced by any
                ' <token> in any expanded key. This is what replaces the parser's former
                ' "Unexpected key type" warning now that the vocabulary is open: a misspelled
                ' reserved key (Pakcage=) lands here as an unreferenced declaration.
                For Each unused In app.Variables.UnreferencedNames()

                    Dim unusedMsg = $"Variable '{unused}=' declared in [{app.Name}] but never referenced by any <{unused}> token; possible typo"
                    gLog(unusedMsg)
                    menuOutput.AddWarning(unusedMsg)

                Next

            Next

            Dim generatedMsg = $"Generated {outputFile.Count} UWP entries"
            menuOutput.AddColoredLine(generatedMsg, ConsoleColor.Yellow)
            gLog($" {generatedMsg}")

            outputFile = remotedebugGuarded(outputFile, NameOf(UWPBuilder), menuOutput)

            Dim sb As New StringBuilder()
            sb.AppendLine($"; # of entries: {outputFile.Count:#,###}")
            sb.AppendLine($"; {UWPFile2.Name} is generated by the Winapp2ool UWP Builder")
            sb.AppendLine("; Entries in this file may be incomplete and are not intended to be used directly with any cleaning software")
            sb.AppendLine("; They are utilized by winapp2ool to create the final winapp2.ini file for distribution")
            sb.AppendLine("; If you are not maintaining winapp2.ini for distribution, you probably don't need this file!")
            sb.AppendLine("; Refer to the Winapp2ool documentation for more information: " & readMeUrl)
            sb.AppendLine("; You can find the complete winapp2.ini file here: " & baseFlavorLink)
            sb.AppendLine()
            sb.Append(outputFile.ToString())

            outputFile.OverwriteToFile(sb.ToString())

            gLog("UWP builder files processed successfully")

        End Using

    End Sub

    ''' <summary>
    ''' Parses an <c> [EntryScaffold: ...] </c> section from the template file, collecting
    ''' its <c> DetectFileBase= </c> values into <paramref name="scaffoldDetectFiles"/> and
    ''' its <c> FileKeyBase= </c> values into <paramref name="scaffoldFileKeys"/>.
    ''' Warns on any unrecognized key types.
    ''' </summary>
    '''
    ''' <param name="scaffoldSection">
    ''' The <c> [EntryScaffold: ...] </c> section to parse
    ''' </param>
    '''
    ''' <param name="scaffoldFileKeys">
    ''' The accumulator list to which parsed <c> FileKeyBase= </c> values are appended
    ''' </param>
    '''
    ''' <param name="scaffoldDetectFiles">
    ''' The accumulator list to which parsed <c> DetectFileBase= </c> values are appended
    ''' </param>
    '''
    ''' <param name="menuOutput">
    ''' The <c> MenuSection </c> receiving progress lines and warnings for display
    ''' </param>
    Private Sub parseScaffold(scaffoldSection As iniSection,
                               scaffoldFileKeys As List(Of String),
                               scaffoldDetectFiles As List(Of String),
                               menuOutput As MenuSection)

        Dim scaffoldName = scaffoldSection.Name.Substring("EntryScaffold:".Length).Trim()
        Dim scaffoldMsg = $"Processing scaffold: {scaffoldName}"
        menuOutput.AddColoredLine(scaffoldMsg, ConsoleColor.Magenta)
        gLog($"  {scaffoldMsg}")

        For Each key In scaffoldSection.Keys

            Select Case key.KeyType.ToUpperInvariant()

                Case "DETECTFILEBASE" : scaffoldDetectFiles.Add(key.Value)

                Case "FILEKEYBASE" : scaffoldFileKeys.Add(key.Value)

                Case Else

                    Dim errMsg = $"Unexpected key in scaffold [{scaffoldSection.Name}]: {key.Name}"
                    gLog(errMsg)
                    menuOutput.AddWarning(errMsg)

            End Select

        Next

    End Sub

    ''' <summary>
    ''' Splits a comma-separated value into a trimmed, non-empty list of tokens.
    ''' An empty input yields an empty list.
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
    ''' Parses an AppInfo section and returns a populated <c> UWPAppInfo </c> structure.
    ''' Issues warnings and sets <c> ShouldSkip </c> for entries that are structurally
    ''' invalid (missing package or missing category). An entry declaring both
    ''' <c> LangSecRef= </c> and <c> Section= </c> warns and keeps <c> LangSecRef </c>.
    ''' <br /><br />
    '''
    ''' <c> SpecialDetect= </c> and <c> Default= </c> warn and are dropped. Any key type
    ''' without a branch here becomes a variable declaration, and we resolve the
    ''' declarations' nested references through <see cref="VariableExpander.ResolveAll"/>
    ''' before returning. A single-valued key that repeats (<c> DetectOS= </c>,
    ''' <c> LangSecRef= </c>, a scaffold selection list) keeps its last value.
    ''' </summary>
    '''
    ''' <param name="appSection">
    ''' The AppInfo section describing one UWP application
    ''' </param>
    '''
    ''' <param name="menuOutput">
    ''' The <c> MenuSection </c> receiving warnings for display
    ''' </param>
    '''
    ''' <returns>
    ''' A <c> UWPAppInfo </c> populated from <paramref name="appSection"/>.
    ''' Check <c> ShouldSkip </c> before using the result.
    ''' </returns>
    Friend Function parseAppInfo(appSection As iniSection,
                                  menuOutput As MenuSection) As UWPAppInfo

        Dim app As New UWPAppInfo(appSection.Name)

        For Each key In appSection.Keys

            Select Case key.KeyType.ToUpperInvariant()

                Case "PACKAGE"

                    ' Package values are not variable-expanded: a fan-out here would change
                    ' the package list's length and silently shift the meaning of every
                    ' %PackageN% selector in the entry. Warn rather than emit a junk path.
                    If key.Value.Contains("<") OrElse key.Value.Contains(">") Then

                        Dim pkgTokenMsg = $"Package= value in [{app.Name}] contains '<' or '>'; Package is never variable-expanded and the value is used literally"
                        gLog(pkgTokenMsg)
                        menuOutput.AddWarning(pkgTokenMsg)

                    End If

                    app.Packages.Add(key.Value)

                Case "LANGSECREF" : app.LangSecRef = key.Value

                Case "SECTION" : app.SectionName = key.Value

                Case "DETECT" : app.DetectKeys.Add(key.Value)

                Case "DETECTFILE" : app.DetectFileKeys.Add(key.Value)

                Case "DETECTOS" : app.DetectOS = key.Value

                Case "SPECIALDETECT"

                    Dim sdMsg = $"SpecialDetect is deprecated; key dropped from [{app.Name}]. Replace with Detect or DetectFile"
                    gLog(sdMsg)
                    menuOutput.AddWarning(sdMsg)

                Case "WARNING" : app.Warnings.Add(key.Value)

                Case "DEFAULT"

                    Dim defaultMsg = $"Default= declared in [{app.Name}]; UWPBuilder never emits Default, ignoring"
                    gLog(defaultMsg)
                    menuOutput.AddWarning(defaultMsg)

                Case "FILEKEYBASE", "FILEKEY" : app.AppKeys.Add(key.Value)

                Case "REGKEY", "REGKEYBASE" : app.RegKeys.Add(key.Value)

                Case "EXCLUDEKEY", "EXCLUDEKEYBASE" : app.ExcludeKeys.Add(key.Value)

                Case "SKIP" : app.ShouldSkip = True

                Case "SKIPUWPFILEKEYS" : app.SkipUWPFileKeys = True

                Case "WEBVIEWPATH", "WEBVIEWROOT" : app.WebViewPaths.Add(key.Value)

                Case "WEBVIEWSCAFFOLDS"

                    app.WebViewScaffoldNames = splitCsv(key.Value)
                    app.WebViewScaffoldsKeyPresent = True

                Case "EXCLUDEWEBVIEWSCAFFOLDS" : app.ExcludedWebViewScaffolds = splitCsv(key.Value)

                Case "QTWEBENGINEPATH", "QTWEBENGINEROOT" : app.QtWebEnginePaths.Add(key.Value)

                Case "QTWEBENGINECACHEPATH", "QTWEBENGINECACHEROOT" : app.QtWebEngineCachePaths.Add(key.Value)

                Case "QTWEBENGINESCAFFOLDS"

                    app.QtWebEngineScaffoldNames = splitCsv(key.Value)
                    app.QtWebEngineScaffoldsKeyPresent = True

                Case "EXCLUDEQTWEBENGINESCAFFOLDS" : app.ExcludedQtWebEngineScaffolds = splitCsv(key.Value)

                Case "ELECTRONPATH", "ELECTRONROOT" : app.ElectronPaths.Add(key.Value)

                Case "ELECTRONUPDATERPATH", "ELECTRONUPDATERROOT" : app.ElectronUpdaterPaths.Add(key.Value)

                Case "ELECTRONSCAFFOLDS"

                    app.ElectronScaffoldNames = splitCsv(key.Value)
                    app.ElectronScaffoldsKeyPresent = True

                Case "EXCLUDEELECTRONSCAFFOLDS" : app.ExcludedElectronScaffolds = splitCsv(key.Value)

                Case Else

                    ' Open-vocabulary: any unrecognised key is a list-variable declaration
                    ' (e.g. Version=11.0,16.0). This replaces the former "Unexpected key type"
                    ' warning, whose typo-catching role is taken over by two backstops:
                    ' (a) undeclared-token warnings during expansion catch <Verison> for a
                    ' Version= declaration, and (b) the UnreferencedNames sweep after
                    ' generation catches Versoin= that nothing references.
                    '
                    ' Defensive collision check mirroring EntryBuilder: with the current
                    ' Select Case structure this is dead code, but it keeps UWPReservedKeys
                    ' and the parser's Case set explicitly cross-referenced so a future
                    ' refactor that drops a branch surfaces the omission.
                    If UWPReservedKeys.Any(Function(r) String.Equals(r, key.KeyType, StringComparison.InvariantCultureIgnoreCase)) Then

                        Dim shadowMsg = $"Variable declaration '{key.KeyType}=' in [{app.Name}] shadows reserved key name; possible typo"
                        gLog(shadowMsg)
                        menuOutput.AddWarning(shadowMsg)

                    End If

                    app.Variables.Add(key.KeyType, key.Value)

            End Select

        Next

        ' Capture the symbol table's nesting depth before ResolveAll flattens it, then
        ' resolve nested <token> references inside the declarations themselves so that
        ' generation-time expansion sees fully-literal value lists.
        app.NestedVariableRefs = app.Variables.NestedReferenceCount()

        For Each diag In VariableExpander.ResolveAll(app.Variables)

            Dim contextualMsg = $"[{app.Name}] {diag.Message}"
            gLog(contextualMsg)
            If diag.Severity = DiagnosticSeverity.Warning Then menuOutput.AddWarning(contextualMsg)

        Next

        Dim noPackage = app.Packages.Count = 0
        Dim noCategory = app.LangSecRef.Length = 0 AndAlso app.SectionName.Length = 0
        Dim bothCategories = app.LangSecRef.Length > 0 AndAlso app.SectionName.Length > 0

        Dim msgs = {$"No Package key in [{app.Name}], skipping",
                    $"No LangSecRef or Section in [{app.Name}], skipping"}

        Dim bools = {noPackage, noCategory}
        For i = 0 To bools.Count - 1

            If Not bools(i) Then Continue For

            gLog(msgs(i))
            menuOutput.AddWarning(msgs(i))
            app.ShouldSkip = True

        Next

        If Not bothCategories Then Return app

        Dim bothMsg = $"Both LangSecRef and Section present in [{app.Name}], using LangSecRef"
        gLog(bothMsg)
        menuOutput.AddWarning(bothMsg)
        app.SectionName = ""

        Return app

    End Function

    ''' <summary>
    ''' One scaffold family's per-entry state, gathered so <see cref="generateUWPEntry"/> can
    ''' emit every family from a single loop instead of one hand-written block per family.
    ''' Adding a family becomes a case in <see cref="scaffoldFamiliesFor"/> rather than another
    ''' copy of the emission nesting.
    ''' </summary>
    Private Structure UWPScaffoldFamily

        ''' <summary>
        ''' The family token, e.g. <c> Electron </c>. We look up the family's catalog and default
        ''' set by it, and it phrases diagnostics.
        ''' </summary>
        Public Label As String

        ''' <summary> Names from <c> {Family}Scaffolds= </c>, empty when undeclared </summary>
        Public Selection As List(Of String)

        ''' <summary> Names from <c> Exclude{Family}Scaffolds= </c> </summary>
        Public Excluded As List(Of String)

        ''' <summary> Indicates whether <c> {Family}Scaffolds= </c> was declared at all </summary>
        Public KeyPresent As Boolean

        ''' <summary>
        ''' The family's (placeholder, roots) pairs, with <c> %Package% </c> already expanded in
        ''' the roots. A root can still hold <c> &lt;token&gt; </c> references.
        ''' QtWebEngine and Electron carry two pairs; WebView one.
        ''' </summary>
        Public Bindings As List(Of ScaffoldCatalogs.ScaffoldRootBinding)

        ''' <summary>
        ''' Indicates whether the entry opted into this family by declaring a root for any of
        ''' the family's placeholders. QtWebEngine and Electron opt in on either root, since an
        ''' entry may want only the cache or updater templates.
        ''' </summary>
        Public ReadOnly Property IsDeclared As Boolean
            Get
                Return Bindings.Any(Function(b) b.Roots.Count > 0)
            End Get
        End Property

        ''' <summary>Creates a new <c> UWPScaffoldFamily </c> holding one family's per-entry state</summary>
        '''
        ''' <param name="label">
        ''' The family token
        ''' </param>
        '''
        ''' <param name="selection">
        ''' Names from the family's <c> {Family}Scaffolds= </c> key
        ''' </param>
        '''
        ''' <param name="excluded">
        ''' Names from the family's <c> Exclude{Family}Scaffolds= </c> key
        ''' </param>
        '''
        ''' <param name="keyPresent">
        ''' Indicates whether <c> {Family}Scaffolds= </c> was declared
        ''' </param>
        '''
        ''' <param name="bindings">
        ''' The family's (placeholder, roots) pairs with <c> %Package% </c> already expanded in
        ''' the roots
        ''' </param>
        Public Sub New(label As String,
                       selection As List(Of String),
                       excluded As List(Of String),
                       keyPresent As Boolean,
                       bindings As List(Of ScaffoldCatalogs.ScaffoldRootBinding))

            Me.Label = label
            Me.Selection = selection
            Me.Excluded = excluded
            Me.KeyPresent = keyPresent
            Me.Bindings = bindings

        End Sub

    End Structure

    ''' <summary>
    ''' Builds the per-entry state of every scaffold family for <paramref name="app"/>, whether
    ''' or not the entry declared it, expanding <c> %Package% </c> in each declared root. This
    ''' list has to cover every family in <see cref="ScaffoldCatalogs.ScaffoldFamilies"/>. The
    ''' <c> ScaffoldFamilies_KeyVocabularyPresentInBothBuilders </c> test checks only
    ''' <see cref="UWPReservedKeys"/>, so a family missing from this list would not fail it.
    ''' <br /><br />
    '''
    ''' Package expansion (phase 0) must happen here, before placeholder substitution (phase 1),
    ''' because roots themselves reference <c> %Package% </c>
    ''' (<c> WebViewRoot=%Package%\LocalState\EBWebView </c>). Phase 2 (<c> &lt;&gt; </c>
    ''' expansion) then runs last on the merged string, so a token inside a root declaration
    ''' joins the same cartesian product as the scaffold template's own tokens.
    ''' </summary>
    '''
    ''' <param name="app">
    ''' The parsed app definition whose root declarations are being gathered
    ''' </param>
    '''
    ''' <returns>
    ''' One entry per scaffold family, in <see cref="ScaffoldCatalogs.ScaffoldFamilies"/> order
    ''' </returns>
    Private Function scaffoldFamiliesFor(app As UWPAppInfo) As List(Of UWPScaffoldFamily)

        Dim expandRoots = Function(templates As List(Of String)) templates _
            .SelectMany(Function(t) expandPackageKey(t, app.Packages)) _
            .ToList()

        Return New List(Of UWPScaffoldFamily) From {
            New UWPScaffoldFamily("WebView", app.WebViewScaffoldNames, app.ExcludedWebViewScaffolds,
                                  app.WebViewScaffoldsKeyPresent,
                                  New List(Of ScaffoldCatalogs.ScaffoldRootBinding) From {
                                      New ScaffoldCatalogs.ScaffoldRootBinding("%WebViewRoot%", expandRoots(app.WebViewPaths))}),
            New UWPScaffoldFamily("QtWebEngine", app.QtWebEngineScaffoldNames, app.ExcludedQtWebEngineScaffolds,
                                  app.QtWebEngineScaffoldsKeyPresent,
                                  New List(Of ScaffoldCatalogs.ScaffoldRootBinding) From {
                                      New ScaffoldCatalogs.ScaffoldRootBinding("%QtWebEngineRoot%", expandRoots(app.QtWebEnginePaths)),
                                      New ScaffoldCatalogs.ScaffoldRootBinding("%QtWebEngineCacheRoot%", expandRoots(app.QtWebEngineCachePaths))}),
            New UWPScaffoldFamily("Electron", app.ElectronScaffoldNames, app.ExcludedElectronScaffolds,
                                  app.ElectronScaffoldsKeyPresent,
                                  New List(Of ScaffoldCatalogs.ScaffoldRootBinding) From {
                                      New ScaffoldCatalogs.ScaffoldRootBinding("%ElectronRoot%", expandRoots(app.ElectronPaths)),
                                      New ScaffoldCatalogs.ScaffoldRootBinding("%ElectronUpdaterRoot%", expandRoots(app.ElectronUpdaterPaths))})}

    End Function

    ''' <summary>
    ''' Emits one scaffold family's FileKey values for an app. Resolution follows the shared
    ''' grammar in <see cref="ScaffoldCatalogs.ResolveScaffolds"/>: an explicit
    ''' <c> {Family}Scaffolds= </c> wins (a present-but-empty value yields nothing, distinct
    ''' from the key being absent), the <c> All </c> sentinel expands the catalog minus its
    ''' <c> Tier=Legacy </c> scaffolds,
    ''' otherwise the family's default set applies; exclusions are then subtracted and unknown
    ''' names dropped with a warning. Substitution is delegated to
    ''' <see cref="ScaffoldCatalogs.BindFamilyTemplates"/>, so a template whose placeholder has
    ''' no declared root is dropped rather than emitted literally. That keeps Electron's
    ''' default-on <c> UpdaterCache </c> inert for an entry with no updater root.
    ''' <br /><br />
    '''
    ''' A non-empty scaffold that produced nothing warns only when the entry declared
    ''' <c> {Family}Scaffolds= </c>, <c> All </c> included; a default-set member yielding
    ''' nothing is by design and would otherwise warn on nearly every Electron entry. The
    ''' warning blames a missing root even when phase 2 dropped the keys instead. Mirrors
    ''' EntryBuilder's <c> expandScaffoldFamily </c>.
    ''' </summary>
    '''
    ''' <param name="family">
    ''' The family's per-entry state, with roots already package-expanded
    ''' </param>
    '''
    ''' <param name="app">
    ''' The entry being generated, supplying the variable set for phase-2 expansion
    ''' </param>
    '''
    ''' <param name="catalog">
    ''' The family's scaffold catalog
    ''' </param>
    '''
    ''' <param name="menuOutput">
    ''' The <c> MenuSection </c> receiving warnings for display
    ''' </param>
    '''
    ''' <returns>
    ''' Every expanded FileKey value produced by this family, in fan-out order
    ''' </returns>
    Private Function expandScaffoldFamily(family As UWPScaffoldFamily,
                                          app As UWPAppInfo,
                                          catalog As Dictionary(Of String, List(Of String)),
                                          menuOutput As MenuSection) As List(Of String)

        Dim selected = ScaffoldCatalogs.ResolveScaffolds(family.Selection, family.KeyPresent,
                                                         family.Excluded, catalog,
                                                         ScaffoldCatalogs.DefaultsForFamily(family.Label),
                                                         family.Label, app.Name, menuOutput)

        Dim result As New List(Of String)

        For Each bound In ScaffoldCatalogs.BindFamilyTemplates(family.Bindings, selected, catalog)

            Dim scaffoldYield = 0

            For Each expanded In bound.Templates

                Dim produced = expandPhase2(expanded, app, ExpansionDomain.Filesystem, "FileKey", menuOutput)
                scaffoldYield += produced.Count
                result.AddRange(produced)

            Next

            If scaffoldYield = 0 AndAlso bound.TemplateCount > 0 AndAlso family.KeyPresent Then

                Dim noRootMsg = $"Scaffold '{bound.ScaffoldName}' selected by [{app.Name}] produced no keys; its templates reference a root the entry did not declare"
                gLog(noRootMsg)
                menuOutput.AddWarning(noRootMsg)

            End If

        Next

        Return result

    End Function

    ''' <summary>
    ''' Generates a winapp2.ini entry section for the given <c> UWPAppInfo </c> by applying
    ''' the scaffold template keys and app-specific keys, expanding any
    ''' <c> %Package% </c> / <c> %PackageN% </c> variables and <c> &lt;token&gt; </c>
    ''' references along the way. Keys are emitted in this order: category, Detect,
    ''' DetectFile, DetectOS, Warning, FileKey, RegKey, ExcludeKey. FileKeys run on from
    ''' <c> UWP.ini </c>'s scaffold templates to the app's own keys and then the shared
    ''' catalog families.
    ''' </summary>
    '''
    ''' <param name="app">
    ''' The parsed app definition to generate an entry for
    ''' </param>
    '''
    ''' <param name="scaffoldFileKeys">
    ''' The <c> FileKeyBase= </c> templates from <c> UWP.ini </c>, applied to every app
    ''' </param>
    '''
    ''' <param name="scaffoldDetectFiles">
    ''' The <c> DetectFileBase= </c> templates from <c> UWP.ini </c>, expanded per package
    ''' to produce the entry's detection paths
    ''' </param>
    '''
    ''' <param name="scaffoldSet">
    ''' The shared scaffold catalogs loaded from the scaffold directory. An app draws additional
    ''' FileKeys from every family it declared a root for, with that family's placeholders
    ''' substituted per declared path.
    ''' </param>
    '''
    ''' <param name="menuOutput">
    ''' The <c> MenuSection </c> receiving progress lines and warnings for display
    ''' </param>
    '''
    ''' <returns>
    ''' A fully populated <c> iniSection </c> ready to be added to the output file
    ''' </returns>
    Friend Function generateUWPEntry(app As UWPAppInfo,
                                      scaffoldFileKeys As List(Of String),
                                      scaffoldDetectFiles As List(Of String),
                                      scaffoldSet As ScaffoldCatalogSet,
                                      menuOutput As MenuSection) As iniSection

        Dim generatingMsg = $"Generating entry: {app.Name}"
        menuOutput.AddColoredLine(generatingMsg, ConsoleColor.Magenta)
        gLog($"  {generatingMsg}")

        Dim section As New iniSection(app.Name)
        Dim fileKeyNum As Integer = 1

        ' 1. Category
        Select Case True

            Case app.LangSecRef.Length > 0 : section.AddKey(New iniKey($"LangSecRef={app.LangSecRef}"))

            Case app.SectionName.Length > 0 : section.AddKey(New iniKey($"Section={app.SectionName}"))

            Case Else : gLog($"{app.Name} has no category key")

        End Select

        ' 2. Detect keys (Registry domain: an undeclared token stays literal rather than
        '    dropping the key), unnumbered if single, numbered from 1 if multiple.
        '    Detect values carry no %Package% reference, so phase 0 does not apply.
        Dim detectValues As New List(Of String)
        For Each d In app.DetectKeys

            detectValues.AddRange(expandPhase2(d, app, ExpansionDomain.Registry, "Detect", menuOutput))

        Next

        If detectValues.Count = 1 Then

            section.AddKey(New iniKey($"Detect={detectValues(0)}"))

        Else

            For i = 0 To detectValues.Count - 1

                section.AddKey(New iniKey($"Detect{i + 1}={detectValues(i)}"))

            Next

        End If

        ' 3. DetectFile keys, scaffold templates expanded per package first, then app-specific,
        '    unnumbered if only one total. App-specific values are deliberately not phase-0
        '    expanded, since only the scaffold templates reference %Package%.
        Dim allDetectFiles As New List(Of String)
        allDetectFiles.AddRange(expandPackageAndVars(scaffoldDetectFiles, app, ExpansionDomain.Filesystem, "DetectFile", menuOutput))

        For Each df In app.DetectFileKeys

            allDetectFiles.AddRange(expandPhase2(df, app, ExpansionDomain.Filesystem, "DetectFile", menuOutput))

        Next

        If allDetectFiles.Count = 1 Then

            section.AddKey(New iniKey($"DetectFile={allDetectFiles(0)}"))

        Else

            For i = 0 To allDetectFiles.Count - 1

                section.AddKey(New iniKey($"DetectFile{i + 1}={allDetectFiles(i)}"))

            Next

        End If

        ' 3b. DetectOS, emitted verbatim: a kernel version range, not a path template
        If app.DetectOS.Length > 0 Then section.AddKey(New iniKey($"DetectOS={app.DetectOS}"))

        ' 3c. Warnings, emitted verbatim: prose is never variable-expanded
        For Each w In app.Warnings

            section.AddKey(New iniKey($"Warning={w}"))

        Next

        ' 4. Scaffold template FileKeys applied to all packages (suppressed when SkipUWPFileKeys is set)
        If app.SkipUWPFileKeys Then

            Dim skipMsg = $"Skipping scaffold FileKeys for: {app.Name}"
            menuOutput.AddColoredLine(skipMsg, ConsoleColor.DarkYellow)
            gLog($"  {skipMsg}")

        Else

            For Each expanded In expandPackageAndVars(scaffoldFileKeys, app, ExpansionDomain.Filesystem, "FileKey", menuOutput)

                section.AddKey(New iniKey($"FileKey{fileKeyNum}={expanded}"))
                fileKeyNum += 1

            Next

        End If

        ' 5. App-specific FileKey / FileKeyBase values in document order
        For Each expanded In expandPackageAndVars(app.AppKeys, app, ExpansionDomain.Filesystem, "FileKey", menuOutput)

            section.AddKey(New iniKey($"FileKey{fileKeyNum}={expanded}"))
            fileKeyNum += 1

        Next

        ' 5b. Scaffold FileKeys for every family the entry opted into by declaring a root.
        '     Phase 0 (%Package% fan-out) already ran in scaffoldFamiliesFor; expandScaffoldFamily
        '     does phase 1 (root substitution) then phase 2 (<> expansion) on the merged string,
        '     so a <token> inside a root declaration joins the same cartesian product as the
        '     template's own tokens. A family with no declared root emits nothing.
        For Each family In scaffoldFamiliesFor(app)

            If Not family.IsDeclared Then Continue For

            For Each emitted In expandScaffoldFamily(family, app, scaffoldSet.ForFamily(family.Label), menuOutput)

                section.AddKey(New iniKey($"FileKey{fileKeyNum}={emitted}"))
                fileKeyNum += 1

            Next

        Next

        ' 6. RegKeys, variable-expanded in the Registry domain and renumbered from 1. Registry
        '    paths never reference %Package% (which resolves to a file system path), so phase 0
        '    does not apply here.
        Dim regKeyNum As Integer = 1
        For Each regKey In app.RegKeys

            For Each expanded In expandPhase2(regKey, app, ExpansionDomain.Registry, "RegKey", menuOutput)

                section.AddKey(New iniKey($"RegKey{regKeyNum}={expanded}"))
                regKeyNum += 1

            Next

        Next

        ' 7. ExcludeKeys expanded for %Package% / %PackageN% then variables, renumbered from 1.
        '    Domain is classified per expanded value from the key's own flag.
        Dim exclNum As Integer = 1
        For Each expanded In expandExcludeKeys(app.ExcludeKeys, app, menuOutput)

            section.AddKey(New iniKey($"ExcludeKey{exclNum}={expanded}"))
            exclNum += 1

        Next

        Dim generatedMsg = $"Generated entry: {app.Name}"
        menuOutput.AddColoredLine(generatedMsg, ConsoleColor.Yellow)
        gLog($"        {generatedMsg}")

        Return section

    End Function

    ''' <summary>
    ''' Combines all <c> *.ini </c> files in <paramref name="appInfoDir"/> into a single
    ''' in-memory <c> iniFile </c>. Files are processed in sorted order, top level only.
    ''' A section whose name an earlier file already used is silently dropped (first file
    ''' wins), and so is a repeated section within one file.
    ''' </summary>
    '''
    ''' <param name="appInfoDir">
    ''' Path to the <c> AppInfo\ </c> directory containing per-letter <c> *.ini </c> files
    ''' </param>
    '''
    ''' <returns>
    ''' The merged <c> iniFile </c>, or an empty one if the directory does not exist
    ''' or contains no parseable sections
    ''' </returns>
    Private Function combineAppInfoDir(appInfoDir As String) As iniFile

        Dim combined = iniFile.Empty("", "")

        If Not Directory.Exists(appInfoDir) Then Return combined

        Dim files = Directory.GetFiles(appInfoDir, "*.ini", SearchOption.TopDirectoryOnly).ToList()
        files.Sort()

        For Each filePath In files

            Dim f = iniFile.FromFile(filePath)
            For Each section In f : combined.AddSection(section) : Next

        Next

        Return combined

    End Function

    ''' <summary>
    ''' Runs the phase-2 variable-expansion pass on <paramref name="template"/> and routes any
    ''' returned diagnostics onto <c> gLog </c> and <paramref name="menuOutput"/>. The caller is
    ''' responsible for phase-0 <c> %Package% </c> fan-out and phase-1 root substitution
    ''' beforehand. Running last means a root value's own <c> &lt;token&gt; </c> references are
    ''' already inlined into <paramref name="template"/> and participate in the same cartesian
    ''' product as the surrounding template's tokens, matching EntryBuilder's semantics.
    ''' </summary>
    '''
    ''' <param name="template">
    ''' One key value with phases 0 and 1 already applied, possibly containing
    ''' <c> &lt;var&gt; </c> tokens
    ''' </param>
    '''
    ''' <param name="app">
    ''' The entry whose <see cref="UWPAppInfo.Variables"/> set drives the expansion; tokens
    ''' observed update reference tracking for the post-generation typo backstop
    ''' </param>
    '''
    ''' <param name="domain">
    ''' The domain of the emitting key, determining undeclared-token behavior
    ''' </param>
    '''
    ''' <param name="keyLabel">
    ''' Short tag (e.g. <c> Detect </c>, <c> FileKey </c>) embedded into diagnostic messages
    ''' so a warning names the key it came from
    ''' </param>
    '''
    ''' <param name="menuOutput">
    ''' The <c> MenuSection </c> receiving Warning-severity diagnostics for display
    ''' </param>
    '''
    ''' <returns>
    ''' The fan-out values; empty when the template was dropped
    ''' </returns>
    Friend Function expandPhase2(template As String,
                                  app As UWPAppInfo,
                                  domain As ExpansionDomain,
                                  keyLabel As String,
                                  menuOutput As MenuSection) As List(Of String)

        Dim result = VariableExpander.Expand(template, app.Variables, domain, $"[{app.Name}].{keyLabel}")

        For Each diag In result.Diagnostics

            gLog(diag.Message)
            If diag.Severity = DiagnosticSeverity.Warning Then menuOutput.AddWarning(diag.Message)

        Next

        Return result.Values

    End Function

    ''' <summary>
    ''' Applies phase 0 (<c> %Package% </c> / <c> %PackageN% </c> fan-out) followed by phase 2
    ''' (<c> &lt;var&gt; </c> expansion) to every template in <paramref name="templates"/>,
    ''' returning the flattened result in fan-out order. Used for the key classes that carry no
    ''' root placeholder; the scaffold blocks interleave phase-1 substitution themselves.
    ''' </summary>
    '''
    ''' <param name="templates">
    ''' The key value templates to expand, in declaration order
    ''' </param>
    '''
    ''' <param name="app">
    ''' The entry supplying both the package list and the variable symbol table
    ''' </param>
    '''
    ''' <param name="domain">
    ''' The domain of the emitting key class
    ''' </param>
    '''
    ''' <param name="keyLabel">
    ''' Short tag embedded into diagnostic messages
    ''' </param>
    '''
    ''' <param name="menuOutput">
    ''' The <c> MenuSection </c> receiving Warning-severity diagnostics for display
    ''' </param>
    '''
    ''' <returns>
    ''' Every string produced by phase-0 × phase-2 expansion, in fan-out order
    ''' </returns>
    Friend Function expandPackageAndVars(templates As List(Of String),
                                          app As UWPAppInfo,
                                          domain As ExpansionDomain,
                                          keyLabel As String,
                                          menuOutput As MenuSection) As List(Of String)

        Dim values As New List(Of String)

        For Each template In templates

            For Each packageExpanded In expandPackageKey(template, app.Packages)

                values.AddRange(expandPhase2(packageExpanded, app, domain, keyLabel, menuOutput))

            Next

        Next

        Return values

    End Function

    ''' <summary>
    ''' Expands ExcludeKey templates through phases 0 and 2, classifying the expansion domain
    ''' per phase-0 result rather than per template: an ExcludeKey's flag determines whether its
    ''' value lands in the registry or the file system, and a <c> %Package% </c> fan-out can in
    ''' principle yield values whose flags differ. Mirrors EntryBuilder's <c> expandExcludeKeys </c>.
    ''' </summary>
    '''
    ''' <param name="templates">
    ''' The ExcludeKey value templates to expand, in declaration order
    ''' </param>
    '''
    ''' <param name="app">
    ''' The entry supplying both the package list and the variable symbol table
    ''' </param>
    '''
    ''' <param name="menuOutput">
    ''' The <c> MenuSection </c> receiving Warning-severity diagnostics for display
    ''' </param>
    '''
    ''' <returns>
    ''' Every expanded ExcludeKey value, in fan-out order
    ''' </returns>
    Friend Function expandExcludeKeys(templates As List(Of String),
                                       app As UWPAppInfo,
                                       menuOutput As MenuSection) As List(Of String)

        Dim values As New List(Of String)

        For Each template In templates

            For Each packageExpanded In expandPackageKey(template, app.Packages)

                Dim parsed As New excludeKeyParams(packageExpanded)
                Dim exclDomain As ExpansionDomain = If(parsed.Flag = excludeKeyFlag.Reg,
                                                       ExpansionDomain.Registry,
                                                       ExpansionDomain.Filesystem)

                values.AddRange(expandPhase2(packageExpanded, app, exclDomain, "ExcludeKey", menuOutput))

            Next

        Next

        Return values

    End Function

    ''' <summary>
    ''' Expands <c> %Package% </c> and <c> %PackageN% </c> variables in a key value template,
    ''' returning one expanded string per applicable package.
    ''' <br /><br />
    ''' 
    ''' Rules:
    ''' <list type="bullet">
    ''' 
    ''' <item>
    ''' <c> %PackageN% </c> (numbered): expands exactly once using the Nth package.
    ''' Only the lowest such N in the template is replaced, and any other package
    ''' reference in it, numbered or not, stays literal. An N past the package count
    ''' isn't recognized and stays literal too, without a warning.
    ''' </item>
    ''' 
    ''' <item>
    ''' <c> %Package% </c> (unnumbered): expands once per package, producing one
    ''' output string per package in order
    ''' </item>
    ''' 
    ''' <item>
    ''' No package variable: returned verbatim as a single-element list
    ''' </item>
    ''' 
    ''' </list>
    '''
    ''' Matching is case-sensitive, so <c> %package% </c> isn't expanded.
    ''' </summary>
    '''
    ''' <param name="template">
    ''' A key value string optionally containing <c> %Package% </c> or
    ''' <c> %PackageN% </c> variable references
    ''' </param>
    '''
    ''' <param name="packages">
    ''' The ordered list of package folder names for the current app
    ''' </param>
    '''
    ''' <returns>
    ''' One expanded string per applicable package, or a single-element list containing
    ''' <paramref name="template"/> verbatim if no package variable is present
    ''' </returns>
    Friend Function expandPackageKey(template As String,
                                      packages As List(Of String)) As List(Of String)

        Dim result As New List(Of String)

        ' Check for numbered package references (%Package1%, %Package2%, …)
        ' These appear only in AppInfo FileKey values as package selectors. scaffold templates
        ' always use unnumbered %Package%. A numbered reference selects exactly one package by position.
        For i = 1 To packages.Count

            Dim numberedVar = $"%Package{i}%"
            If Not template.Contains(numberedVar) Then Continue For

            result.Add(template.Replace(numberedVar, $"%LocalAppData%\Packages\{packages(i - 1)}"))
            Return result

        Next

        ' Unnumbered %Package%, expand for every package
        If template.Contains("%Package%") Then

            For Each pkg In packages
                result.Add(template.Replace("%Package%", $"%LocalAppData%\Packages\{pkg}"))
            Next

            Return result

        End If

        ' No package variable, pass through verbatim
        result.Add(template)
        Return result

    End Function

End Module
