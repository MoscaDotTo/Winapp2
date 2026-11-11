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
''' BrowserBuilder is a winapp2ool module which handles the generation of bespoke winapp2.ini
''' entries for a very large number of web browsers. It provides a small scripting interface
''' which allows for the dynamic creation of entire groups of winapp2.ini entries using a limited
''' set of easily accessed information. <br /><br />
'''
''' All input files are read from a single configurable source directory. At minimum, that
''' directory must contain at least one of <c> chromium.ini </c> or <c> gecko.ini </c>: the
''' specially formatted ruleset files that drive entry generation. BrowserBuilder is primarily
''' intended as a small devops tool but end users might find it useful as it enables them to
''' generate entries for non-standard installation paths or portable applications while keeping
''' winapp2.ini up to date separately <br /><br />
'''
''' Additionally, and because it is intended to simplify the devops, Browser Builder expects to
''' "Flavorize" the browsers.ini that it produces. This is done so as to resolve over and under
''' coverages by generated entries as well as resolve incompatibilities. Accordingly, it expects
''' but does not require a set of Flavor correction files in the source directory (up to 1 "Add"
''' file, up to 3 "Remove" files, and up to 2 "Replace" files), each found by its fixed filename.
''' Consult the Transmute documentation for more information about Flavors <br /><br />
''' Unflavored <c> browsers.ini </c> output may be broken or incomplete depending on use case
''' </summary>
Public Module BrowserBuilder

    ''' <summary>
    ''' The total number of entries generated for chromium-based browsers
    ''' </summary>
    Private Property totalChromiumCount As Integer = 0

    ''' <summary>
    ''' The total number of entries generated for gecko-based browsers
    ''' </summary>
    Private Property totalGeckoCount As Integer = 0

    ''' <summary>
    ''' The start of a BrowserInfo section's name
    ''' </summary>
    Private Const BrowserInfoPrefix As String = "BrowserInfo:"

    ''' <summary>
    ''' The start of an EntryScaffold section's name
    ''' </summary>
    Private Const EntryScaffoldPrefix As String = "EntryScaffold:"

    ''' <summary>
    ''' Binds the command-line file arguments to <see cref="BuilderFile1"/> and
    ''' <see cref="BuilderFile2"/>, starting from the default settings rather than any saved
    ''' ones, then runs the build
    ''' </summary>
    '''
    ''' <remarks>
    ''' <b> File arguments: </b> <br />
    ''' <c> -1d </c> <br /> Source directory containing <c> chromium.ini </c>,
    ''' <c> gecko.ini </c>, and flavor correction files 
    ''' <br /> <br /> 
    ''' <c> -2d </c> / <c> -2f </c> <br /> Output file path and name.
    ''' Defaults to <c> browsers.ini </c> in the current directory.
    ''' </remarks>
    Public Sub handleCmdLine()

        InitDefaultBrowserBuilderSettings()

        Dim spec As New CliArgSpec(NameOf(BrowserBuilder))
        spec.WithFile(1, BuilderFile1) _
            .WithFile(2, BuilderFile2) _
            .Parse()

        initBrowserBuilder()

    End Sub

    ''' <summary>
    ''' Reads <c> chromium.ini </c> and <c> gecko.ini </c> from <see cref="BuilderFile1"/>'s
    ''' directory, builds and writes the output file, and displays the results unless
    ''' <c> SuppressOutput </c> is set. A missing ruleset counts as an empty one. When neither
    ''' has any sections, which includes a missing source directory, we set a header message,
    ''' fail the run and return without writing.
    ''' </summary>
    Public Sub initBrowserBuilder()

        clrConsole()

        Dim sourceDir = BuilderFile1.Dir

        Dim chromiumIni = If(loadIfPresent(sourceDir, "chromium.ini"), iniFile.Empty(sourceDir, "chromium.ini"))
        Dim geckoIni = If(loadIfPresent(sourceDir, "gecko.ini"), iniFile.Empty(sourceDir, "gecko.ini"))
        Dim noRules = chromiumIni.Count + geckoIni.Count = 0

        If noRules Then

            Dim noRulesMsg = "No valid generative rulesets found"
            gLog($"{noRulesMsg} in {sourceDir}")
            setNextMenuHeaderText(noRulesMsg, printColor:=ConsoleColor.Red)
            markRunFailed()
            Return

        End If

        Dim browserBuilderStartPhrase = "Building browser configuration entries"
        Dim output As New MenuSection
        output.AddBoxWithText(browserBuilderStartPhrase)

        Using gLogScope(browserBuilderStartPhrase)

            processBrowserBuilder(chromiumIni, geckoIni, output)

            Dim browserBuilderEndPhrase = "Browser configuration entries built successfully"
            output.AddBoxWithText(browserBuilderEndPhrase)
            gLog(browserBuilderEndPhrase, buffr:=True)

        End Using

        output.AddAnyKeyPrompt()

        If Not SuppressOutput Then output.Print()
        crk()

    End Sub

    ''' <summary>
    ''' Generates the entries from both rulesets, Chromium first, then flavorizes them with
    ''' whichever correction files exist in the source directory, lints the result through
    ''' <see cref="remotedebugGuarded"/>, and writes it to <see cref="BuilderFile2"/> under a
    ''' <see cref="ArtifactHeader.Text"/>'s header. <see cref="Flavorize"/> sorts the entries and writes the file
    ''' once on its own, without the header, before we lint it and write it again.
    ''' </summary>
    '''
    ''' <param name="chromiumIni">
    ''' The parsed chromium.ini ruleset
    ''' </param>
    '''
    ''' <param name="geckoIni">
    ''' The parsed gecko.ini ruleset
    ''' </param>
    '''
    ''' <param name="output">
    ''' The <c> MenuSection </c> accumulating user-visible output for this run
    ''' </param>
    Private Sub processBrowserBuilder(chromiumIni As iniFile,
                                      geckoIni As iniFile,
                                      ByRef output As MenuSection)

        Using gLogScope("Processing browser builder files")

            Dim outputFile = iniFile.Empty(BuilderFile2.Dir, BuilderFile2.Name)

            buildScaffolds(chromiumIni, False, outputFile, output)
            totalChromiumCount = outputFile.Count

            buildScaffolds(geckoIni, True, outputFile, output)
            totalGeckoCount = outputFile.Count - totalChromiumCount

            Dim sourceDir = BuilderFile1.Dir

            Dim flavorAdd = loadIfPresent(sourceDir, "browser_additions.ini")
            Dim flavorSecRem = loadIfPresent(sourceDir, "browser_section_removals.ini")
            Dim flavorNameRem = loadIfPresent(sourceDir, "browser_name_removals.ini")
            Dim flavorValRem = loadIfPresent(sourceDir, "browser_value_removals.ini")
            Dim flavorSecRep = loadIfPresent(sourceDir, "browser_section_replacements.ini")
            Dim flavorKeyRep = loadIfPresent(sourceDir, "browser_key_replacements.ini")

            Flavorize(outputFile, outputFile, output, flavorAdd, flavorSecRem, flavorNameRem, flavorValRem, flavorSecRep, flavorKeyRep)

            outputFile = remotedebugGuarded(outputFile, NameOf(BrowserBuilder), output)

            Dim header = ArtifactHeader.Text(outputFile.Name, NameOf(BrowserBuilder), outputFile.Count, BrowserBuilderSources, sourceIsFolder:=True)
            outputFile.OverwriteToFile(header & outputFile.ToString())

            gLog("Browser builder files processed successfully")

        End Using

    End Sub

    ''' <summary>
    ''' Sorts the sections of <paramref name="rulesetFile"/> into <c> BrowserInfo: </c> and
    ''' <c> EntryScaffold: </c> sections, then generates one entry per scaffold for each
    ''' browser that isn't skipped. We parse every browser before any scaffold, so their order
    ''' in the file doesn't matter. Any other section, including one with no name after its
    ''' prefix, is ignored with a warning and fails the run.
    ''' </summary>
    '''
    ''' <param name="rulesetFile">
    ''' The <c> iniFile </c> containing the set of generative rules for a particular group
    ''' of web browsers
    ''' </param>
    '''
    ''' <param name="isGecko">
    ''' Indicates whether the browsers in <paramref name="rulesetFile"/> are Gecko-based
    ''' </param>
    '''
    ''' <param name="outputFile">
    ''' The file receiving the generated entries. A generated name it already holds is dropped.
    ''' </param>
    '''
    ''' <param name="menuOutput">
    ''' The <c> MenuSection </c> accumulating user-visible output for this run
    ''' </param>
    Private Sub buildScaffolds(rulesetFile As iniFile,
                               isGecko As Boolean,
                         ByRef outputFile As iniFile,
                         ByRef menuOutput As MenuSection)

        Dim browsers As New List(Of BrowserInfo)
        Dim scaffoldSections As New List(Of iniSection)

        For Each section In rulesetFile

            Select Case True

                Case nameAfterPrefix(section.Name, BrowserInfoPrefix).Length > 0

                    Dim browserInfo As BrowserInfo = parseBrowserInfo(section, menuOutput)
                    If browserInfo.ShouldSkip Then Continue For
                    browsers.Add(browserInfo)

                Case nameAfterPrefix(section.Name, EntryScaffoldPrefix).Length > 0 : scaffoldSections.Add(section)

                Case Else

                    Dim logMsg = $"Invalid section found and ignored: [{section.Name}]; a section is named `{BrowserInfoPrefix} <name>` or `{EntryScaffoldPrefix} <name>`"
                    gLog(logMsg)
                    menuOutput.AddWarning(logMsg)
                    markRunFailed()

            End Select

        Next

        Dim configCount = $"Found {browsers.Count} browser configurations"
        menuOutput.AddColoredLine(configCount, ConsoleColor.Yellow)
        gLog($"  {configCount}")

        For Each scaffoldSection In scaffoldSections

            processEntryScaffold(scaffoldSection, browsers, isGecko, outputFile, menuOutput)

        Next

    End Sub

    ''' <summary>
    ''' Parses a BrowserInfo section into a <see cref="BrowserInfo"/>. The browser's name is
    ''' the header after <c> BrowserInfo: </c>, trimmed. A missing <c> UserDataPath= </c> or
    ''' <c> Section= </c> warns but doesn't skip the browser. <c> TruncateDetect </c> and
    ''' <c> Skip </c> take effect by being present, whatever their value. Any other key
    ''' declares one of the browser's variables (see <see cref="declareVariable"/>). Once
    ''' every key is read we resolve the variables, then expand each <c> UserDataPath= </c> and
    ''' <c> RegistryRoot= </c> value into plain paths (see <see cref="expandBrowserValue"/>),
    ''' so nothing after this sees a variable. An empty value warns and is skipped. A path with
    ''' no parent folder for <c> &lt;UserDataParent&gt; </c> is skipped with a warning and fails
    ''' the run.
    ''' </summary>
    '''
    ''' <param name="browserSection">
    ''' The <c> iniSection </c> containing the BrowserInfo data
    ''' </param>
    '''
    ''' <param name="menuOutput">
    ''' The <c> MenuSection </c> accumulating user-visible output for this run
    ''' </param>
    '''
    ''' <returns>
    ''' A <c> BrowserInfo </c> structure containing all parsed browser parameters
    ''' </returns>
    Friend Function parseBrowserInfo(browserSection As iniSection,
                               ByRef menuOutput As MenuSection) As BrowserInfo

        Dim browserName As String = nameAfterPrefix(browserSection.Name, BrowserInfoPrefix)
        Dim browserInfo As New BrowserInfo(browserName)
        Dim variables As New VariableSet
        Dim pathKeys As New List(Of iniKey)
        Dim rootKeys As New List(Of iniKey)

        For Each key In browserSection.Keys

            Select Case key.TypeOrName().ToUpperInvariant()

                Case "USERDATAPATH" : pathKeys.Add(key)

                Case "REGISTRYROOT" : rootKeys.Add(key)

                Case "SECTION" : browserInfo.SectionName = key.Value

                Case "TRUNCATEDETECT" : browserInfo.TruncateDetect = True

                Case "SKIP" : browserInfo.ShouldSkip = True

                Case Else : declareVariable(key, browserSection.Name, variables, menuOutput)

            End Select

        Next

        reportResolveDiagnostics(VariableExpander.ResolveAll(variables), browserSection.Name, menuOutput)

        For Each key In pathKeys

            For Each path In expandBrowserValue(key, browserSection.Name, variables, ExpansionDomain.Filesystem, menuOutput)

                Dim lastSlash = path.LastIndexOf("\"c)

                If lastSlash <= 0 Then

                    Dim noParentMsg = $"[{browserSection.Name}].{key.Name} gives `{path}`, which has no parent folder for <UserDataParent>; skipping it"
                    gLog(noParentMsg)
                    menuOutput.AddWarning(noParentMsg)
                    markRunFailed()
                    Continue For

                End If

                browserInfo.UserDataPaths.Add(path)
                browserInfo.UserDataParentPaths.Add(path.Substring(0, lastSlash))

            Next

        Next

        For Each key In rootKeys

            browserInfo.RegistryRoots.AddRange(expandBrowserValue(key, browserSection.Name, variables, ExpansionDomain.Registry, menuOutput))

        Next

        warnUnreferenced(variables, browserSection.Name, menuOutput)

        Dim noUserData = browserInfo.UserDataPaths.Count = 0
        Dim noSection = browserInfo.SectionName = ""
        Dim noUserDataErr = $"No valid UserDataPath key found in {browserName}"
        Dim noValidSectionErr = $"No valid Section key found in {browserName}"

        menuOutput.AddWarning(noUserDataErr, condition:=noUserData)
        menuOutput.AddWarning(noValidSectionErr, condition:=noSection)
        gLog(noUserDataErr, cond:=noUserData)
        gLog(noValidSectionErr, cond:=noSection)

        Return browserInfo

    End Function

    ''' <summary>
    ''' Returns the plain values a BrowserInfo <c> UserDataPath= </c> or
    ''' <c> RegistryRoot= </c> key expands to over the browser's variables and inline lists,
    ''' one per combination. A value <see cref="findBrowserValueError"/> or
    ''' <see cref="expansionError"/> refuses is skipped with a warning and fails the run, and
    ''' we return nothing for it.
    ''' </summary>
    '''
    ''' <param name="key">
    ''' The <c> UserDataPath= </c> or <c> RegistryRoot= </c> key
    ''' </param>
    '''
    ''' <param name="browserSectionName">
    ''' The full name of the BrowserInfo section, used in messages
    ''' </param>
    '''
    ''' <param name="variables">
    ''' The browser's declared variables, already resolved
    ''' </param>
    '''
    ''' <param name="domain">
    ''' <c> Filesystem </c> for a user data path, <c> Registry </c> for a registry root
    ''' </param>
    '''
    ''' <param name="menuOutput">
    ''' The <c> MenuSection </c> accumulating user-visible output for this run
    ''' </param>
    Private Function expandBrowserValue(key As iniKey,
                                        browserSectionName As String,
                                        variables As VariableSet,
                                        domain As ExpansionDomain,
                                        menuOutput As MenuSection) As List(Of String)

        Dim context = $"[{browserSectionName}].{key.Name}"

        If key.Value.Trim().Length = 0 Then

            Dim emptyMsg = $"{context} has no value; skipping it"
            gLog(emptyMsg)
            menuOutput.AddWarning(emptyMsg)
            Return New List(Of String)

        End If

        Dim valueError = findBrowserValueError(key.Value, variables, context)
        Dim result As ExpansionResult = Nothing

        If valueError.Length = 0 Then

            result = VariableExpander.Expand(key.Value, variables, domain, context)
            valueError = expansionError(result, context)

        End If

        If valueError.Length > 0 Then

            Dim skipMsg = $"{valueError}; skipping `{key.Name}={key.Value}`"
            gLog(skipMsg)
            menuOutput.AddWarning(skipMsg)
            markRunFailed()
            Return New List(Of String)

        End If

        For Each diag In result.Diagnostics

            gLog(diag.Message)

        Next

        Return result.Values

    End Function

    ''' <summary>
    ''' Declares <paramref name="key"/> as a variable named after its key type, with its
    ''' comma-separated value as the list. The key type drops trailing digits, so <c> Dirs1= </c>
    ''' and <c> Dirs2= </c> both declare <c> Dirs </c>. A second declaration replaces the first
    ''' with a warning. A key named after one of the four scaffold tokens is dropped with a
    ''' warning and fails the run.
    ''' </summary>
    '''
    ''' <param name="key">
    ''' A key that isn't one of the section's reserved keys
    ''' </param>
    '''
    ''' <param name="sectionName">
    ''' The full name of the section, used in the message
    ''' </param>
    '''
    ''' <param name="variables">
    ''' The section's variables
    ''' </param>
    '''
    ''' <param name="menuOutput">
    ''' The <c> MenuSection </c> accumulating user-visible output for this run
    ''' </param>
    Private Sub declareVariable(key As iniKey,
                                sectionName As String,
                                variables As VariableSet,
                                menuOutput As MenuSection)

        Dim name = key.TypeOrName()

        If isTokenName(name) Then

            Dim dropMsg = $"[{sectionName}] declares `{name}`, which is a token, not a variable; dropping `{key.Name}={key.Value}`"
            gLog(dropMsg)
            menuOutput.AddWarning(dropMsg)
            markRunFailed()
            Return

        End If

        If variables.IsDeclared(name) Then

            Dim redeclaredMsg = $"[{sectionName}] declares `{name}` more than once; `{key.Name}={key.Value}` replaces the earlier values"
            gLog(redeclaredMsg)
            menuOutput.AddWarning(redeclaredMsg)

        End If

        variables.Add(name, key.Value)

    End Sub

    ''' <summary>
    ''' Logs the diagnostics from resolving a section's variables against each other. A
    ''' warning, such as a cycle, is shown and fails the run.
    ''' </summary>
    '''
    ''' <param name="diagnostics">
    ''' What <see cref="VariableExpander.ResolveAll"/> returned
    ''' </param>
    '''
    ''' <param name="sectionName">
    ''' The full name of the section, added to each message
    ''' </param>
    '''
    ''' <param name="menuOutput">
    ''' The <c> MenuSection </c> accumulating user-visible output for this run
    ''' </param>
    Private Sub reportResolveDiagnostics(diagnostics As List(Of ExpansionDiagnostic),
                                         sectionName As String,
                                         menuOutput As MenuSection)

        For Each diag In diagnostics

            Dim msg = $"[{sectionName}] {diag.Message}"
            gLog(msg)
            If diag.Severity <> DiagnosticSeverity.Warning Then Continue For

            menuOutput.AddWarning(msg)
            markRunFailed()

        Next

    End Sub

    ''' <summary>
    ''' Warns about each variable declared in a section that nothing references, which is
    ''' how a misspelled reserved key shows up
    ''' </summary>
    '''
    ''' <param name="variables">
    ''' The section's variables, after every value or template that could use them has
    ''' </param>
    '''
    ''' <param name="sectionName">
    ''' The full name of the section, used in the message
    ''' </param>
    '''
    ''' <param name="menuOutput">
    ''' The <c> MenuSection </c> accumulating user-visible output for this run
    ''' </param>
    Private Sub warnUnreferenced(variables As VariableSet,
                                 sectionName As String,
                                 menuOutput As MenuSection)

        For Each unused In variables.UnreferencedNames()

            Dim unusedMsg = $"Variable '{unused}=' declared in [{sectionName}] but never referenced by any <{unused}> token; possible typo"
            gLog(unusedMsg)
            menuOutput.AddWarning(unusedMsg)

        Next

    End Sub

    ''' <summary>
    ''' Returns an <c> iniFile </c> loaded from the given directory and filename,
    ''' or <c> Nothing </c> if the file does not exist. We check first because
    ''' <see cref="iniFile.FromFile"/> fails the run on a missing file.
    ''' </summary>
    '''
    ''' <param name="dir">The directory containing the file</param>
    '''
    ''' <param name="name">The filename</param>
    '''
    ''' <returns>
    ''' A loaded <c> iniFile </c>, or <c> Nothing </c> if the file is absent
    ''' </returns>
    Private Function loadIfPresent(dir As String, name As String) As iniFile

        Dim path = $"{dir}\{name}"
        Return If(IO.File.Exists(path), iniFile.FromFile(path), Nothing)

    End Function

    ''' <summary>
    ''' Returns the trimmed name that follows <paramref name="prefix"/> in
    ''' <paramref name="sectionName"/>, or an empty string when the section doesn't start with
    ''' the prefix or names nothing after it. The prefix matches in any case, and a space after
    ''' its colon is optional.
    ''' </summary>
    '''
    ''' <param name="sectionName">
    ''' The name of a ruleset section
    ''' </param>
    '''
    ''' <param name="prefix">
    ''' <see cref="BrowserInfoPrefix"/> or <see cref="EntryScaffoldPrefix"/>
    ''' </param>
    Friend Function nameAfterPrefix(sectionName As String, prefix As String) As String

        If Not sectionName.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) Then Return ""

        Return sectionName.Substring(prefix.Length).Trim()

    End Function

    ''' <summary>
    ''' Processes an EntryScaffold section and generates entries for each browser.
    ''' Browsers with no <c> RegistryRoot </c> are skipped, with a warning, when the scaffold
    ''' contains a <c> RequiresRegistryRoot </c> key. Any key other than
    ''' <c> FileKeyBase </c>, <c> RegKeyBase </c> and <c> RequiresRegistryRoot </c> declares
    ''' one of the scaffold's variables (see <see cref="declareVariable"/>), wherever it sits
    ''' in the section, and we resolve them before checking any template.
    ''' <br /><br />
    '''
    ''' We check each template once here, before any browser sees it, with
    ''' <see cref="findTemplateError"/>. A template it refuses (a retired <c> %...% </c> token,
    ''' an unknown <c> &lt;name&gt; </c> token, a misplaced or missing token) is dropped with a
    ''' warning and fails the run through <see cref="markRunFailed"/>, and the scaffold's other
    ''' templates still build.
    ''' </summary>
    '''
    ''' <param name="scaffoldSection">
    ''' A particular <c> EntryScaffold </c> section to be generated for each browser
    ''' </param>
    '''
    ''' <param name="browsers">
    ''' The set of all BrowserInfo sections from the ruleset file
    ''' </param>
    '''
    ''' <param name="isGecko">
    ''' Indicates whether the browsers are Gecko-based
    ''' </param>
    '''
    ''' <param name="outputFile">
    ''' The file receiving the generated entries
    ''' </param>
    '''
    ''' <param name="menuOutput">
    ''' The <c> MenuSection </c> accumulating user-visible output for this run
    ''' </param>
    Friend Sub processEntryScaffold(scaffoldSection As iniSection,
                                    browsers As List(Of BrowserInfo),
                                    isGecko As Boolean,
                              ByRef outputFile As iniFile,
                              ByRef menuOutput As MenuSection)

        Dim scaffoldName As String = nameAfterPrefix(scaffoldSection.Name, EntryScaffoldPrefix)

        Dim processingMsg = $"Processing EntryScaffold: {scaffoldName}"
        menuOutput.AddColoredLine(processingMsg, ConsoleColor.Magenta)

        Using gLogScope(processingMsg)

            Dim fileKeyBases As New List(Of String)
            Dim regKeyBases As New List(Of String)
            Dim templates As New List(Of iniKey)
            Dim variables As New VariableSet

            Dim requiresRegistryRoot As Boolean = False

            For Each key In scaffoldSection.Keys

                Select Case key.TypeOrName().ToUpperInvariant()

                    Case "FILEKEYBASE", "REGKEYBASE" : templates.Add(key)

                    Case "REQUIRESREGISTRYROOT" : requiresRegistryRoot = True

                    Case Else : declareVariable(key, scaffoldSection.Name, variables, menuOutput)

                End Select

            Next

            reportResolveDiagnostics(VariableExpander.ResolveAll(variables), scaffoldSection.Name, menuOutput)

            For Each template In templates

                Dim isRegKey = template.TypeOrName().Equals("RegKeyBase", StringComparison.OrdinalIgnoreCase)
                If Not templateIsUsable(template, isRegKey, isGecko, scaffoldSection.Name, variables, menuOutput) Then Continue For

                If isRegKey Then

                    regKeyBases.Add(template.Value)

                Else

                    fileKeyBases.Add(template.Value)

                End If

            Next

            warnUnreferenced(variables, scaffoldSection.Name, menuOutput)

            For Each browser In browsers

                If requiresRegistryRoot AndAlso browser.RegistryRoots.Count = 0 Then

                    Dim skipMsg = $"Skipping {browser.Name} for {scaffoldName}: scaffold requires RegistryRoot"
                    gLog(skipMsg)
                    menuOutput.AddWarning(skipMsg)
                    Continue For

                End If

                generateBrowserEntry(browser, scaffoldName, fileKeyBases, regKeyBases, variables, outputFile, isGecko, menuOutput)

            Next

            Dim finishedMsg = $"Finished processing EntryScaffold: {scaffoldName}"
            menuOutput.AddColoredLine(finishedMsg, ConsoleColor.Yellow)
            gLog(finishedMsg, buffr:=True)

        End Using

    End Sub

    ''' <summary>
    ''' Returns whether a <c> FileKeyBase= </c> or <c> RegKeyBase= </c> template passes
    ''' <see cref="findTemplateError"/>. When it doesn't, we log and show why, name the dropped
    ''' key, and fail the run.
    ''' </summary>
    '''
    ''' <param name="key">
    ''' The template key
    ''' </param>
    '''
    ''' <param name="isRegKey">
    ''' Indicates whether <paramref name="key"/> is a <c> RegKeyBase= </c>
    ''' </param>
    '''
    ''' <param name="isGecko">
    ''' Indicates whether the scaffold belongs to the Gecko ruleset
    ''' </param>
    '''
    ''' <param name="scaffoldSectionName">
    ''' The full name of the scaffold's section, used in the message
    ''' </param>
    '''
    ''' <param name="variables">
    ''' The scaffold's declared variables, already resolved
    ''' </param>
    '''
    ''' <param name="menuOutput">
    ''' The <c> MenuSection </c> accumulating user-visible output for this run
    ''' </param>
    Private Function templateIsUsable(key As iniKey,
                                      isRegKey As Boolean,
                                      isGecko As Boolean,
                                      scaffoldSectionName As String,
                                      variables As VariableSet,
                                      menuOutput As MenuSection) As Boolean

        Dim templateError = findTemplateError(key.Value, isRegKey, isGecko, variables, $"[{scaffoldSectionName}].{key.Name}")
        If templateError.Length = 0 Then Return True

        Dim dropMsg = $"{templateError}; dropping `{key.Name}={key.Value}`"
        gLog(dropMsg)
        menuOutput.AddWarning(dropMsg)
        markRunFailed()

        Return False

    End Function

End Module
