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
''' <c> FileKeyBase= </c> lines are applied to every application as a consistent
''' baseline set of cleaning targets.
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
''' <c> %LocalAppData%\Packages\&lt;PACKAGE_FOLDER&gt; </c>.
''' Apps with multiple packages use numbered variants:
''' <c> Package1= </c> / <c> Package2= </c> in the source and
''' <c> %Package1% </c> / <c> %Package2% </c> in key values.
''' </summary>
Public Module UWPBuilder

    ''' <summary>
    ''' Stores the parsed information for a single UWP application entry
    ''' </summary>
    Private Structure UWPAppInfo

        ''' <summary>
        ''' The entry name — used verbatim as the section header in the output
        ''' </summary>
        Public Name As String

        ''' <summary>
        ''' The package folder names under <c> %LocalAppData%\Packages </c>.
        ''' Single-package apps use <c> Package= </c>; multi-package apps use
        ''' <c> Package1= </c>, <c> Package2= </c>, etc.
        ''' </summary>
        Public Packages As List(Of String)

        ''' <summary>
        ''' The LangSecRef= value, or empty if Section= is used
        ''' </summary>
        Public LangSecRef As String

        ''' <summary>
        ''' The Section= value, or empty if LangSecRef= is used
        ''' </summary>
        Public SectionName As String

        ''' <summary>
        ''' Detect= registry key values in their original file order, passed through verbatim
        ''' and renumbered from 1. Supports hybrid win32+UWP entries that need multiple
        ''' detection conditions.
        ''' </summary>
        Public DetectKeys As List(Of String)

        ''' <summary>
        ''' Additional DetectFile= values in their original file order, appended after the
        ''' scaffold-generated DetectFile keys. Supports hybrid win32+UWP entries that need
        ''' file system detection for the win32 installation alongside package detection.
        ''' </summary>
        Public DetectFileKeys As List(Of String)

        ''' <summary>
        ''' App-specific FileKey and FileKeyBase values in their original file order.
        ''' Values containing <c> %Package% </c> or <c> %PackageN% </c> are expanded
        ''' during entry generation; all others are passed through verbatim.
        ''' </summary>
        Public AppKeys As List(Of String)

        ''' <summary>
        ''' RegKey= values in their original file order, passed through verbatim
        ''' and renumbered from 1
        ''' </summary>
        Public RegKeys As List(Of String)

        ''' <summary>
        ''' <c> ExcludeKey= </c> and <c> ExcludeKeyBase= </c> values in their original file order.
        ''' Values containing <c> %Package% </c> or <c> %PackageN% </c> are expanded
        ''' during entry generation; all others are passed through verbatim.
        ''' </summary>
        Public ExcludeKeys As List(Of String)

        ''' <summary>
        ''' When True, this entry is omitted from generation
        ''' </summary>
        Public ShouldSkip As Boolean

        ''' <summary>
        ''' Creates a new <c> UWPAppInfo </c> for an entry with the given name,
        ''' initialising all list fields to empty collections
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
            ShouldSkip = False

        End Sub

    End Structure

    ''' <summary>
    ''' Handles the command-line arguments for <c>UWPBuilder</c>
    ''' </summary>
    Public Sub handleCmdLine()

        getFileAndDirParams({UWPFile1, UWPFile2})
        initUWPBuilder()

    End Sub

    ''' <summary>
    ''' Reads the source directory, runs the generation pipeline, writes the output file,
    ''' and displays the results. Bails early with a header message if the template or
    ''' app definitions are missing.
    ''' </summary>
    Public Sub initUWPBuilder()

        clrConsole()

        Dim sourceDir = UWPFile1.Dir
        Dim templateIni = iniFile2.FromFile($"{sourceDir}\UWP.ini")

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

        Dim startPhrase = "Building UWP app entries"
        Dim output As New MenuSection
        output.AddBoxWithText(startPhrase)
        gLog(startPhrase, buffr:=True, ascend:=True)

        processUWPBuilder(templateIni, appsIni, output)

        Dim endPhrase = "UWP app entries built successfully"
        output.AddBoxWithText(endPhrase)
        gLog(endPhrase, descend:=True, buffr:=True)

        output.AddAnyKeyPrompt()

        If Not SuppressOutput Then output.Print()
        crk()

    End Sub

    ''' <summary>
    ''' Orchestrates the UWP builder process: parses scaffold templates and app definitions,
    ''' generates one <c> iniSection2 </c> per app, serialises the result with a header
    ''' comment block, and writes it to the output file.
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
    ''' <param name="menuOutput">
    ''' The <c> MenuSection </c> receiving progress lines and warnings for display
    ''' </param>
    Private Sub processUWPBuilder(templateIni As iniFile2,
                                   appsIni As iniFile2,
                                   menuOutput As MenuSection)

        gLog("Processing UWP builder files", ascend:=True, buffr:=True)

        Dim scaffoldFileKeys As New List(Of String)
        Dim scaffoldDetectFiles As New List(Of String)
        Dim apps As New List(Of UWPAppInfo)

        For Each section In templateIni

            If section.Name.StartsWith("EntryScaffold:", StringComparison.InvariantCulture) Then

                parseScaffold(section, scaffoldFileKeys, scaffoldDetectFiles, menuOutput)

            Else

                Dim logMsg = $"Unexpected section in template file: [{section.Name}]"
                gLog(logMsg)
                menuOutput.AddWarning(logMsg)

            End If

        Next

        Dim scaffoldMsg = $"Loaded {scaffoldDetectFiles.Count} scaffold DetectFile template(s), {scaffoldFileKeys.Count} scaffold FileKey template(s)"
        menuOutput.AddColoredLine(scaffoldMsg, ConsoleColor.Yellow)
        gLog(scaffoldMsg, indent:=True)

        For Each section In appsIni

            Dim app As UWPAppInfo = parseAppInfo(section, menuOutput)
            If Not app.ShouldSkip Then apps.Add(app)

        Next

        Dim appsMsg = $"Loaded {apps.Count} app definition(s)"
        menuOutput.AddColoredLine(appsMsg, ConsoleColor.Yellow)
        gLog(appsMsg, indent:=True)

        Dim outputFile = iniFile2.Empty(UWPFile2.Dir, UWPFile2.Name)

        For Each app In apps

            Dim entrySection = generateUWPEntry(app, scaffoldFileKeys, scaffoldDetectFiles, menuOutput)
            outputFile.AddSection(entrySection)

        Next

        Dim generatedMsg = $"Generated {outputFile.Count} UWP entries"
        menuOutput.AddColoredLine(generatedMsg, ConsoleColor.Yellow)
        gLog(generatedMsg, indent:=True)

        Dim sb As New StringBuilder()
        sb.AppendLine($"; Version {DateTime.Now.ToString("yyMMdd")}")
        sb.AppendLine($"; # of entries: {outputFile.Count:#,###}")
        sb.AppendLine($"; {UWPFile2.Name} is generated by the Winapp2ool UWP Builder")
        sb.AppendLine("; Entries in this file may be incomplete and are not intended to be used directly with any cleaning software")
        sb.AppendLine("; They are utilized by winapp2ool to create the final winapp2.ini file for distribution")
        sb.AppendLine("; If you are not maintaining winapp2.ini for distribution, you probably don't need this file!")
        sb.AppendLine("; Refer to the Winapp2ool documentation for more information: " & readMeUrl)
        sb.AppendLine()
        sb.Append(outputFile.ToString())

        outputFile.OverwriteToFile(sb.ToString())

        gLog("UWP builder files processed successfully", descend:=True)

    End Sub

    ''' <summary>
    ''' Parses an <c> [EntryScaffold: ...] </c> section from the template file, collecting
    ''' its <c> DetectFileBase= </c> values into <paramref name="scaffoldDetectFiles"/> and
    ''' its <c> FileKeyBase= </c> values into <paramref name="scaffoldFileKeys"/>.
    ''' Warns on any unrecognised key types.
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
    Private Sub parseScaffold(scaffoldSection As iniSection2,
                               scaffoldFileKeys As List(Of String),
                               scaffoldDetectFiles As List(Of String),
                               menuOutput As MenuSection)

        Dim scaffoldName = scaffoldSection.Name.Substring("EntryScaffold:".Length).Trim()
        Dim scaffoldMsg = $"Processing scaffold: {scaffoldName}"
        menuOutput.AddColoredLine(scaffoldMsg, ConsoleColor.Magenta)
        gLog(scaffoldMsg, indent:=True)

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
    ''' Parses an AppInfo section and returns a populated <c> UWPAppInfo </c> structure.
    ''' Issues warnings and sets <c> ShouldSkip </c> for entries that are structurally
    ''' invalid (missing package, missing category, or conflicting categories).
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
    Private Function parseAppInfo(appSection As iniSection2,
                                   menuOutput As MenuSection) As UWPAppInfo

        Dim app As New UWPAppInfo(appSection.Name)

        For Each key In appSection.Keys

            Select Case key.KeyType.ToUpperInvariant()

                Case "PACKAGE" : app.Packages.Add(key.Value)

                Case "LANGSECREF" : app.LangSecRef = key.Value

                Case "SECTION" : app.SectionName = key.Value

                Case "DETECT" : app.DetectKeys.Add(key.Value)

                Case "DETECTFILE" : app.DetectFileKeys.Add(key.Value)

                Case "FILEKEYBASE", "FILEKEY" : app.AppKeys.Add(key.Value)

                Case "REGKEY" : app.RegKeys.Add(key.Value)

                Case "EXCLUDEKEY", "EXCLUDEKEYBASE" : app.ExcludeKeys.Add(key.Value)

                Case "SKIP" : app.ShouldSkip = True

                Case Else

                    Dim errMsg = $"Unexpected key type in [{app.Name}]: {key.Name}"
                    gLog(errMsg)
                    menuOutput.AddWarning(errMsg)

            End Select

        Next

        Dim noPackage = app.Packages.Count = 0
        Dim noCategory = app.LangSecRef.Length = 0 AndAlso app.SectionName.Length = 0
        Dim bothCategories = app.LangSecRef.Length > 0 AndAlso app.SectionName.Length > 0

        If noPackage Then

            Dim msg = $"No Package key in [{app.Name}] — skipping"
            gLog(msg)
            menuOutput.AddWarning(msg)
            app.ShouldSkip = True

        End If

        If noCategory Then

            Dim msg = $"No LangSecRef or Section in [{app.Name}] — skipping"
            gLog(msg)
            menuOutput.AddWarning(msg)
            app.ShouldSkip = True

        End If

        If bothCategories Then

            Dim msg = $"Both LangSecRef and Section present in [{app.Name}] — using LangSecRef"
            gLog(msg)
            menuOutput.AddWarning(msg)
            app.SectionName = ""

        End If

        Return app

    End Function

    ''' <summary>
    ''' Generates a winapp2.ini entry section for the given <c> UWPAppInfo </c> by applying
    ''' the scaffold template keys and app-specific keys, expanding any
    ''' <c> %Package% </c> / <c> %PackageN% </c> variables along the way.
    ''' Keys are emitted in winapp2.ini order: category, Detect, DetectFile,
    ''' FileKey, RegKey, ExcludeKey.
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
    ''' <param name="menuOutput">
    ''' The <c> MenuSection </c> receiving progress lines for display
    ''' </param>
    '''
    ''' <returns>
    ''' A fully populated <c> iniSection2 </c> ready to be added to the output file
    ''' </returns>
    Private Function generateUWPEntry(app As UWPAppInfo,
                                      scaffoldFileKeys As List(Of String),
                                      scaffoldDetectFiles As List(Of String),
                                      menuOutput As MenuSection) As iniSection2

        Dim generatingMsg = $"Generating entry: {app.Name}"
        menuOutput.AddColoredLine(generatingMsg, ConsoleColor.Magenta)
        gLog(generatingMsg, indent:=True)

        Dim section As New iniSection2(app.Name)
        Dim fileKeyNum As Integer = 1

        ' 1. Category
        Select Case True

            Case app.LangSecRef.Length > 0 : section.AddKey(New iniKey2($"LangSecRef={app.LangSecRef}"))

            Case app.SectionName.Length > 0 : section.AddKey(New iniKey2($"Section={app.SectionName}"))

            Case Else : gLog($"{app.Name} has no category key")

        End Select

        ' 2. Detect keys — unnumbered if single, numbered from 1 if multiple
        If app.DetectKeys.Count = 1 Then

            section.AddKey(New iniKey2($"Detect={app.DetectKeys(0)}"))

        Else

            For i = 0 To app.DetectKeys.Count - 1
                section.AddKey(New iniKey2($"Detect{i + 1}={app.DetectKeys(i)}"))
            Next

        End If

        ' 3. DetectFile keys — scaffold templates expanded per package first, then app-specific, unnumbered if only one total
        Dim allDetectFiles As New List(Of String)
        For Each template In scaffoldDetectFiles

            allDetectFiles.AddRange(expandPackageKey(template, app.Packages))

        Next
        allDetectFiles.AddRange(app.DetectFileKeys)

        If allDetectFiles.Count = 1 Then

            section.AddKey(New iniKey2($"DetectFile={allDetectFiles(0)}"))

        Else

            For i = 0 To allDetectFiles.Count - 1
                section.AddKey(New iniKey2($"DetectFile{i + 1}={allDetectFiles(i)}"))
            Next

        End If

        ' 4. Scaffold template FileKeys — applied to all packages
        For Each scaffoldKey In scaffoldFileKeys

            For Each expanded In expandPackageKey(scaffoldKey, app.Packages)

                section.AddKey(New iniKey2($"FileKey{fileKeyNum}={expanded}"))
                fileKeyNum += 1

            Next

        Next

        ' 5. App-specific FileKey / FileKeyBase values in document order
        For Each appKey In app.AppKeys

            For Each expanded In expandPackageKey(appKey, app.Packages)

                section.AddKey(New iniKey2($"FileKey{fileKeyNum}={expanded}"))
                fileKeyNum += 1

            Next

        Next

        ' 6. RegKeys — passed through verbatim, renumbered from 1
        Dim regKeyNum As Integer = 1
        For Each regKey In app.RegKeys

            section.AddKey(New iniKey2($"RegKey{regKeyNum}={regKey}"))
            regKeyNum += 1

        Next

        ' 7. ExcludeKeys — expanded for %Package% / %PackageN%, renumbered from 1
        Dim exclNum As Integer = 1
        For Each exclKey In app.ExcludeKeys

            For Each expanded In expandPackageKey(exclKey, app.Packages)

                section.AddKey(New iniKey2($"ExcludeKey{exclNum}={expanded}"))
                exclNum += 1

            Next

        Next

        Dim generatedMsg = $"Generated entry: {app.Name}"
        menuOutput.AddColoredLine(generatedMsg, ConsoleColor.Yellow)
        gLog(generatedMsg, indent:=True, indAmt:=4)

        Return section

    End Function

    ''' <summary>
    ''' Combines all <c> *.ini </c> files in <paramref name="appInfoDir"/> into a single
    ''' in-memory <c> iniFile2 </c>. Files are processed in alphabetical order.
    ''' Sections with duplicate names across files are silently ignored (first-file-wins).
    ''' </summary>
    '''
    ''' <param name="appInfoDir">
    ''' Path to the <c> AppInfo\ </c> directory containing per-letter <c> *.ini </c> files
    ''' </param>
    '''
    ''' <returns>
    ''' The merged <c> iniFile2 </c>, or an empty one if the directory does not exist
    ''' or contains no parseable sections
    ''' </returns>
    Private Function combineAppInfoDir(appInfoDir As String) As iniFile2

        Dim combined = iniFile2.Empty("", "")

        If Not Directory.Exists(appInfoDir) Then Return combined

        Dim files = Directory.GetFiles(appInfoDir, "*.ini", SearchOption.TopDirectoryOnly).ToList()
        files.Sort()

        For Each filePath In files

            Dim f = iniFile2.FromFile(filePath)
            For Each section In f : combined.AddSection(section) : Next

        Next

        Return combined

    End Function

    ''' <summary>
    ''' Expands <c>%Package%</c> and <c>%PackageN%</c> variables in a key value template,
    ''' returning one expanded string per applicable package.
    ''' <br /><br />
    ''' 
    ''' Rules:
    ''' <list type="bullet">
    ''' 
    ''' <item>
    ''' <c> %PackageN% </c> (numbered) — expands exactly once using the Nth package;
    ''' any other packages are ignored
    ''' </item>
    ''' 
    ''' <item>
    ''' <c> %Package% </c> (unnumbered) — expands once per package, producing one
    ''' output string per package in order
    ''' </item>
    ''' 
    ''' <item>
    ''' No package variable — returned verbatim as a single-element list
    ''' </item>
    ''' 
    ''' </list>
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
    Private Function expandPackageKey(template As String,
                                      packages As List(Of String)) As List(Of String)

        Dim result As New List(Of String)

        ' Check for numbered package references (%Package1%, %Package2%, …)
        ' These appear only in AppInfo FileKey values as package selectors — scaffold templates
        ' always use unnumbered %Package%. A numbered reference selects exactly one package by position.
        For i = 1 To packages.Count

            Dim numberedVar = $"%Package{i}%"
            If Not template.Contains(numberedVar) Then Continue For

            result.Add(template.Replace(numberedVar, $"%LocalAppData%\Packages\{packages(i - 1)}"))
            Return result

        Next

        ' Unnumbered %Package% — expand for every package
        If template.Contains("%Package%") Then

            For Each pkg In packages
                result.Add(template.Replace("%Package%", $"%LocalAppData%\Packages\{pkg}"))
            Next

            Return result

        End If

        ' No package variable — pass through verbatim
        result.Add(template)
        Return result

    End Function

End Module
