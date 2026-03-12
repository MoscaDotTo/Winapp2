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
''' BrowserBuilder is a winapp2ool module which handles the generation of bespoke winapp2.ini
''' entries for a very large number of web browsers. It provides a small scripting interface
''' which allows for the dynamic creation of entire groups of winapp2.ini entries using a limited
''' set of easily accessed information. <br /><br />
'''
''' All input files are read from a single configurable source directory. At minimum, that
''' directory must contain at least one of <c>chromium.ini</c> or <c>gecko.ini</c> - the
''' specially formatted ruleset files that drive entry generation. BrowserBuilder is primarily
''' intended as a small devops tool but end users might find it useful as it enables them to
''' generate entries for non-standard installation paths or portable applications while keeping
''' winapp2.ini up to date separately <br /><br />
'''
''' Additionally, and because it is intended to simplify the devops, Browser Builder expects to
''' "Flavorize" the browsers.ini that it produces. This is done so as to resolve over and under
''' coverages by generated entries as well as resolve incompatibilities. Accordingly, it expects
''' but does not require a set of Flavor correction files in the source directory (up to 1 "Add"
''' file, up to 3 "Remove" files, and up to 2 "Replace" files) - all resolved by canonical name.
''' Consult the Transmute documentation for more information about Flavors <br /><br />
''' Unflavored browser.ini files may be broken or incomplete depending on use case
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
    ''' Handles the commandline arguments for <c>BrowserBuilder</c>
    ''' </summary>
    '''
    ''' <remarks>
    ''' <b> File arguments: </b> <br />
    ''' <c> -1d </c> <br /> Source directory containing <c> chromium.ini </c>,
    ''' <c> gecko.ini </c>, and flavor correction files 
    ''' <br /> <br /> 
    ''' <c>-2d</c> / <c>-2f</c> <br /> Output file path and name 
    ''' (default: <c> browsers.ini </c> in the current directory)
    ''' </remarks>
    Public Sub handleCmdLine()

        getFileAndDirParams({BuilderFile1, BuilderFile2})

        initBrowserBuilder()

    End Sub

    ''' <summary>
    ''' Initializes the browser builder process
    ''' </summary>
    Public Sub initBrowserBuilder()

        clrConsole()

        Dim sourceDir = BuilderFile1.Dir

        Dim chromiumIni = iniFile2.FromFile($"{sourceDir}\chromium.ini")
        Dim geckoIni = iniFile2.FromFile($"{sourceDir}\gecko.ini")
        Dim noRules = chromiumIni.Count + geckoIni.Count = 0

        If noRules Then

            setNextMenuHeaderText("No valid generative rulesets found", printColor:=ConsoleColor.Red)
            Return

        End If

        Dim browserBuilderStartPhrase = "Building browser configuration entries"
        Dim output As New MenuSection
        output.AddBoxWithText(browserBuilderStartPhrase)
        gLog(browserBuilderStartPhrase, buffr:=True, ascend:=True)

        processBrowserBuilder(chromiumIni, geckoIni, output)

        Dim browserBuilderEndPhrase = "Browser configuration entries built successfully"
        output.AddBoxWithText(browserBuilderEndPhrase)
        gLog(browserBuilderEndPhrase, descend:=True, buffr:=True)

        output.AddAnyKeyPrompt()

        If Not SuppressOutput Then output.Print()
        crk()

    End Sub

    ''' <summary>
    ''' Processes the browser builder files and generates the output
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
    ''' The <c>MenuSection</c> accumulating user-visible output for this run
    ''' </param>
    Private Sub processBrowserBuilder(chromiumIni As iniFile2,
                                      geckoIni As iniFile2,
                                      ByRef output As MenuSection)

        gLog("Processing browser builder files", ascend:=True, buffr:=True)

        Dim outputFile = iniFile2.Empty(BuilderFile2.Dir, BuilderFile2.Name)

        buildScaffolds(chromiumIni, False, outputFile, output)
        totalChromiumCount = outputFile.Count

        buildScaffolds(geckoIni, True, outputFile, output)
        totalGeckoCount = outputFile.Count - totalChromiumCount

        ' Convert to legacy iniFile at the Flavorize/remotedebug boundary
        ' We'll remove this once Flavorize and WinappDebug operate on iniFile2
        Dim legacyOutput = IniFileBridge.ToIniFile(outputFile)

        Dim sourceDir = BuilderFile1.Dir
        Dim flavorAdd As New iniFile(sourceDir, "browser_additions.ini")
        Dim flavorSecRem As New iniFile(sourceDir, "browser_section_removals.ini")
        Dim flavorNameRem As New iniFile(sourceDir, "browser_name_removals.ini")
        Dim flavorValRem As New iniFile(sourceDir, "browser_value_removals.ini")
        Dim flavorSecRep As New iniFile(sourceDir, "browser_section_replacements.ini")
        Dim flavorKeyRep As New iniFile(sourceDir, "browser_key_replacements.ini")

        flavorAdd.init()
        flavorSecRem.init()
        flavorNameRem.init()
        flavorValRem.init()
        flavorSecRep.init()
        flavorKeyRep.init()

        Flavorize(legacyOutput, legacyOutput, output, flavorAdd, flavorSecRem, flavorNameRem, flavorValRem, flavorSecRep, flavorKeyRep)

        legacyOutput.Sections = remotedebug(legacyOutput, True).Sections

        Dim leadingComments As New List(Of String) From {
            $"; Version {DateTime.Now.ToString("yyMMdd")}",
            $"; # of entries: {legacyOutput.Sections.Count:#,###}",
            $"; {legacyOutput.Name} is generated by the Winapp2ool Browser Builder",
            "; Entries in this file may be incomplete and are not intended to be used directly with any cleaning software",
            "; They are utilized by winapp2ool to create the final winapp2.ini file for distribution",
            "; If you are not maintaining winapp2.ini for distribution, you probably don't need this file!",
            "; Refer to the Winapp2ool documentation for more information: " & readMeUrl,
            "; You can find the complete winapp2.ini file here: " & baseFlavorLink
        }

        Dim sb As New StringBuilder()

        For Each comment In leadingComments : sb.AppendLine(comment) : Next

        sb.AppendLine()
        sb.Append(legacyOutput.toString)

        legacyOutput.overwriteToFile(sb.ToString())

        gLog("Browser builder files processed successfully", descend:=True)

    End Sub

    ''' <summary>
    ''' Builds each <c>EntryScaffold</c> and then generates an appropriate entry for each
    ''' browser provided in <paramref name="rulesetFile"/>
    ''' </summary>
    '''
    ''' <param name="rulesetFile">
    ''' The <c>iniFile2</c> containing the set of generative rules for a particular group
    ''' of web browsers
    ''' </param>
    '''
    ''' <param name="isGecko">
    ''' Indicates whether or not the current group of web browsers being operated on is gecko-based
    ''' </param>
    '''
    ''' <param name="outputFile">
    ''' The location to which the generative output of Browser Builder will be stored in memory
    ''' </param>
    '''
    ''' <param name="menuOutput">
    ''' The <c>MenuSection</c> accumulating user-visible output for this run
    ''' </param>
    Private Sub buildScaffolds(rulesetFile As iniFile2,
                               isGecko As Boolean,
                         ByRef outputFile As iniFile2,
                         ByRef menuOutput As MenuSection)

        Dim browsers As New List(Of BrowserInfo)
        Dim scaffoldSections As New List(Of iniSection2)

        For Each section In rulesetFile

            Select Case True

                Case section.Name.StartsWith("BrowserInfo:", StringComparison.InvariantCulture)

                    Dim browserInfo As BrowserInfo = parseBrowserInfo(section, menuOutput)
                    If browserInfo.ShouldSkip Then Continue For
                    browsers.Add(browserInfo)

                Case section.Name.StartsWith("EntryScaffold:", StringComparison.InvariantCulture) : scaffoldSections.Add(section)

                Case Else

                    Dim logMsg = $"Invalid section found and ignored: [{section.Name}]"
                    gLog(logMsg)
                    menuOutput.AddWarning(logMsg)

            End Select

        Next

        Dim configCount = $"Found {browsers.Count} browser configurations"
        menuOutput.AddColoredLine(configCount, ConsoleColor.Yellow)
        gLog(configCount, indent:=True)

        For Each scaffoldSection In scaffoldSections

            processEntryScaffold(scaffoldSection, browsers, isGecko, outputFile, menuOutput)

        Next

    End Sub

    ''' <summary>
    ''' Parses a BrowserInfo section and returns a pre-computed BrowserInfo structure
    ''' </summary>
    '''
    ''' <param name="browserSection">
    ''' The <c>iniSection2</c> containing the BrowserInfo data
    ''' </param>
    '''
    ''' <param name="menuOutput">
    ''' The <c>MenuSection</c> accumulating user-visible output for this run
    ''' </param>
    '''
    ''' <returns>
    ''' A <c>BrowserInfo</c> structure containing all parsed browser parameters
    ''' </returns>
    Private Function parseBrowserInfo(browserSection As iniSection2,
                                ByRef menuOutput As MenuSection) As BrowserInfo

        Dim browserName As String = browserSection.Name.Substring("BrowserInfo: ".Length)
        Dim browserInfo As New BrowserInfo(browserName)

        For Each key In browserSection.Keys

            Select Case key.KeyType.ToUpperInvariant()

                Case "USERDATAPATH"

                    browserInfo.UserDataPaths.Add(key.Value)
                    browserInfo.UserDataParentPaths.Add(key.Value.Substring(0, key.Value.LastIndexOf("\"c)))

                Case "SECTION" : browserInfo.SectionName = key.Value

                Case "TRUNCATEDETECT" : browserInfo.TruncateDetect = True

                Case "SKIP" : browserInfo.ShouldSkip = True

                Case "REGISTRYROOT" : browserInfo.RegistryRoots.Add(key.Value)

                Case Else

                    Dim logMsg = $"Unexpected KeyType in {browserSection.Name}: {key.KeyType}"
                    menuOutput.AddWarning(logMsg)
                    gLog(logMsg)

            End Select

        Next

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
    ''' Processes an EntryScaffold section and generates entries for each browser
    ''' </summary>
    '''
    ''' <param name="scaffoldSection">
    ''' A particular <c>EntryScaffold</c> section to be generated for each browser
    ''' </param>
    '''
    ''' <param name="browsers">
    ''' The set of all BrowserInfo sections from the ruleset file
    ''' </param>
    '''
    ''' <param name="isGecko">
    ''' Indicates whether or not the current browser is gecko-based
    ''' </param>
    '''
    ''' <param name="outputFile">
    ''' The location to which the generative output of Browser Builder will be stored in memory
    ''' </param>
    '''
    ''' <param name="menuOutput">
    ''' The <c>MenuSection</c> accumulating user-visible output for this run
    ''' </param>
    Private Sub processEntryScaffold(scaffoldSection As iniSection2,
                                     browsers As List(Of BrowserInfo),
                                     isGecko As Boolean,
                               ByRef outputFile As iniFile2,
                               ByRef menuOutput As MenuSection)

        Dim scaffoldName As String = scaffoldSection.Name.Substring("EntryScaffold: ".Length)

        Dim processingMsg = $"Processing EntryScaffold: {scaffoldName}"
        menuOutput.AddColoredLine(processingMsg, ConsoleColor.Magenta)
        gLog(processingMsg, ascend:=True)

        Dim fileKeyBases As New List(Of String)
        Dim regKeyBases As New List(Of String)

        For Each key In scaffoldSection.Keys

            Select Case key.KeyType.ToUpperInvariant()

                Case "FILEKEYBASE" : fileKeyBases.Add(key.Value)

                Case "REGKEYBASE" : regKeyBases.Add(key.Value)

                Case Else

                    Dim errMsg = $"Unexpected KeyType in {scaffoldSection.Name}: {key.KeyType}"
                    gLog(errMsg)
                    menuOutput.AddWarning(errMsg)

            End Select

        Next

        For Each browser In browsers

            generateBrowserEntry(browser, scaffoldName, fileKeyBases, regKeyBases, outputFile, isGecko, menuOutput)

        Next

        Dim finishedMsg = $"Finished processing EntryScaffold: {scaffoldName}"
        menuOutput.AddColoredLine(finishedMsg, ConsoleColor.Yellow)
        gLog(finishedMsg, descend:=True, buffr:=True)

    End Sub

End Module
