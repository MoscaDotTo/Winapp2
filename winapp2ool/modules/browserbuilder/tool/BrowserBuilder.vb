'    Copyright (C) 2018-2025 Hazel Ward
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
''' This module expects to interact with at least one of two specially formatted ini files. 
''' These files, chromium.ini and gecko.ini, provide the rulesets needed to generate entries for
''' web browsers running with these respective engines. Refer to these files for more specific 
''' documentation on how they are used by the user. BrowserBuilder is primarily intended as a 
''' small devops tool but end users might find it useful as it enables them to generate entries 
''' for non standard installation paths or portable applications while keeping winapp2.ini up
''' to date separately <br /><br />
''' 
''' Additionally, and because it is intended to simplify the devops, Browser Builder expects to 
''' "Flavorize" the browsers.ini that it produces. This is done so as to resolve over and under 
''' coverages by generated entries as well as resolve incompatibilities. Accordingly, it expects
''' but does not require a set of Flavor files (up to 1 "Add" file, up to 3 "Remove" files, and
''' up to 2 "Replace" files). Consult the Transmute documentation for more information about Flavors
''' <br/><br/> Unflavored browser.ini files may be broken or incomplete depending on use case
''' </summary>
''' 
''' Docs last updated: 2025-08-01 | Code last updated: 2025-08-01
Public Module BrowserBuilder

    ''' <summary>
    ''' Stores the information required by EntryScaffold sections to 
    ''' generate individual browser entries
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-08-01 | Code last updated: 2025-08-01
    Private Structure BrowserInfo

        ''' <summary>
        ''' The name of the browser, will be preprended to the name of each entry scaffold
        ''' </summary>
        ''' 
        ''' Docs last updated: 2025-08-01 | Code last updated: 2025-08-01
        Public Name As String

        ''' <summary>
        ''' The set of provided user data (chromium) or profiles (gecko) paths
        ''' </summary>
        ''' 
        ''' Docs last updated: 2025-08-01 | Code last updated: 2025-08-01
        Public UserDataPaths As List(Of String)

        ''' <summary>
        ''' The set of parent paths to the <c> UserDataPaths </c>
        ''' </summary>
        ''' 
        ''' Docs last updated: 2025-08-01 | Code last updated: 2025-08-01
        Public UserDataParentPaths As List(Of String)

        ''' <summary>
        ''' The name of the CCleaner "Section" that all entries for this browser will be grouped into
        ''' </summary>
        ''' 
        ''' Docs last updated: 2025-08-01 | Code last updated: 2025-08-01
        Public SectionName As String

        ''' <summary>
        ''' Indicates that the User Data path should be truncated off for the DetectFile <br />
        ''' Useful for easily supporting multiple versions of a single browser
        ''' </summary>
        ''' 
        ''' Docs last updated: 2025-08-01 | Code last updated: 2025-08-01
        Public TruncateDetect As Boolean

        ''' <summary>
        ''' The set of parent paths in the registry for the browser, used to generate RegKeys
        ''' </summary>
        ''' 
        ''' Docs last updated: 2025-08-01 | Code last updated: 2025-08-01
        Public RegistryRoots As List(Of String)

        ''' <summary>
        ''' Indicates whether or not the current browser should be omitted from the generation
        ''' process <br /><br />
        ''' Allows the easy enabling and disabling of browser support over time without requiring 
        ''' any information to be truly lost 
        ''' </summary>
        ''' 
        ''' Docs last updated: 2025-08-01 | Code last updated: 2025-08-01
        Public ShouldSkip As Boolean

        ''' <summary>
        ''' Creates a new BrowserInfo object for a particular browser
        ''' </summary>
        ''' 
        ''' <param name="name">
        ''' The name of the web browser as it appears in the BrowserInfo section name 
        ''' </param>
        ''' 
        ''' Docs last updated: 2025-08-01 | Code last updated: 2025-08-01
        Public Sub New(name As String)

            Me.Name = name
            UserDataPaths = New List(Of String)
            UserDataParentPaths = New List(Of String)
            SectionName = ""
            TruncateDetect = False
            RegistryRoots = New List(Of String)
            ShouldSkip = False

        End Sub

    End Structure

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <returns></returns>
    Private Property totalChromiumCount As Integer = 0

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <returns></returns>
    Private Property totalGeckoCount As Integer = 0

    ''' <summary>
    ''' Handles the commandline arguments for <c> BrowserBuilder </c>
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-08-23 | Code last updated: 2025-08-23
    Public Sub handleCmdLine()

        getFileAndDirParams({BuilderFile1, BuilderFile2, BuilderFile3,
                    BuilderFile4, BuilderFile5, BuilderFile6,
                    BuilderFile7, BuilderFile8, BuilderFile9})

        initBrowserBuilder()

    End Sub

    ''' <summary>
    ''' Initializes the browser builder process
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-07-02 | Code last updated: 2025-08-27
    Public Sub initBrowserBuilder()

        clrConsole()
        BuilderFile3.Sections.Clear()

        BuilderFile1.init()
        BuilderFile2.init()

        Dim noRules = BuilderFile1.Sections.Count + BuilderFile2.Sections.Count = 0

        setHeaderText("No valid generative rulesets found", True, noRules)
        If noRules Then Return

        Dim browserBuilderStartPhrase = "Building browser configuration entries"
        Dim output As New MenuSection
        output.AddBoxWithText(browserBuilderStartPhrase)
        gLog(browserBuilderStartPhrase, buffr:=True, ascend:=True)

        processBrowserBuilder(output)

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
    ''' Docs last updated: 2025-07-02 | Code last updated: 2025-07-02
    Private Sub processBrowserBuilder(ByRef output As MenuSection)

        gLog("Processing browser builder files", ascend:=True, buffr:=True)

        ' Create the output ini file
        Dim outputFile As New iniFile With {
            .Dir = BuilderFile3.Dir,
            .Name = BuilderFile3.Name
        }

        ' Process each scaffold entry for Chromium
        buildScaffolds(BuilderFile1, False, outputFile, output)
        totalChromiumCount = outputFile.Sections.Count

        ' Process each scaffold entry for Gecko 
        buildScaffolds(BuilderFile2, True, outputFile, output)
        totalGeckoCount = outputFile.Sections.Count - totalChromiumCount

        ' Apply corrections to our generated output and save
        Flavorize(outputFile, outputFile, output, BuilderFile4, BuilderFile5, BuilderFile6, BuilderFile9, BuilderFile8, BuilderFile7)

        outputFile.Sections = remotedebug(outputFile, True).Sections

        Dim leadingComments As New List(Of String) From {
            $"; Version {DateTime.Now.ToString("yyMMdd")}",
            $"; # of entries: {outputFile.Sections.Count:#,###}",
            $"; {outputFile.Name} is generated by the Winapp2ool Browser Builder",
            "; Entries in this file may be incomplete and are not intended to be used directly with any cleaning software",
            "; They are utilized by winapp2ool to create the final winapp2.ini file for distribution",
            "; If you are not maintaining winapp2.ini for distribution, you probably don't need this file!",
            "; Refer to the Winapp2ool documentation for more information: " & readMeUrl,
            "; You can find the complete winapp2.ini file here: " & winapp2link()
        }

        Dim sb As New StringBuilder()

        For Each comment In leadingComments

            sb.AppendLine(comment)

        Next

        sb.AppendLine()
        sb.Append(outputFile.toString)

        outputFile.overwriteToFile(sb.ToString())

        gLog("Browser builder files processed successfully", descend:=True)

    End Sub

    ''' <summary>
    ''' Builds each <c> EntryScaffold </c> and then generates an appropriate entry for each 
    ''' Browser provided in the <c> <paramref name="rulesetFile"/> </c>
    ''' </summary>
    ''' 
    ''' <param name="rulesetFile">
    ''' The <c> iniFile </c> containing the set of generative rules for a particular group
    ''' of web browsers
    ''' </param>
    ''' 
    ''' <param name="isGecko">
    ''' Indicates whether or not the current group of web browsers being operated on is gecko-based
    ''' </param>
    ''' 
    ''' <param name="outputFile">
    ''' The location to which the generative output of Browser Builder will be be stored in 
    ''' memory and also saved to disk 
    ''' </param>
    ''' 
    ''' Docs last updated: 2025-08-01 | Code last updated: 2025-08-01
    Private Sub buildScaffolds(rulesetFile As iniFile,
                               isGecko As Boolean,
                         ByRef outputFile As iniFile,
                         ByRef menuOutput As MenuSection)

        Dim browsers As New List(Of BrowserInfo)
        Dim scaffoldSections As New List(Of iniSection)

        For Each section In rulesetFile.Sections.Values

            Select Case True

                Case section.Name.StartsWith("BrowserInfo:", StringComparison.InvariantCulture)

                    Dim browserInfo As BrowserInfo = parseBrowserInfo(section, menuOutput)
                    If browserInfo.ShouldSkip Then Continue For
                    browsers.Add(browserInfo)

                Case section.Name.StartsWith("EntryScaffold:", StringComparison.InvariantCulture)

                    scaffoldSections.Add(section)

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
    ''' The <c> iniSection </c> containing the BrowserInfo data
    ''' </param>
    ''' 
    ''' <returns>
    ''' A <c> BrowserInfo </c> structure containing all parsed browser parameters
    ''' </returns>
    ''' 
    ''' Docs last updated: 2025-08-01 | Code last updated: 2025-08-01
    Private Function parseBrowserInfo(browserSection As iniSection,
                                      ByRef menuOutput As MenuSection) As BrowserInfo

        Dim browserName As String = browserSection.Name.Substring("BrowserInfo: ".Length)
        Dim browserInfo As New BrowserInfo(browserName)

        For Each key In browserSection.Keys.Keys

            Select Case key.KeyType.ToUpperInvariant()

                Case "USERDATAPATH"

                    browserInfo.UserDataPaths.Add(key.Value)
                    browserInfo.UserDataParentPaths.Add(key.Value.Substring(0, key.Value.LastIndexOf("\"c)))

                Case "SECTION"

                    browserInfo.SectionName = key.Value

                Case "TRUNCATEDETECT"

                    browserInfo.TruncateDetect = True

                Case "SKIP"

                    browserInfo.ShouldSkip = True

                Case "REGISTRYROOT"

                    browserInfo.RegistryRoots.Add(key.Value)

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
    ''' A particular <c> EntryScaffold </c> section to be generated for each browser
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
    ''' The location to which the generative output of Browser Builder will be be stored in 
    ''' memory and also saved to disk 
    ''' </param>
    ''' 
    ''' Docs last updated: 2025-08-01 | Code last updated: 2025-08-01
    Private Sub processEntryScaffold(scaffoldSection As iniSection,
                                     browsers As List(Of BrowserInfo),
                                     isGecko As Boolean,
                               ByRef outputFile As iniFile,
                               ByRef menuOutput As MenuSection)

        Dim scaffoldName As String = scaffoldSection.Name.Substring("EntryScaffold: ".Length)

        Dim processingMsg = $"Processing EntryScaffold: {scaffoldName}"
        menuOutput.AddColoredLine(processingMsg, ConsoleColor.Magenta)
        gLog(processingMsg, ascend:=True)

        Dim fileKeyBases As New List(Of String)
        Dim regKeyBases As New List(Of String)

        For Each key In scaffoldSection.Keys.Keys

            Select Case key.KeyType.ToUpperInvariant()

                Case "FILEKEYBASE"

                    fileKeyBases.Add(key.Value)

                Case "REGKEYBASE"

                    regKeyBases.Add(key.Value)

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

    ''' <summary>
    ''' Produces entries for the Opera GX browser by transforming the EntryScaffold template
    ''' paths in accordance with the non-standard way Opera GX stores itself on the disk by default
    ''' </summary>
    ''' 
    ''' <param name="browserName"> The name of the Browser being Processed <br /> 
    ''' In this particular instance, this will always be Opera GX
    ''' </param>
    ''' 
    ''' <param name="fileKeyBase">
    ''' A FileKey template to be transformed such that it is appropriate for Opera GX
    ''' </param>
    ''' 
    ''' <param name="fileKeyNum">
    ''' The number of the current FileKey being generated
    ''' </param>
    ''' 
    ''' <param name="newSection">
    ''' The <c> iniSection </c> currently being generated by Browser Builder 
    ''' </param>
    ''' 
    ''' <param name="userDataPath">
    ''' The set of values for %UserDataPath%, provided by the BrowserInfo section for the 
    ''' current browser
    ''' </param>
    ''' 
    ''' Docs last updated: 2025-07-31 | Code last updated: 2025-07-31
    Private Function processOperaGX(browserName As String,
                                    fileKeyBase As String,
                              ByRef fileKeyNum As Integer,
                              ByRef newSection As iniSection,
                                    userDataPath As String) As Boolean

        If Not browserName.Equals("Opera GX") Then Return False

        Dim baseKeyValue As String
        Dim sideProfileKeyValue As String

        ' Opera GX stores files both in the main user data directory (without a folder such as Default)
        ' and any profiles created in _side_profiles\* subdirectories instead of the usual Profile1, etc.
        ' Accordingly, we transform our scaffold templates to accommodate this 
        ' %UserDataPath%\*|<params> -> %UserDataPath%|<params> + %UserDataPath%\_side_profiles\*|<params>
        ' %UserDataPath%\*\<directory> -> %UserDataPath%\<directory> + %UserDataPath%\_side_profiles\*\<directory>
        ' %UserDataPath%|<params> -> %UserDataPath%|<params> + %UserDataPath%\_side_profiles\*|<params>
        ' %UserDataPath%\<directory> -> %UserDataPath%\<directory> + %UserDataPath%\_side_profiles\*\<directory>

        Select Case True

            '%UserDataPath%\*|<params> -> remove \* + create side profiles variant
            Case fileKeyBase.Contains("%UserDataPath%\*|")

                Dim afterWildcardPipe As String = fileKeyBase.Substring(fileKeyBase.IndexOf("%UserDataPath%\*|") + "%UserDataPath%\*|".Length)
                Dim basePattern As String = "%UserDataPath%|" & afterWildcardPipe
                baseKeyValue = basePattern.Replace("%UserDataPath%", userDataPath)
                Dim sideProfilesPattern As String = "%UserDataPath%\_side_profiles\*|" & afterWildcardPipe
                sideProfileKeyValue = sideProfilesPattern.Replace("%UserDataPath%", userDataPath)

            ' %UserDataPath%\*\<directory> -> unnest one level + create side profiles variant
            Case fileKeyBase.Contains("%UserDataPath%\*\")

                Dim afterWildcard As String = fileKeyBase.Substring(fileKeyBase.IndexOf("%UserDataPath%\*\") + "%UserDataPath%\*\".Length)
                Dim unnestedBase As String = "%UserDataPath%\" & afterWildcard
                baseKeyValue = unnestedBase.Replace("%UserDataPath%", userDataPath)
                Dim sideProfilesBase As String = "%UserDataPath%\_side_profiles\*\" & afterWildcard
                sideProfileKeyValue = sideProfilesBase.Replace("%UserDataPath%", userDataPath)

            ' %UserDataPath%|<params> -> keep + create side profiles variant 
            Case fileKeyBase.StartsWith("%UserDataPath%|")

                Dim afterPipe As String = fileKeyBase.Substring("%UserDataPath%|".Length)
                baseKeyValue = fileKeyBase.Replace("%UserDataPath%", userDataPath)
                Dim sideProfilesBase As String = "%UserDataPath%\_side_profiles\*|" & afterPipe
                sideProfileKeyValue = sideProfilesBase.Replace("%UserDataPath%", userDataPath)

            ' %UserDataPath%\<directory> -> (non-wildcard subdirectory - RETAINS wildcards in directory names)
            Case fileKeyBase.StartsWith("%UserDataPath%\")

                Dim afterBackslash As String = fileKeyBase.Substring("%UserDataPath%\".Length)
                baseKeyValue = fileKeyBase.Replace("%UserDataPath%", userDataPath)
                Dim sideProfilesBase As String = "%UserDataPath%\_side_profiles\*\" & afterBackslash
                sideProfileKeyValue = sideProfilesBase.Replace("%UserDataPath%", userDataPath)

            Case Else

                Dim errMsg = $"Unsupported OperaGX path provided in {fileKeyBase}"
                gLog(errMsg)

        End Select

        CreateAndAddKey(newSection, "FileKey", fileKeyNum, baseKeyValue)
        CreateAndAddKey(newSection, "FileKey", fileKeyNum, sideProfileKeyValue)

        Return True

    End Function

    ''' <summary>
    ''' Generates a browser entry using pre-computed browser information
    ''' </summary>
    ''' 
    ''' <param name="browserInfo"> 
    ''' Pre-computed <c> BrowserInfo </c> structure containing all browser parameters
    ''' </param>
    ''' 
    ''' <param name="scaffoldName"> 
    ''' The name of the <c> EntryScaffold </c> used as a template for the current entry being 
    ''' generated 
    ''' </param>
    ''' 
    ''' <param name="fileKeyBases">
    ''' A set of FileKey templates that will be personalized for the current browser
    ''' </param>
    ''' 
    ''' <param name="regKeyBases">
    ''' A set of RegKey templates that will be personalized for the current browser
    ''' </param>
    ''' 
    ''' <param name="outputFile">
    ''' The <c> iniFile </c> containing all the entries being generated by Browser Builder
    ''' and to which the entry built by this function will be added 
    ''' </param>
    ''' 
    ''' <param name="isGecko">
    ''' Indicates whether or not the Browser being processed is a Gecko-based browser
    ''' </param>
    ''' 
    ''' Docs last updated: 2025-08-01 | Code last updated: 2025-08-01
    Private Sub generateBrowserEntry(browserInfo As BrowserInfo,
                                     scaffoldName As String,
                                     fileKeyBases As List(Of String),
                                     regKeyBases As List(Of String),
                               ByRef outputFile As iniFile,
                                     isGecko As Boolean,
                               ByRef menuOutput As MenuSection)

        Dim generatingMsg = $"Generating entry for browser: {browserInfo.Name}"
        gLog(generatingMsg, indent:=True)
        menuOutput.AddColoredLine(generatingMsg, ConsoleColor.Magenta)

        ' Create the new entry section
        Dim entryName As String = $"{browserInfo.Name} {scaffoldName} *"
        Dim newSection As New iniSection With {.Name = entryName}
        newSection.Keys.add(New iniKey($"Section={browserInfo.SectionName}"))

        processUserDataPaths(newSection, browserInfo.UserDataPaths, browserInfo.UserDataParentPaths, browserInfo.SectionName, browserInfo.TruncateDetect, fileKeyBases, browserInfo.Name, isGecko)

        ' Process registry keys
        processRegKeyBases(regKeyBases, newSection, browserInfo.RegistryRoots)

        ' Add the section to the output file
        outputFile.Sections.Add(entryName, newSection)

        Dim generatedText = $"Generated entry: {entryName}"
        menuOutput.AddColoredLine(generatedText, ConsoleColor.Yellow)
        gLog(generatedText, indent:=True, indAmt:=4)

    End Sub

    ''' <summary>
    ''' Generates any necessary RegKeys entries based on the provided RegKeyBases 
    ''' and adds them into <c> <paramref name="newSection"/> </c>
    ''' </summary>
    ''' 
    ''' <param name="regKeyBases">
    ''' The set of templates for RegKeys that will be personalized for the current browser
    ''' </param>
    ''' 
    ''' <param name="newSection">
    ''' The <c> iniSection </c> to which the generated RegKeys will be added
    ''' </param>
    ''' 
    ''' <param name="RegistryRoots">
    ''' The set of root paths in the registry for the current browser's RegKeys
    ''' </param>
    ''' 
    ''' Docs last updated: 2025-08-01 | Code last updated: 2025-08-01
    Private Sub processRegKeyBases(regKeyBases As List(Of String),
                             ByRef newSection As iniSection,
                                   RegistryRoots As List(Of String))

        Dim regKeyNum As Integer = 1

        For Each regKeyBase In regKeyBases

            For Each root In RegistryRoots

                CreateAndAddKey(newSection, "RegKey", regKeyNum, regKeyBase.Replace("%RegistryRoot%", root))

            Next

        Next

    End Sub

    ''' <summary>
    ''' Creates a numbered <c> iniKey </c> with the provided <c> <paramref name="keyName"/> </c>,
    ''' <c> <paramref name="keyNum"/> </c>, and <paramref name="keyValue"/>, then adds it to the
    ''' provided <c> <paramref name="section"/> </c>, Incrementing the <c> <paramref name="keyNum"/> </c>
    ''' </summary>
    ''' 
    ''' <param name="section">
    ''' An <c> iniSection </c> to which a key will be added
    ''' </param>
    ''' 
    ''' <param name="keyName">
    ''' The name of the <c> iniKey </c>, does not include numbers
    ''' </param>
    ''' 
    ''' <param name="keyNum">
    ''' The number after the <c> <paramref name="keyName"/> </c>
    ''' </param>
    ''' 
    ''' <param name="keyValue">
    ''' The value of the <c> iniKey </c>
    ''' </param>
    ''' 
    ''' Docs last updated: 2025-08-01 | Code last updated: 2025-08-01
    Private Sub CreateAndAddKey(ByRef section As iniSection,
                                      keyName As String,
                                ByRef keyNum As Integer,
                                      keyValue As String)

        Dim key As New iniKey($"{keyName}{keyNum}={keyValue}")
        section.Keys.add(key)
        keyNum += 1

    End Sub

    ''' <summary>
    ''' Iterates through the user data paths and generates DetectFile and FileKey entries
    ''' for each path, adding them to the provided <c> newSection </c>
    ''' </summary>
    ''' 
    ''' <param name="newSection">
    ''' The <c> inSection </c> currently being built by Browser Builder and to which keys will be added
    ''' </param>
    ''' 
    ''' <param name="userDataPaths">
    ''' The set of user data paths provided from the ruleset file for the current browser
    ''' </param>
    ''' 
    ''' <param name="userDataParentPaths">
    ''' The set of parent paths for the user data paths, used for truncation
    ''' </param>
    ''' 
    ''' <param name="sectionName">
    ''' The name of the section currently being built
    ''' </param>
    ''' 
    ''' <param name="truncate">
    ''' Indicates that the parent path should be used instead of the user data path for the purpose 
    ''' of creating DetectFile keys. This is useful for browsers that may have multiple versions 
    ''' in the same parent folder 
    ''' </param>
    ''' 
    ''' <param name="fileKeyBases">
    ''' The set of FileKey templates provided by the current <c> EntryScaffold </c>
    ''' </param>
    ''' 
    ''' <param name="BrowserName">
    ''' The name of the Browser for which the current entry is being built
    ''' </param>
    ''' 
    ''' <param name="isGecko">
    ''' Indicates whether or not the current browser is gecko-based 
    ''' </param>
    ''' 
    ''' Docs last updated: 2025-07-31 | Code last updated: 2025-07-31
    Private Sub processUserDataPaths(ByRef newSection As iniSection,
                                     ByRef userDataPaths As List(Of String),
                                     ByRef userDataParentPaths As List(Of String),
                                           sectionName As String,
                                           truncate As Boolean,
                                     ByRef fileKeyBases As List(Of String),
                                           BrowserName As String,
                                           isGecko As Boolean)

        Dim fileKeyNum As Integer = 1

        For i = 0 To userDataPaths.Count - 1

            Dim DetectFileNum = If(userDataPaths.Count > 1, (i + 1).ToString, "")
            newSection.Keys.add(New iniKey($"DetectFile{DetectFileNum}={If(truncate, userDataParentPaths(i), userDataPaths(i))}"))

            For Each fileKeyBase In fileKeyBases

                Dim fileKeyValue As String

                Select Case True

                    Case fileKeyBase.Contains("%UserDataPath%")

                        ' OperaGX is a nightmare!! We need to process it separately 
                        If processOperaGX(BrowserName, fileKeyBase, fileKeyNum, newSection, userDataPaths(i)) Then Continue For

                        ' Replace %UserDataPath% with the actual user data path
                        fileKeyValue = fileKeyBase.Replace("%UserDataPath%", userDataPaths(i))


                    Case fileKeyBase.Contains("%BrowserPath%")

                        ' For both types of browsers we want this to capture the parent of the user data path
                        fileKeyValue = fileKeyBase.Replace("%BrowserPath%", userDataParentPaths(i))

                        If isGecko Then Exit Select

                        ' For Chromium browsers, this is typically being used to point to the place where 
                        ' the application files can be found. We need to also support %ProgramFiles%
                        ' for this because many distributions of Chromium install to this location 
                        ' on Windows when installed for All Users
                        Dim fileKeyValue2 = fileKeyValue.Replace("%LocalAppData%", "%ProgramFiles%")

                        ' A small numer of browsers live in %AppData% instead of %LocalAppData% by default 
                        fileKeyValue2 = fileKeyValue2.Replace("%AppData%", "%ProgramFiles%")
                        CreateAndAddKey(newSection, "FileKey", fileKeyNum, fileKeyValue2)

                    ' We only support %LocalDataPath% on Gecko 
                    Case fileKeyBase.Contains("%LocalDataPath%") AndAlso isGecko

                        Dim localdatapath = userDataPaths(i).Replace("%AppData%", "%LocalAppData%")
                        fileKeyValue = fileKeyBase.Replace("%LocalDataPath%", localdatapath)

                End Select

                CreateAndAddKey(newSection, "FileKey", fileKeyNum, fileKeyValue)

            Next

        Next

    End Sub

End Module
