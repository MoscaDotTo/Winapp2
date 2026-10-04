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
Imports System.Text.RegularExpressions

''' <summary>
''' Observes, reports, and attempts to repair errors in winapp2.ini
''' </summary>
Public Module WinappDebug

    ''' <summary>
    ''' The number of errors found during the most recent <see cref="Debug"/> call. Optimization
    ''' merge reports aren't counted.
    ''' </summary>
    Public Property ErrorsFound As Integer = 0

    ''' <summary>
    ''' The error report from the most recent <see cref="Debug"/> call as plain text: each
    ''' error's header, message and detail lines, without the optimization merge blocks.
    ''' <see cref="EmitEntryResult"/> and <see cref="EmitEntryAlphabetizationErrors"/> both write here
    ''' from the orchestrating thread once the parallel pass is done, so this needs no
    ''' locking. Stays empty when the run found no errors, which is how the menu decides
    ''' whether to offer the Log Viewer option
    ''' </summary>
    Public Property MostRecentLintLog As New System.Text.StringBuilder

    ''' <summary>
    ''' Meant to hold the milliseconds spent in the parallel per-entry block of the most recent
    ''' <see cref="Debug"/> call. Nothing assigns it, so it stays <c> 0 </c>.
    ''' </summary>
    Public Property LastParallelElapsedMs As Long = 0

    ''' <summary>
    ''' Meant to hold the milliseconds spent alphabetizing entries in the most recent
    ''' <see cref="Debug"/> call. Nothing assigns it, so it stays <c> 0 </c>.
    ''' </summary>
    Public Property LastAlphabetizeElapsedMs As Long = 0

    ''' <summary>
    ''' Meant to hold the total milliseconds of the most recent <see cref="Debug"/> call.
    ''' Nothing assigns it, so it stays <c> 0 </c>.
    ''' </summary>
    Public Property LastLintElapsedMs As Long = 0

    ''' <summary>
    ''' The lint rules, in Scan Settings menu order. Each <c> lint* </c> property below picks
    ''' its rule from this list by index, so reordering the list reassigns them.
    ''' </summary>
    Public Property Rules As New List(Of lintRule) From {
        New lintRule(True, True, "Casing", "improper CamelCasing", "fixing improper CamelCasing"),
        New lintRule(True, True, "Alphabetization", "improper alphabetization", "fixing improper alphabetization"),
        New lintRule(True, True, "Improper Numbering", "improper key numbering", "fixing improper key numbering"),
        New lintRule(True, True, "Parameters", "improper parameterization on FileKeys", "fixing improper parameterization on FileKeys"),
        New lintRule(True, True, "Flags", "improper FileKey/ExcludeKey flag formatting", "fixing improper FileKey/ExcludeKey flag formatting"),
        New lintRule(True, True, "Slashes", "improper use of slashes (\)", "fixing improper use of slashes (\)"),
        New lintRule(True, True, "Defaults", "Default=True", "enforcing no default key"),
        New lintRule(True, True, "Duplicates", "duplicate key values", "removing keys with duplicated values"),
        New lintRule(True, True, "Unneeded Numbering", "use of numbers where there should not be", "removing numbers used where they shouldn't be"),
        New lintRule(True, True, "Multiples", "multiples of key types that should only occur once in an entry", "removing unneeded multiples of key types that should occur only once"),
        New lintRule(True, True, "Invalid Values", "invalid key values", "fixing certain types of invalid key values"),
        New lintRule(True, True, "Syntax Errors", "some entries whose configuration will not run in CCleaner", "attempting to fix certain types of syntax errors"),
        New lintRule(True, True, "Path Validity", "invalid filesystem or registry locations", "attempting to repair some basic invalid parameters in paths"),
        New lintRule(True, True, "Semicolons", "improper use of semicolons (;)", "fixing some improper uses of semicolons(;)"),
        New lintRule(False, False, "Optimizations", "situations where keys can be merged (experimental)", "automatic merging of keys (experimental)")
    }

    ''' <summary>
    ''' Controls scan/repairs for CamelCasing issues
    ''' <br /> Default: <c> True </c>
    ''' </summary>
    Private Property lintCasing As lintRule = Rules(0)

    ''' <summary>
    ''' Controls scan/repairs for alphabetization issues
    ''' <br /> Default: <c> True </c>
    ''' </summary>
    Private Property lintAlpha As lintRule = Rules(1)

    ''' <summary>
    ''' Controls scan/repairs for incorrectly numbered keys
    ''' <br /> Default: <c> True </c>
    ''' </summary>
    Private Property lintWrongNums As lintRule = Rules(2)


    ''' <summary>
    ''' Controls scan/repairs for parameters inside of FileKeys
    ''' <br /> Default: <c> True </c>
    ''' </summary>
    Private Property lintParams As lintRule = Rules(3)


    ''' <summary>
    ''' Controls scan/repairs for flags in ExcludeKeys and FileKeys
    ''' <br /> Default: <c> True </c>
    ''' </summary>
    Private Property lintFlags As lintRule = Rules(4)

    ''' <summary>
    ''' Controls scan/repairs for improper slash usage
    ''' <br /> Default: <c> True </c>
    ''' </summary>
    Private Property lintSlashes As lintRule = Rules(5)

    ''' <summary>
    ''' Controls scan/repairs for Default keys of any value, or for the value check when
    ''' <see cref="overrideDefaultVal"/> is on
    ''' <br /> Default: <c> True </c>
    ''' </summary>
    Private Property lintDefaults As lintRule = Rules(6)

    ''' <summary>
    ''' Controls scan/repairs for duplicate values
    ''' <br /> Default: <c> True </c>
    ''' </summary>
    Private Property lintDupes As lintRule = Rules(7)

    ''' <summary>
    ''' Controls scan/repairs for keys with numbers they shouldn't have
    ''' <br /> Default: <c> True </c>
    ''' </summary>
    Private Property lintExtraNums As lintRule = Rules(8)

    ''' <summary>
    ''' Controls scan/repairs for keys which should only occur once
    ''' <br /> Default: <c> True </c>
    ''' </summary>
    Private Property lintMulti As lintRule = Rules(9)

    ''' <summary>
    ''' Controls scan/repairs for keys with invalid values
    ''' <br /> Default: <c> True </c>
    ''' </summary>
    Private Property lintInvalid As lintRule = Rules(10)

    ''' <summary>
    ''' Controls scan/repairs for winapp2.ini syntax errors
    ''' <br /> Default: <c> True </c>
    ''' </summary>
    Private Property lintSyntax As lintRule = Rules(11)

    ''' <summary>
    ''' Controls scan/repairs for invalid file or registry paths
    ''' <br /> Default: <c> True </c>
    ''' </summary>
    Private Property lintPathValidity As lintRule = Rules(12)

    ''' <summary>
    ''' Controls scan/repairs for improper use of semicolons
    ''' <br /> Default: <c> True </c>
    ''' </summary>
    Private Property lintSemis As lintRule = Rules(13)

    ''' <summary>
    ''' Controls the FileKey merge in <see cref="cOptimization"/>. The <c> -opti </c> flag and
    ''' <see cref="remotedebug"/>'s <c> forceOpti </c> turn it on.
    ''' <br /> Default: <c> False </c>
    ''' </summary>
    Public Property lintOpti As lintRule = Rules(14)

    ''' <summary>
    ''' Regex to detect long form registry paths
    ''' </summary>
    Private ReadOnly longReg As New Regex("HKEY_(C(URRENT_(USER$|CONFIG$)|LASSES_ROOT$)|LOCAL_MACHINE$|USERS$)",
                                          RegexOptions.Compiled Or RegexOptions.CultureInvariant)

    ''' <summary>
    ''' Regex to detect short form registry paths
    ''' </summary>
    Private ReadOnly shortReg As New Regex("HK(C(C$|R$|U$)|LM$|U$)",
                                           RegexOptions.Compiled Or RegexOptions.CultureInvariant)

    ''' <summary>
    ''' Regex to detect valid LangSecRef numbers. It isn't anchored, so it matches a valid
    ''' number anywhere in the value.
    ''' </summary>
    Private ReadOnly secRefNums As New Regex("30(0([1-6])|2([1-9])|3([0-9])|4([0-4]))",
                                             RegexOptions.Compiled Or RegexOptions.CultureInvariant)

    ''' <summary>
    ''' Regex to detect valid drive letter parameters
    ''' </summary>
    Private ReadOnly driveLtrs As New Regex("[a-zA-Z]:",
                                            RegexOptions.Compiled Or RegexOptions.CultureInvariant)

    ''' <summary>
    ''' Regex to detect potential %EnvironmentVariables%
    ''' </summary>
    Private ReadOnly envVarRegex As New Regex("%[A-Za-z0-9]*%",
                                              RegexOptions.Compiled Or RegexOptions.CultureInvariant)

    ''' <summary>
    ''' Regex to detect ExcludeKey flags
    ''' </summary>
    Private ReadOnly HasFlagRegex As New Regex("^(FILE|PATH|REG)",
                                               RegexOptions.Compiled Or RegexOptions.CultureInvariant)

    ''' <summary>
    ''' Valid environment variable names for winapp2.ini paths
    ''' </summary>
    Private ReadOnly EnVars As String() = {"AllUsersProfile", "AppData", "CommonAppData", "CommonProgramFiles",
        "Documents", "HomeDrive", "LocalAppData", "LocalLowAppData", "Music", "Pictures", "ProgramData", "ProgramFiles", "Public",
        "RootDir", "SystemDrive", "SystemRoot", "Temp", "Tmp", "UserName", "UserProfile", "Video", "WinDir"}

    ''' <summary>
    ''' Case sensitive regex for an env var prefix that may be missing its percent signs:
    ''' an optional <c> % </c>, the variable name, another optional <c> % </c>, then a <c> \ </c>,
    ''' at the start of the text. <see cref="fixBrokenEnVars"/> uses this to catch a missing
    ''' leading or trailing % in one pass.
    ''' </summary>
    Private ReadOnly enVarBrokenPrefix As New Regex(
        "^(%?)(" & String.Join("|", EnVars) & ")(%?)\\",
        RegexOptions.Compiled Or RegexOptions.CultureInvariant)

    ''' <summary>
    ''' Valid key type names for winapp2.ini entries
    ''' </summary>
    Private ReadOnly ValidCmds As String() = {"Default", "DetectOS", "DetectFile", "Detect", "ExcludeKey",
        "FileKey", "LangSecRef", "RegKey", "Section", "SpecialDetect", "Warning"}

    ''' <summary>
    ''' Valid <c> SpecialDetect </c> values, properly cased. <c> chkCasing </c> uses these to
    ''' find and fix casing errors in <c> SpecialDetect </c> values.
    ''' </summary>
    Private ReadOnly SpecialDetectVals As String() = {"DET_CHROME", "DET_MOZILLA", "DET_THUNDERBIRD", "DET_OPERA"}

    ''' <summary>
    ''' Maps valid env var names to their properly cased form, case insensitively. Built once
    ''' at module init so <c> chkCasing </c> doesn't have to scan the whole array each time.
    ''' </summary>
    Private ReadOnly EnVarsLookup As Dictionary(Of String, String) = BuildCasedLookup(EnVars)

    ''' <summary>
    ''' Maps valid winapp2.ini key types to their properly cased form, case insensitively.
    ''' </summary>
    Private ReadOnly ValidCmdsLookup As Dictionary(Of String, String) = BuildCasedLookup(ValidCmds)

    ''' <summary>
    ''' Maps valid <c> SpecialDetect </c> values to their properly cased form, case insensitively.
    ''' </summary>
    Private ReadOnly SpecialDetectLookup As Dictionary(Of String, String) = BuildCasedLookup(SpecialDetectVals)

    ''' <summary>
    ''' The valid value lists, joined up ahead of time for the "Invalid data provided"
    ''' message that <c> chkCasing </c> prints.
    ''' </summary>
    Private ReadOnly EnVarsJoined As String = String.Join(", ", EnVars)
    Private ReadOnly ValidCmdsJoined As String = String.Join(", ", ValidCmds)
    Private ReadOnly SpecialDetectJoined As String = String.Join(", ", SpecialDetectVals)

    ''' <summary>
    ''' Builds the OrdinalIgnoreCase dictionary <c> chkCasing </c> uses to find the properly
    ''' cased form of a value, mapping each item in <paramref name="cased"/> to itself.
    ''' Items that differ from one already in the dictionary only by case get skipped, so
    ''' adding one later won't throw.
    ''' </summary>
    Private Function BuildCasedLookup(cased As String()) As Dictionary(Of String, String)
        Dim out As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        For Each s In cased
            If Not out.ContainsKey(s) Then out(s) = s
        Next
        Return out
    End Function

    ''' <summary>
    ''' Key types the general forward-slash check skips. RegKeys get a separate check in
    ''' <see cref="cFormat"/> that only looks at the path before the first pipe.
    ''' </summary>
    Private ReadOnly NoSlashCheckTypes As New HashSet(Of String)({"RegKey", "Section", "Warning"}, StringComparer.OrdinalIgnoreCase)

    ''' <summary>
    ''' Key types for which environment variable checks apply
    ''' </summary>
    Private ReadOnly EnVarCheckTypes As New HashSet(Of String)({"FileKey", "ExcludeKey", "DetectFile"}, StringComparer.OrdinalIgnoreCase)

    ''' <summary>
    ''' Indicates whether <see cref="InitDebug"/> stamps the version line with today's date
    ''' instead of keeping the file's own. The <c> -usedate </c> flag toggles it.
    ''' <br /> Default: <c> False </c>
    ''' </summary>
    Private Property UseCurrentDate As Boolean = False

    ''' <summary>
    ''' Handles the commandline args for <c> WinappDebug </c>. We reset the module settings to
    ''' their defaults first, so a command line run ignores the saved scan settings.
    ''' </summary>
    '''
    ''' <remarks>
    ''' Supported args: <br />
    ''' <c> -1f </c> / <c> -1d </c> input winapp2.ini (slot 1) <br />
    ''' <c> -3f </c> / <c> -3d </c> save target (slot 3) <br />
    ''' <c> -c </c> toggle saving of changes made by the linter <br />
    ''' <c> -usedate </c> toggle using the current date in the version string <br />
    ''' <c> -opti </c> enable the experimental Optimizations rule (FileKey merger) for this run <br />
    ''' <c> -keepdefaults </c> preserve existing Default keys instead of removing them (for flavors
    ''' that deliberately set Default values, such as FluentCleaner)
    ''' </remarks>
    Public Sub HandleLintCmdLine()

        InitDefaultLintSettings()

        Dim spec As New CliArgSpec(NameOf(WinappDebug))
        spec.WithFile(1, winappDebugFile1) _
            .WithFile(3, winappDebugFile3) _
            .WithFlag("-c", Sub() SaveChanges = Not SaveChanges) _
            .WithFlag("-usedate", Sub() UseCurrentDate = Not UseCurrentDate) _
            .WithFlag("-opti", Sub() lintOpti.turnOn()) _
            .WithFlag("-keepdefaults", Sub() PreserveDefaultKeys = True) _
            .Parse()

        If cmdargs.Contains("UNIT_TESTING_HALT") Then Return

        InitDebug()

    End Sub

    ''' <summary>
    ''' Lints <paramref name="givenIni"/>'s winapp2.ini formatting from outside the module's UI,
    ''' with the current scan settings, and throws the report away. Like any
    ''' <see cref="Debug"/> call, it resets <see cref="ErrorsFound"/> and
    ''' <see cref="MostRecentLintLog"/>.
    ''' <br /><br />
    '''
    ''' The lint edits <paramref name="givenIni"/>'s key objects in place, but keys it removes
    ''' stay in <paramref name="givenIni"/>'s sections, so use the returned file afterwards.
    ''' </summary>
    '''
    ''' <param name="givenIni">
    ''' The winapp2.ini syntax <c> iniFile </c> to be linted
    ''' </param>
    '''
    ''' <param name="forceOpti">
    ''' Indicates whether to turn on the Optimizations scan and repair for this call. We restore
    ''' their previous states afterwards <br /><br />
    ''' Optional, Default: <c> False </c>
    ''' </param>
    '''
    ''' <returns>
    ''' A new <c> iniFile </c> built from the linted entries, or <c> Nothing </c> if
    ''' <paramref name="givenIni"/> is <c> Nothing </c>
    ''' </returns>
    Public Function remotedebug(givenIni As iniFile,
                                Optional forceOpti As Boolean = False) As iniFile

        If givenIni Is Nothing Then argIsNull(NameOf(givenIni)) : Return Nothing

        Dim prevScan = lintOpti.ShouldScan
        Dim prevRepair = lintOpti.ShouldRepair

        If forceOpti Then
            lintOpti.ShouldScan = True
            lintOpti.ShouldRepair = True
        End If

        Dim wa2 As New winapp2file(givenIni)
        Debug(wa2)

        lintOpti.ShouldScan = prevScan
        lintOpti.ShouldRepair = prevRepair

        Return wa2.ToIni()

    End Function

    ''' <summary>
    ''' Loads <see cref="winappDebugFile1"/>, lints it, and prints the report and a summary.
    ''' When <see cref="SaveChanges"/> is on, writes the linted file to
    ''' <see cref="winappDebugFile3"/>. Returns without linting when the file is missing or
    ''' has no sections.
    ''' </summary>
    Public Sub InitDebug()

        Dim inputFile = winappDebugFile1.Load(LintModuleSettingsChanged, NameOf(WinappDebug), NameOf(winappDebugFile1), NameOf(LintModuleSettingsChanged))

        If Not enforceFileHasContent(inputFile) Then Return

        Dim wa2 As New winapp2file(inputFile, UseCurrentDate)

        clrConsole()
        gLog("")

        Dim output As New List(Of MenuSection)

        Dim header As New MenuSection()
        header.AddTopBorder() _
              .AddColoredLine("Linting winapp2.ini", ConsoleColor.Cyan, centered:=True) _
              .AddDivider(solid:=False)
        output.Add(header)

        Dim lintSw = Stopwatch.StartNew()

        Using gLogScope("Beginning lint")

            gLog("")
            output.AddRange(Debug(wa2))

        End Using

        lintSw.Stop()

        gLog("Lint complete")
        gLog($"Entry count: {wa2.Count}")
        gLog($"{ErrorsFound} errors detected")
        gLog($"{lintSw.ElapsedMilliseconds} ms")
        setNextMenuHeaderText("Lint complete", printColor:=ConsoleColor.Green)

        Dim errColor = If(ErrorsFound = 0, ConsoleColor.Green, If(ErrorsFound < 10, ConsoleColor.DarkYellow, ConsoleColor.DarkRed))
        Dim summary As New MenuSection()
        summary.AddDivider(solid:=False) _
               .AddColoredLine("Lint Complete!", ConsoleColor.Green, centered:=True) _
               .AddDivider(solid:=False) _
               .AddLine($"Entry count: {wa2.Count}", centered:=True) _
               .AddColoredLine($"{ErrorsFound} possible errors detected.", errColor, centered:=True)

        If SaveChanges Then

            If iniFile.Empty(winappDebugFile3.Dir, winappDebugFile3.Name).OverwriteToFile(wa2.ToWinapp2String()) Then

                summary.AddColoredLine($"{winappDebugFile3.Name} saved with any corrections made", ConsoleColor.DarkGreen, centered:=True)

            Else

                summary.AddColoredLine($"{winappDebugFile3.Name} was not saved", ConsoleColor.Red, centered:=True)

            End If

        End If

        summary.AddBlank.AddColoredLine(anyKeyStr, ConsoleColor.Cyan, centered:=True)
        summary.AddBottomBorder()

        output.Add(summary)

        clrConsole()
        output.ForEach(Sub(s) s.Print(withDivider:=False))

        crk()

    End Sub

    ''' <summary>
    ''' Lints every entry of <paramref name="fileToBeDebugged"/> in parallel, then checks entry
    ''' order within each category. Enabled repairs edit the entries in place. Resets
    ''' <see cref="ErrorsFound"/> and <see cref="MostRecentLintLog"/> before counting this run.
    ''' </summary>
    '''
    ''' <param name="fileToBeDebugged">
    ''' A <c> winapp2file </c> to be linted. It's passed <c> ByRef </c> but never reassigned
    ''' </param>
    '''
    ''' <returns>
    ''' The report in render order: each entry's errors followed by its merge block, then the
    ''' entry alphabetization errors. Empty if <paramref name="fileToBeDebugged"/> is <c> Nothing </c>
    ''' </returns>
    Public Function Debug(ByRef fileToBeDebugged As winapp2file) As List(Of MenuSection)

        If fileToBeDebugged Is Nothing Then argIsNull(NameOf(fileToBeDebugged)) : Return New List(Of MenuSection)

        ErrorsFound = 0
        MostRecentLintLog.Clear()

        Dim output As New List(Of MenuSection)

        Dim duplicateNames = FindDuplicateEntryNames(fileToBeDebugged)

        Dim results = fileToBeDebugged.Entries _
            .AsParallel() _
            .AsOrdered() _
            .Select(Function(entry) ProcessEntry(entry, duplicateNames)) _
            .ToList()

        For Each result In results : output.AddRange(EmitEntryResult(result)) : Next

        output.AddRange(AlphabetizeEntries(fileToBeDebugged))

        Return output

    End Function

    ''' <summary>
    ''' Returns the set of entry names that appear more than once in a <c> winapp2file </c>,
    ''' ignoring case. <c> iniFile </c> drops repeated sections while parsing, so a file read
    ''' from disk never has any.
    ''' </summary>
    '''
    ''' <param name="winapp">
    ''' The <c> winapp2file </c> whose entries will be scanned for duplicate names
    ''' </param>
    Private Function FindDuplicateEntryNames(winapp As winapp2file) As HashSet(Of String)

        Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Dim duplicates As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        For Each entry In winapp.Entries
            If Not seen.Add(entry.Name) Then duplicates.Add(entry.Name)
        Next

        Return duplicates

    End Function

    ''' <summary>
    ''' Collects the errors from an <c> EntryLintResult </c> into a <c> MenuSection </c>, logs them
    ''' along with the result's captured log lines, appends them to <see cref="MostRecentLintLog"/>,
    ''' and adds its error count to <see cref="ErrorsFound"/>. The error section comes back
    ''' followed by the result's deferred sections, to be rendered later.
    ''' </summary>
    '''
    ''' <param name="result">
    ''' The <c> EntryLintResult </c> whose errors will be collected
    ''' </param>
    Private Function EmitEntryResult(result As EntryLintResult) As List(Of MenuSection)

        Dim sections As New List(Of MenuSection)

        ErrorsFound += result.ErrorCount

        If result.LogLines.Count = 0 AndAlso result.ErrorCount = 0 Then

            sections.AddRange(result.DeferredSections)
            Return sections

        End If

        Using gLogScope($"Processing {result.EntryName}")
            gLog("")

            EmitCaptured(result.LogLines)

            If result.ErrorCount > 0 Then

                Dim section As New MenuSection()

                For Each Errr In result.Errors

                    Dim out = $"Error in {result.EntryName}:"
                    Dim out2 = $"{Errr.Message}"
                    gLog(out)
                    section.AddColoredLine(out, ConsoleColor.Red).AddColoredLine(out2, ConsoleColor.DarkYellow)
                    MostRecentLintLog.AppendLine(out).AppendLine(out2)

                    Using gLogScope(out2)

                        For Each detail In Errr.Details
                            gLog($"{detail}")
                            section.AddColoredLine($"{detail}", ConsoleColor.Yellow)
                            MostRecentLintLog.AppendLine(detail)
                        Next

                    End Using

                    section.AddBlank()
                    MostRecentLintLog.AppendLine()

                Next

                sections.Add(section)

            End If

            gLog("")

        End Using

        sections.AddRange(result.DeferredSections)
        Return sections

    End Function

    ''' <summary>
    ''' Lints one entry: its name, then every key (<see cref="ValidateKeys"/>), then each typed
    ''' bucket in winapp2.ini key order (<see cref="processKeyList"/>), then the entry's overall
    ''' structure and the Defaults rule. Edits <paramref name="entry"/> in place and captures
    ''' the log lines it writes into the result.
    ''' <br /><br />
    '''
    ''' A duplicate name, a name without the <c> * </c> suffix, and a missing detection key are
    ''' reported whatever the scan settings say.
    ''' While <see cref="overrideDefaultVal"/> is on, a missing Default key is reported and added
    ''' whatever the Defaults rule says.
    ''' </summary>
    '''
    ''' <param name="entry">
    ''' A <c> winapp2entry </c> to be audited for syntax errors
    ''' </param>
    '''
    ''' <param name="duplicateNames">The entry names that appear more than once in the file</param>
    Private Function ProcessEntry(entry As winapp2entry,
                                  duplicateNames As HashSet(Of String)) As EntryLintResult

        Dim result As New EntryLintResult(entry.FullName)

        Using cap = gLogCapture()

            Dim hasFileExcludes = False
            Dim hasRegExcludes = False

            result.RecordError("Duplicate entry name detected", Array.Empty(Of String)(), duplicateNames.Contains(entry.Name))
            result.RecordError("All entries must end in ' *'", Array.Empty(Of String)(), Not entry.HasValidNameSuffix)

            ValidateKeys(result, entry)

            Dim bc = Function(typeName As String) As Integer
                         Dim idx = winapp2entry.GetBucketIndex(typeName)
                         Return If(idx >= 0, entry.KeyLists(idx).Count, 0)
                     End Function

            processKeyList(result, entry, New KeyListSpec("DetectOS", noNumbers:=True, oneOnly:=True), AddressOf voidDelegate)
            processKeyList(result, entry, New KeyListSpec("LangSecRef", noNumbers:=True, oneOnly:=True), AddressOf voidDelegate)
            processKeyList(result, entry, New KeyListSpec("Section", noNumbers:=True, oneOnly:=True), AddressOf voidDelegate)
            processKeyList(result, entry, New KeyListSpec("SpecialDetect", noNumbers:=True, oneOnly:=True), AddressOf voidDelegate)
            processKeyList(result, entry, New KeyListSpec("Detect", noNumbers:=(bc("Detect") = 1), checkPathValidity:=True, isRegistryPath:=True), AddressOf voidDelegate)
            processKeyList(result, entry, New KeyListSpec("DetectFile", noNumbers:=(bc("DetectFile") = 1)), Function(k) pDetectFile(result, k))
            processKeyList(result, entry, New KeyListSpec("Default", noNumbers:=True, oneOnly:=True), AddressOf voidDelegate)
            processKeyList(result, entry, New KeyListSpec("Warning", noNumbers:=True, oneOnly:=True), AddressOf voidDelegate)
            processKeyList(result, entry, New KeyListSpec("FileKey"), Function(k) pFileKey(result, k))
            processKeyList(result, entry, New KeyListSpec("RegKey", checkPathValidity:=True, isRegistryPath:=True), AddressOf voidDelegate)
            processKeyList(result, entry, New KeyListSpec("ExcludeKey", isExcludeKey:=True), AddressOf voidDelegate, hasFileExcludes, hasRegExcludes)

            Dim hasSectionKey = entry.SectionKey.Count <> 0
            Dim hasLangSecRef = entry.LangSecRef.Count <> 0
            Dim hasDetectFiles = entry.DetectFiles.Count <> 0
            Dim hasDetects = entry.Detects.Count <> 0
            Dim hasDetectOS = entry.DetectOS.Count <> 0
            Dim hasSpecialDetect = entry.SpecialDetect.Count <> 0
            Dim hasFileKeys = entry.FileKeys.Count <> 0
            Dim hasRegKeys = entry.RegKeys.Count <> 0
            Dim hasDefaultKey = entry.DefaultKey.Count > 0

            result.RecordError("Section key found alongside LangSecRef key, but only one should be present", Array.Empty(Of String)(), lintSyntax.ShouldScan AndAlso hasSectionKey AndAlso hasLangSecRef)
            result.RecordError("Entry has no valid classifier key (LangSecRef, Section)", Array.Empty(Of String)(), lintSyntax.ShouldScan AndAlso Not (hasSectionKey OrElse hasLangSecRef))
            result.RecordError("Entry has no valid detection keys (Detect, DetectFile, DetectOS, SpecialDetect)", Array.Empty(Of String)(), Not (hasDetectFiles OrElse hasDetects OrElse hasDetectOS OrElse hasSpecialDetect))
            result.RecordError("Entry has no valid deletion keys (FileKey, RegKey)", Array.Empty(Of String)(), lintSyntax.ShouldScan AndAlso Not (hasFileKeys OrElse hasRegKeys))
            result.RecordError("Entry has ExcludeKeys but no valid FileKeys or RegKeys", Array.Empty(Of String)(), lintSyntax.ShouldScan AndAlso hasFileExcludes AndAlso Not (hasFileKeys OrElse hasRegKeys))
            result.RecordError("Entry has ExcludeKeys pointing to file system locations but no FileKeys", Array.Empty(Of String)(), hasFileExcludes AndAlso Not hasFileKeys)
            result.RecordError("Entry has ExcludeKeys pointing to registry locations but no RegKeys", Array.Empty(Of String)(), hasRegExcludes AndAlso Not hasRegKeys)
            result.RecordError("Entry has a Default key where there should be none", Array.Empty(Of String)(), lintDefaults.ShouldScan AndAlso hasDefaultKey AndAlso Not overrideDefaultVal AndAlso Not PreserveDefaultKeys)

            If lintDefaults.fixFormat AndAlso hasDefaultKey AndAlso Not overrideDefaultVal AndAlso Not PreserveDefaultKeys Then

                For Each k In entry.DefaultKey.ToList() : entry.RemoveKey(k) : Next

            End If

            If overrideDefaultVal Then

                Dim expected = tsInvariant(expectedDefaultValue)

                If entry.DefaultKey.Count > 0 Then

                    Dim key = entry.DefaultKey(0)
                    fullKeyErr(result, key, "Incorrect value for Default Key found", lintDefaults.ShouldScan AndAlso Not key.Value = expected, lintDefaults.fixFormat, key.Value, expected)

                Else

                    result.RecordError("No Default Key found", Array.Empty(Of String)())
                    entry.AddKey(New iniKey($"Default={expected}"))

                End If

            End If

            result.LogLines = New List(Of String)(cap.Lines)

        End Using

        Return result

    End Function

    ''' <summary>
    ''' Runs <see cref="cValidity"/> over every <c> iniKey </c> in a <c> winapp2entry </c>,
    ''' removes the keys it rejects, and moves any unrecognized key whose type a repair made
    ''' valid into its typed bucket. Unrecognized keys that stay unrecognized are kept.
    ''' </summary>
    '''
    ''' <param name="result">The lint result that collects the errors</param>
    '''
    ''' <param name="entry">
    ''' A <c> winapp2entry </c> whose <c> iniKey </c>s will be audited for basic syntax correctness
    ''' </param>
    Private Sub ValidateKeys(result As EntryLintResult, entry As winapp2entry)

        ' Run cValidity over the normal buckets and drop anything that fails
        For i = 0 To entry.KeyLists.Count - 2
            For Each key In entry.KeyLists(i).Where(Function(k) Not cValidity(result, k)).ToList()
                entry.RemoveKey(key)
            Next
        Next

        ' Same for the error keys, except these have to be force removed. A partial repair
        ' may have changed the KeyType out from under the normal routing
        For Each key In entry.ErrorKeys.Where(Function(k) Not cValidity(result, k)).ToList()
            entry.ForceRemoveErrorKey(key)
        Next

        ' Anything we managed to repair into a real winapp2.ini key type goes back in a bucket
        entry.ReclassifyErrorKeys()

    End Sub

    ''' <summary>
    ''' Reports entries that are out of alphabetical order within their category, under the
    ''' Alphabetization scan, and sorts every category when the Alphabetization repair is on.
    ''' Hyphens sort as spaces.
    ''' </summary>
    '''
    ''' <param name="winapp">
    ''' The <c> winapp2file </c> whose entries will be alphabetized
    ''' </param>
    Private Function AlphabetizeEntries(winapp As winapp2file) As List(Of MenuSection)

        Dim sections As New List(Of MenuSection)

        For Each category In winapp.Categories

            If category.Count < 2 Then Continue For

            Dim unsortedNames As New List(Of String)
            For i = 0 To category.Count - 1
                unsortedNames.Add(category(i).Name)
            Next

            Dim sortedNames = replaceAndSort(unsortedNames, "-", "  ")

            If lintAlpha.ShouldScan Then
                Dim alreadySorted = True
                For i = 0 To unsortedNames.Count - 1
                    If Not unsortedNames(i).Equals(sortedNames(i), StringComparison.Ordinal) Then
                        alreadySorted = False
                        Exit For
                    End If
                Next
                If Not alreadySorted Then sections.AddRange(EmitEntryAlphabetizationErrors(findOutOfPlace(unsortedNames, sortedNames)))
            End If

        Next

        If lintAlpha.fixFormat Then winapp.SortEntries()

        Return sections

    End Function

    ''' <summary>
    ''' Renders entry level alphabetization misplacements into <c> MenuSection </c>s and the
    ''' global log, and bumps <c> ErrorsFound </c> for each one
    ''' </summary>
    '''
    ''' <param name="misplacements">
    ''' The misplaced entries, as returned by <c> findOutOfPlace </c>
    ''' </param>
    Private Function EmitEntryAlphabetizationErrors(misplacements As List(Of AlphaMisplacement)) As List(Of MenuSection)

        Dim sections As New List(Of MenuSection)
        If misplacements.Count = 0 Then Return sections

        Dim section As New MenuSection()

        For Each m In misplacements

            ErrorsFound += 1

            Dim out = $"Error in {m.Item}:"
            Dim out2 = "Entry alphabetization"
            gLog(out)
            section.AddColoredLine(out, ConsoleColor.Red).AddColoredLine(out2, ConsoleColor.DarkYellow)
            MostRecentLintLog.AppendLine(out).AppendLine(out2)

            Using gLogScope(out2)

                Dim d1 = $"{m.Item} appears to be out of place"
                Dim d2 = $"Expected position: {m.ExpectedPos + 1}"
                gLog(d1)
                gLog(d2)
                section.AddColoredLine(d1, ConsoleColor.Yellow).AddColoredLine(d2, ConsoleColor.Yellow)
                MostRecentLintLog.AppendLine(d1).AppendLine(d2)

            End Using

            section.AddBlank()
            MostRecentLintLog.AppendLine()

        Next

        sections.Add(section)
        Return sections

    End Function

    ''' <summary>
    ''' One out-of-place item discovered by <c> findOutOfPlace </c>
    ''' </summary>
    Private Structure AlphaMisplacement

        ''' <summary>
        ''' The raw item value (entry name or key value) that is out of place
        ''' </summary>
        Public ReadOnly Property Item As String

        ''' <summary>
        ''' The 0-based index of the first item in the unsorted list with this value
        ''' </summary>
        Public ReadOnly Property ActualPos As Integer

        ''' <summary>
        ''' The 0-based index the item should occupy in the sorted list
        ''' </summary>
        Public ReadOnly Property ExpectedPos As Integer

        ''' <summary>Creates a new <c> AlphaMisplacement </c></summary>
        ''' <param name="item">The out-of-place value</param>
        ''' <param name="actualPos">Its 0-based index in the unsorted list</param>
        ''' <param name="expectedPos">Its 0-based index in the sorted list</param>
        Public Sub New(item As String, actualPos As Integer, expectedPos As Integer)
            Me.Item = item
            Me.ActualPos = actualPos
            Me.ExpectedPos = expectedPos
        End Sub

    End Structure

    ''' <summary>
    ''' Returns the items from <paramref name="someList"/> that are out of order with respect
    ''' to <paramref name="sortedList"/>, paired with their actual and expected positions.
    ''' It's up to the caller to report them.
    ''' </summary>
    '''
    ''' <param name="someList">
    ''' An unsorted list of strings (iniKey values or iniSection names)
    ''' </param>
    '''
    ''' <param name="sortedList">
    ''' The sorted state of <paramref name="someList"/>
    ''' </param>
    Private Function findOutOfPlace(someList As List(Of String),
                                    sortedList As List(Of String)) As List(Of AlphaMisplacement)

        Dim findings As New List(Of AlphaMisplacement)
        If someList.Count < 2 Then Return findings

        Dim sortedIndices As New Dictionary(Of String, Integer)
        For i = 0 To sortedList.Count - 1
            sortedIndices(sortedList(i)) = i
        Next

        ' Line up each item's position in the sorted list, then work out which of the
        ' actual-list indices form the LIS. Whatever falls outside the LIS is the smallest
        ' set of things actually out of place. Everything else is just getting shoved
        ' around by those, so there's no point reporting it.
        Dim sortedPosSequence As New List(Of Integer)
        For Each item In someList
            sortedPosSequence.Add(sortedIndices(item))
        Next

        Dim lisIndices = FindLISIndices(sortedPosSequence)

        For i = 0 To someList.Count - 1

            If lisIndices.Contains(i) Then Continue For

            Dim item = someList(i)
            Dim recInd = someList.IndexOf(item)
            Dim sortInd = sortedIndices(item)

            If recInd = sortInd Then Continue For

            findings.Add(New AlphaMisplacement(item, recInd, sortInd))

        Next

        Return findings

    End Function

    ''' <summary>
    ''' Returns the set of indices (into <paramref name="sequence"/>) that form its Longest
    ''' Increasing Subsequence. Anything whose index is missing from that set is out of order.
    ''' </summary>
    '''
    ''' <param name="sequence">
    ''' A sequence of integers representing the sorted-list position of each item in the actual list
    ''' </param>
    Private Function FindLISIndices(sequence As List(Of Integer)) As HashSet(Of Integer)

        Dim n = sequence.Count
        If n = 0 Then Return New HashSet(Of Integer)

        ' Patience sort: tails(i) = smallest tail value of any IS of length i+1
        Dim tails As New List(Of Integer)
        Dim tailPos As New List(Of Integer)    ' actual-list index of each tail
        Dim parent As Integer() = New Integer(n - 1) {}
        For i = 0 To n - 1
            parent(i) = -1
        Next

        For i = 0 To n - 1

            Dim val = sequence(i)

            Dim lo = 0, hi = tails.Count
            While lo < hi
                Dim mid = (lo + hi) \ 2
                If tails(mid) < val Then lo = mid + 1 Else hi = mid
            End While

            parent(i) = If(lo > 0, tailPos(lo - 1), -1)

            If lo = tails.Count Then
                tails.Add(val)
                tailPos.Add(i)
            Else
                tails(lo) = val
                tailPos(lo) = i
            End If

        Next

        ' Backtrack from the tail of the longest IS to recover member indices
        Dim lisIndices As New HashSet(Of Integer)
        Dim cur = tailPos(tails.Count - 1)
        While cur >= 0
            lisIndices.Add(cur)
            cur = parent(cur)
        End While

        Return lisIndices

    End Function

    ''' <summary>
    ''' The per-call configuration for <see cref="processKeyList"/>, holding the behavior that
    ''' varies by key type.
    ''' </summary>
    Private Structure KeyListSpec

        ''' <summary>The key type whose bucket is processed</summary>
        Public ReadOnly Property TypeName As String
        ''' <summary>Indicates whether keys of this type should carry no number</summary>
        Public ReadOnly Property NoNumbers As Boolean
        ''' <summary>Indicates whether an entry may hold only one key of this type</summary>
        Public ReadOnly Property OneOnly As Boolean
        ''' <summary>Indicates whether each key's path goes through <see cref="chkPathFormatValidity"/></summary>
        Public ReadOnly Property CheckPathValidity As Boolean
        ''' <summary>Indicates whether that path check expects a registry path</summary>
        Public ReadOnly Property IsRegistryPath As Boolean
        ''' <summary>Indicates whether each key also goes through <see cref="pExcludeKey"/></summary>
        Public ReadOnly Property IsExcludeKey As Boolean

        ''' <summary>Creates a new <c> KeyListSpec </c>. Every option defaults to <c> False </c>.</summary>
        ''' <param name="typeName">The key type whose bucket is processed</param>
        ''' <param name="noNumbers">Indicates whether keys of this type should carry no number</param>
        ''' <param name="oneOnly">Indicates whether an entry may hold only one key of this type</param>
        ''' <param name="checkPathValidity">Indicates whether to check each key's path</param>
        ''' <param name="isRegistryPath">Indicates whether that path is a registry path</param>
        ''' <param name="isExcludeKey">Indicates whether to run the ExcludeKey checks</param>
        Public Sub New(typeName As String,
                       Optional noNumbers As Boolean = False,
                       Optional oneOnly As Boolean = False,
                       Optional checkPathValidity As Boolean = False,
                       Optional isRegistryPath As Boolean = False,
                       Optional isExcludeKey As Boolean = False)
            Me.TypeName = typeName
            Me.NoNumbers = noNumbers
            Me.OneOnly = oneOnly
            Me.CheckPathValidity = checkPathValidity
            Me.IsRegistryPath = isRegistryPath
            Me.IsExcludeKey = isExcludeKey
        End Sub

    End Structure

    ''' <summary>
    ''' Runs the per-key checks over one typed bucket of <paramref name="entry"/>, removes the
    ''' duplicates and multiples marked for removal, then hands the bucket to
    ''' <see cref="sortKeys"/>. For the FileKey bucket, also runs the Optimizations merge
    ''' while that scan is on.
    ''' </summary>
    '''
    ''' <param name="result">The lint result that collects the errors</param>
    '''
    ''' <param name="entry">
    ''' The <c> winapp2entry </c> whose keys are being processed
    ''' </param>
    '''
    ''' <param name="spec">
    ''' Configuration encoding the type-specific behavior for this bucket
    ''' </param>
    '''
    ''' <param name="processKey">
    ''' The type-specific check for each key, or <see cref="voidDelegate"/> if there isn't one.
    ''' We ignore what it returns, so it has to change the key in place.
    ''' </param>
    '''
    ''' <param name="hasF">
    ''' Set to <c> True </c> if an ExcludeKey in the bucket excludes a file system location <br /><br />
    ''' Optional, Default: <c> False </c>
    ''' </param>
    '''
    ''' <param name="hasR">
    ''' Set to <c> True </c> if an ExcludeKey in the bucket excludes a registry location <br /><br />
    ''' Optional, Default: <c> False </c>
    ''' </param>
    Private Sub processKeyList(result As EntryLintResult,
                               entry As winapp2entry,
                               spec As KeyListSpec,
                               processKey As Func(Of iniKey, iniKey),
                               Optional ByRef hasF As Boolean = False,
                               Optional ByRef hasR As Boolean = False)

        Dim bucketIdx = winapp2entry.GetBucketIndex(spec.TypeName)
        If bucketIdx < 0 Then Return

        Dim bucket = entry.KeyLists(bucketIdx)

        If bucket.Count = 0 Then Return

        Dim curNum = 1
        ' Don't bother allocating the dupe tracker for a bucket holding one key, since there's
        ' nothing to duplicate, or when the user has turned off both the scan and the repair
        ' for duplicate values. 
        Dim seenValues As Dictionary(Of String, iniKey) = Nothing
        Dim dupeKeys As List(Of iniKey) = Nothing  ' lazily allocated on first duplicate
        If bucket.Count > 1 AndAlso (lintDupes.ShouldScan OrElse lintDupes.fixFormat) Then
            seenValues = New Dictionary(Of String, iniKey)(StringComparer.OrdinalIgnoreCase)
        End If

        For Each key In bucket

            If spec.CheckPathValidity Then chkPathFormatValidity(result, key, spec.IsRegistryPath)

            If spec.OneOnly AndAlso curNum > 1 AndAlso lintMulti.ShouldScan Then
                fullKeyErr(result, key, $"Multiple {key.KeyType} detected.")
                If lintMulti.fixFormat Then
                    If dupeKeys Is Nothing Then dupeKeys = New List(Of iniKey)
                    dupeKeys.Add(key)
                End If
            End If

            cFormat(result, key, curNum, seenValues, dupeKeys, spec.NoNumbers)

            If key.typeIs("SpecialDetect") Then chkCasing(result, key, SpecialDetectLookup, SpecialDetectJoined, key.Value)
            fullKeyErr(result, key, "LangSecRef holds an invalid value.", lintInvalid.ShouldScan AndAlso key.typeIs("LangSecRef") AndAlso Not secRefNums.IsMatch(key.Value))

            If spec.IsExcludeKey Then pExcludeKey(result, key, hasF, hasR)

            key = processKey(key)

        Next

        Dim removedDupes = dupeKeys IsNot Nothing AndAlso dupeKeys.Count > 0
        If removedDupes Then
            For Each dupe In dupeKeys
                entry.RemoveKey(dupe)
            Next
        End If

        sortKeys(result, entry, spec.TypeName, removedDupes)

        If spec.TypeName = "FileKey" AndAlso lintOpti.ShouldScan Then cOptimization(result, entry)

    End Sub

    ''' <summary>
    ''' Returns <paramref name="key"/> unchanged, for the key types that have no type-specific
    ''' check in <see cref="processKeyList"/>
    ''' </summary>
    '''
    ''' <param name="key">
    ''' An <c> iniKey </c> with which to do nothing
    ''' </param>
    Private Function voidDelegate(key As iniKey) As iniKey

        Return key

    End Function

    ''' <summary>
    ''' Runs the checks every key in a typed bucket gets: a value that repeats an earlier key's,
    ''' ignoring case (Duplicates), the key's number (Improper Numbering or Unneeded Numbering),
    ''' forward slashes (Slashes), and a trailing semicolon (Semicolons). FileKeys, ExcludeKeys
    ''' and DetectFiles also go through <see cref="cEnVar"/>.
    ''' </summary>
    '''
    ''' <param name="result">The lint result that collects the errors</param>
    '''
    ''' <param name="key">
    ''' An <c> iniKey </c> whose format will be audited
    ''' </param>
    '''
    ''' <param name="keyNumber">
    ''' The number this key should carry. We increment it for the next key
    ''' </param>
    '''
    ''' <param name="seenValues">
    ''' A map of the key values we've seen so far to the first <c> iniKey </c> that held each
    ''' one, used to spot duplicates. <c> Nothing </c> skips the duplicate check
    ''' </param>
    '''
    ''' <param name="dupeKeys">
    ''' The keys marked for removal. We create it on the first duplicate if it's <c> Nothing </c>
    ''' </param>
    '''
    ''' <param name="noNumbers">
    ''' Indicates whether the current set of keys should not be numbered <br /><br />
    ''' Optional, Default: <c> False </c>
    ''' </param>
    Private Sub cFormat(result As EntryLintResult,
                        key As iniKey,
                        ByRef keyNumber As Integer,
                        ByRef seenValues As Dictionary(Of String, iniKey),
                        ByRef dupeKeys As List(Of iniKey),
                        Optional noNumbers As Boolean = False)

        If seenValues IsNot Nothing Then

            Dim firstSeen As iniKey = Nothing
            If seenValues.TryGetValue(key.Value, firstSeen) Then

                result.RecordError("Duplicate key value found", {$"Key:            {key.ToString()}", $"Duplicates:     {firstSeen.ToString()}"}, lintDupes.ShouldScan)
                If lintDupes.fixFormat Then
                    If dupeKeys Is Nothing Then dupeKeys = New List(Of iniKey)
                    dupeKeys.Add(key)
                End If

            Else

                seenValues(key.Value) = key

            End If

        End If

        ' Check for both types of numbering errors (incorrect and unneeded)
        Dim hasNumberingError = If(noNumbers, Not key.nameIs(key.KeyType), Not key.nameIs(key.KeyType & keyNumber))

        If hasNumberingError Then
            Dim numberingErrStr = If(noNumbers, "Detected unnecessary numbering.", $"{key.KeyType} entry is incorrectly numbered.")
            Dim fixedStr = If(noNumbers, key.KeyType, key.KeyType & keyNumber)
            gLog($"Input mismatch error in {key.ToString()}")
            inputMismatchErr(result, numberingErrStr, key.Name, fixedStr, If(noNumbers, lintExtraNums.ShouldScan, lintWrongNums.ShouldScan))
            fixStr(If(noNumbers, lintExtraNums.fixFormat, lintWrongNums.fixFormat), key.Name, fixedStr)
        End If

        ' Scan for and fix any use of incorrect slashes or trailing semicolons
        Dim fwdSlashErr = "Forward slash (/) detected in lieu of backslash (\)."
        fullKeyErr(result, key, fwdSlashErr, Not NoSlashCheckTypes.Contains(key.KeyType) AndAlso lintSlashes.ShouldScan AndAlso key.vHas("/"),
                                             lintSlashes.fixFormat, key.Value, Function() key.Value.Replace("/", "\"))
        If key.typeIs("RegKey") AndAlso lintSlashes.ShouldScan AndAlso key.vHas("/") Then
            Dim pipeIdx = key.Value.IndexOf("|"c)
            Dim pathPart = If(pipeIdx >= 0, key.Value.Substring(0, pipeIdx), key.Value)
            If pathPart.Contains("/") Then
                fullKeyErr(result, key, fwdSlashErr, True, lintSlashes.fixFormat, key.Value,
                           Function() pathPart.Replace("/", "\") & If(pipeIdx >= 0, key.Value.Substring(pipeIdx), ""))
            End If
        End If
        fullKeyErr(result, key, "Trailing semicolon (;).", key.Value.Length > 0 AndAlso key.Value(key.Value.Length - 1) = ";"c AndAlso lintSemis.ShouldScan, lintSemis.fixFormat, key.Value, Function() key.Value.TrimEnd(";"c))

        ' Do some formatting checks for environment variables if needed
        If EnVarCheckTypes.Contains(key.KeyType) Then cEnVar(result, key)
        keyNumber += 1

    End Sub

    ''' <summary>
    ''' Reports a known environment variable at the start of the key's path that is missing its
    ''' leading <c> % </c>, its trailing one, or both, and adds them back. An ExcludeKey's path
    ''' starts after its first pipe. The report follows the Syntax Errors scan, and the repair
    ''' follows that rule's <c> ShouldRepair </c> directly rather than <see cref="lintRule.fixFormat"/>.
    ''' </summary>
    '''
    ''' <param name="result">The lint result that collects the errors</param>
    '''
    ''' <param name="key">
    ''' An <c> iniKey </c> whose value will be audited for syntax errors
    ''' </param>
    '''
    ''' <param name="enVars">
    ''' Unused. The check goes through <see cref="enVarBrokenPrefix"/> instead
    ''' </param>
    '''
    ''' <param name="cond">
    ''' Indicates whether to run the check
    ''' </param>
    Private Sub fixBrokenEnVars(result As EntryLintResult, key As iniKey, enVars As String(), cond As Boolean)

        If Not cond Then Return

        Dim pipeInd = If(key.typeIs("ExcludeKey"), key.Value.IndexOf("|"c), -1)
        Dim pathPart = If(pipeInd >= 0, key.Value.Substring(pipeInd + 1), key.Value)

        Dim m = enVarBrokenPrefix.Match(pathPart)
        If Not m.Success Then Return

        Dim hasLeading = m.Groups(1).Value = "%"
        Dim enVar = m.Groups(2).Value
        Dim hasTrailing = m.Groups(3).Value = "%"

        If hasLeading AndAlso hasTrailing Then Return

        Dim msg As String
        Dim brokenForm As String

        If hasLeading Then
            msg = "Environment Variable is missing trailing %"
            brokenForm = $"%{enVar}\"
        ElseIf hasTrailing Then
            msg = "Environment Variable is missing leading %"
            brokenForm = $"{enVar}%\"
        Else
            msg = "Environment Variable is missing leading and trailing %"
            brokenForm = $"{enVar}\"
        End If

        Dim fixedPath = pathPart.Replace(brokenForm, $"%{enVar}%\")
        Dim fixedValue = If(pipeInd >= 0, key.Value.Substring(0, pipeInd + 1) & fixedPath, fixedPath)
        fullKeyErr(result, key, msg, lintSyntax.ShouldScan, lintSyntax.ShouldRepair, key.Value, fixedValue)

    End Sub

    ''' <summary>
    ''' Validates the formatting of any %EnvironmentVariables% in a given <c> iniKey </c>: a
    ''' doubled <c> %% </c> (reported always, repaired under Syntax Errors), missing percent signs
    ''' (<see cref="fixBrokenEnVars"/>), the casing and validity of each variable name
    ''' (<see cref="chkCasing"/>), and a variable with no backslash after it (Slashes, report only)
    ''' </summary>
    '''
    ''' <param name="result">The lint result that collects the errors</param>
    '''
    ''' <param name="key">
    ''' The <c> iniKey </c> whose data will be audited for environment variable correctness
    ''' </param>
    Private Sub cEnVar(result As EntryLintResult, key As iniKey)

        fullKeyErr(result, key, "Double '%' found in environment variable", key.vHas("%%"), lintSyntax.ShouldRepair, key.Value, Function() key.Value.Replace("%%", "%"))

        Dim envMatches = envVarRegex.Matches(key.Value)
        fixBrokenEnVars(result, key, EnVars, lintSyntax.ShouldScan AndAlso key.vHas("%") AndAlso envMatches.Count = 0 OrElse key.vHasAny(EnVars) AndAlso Not key.vHas("%"))

        For Each m As Match In envMatches

            Dim strippedText = m.ToString.Trim(CChar("%"))
            chkCasing(result, key, EnVarsLookup, EnVarsJoined, strippedText)

        Next

        ' Environment variables should be trailed by a backslash
        fullKeyErr(result, key, "Missing backslash (\) after %EnvironmentVariable%.", lintSlashes.ShouldScan And key.vHas("%") And Not key.vHasAny({"%|", "%\"}))

    End Sub

    ''' <summary>
    ''' Splits a key that has no <c> = </c> into a name and a value. The type is the first entry
    ''' in <paramref name="cmds"/> that the key's name starts with, ignoring case. When the name
    ''' spells the type with the same casing, the new name is the type plus any digits that
    ''' follow it (none for types that take no number), and the rest becomes the value. The strip
    ''' is case-sensitive and removes every occurrence, not just the prefix, so with different
    ''' casing nothing is stripped: <c> filekey1%AppData%\Test </c> becomes
    ''' <c> FileKey=filekey1%AppData%\Test </c>. Once a type matches, we change
    ''' <paramref name="key"/> even if the repair then fails.
    ''' </summary>
    '''
    ''' <param name="result">
    ''' Unused
    ''' </param>
    '''
    ''' <param name="key">
    ''' A misformatted <c> iniKey </c> to attempt to repair
    ''' </param>
    '''
    ''' <param name="cmds">
    ''' An array containing valid winapp2.ini <c> keyTypes </c>
    ''' </param>
    '''
    ''' <returns>
    ''' <c> True </c> if the key now has a value, <c> False </c> if no type matched or
    ''' nothing was left for the value
    ''' </returns>
    Private Function fixMissingEquals(result As EntryLintResult,
                                      key As iniKey,
                                      cmds As String()) As Boolean

        gLog("Attempting missing equals repair")

        For Each cmd In cmds

            If Not key.Name.ToUpperInvariant.StartsWith(cmd.ToUpperInvariant) Then Continue For

            Select Case cmd

                ' We don't expect numbers in these keys
                Case "Default", "DetectOS", "Section", "LangSecRef", "Section", "SpecialDetect"

                    key.Value = key.Name.Replace(cmd, "")
                    key.Name = cmd

                Case Else

                    Dim newName = cmd
                    Dim withNums = key.Name.Replace(cmd, "")

                    For Each c As Char In withNums.ToCharArray

                        If Char.IsNumber(c) Then newName += c : Else Exit For

                    Next

                    key.Value = key.Name.Replace(newName, "")
                    key.Name = newName

            End Select

            gLog($"Repair complete. Result: {key.ToString()}")

            ' Don't allow valueless keys in winapp2.ini
            If key.Value.Length = 0 Then gLog("Repair failed, key will be removed.") : Return False
            Return True

        Next

        ' Return false if no valid command is found
        gLog("Repair failed, key will be removed.")
        Return False

    End Function

    ''' <summary>
    ''' Runs the checks every key gets before bucket processing. A key with no <c> = </c> or no
    ''' value goes through <see cref="fixMissingEquals"/>, whatever the scan settings say. Doubled
    ''' backslashes are reported and collapsed under Slashes. Leading and trailing whitespace is
    ''' reported and trimmed from the name and value whatever the settings say. Last,
    ''' <see cref="chkCasing"/> checks the key type.
    ''' </summary>
    '''
    ''' <param name="result">The lint result that collects the errors</param>
    '''
    ''' <param name="key">
    ''' An <c> iniKey </c> whose basic syntactic validity will be assessed
    ''' </param>
    '''
    ''' <returns>
    ''' <c> False </c> if <paramref name="key"/> is <c> Nothing </c> or couldn't be given a
    ''' value, which tells the caller to remove it, <c> True </c> otherwise
    ''' </returns>
    Private Function cValidity(result As EntryLintResult, key As iniKey) As Boolean

        If key Is Nothing Then argIsNull(NameOf(key)) : Return False

        ' Attempt to fix the case where keys are missing an equal sign to delineate name and value
        If key.typeIs("DeleteMe") Then

            gLog($"Broken Key Found: {key.Name}")

            ' If we didn't find a fixable situation, delete the key
            Dim fixedMsngEq = fixMissingEquals(result, key, ValidCmds)

            fullKeyErr(result, key, "Missing '=' detected and repaired in key.", fixedMsngEq)

            If Not fixedMsngEq Then

                result.RecordError($"{key.Name} is missing a '=' or was not provided with a value. It will be deleted.", Array.Empty(Of String)())
                Return False

            End If

        End If

        ' Remove any instances of double backslashes because we don't expect them

        If key.vHas("\\") Then

            fullKeyErr(result, key, "Extraneous backslashes (\\) detected", lintSlashes.ShouldScan)

            While (key.Value.Contains("\\") And lintSlashes.fixFormat) : key.Value = key.Value.Replace("\\", "\") : End While

        End If

        ' Check for leading or trailing whitespace, do this always as spaces in the name interfere with proper keyType identification
        If key.Name.StartsWith(" ", StringComparison.InvariantCulture) OrElse key.Name.EndsWith(" ", StringComparison.InvariantCulture) OrElse
            key.Value.StartsWith(" ", StringComparison.InvariantCulture) OrElse key.Value.EndsWith(" ", StringComparison.InvariantCulture) Then

            fullKeyErr(result, key, "Detected unwanted whitespace in iniKey", True)
            fixStr(True, key.Value, key.Value.Trim)
            fixStr(True, key.Name, key.Name.Trim)

        End If

        ' Make sure the keyType is valid
        chkCasing(result, key, ValidCmdsLookup, ValidCmdsJoined, key.KeyType)

        Return True

    End Function

    ''' <summary>
    ''' Checks text from the <c> Value </c> or the <c> KeyType </c> of an <c> iniKey </c> against a
    ''' lookup of properly cased values. A casing error is reported under Casing and repaired by
    ''' replacing every occurrence of <paramref name="strToChk"/> in the value, or the key type in
    ''' the name. Text missing from the lookup is reported under Invalid Values.
    ''' </summary>
    '''
    ''' <param name="result">
    ''' The <c> EntryLintResult </c> collecting errors for the entry being linted
    ''' </param>
    '''
    ''' <param name="key">
    ''' The <c> iniKey </c> whose casing will be audited
    ''' </param>
    '''
    ''' <param name="casedLookup">
    ''' Maps each valid value, case-insensitively, to its properly cased form
    ''' </param>
    '''
    ''' <param name="casedJoined">
    ''' The valid values joined into one string for the invalid data error
    ''' </param>
    '''
    ''' <param name="strToChk">
    ''' The text being audited, taken from either the key's value or its type
    ''' </param>
    Private Sub chkCasing(result As EntryLintResult,
                          key As iniKey,
                          casedLookup As Dictionary(Of String, String),
                          casedJoined As String,
                          strToChk As String)

        ' Look up the properly cased string. If it isn't in there, the value is invalid
        Dim casedString As String = Nothing
        Dim found = casedLookup.TryGetValue(strToChk, casedString)
        If Not found Then casedString = strToChk

        Dim hasCasingErr = found AndAlso Not casedString.Equals(strToChk, StringComparison.InvariantCulture)

        ' Inform the user if there are casing errors and fix them
        fullKeyErr(result, key, $"{casedString} has a casing error.", hasCasingErr And lintCasing.ShouldScan, False, "", "")
        fixStr(hasCasingErr AndAlso lintCasing.fixFormat AndAlso key.Value.Contains(strToChk), key.Value, Function() key.Value.Replace(strToChk, casedString))
        fixStr(hasCasingErr AndAlso lintCasing.fixFormat AndAlso key.Name.Contains(strToChk), key.Name, Function() key.Name.Replace(key.KeyType, casedString))

        ' Inform the user about invalid data
        If Not found AndAlso lintInvalid.ShouldScan Then
            fullKeyErr(result, key, $"Invalid data provided: {strToChk} in {key.ToString()}{Environment.NewLine}Valid data: {casedJoined}", True)
        End If

    End Sub

    ''' <summary>
    ''' Processes a FileKey format winapp2.ini <c> iniKey </c> and checks it for errors, correcting
    ''' them where possible. Covers flag casing and spelling, missing pipes, colons used as
    ''' semicolons (skipped when the path's first segment holds a drive letter), stray
    ''' semicolons and backslashes before a pipe, and the path itself. Under Parameters, a
    ''' pattern that repeats an earlier one (ignoring case), including a second empty one, is
    ''' reported, and so is a pattern list out of order. When either turns up, the repair drops
    ''' repeated and empty patterns and sorts the rest.
    ''' </summary>
    '''
    ''' <param name="result">The lint result that collects the errors</param>
    '''
    ''' <param name="key">
    ''' A winapp2.ini FileKey format <c> iniKey </c> to be checked for correctness
    ''' </param>
    '''
    ''' <returns><paramref name="key"/> itself, edited in place</returns>
    Public Function pFileKey(result As EntryLintResult, key As iniKey) As iniKey

        If key Is Nothing Then argIsNull(NameOf(key)) : Return key

        ' Pipe symbol checks
        Dim iteratorCheckerList = key.PipeSplit

        If iteratorCheckerList.Length > 2 Then

            Dim rawFlag = iteratorCheckerList(iteratorCheckerList.Length - 1)
            Dim upperFlag = rawFlag.ToUpperInvariant()
            If rawFlag <> upperFlag AndAlso (upperFlag = "RECURSE" OrElse upperFlag = "REMOVESELF") Then
                fullKeyErr(result, key, $"{upperFlag} has a casing error.", lintCasing.ShouldScan, lintCasing.fixFormat,
                           key.Value, Function() key.Value.Substring(0, key.Value.Length - rawFlag.Length) & upperFlag)
            End If

            iteratorCheckerList = key.PipeSplit

        End If

        fullKeyErr(result, key, "Missing pipe (|) in FileKey.", Not key.vHas("|"))

        ' The driveLtr check to allow entries that contain hard coded drive letters to contain colons. Since this is an edge case only likely to pop up in winapp3.ini (as far as official releases go)
        ' We'll assume that if the path contains a hard coded drive letter, any colon use is intentional and disable this check.
        fullKeyErr(result, key, "Colon (:) found where there should be a semicolon (;)", key.Value.Contains(":") And Not driveLtrs.IsMatch(getFirstDir(key.Value)), lintSemis.fixFormat, key.Value, Function() key.Value.Replace(":", ";"))

        ' Trailing semicolon check only applies to the parameters section (after the first pipe)
        Dim firstPipe = key.Value.IndexOf("|"c)
        Dim afterFirstPipe = If(firstPipe >= 0, key.Value.Substring(firstPipe), "")
        fullKeyErr(result, key, "Trailing semicolon (;) in parameters", lintSemis.ShouldScan And afterFirstPipe.Contains(";|"), lintSemis.fixFormat, key.Value, Function() key.Value.Replace(";|", "|"))

        ' Check for incorrect spellings of RECURSE or REMOVESELF
        If iteratorCheckerList.Length > 2 Then fullKeyErr(result, key, "RECURSE or REMOVESELF is incorrectly spelled, or there are too many pipe (|) symbols.", Not iteratorCheckerList(2).Equals("RECURSE", StringComparison.OrdinalIgnoreCase) And Not iteratorCheckerList(2).Equals("REMOVESELF", StringComparison.OrdinalIgnoreCase))

        ' Check for missing pipe symbol on recurse and removeself, fix them if detected
        fullKeyErr(result, key, "Missing pipe (|) before RECURSE.", lintFlags.ShouldScan And key.vHas("RECURSE") And Not key.vHas("|RECURSE"), lintFlags.fixFormat, key.Value, Function() key.Value.Replace("RECURSE", "|RECURSE"))
        fullKeyErr(result, key, "Missing pipe (|) before REMOVESELF.", lintFlags.ShouldScan And key.vHas("REMOVESELF") And Not key.vHas("|REMOVESELF"), lintFlags.fixFormat, key.Value, Function() key.Value.Replace("REMOVESELF", "|REMOVESELF"))

        ' Backslash checks, fix if detected
        fullKeyErr(result, key, "Backslash (\) found before pipe (|).", lintSlashes.ShouldScan And key.vHas("\|"), lintSlashes.fixFormat, key.Value, Function() key.Value.Replace("\|", "|"))

        ' Check for duplicate or empty parameters using fileKeyParams
        Dim keyParams As New fileKeyParams(key.Value)
        Dim seenArgs As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Dim dupedArgs As New List(Of String)

        For Each arg In keyParams.Patterns

            If seenArgs.Add(arg) Then Continue For
            result.RecordError($"{If(arg.Length = 0, "Empty", "Duplicate")} FileKey parameter found", {$"Key:     {key.ToString()}", $"Parameter: {arg}"}, lintParams.ShouldScan)
            dupedArgs.Add(arg)

        Next

        Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Dim deduped = keyParams.Patterns.Where(Function(p) p.Length > 0 AndAlso seen.Add(p)).ToList()
        Dim sortedPatterns = deduped _
            .OrderBy(Function(p) p, StringComparer.OrdinalIgnoreCase) _
            .ThenBy(Function(p) p, StringComparer.Ordinal) _
            .ToList()

        Dim outOfOrder = Not deduped.SequenceEqual(sortedPatterns, StringComparer.Ordinal)
        If outOfOrder Then
            result.RecordError("FileKey parameters are not in alphabetical order",
                               {$"Key:      {key.ToString()}",
                                $"Current:  {String.Join(";", deduped)}",
                                $"Sorted:   {String.Join(";", sortedPatterns)}"}, lintParams.ShouldScan)
        End If

        If lintParams.fixFormat AndAlso (dupedArgs.Count > 0 OrElse outOfOrder) Then key.Value = keyParams.Reconstruct(sortedPatterns)

        ' Make sure that FileKey paths point to a filesystem location
        chkPathFormatValidity(result, key, False)

        Return key

    End Function

    ''' <summary>
    ''' Processes a DetectFile format <c> iniKey </c> and checks it for errors, correcting where
    ''' possible. A trailing backslash is reported and removed under Slashes. A wildcard is
    ''' reported only for the System Ninja flavor, and a wildcard before the last path segment
    ''' is always reported. Both wildcard reports ignore the scan settings.
    ''' </summary>
    '''
    ''' <param name="result">The lint result that collects the errors</param>
    '''
    ''' <param name="key">
    ''' A winapp2.ini DetectFile format <c> iniKey </c> to be checked for correctness
    ''' </param>
    Private Function pDetectFile(result As EntryLintResult, key As iniKey) As iniKey

        ' Trailing Backslashes & nested wildcards
        fullKeyErr(result, key, "Trailing backslash (\) found in DetectFile",
                   lintSlashes.ShouldScan AndAlso key.Value.Length > 0 AndAlso key.Value(key.Value.Length - 1) = "\"c, lintSlashes.fixFormat, key.Value, Function() key.Value.TrimEnd("\"c))

        If key.vHas("*") Then

            ' System Ninja doesn't support wildcards in DetectFile keys, so we won't when the Flavor is set to that
            fullKeyErr(result, key, "Wildcard (*) found in DetectFile", CurrentWinappFlavor = WinappFlavor.SystemNinja)

            Dim splitDir = key.Value.Split(CChar("\"))

            Dim hasNestedWildcard = splitDir.Take(splitDir.Length - 1).Any(Function(dir) dir.Contains("*"))
            fullKeyErr(result, key, "Nested wildcard found in DetectFile", hasNestedWildcard)

        End If

        ' Make sure that DetectFile paths point to a filesystem location
        chkPathFormatValidity(result, key, False)

        Return key

    End Function

    ''' <summary>
    ''' Audits the root of a file system or registry path, ignoring the key's patterns and an
    ''' ExcludeKey's flag. A registry hive in the wrong case is reported and fixed under Casing,
    ''' and an unknown hive is reported under Path Validity. A file system path must start with a <c> % </c> or
    ''' contain a drive letter in its first segment, and must not contain <c> &lt; </c>,
    ''' <c> &gt; </c> or <c> " </c>; both are report-only Path Validity checks. Does nothing
    ''' while the Path Validity and Casing scans are both off.
    ''' </summary>
    '''
    ''' <param name="result">The lint result that collects the errors</param>
    '''
    ''' <param name="key">
    ''' An <c> iniKey </c> containing a registry or filesystem path to have its syntax validated
    ''' </param>
    '''
    ''' <param name="isRegistry">
    ''' Indicates whether <paramref name="key"/> is expected to hold a registry path
    ''' </param>
    Private Sub chkPathFormatValidity(result As EntryLintResult, key As iniKey, isRegistry As Boolean)

        If Not (lintPathValidity.ShouldScan OrElse lintCasing.ShouldScan) Then Return

        ' Strip the pattern suffix (everything after |) and ExcludeKey flags before inspecting the path
        Dim pathPortion = If(Not key.typeIs("ExcludeKey"), key.Value.Split(CChar("|"))(0), New excludeKeyParams(key.Value).Path)
        Dim rootStr = getFirstDir(pathPortion)

        ' Ensure that registry paths have a valid hive and file paths have either a variable or a drive letter
        If isRegistry AndAlso Not (shortReg.IsMatch(rootStr) OrElse longReg.IsMatch(rootStr)) Then

            Dim validHives = {"HKCU", "HKLM", "HKCR", "HKU", "HKCC",
                              "HKEY_CURRENT_USER", "HKEY_LOCAL_MACHINE", "HKEY_CLASSES_ROOT", "HKEY_USERS", "HKEY_CURRENT_CONFIG"}
            Dim correctHive = validHives.FirstOrDefault(Function(h) h.Equals(rootStr, StringComparison.OrdinalIgnoreCase))

            If correctHive IsNot Nothing Then
                fullKeyErr(result, key, "Incorrect registry root casing.", lintCasing.ShouldScan, lintCasing.fixFormat, key.Value,
                           Function()
                               Dim idx = key.Value.IndexOf(rootStr)
                               Return key.Value.Remove(idx, rootStr.Length).Insert(idx, correctHive)
                           End Function)
            Else
                fullKeyErr(result, key, "Invalid registry path detected.", lintPathValidity.ShouldScan)
            End If

        End If

        fullKeyErr(result, key, "Invalid file system path detected.",
                   Not isRegistry AndAlso lintPathValidity.ShouldScan AndAlso Not (rootStr.StartsWith("%", StringComparison.InvariantCultureIgnoreCase) OrElse driveLtrs.IsMatch(rootStr)))

        fullKeyErr(result, key, "Illegal characters (< > "") detected in filesystem path.",
                   Not isRegistry AndAlso lintPathValidity.ShouldScan AndAlso pathPortion.IndexOfAny({CChar("<"), CChar(">"), CChar("""")}) >= 0)

    End Sub

    ''' <summary>
    ''' Checks one ExcludeKey format <c> iniKey </c> for errors, correcting where possible: its
    ''' flag (<see cref="checkExcludeFlags"/>) and flag count under Flags, then its path. A FILE or
    ''' PATH key also needs a backslash before its pattern pipe (Path Validity), and a REG key is
    ''' reported for the BleachBit flavor. Everything after the flag checks, the BleachBit report
    ''' included, only runs while the Path Validity or Casing scan is on and the key has a valid
    ''' flag.
    ''' </summary>
    '''
    ''' <param name="result">The lint result that collects the errors</param>
    '''
    ''' <param name="key">
    ''' A winapp2.ini ExcludeKey format <c> iniKey </c> to be checked for correctness
    ''' </param>
    '''
    ''' <param name="hasF">
    ''' Set to <c> True </c> if the key excludes a file system location. Only set while the path
    ''' checks run
    ''' </param>
    '''
    ''' <param name="hasR">
    ''' Set to <c> True </c> if the key excludes a registry location. Only set while the path
    ''' checks run
    ''' </param>
    Private Sub pExcludeKey(result As EntryLintResult,
                            key As iniKey,
                            ByRef hasF As Boolean,
                            ByRef hasR As Boolean)

        Dim hasValidFlags = key.vHasAny({"FILE|", "PATH|", "REG|"})
        If Not hasValidFlags Then hasValidFlags = checkExcludeFlags(result, key)

        Dim pipeParts = key.PipeSplit
        fullKeyErr(result, key, "ExcludeKey has too many flags", lintFlags.ShouldScan AndAlso If(key.vHas("REG|"), pipeParts.Length > 2, pipeParts.Length > 3))


        If Not ((lintPathValidity.ShouldScan OrElse lintCasing.ShouldScan) AndAlso hasValidFlags) Then Return

        Select Case True

            Case key.vHasAny({"FILE|", "PATH|"})

                hasF = True

                chkPathFormatValidity(result, key, False)

                Dim flagPipe = key.Value.IndexOf("|"c)
                Dim patternPipe = If(flagPipe >= 0, key.Value.IndexOf("|"c, flagPipe + 1), -1)
                fullKeyErr(result, key, "Missing backslash (\) before pipe (|) in ExcludeKey.",
                           lintPathValidity.ShouldScan AndAlso Not key.vHas("\|"),
                           lintPathValidity.fixFormat AndAlso patternPipe >= 0,
                           key.Value, Function() If(patternPipe >= 0, key.Value.Insert(patternPipe, "\"), ""))

            Case key.vHas("REG|")

                hasR = True

                chkPathFormatValidity(result, key, True)
                fullKeyErr(result, key, "ExcludeKey contains REG flag in BleachBit flavor", CurrentWinappFlavor = WinappFlavor.BleachBit)

            Case Else

                checkExcludeFlags(result, key)

        End Select


    End Sub

    ''' <summary>
    ''' Checks an ExcludeKey whose value has no <c> FILE| </c>, <c> PATH| </c> or <c> REG| </c>.
    ''' While the Flags scan is on, a value that doesn't start with one of the flags is reported,
    ''' and one that does is reported as missing the pipe after it. The pipe goes in when the
    ''' Flags rule's <c> ShouldRepair </c> is on, checked directly rather than through
    ''' <see cref="lintRule.fixFormat"/>.
    ''' </summary>
    '''
    ''' <param name="result">The lint result that collects the errors</param>
    '''
    ''' <param name="key">
    ''' A winapp2.ini ExcludeKey format <c> iniKey </c> to be checked for correctness
    ''' </param>
    '''
    ''' <returns>Whether the value starts with <c> FILE </c>, <c> PATH </c> or <c> REG </c></returns>
    Private Function checkExcludeFlags(result As EntryLintResult, key As iniKey) As Boolean

        Dim matches = HasFlagRegex.Matches(key.Value)

        ' If we're not checking flags, we should at least indicate whether or not valid ones are present
        If Not lintFlags.ShouldScan Then Return matches.Count > 0

        If matches.Count = 0 Then

            fullKeyErr(result, key, "No valid exclude flag (FILE, PATH, or REG) found in ExcludeKey.")
            Return False

        End If

        Dim foundFlag = matches(0)
        Dim fixedValue = key.Value.Insert(foundFlag.Length, "|")
        fullKeyErr(result, key, "Missing pipe (|) after ExcludeKey flag", repCond:=lintFlags.ShouldRepair, repairVal:=key.Value, newVal:=fixedValue)

        Return True

    End Function

    ''' <summary>
    ''' Reports keys in a typed bucket that are out of alphabetical order, comparing values with
    ''' each pipe swapped for <c> " \ \" </c> and numbers padded, and rewrites the bucket's values
    ''' in sorted order, renamed <c> &lt;keyType&gt;1 </c> onward. The rewrite runs whenever a
    ''' key is out of place or keys were removed, whatever the Alphabetization and numbering
    ''' repairs say.
    ''' <br /><br />
    '''
    ''' Does nothing while the Alphabetization scan is off or the bucket holds one key, so
    ''' then nothing renumbers the bucket, even after a removal.
    ''' </summary>
    '''
    ''' <param name="result">The lint result that collects the errors</param>
    '''
    ''' <param name="entry">
    ''' The <c> winapp2entry </c> whose bucket will be sorted
    ''' </param>
    '''
    ''' <param name="keyType">
    ''' The key type name identifying which bucket to sort
    ''' </param>
    '''
    ''' <param name="hadDuplicatesRemoved">
    ''' Indicates whether keys have just been removed from the bucket, which forces the rewrite
    ''' so the remaining keys get renumbered
    ''' </param>
    Private Sub sortKeys(result As EntryLintResult,
                          entry As winapp2entry,
                          keyType As String,
                          hadDuplicatesRemoved As Boolean)

        Dim bucketIdx = winapp2entry.GetBucketIndex(keyType)
        If bucketIdx < 0 Then Return

        Dim bucket = entry.KeyLists(bucketIdx)

        If bucket.Count <= 1 OrElse Not lintAlpha.ShouldScan Then Return

        Dim keyValues As New List(Of String)

        For i = 0 To bucket.Count - 1
            keyValues.Add(bucket(i).Value)
        Next

        Dim sortedKeyValues = replaceAndSort(keyValues, "|", " \ \")

        ' If the bucket is already sorted and nothing got removed, there's nothing to report and nothing to rewrite
        If Not hadDuplicatesRemoved Then
            Dim alreadySorted = True
            For i = 0 To keyValues.Count - 1
                If Not keyValues(i).Equals(sortedKeyValues(i), StringComparison.Ordinal) Then
                    alreadySorted = False
                    Exit For
                End If
            Next
            If alreadySorted Then Return
        End If

        Dim findings = findOutOfPlace(keyValues, sortedKeyValues)

        For Each m In findings

            Dim contextStr = $"{keyType}{m.ActualPos + 1}={m.Item}"
            result.RecordError(
                $"{keyType} alphabetization",
                {$"{contextStr} appears to be out of place",
                 $"Expected position: {m.ExpectedPos + 1}"})

        Next

        If Not (findings.Count > 0 OrElse hadDuplicatesRemoved) AndAlso
               (lintAlpha.fixFormat OrElse lintWrongNums.fixFormat OrElse lintExtraNums.fixFormat) Then Return

        For i = 0 To bucket.Count - 1
            bucket(i).Value = sortedKeyValues(i)
            bucket(i).Name = keyType & CStr(i + 1)
        Next

    End Sub

    ''' <summary>
    ''' Records an error when data is received that does not match an expected value
    ''' </summary>
    '''
    ''' <param name="result">The lint result that collects the error</param>
    '''
    ''' <param name="err">
    ''' A description of the error as it will be displayed to the user
    ''' </param>
    '''
    ''' <param name="received">
    ''' The (erroneous) input data
    ''' </param>
    '''
    ''' <param name="expected">
    ''' The expected data
    ''' </param>
    '''
    ''' <param name="cond">
    ''' Indicates whether the error condition is present <br /><br />
    ''' Optional, Default: <c> True </c>
    ''' </param>
    Private Sub inputMismatchErr(result As EntryLintResult,
                                 err As String,
                                 received As String,
                                 expected As String,
                                 Optional cond As Boolean = True)

        result.RecordError(err, {$"Expected: {expected}", $"Found:    {received}"}, cond)

    End Sub

    ''' <summary>
    ''' Records an error that names <paramref name="key"/> when <paramref name="cond"/> is
    ''' <c> True </c>, and when <paramref name="repCond"/> is too, sets
    ''' <paramref name="repairVal"/> to <paramref name="newVal"/>. Nothing is repaired while
    ''' <paramref name="cond"/> is <c> False </c>.
    ''' </summary>
    '''
    ''' <param name="result">The lint result that collects the error</param>
    '''
    ''' <param name="key">
    ''' The <c> iniKey </c> containing an error
    ''' </param>
    '''
    ''' <param name="err">
    ''' A description of the error as it will be displayed to the user
    ''' </param>
    '''
    ''' <param name="cond">
    ''' Indicates whether the error is present, with any scan setting already folded in <br /><br />
    ''' Optional, Default: <c> True </c>
    ''' </param>
    '''
    ''' <param name="repCond">
    ''' Indicates whether to apply the repair <br /><br />
    ''' Optional, Default: <c> False </c>
    ''' </param>
    '''
    ''' <param name="repairVal">
    ''' The value to repair. Callers pass a property such as <c> key.Value </c>, which VB writes
    ''' back when the call returns <br /><br />
    ''' Optional, Default: <c> "" </c>
    ''' </param>
    '''
    ''' <param name="newVal">
    ''' The corrected value for <paramref name="repairVal"/> <br /><br />
    ''' Optional, Default: <c> "" </c>
    ''' </param>
    Private Sub fullKeyErr(result As EntryLintResult,
                           key As iniKey,
                           err As String,
                  Optional cond As Boolean = True,
                  Optional repCond As Boolean = False,
            Optional ByRef repairVal As String = "",
                  Optional newVal As String = "")

        If Not cond Then Return

        result.RecordError(err, {$"Key: {key.ToString()}"})
        fixStr(cond And repCond, repairVal, newVal)

    End Sub

    ''' <summary>
    ''' Lazy version of <c> fullKeyErr </c>. <paramref name="newValFactory"/> only runs when
    ''' both <paramref name="cond"/> and <paramref name="repCond"/> are true. Worth using
    ''' where building the replacement costs something, like <c> key.Value.Replace(...) </c>,
    ''' which would otherwise allocate for every key whether or not the repair fires.
    ''' </summary>
    Private Sub fullKeyErr(result As EntryLintResult,
                           key As iniKey,
                           err As String,
                           cond As Boolean,
                           repCond As Boolean,
                     ByRef repairVal As String,
                           newValFactory As Func(Of String))

        If Not cond Then Return

        result.RecordError(err, {$"Key: {key.ToString()}"})
        fixStr(cond And repCond, repairVal, newValFactory)

    End Sub

    ''' <summary>
    ''' Replaces <paramref name="currentValue"/> with <paramref name="newValue"/> and logs the
    ''' change, when <paramref name="param"/> is <c> True </c>
    ''' </summary>
    '''
    ''' <param name="param">
    ''' Indicates whether to replace the string
    ''' </param>
    '''
    ''' <param name="currentValue">
    ''' The string to be replaced, passed by reference
    ''' </param>
    '''
    ''' <param name="newValue">
    ''' The replacement value for <paramref name="currentValue"/>
    ''' </param>
    Private Sub fixStr(param As Boolean,
                 ByRef currentValue As String,
                       newValue As String)

        If Not param Then Return

        gLog($"Changing '{currentValue}' to '{newValue}'")
        currentValue = newValue

    End Sub

    ''' <summary>
    ''' Lazy version of <c> fixStr </c>. <paramref name="newValueFactory"/> only runs when
    ''' <paramref name="param"/> is true, so a gated call site doesn't allocate a replacement
    ''' string it's never going to use.
    ''' </summary>
    Private Sub fixStr(param As Boolean,
                 ByRef currentValue As String,
                       newValueFactory As Func(Of String))

        If Not param Then Return

        Dim resolved = newValueFactory()
        gLog($"Changing '{currentValue}' to '{resolved}'")
        currentValue = resolved

    End Sub

End Module
