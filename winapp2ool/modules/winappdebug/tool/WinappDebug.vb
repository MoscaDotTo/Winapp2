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

Imports System.Globalization
Imports System.Text
Imports System.Text.RegularExpressions

''' <summary>
''' Observes, reports, and attempts to repair errors in winapp2.ini
''' </summary>
Public Module WinappDebug

    ''' <summary>
    ''' The number of errors found during the lint
    ''' </summary>
    Public Property ErrorsFound As Integer = 0

    ''' <summary>
    ''' The winapp2ool logslice from the most recent Lint run
    ''' </summary>
    Public Property MostRecentLintLog As New System.Text.StringBuilder

    ''' <summary>
    ''' The current rules for scans and repairs
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
    ''' Controls scan/repairs for missing or True Default values
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
    ''' Controls scan/repairs for keys with invlaid values
    ''' <br /> Default: <c> True </c>
    ''' </summary>
    Private Property lintInvalid As lintRule = Rules(10)

    ''' <summary>
    ''' Controls scan/repairs for winapp2.ini syntax errors
    ''' <br /> Default: <c> True </c>
    ''' </summary>
    Private Property lintSyntax As lintRule = Rules(11)

    ''' <summary>
    ''' Controls scan/repairs for invalid file or regsitry paths
    ''' <br /> Default: <c> True </c>
    ''' </summary>
    Private Property lintPathValidity As lintRule = Rules(12)

    ''' <summary>
    ''' Controls scan/repairs for improper use of semicolons
    ''' <br /> Default: <c> True </c>
    ''' </summary>
    Private Property lintSemis As lintRule = Rules(13)

    ''' <summary>
    ''' Controls scan/repairs for keys that can be merged into eachother (FileKeys only currently)
    ''' <br /> Default: <c> False </c>
    ''' </summary>
    Public Property lintOpti As lintRule = Rules(14)

    ''' <summary>
    ''' Regex to detect long form registry paths
    ''' </summary>
    Private Property longReg As New Regex("HKEY_(C(URRENT_(USER$|CONFIG$)|LASSES_ROOT$)|LOCAL_MACHINE$|USERS$)")

    ''' <summary>
    ''' Regex to detect short form registry paths
    ''' </summary>
    Private Property shortReg As New Regex("HK(C(C$|R$|U$)|LM$|U$)")

    ''' <summary>
    ''' Regex to detect valid LangSecRef numbers
    ''' </summary>
    Private Property secRefNums As New Regex("30(0([1-6])|2([1-9])|3([0-9])|4([0-4]))")

    ''' <summary>
    ''' Regex to detect valid drive letter parameters
    ''' </summary>
    Private Property driveLtrs As New Regex("[a-zA-Z]:")

    ''' <summary>
    ''' Regex to detect potential %EnvironmentVariables%
    ''' </summary>
    Private Property envVarRegex As New Regex("%[A-Za-z0-9]*%")

    ''' <summary>
    ''' Regex to detect ExcludeKey flags
    ''' </summary>
    Private Property HasFlagRegex As New Regex("^(FILE|PATH|REG)")

    ''' <summary>
    ''' commandline runtime parameter for creating winapp2.ini with a version string
    ''' reflecting the current date <br />
    ''' Default: <c> True </c> - uses current date; <c> False </c> uses static version string
    ''' </summary>
    Private Property UseCurrentDate As Boolean = False

    ''' <summary>
    ''' Handles the commandline args for <c> WinappDebug </c>
    ''' </summary>
    '''
    ''' <remarks>
    ''' Supported args: <br />
    ''' <c> -1f </c> / <c> -1d </c> — input winapp2.ini (slot 1) <br />
    ''' <c> -3f </c> / <c> -3d </c> — save target (slot 3) <br />
    ''' <c> -c </c> — enable saving of changes made by the linter <br />
    ''' <c> -usedate </c> — use current date in version string
    ''' </remarks>
    Public Sub HandleLintCmdLine()

        InitDefaultLintSettings()

        Dim spec As New CliArgSpec(NameOf(WinappDebug))
        spec.WithFile(1, winappDebugFile1) _
            .WithFile(3, winappDebugFile3) _
            .WithFlag("-c", Sub() SaveChanges = Not SaveChanges) _
            .WithFlag("-usedate", Sub() UseCurrentDate = Not UseCurrentDate) _
            .Parse()

        If Not cmdargs.Contains("UNIT_TESTING_HALT") Then InitDebug()

    End Sub

    ''' <summary>
    ''' Lints <paramref name="givenIni"/>'s winapp2.ini formatting from outside the module's UI.
    ''' Returns the linted <c>iniFile2</c>.
    ''' </summary>
    '''
    ''' <param name="givenIni">
    ''' The winapp2.ini syntax <c>iniFile2</c> to be linted
    ''' </param>
    '''
    ''' <param name="forceOpti">
    ''' Indicates whether or not the linter should attempt to optimize entries <br />
    ''' Optional, Default: <c> False </c>
    ''' </param>
    Public Function remotedebug(givenIni As iniFile2,
                                Optional forceOpti As Boolean = False) As iniFile2

        If givenIni Is Nothing Then argIsNull(NameOf(givenIni)) : Return Nothing

        If forceOpti Then
            lintOpti.ShouldScan = True
            lintOpti.ShouldRepair = True
        End If

        Dim wa2 As New winapp2file2(givenIni)
        Debug(wa2)
        Return wa2.ToIni()

    End Function

    ''' <summary>
    ''' Validates winapp2.ini, then sets up the output window before sending it off to the linter.
    ''' After linting, reports the results of the lint to the user
    ''' </summary>
    Public Sub InitDebug()

        Dim inputFile = winappDebugFile1.Load()

        If Not enforceFileHasContent(inputFile) Then Return

        Dim wa2 As New winapp2file2(inputFile, UseCurrentDate)

        clrConsole()

        print(3, "Beginning analysis of winapp2.ini", trailr:=True)
        gLog("Beginning lint", leadr:=True, ascend:=True)

        MostRecentLintLog.Clear()

        Debug(wa2)

        gLog(descend:=True)
        gLog("Lint complete")
        setNextMenuHeaderText("Lint complete", printColor:=ConsoleColor.Green)
        print(4, "Completed analysis of winapp2.ini", conjoin:=True)
        print(0, $"{ErrorsFound} possible errors were detected.")
        print(0, $"Number of entries {wa2.Count}", trailingBlank:=True)

        RewriteChanges(wa2)

        print(0, anyKeyStr, closeMenu:=True)
        crk()

    End Sub

    ''' <summary>
    ''' Sends the entries in a winapp2.ini format <c>iniFile2</c> into specific format and syntax checking routines
    ''' </summary>
    '''
    ''' <param name="fileToBeDebugged">
    ''' A <c> winapp2file2 </c> to be linted
    ''' </param>
    Public Sub Debug(ByRef fileToBeDebugged As winapp2file2)

        If fileToBeDebugged Is Nothing Then argIsNull(NameOf(fileToBeDebugged)) : Return

        ErrorsFound = 0
        MostRecentLintLog.Clear()
        gLog(ascend:=True)

        Dim duplicateNames = FindDuplicateEntryNames(fileToBeDebugged)

        For Each entry In fileToBeDebugged.Entries
            EmitEntryResult(ProcessEntry(entry, duplicateNames))
        Next

        resetKeyTrackers()
        AlphabetizeEntries(fileToBeDebugged)

    End Sub

    ''' <summary>
    ''' Returns the set of entry names that appear more than once in a <c>winapp2file2</c>
    ''' </summary>
    '''
    ''' <param name="winapp">
    ''' The <c>winapp2file2</c> whose entries will be scanned for duplicate names
    ''' </param>
    Private Function FindDuplicateEntryNames(winapp As winapp2file2) As HashSet(Of String)

        Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Dim duplicates As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        For Each entry In winapp.Entries
            If Not seen.Add(entry.Name) Then duplicates.Add(entry.Name)
        Next

        Return duplicates

    End Function

    ''' <summary>
    ''' Writes the errors from an <c>EntryLintResult</c> to the console and the lint log,
    ''' and adds its error count to <c>ErrorsFound</c>
    ''' </summary>
    '''
    ''' <param name="result">
    ''' The <c>EntryLintResult</c> whose errors will be emitted
    ''' </param>
    Private Sub EmitEntryResult(result As EntryLintResult)

        For Each Errr In result.Errors

            gLog(Errr.Message, ascend:=True)

            Dim out = $"Error in {result.EntryName}: {Errr.Message}"
            cwl(out)
            MostRecentLintLog.AppendLine(out)

            For Each detail In Errr.Details
                cwl(detail)
                gLog($"  {detail}")
                MostRecentLintLog.AppendLine(detail)
            Next

            gLog(descend:=True)
            cwl()
            MostRecentLintLog.AppendLine()

        Next

        ErrorsFound += result.ErrorCount

    End Sub

    ''' <summary>
    ''' Validates the basic structure of a <c> winapp2entry2 </c> and sends off its individual keys for more specific analysis
    ''' </summary>
    '''
    ''' <param name="entry">
    ''' A <c> winapp2entry2 </c> to be audited for syntax errors
    ''' </param>
    Private Function ProcessEntry(entry As winapp2entry2, duplicateNames As HashSet(Of String)) As EntryLintResult

        Dim result As New EntryLintResult(entry.FullName)
        gLog($"Processing entry {entry.Name}", buffr:=True)

        Dim hasFileExcludes = False
        Dim hasRegExcludes = False

        result.RecordError("Duplicate entry name detected", Array.Empty(Of String)(), duplicateNames.Contains(entry.Name))
        result.RecordError("All entries must end in ' *'", Array.Empty(Of String)(), Not entry.HasValidNameSuffix)

        ValidateKeys(result, entry)

        Dim keyTypes = {"DetectOS", "LangSecRef", "Section", "SpecialDetect", "Detect", "DetectFile",
                        "Default", "Warning", "FileKey", "RegKey", "ExcludeKey"}

        For Each keyType In keyTypes

            Select Case keyType

                Case "DetectFile"

                    processKeyList(result, entry, keyType, Function(k) pDetectFile(result, k))

                Case "FileKey"

                    processKeyList(result, entry, keyType, Function(k) pFileKey(result, k))

                Case Else

                    processKeyList(result, entry, keyType, AddressOf voidDelegate, hasFileExcludes, hasRegExcludes)

            End Select

        Next

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
        result.RecordError("Entry has a Default key where there should be none", Array.Empty(Of String)(), lintDefaults.ShouldScan AndAlso hasDefaultKey AndAlso Not overrideDefaultVal)

        If lintDefaults.fixFormat AndAlso hasDefaultKey AndAlso Not overrideDefaultVal Then
            For Each k In entry.DefaultKey.ToList()
                entry.RemoveKey(k)
            Next
        End If

        If Not overrideDefaultVal Then gLog($"Finished processing {entry.Name}", buffr:=True) : Return result

        Dim expected = tsInvariant(expectedDefaultValue)

        If entry.DefaultKey.Count > 0 Then

            Dim key = entry.DefaultKey(0)
            fullKeyErr(result, key, "Incorrect value for Default Key found", lintDefaults.ShouldScan AndAlso Not key.Value = expected, lintDefaults.fixFormat, key.Value, expected)
            Return result

        End If

        result.RecordError("No Default Key found", Array.Empty(Of String)())
        entry.AddKey(New iniKey2($"Default={expected}"))

        gLog($"Finished processing {entry.Name}", buffr:=True)
        Return result

    End Function

    ''' <summary>
    ''' Checks the basic structure of all <c>iniKey2</c>s in a <c>winapp2entry2</c>,
    ''' attempts to repair some keys and place them back into their appropriate typed bucket,
    ''' and removes any that are too problematic to continue with
    ''' </summary>
    '''
    ''' <param name="entry">
    ''' A <c>winapp2entry2</c> whose <c>iniKey2</c>s will be audited for basic syntax correctness
    ''' </param>
    Private Sub ValidateKeys(result As EntryLintResult, entry As winapp2entry2)

        ' Run cValidity on all non-error buckets; remove failures
        For i = 0 To entry.KeyLists.Count - 2
            For Each key In entry.KeyLists(i).Where(Function(k) Not cValidity(result, k)).ToList()
                entry.RemoveKey(key)
            Next
        Next

        ' Run cValidity on error keys; force-remove failures
        ' (KeyType may have changed during partial repair, so normal routing would be wrong)
        For Each key In entry.ErrorKeys.Where(Function(k) Not cValidity(result, k)).ToList()
            entry.ForceRemoveErrorKey(key)
        Next

        ' Promote error keys whose KeyType is now a recognised winapp2.ini type
        entry.ReclassifyErrorKeys()

    End Sub

    ''' <summary>
    ''' Alphabetizes all the entries in a winapp2.ini file and observes any that were out of place
    ''' </summary>
    '''
    ''' <param name="winapp">
    ''' The <c>winapp2file2</c> whose entries will be alphabetized
    ''' </param>
    Private Sub AlphabetizeEntries(winapp As winapp2file2)

        For Each category In winapp.Categories

            If category.Count < 2 Then Continue For

            Dim unsortedNames As New strList
            Dim indices As New List(Of Integer)

            For i = 0 To category.Count - 1
                unsortedNames.add(category(i).Name)
                indices.Add(i)
            Next

            Dim sortedNames = replaceAndSort(unsortedNames, "-", "  ")

            If lintAlpha.ShouldScan Then findOutOfPlace(unsortedNames, sortedNames, "Entry", indices)

        Next

        If lintAlpha.fixFormat Then winapp.SortEntries()

    End Sub

    ''' <summary>
    ''' Writes any changes made during the lint back to disk, correcting any errors that were found and repaired
    ''' </summary>
    '''
    ''' <param name="winapp2file">
    ''' The <c> winapp2file2 </c> that was linted
    ''' </param>
    Private Sub RewriteChanges(winapp2file As winapp2file2)

        If Not SaveChanges Then Return

        print(0, "Saving changes, do not close winapp2ool or data loss may occur...", leadingBlank:=True)
        iniFile2.Empty(winappDebugFile3.Dir, winappDebugFile3.Name).OverwriteToFile(winapp2file.ToWinapp2String())
        print(0, "Finished saving changes.", trailingBlank:=True)

    End Sub

    ''' <summary>
    ''' Assess a list and its sorted state to observe changes in neighboring strings,
    ''' such as the changes made while sorting the strings alphabetically
    ''' </summary>
    '''
    ''' <param name="someList">
    ''' An unsorted list of strings (iniKey values or iniSection names)
    ''' </param>
    '''
    ''' <param name="sortedList">
    ''' The sorted state of <c> <paramref name="someList"/> </c>
    ''' </param>
    '''
    ''' <param name="findType">
    ''' The type of neighbor checking
    ''' <br/> <br/> When checking iniKeys (as opposed to entries),
    ''' <paramref name="findType"/> contains a <c> keyType </c>
    ''' </param>
    '''
    ''' <param name="indices">
    ''' The 0-based positions of the items in <c> <paramref name="someList"/> </c> within their containing collection
    ''' </param>
    '''
    ''' <param name="oopBool">
    ''' Tracking variable indicating that alphabetization errors have been found
    ''' <br/>  Optional, Default: <c> False </c>
    ''' </param>
    Private Sub findOutOfPlace(ByRef someList As strList,
                               ByRef sortedList As strList,
                               findType As String,
                               ByRef indices As List(Of Integer),
                               Optional ByRef oopBool As Boolean = False)

        If someList.Count < 2 Then Return

        Dim sortedIndices As New Dictionary(Of String, Integer)
        For i = 0 To sortedList.Count - 1
            sortedIndices(sortedList.Items(i)) = i
        Next

        ' Build the sequence of each item's position in the sorted list,
        ' then find which actual-list indices form the LIS. Entries outside
        ' the LIS are the minimal set that is genuinely out of place; the
        ' rest are just displaced by those entries and should not be reported.
        Dim sortedPosSequence As New List(Of Integer)
        For Each item In someList.Items
            sortedPosSequence.Add(sortedIndices(item))
        Next

        Dim lisIndices = FindLISIndices(sortedPosSequence)

        Dim misplacedEntries As New strList
        For i = 0 To someList.Count - 1
            If Not lisIndices.Contains(i) Then
                misplacedEntries.add(someList.Items(i))
            End If
        Next

        For Each entry In misplacedEntries.Items

            Dim recInd = someList.indexOf(entry)
            Dim sortInd = sortedIndices(entry)

            If recInd = sortInd Then Continue For

            Dim contextStr = If(findType = "Entry", entry, $"{findType}{recInd + 1}={entry}")
            entry = contextStr
            oopBool = True

            customErr(contextStr, $"{findType} alphabetization",
                                  {$"{entry} appears to be out of place",
                                   $"Expected position: {sortInd + 1}"})

        Next

    End Sub

    ''' <summary>
    ''' Returns the set of indices (into <paramref name="sequence"/>) that form its Longest Increasing Subsequence.
    ''' Entries whose indices are absent are the minimal set that is out of order.
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
    ''' Hands off each <c>iniKey2</c> in a winapp2.ini format typed bucket to be audited for correctness
    ''' </summary>
    '''
    ''' <param name="entry">
    ''' The <c>winapp2entry2</c> whose keys are being processed
    ''' </param>
    '''
    ''' <param name="keyType">
    ''' The <c>KeyType</c> name of the bucket to process
    ''' </param>
    '''
    ''' <param name="processKey">
    ''' The <c> function </c> that audits the keys of the <c> KeyType </c> provided <br/>
    ''' <c> voidDelegate </c> if no further operations are needed outside of the basic formatting checks
    ''' </param>
    '''
    ''' <param name="hasF">
    ''' Tracking variable indicating that there exist ExcludeKeys for file system locations
    ''' <br/> Optional, Default: <c> False </c>
    ''' </param>
    '''
    ''' <param name="hasR">
    ''' Tracking variable indicating that there exist ExcludeKeys contain registry locations
    ''' <br/> Optional, Default: <c> False </c>
    ''' </param>
    Private Sub processKeyList(result As EntryLintResult,
                               entry As winapp2entry2,
                               keyType As String,
                               processKey As Func(Of iniKey2, iniKey2),
                               Optional ByRef hasF As Boolean = False,
                               Optional ByRef hasR As Boolean = False)

        Dim bucketIdx = winapp2entry2.GetBucketIndex(keyType)
        If bucketIdx < 0 Then Return

        Dim bucket = entry.KeyLists(bucketIdx)

        If bucket.Count = 0 Then Return

        gLog($"Processing {keyType}s", ascend:=True, buffr:=True)

        Dim curNum = 1
        Dim seenValues As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        Dim dupeKeys As New List(Of iniKey2)

        For Each key In bucket.ToList()

            Select Case keyType

                Case "ExcludeKey"

                    cFormat(result, key, curNum, seenValues, dupeKeys)
                    pExcludeKey(result, key, hasF, hasR)

                Case "Detect", "DetectFile"

                    If key.typeIs("Detect") Then chkPathFormatValidity(result, key, True)
                    cFormat(result, key, curNum, seenValues, dupeKeys, bucket.Count = 1)

                Case "RegKey"

                    chkPathFormatValidity(result, key, True)
                    cFormat(result, key, curNum, seenValues, dupeKeys)

                Case "Warning", "DetectOS", "SpecialDetect", "LangSecRef", "Section", "Default"

                    If curNum > 1 AndAlso lintMulti.ShouldScan Then

                        fullKeyErr(result, key, $"Multiple {key.KeyType} detected.")
                        If lintMulti.fixFormat Then dupeKeys.Add(key)

                    End If

                    cFormat(result, key, curNum, seenValues, dupeKeys, True)

                    If key.typeIs("SpecialDetect") Then chkCasing(result, key, {"DET_CHROME", "DET_MOZILLA", "DET_THUNDERBIRD", "DET_OPERA"}, key.Value)
                    fullKeyErr(result, key, "LangSecRef holds an invalid value.", lintInvalid.ShouldScan AndAlso key.typeIs("LangSecRef") AndAlso Not secRefNums.IsMatch(key.Value))

                Case Else

                    cFormat(result, key, curNum, seenValues, dupeKeys)

            End Select

            key = processKey(key)

        Next

        For Each dupe In dupeKeys
            entry.RemoveKey(dupe)
        Next

        sortKeys2(entry, keyType, dupeKeys.Count > 0)

        If keyType = "FileKey" AndAlso lintOpti.ShouldScan Then cOptimization(entry)

        gLog(descend:=True)

    End Sub

    ''' <summary>
    ''' This function does nothing by design, used when a method or function expects to be passed a function
    ''' who modifies an iniKey2 on a KeyType where we don't want to modify the keys
    ''' </summary>
    '''
    ''' <param name="key">
    ''' An <c>iniKey2</c> with which to do nothing
    ''' </param>
    Private Function voidDelegate(key As iniKey2) As iniKey2

        Return key

    End Function

    ''' <summary>
    ''' Does some basic formatting checks that apply to all winapp2.ini format <c>iniKey2</c>s
    ''' </summary>
    '''
    ''' <param name="key">
    ''' An <c>iniKey2</c> whose format will be audited
    ''' </param>
    '''
    ''' <param name="keyNumber">
    ''' The current expected key number for numbered keys
    ''' </param>
    '''
    ''' <param name="seenValues">
    ''' A map of already-observed key values to their first-seen key string, used to detect duplicates
    ''' </param>
    '''
    ''' <param name="dupeKeys">
    ''' A tracking list of <c>iniKey2</c>s with duplicate values
    ''' </param>
    '''
    ''' <param name="noNumbers">
    ''' Indicates that the current set of keys should not be numbered
    ''' </param>
    Private Sub cFormat(result As EntryLintResult,
                        key As iniKey2,
                        ByRef keyNumber As Integer,
                        ByRef seenValues As Dictionary(Of String, String),
                        dupeKeys As List(Of iniKey2),
                        Optional noNumbers As Boolean = False)

        ' Check for duplicates
        If seenValues.ContainsKey(key.Value) Then

            result.RecordError("Duplicate key value found", {$"Key:            {key.ToString()}", $"Duplicates:     {seenValues(key.Value)}"}, lintDupes.ShouldScan)
            If lintDupes.fixFormat Then dupeKeys.Add(key)

        Else

            seenValues(key.Value) = key.ToString()

        End If

        ' Check for both types of numbering errors (incorrect and unneeded)
        Dim hasNumberingError = If(noNumbers, Not key.nameIs(key.KeyType), Not key.nameIs(key.KeyType & keyNumber))
        Dim numberingErrStr = If(noNumbers, "Detected unnecessary numbering.", $"{key.KeyType} entry is incorrectly numbered.")
        Dim fixedStr = If(noNumbers, key.KeyType, key.KeyType & keyNumber)

        gLog($"  Input mismatch error in {key.ToString()}", hasNumberingError)
        inputMismatchErr(result, numberingErrStr, key.Name, fixedStr, If(noNumbers, lintExtraNums.ShouldScan, lintWrongNums.ShouldScan) And hasNumberingError)
        fixStr(If(noNumbers, lintExtraNums.fixFormat, lintWrongNums.fixFormat) And hasNumberingError, key.Name, fixedStr)

        ' Scan for and fix any use of incorrect slashes (except in Warning keys) or trailing semicolons
        Dim fwdSlashErr = "Forward slash (/) detected in lieu of backslash (\)."
        fullKeyErr(result, key, fwdSlashErr, Not (key.typeIs("RegKey") OrElse key.typeIs("Section") OrElse key.typeIs("Warning")) AndAlso lintSlashes.ShouldScan AndAlso key.vHas("/"),
                                             lintSlashes.fixFormat, key.Value, key.Value.Replace("/", "\"))
        If key.typeIs("RegKey") AndAlso lintSlashes.ShouldScan AndAlso key.vHas("/") Then
            Dim pipeIdx = key.Value.IndexOf("|"c)
            Dim pathPart = If(pipeIdx >= 0, key.Value.Substring(0, pipeIdx), key.Value)
            If pathPart.Contains("/") Then
                fullKeyErr(result, key, fwdSlashErr, True, lintSlashes.fixFormat, key.Value,
                           pathPart.Replace("/", "\") & If(pipeIdx >= 0, key.Value.Substring(pipeIdx), ""))
            End If
        End If
        fullKeyErr(result, key, "Trailing semicolon (;).", key.ToString().Last = CChar(";") AndAlso lintSemis.ShouldScan, lintSemis.fixFormat, key.Value, key.Value.TrimEnd(CChar(";")))

        ' Do some formatting checks for environment variables if needed
        If {"FileKey", "ExcludeKey", "DetectFile"}.Contains(key.KeyType) Then cEnVar(result, key)
        keyNumber += 1

    End Sub

    ''' <summary>
    ''' Attempts to fix any broken environment variables in a given <c>iniKey2</c> <br/> <br/>
    ''' This function will attempt to repair any environment variables that are missing leading or trailing % characters
    ''' </summary>
    '''
    ''' <param name="key">
    ''' An <c>iniKey2</c> whose value will be audited for syntax errors
    ''' </param>
    '''
    ''' <param name="enVars">
    ''' The list of valid Environment Variables for Winapp2.ini
    ''' </param>
    '''
    ''' <param name="cond">
    ''' The condition under which this scan should be run
    ''' </param>
    Private Sub fixBrokenEnVars(result As EntryLintResult, key As iniKey2, enVars As String(), cond As Boolean)

        If Not cond Then Return

        Dim pipeInd = If(key.typeIs("ExcludeKey"), key.Value.IndexOf("|"c), -1)
        Dim pathPart = If(pipeInd >= 0, key.Value.Substring(pipeInd + 1), key.Value)

        For Each enVar In enVars

            If Not pathPart.Contains(enVar) Then Continue For
            If pathPart.Contains($"%{enVar}%") Then Continue For

            Dim msg = ""
            Dim fixedPath = ""
            Dim repairMade = False

            If pathPart.StartsWith($"%{enVar}\", StringComparison.InvariantCulture) Then

                msg = "Environment Variable is missing trailing %"
                fixedPath = pathPart.Replace($"%{enVar}\", $"%{enVar}%\")
                repairMade = True

            ElseIf pathPart.StartsWith($"{enVar}%\", StringComparison.InvariantCulture) Then

                msg = "Environment Variable is missing leading %"
                fixedPath = pathPart.Replace($"{enVar}%\", $"%{enVar}%\")
                repairMade = True

            ElseIf pathPart.StartsWith($"{enVar}\", StringComparison.InvariantCulture) Then

                msg = "Environment Variable is missing leading and trailing %"
                fixedPath = pathPart.Replace($"{enVar}\", $"%{enVar}%\")
                repairMade = True

            End If

            If Not repairMade Then Continue For

            Dim fixedValue = If(pipeInd >= 0, key.Value.Substring(0, pipeInd + 1) & fixedPath, fixedPath)
            fullKeyErr(result, key, msg, lintSyntax.ShouldScan, lintSyntax.ShouldRepair, key.Value, fixedValue)
            Exit For

        Next

    End Sub

    ''' <summary>
    ''' Validates the formatting of any %EnvironmentVariables% in a given <c>iniKey2</c>
    ''' </summary>
    '''
    ''' <param name="key">
    ''' The <c>iniKey2</c> whose data will be audited for environment variable correctness
    ''' </param>
    Private Sub cEnVar(result As EntryLintResult, key As iniKey2)

        ' Valid Environmental Variables for winapp2.ini
        Dim enVars = {"AllUsersProfile", "AppData", "CommonAppData", "CommonProgramFiles",
        "Documents", "HomeDrive", "LocalAppData", "LocalLowAppData", "Music", "Pictures", "ProgramData", "ProgramFiles", "Public",
        "RootDir", "SystemDrive", "SystemRoot", "Temp", "Tmp", "UserName", "UserProfile", "Video", "WinDir"}

        fullKeyErr(result, key, "Double '%' found in environment variable", key.vHas("%%"), lintSyntax.ShouldRepair, key.Value, key.Value.Replace("%%", "%"))

        fixBrokenEnVars(result, key, enVars, lintSyntax.ShouldScan AndAlso key.vHas("%") AndAlso envVarRegex.Matches(key.Value).Count = 0 OrElse key.vHasAny(enVars) AndAlso Not key.vHas("%"))

        For Each m As Match In envVarRegex.Matches(key.Value)

            Dim strippedText = m.ToString.Trim(CChar("%"))
            chkCasing(result, key, enVars, strippedText)

        Next

        ' Environment variables should be trailed by a backslash
        fullKeyErr(result, key, "Missing backslash (\) after %EnvironmentVariable%.", lintSlashes.ShouldScan And key.vHas("%") And Not key.vHasAny({"%|", "%\"}))

    End Sub

    ''' <summary>
    ''' Attempts to insert missing equal signs (=) into <c>iniKey2</c>s <br/> <br/> Returns <c> True </c> if the repair is
    '''  successful, <c> False </c> otherwise
    '''  </summary>
    '''
    ''' <param name="key">
    ''' A misformatted <c>iniKey2</c> to attempt to repair
    ''' </param>
    '''
    ''' <param name="cmds">
    ''' An array containing valid winapp2.ini <c> keyTypes </c>
    ''' </param>
    Private Function fixMissingEquals(key As iniKey2,
                                      cmds As String()) As Boolean

        gLog("Attempting missing equals repair", ascend:=True)

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

            gLog($"  Repair complete. Result: {key.ToString()}", descend:=True)

            ' Don't allow valueless keys in winapp2.ini

            If key.Value.Length = 0 Then gLog("Repair failed, key will be removed.", descend:=True) : Return False

            Return True

        Next

        ' Return false if no valid command is found
        gLog("Repair failed, key will be removed.", descend:=True)
        Return False

    End Function

    ''' <summary>
    ''' Does basic syntax and formatting audits that apply across all keys, returns <c> False </c>
    ''' if a key is malformed or if a null argument is given
    ''' </summary>
    '''
    ''' <param name="key">
    ''' An <c>iniKey2</c> whose basic syntactic validity will be assessed
    ''' </param>
    Private Function cValidity(result As EntryLintResult, key As iniKey2) As Boolean

        If key Is Nothing Then argIsNull(NameOf(key)) : Return False

        Dim validCmds = {"Default",
                         "DetectOS",
                         "DetectFile",
                         "Detect",
                         "ExcludeKey",
                         "FileKey",
                         "LangSecRef",
                         "RegKey",
                         "Section",
                         "SpecialDetect",
                         "Warning"
                        }

        ' Attempt to fix the case where keys are missing an equal sign to delineate name and value
        If key.typeIs("DeleteMe") Then

            gLog($"  Broken Key Found: {key.Name}", ascend:=True)

            ' If we didn't find a fixable situation, delete the key
            Dim fixedMsngEq = fixMissingEquals(key, validCmds)

            fullKeyErr(result, key, "Missing '=' detected and repaired in key.", fixedMsngEq)

            If Not fixedMsngEq Then
                result.RecordError($"{key.Name} is missing a '=' or was not provided with a value. It will be deleted.", Array.Empty(Of String)())
                Return False
            End If

        End If

        ' Remove any instances of double backslashes because we don't expect them

        If key.vHas("\\", True) Then

            fullKeyErr(result, key, "Extraneous backslashes (\\) detected", lintSlashes.ShouldScan)

            While (key.Value.Contains("\\") And lintSlashes.fixFormat)

                key.Value = key.Value.Replace("\\", "\")

            End While

        End If

        ' Check for leading or trailing whitespace, do this always as spaces in the name interfere with proper keyType identification
        If key.Name.StartsWith(" ", StringComparison.InvariantCulture) OrElse key.Name.EndsWith(" ", StringComparison.InvariantCulture) OrElse
            key.Value.StartsWith(" ", StringComparison.InvariantCulture) OrElse key.Value.EndsWith(" ", StringComparison.InvariantCulture) Then

            fullKeyErr(result, key, "Detected unwanted whitespace in iniKey", True)
            fixStr(True, key.Value, key.Value.Trim)
            fixStr(True, key.Name, key.Name.Trim)

        End If

        ' Make sure the keyType is valid
        chkCasing(result, key, validCmds, key.KeyType)

        Return True

    End Function

    ''' <summary>
    ''' Checks the <c> Value </c> or the <c> KeyType </c> of an <c>iniKey2</c> against a given array of expected cased values, attempts
    ''' to repair casing errors if possible
    ''' </summary>
    '''
    ''' <param name="key">
    ''' The <c>iniKey2</c> whose casing will be audited
    ''' </param>
    '''
    ''' <param name="casedArray">
    ''' The array of expected cased values
    ''' </param>
    '''
    ''' <param name="strToChk">
    ''' A pointer to the value being audited
    ''' </param>
    Private Sub chkCasing(result As EntryLintResult,
                          key As iniKey2,
                          casedArray As String(),
                          strToChk As String)

        ' Get the properly cased string
        Dim casedString As String = strToChk
        For Each casedText In casedArray

            If strToChk.Equals(casedText, StringComparison.InvariantCultureIgnoreCase) Then casedString = casedText

        Next

        ' Determine if there's a casing error
        Dim hasCasingErr = Not casedString.Equals(strToChk, StringComparison.InvariantCulture) And casedArray.Contains(casedString)
        Dim validData = String.Join(", ", casedArray)

        ' Inform the user if there are casing errors and fix them
        fullKeyErr(result, key, $"{casedString} has a casing error.", hasCasingErr And lintCasing.ShouldScan, False, "", "")
        fixStr(hasCasingErr AndAlso key.Value.Contains(strToChk), key.Value, key.Value.Replace(strToChk, casedString))
        fixStr(hasCasingErr AndAlso key.Name.Contains(strToChk), key.Name, key.Name.Replace(key.KeyType, casedString))

        ' Inform the user about invalid data
        fullKeyErr(result, key, $"Invalid data provided: {strToChk} in {key.ToString()}{Environment.NewLine}Valid data: {validData}", Not casedArray.Contains(casedString) And lintInvalid.ShouldScan)

    End Sub

    ''' <summary>
    ''' Processes a FileKey format winapp2.ini <c>iniKey2</c> and checks it for errors, correcting them where possible
    ''' </summary>
    '''
    ''' <param name="key">
    ''' A winapp2.ini FileKey format <c>iniKey2</c> to be checked for correctness
    ''' </param>
    Public Function pFileKey(result As EntryLintResult, key As iniKey2) As iniKey2

        If key Is Nothing Then argIsNull(NameOf(key)) : Return key

        ' Pipe symbol checks
        Dim iteratorCheckerList = Split(key.Value, "|")

        If iteratorCheckerList.Length > 2 Then

            Dim rawFlag = iteratorCheckerList.Last
            Dim upperFlag = rawFlag.ToUpperInvariant()
            If rawFlag <> upperFlag AndAlso (upperFlag = "RECURSE" OrElse upperFlag = "REMOVESELF") Then
                fullKeyErr(result, key, $"{upperFlag} has a casing error.", lintCasing.ShouldScan, lintCasing.fixFormat,
                           key.Value, key.Value.Substring(0, key.Value.Length - rawFlag.Length) & upperFlag)
            End If

            iteratorCheckerList = Split(key.Value, "|")

        End If

        fullKeyErr(result, key, "Missing pipe (|) in FileKey.", Not key.vHas("|"))

        ' The driveLtr check to allow entries that contain hard coded drive letters to contain colons. Since this is an edge case only likely to pop up in winapp3.ini (as far as official releases go)
        ' We'll assume that if the path contains a hard coded drive letter, any colon use is intentional and disable this check.
        fullKeyErr(result, key, "Colon (:) found where there should be a semicolon (;)", key.Value.Contains(":") And Not driveLtrs.IsMatch(getFirstDir(key.Value)), lintSemis.fixFormat, key.Value, key.Value.Replace(":", ";"))

        ' Trailing semicolon check only applies to the parameters section (after the first pipe)
        Dim firstPipe = key.Value.IndexOf("|"c)
        Dim afterFirstPipe = If(firstPipe >= 0, key.Value.Substring(firstPipe), "")
        fullKeyErr(result, key, "Trailing semicolon (;) in parameters", lintSemis.ShouldScan And afterFirstPipe.Contains(";|"), lintSemis.fixFormat, key.Value, key.Value.Replace(";|", "|"))

        ' Check for incorrect spellings of RECURSE or REMOVESELF
        If iteratorCheckerList.Length > 2 Then fullKeyErr(result, key, "RECURSE or REMOVESELF is incorrectly spelled, or there are too many pipe (|) symbols.", Not iteratorCheckerList(2).Equals("RECURSE", StringComparison.OrdinalIgnoreCase) And Not iteratorCheckerList(2).Equals("REMOVESELF", StringComparison.OrdinalIgnoreCase))

        ' Check for missing pipe symbol on recurse and removeself, fix them if detected
        Dim flags As New List(Of String) From {"RECURSE", "REMOVESELF"}
        flags.ForEach(Sub(flagStr) fullKeyErr(result, key, $"Missing pipe (|) before {flagStr}.", lintFlags.ShouldScan And key.vHas(flagStr) And Not key.vHas($"|{flagStr}"), lintFlags.fixFormat, key.Value, key.Value.Replace(flagStr, $"|{flagStr}")))

        ' Backslash checks, fix if detected
        fullKeyErr(result, key, "Backslash (\) found before pipe (|).", lintSlashes.ShouldScan And key.vHas("\|"), lintSlashes.fixFormat, key.Value, key.Value.Replace("\|", "|"))

        ' Check for duplicate or empty parameters using fileKeyParams2
        Dim keyParams As New fileKeyParams2(key.Value)
        Dim seenArgs As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Dim dupedArgs As New List(Of String)

        For Each arg In keyParams.Patterns

            If seenArgs.Add(arg) Then Continue For
            result.RecordError($"{If(arg.Length = 0, "Empty", "Duplicate")} FileKey parameter found", {$"Key:     {key.ToString()}", $"Command: {arg}"}, lintParams.ShouldScan)
            dupedArgs.Add(arg)

        Next

        If lintParams.fixFormat AndAlso dupedArgs.Count > 0 Then

            Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            Dim deduped = keyParams.Patterns.Where(Function(p) p.Length > 0 AndAlso seen.Add(p)).ToList()
            key.Value = keyParams.Path

            If deduped.Count > 0 Then key.Value &= "|" & String.Join(";", deduped)

            Select Case keyParams.Flag
                Case fileKeyFlag.Recurse : key.Value &= "|RECURSE"
                Case fileKeyFlag.RemoveSelf : key.Value &= "|REMOVESELF"
                Case fileKeyFlag.Unknown : key.Value &= "|" & keyParams.RawFlag
            End Select

        End If

        ' Make sure that FileKey paths point to a filesystem location
        chkPathFormatValidity(result, key, False)

        Return key

    End Function

    ''' <summary>
    ''' Processes a DetectFile format <c>iniKey2</c> and checks it for errors, correcting where possible
    ''' </summary>
    '''
    ''' <param name="key">
    ''' A winapp2.ini DetectFile format <c>iniKey2</c> to be checked for correctness
    ''' </param>
    Private Function pDetectFile(result As EntryLintResult, key As iniKey2) As iniKey2

        ' Trailing Backslashes & nested wildcards
        fullKeyErr(result, key, "Trailing backslash (\) found in DetectFile",
                   lintSlashes.ShouldScan AndAlso key.Value.Last = CChar("\"), lintSlashes.fixFormat, key.Value, key.Value.TrimEnd(CChar("\")))

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
    ''' Audits the syntax of file system and registry paths
    ''' </summary>
    '''
    ''' <param name="key">
    ''' An <c>iniKey2</c> containing a registry or filesystem path to have its syntax validated
    ''' </param>
    '''
    ''' <param name="isRegistry">
    ''' Indicates that the given <paramref name="key"/> is expected to hold a registry path
    ''' </param>
    Private Sub chkPathFormatValidity(result As EntryLintResult, key As iniKey2, isRegistry As Boolean)

        If Not (lintPathValidity.ShouldScan OrElse lintCasing.ShouldScan) Then Return

        ' Strip the pattern suffix (everything after |) and ExcludeKey flags before inspecting the path
        Dim pathPortion = If(Not key.typeIs("ExcludeKey"), key.Value.Split(CChar("|"))(0), New excludeKeyParams2(key.Value).Path)
        Dim rootStr = getFirstDir(pathPortion)

        ' Ensure that registry paths have a valid hive and file paths have either a variable or a drive letter
        If isRegistry AndAlso Not (shortReg.IsMatch(rootStr) OrElse longReg.IsMatch(rootStr)) Then

            Dim validHives = {"HKCU", "HKLM", "HKCR", "HKU", "HKCC",
                              "HKEY_CURRENT_USER", "HKEY_LOCAL_MACHINE", "HKEY_CLASSES_ROOT", "HKEY_USERS", "HKEY_CURRENT_CONFIG"}
            Dim correctHive = validHives.FirstOrDefault(Function(h) h.Equals(rootStr, StringComparison.OrdinalIgnoreCase))

            If correctHive IsNot Nothing Then
                Dim idx = key.Value.IndexOf(rootStr)
                Dim fixedValue = key.Value.Remove(idx, rootStr.Length).Insert(idx, correctHive)
                fullKeyErr(result, key, "Incorrect registry root casing.", lintCasing.ShouldScan, lintCasing.fixFormat, key.Value, fixedValue)
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
    ''' Processes a list of ExcludeKey format <c>iniKey2</c>s and checks them for errors, correcting where possible
    ''' </summary>
    '''
    ''' <param name="key">
    ''' A winapp2.ini ExcludeKey format <c>iniKey2</c> to be checked for correctness
    ''' </param>
    '''
    ''' <param name="hasF">
    ''' Indicates whether the entry excludes any filesystem locations
    ''' </param>
    '''
    ''' <param name="hasR">
    ''' Indicates whether the entry excludes any registry locations
    ''' </param>
    Private Sub pExcludeKey(result As EntryLintResult,
                            key As iniKey2,
                            ByRef hasF As Boolean,
                            ByRef hasR As Boolean)

        Dim hasValidFlags = key.vHasAny({"FILE|", "PATH|", "REG|"})
        If Not hasValidFlags Then hasValidFlags = checkExcludeFlags(result, key)

        fullKeyErr(result, key, "ExcludeKey has too many flags", lintFlags.ShouldScan AndAlso If(key.vHas("REG|"), key.Value.Split(CChar("|")).Length > 2, key.Value.Split(CChar("|")).Length > 3))


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
                           key.Value, If(patternPipe >= 0, key.Value.Insert(patternPipe, "\"), ""))

            Case key.vHas("REG|")

                hasR = True

                chkPathFormatValidity(result, key, True)
                fullKeyErr(result, key, "ExcludeKey contains REG flag in BleachBit flavor", CurrentWinappFlavor = WinappFlavor.BleachBit)

            Case Else

                checkExcludeFlags(result, key)

        End Select


    End Sub

    ''' <summary>
    ''' Assesses the formatting of ExcludeKey format <c>iniKey2</c>s to see if the flag (FILE, PATH, REG)
    ''' is malformatted. Attempts to repair when possible.
    ''' </summary>
    '''
    ''' <param name="key">
    ''' A winapp2.ini ExcludeKey format <c>iniKey2</c> to be checked for correctness
    ''' </param>
    Private Function checkExcludeFlags(result As EntryLintResult, key As iniKey2) As Boolean

        Dim matches = HasFlagRegex.Matches(key.Value)

        ' If we're not checking flags, we should at least indicate whether or not valid ones are present
        If Not lintFlags.ShouldScan Then Return HasFlagRegex.IsMatch(key.Value)

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
    ''' Sorts a typed bucket alphabetically with winapp2.ini precedence applied to the key values
    ''' </summary>
    '''
    ''' <param name="entry">
    ''' The <c>winapp2entry2</c> whose bucket will be sorted
    ''' </param>
    '''
    ''' <param name="keyType">
    ''' The key type name identifying which bucket to sort
    ''' </param>
    '''
    ''' <param name="hadDuplicatesRemoved">
    ''' Indicates that keys have been removed from the bucket
    ''' </param>
    Private Sub sortKeys2(entry As winapp2entry2,
                          keyType As String,
                          hadDuplicatesRemoved As Boolean)

        Dim bucketIdx = winapp2entry2.GetBucketIndex(keyType)
        If bucketIdx < 0 Then Return

        Dim bucket = entry.KeyLists(bucketIdx)

        If bucket.Count <= 1 OrElse Not lintAlpha.ShouldScan Then Return

        Dim keyValues As New strList
        Dim indices As New List(Of Integer)

        For i = 0 To bucket.Count - 1
            keyValues.add(bucket(i).Value)
            indices.Add(i)
        Next

        Dim sortedKeyValues = replaceAndSort(keyValues, "|", " \ \")

        Dim keysOutOfPlace = False
        findOutOfPlace(keyValues, sortedKeyValues, keyType, indices, keysOutOfPlace)

        If Not (keysOutOfPlace OrElse hadDuplicatesRemoved) AndAlso
               (lintAlpha.fixFormat OrElse lintWrongNums.fixFormat OrElse lintExtraNums.fixFormat) Then Return

        ' Reassign values and renumber, mirroring keyList.renumberKeys semantics
        For i = 0 To bucket.Count - 1
            bucket(i).Value = sortedKeyValues.Items(i)
            bucket(i).Name = keyType & CStr(i + 1)
        Next

    End Sub

    ''' <summary>
    ''' Prints an error when data is received that does not match an expected value
    ''' </summary>
    '''
    ''' <param name="context">
    ''' The key or entry that contains the error, used to identify the error location in output
    ''' </param>
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
    ''' Indicates that the error condition is present
    ''' <br/> Optional, Default: <c> True </c>
    ''' </param>
    Private Sub inputMismatchErr(result As EntryLintResult,
                                 err As String,
                                 received As String,
                                 expected As String,
                                 Optional cond As Boolean = True)

        result.RecordError(err, {$"Expected: {expected}", $"Found:    {received}"}, cond)

    End Sub

    ''' <summary>
    ''' Prints an error followed by the [Full Name *] of the entry to which it belongs
    ''' </summary>
    '''
    ''' <param name="cond">
    ''' Indicates that the error condition is present
    ''' </param>
    '''
    ''' <param name="entry">
    ''' The <c>winapp2entry2</c> containing an error
    ''' </param>
    '''
    ''' <param name="errTxt">
    ''' A description of the error as it will be displayed to the user
    ''' </param>
    Private Sub fullNameErr(result As EntryLintResult,
                            cond As Boolean,
                            errTxt As String)

        result.RecordError(errTxt, Array.Empty(Of String)(), cond)

    End Sub

    ''' <summary>
    ''' Prints an error whose output text contains an <c>iniKey2</c> string, optionally correcting that value with one that is provided
    ''' </summary>
    '''
    ''' <param name="key">
    ''' The <c>iniKey2</c> containing an error
    ''' </param>
    '''
    ''' <param name="err">
    ''' A description of the error as it will be displayed to the user
    ''' </param>
    '''
    ''' <param name="cond">
    ''' Indicates that the error condition(s) are present (including any <c> lintRule.shouldScans </c>)
    ''' <br/> Optional, Default: <c> True </c>
    ''' </param>
    '''
    ''' <param name="repCond">
    ''' Indicates that the repair function should run
    ''' <br/> Optional, Default: <c> False </c>
    ''' </param>
    '''
    ''' <param name="newVal">
    ''' The corrected value with which to replace the incorrect correct value held by <paramref name="repairVal"/>
    ''' <br/> Optional, Default: <c> "" </c>
    ''' </param>
    '''
    ''' <param name="repairVal">
    ''' The incorrect value
    ''' <br/> Optional, Default: <c> "" </c>
    ''' </param>
    Private Sub fullKeyErr(result As EntryLintResult,
                           key As iniKey2,
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
    ''' Prints arbitrarily defined errors without a precondition
    ''' </summary>
    '''
    ''' <param name="context">
    ''' The entry name or key string identifying where the error was found
    ''' </param>
    '''
    ''' <param name="err">
    ''' A description of the error as it will be displayed to the user
    ''' </param>
    '''
    ''' <param name="lines">
    ''' Any additional error information to be printed alongside the description
    ''' </param>
    Private Sub customErr(context As String,
                          err As String,
                          lines As String())

        gLog(err, ascend:=True)

        Dim out = $"Error in {context}: {Environment.NewLine} {err}"
        cwl(out)
        MostRecentLintLog.AppendLine(out)

        For Each errStr In lines

            cwl(errStr)
            gLog($"  {errStr}")
            MostRecentLintLog.AppendLine(errStr)

        Next

        gLog(descend:=True)
        cwl()

        MostRecentLintLog.AppendLine()

        ErrorsFound += 1

    End Sub

    ''' <summary>
    ''' Replace a given string with a new value if the fix condition is met
    ''' </summary>
    '''
    ''' <param name="param">
    ''' The condition under which the string should be replaced
    ''' </param>
    '''
    ''' <param name="currentValue">
    ''' A pointer to the string to be replaced
    ''' </param>
    '''
    ''' <param name="newValue">
    ''' The replacement value for <paramref name="currentValue"/>
    ''' </param>
    Private Sub fixStr(param As Boolean,
                       ByRef currentValue As String,
                       newValue As String)

        If Not param Then Return

        gLog($"  Changing '{currentValue}' to '{newValue}'", ascend:=True, descend:=True, buffr:=True)
        currentValue = newValue

    End Sub

End Module
