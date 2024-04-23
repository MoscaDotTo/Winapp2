'    Copyright (C) 2018-2024 Hazel Ward
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
    ''' The winapp2.ini file that will be linted 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2021-11-13 | Code last updated: 2021-11-13
    Public Property winappDebugFile1 As New iniFile(Environment.CurrentDirectory, "winapp2.ini", mExist:=True)

    ''' <summary> 
    ''' The save path for the linted file. Overwrites the input file by default 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2021-11-13 | Code last updated: 2021-11-13
    Public Property winappDebugFile3 As New iniFile(Environment.CurrentDirectory, "winapp2.ini", "winapp2-debugged.ini")

    ''' <summary> 
    ''' Indicates that some but not all repairs will run 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2021-11-13 | Code last updated: 2021-11-13
    Public Property RepairSomeErrsFound As Boolean = False

    ''' <summary>
    ''' Indicates that the scan settings have been modified from their defaults 
    ''' <br /> Default: <c> False </c> 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2021-11-13 | Code last updated: 2021-11-13
    Public Property ScanSettingsChanged As Boolean = False

    ''' <summary>
    ''' Indicates that the module settings have been modified from their defaults 
    ''' <br/> Default: <c> False </c> 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2021-11-13 | Code last updated: 2021-11-13
    Public Property ModuleSettingsChanged As Boolean = False

    ''' <summary> 
    ''' Indicates that the any changes made by the linter should be saved back to disk 
    ''' <br/> Default: <c> False </c> 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2021-11-13 | Code last updated: 2021-11-13
    Public Property SaveChanges As Boolean = False

    ''' <summary> 
    ''' Indicates that the linter should attempt to repair errors it finds 
    ''' <br/> Default: <c> True </c>
    ''' </summary>
    ''' 
    ''' Docs last updated: 2021-11-13 | Code last updated: 2021-11-13
    Public Property RepairErrsFound As Boolean = True

    ''' <summary> 
    ''' The number of errors found during the lint 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2021-11-13 | Code last updated: 2021-11-13
    Public Property ErrorsFound As Integer = 0

    ''' <summary> 
    ''' The list of all entry names found during the lint, used to check for duplicates 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2021-11-13 | Code last updated: 2021-11-13
    Private Property allEntryNames As New strList

    ''' <summary> 
    ''' The winapp2ool logslice from the most recent Lint run 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2021-11-13 | Code last updated: 2021-11-13
    Public Property MostRecentLintLog As String = ""

    ''' <summary> 
    ''' The current rules for scans and repairs 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2021-11-13 | Code last updated: 2021-11-13
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
        New lintRule(False, False, "Optimizations", "situations where keys can be merged (experimental)", "automatic merging of keys (experimental)"),
        New lintRule(False, False, "Potentially Duplicate Keys", "duplicated keys between multiple entries", "repair not yet supported")
    }

    ''' <summary> 
    ''' Controls scan/repairs for CamelCasing issues 
    ''' <br /> Default: <c> True </c> 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2021-11-13 | Code last updated: 2021-11-13
    Private Property lintCasing As lintRule = Rules(0)

    ''' <summary> 
    ''' Controls scan/repairs for alphabetization issues 
    ''' <br /> Default: <c> True </c>
    ''' </summary>
    ''' 
    ''' Docs last updated: 2021-11-13 | Code last updated: 2021-11-13
    Private Property lintAlpha As lintRule = Rules(1)

    ''' <summary> 
    ''' Controls scan/repairs for incorrectly numbered keys 
    ''' <br /> Default: <c> True </c>
    ''' </summary>
    ''' 
    ''' Docs last updated: 2021-11-13 | Code last updated: 2021-11-13
    Private Property lintWrongNums As lintRule = Rules(2)


    ''' <summary> 
    ''' Controls scan/repairs for parameters inside of FileKeys 
    ''' <br /> Default: <c> True </c>
    ''' </summary>
    ''' 
    ''' Docs last updated: 2021-11-13 | Code last updated: 2021-11-13
    Private Property lintParams As lintRule = Rules(3)


    ''' <summary> 
    ''' Controls scan/repairs for flags in ExcludeKeys and FileKeys 
    ''' <br /> Default: <c> True </c> 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2021-11-13 | Code last updated: 2021-11-13
    Private Property lintFlags As lintRule = Rules(4)

    ''' <summary> 
    ''' Controls scan/repairs for improper slash usage 
    ''' <br /> Default: <c> True </c> 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2021-11-13 | Code last updated: 2021-11-13
    Private Property lintSlashes As lintRule = Rules(5)

    ''' <summary> 
    ''' Controls scan/repairs for missing or True Default values 
    ''' <br /> Default: <c> True </c> 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2021-11-13 | Code last updated: 2021-11-13
    Private Property lintDefaults As lintRule = Rules(6)

    ''' <summary>
    ''' Controls scan/repairs for duplicate values 
    ''' <br /> Default: <c> True </c>
    ''' </summary>
    ''' 
    ''' Docs last updated: 2021-11-13 | Code last updated: 2021-11-13
    Private Property lintDupes As lintRule = Rules(7)

    ''' <summary> 
    ''' Controls scan/repairs for keys with numbers they shouldn't have 
    ''' <br /> Default: <c> True </c>
    ''' </summary>
    ''' 
    ''' Docs last updated: 2021-11-13 | Code last updated: 2021-11-13
    Private Property lintExtraNums As lintRule = Rules(8)

    ''' <summary> 
    ''' Controls scan/repairs for keys which should only occur once 
    ''' <br /> Default: <c> True </c>
    ''' </summary>
    ''' 
    ''' Docs last updated: 2021-11-13 | Code last updated: 2021-11-13
    Private Property lintMulti As lintRule = Rules(9)

    ''' <summary> 
    ''' Controls scan/repairs for keys with invlaid values 
    ''' <br /> Default: <c> True </c> 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2021-11-13 | Code last updated: 2021-11-13
    Private Property lintInvalid As lintRule = Rules(10)

    ''' <summary> 
    ''' Controls scan/repairs for winapp2.ini syntax errors 
    ''' <br /> Default: <c> True </c>
    ''' </summary>
    ''' 
    ''' Docs last updated: 2021-11-13 | Code last updated: 2021-11-13
    Private Property lintSyntax As lintRule = Rules(11)

    ''' <summary> 
    ''' Controls scan/repairs for invalid file or regsitry paths 
    ''' <br /> Default: <c> True </c> 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2021-11-13 | Code last updated: 2021-11-13
    Private Property lintPathValidity As lintRule = Rules(12)

    ''' <summary> 
    ''' Controls scan/repairs for improper use of semicolons 
    ''' <br /> Default: <c> True </c>
    ''' </summary>
    ''' 
    ''' Docs last updated: 2021-11-13 | Code last updated: 2021-11-13
    Private Property lintSemis As lintRule = Rules(13)

    ''' <summary> 
    ''' Controls scan/repairs for keys that can be merged into eachother (FileKeys only currently) 
    ''' <br /> Default: <c> False </c> 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2021-11-13 | Code last updated: 2021-11-13
    Public Property lintOpti As lintRule = Rules(14)

    ''' <summary> 
    ''' Controls scan/repairs for keys that may possibly exist in more than one entry 
    ''' <br /> Default: <c> False </c> 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2021-11-13 | Code last updated: 2021-11-13
    Private Property lintMultiDupe As lintRule = Rules(15)

    ''' <summary> 
    ''' Regex to detect long form registry paths 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2021-11-13 | Code last updated: 2021-11-13
    Private Property longReg As New Regex("HKEY_(C(URRENT_(USER$|CONFIG$)|LASSES_ROOT$)|LOCAL_MACHINE$|USERS$)")

    ''' <summary> 
    ''' Regex to detect short form registry paths 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2021-11-13 | Code last updated: 2021-11-13
    Private Property shortReg As New Regex("HK(C(C$|R$|U$)|LM$|U$)")

    ''' <summary> 
    ''' Regex to detect valid LangSecRef numbers 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2021-11-13 | Code last updated: 2021-11-13
    Private Property secRefNums As New Regex("30(0([1-6])|2([1-9])|3([0-8]))")

    ''' <summary> 
    ''' Regex to detect valid drive letter parameters 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2021-11-13 | Code last updated: 2021-11-13
    Private Property driveLtrs As New Regex("[a-zA-Z]:")

    ''' <summary> 
    ''' Regex to detect potential %EnvironmentVariables% 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2021-11-13 | Code last updated: 2021-11-13
    Private Property envVarRegex As New Regex("%[A-Za-z0-9]*%")

    ''' <summary> 
    ''' Indicates that Default keys should have their values auited instead of being considered invalid for existing 
    ''' <br /> Default: <c> False </c> 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2021-11-13 | Code last updated: 2021-11-13
    Public Property overrideDefaultVal As Boolean = False

    ''' <summary> 
    ''' The expected value for Default keys when auditing their values 
    ''' <br/> Default: <c> Faalse </c> 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2021-11-13 | Code last updated: 2021-11-13
    Public Property expectedDefaultValue As Boolean = False

    ''' <summary>
    ''' Handles the commandline args for <c> WinappDebug </c> <br />
    ''' WinappDebug commandline args: <br />
    ''' <c> -c </c> enable saving of changes made by the linter
    ''' </summary>
    ''' 
    ''' Docs last updated: 2021-11-13 | Code last updated: 2021-11-13
    Public Sub HandleLintCmdLine()

        InitDefaultLintSettings()
        invertSettingAndRemoveArg(SaveChanges, "-c")
        getFileAndDirParams(winappDebugFile1, New iniFile, winappDebugFile3)

        If Not cmdargs.Contains("UNIT_TESTING_HALT") Then InitDebug()

    End Sub

    ''' <summary> 
    ''' Validates winapp2.ini, then sets up the output window before sending it off to the linter.
    ''' After linting, reports the results of the lint to the user 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2021-11-13 | Code last updated: 2021-11-13
    Public Sub InitDebug()

        If Not enforceFileHasContent(winappDebugFile1) Then Return

        Dim wa2 As New winapp2file(winappDebugFile1)

        clrConsole()

        print(3, "Beginning analysis of winapp2.ini", trailr:=True)
        gLog("Beginning lint", leadr:=True, ascend:=True)

        MostRecentLintLog = ""

        Debug(wa2)

        gLog(descend:=True)
        gLog("Lint complete")
        setHeaderText("Lint complete")
        print(4, "Completed analysis of winapp2.ini", conjoin:=True)
        print(0, $"{ErrorsFound} possible errors were detected.")
        print(0, $"Number of entries {winappDebugFile1.Sections.Count}", trailingBlank:=True)

        RewriteChanges(wa2)

        print(0, anyKeyStr, closeMenu:=True)
        crk()

    End Sub

    ''' <summary> 
    ''' Sends the entries in a winapp2.ini format <c> iniFile </c> into specific format and syntax checking routines 
    ''' </summary>
    ''' 
    ''' <param name="fileToBeDebugged"> 
    ''' A <c> winapp2file </c> to be linted 
    ''' </param>
    ''' 
    ''' Docs last updated: 2021-11-13 | Code last updated: 2021-11-13
    Public Sub Debug(ByRef fileToBeDebugged As winapp2file)

        If fileToBeDebugged Is Nothing Then argIsNull(NameOf(fileToBeDebugged)) : Return

        ErrorsFound = 0
        allEntryNames = New strList
        gLog(ascend:=True)

        For Each entryList In fileToBeDebugged.Winapp2entries

            If entryList.Count = 0 Then Continue For
            entryList.ForEach(Sub(entry) ProcessEntry(entry))

        Next

        resetKeyTrackers()
        fileToBeDebugged.rebuildToIniFiles()
        AlphabetizeEntries(fileToBeDebugged)

    End Sub

    ''' <summary> Validates the basic structure of a <c> winapp2entry </c> and sends off its individual keys for more specific analysis </summary>
    ''' <param name="entry"> A <c> winapp2entry </c> to be audited for syntax errors </param>
    ''' Docs last updated: 2021-11-13 | Code last updated: 2022-01-11
    Private Sub ProcessEntry(ByRef entry As winapp2entry)

        gLog($"Processing entry {entry.Name}", buffr:=True)

        Dim hasFileExcludes = False
        Dim hasRegExcludes = False

        fullNameErr(allEntryNames.chkDupes(entry.Name), entry,
                    "Duplicate entry name detected")

        fullNameErr(Not entry.Name.EndsWith(" *", StringComparison.InvariantCulture), entry,
                    "All entries must end in ' *'")

        ValidateKeys(entry)

        For Each lst In entry.KeyListList

            If lst.KeyType = "Error" Then Continue For

            Select Case lst.KeyType

                Case "DetectFile"

                    processKeyList(lst, AddressOf pDetectFile)

                Case "FileKey"

                    processKeyList(lst, AddressOf pFileKey)

                Case Else

                    processKeyList(lst, AddressOf voidDelegate, hasFileExcludes, hasRegExcludes)

            End Select

        Next

        fullNameErr(lintSyntax.ShouldScan AndAlso entry.SectionKey.KeyCount <> 0 AndAlso entry.LangSecRef.KeyCount <> 0, entry,
                    "Section key found alongside LangSecRef key, but only one should be present")

        fullNameErr(lintSyntax.ShouldScan AndAlso Not entry.SectionKey.KeyCount <> 0 AndAlso Not entry.LangSecRef.KeyCount <> 0, entry,
                    "Entry has no valid classifier key (LangSecRef, Section)")

        fullNameErr(Not (entry.DetectFiles.KeyCount <> 0 OrElse entry.Detects.KeyCount <> 0 OrElse entry.DetectOS.KeyCount <> 0 OrElse entry.SpecialDetect.KeyCount <> 0), entry,
                    "Entry has no valid detection keys (Detect, DetectFile, DetectOS, SpecialDetect)")

        fullNameErr(lintSyntax.ShouldScan AndAlso Not (entry.FileKeys.KeyCount <> 0 OrElse entry.RegKeys.KeyCount <> 0), entry,
                    "Entry has no valid deletion keys (FileKey, RegKey)")

        fullNameErr(lintSyntax.ShouldScan AndAlso hasFileExcludes AndAlso Not (entry.FileKeys.KeyCount <> 0 OrElse entry.RegKeys.KeyCount <> 0), entry,
                    "Entry has ExcludeKeys but no valid FileKeys or RegKeys")

        fullNameErr(hasFileExcludes AndAlso Not entry.FileKeys.KeyCount <> 0, entry,
                    "Entry has ExcludeKeys pointing to file system locations but no FileKeys")

        fullNameErr(hasRegExcludes AndAlso Not entry.RegKeys.KeyCount <> 0, entry,
                    "Entry has ExcludeKeys pointing to registry locations but no RegKeys")

        fullNameErr(lintDefaults.ShouldScan AndAlso entry.DefaultKey.KeyCount > 0 AndAlso Not overrideDefaultVal, entry,
                    "Entry has a Default key where there should be none")

        If lintDefaults.fixFormat And entry.DefaultKey.KeyCount > 0 And Not overrideDefaultVal Then entry.DefaultKey.Keys.Clear()

        If Not overrideDefaultVal Then gLog($"Finished processing {entry.Name}", buffr:=True) : Return

        Dim expected = tsInvariant(expectedDefaultValue)

        If entry.DefaultKey.KeyCount > 0 Then

            Dim key = entry.DefaultKey.Keys(0)
            fullKeyErr(key, "Incorrect value for Default Key found", lintDefaults.ShouldScan AndAlso Not key.Value = expected, lintDefaults.fixFormat, key.Value, expected)
            Return

        End If

        Dim NoDefaultKeyErrorText = "No Default Key found"
        fullNameErr(True, entry, NoDefaultKeyErrorText)
        entry.DefaultKey.add(New iniKey($"Default={expected}"))

        gLog($"Finished processing {entry.Name}", buffr:=True)

    End Sub

    ''' <summary> 
    ''' Checks the basic structure of all <c>iniKeys </c> in a <c> winapp2entry </c>,
    ''' attempts to repair some keys and place them back into their appropriate <c> keyList </c>,
    ''' and removes any that are too problematic to continue with 
    ''' </summary>
    ''' 
    ''' <param name="entry"> 
    ''' A <c> winapp2entry </c> whose <c> iniKeys </c> will be audited for basic syntax correctness 
    ''' </param>
    ''' 
    ''' Docs last updated: 2021-11-13 | Code last updated: 2021-11-13
    Private Sub ValidateKeys(ByRef entry As winapp2entry)

        For Each lst In entry.KeyListList

            Dim brokenKeys As New keyList
            lst.Keys.ForEach(Sub(key) brokenKeys.add(key, Not cValidity(key)))
            lst.remove(brokenKeys.Keys)
            entry.ErrorKeys.remove(brokenKeys.Keys)

        Next

        Dim toRemove As New keyList

        For Each key In entry.ErrorKeys.Keys

            For Each lst In entry.KeyListList

                If lst.KeyType = "Error" Then Continue For

                Dim TypeMatch = key.typeIs(lst.KeyType)
                lst.add(key, TypeMatch)
                toRemove.add(key, TypeMatch)

            Next

        Next

        entry.ErrorKeys.remove(toRemove.Keys)

    End Sub

    ''' <summary> 
    ''' Alphabetizes all the entries in a winapp2.ini file and observes any that were out of place 
    ''' </summary>
    ''' 
    ''' <param name="winapp">
    ''' The <c> winapp2file </c> whose entries will be alphabetized 
    ''' </param>
    ''' 
    ''' Docs last updated: 2021-11-13 | Code last updated: 2021-11-13
    Private Sub AlphabetizeEntries(ByRef winapp As winapp2file)

        For Each innerFile In winapp.EntrySections

            Dim unsortedEntryList = innerFile.namesToStrList
            Dim sortedEntryList = sortEntryNames(innerFile)
            If lintAlpha.ShouldScan Then findOutOfPlace(unsortedEntryList, sortedEntryList, "Entry", innerFile.getLineNumsFromSections)
            If lintAlpha.fixFormat Then innerFile.sortSections(sortedEntryList)

        Next

    End Sub

    ''' <summary> 
    ''' Writes any changes made during the lint back to disk, correcting any errors that were found and repaired 
    ''' </summary>
    ''' 
    ''' <param name="winapp2file">
    ''' The <c> winapp2file </c> that was linted 
    ''' </param>
    ''' 
    ''' Docs last updated: 2021-11-13 | Code last updated: 2021-11-13
    Private Sub RewriteChanges(ByRef winapp2file As winapp2file)

        If SaveChanges Then

            print(0, "Saving changes, do not close winapp2ool or data loss may occur...", leadingBlank:=True)
            winappDebugFile3.overwriteToFile(winapp2file.winapp2string)
            print(0, "Finished saving changes.", trailingBlank:=True)

        End If

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
    ''' <param name="LineCountList"> 
    ''' The line numbers associated with the lines in <c> <paramref name="someList"/> </c> 
    ''' </param>
    ''' 
    ''' <param name="oopBool"> 
    ''' Tracking variable indicating that alphabetization errors have been found 
    ''' <br/>  Optional, Default: <c> False </c> 
    ''' </param>
    ''' 
    ''' Docs last updated: 2023-07-22 | Code last updated: 2021-11-13
    Private Sub findOutOfPlace(ByRef someList As strList,
                               ByRef sortedList As strList,
                               findType As String,
                               ByRef LineCountList As List(Of Integer),
                               Optional ByRef oopBool As Boolean = False)


        If someList.Count < 2 Then Return

        Dim misplacedEntries As New strList
        Dim initialNeighbors = someList.getNeighborList
        Dim sortedNeighbors = sortedList.getNeighborList

        For i = 0 To someList.Count - 1

            Dim HasSameNeighbor = initialNeighbors(i).Key = sortedNeighbors(sortedList.Items.IndexOf(someList.Items(i))).Key
            Dim HasSamePosition = initialNeighbors(i).Value = sortedNeighbors(sortedList.Items.IndexOf(someList.Items(i))).Value
            misplacedEntries.add(someList.Items(i), Not (HasSameNeighbor AndAlso HasSamePosition))

        Next

        For Each entry In misplacedEntries.Items

            Dim recInd = someList.indexOf(entry)
            Dim sortInd = sortedList.indexOf(entry)
            Dim curLine = LineCountList(recInd)
            Dim sortLine = LineCountList(sortInd)

            If (recInd = sortInd OrElse curLine = sortLine) Then Continue For

            entry = If(findType = "Entry", entry, $"{findType & (recInd + 1)}={entry}")
            oopBool = True

            customErr(LineCountList(recInd), $"{findType} alphabetization",
                                            {$"{entry} appears to be out of place",
                                             $"Current line: {curLine}",
                                             $"Expected line: {sortLine}"})

        Next

    End Sub

    ''' <summary> 
    ''' Hands off each <c> iniKey </c> in a winapp2.ini format <c> keyList </c> to be audited for correctness 
    ''' </summary>
    ''' 
    ''' <param name="kl">
    ''' A <c> keyList </c> of a particular <c> keyType </c> to be audited 
    ''' </param>
    ''' 
    ''' <param name="processKey"> 
    ''' The <c> function </c> that audits the keys of the <c> KeyType </c> provided in <c> <paramref name="kl"/> </c> <br/> 
    ''' <c> VoidDelegate </c> if no further operations are needed outside of the basic formatting checks 
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
    ''' 
    ''' Docs last updated: 2021-11-13 | Code last updated: 2021-11-13
    Private Sub processKeyList(ByRef kl As keyList,
                               processKey As Func(Of iniKey, iniKey),
                               Optional ByRef hasF As Boolean = False,
                               Optional ByRef hasR As Boolean = False)

        If kl.KeyCount = 0 Then Return

        gLog($"Processing {kl.KeyType}s", ascend:=True, buffr:=True)

        Dim curNum = 1
        Dim curStrings As New strList
        Dim dupes As New keyList
        Dim kt = kl.KeyType

        For Each key In kl.Keys

            Select Case kt

                Case "ExcludeKey"

                    cFormat(key, curNum, curStrings, dupes)
                    pExcludeKey(key, hasF, hasR)

                Case "Detect", "DetectFile"

                    If key.typeIs("Detect") Then chkPathFormatValidity(key, True)
                    cFormat(key, curNum, curStrings, dupes, kl.KeyCount = 1)
                    If lintMultiDupe.ShouldScan Then cDuplicateKeysBetweenEntries(key)

                Case "RegKey"

                    chkPathFormatValidity(key, True)
                    cFormat(key, curNum, curStrings, dupes)
                    If lintMultiDupe.ShouldScan Then cDuplicateKeysBetweenEntries(key)

                Case "Warning", "DetectOS", "SpecialDetect", "LangSecRef", "Section", "Default"


                    If curNum > 1 AndAlso lintMulti.ShouldScan Then

                        fullKeyErr(key, $"Multiple {key.KeyType} detected.")
                        dupes.add(key, lintMulti.fixFormat)

                    End If

                    cFormat(key, curNum, curStrings, dupes, True)

                    If key.typeIs("SpecialDetect") Then chkCasing(key, {"DET_CHROME", "DET_MOZILLA", "DET_THUNDERBIRD", "DET_OPERA"}, key.Value)
                    fullKeyErr(key, "LangSecRef holds an invalid value.", lintInvalid.ShouldScan And key.typeIs("LangSecRef") And Not secRefNums.IsMatch(key.Value))

                Case Else

                    cFormat(key, curNum, curStrings, dupes)

            End Select

            key = processKey(key)

        Next

        kl.remove(dupes.Keys)
        sortKeys(kl, dupes.KeyCount > 0)

        If kl.typeIs("FileKey") And lintOpti.ShouldScan Then cOptimization(kl)

        gLog(descend:=True)

    End Sub

    ''' <summary> 
    ''' This function does nothing by design, used when a method or function expects to be passed a function 
    ''' who modifies and iniKey on a KeyType where we don't want to modify the keys 
    ''' </summary>
    ''' 
    ''' <param name="key"> 
    ''' An <c> iniKey </c> with which to do nothing 
    ''' </param>
    ''' Docs last updated: 2021-11-13 | Code last updated: 2021-11-13
    Private Function voidDelegate(key As iniKey) As iniKey

        Return key

    End Function

    ''' <summary> 
    ''' Does some basic formatting checks that apply to all winapp2.ini format <c> iniKeys </c>
    ''' </summary>
    ''' 
    ''' <param name="key"> 
    ''' An <c> iniKey </c> whose format will be audited 
    ''' </param>
    ''' 
    ''' <param name="keyNumber"> 
    ''' The current expected key number for numbered keys
    ''' </param>
    ''' 
    ''' <param name="keyValues"> 
    ''' The current list of observed <c> iniKey </c> values
    ''' </param>
    ''' 
    ''' <param name="dupeList">
    ''' A tracking list of <c> iniKeys </c> with duplicate values 
    ''' </param>
    ''' 
    ''' <param name="noNumbers"> 
    ''' Indicates that the current set of keys should not be numbered 
    ''' </param>
    ''' 
    ''' Docs last updated: 2021-11-13 | Code last updated: 2021-11-13
    Private Sub cFormat(ByRef key As iniKey,
                        ByRef keyNumber As Integer,
                        ByRef keyValues As strList,
                        ByRef dupeList As keyList,
                        Optional noNumbers As Boolean = False)

        ' Check for duplicates
        If keyValues.contains(key.Value, True) Then
            Dim dupeKeyStr = $"{key.KeyType}{If(Not noNumbers, (keyValues.Items.IndexOf(key.Value) + 1).ToString(Globalization.CultureInfo.InvariantCulture), "")}={key.Value}"
            If lintDupes.ShouldScan Then customErr(key.LineNumber, "Duplicate key value found", {$"Key:            {key.toString}", $"Duplicates:     {dupeKeyStr}"})
            dupeList.add(key, lintDupes.fixFormat)
        Else
            keyValues.add(key.Value)
        End If
        ' Check for both types of numbering errors (incorrect and unneeded) 
        Dim hasNumberingError = If(noNumbers, Not key.nameIs(key.KeyType), Not key.nameIs(key.KeyType & keyNumber))
        Dim numberingErrStr = If(noNumbers, "Detected unnecessary numbering.", $"{key.KeyType} entry is incorrectly numbered.")
        Dim fixedStr = If(noNumbers, key.KeyType, key.KeyType & keyNumber)
        gLog($"Input mismatch error in {key.toString}", hasNumberingError, indent:=True)
        inputMismatchErr(key.LineNumber, numberingErrStr, key.Name, fixedStr, If(noNumbers, lintExtraNums.ShouldScan, lintWrongNums.ShouldScan) And hasNumberingError)
        fixStr(If(noNumbers, lintExtraNums.fixFormat, lintWrongNums.fixFormat) And hasNumberingError, key.Name, fixedStr)
        ' Scan for and fix any use of incorrect slashes (except in Warning keys) or trailing semicolons
        fullKeyErr(key, "Forward slash (/) detected in lieu of backslash (\).", Not (key.typeIs("Warning") Or key.typeIs("RegKey")) And lintSlashes.ShouldScan And key.vHas("/"),
                                                                                                        lintSlashes.fixFormat, key.Value, key.Value.Replace("/", "\"))
        fullKeyErr(key, "Trailing semicolon (;).", key.toString.Last = CChar(";") And lintSemis.ShouldScan, lintSemis.fixFormat, key.Value, key.Value.TrimEnd(CChar(";")))
        ' Do some formatting checks for environment variables if needed
        If {"FileKey", "ExcludeKey", "DetectFile"}.Contains(key.KeyType) Then cEnVar(key)
        keyNumber += 1
    End Sub

    ''' <summary>
    ''' Attempts to fix any broken environment variables in a given <c> iniKey </c> <br/> <br/>
    ''' This function will attempt to repair any environment variables that are missing leading or trailing % characters    
    ''' </summary>
    ''' <param name="key"> An <c> iniKey </c> whose value will be audited for syntax errors </param>
    ''' <param name="enVars"> The list of valid Environment Variables for Winapp2.ini </param>
    ''' <param name="cond"> The condition under which this scan should be run </param>
    ''' Docs last updated: 2024-04-22 | Code last updated: 2024-04-22
    Private Sub fixBrokenEnVars(ByRef key As iniKey, enVars As String(), cond As Boolean)

        If Not cond Then Return

        For Each enVar In enVars

            If Not key.vHas(enVar) Then Continue For

            Dim tmpRegex As New Regex(enVar)

            Dim trailingCharMissing As New Regex($"%{enVar}\\")
            Dim leadingCharMissing As New Regex($"^{enVar}%")
            Dim bothCharsMissing As New Regex($"^{enVar}\\")

            Dim msg = ""
            Dim replValue = ""
            Dim repairMade = False

            Select Case True

                Case trailingCharMissing.IsMatch(key.Value)

                    msg = "Environment Variable is missing trailing %"
                    replValue = key.Value.Replace($"%{enVar}", $"%{enVar}%")
                    repairMade = True

                Case leadingCharMissing.IsMatch(key.Value)

                    msg = "Environment Variable is missing leading %"
                    replValue = key.Value.Replace($"{enVar}%", $"%{enVar}%")
                    repairMade = True

                Case bothCharsMissing.IsMatch(key.Value)

                    msg = "Environment Variable is missing leading and trailing %"
                    replValue = key.Value.Replace($"{enVar}\", $"%{enVar}%\")
                    repairMade = True

                Case Else

                    ' This only happens because "AppData" is a substring of "LocalAppData" and will result in this code path being hit 
                    ' We can silently ignore this case 

            End Select

            fullKeyErr(key, msg, lintSyntax.ShouldScan AndAlso repairMade, lintSyntax.ShouldRepair, key.Value, replValue)

            If repairMade Then Exit For

        Next

    End Sub


    ''' <summary> Validates the formatting of any %EnvironmentVariables% in a given <c> iniKey </c> </summary>
    ''' <param name="key">The <c> iniKey </c> whose data will be audited for environment variable correctness </param>
    ''' Docs last updated: 2021-11-13 | Code last updated: 2024-04-22
    Private Sub cEnVar(ByRef key As iniKey)

        ' Valid Environmental Variables for winapp2.ini
        Dim enVars = {"AllUsersProfile", "AppData", "CommonAppData", "CommonProgramFiles",
        "Documents", "HomeDrive", "LocalAppData", "LocalLowAppData", "Music", "Pictures", "ProgramData", "ProgramFiles", "Public",
        "RootDir", "SystemDrive", "SystemRoot", "Temp", "Tmp", "UserName", "UserProfile", "Video", "WinDir"}

        fullKeyErr(key, "Double '%' found in environment variable", key.vHas("%%"), lintSyntax.ShouldRepair, key.Value, key.Value.Replace("%%", "%"))

        fixBrokenEnVars(key, enVars, key.vHas("%") AndAlso envVarRegex.Matches(key.Value).Count = 0 OrElse key.vHasAny(enVars) AndAlso Not key.vHas("%"))

        For Each m As Match In envVarRegex.Matches(key.Value)
            Dim strippedText = m.ToString.Trim(CChar("%"))
            chkCasing(key, enVars, strippedText)
        Next

        ' Environment variables should be trailed by a backslash 
        fullKeyErr(key, "Missing backslash (\) after %EnvironmentVariable%.", lintSlashes.ShouldScan And key.vHas("%") And Not key.vHasAny({"%|", "%\"}))

    End Sub

    ''' <summary> Attempts to insert missing equal signs (=) into <c> iniKeys </c> <br/> <br/> Returns <c> True </c> if the repair is 
    '''  successful, <c> False </c> otherwise </summary>
    ''' <param name="key"> A misformatted <c> iniKey </c> to attempt to repair </param>
    ''' <param name="cmds"> An array containing valid winapp2.ini <c> keyTypes </c> </param>
    ''' Docs last updated: 2021-11-13 | Code last updated: 2021-11-13
    Private Function fixMissingEquals(ByRef key As iniKey, cmds As String()) As Boolean
        gLog("Attempting missing equals repair", ascend:=True)
        For Each cmd In cmds
            If key.Name.ToUpperInvariant.Contains(cmd.ToUpperInvariant) Then
                Select Case cmd
                ' We don't expect numbers in these keys
                    Case "Default", "DetectOS", "Section", "LangSecRef", "Section", "SpecialDetect"
                        key.Value = key.Name.Replace(cmd, "")
                        key.Name = cmd
                        key.KeyType = cmd
                    Case Else
                        Dim newName = cmd
                        Dim withNums = key.Name.Replace(cmd, "")
                        For Each c As Char In withNums.ToCharArray
                            If Char.IsNumber(c) Then newName += c : Else Exit For
                        Next
                        key.Value = key.Name.Replace(newName, "")
                        key.Name = newName
                        key.KeyType = cmd
                End Select
                gLog($"Repair complete. Result: {key.toString}", indent:=True, descend:=True)
                ' Don't allow valueless keys in winapp2.ini 
                If key.Value.Length = 0 Then gLog("Repair failed, key will be removed.", descend:=True) : Return False
                Return True
            End If
        Next
        ' Return false if no valid command is found
        gLog("Repair failed, key will be removed.", descend:=True)
        Return False
    End Function

    ''' <summary> Does basic syntax and formatting audits that apply across all keys, returns <c> False </c> 
    ''' if a key is malformed or if a null argument is given </summary>
    ''' <param name="key"> A <c> iniKey </c> whose basic syntactic validity will be assessed </param>
    ''' Docs last updated: 2021-11-13 | Code last updated: 2022-12-01
    Private Function cValidity(ByRef key As iniKey) As Boolean
        If key Is Nothing Then argIsNull(NameOf(key)) : Return False
        Dim validCmds = {"Default", "DetectOS", "DetectFile", "Detect", "ExcludeKey",
                        "FileKey", "LangSecRef", "RegKey", "Section", "SpecialDetect", "Warning"}
        ' Attempt to fix the case where keys are missing an equal sign to delineate name and value 
        If key.typeIs("DeleteMe") Then
            gLog($"Broken Key Found: {key.Name}", indent:=True, ascend:=True)
            ' If we didn't find a fixable situation, delete the key
            Dim fixedMsngEq = fixMissingEquals(key, validCmds)
            If Not fixedMsngEq Then customErr(key.LineNumber, $"{key.Name} is missing a '=' or was not provided with a value. It will be deleted.", Array.Empty(Of String)()) : Return False
            fullKeyErr(key, "Missing '=' detected and repaired in key.", fixedMsngEq)
        End If
        ' Remove any instances of double backlashes because we don't expect them 
        If key.vHas("\\", True) Then
            fullKeyErr(key, "Extraneous backslashes (\\) detected", lintSlashes.ShouldScan)
            While (key.Value.Contains("\\") And lintSlashes.fixFormat)
                key.Value = key.Value.Replace("\\", "\")
            End While
        End If
        ' Check for leading or trailing whitespace, do this always as spaces in the name interfere with proper keyType identification
        If key.Name.StartsWith(" ", StringComparison.InvariantCulture) Or key.Name.EndsWith(" ", StringComparison.InvariantCulture) Or
            key.Value.StartsWith(" ", StringComparison.InvariantCulture) Or key.Value.EndsWith(" ", StringComparison.InvariantCulture) Then
            fullKeyErr(key, "Detected unwanted whitespace in iniKey", True)
            fixStr(True, key.Value, key.Value.Trim)
            fixStr(True, key.Name, key.Name.Trim)
            fixStr(True, key.KeyType, key.KeyType.Trim)
        End If
        ' Make sure the keyType is valid
        chkCasing(key, validCmds, key.KeyType)
        Return True
    End Function

    ''' <summary> Checks the <c> Value </c> or the <c> KeyType </c> of an <c> iniKey </c> against a given array of expected cased values, attempts 
    ''' to repair casing errors if possible </summary>
    ''' <param name="key"> The <c> iniKey </c> whose casing will be audited </param>
    ''' <param name="casedArray"> The array of expected cased values </param>
    ''' <param name="strToChk"> A pointer to the value being audited </param>
    ''' Docs last updated: 2021-11-13 | Code last updated: 2022-12-01
    Private Sub chkCasing(ByRef key As iniKey, casedArray As String(), strToChk As String)

        ' Get the properly cased string 
        Dim casedString As String = strToChk
        For Each casedText In casedArray
            If strToChk.Equals(casedText, StringComparison.InvariantCultureIgnoreCase) Then casedString = casedText
        Next

        ' Determine if there's a casing error
        Dim hasCasingErr = Not casedString.Equals(strToChk, StringComparison.InvariantCulture) And casedArray.Contains(casedString)
        Dim validData = String.Join(", ", casedArray)

        ' Inform the user if there are casing errors and fix them 
        fullKeyErr(key, $"{casedString} has a casing error.", hasCasingErr And lintCasing.ShouldScan, False, "", "")
        fixStr(hasCasingErr AndAlso key.Value.Contains(strToChk), key.Value, key.Value.Replace(strToChk, casedString))
        fixStr(hasCasingErr AndAlso key.Name.Contains(strToChk), key.Name, key.Name.Replace(key.KeyType, casedString))
        fixStr(hasCasingErr AndAlso key.KeyType.Contains(strToChk), key.KeyType, key.KeyType.Replace(key.KeyType, casedString))
        ' Inform the user about invalid data 
        fullKeyErr(key, $"Invalid data provided: {strToChk} in {key.toString}{Environment.NewLine}Valid data: {validData}", Not casedArray.Contains(casedString) And lintInvalid.ShouldScan)

    End Sub

    ''' <summary> Processes a FileKey format winapp2.ini <c> iniKey </c> and checks it for errors, correcting them where possible </summary>
    ''' <param name="key"> A winapp2.ini FileKey format <c> iniKey </c> to be checked for correctness </param>
    ''' Docs last updated: 2021-11-13 | Code last updated: 2021-11-13
    Public Function pFileKey(key As iniKey) As iniKey
        If key Is Nothing Then argIsNull(NameOf(key)) : Return key
        ' Pipe symbol checks
        Dim iteratorCheckerList = Split(key.Value, "|")
        If iteratorCheckerList.Length > 2 Then
            chkCasing(key, {"RECURSE", "REMOVESELF"}, iteratorCheckerList.Last)
            iteratorCheckerList = Split(key.Value, "|")
        End If
        fullKeyErr(key, "Missing pipe (|) in FileKey.", Not key.vHas("|"))
        ' The driveLtr check to allow entries that contain hard coded drive letters to contain colons. Since this is an edge case only likely to pop up in winapp3.ini (as far as official releases go)
        ' We'll assume that if the path contains a hard coded drive letter, any colon use is intentional and disable this check. 
        fullKeyErr(key, "Colon (:) found where there should be a semicolon (;)", key.Value.Contains(":") And Not driveLtrs.IsMatch(getFirstDir(key.Value)), lintSemis.fixFormat, key.Value, key.Value.Replace(":", ";"))
        ' Captures any incident of semi colons coming before the first pipe symbol
        fullKeyErr(key, "Semicolon (;) found before pipe (|).", lintSemis.ShouldScan And key.vHas(";") And (key.Value.IndexOf(";", StringComparison.InvariantCultureIgnoreCase) < key.Value.IndexOf("|", StringComparison.InvariantCultureIgnoreCase)))
        fullKeyErr(key, "Trailing semicolon (;) in parameters", lintSemis.ShouldScan And key.vHas(";|"), lintSemis.fixFormat, key.Value, key.Value.Replace(";|", "|"))
        ' Check for incorrect spellings of RECURSE or REMOVESELF
        If iteratorCheckerList.Length > 2 Then fullKeyErr(key, "RECURSE or REMOVESELF is incorrectly spelled, or there are too many pipe (|) symbols.", Not iteratorCheckerList(2).Contains("RECURSE") And Not iteratorCheckerList(2).Contains("REMOVESELF"))
        ' Check for missing pipe symbol on recurse and removeself, fix them if detected
        Dim flags As New List(Of String) From {"RECURSE", "REMOVESELF"}
        flags.ForEach(Sub(flagStr) fullKeyErr(key, $"Missing pipe (|) before {flagStr}.", lintFlags.ShouldScan And key.vHas(flagStr) And Not key.vHas($"|{flagStr}"), lintFlags.fixFormat, key.Value, key.Value.Replace(flagStr, $"|{flagStr}")))
        ' Make sure VirtualStore folders point to the correct place
        inputMismatchErr(key.LineNumber, "Incorrect VirtualStore location.", key.Value, "%LocalAppData%\VirtualStore\Program Files*\", key.vHas("\virtualStore\p", True) And Not key.vHasAny({"programdata", "program files*", "program*"}, True))
        ' Backslash checks, fix if detected
        fullKeyErr(key, "Backslash (\) found before pipe (|).", lintSlashes.ShouldScan And key.vHas("\|"), lintSlashes.fixFormat, key.Value, key.Value.Replace("\|", "|"))
        ' Get the parameters given to the file key and sort them 
        Dim keyParams As New winapp2KeyParameters(key)
        Dim argsStrings As New strList
        Dim dupeArgs As New strList
        ' Check for duplicate args
        For Each arg In keyParams.ArgsList
            If argsStrings.chkDupes(arg) And lintParams.ShouldScan Then
                customErr(key.LineNumber, $"{If(arg.Length = 0, "Empty", "Duplicate")} FileKey parameter found", {$"Command: {arg}"})
                dupeArgs.add(arg, lintParams.fixFormat)
            End If
        Next
        ' Remove any duplicate arguments from the key parameters and reconstruct keys we've modified above
        If lintParams.fixFormat Then
            dupeArgs.Items.ForEach(Sub(arg) keyParams.ArgsList.Remove(arg))
            keyParams.reconstructKey(key)
        End If
        If lintMultiDupe.ShouldScan Then cDuplicateKeysBetweenEntries(key)
        Return key
    End Function

    ''' <summary> Processes a DetectFile format <c> iniKey </c> and checks it for errors, correcting where possible </summary>
    ''' <param name="key"> A winapp2.ini DetectFile format <c> iniKey </c> to be checked for correctness </param>
    ''' Docs last updated: 2021-11-13 | Code last updated: 2021-11-13
    Private Function pDetectFile(key As iniKey) As iniKey
        ' Trailing Backslashes & nested wildcards
        fullKeyErr(key, "Trailing backslash (\) found in DetectFile", lintSlashes.ShouldScan _
    And key.Value.Last = CChar("\"), lintSlashes.fixFormat, key.Value, key.Value.TrimEnd(CChar("\")))
        If key.vHas("*") Then
            Dim splitDir = key.Value.Split(CChar("\"))
            For i = 0 To splitDir.Length - 1
                fullKeyErr(key, "Nested wildcard found in DetectFile", splitDir(i).Contains("*") And i <> splitDir.Length - 1)
            Next
        End If
        ' Make sure that DetectFile paths point to a filesystem location
        chkPathFormatValidity(key, False)
        Return key
    End Function

    ''' <summary> Audits the syntax of file system and registry paths </summary>
    ''' <param name="key"> An <c> iniKey </c> containing a registry or filesystem path to have its syntax validated </param>
    ''' <param name="isRegistry"> Indicates that the given <c> <paramref name="key"/> </c> is expected to hold a registry path </param>
    ''' Docs last updated: 2021-11-13 | Code last updated: 2021-11-13
    Private Sub chkPathFormatValidity(key As iniKey, isRegistry As Boolean)
        If Not lintPathValidity.ShouldScan Then Return
        ' Remove the flags from ExcludeKeys if we have them before getting the first directory portion
        Dim rootStr = If(key.KeyType <> "ExcludeKey", getFirstDir(key.Value), getFirstDir(pathFromExcludeKey(key)))
        ' Ensure that registry paths have a valid hive and file paths have either a variable or a drive letter
        fullKeyErr(key, "Invalid registry path detected.", isRegistry And Not longReg.IsMatch(rootStr) And Not shortReg.IsMatch(rootStr))
        fullKeyErr(key, "Invalid file system path detected.", Not isRegistry And Not driveLtrs.IsMatch(rootStr) And Not rootStr.StartsWith("%", StringComparison.InvariantCultureIgnoreCase))
    End Sub

    ''' <summary> Processes a list of ExcludeKey format <c> iniKeys </c> and checks them for errors, correcting where possible </summary>
    ''' <param name="key"> A winapp2.ini ExcludeKey format <c> iniKey </c> to be checked for correctness </param>
    ''' <param name="hasF"> Indicates whether the entry excludes any filesystem locations </param>
    ''' <param name="hasR"> Indicates whether the entry excludes any registry locations </param>
    ''' Docs last updated: 2021-11-13 | Code last updated: 2021-11-13
    Private Sub pExcludeKey(ByRef key As iniKey, ByRef hasF As Boolean, ByRef hasR As Boolean)
        Select Case True
            Case key.vHasAny({"FILE|", "PATH|"})
                hasF = True
                If lintPathValidity.ShouldScan Then
                    chkPathFormatValidity(key, False)
                    fullKeyErr(key, "Missing backslash (\) before pipe (|) in ExcludeKey.", key.vHas("|") And Not key.vHas("\|"))
                End If
            Case key.vHas("REG|")
                hasR = True
                chkPathFormatValidity(key, True)
            Case Else
                If key.Value.StartsWith("FILE", StringComparison.InvariantCulture) Or
                        key.Value.StartsWith("PATH", StringComparison.InvariantCulture) Or
                        key.Value.StartsWith("REG", StringComparison.InvariantCulture) Then
                    fullKeyErr(key, "Missing pipe symbol after ExcludeKey flag)")
                    Return
                End If
                fullKeyErr(key, "No valid exclude flag (FILE, PATH, or REG) found in ExcludeKey.")
        End Select
        fullKeyErr(key, "ExcludeKey has too many flags", key.Value.Split(CChar("|")).Length > 3)
    End Sub

    ''' <summary> Sorts a <c> keyList </c> alphabetically with winapp2.ini precedence applied to the key values </summary>
    ''' <param name="kl"> A <c> keyList </c> to be sorted alphabetically (with numbers having precedence) </param>
    ''' <param name="hadDuplicatesRemoved"> Indicates that keys have been removed from <c> <paramref name="kl"/> </c> </param>
    ''' Docs last updated: 2021-11-13 | Code last updated: 2021-11-13
    Private Sub sortKeys(ByRef kl As keyList, hadDuplicatesRemoved As Boolean)
        If Not lintAlpha.ShouldScan Or kl.KeyCount <= 1 Then Return
        Dim keyValues = kl.toStrLst(True)
        Dim sortedKeyValues = replaceAndSort(keyValues, "|", " \ \")
        ' Rewrite the alphabetized keys back into the keylist (also fixes numbering)
        Dim keysOutOfPlace = False
        findOutOfPlace(keyValues, sortedKeyValues, kl.KeyType, kl.lineNums, keysOutOfPlace)
        If (keysOutOfPlace Or hadDuplicatesRemoved) And (lintAlpha.fixFormat Or lintWrongNums.fixFormat Or lintExtraNums.fixFormat) Then
            kl.renumberKeys(sortedKeyValues)
        End If
    End Sub

    ''' <summary> Prints an error when data is received that does not match an expected value </summary>
    ''' <param name="linecount"> The line number on which the error was detected </param>
    ''' <param name="err"> A description of the error as it will be displayed to the user </param>
    ''' <param name="received"> The (erroneous) input data </param>
    ''' <param name="expected"> The expected data </param>
    ''' <param name="cond"> Indicates that the error condition is present <br/> Optional, Default: <c> True </c> </param>
    ''' Docs last updated: 2021-11-13 | Code last updated: 2021-11-13
    Private Sub inputMismatchErr(linecount As Integer, err As String, received As String, expected As String, Optional cond As Boolean = True)
        If cond Then customErr(linecount, err, {$"Expected: {expected}", $"Found: {received}"})
    End Sub

    ''' <summary> Prints an error followed by the [Full Name *] of the entry to which it belongs </summary>
    ''' <param name="cond"> Indicates that the error condition is present </param>
    ''' <param name="entry"> The <c> winapp2entry </c> containing an error </param>
    ''' <param name="errTxt"> A description of the error as it will be displayed to the user </param>
    ''' Docs last updated: 2021-11-13 | Code last updated: 2021-11-13
    Private Sub fullNameErr(cond As Boolean, entry As winapp2entry, errTxt As String)
        If cond Then customErr(entry.LineNum, errTxt, {$"Entry Name: {entry.FullName}"})
    End Sub

    ''' <summary> Prints an error whose output text contains an <c> iniKey </c> string, optionally correcting that value with one that is provided </summary>
    ''' <param name="key"> The <c> iniKey </c> containing an error </param>
    ''' <param name="err"> A description of the error as it will be displayed to the user </param>
    ''' <param name="cond"> Indicates that the error condition(s) are present (including any <c> lintRule.shouldScans </c>) <br/> Optional, Default: <c> True </c> </param>
    ''' <param name="repCond"> Indicates that the repair function should run <br/> Optional, Default: <c> False </c> </param>
    ''' <param name="newVal"> The corrected value with which to replace the incorrect correct value held by <c> <paramref name="repairVal"/> </c> <br/> Optional, Default: <c> "" </c> </param>
    ''' <param name="repairVal"> The incorrect value <br/> Optional, Default: <c> "" </c> </param>
    ''' Docs last updated: 2021-11-13 | Code last updated: 2024-04-22
    Private Sub fullKeyErr(key As iniKey, err As String, Optional cond As Boolean = True, Optional repCond As Boolean = False, Optional ByRef repairVal As String = "", Optional newVal As String = "")

        If Not cond Then Return

        customErr(key.LineNumber, err, {$"Key: {key.toString}"})
        fixStr(cond And repCond, repairVal, newVal)

    End Sub

    ''' <summary> Prints arbitrarily defined errors without a precondition </summary>
    ''' <param name="lineCount"> The line number on which the error was detected </param>
    ''' <param name="err"> A description of the error as it will be displayed to the user </param>
    ''' <param name="lines"> Any additional error information to be printed alongside the description </param>
    ''' Docs last updated: 2021-11-13 | Code last updated: 2021-11-13
    Private Sub customErr(lineCount As Integer, err As String, lines As String())
        gLog(err, ascend:=True)
        cwl($"Line: {lineCount} - Error: {err}")
        MostRecentLintLog += $"Line: {lineCount} - Error: {err}" & Environment.NewLine
        For Each errStr In lines
            cwl(errStr)
            gLog(errStr, indent:=True)
            MostRecentLintLog += errStr & Environment.NewLine
        Next
        gLog(descend:=True)
        cwl()
        MostRecentLintLog += Environment.NewLine
        ErrorsFound += 1
    End Sub

    ''' <summary> Replace a given string with a new value if the fix condition is met </summary>
    ''' <param name="param"> The condition under which the string should be replaced </param>
    ''' <param name="currentValue"> A pointer to the string to be replaced </param>
    ''' <param name="newValue"> The replacement value for <c> <paramref name="currentValue"/> </c> </param>
    ''' Docs last updated: 2021-11-13 | Code last updated: 2024-04-22
    Private Sub fixStr(param As Boolean, ByRef currentValue As String, newValue As String)

        If Not param Then Return

        gLog($"Changing '{currentValue}' to '{newValue}'", ascend:=True, descend:=True, indent:=True, buffr:=True)
        currentValue = newValue

    End Sub

End Module