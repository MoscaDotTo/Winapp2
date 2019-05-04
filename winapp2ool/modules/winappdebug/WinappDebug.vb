'    Copyright (C) 2018-2019 Robbie Ward
' 
'    This file is a part of Winapp2ool
' 
'    Winapp2ool is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    Winap2ool is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with Winapp2ool.  If not, see <http://www.gnu.org/licenses/>.
Option Strict On
Imports System.Text.RegularExpressions
''' <summary>
''' A program whose purpose is to observe, report, and attempt to repair errors in winapp2.ini
''' </summary>
Public Module WinappDebug
    ''' <summary> The winapp2.ini file that will be linted </summary>
    Public Property winappDebugFile1 As New iniFile(Environment.CurrentDirectory, "winapp2.ini")
    ''' <summary> Holds the save path for the linted file (overwrites the input file by default) </summary>
    Public Property winappDebugFile3 As New iniFile(Environment.CurrentDirectory, "winapp2.ini", "winapp2-debugged.ini")

    ''' <summary> Indicates whether some but not all repairs will run </summary>
    Public Property RepairSomeErrsFound As Boolean
    ''' <summary> Indicates whether scan settings have been changed </summary>
    Public Property ScanSettingsChanged As Boolean
    ''' <summary> Indicates whether module settings have been changed </summary>
    Public Property ModuleSettingsChanged As Boolean
    ''' <summary> Indicates whether or not changes should be saved back to disk (Default: False) </summary>
    Public Property SaveChanges As Boolean
    ''' <summary> Indicates whether or not winappdebug should attempt to repair errors it finds (Default: False)</summary>
    Public Property RepairErrsFound As Boolean = True
    ''' <summary> The number of errors found by the linter </summary>
    Public Property ErrorsFound As Integer

    ''' <summary>The list of all entry names found during a given run of the linter</summary>
    Private Property allEntryNames As New List(Of String)

    ''' <summary>Controls scan/repairs for CamelCasing issues</summary>
    Private Property lintCasing As New lintRule(True, True, "Casing", "improper CamelCasing", "fixing improper CamelCasing")
    ''' <summary>Controls scan/repairs for alphabetization issues</summary>
    Private Property lintAlpha As New lintRule(True, True, "Alphabetization", "improper alphabetization", "fixing improper alphabetization")
    ''' <summary>Controls scan/repairs for incorrectly numbered keys</summary>
    Private Property lintWrongNums As New lintRule(True, True, "Improper Numbering", "improper key numbering", "fixing improper key numbering")
    ''' <summary>Controls scan/repairs for parameters inside of FileKeys</summary>
    Private Property lintParams As New lintRule(True, True, "Parameters", "improper parameterization on FileKeys", "fixing improper parameterization on FileKeys")
    ''' <summary>Controls scan/repairs for flags in ExcludeKeys and FileKeys</summary>
    Private Property lintFlags As New lintRule(True, True, "Flags", "improper FileKey/ExcludeKey flag formatting", "fixing improper FileKey/ExcludeKey flag formatting")
    ''' <summary>Controls scan/repairs for improper slash usage</summary>
    Private Property lintSlashes As New lintRule(True, True, "Slashes", "improper use of slashes (\)", "fixing improper use of slashes (\)")
    ''' <summary>Controls scan/repairs for missing or True Default values</summary>
    Private Property lintDefaults As New lintRule(True, True, "Defaults", "Default=True or missing Default key", "enforcing Default=False")
    ''' <summary>Controls scan/repairs for duplicate values</summary>
    Private Property lintDupes As New lintRule(True, True, "Duplicates", "duplicate key values", "removing keys with duplicated values")
    '''<summary>Controls scan/repairs for keys with numbers they shouldn't have</summary>
    Private Property lintExtraNums As New lintRule(True, True, "Unneeded Numbering", "use of numbers where there should not be", "removing numbers used where they shouldn't be")
    ''' <summary>Controls scan/repairs for keys which should only occur once</summary>
    Private Property lintMulti As New lintRule(True, True, "Multiples", "multiples of key types that should only occur once in an entry", "removing unneeded multiples of key types that should occur only once")
    ''' <summary>Controls scan/repairs for keys with invlaid values</summary>
    Private Property lintInvalid As New lintRule(True, True, "Invalid Values", "invalid key values", "fixing certain types of invalid key values")
    ''' <summary>Controls scan/repairs for winapp2.ini syntax errors</summary>
    Private Property lintSyntax As New lintRule(True, True, "Syntax Errors", "some entries whose configuration will not run in CCleaner", "attempting to fix certain types of syntax errors")
    ''' <summary>Controls scan/repairs for invalid file or regsitry paths</summary>
    Private Property lintPathValidity As New lintRule(True, True, "Path Validity", "invalid filesystem or registry locations", "attempting to repair some basic invalid parameters in paths")
    ''' <summary>Controls scan/repairs for keys that can be merged into eachother (FileKeys only currently)</summary>
    Private Property lintOpti As New lintRule(False, False, "Optimizations", "situations where keys can be merged (experimental)", "automatic merging of keys (experimental)")
    Private currentLintRules As lintRule() = {lintCasing, lintAlpha, lintWrongNums, lintParams, lintFlags, lintSlashes,
            lintDefaults, lintDupes, lintExtraNums, lintMulti, lintInvalid, lintSyntax, lintPathValidity, lintOpti}

    '''<summary>Regex to detect long form registry paths</summary>
    Private Property longReg As New Regex("HKEY_(C(URRENT_(USER$|CONFIG$)|LASSES_ROOT$)|LOCAL_MACHINE$|USERS$)")
    '''<summary>Regex to detect short form registry paths</summary>
    Private Property shortReg As New Regex("HK(C(C$|R$|U$)|LM$|U$)")
    ''' <summary>Regex to detect valid LangSecRef numbers</summary>
    Private Property secRefNums As New Regex("30(05|2([0-9])|3(0|1))")
    '''<summary>Regex to detect valid drive letter parameters</summary>
    Private Property driveLtrs As New Regex("[a-zA-z]:")
    '''<summary>Regex to detect potential %EnvironmentVariables%</summary>
    Private Property envVarRegex As New Regex("%[A-Za-z0-9]*%")

    ''' <summary> The current rules for scans and repairs </summary>
    Public Property Rules As lintRule()
        Get
            Return currentLintRules
        End Get
        Set
            currentLintRules = Value
        End Set
    End Property

    ''' <summary> Handles the commandline args for WinappDebug </summary>
    ''' WinappDebug specific command line args
    ''' -c          : enable autocorrect
    Public Sub handleCmdLine()
        initDefaultSettings()
        invertSettingAndRemoveArg(SaveChanges, "-c")
        getFileAndDirParams(winappDebugFile1, New iniFile, winappDebugFile3)
        If Not cmdargs.Contains("UNIT_TESTING_HALT") Then initDebug()
    End Sub

    ''' <summary> Restore the default state of the module's parameters </summary>
    Private Sub initDefaultSettings()
        winappDebugFile1.resetParams()
        winappDebugFile3.resetParams()
        ModuleSettingsChanged = False
        RepairErrsFound = True
        SaveChanges = False
        resetScanSettings()
    End Sub

    ''' <summary> Prints the main menu to the user </summary>
    Public Sub printMenu()
        printMenuTop({"Scan winapp2.ini For syntax and style errors, And attempt to repair them."})
        print(1, "Run (Default)", "Run the debugger")
        print(1, "File Chooser (winapp2.ini)", "Choose a new winapp2.ini name or path", leadingBlank:=True, trailingBlank:=True)
        print(5, "Toggle Saving", "saving the file after correcting errors", enStrCond:=SaveChanges)
        print(1, "File Chooser (save)", "Save a copy of changes made instead of overwriting winapp2.ini", SaveChanges, trailingBlank:=True)
        print(1, "Toggle Scan Settings", "Enable or disable individual scan and correction routines", leadingBlank:=Not SaveChanges, trailingBlank:=True)
        print(0, $"Current winapp2.ini:  {replDir(winappDebugFile1.path)}", closeMenu:=Not SaveChanges And Not ModuleSettingsChanged)
        print(0, $"Current save target:  {replDir(winappDebugFile3.path)}", cond:=SaveChanges, closeMenu:=Not ModuleSettingsChanged)
        print(2, "WinappDebug", cond:=ModuleSettingsChanged, closeMenu:=True)
    End Sub

    ''' <summary> Handles the user's input from the menu </summary>
    ''' <param name="input">The String containing the user's input</param>
    Public Sub handleUserInput(input As String)
        Select Case True
            Case input = "0"
                exitModule()
            Case input = "1" Or input = ""
                initDebug()
            Case input = "2"
                changeFileParams(winappDebugFile1, ModuleSettingsChanged)
            Case input = "3"
                toggleSettingParam(SaveChanges, "Saving", ModuleSettingsChanged)
            Case input = "4" And SaveChanges
                changeFileParams(winappDebugFile3, ModuleSettingsChanged)
            Case (input = "4" And Not SaveChanges) Or (input = "5" And SaveChanges)
                initModule("Scan Settings", AddressOf advSettings.printMenu, AddressOf advSettings.handleUserInput)
                Console.WindowHeight = 30
            Case ModuleSettingsChanged And ((input = "5" And Not SaveChanges) Or (input = "6" And SaveChanges))
                resetModuleSettings("WinappDebug", AddressOf initDefaultSettings)
            Case Else
                setHeaderText(invInpStr, True)
        End Select
    End Sub

    ''' <summary> Validates and debugs the ini file, informs the user upon completion </summary>
    Private Sub initDebug()
        winappDebugFile1.validate()
        If pendingExit() Then Exit Sub
        Dim wa2 As New winapp2file(winappDebugFile1)
        clrConsole()
        print(0, tmenu("Beginning analysis of winapp2.ini"), closeMenu:=True)
        cwl()
        debug(wa2)
        setHeaderText("Debug Complete")
        print(0, tmenu("Completed analysis of winapp2.ini"))
        print(0, menuStr03)
        print(0, $"{ErrorsFound} possible errors were detected.")
        print(0, $"Number of entries: {winappDebugFile1.Sections.Count}", trailingBlank:=True)
        rewriteChanges(wa2)
        print(0, anyKeyStr, closeMenu:=True)
        If Not SuppressOutput Then Console.ReadKey()
    End Sub

    ''' <summary> Performs syntax and format checking on a winapp2.ini format ini file </summary>
    ''' <param name="fileToBeDebugged">A winapp2file object to be processed</param>
    Public Sub debug(ByRef fileToBeDebugged As winapp2file)
        ErrorsFound = 0
        allEntryNames = New List(Of String)
        For Each entryList In fileToBeDebugged.winapp2entries
            If entryList.Count = 0 Then Continue For
            entryList.ForEach(Sub(entry) processEntry(entry))
        Next
        fileToBeDebugged.rebuildToIniFiles()
        alphabetizeEntries(fileToBeDebugged)
    End Sub

    ''' <summary>Alphabetizes all the entries in a winapp2.ini file and observes any that were out of place</summary>
    ''' <param name="winapp">The object containing the winapp2.ini being operated on</param>
    Private Sub alphabetizeEntries(ByRef winapp As winapp2file)
        For Each innerFile In winapp.entrySections
            Dim unsortedEntryList = innerFile.namesToStrList
            Dim sortedEntryList = sortEntryNames(innerFile)
            findOutOfPlace(unsortedEntryList, sortedEntryList, "Entry", innerFile.getLineNumsFromSections)
            If lintAlpha.fixFormat Then innerFile.sortSections(sortedEntryList)
        Next
    End Sub

    ''' <summary> Overwrites the file on disk with any changes we've made if we are saving them </summary>
    ''' <param name="winapp2file">The object representing winapp2.ini</param>
    Private Sub rewriteChanges(ByRef winapp2file As winapp2file)
        If SaveChanges Then
            print(0, "Saving changes, do not close winapp2ool or data loss may occur...", leadingBlank:=True)
            winappDebugFile3.overwriteToFile(winapp2file.winapp2string)
            print(0, "Finished saving changes.", trailingBlank:=True)
        End If
    End Sub

    ''' <summary> Assess a list and its sorted state to observe changes in neighboring strings</summary>
    ''' eg. changes made to string ordering through sorting
    ''' <param name="someList">A list of strings</param>
    ''' <param name="sortedList">The sorted state of someList</param>
    ''' <param name="findType">The type of neighbor checking (keyType for iniKey values)</param>
    ''' <param name="LineCountList">A list containing the line counts of the Strings in someList</param>
    ''' <param name="oopBool">Optional Boolean that reports whether or not there are any out of place entries in the list</param>
    Private Sub findOutOfPlace(ByRef someList As strList, ByRef sortedList As strList, ByVal findType As String, ByRef LineCountList As List(Of Integer), Optional ByRef oopBool As Boolean = False)
        ' Only try to find out of place keys when there's more than one
        If someList.count > 1 Then
            Dim misplacedEntries As New strList
            ' Learn the neighbors of each string in each respective list
            Dim initialNeighbors = someList.getNeighborList
            Dim sortedNeighbors = sortedList.getNeighborList
            ' Make sure at least one of the neighbors of each string are the same in both the sorted and unsorted state, otherwise the string has moved 
            For i = 0 To someList.count - 1
                Dim sind = sortedList.items.IndexOf(someList.items(i))
                misplacedEntries.add(someList.items(i), Not (initialNeighbors(i).Key = sortedNeighbors(sind).Key And initialNeighbors(i).Value = sortedNeighbors(sind).Value))
            Next
            ' Report any misplaced entries back to the user
            For Each entry In misplacedEntries.items
                Dim recInd = someList.indexOf(entry)
                Dim sortInd = sortedList.indexOf(entry)
                Dim curLine = LineCountList(recInd)
                Dim sortLine = LineCountList(sortInd)
                If (recInd = sortInd Or curLine = sortLine) Then Continue For
                entry = If(findType = "Entry", entry, $"{findType & recInd + 1}={entry}")
                If Not oopBool Then oopBool = True
                customErr(LineCountList(recInd), $"{findType} alphabetization", {$"{entry} appears to be out of place", $"Current line: {curLine}", $"Expected line: {sortLine}"})
            Next
        End If
    End Sub

    ''' <summary> Audits a keyList of winapp2.ini format iniKeys for errors, alerting the user and correcting where possible.</summary>
    ''' <param name="kl">A keylist to audit</param>
    ''' <param name="processKey">The function that audits the keys of the keyType provided in the keyList</param>
    ''' <param name="hasF">Optional boolean for the ExcludeKey case</param>
    ''' <param name="hasR">Optional boolean for the ExcludeKey case</param>
    Private Sub processKeyList(ByRef kl As keyList, processKey As Func(Of iniKey, iniKey), Optional ByRef hasF As Boolean = False, Optional ByRef hasR As Boolean = False)
        If kl.keyCount = 0 Then Exit Sub
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
                    cFormat(key, curNum, curStrings, dupes, kl.keyCount = 1)
                Case "RegKey"
                    chkPathFormatValidity(key, True)
                    cFormat(key, curNum, curStrings, dupes)
                Case "Warning", "DetectOS", "SpecialDetect", "LangSecRef", "Section", "Default"
                    ' No keys of these types should occur more than once per entry
                    If curNum > 1 And lintMulti.ShouldScan Then
                        fullKeyErr(key, $"Multiple {key.KeyType} detected.")
                        dupes.add(key, lintMulti.fixFormat)
                    End If
                    cFormat(key, curNum, curStrings, dupes, True)
                    ' Scan for invalid values in LangSecRef and SpecialDetect
                    If key.typeIs("SpecialDetect") Then chkCasing(key, {"DET_CHROME", "DET_MOZILLA", "DET_THUNDERBIRD", "DET_OPERA"}, key.Value, 1)
                    fullKeyErr(key, "LangSecRef holds an invalid value.", lintInvalid.ShouldScan And key.typeIs("LangSecRef") And Not secRefNums.IsMatch(key.Value))
                    ' Enforce that Default=False
                    fullKeyErr(key, "All entries should be disabled by default (Default=False).", lintDefaults.ShouldScan And Not key.vIs("False") And key.typeIs("Default"), lintDefaults.fixFormat, key.Value, "False")
                Case Else
                    cFormat(key, curNum, curStrings, dupes)
            End Select
            ' Any further changes to the key are handled by the given function
            key = processKey(key)
        Next
        ' Remove any duplicates and sort the keys
        kl.remove(dupes.Keys)
        sortKeys(kl, dupes.keyCount > 0)
        ' Run optimization checks on FileKey lists only 
        If kl.typeIs("FileKey") Then cOptimization(kl)
    End Sub

    ''' <summary> Does some basic formatting checks that apply to most winapp2.ini format iniKey objects </summary>
    ''' <param name="key">The current iniKey object being processed</param>
    ''' <param name="keyNumber">The current expected key number</param>
    ''' <param name="keyValues">The current list of observed inikey values</param>
    ''' <param name="dupeList">A tracking list of detected duplicate iniKeys</param>
    Private Sub cFormat(ByRef key As iniKey, ByRef keyNumber As Integer, ByRef keyValues As strList, ByRef dupeList As keyList, Optional noNumbers As Boolean = False)
        ' Check for duplicates
        If keyValues.contains(key.Value, True) Then
            Dim dupeKeyStr = $"{key.KeyType}{If(Not noNumbers, (keyValues.items.IndexOf(key.Value) + 1).ToString, "")}={key.Value}"
            If lintDupes.ShouldScan Then customErr(key.LineNumber, "Duplicate key value found", {$"Key:            {key.toString}", $"Duplicates:     {dupeKeyStr}"})
            dupeList.add(key, lintDupes.fixFormat)
        Else
            keyValues.add(key.Value)
        End If
        ' Check for both types of numbering errors (incorrect and unneeded) 
        Dim hasNumberingError = If(noNumbers, Not key.nameIs(key.KeyType), Not key.nameIs(key.KeyType & keyNumber))
        Dim numberingErrStr = If(noNumbers, "Detected unnecessary numbering.", $"{key.KeyType} entry is incorrectly numbered.")
        Dim fixedStr = If(noNumbers, key.KeyType, key.KeyType & keyNumber)
        inputMismatchErr(key.LineNumber, numberingErrStr, key.Name, fixedStr, If(noNumbers, lintExtraNums.ShouldScan, lintWrongNums.ShouldScan) And hasNumberingError)
        fixStr(If(noNumbers, lintExtraNums.fixFormat, lintWrongNums.fixFormat) And hasNumberingError, key.Name, fixedStr)
        ' Scan for and fix any use of incorrect slashes (except in Warning keys) or trailing semicolons
        fullKeyErr(key, "Forward slash (/) detected in lieu of backslash (\).", Not key.typeIs("Warning") And lintSlashes.ShouldScan And key.vHas(CChar("/")), lintSlashes.fixFormat, key.Value, key.Value.Replace(CChar("/"), CChar("\")))
        fullKeyErr(key, "Trailing semicolon (;).", key.toString.Last = CChar(";") And lintParams.ShouldScan, lintParams.fixFormat, key.Value, key.Value.TrimEnd(CChar(";")))
        ' Do some formatting checks for environment variables if needed
        If {"FileKey", "ExcludeKey", "DetectFile"}.Contains(key.KeyType) Then cEnVar(key)
        keyNumber += 1
    End Sub

    ''' <summary>Assess the formatting of any %EnvironmentVariables% in a given iniKey</summary>
    ''' <param name="key">An iniKey object to be processed</param>
    Private Sub cEnVar(key As iniKey)
        ' Valid Environmental Variables for winapp2.ini
        Dim enVars = {"AllUsersProfile", "AppData", "CommonAppData", "CommonProgramFiles",
            "Documents", "HomeDrive", "LocalAppData", "LocalLowAppData", "Music", "Pictures", "ProgramData", "ProgramFiles", "Public",
            "RootDir", "SystemDrive", "SystemRoot", "Temp", "Tmp", "UserName", "UserProfile", "Video", "WinDir"}
        fullKeyErr(key, "%EnvironmentVariables% must be surrounded on both sides by a single '%' character.", key.vHas("%") And envVarRegex.Matches(key.Value).Count = 0)
        For Each m As Match In envVarRegex.Matches(key.Value)
            Dim strippedText = m.ToString.Trim(CChar("%"))
            chkCasing(key, enVars, strippedText, 1)
        Next
    End Sub

    ''' <summary>Does basic syntax and formatting audits that apply across all keys, returns false iff a key is malformed</summary>
    ''' <param name="key">an iniKey object to be checked for errors</param>
    Private Function cValidity(key As iniKey) As Boolean
        Dim validCmds As String() = {"Default", "Detect", "DetectFile", "DetectOS", "ExcludeKey",
                           "FileKey", "LangSecRef", "RegKey", "Section", "SpecialDetect", "Warning"}
        If key.typeIs("DeleteMe") Then Return False
        ' Check for leading or trailing whitespace, do this always as spaces in the name interfere with proper keyType identification
        If key.Name.StartsWith(" ") Or key.Name.EndsWith(" ") Or key.Value.StartsWith(" ") Or key.Value.EndsWith(" ") Then
            fullKeyErr(key, "Detected unwanted whitespace in iniKey", True)
            fixStr(True, key.Value, key.Value.Trim(CChar(" ")))
            fixStr(True, key.Name, key.Name.Trim(CChar(" ")))
            fixStr(True, key.KeyType, key.KeyType.Trim(CChar(" ")))
        End If
        ' Make sure the keyType is valid
        chkCasing(key, validCmds, key.KeyType, 0)
        Return True
    End Function

    ''' <summary>Checks either the value or the keyType of an iniKey object against a given array of expected cased texts</summary>
    ''' <param name="key">An iniKey object to check casing in</param>
    ''' <param name="casedArray">The array of expected cased texts</param>
    ''' <param name="strToChk">A reference to the given string</param>
    ''' <param name="chkType">0 if checking keyTypes, 1 if checking Values</param>
    Private Sub chkCasing(ByRef key As iniKey, casedArray As String(), ByRef strToChk As String, chkType As Integer)
        ' Get the properly cased string
        Dim casedString = getCasedString(casedArray, strToChk)
        ' Determine if there's a casing error
        Dim hasCasingErr = Not casedString.Equals(strToChk) And casedArray.Contains(casedString)
        Dim replacementText = ""
        Select Case chkType
            Case 0
                replacementText = key.Name.Replace(key.KeyType, casedString)
            Case 1
                replacementText = key.Value.Replace(key.Value, casedString)
        End Select
        fullKeyErr(key, $"{casedString} has a casing error.", hasCasingErr And lintCasing.ShouldScan, lintCasing.fixFormat, strToChk, replacementText)
        fullKeyErr(key, $"Invalid data provided: {strToChk}{Environment.NewLine}Valid data: {arrayToStr(casedArray)}", Not casedArray.Contains(casedString) And lintInvalid.ShouldScan)
    End Sub

    ''' <summary>Returns the contents of the array as a single comma delimited String</summary>
    ''' <param name="given">A given array of Strings</param>
    Private Function arrayToStr(given As String()) As String
        Dim out = ""
        For i = 0 To given.Count - 2
            out += given(i) & ", "
        Next
        out += given.Last
        Return out
    End Function

    ''' <summary>Processes a FileKey format winapp2.ini iniKey object and checks it for errors, correcting where possible</summary>
    ''' <param name="key">A winap2.ini FileKey format iniKey object</param>
    Public Function pFileKey(key As iniKey) As iniKey
        ' Pipe symbol checks
        Dim iteratorCheckerList() As String = Split(key.Value, "|")
        fullKeyErr(key, "Missing pipe (|) in FileKey.", Not key.vHas("|"))
        ' Captures any incident of semi colons coming before the first pipe symbol
        fullKeyErr(key, "Semicolon (;) found before pipe (|).", key.vHas(";") And key.Value.IndexOf(";") < key.Value.IndexOf("|"))
        ' Check for incorrect spellings of RECURSE or REMOVESELF
        If iteratorCheckerList.Length > 2 Then fullKeyErr(key, "RECURSE or REMOVESELF is incorrectly spelled, or there are too many pipe (|) symbols.", Not iteratorCheckerList(2).Contains("RECURSE") And Not iteratorCheckerList(2).Contains("REMOVESELF"))
        ' Check for missing pipe symbol on recurse and removeself, fix them if detected
        Dim flags As New List(Of String) From {"RECURSE", "REMOVESELF"}
        flags.ForEach(Sub(flagStr) fullKeyErr(key, $"Missing pipe (|) before {flagStr}.", lintFlags.ShouldScan And key.vHas(flagStr) And Not key.vHas($"|{flagStr}"), lintFlags.fixFormat, key.Value, key.Value.Replace(flagStr, $"|{flagStr}")))
        ' Make sure VirtualStore folders point to the correct place
        inputMismatchErr(key.LineNumber, "Incorrect VirtualStore location.", key.Value, "%LocalAppData%\VirtualStore\Program Files*\", key.vHas("\virtualStore\p", True) And Not key.vHasAny({"programdata", "program files*", "program*"}, True))
        ' Backslash checks, fix if detected
        fullKeyErr(key, "Backslash (\) found before pipe (|).", lintSlashes.ShouldScan And key.vHas("%\|"), lintSlashes.fixFormat, key.Value, key.Value.Replace("%\|", "%|"))
        fullKeyErr(key, "Missing backslash (\) after %EnvironmentVariable%.", key.vHas("%") And Not key.vHasAny({"%|", "%\"}))
        ' Get the parameters given to the file key and sort them 
        Dim keyParams As New winapp2KeyParameters(key)
        Dim argsStrings As New List(Of String)
        Dim dupeArgs As New List(Of String)
        ' Check for duplicate args
        For Each arg In keyParams.ArgsList
            If chkDupes(argsStrings, arg) And lintParams.ShouldScan Then
                customErr(key.LineNumber, "Duplicate FileKey parameter found", {$"Command: {arg}"})
                If lintParams.fixFormat Then dupeArgs.Add(arg)
            End If
        Next
        ' Remove any duplicate arguments from the key parameters and reconstruct keys we've modified above
        If lintParams.fixFormat Then
            dupeArgs.ForEach(Sub(arg) keyParams.ArgsList.Remove(arg))
            keyParams.reconstructKey(key)
        End If
        Return key
    End Function

    ''' <summary>Checks the validity of the keys in an entry and removes any that are too problematic to continue with</summary>
    ''' <param name="entry">A winapp2entry object to be audited</param>
    Private Sub validateKeys(ByRef entry As winapp2entry)
        For Each lst In entry.KeyListList
            Dim brokenKeys As New keyList
            lst.Keys.ForEach(Sub(key) brokenKeys.add(key, Not cValidity(key)))
            lst.remove(brokenKeys.Keys)
            entry.ErrorKeys.remove(brokenKeys.Keys)
        Next
        ' Attempt to assign keys that had errors to their intended lists
        For Each key In entry.ErrorKeys.Keys
            For Each lst In entry.KeyListList
                If lst.KeyType = "Error" Then Continue For
                lst.add(key, key.typeIs(lst.KeyType))
            Next
        Next
    End Sub

    ''' <summary>Processes a winapp2entry object (generated from a winapp2.ini format iniKey object) for errors</summary>
    ''' <param name="entry">The winapp2entry object to be processed</param>
    Private Sub processEntry(ByRef entry As winapp2entry)
        Dim entryLinesList As New List(Of Integer)
        Dim hasFileExcludes = False
        Dim hasRegExcludes = False
        ' Check for duplicate names that are differently cased 
        fullNameErr(chkDupes(allEntryNames, entry.Name), entry, "Duplicate entry name detected")
        ' Check that the entry is named properly 
        fullNameErr(Not entry.Name.EndsWith(" *"), entry, "All entries must End In ' *'")
        ' Confirm the validity of keys and remove any broken ones before continuing
        validateKeys(entry)
        ' Process the entry's keylists in winapp2.ini order (ignore the last list because it has only errors)
        For i = 0 To entry.KeyListList.Count - 2
            Dim lst = entry.KeyListList(i)
            Select Case lst.KeyType
                Case "DetectFile"
                    processKeyList(lst, AddressOf pDetectFile)
                Case "FileKey"
                    processKeyList(lst, AddressOf pFileKey)
                Case Else
                    processKeyList(lst, AddressOf voidDelegate, hasFileExcludes, hasRegExcludes)
            End Select
        Next
        ' Make sure we only have LangSecRef if we have LangSecRef at all
        fullNameErr(entry.LangSecRef.keyCount <> 0 And entry.SectionKey.keyCount <> 0 And lintSyntax.ShouldScan, entry, "Section key found alongside LangSecRef key, only one is required.")
        ' Make sure we have at least 1 valid detect key and at least one valid cleaning key
        fullNameErr(entry.DetectOS.keyCount + entry.Detects.keyCount + entry.SpecialDetect.keyCount + entry.DetectFiles.keyCount = 0, entry, "Entry has no valid detection keys (Detect, DetectFile, DetectOS, SpecialDetect)")
        fullNameErr(entry.FileKeys.keyCount + entry.RegKeys.keyCount = 0 And lintSyntax.ShouldScan, entry, "Entry has no valid FileKeys or RegKeys")
        ' If we don't have FileKeys or RegKeys, we shouldn't have ExcludeKeys.
        fullNameErr(entry.ExcludeKeys.keyCount > 0 And entry.FileKeys.keyCount + entry.RegKeys.keyCount = 0, entry, "Entry has ExcludeKeys but no valid FileKeys or RegKeys")
        ' Make sure that if we have excludes, we also have corresponding file/reg keys
        fullNameErr(entry.FileKeys.keyCount = 0 And hasFileExcludes, entry, "ExcludeKeys targeting filesystem locations found without any corresponding FileKeys")
        fullNameErr(entry.RegKeys.keyCount = 0 And hasRegExcludes, entry, "ExcludeKeys targeting registry locations found without any corresponding RegKeys")
        ' Make sure we have a Default key.
        fullNameErr(entry.DefaultKey.keyCount = 0 And lintDefaults.ShouldScan, entry, "Entry is missing a Default key")
        entry.DefaultKey.add(New iniKey("Default=False"), lintDefaults.fixFormat And entry.DefaultKey.keyCount = 0)
    End Sub

    ''' <summary> This method does nothing by design </summary>
    ''' <param name="key">An iniKey object to do nothing with</param>
    Private Function voidDelegate(key As iniKey) As iniKey
        Return key
    End Function

    ''' <summary>Processes a DetectFile format winapp2.ini iniKey objects and checks it for errors, correcting where possible </summary>
    ''' <param name="key">A winapp2.ini DetectFile format iniKey</param>
    Private Function pDetectFile(key As iniKey) As iniKey
        ' Trailing Backslashes & nested wildcards
        fullKeyErr(key, "Trailing backslash (\) found in DetectFile", lintSlashes.ShouldScan _
            And key.Value.Last = CChar("\"), lintSlashes.fixFormat, key.Value, key.Value.TrimEnd(CChar("\")))
        If key.vHas("*") Then
            Dim splitDir = key.Value.Split(CChar("\"))
            For i = 0 To splitDir.Count - 1
                fullKeyErr(key, "Nested wildcard found in DetectFile", splitDir(i).Contains("*") And i <> splitDir.Count - 1)
            Next
        End If
        ' Make sure that DetectFile paths point to a filesystem location
        chkPathFormatValidity(key, False)
        Return key
    End Function

    ''' <summary> Audits the syntax of file system and registry paths </summary>
    ''' <param name="key">An iniKey to be audited</param>
    ''' <param name="isRegistry">Specifies whether the given key holds a registry path</param>
    Private Sub chkPathFormatValidity(key As iniKey, isRegistry As Boolean)
        If Not lintPathValidity.ShouldScan Then Exit Sub
        ' Remove the flags from ExcludeKeys if we have them before getting the first directory portion
        Dim rootStr = If(key.KeyType <> "ExcludeKey", getFirstDir(key.Value), getFirstDir(pathFromExcludeKey(key)))
        ' Ensure that registry paths have a valid hive and file paths have either a variable or a drive letter
        fullKeyErr(key, "Invalid registry path detected.", isRegistry And Not longReg.IsMatch(rootStr) And Not shortReg.IsMatch(rootStr))
        fullKeyErr(key, "Invalid file system path detected.", Not isRegistry And Not driveLtrs.IsMatch(rootStr) And Not rootStr.StartsWith("%"))
    End Sub

    ''' <summary>
    ''' Checks whether the current value appears in the given list of strings (case insensitive). Returns true if there is a duplicate,
    ''' otherwise, adds the current value to the list and returns false.
    ''' </summary>
    ''' <param name="valueStrings">A list of strings holding observed values</param>
    ''' <param name="currentValue">The current value to be audited</param>
    Private Function chkDupes(ByRef valueStrings As List(Of String), currentValue As String) As Boolean
        For Each value In valueStrings
            If currentValue.Equals(value, StringComparison.InvariantCultureIgnoreCase) Then Return True
        Next
        valueStrings.Add(currentValue)
        Return False
    End Function

    ''' <summary> Sorts a keylist alphabetically with winapp2.ini precedence applied to the key values </summary>
    ''' <param name="kl">The keylist to be sorted</param>
    ''' <param name="hadDuplicatesRemoved">The boolean indicating whether keys have been removed from this list</param>
    Private Sub sortKeys(ByRef kl As keyList, hadDuplicatesRemoved As Boolean)
        If Not lintAlpha.ShouldScan Or kl.keyCount <= 1 Then Exit Sub
        Dim keyValues = kl.toStrLst(True)
        Dim sortedKeyValues = replaceAndSort(keyValues, "|", " \ \")
        ' Rewrite the alphabetized keys back into the keylist (also fixes numbering)
        Dim keysOutOfPlace = False
        findOutOfPlace(keyValues, sortedKeyValues, kl.KeyType, kl.lineNums, keysOutOfPlace)
        If (keysOutOfPlace Or hadDuplicatesRemoved) And (lintAlpha.fixFormat Or lintWrongNums.fixFormat Or lintExtraNums.fixFormat) Then
            kl.renumberKeys(sortedKeyValues)
        End If
    End Sub

    ''' <summary>Processes a list of ExcludeKey format winapp2.ini iniKey objects and checks them for errors, correcting where possible</summary>
    ''' <param name="key">A winapp2.ini ExcludeKey format iniKey object</param>
    ''' <param name="hasF">Indicates whether the given list excludes filesystem locations</param>
    ''' <param name="hasR">Indicates whether the given list excludes registry locations</param>
    Private Sub pExcludeKey(ByRef key As iniKey, ByRef hasF As Boolean, ByRef hasR As Boolean)
        Select Case True
            Case key.vHasAny({"FILE|", "PATH|"})
                hasF = True
                chkPathFormatValidity(key, False)
                fullKeyErr(key, "Missing backslash (\) before pipe (|) in ExcludeKey.", (key.vHas("|") And Not key.vHas("\|")))
            Case key.vHas("REG|")
                hasR = True
                chkPathFormatValidity(key, True)
            Case Else
                If key.Value.StartsWith("FILE") Or key.Value.StartsWith("PATH") Or key.Value.StartsWith("REG") Then
                    fullKeyErr(key, "Missing pipe symbol after ExcludeKey flag)")
                    Exit Sub
                End If
                fullKeyErr(key, "No valid exclude flag (FILE, PATH, or REG) found in ExcludeKey.")
        End Select
    End Sub

    ''' <summary>Prints an error when data is received that does not match an expected value.</summary>
    ''' <param name="linecount">The line number of the error</param>
    ''' <param name="err">The string containing the output error text</param>
    ''' <param name="received">The (erroneous) input data</param>
    ''' <param name="expected">The expected data</param>
    ''' <param name="cond">Optional condition under which to display the error</param>
    Private Sub inputMismatchErr(linecount As Integer, err As String, received As String, expected As String, Optional cond As Boolean = True)
        If cond Then customErr(linecount, err, {$"Expected: {expected}", $"Found: {received}"})
    End Sub

    ''' <summary>Prints an error followed by the [Full Name *] of the entry to which it belongs</summary>
    ''' <param name="cond">The condition under which to print</param>
    ''' <param name="entry">The entry containing an error</param>
    ''' <param name="errTxt">The String containing the text to be printed to the user</param>
    Private Sub fullNameErr(cond As Boolean, entry As winapp2entry, errTxt As String)
        If cond Then customErr(entry.LineNum, errTxt, {$"Entry Name: {entry.FullName}"})
    End Sub

    ''' <summary>Prints an error whose output text contains an iniKey string</summary>
    ''' <param name="key">The inikey to be printed</param>
    ''' <param name="err">The string containing the output error text</param>
    ''' <param name="cond">Optional condition under which the err should be printed</param>
    ''' <param name="repCond">Optional condition under which to repair the given key</param>
    ''' <param name="newVal">The value to replace the error value</param>
    ''' <param name="repairVal">The error value</param>
    Private Sub fullKeyErr(key As iniKey, err As String, Optional cond As Boolean = True, Optional repCond As Boolean = False, Optional ByRef repairVal As String = "", Optional newVal As String = "")
        If cond Then customErr(key.LineNumber, err, {$"Key: {key.toString}"})
        fixStr(cond And repCond, repairVal, newVal)
    End Sub

    ''' <summary>Prints arbitrarily defined errors without a precondition</summary>
    ''' <param name="lineCount"></param>
    ''' <param name="err"></param>
    ''' <param name="lines"></param>
    Private Sub customErr(lineCount As Integer, err As String, lines As String())
        cwl($"Line: {lineCount} - Error: {err}")
        lines.ToList.ForEach(Sub(errStr As String) cwl(errStr))
        cwl()
        ErrorsFound += 1
    End Sub

    ''' <summary>Replace a given string with a new value if the fix condition is met.</summary>
    ''' <param name="param">The condition under which the string should be replaced</param>
    ''' <param name="currentValue">A reference to the string to be replaced</param>
    ''' <param name="newValue">The replacement value for the given string</param>
    Private Sub fixStr(param As Boolean, ByRef currentValue As String, ByRef newValue As String)
        If param Then currentValue = newValue
    End Sub
End Module