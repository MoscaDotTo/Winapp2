# WinappDebug

WinappDebug is a basic [linter](https://www.wikiwand.com/en/Lint_%28software%29) for winapp2.ini. It performs static analysis on winapp2.ini to ensure and enforce correctness of style and syntax across a wide variety of configurable categories. Optionally and additionally, it will save the linted file back to disk.

## Menu Options

Option|Effect|Notes
:-|:-|:-
Exit|Returns you to the winapp2ool menu|
Run|Runs the tool using your current settings|Default option
File Chooser (winapp2.ini)|Opens the interface to select a new local file for winapp2.ini|Default file path is winapp2.ini in the current working directory
Toggle Saving|Toggles the saving of corrected errors back to disk|Disabled by default
File Chooser (save)|Opens the interface to select a different local path to save changes too|Overwrites the input file by default. <br/> Only shown when saving is enabled.
Toggle Scan Settings|Opens the interface to toggle individual scans and/or repairs|See "Scan Settings" section below
Reset Settings|Restores settings to their default states, undoing any changes the user may have made|Only shown if a setting has or may have been changed (ie. a File Chooser was opened)
Log Viewer|Shows a summary of the last analysis|Only shown if errors are detected

## Scan Settings Menu Options

Option|Type of Scan/Repair controlled
:-|:-
Casing|Instances of improper CamelCasing
Alphabetization|Instances of improper alphabetization
Improper Numbering|Instances of numbered keys having incorrect values
Parameters|Instances of FileKeys having errors in their parameters
Flags|Instances of incorrect flag masks in FileKeys and ExcludeKeys
Slashes|Instances of improper uses of slashes
Defaults|Instances of keys lacking a "Default=False" key
Duplicates|Instances of duplicate key names or values in a single entry
Unneeded Numbering|Instances of keys having numbers that they do not need or should not have
Multiples|Instances of singleton keys occurring more than once
Invalid Values|Instances of invalid values for certain keyTypes
Syntax Errors|Instances of entries whose configuration may not run in CCleaner
Path Validity|Instances of invalid file system or registry locations
Semicolons|Instances of improperly used semicolons (;)
Optimizations|Instances where FileKeys may possibly be merged **(experimental, disabled by default)**

### Detected Errors

**Bold** items are correctable
*italicised* items are covered by tests

#### General

* ***Duplicate key names or values***
* ***Incorrect/Unnecessary key numbering***
* ***Incorrect key alphabetization***
* ***Forward slash (/) use where there should be a backslash (\\)***
* ***Multiple backslashes (\\\\)***
* ***Trailing Semicolons (;)***
* **Invalid %EnvironmentVariable% CamelCasing**
* Invalid %EnvironmentVariable% formatting
* Unsupported %EnvironmentVariable% target
* **Leading or trailing whitespace on key names or values**
* **Invalid CamelCasing on winapp2.ini commands**
* Invalid winapp2.ini commands
* No valid Detect / DetectFile / DetectOS / SpecialDetect provided
* No valid FileKeys or RegKeys provided
* ExcludeKeys specified in the absence of corresponding FileKeys or RegKeys
* **No Default key provided**
* **Missing = between iniKey Name and Value**
  * ie FileKey1%WinDir%\tmp|\*.\*

#### DetectOS

* More than one DetectOS key provided

#### LangSecRef / Section

* More than one (of either) key provided
* Invalid LangSecRef value provided

#### SpecialDetect

* More than one SpecialDetect key provided
* Invalid SpecialDetect provided

#### Detect

* Invalid Registry path provided

#### DetectFile

* **Trailing backslashes (\\)**
* Nested wildcard provided (supported by Trim, not supported by CCleaner)
* Invalid file system path provided

#### Default

* More than one Default key provided
* **Default value other than "False" provided**

#### Warning

* More than one Warning key provided

#### FileKey

* **Duplicate parameters provided in a single FileKey**
* **Empty FileKey parameters**
* Semicolon provided before Pipe symbol
* Incorrect RECURSE/REMOVESELF spelling/too many Pipe symbols
* Incorrect VirtualStore locations
* **Backslash use before Pipe**
* Missing backslash after %EnvironmentVariable%
* **Experimental: Similar FileKey merger**
  * The debugger can attempt to merge FileKeys it thinks can collapse into a single key. Results may vary.

#### RegKey

* Invalid Registry path provided

#### ExcludeKey

* Missing backslash before pipe symbol
* No valid flag (FILE, PATH, REG) provided

## Command-line args

|Arg|Effect|
|:-|:-
`-1f` or `-1d`|Defines the path for winapp2.ini
`-3f` or `-3d`|Defines the path for the save file
`-c`|Enables saving

# Code Documentation

Below is documentation for how **WinappDebug** works under the hood

## Properties

*italicized* properties are *private*

| Name|Type|Default Value|Description|
|:-|:-|:-|:-|
winappDebugFile1|`iniFile`|`winapp2.ini` in the current directory|The winapp2.ini file that will be linted
winappDebugFile3|`iniFile`|`winapp2.ini` in the current directory|The save path for the linted file (overwrites the input file by default)
RepairSomeErrsFound|`Boolean`|`False`|Indicates that some but not all repairs will run
ScanSettingsChanged|`Boolean`|`False`|Indicates that the scan settings have been modified from their defaults
ModuleSettingsChanged|`Boolean`|`False`|Indicates that the module settings have been modified from their defaults
SaveChanges|`Boolean`|`False`|Indicates that the any changes made by the linter should be saved back to disk
RepairErrsFound|`Boolean`|`True`|Indicates that the linter should attempt to repair errors it finds
ErrorsFound|`Integer`|`0`|The number of errors found during the lint
allEntryNames|`strList`|`Nothing`|The list of all entry names found during the lint
MostRecentLintLog|`String`|`""`|The winapp2ool logslice from the most recent Lint run
Rules|`List(Of lintRule)`|See Description|Rules is created to hold the current lint rules which are created as part of Rules' instantiation. They are aliased below
*lintCasing*|`lintRule`|`Rules(0)`|Controls scan/repairs for CamelCasing issues
*lintAlpha*|`lintRule`|`Rules(1)`|Controls scan/repairs for alphabetization issues
*lintWrongNums*|`lintRule`|`Rules(2)`|Controls scan/repairs for incorrectly numbered keys
*lintParams*|`lintRule`|`Rules(3)`|Controls scan/repairs for parameters inside of FileKeys
*lintFlags*|`lintRule`|`Rules(4)`|Controls scan/repairs for flags in ExcludeKeys and FileKeys
*lintSlashes*|`lintRule`|`Rules(5)`|Controls scan/repairs for improper slash usage
*lintDefaults*|`lintRule`|`Rules(6)`|Controls scan/repairs for missing or True Default values
*lintDupes*|`lintRule`|`Rules(7)`|Controls scan/repairs for duplicate values
*lintExtraNums*|`lintRule`|`Rules(8)`|Controls scan/repairs for keys with numbers they shouldn't have
*lintMulti*|`lintRule`|`Rules(9)`|Controls scan/repairs for keys which should only occur once
*lintInvalid*|`lintRule`|`Rules(10)`|Controls scan/repairs for keys with invlaid values
*lintSyntax*|`lintRule`|`Rules(11)`|Controls scan/repairs for winapp2.ini syntax errors
*lintPathValidity*|`lintRule`|`Rules(12)`|Controls scan/repairs for invalid file or regsitry paths
*lintSemis*|`lintRule`|`Rules(13)`|Controls scan/repairs for improper use of semicolons
*lintOpti*|`lintRule`|`Rules(14)`|Controls scan/repairs for keys that can be merged into eachother (FileKeys only currently)

## Regex properties

*italicized* properties are *private*

| Name|Type|Default Value|Description|
|:-|:-|:-|:-|
longReg|`Regex`|`"HKEY_(C(URRENT_(USER$|CONFIG$)|LASSES_ROOT$)|LOCAL_MACHINE$|USERS$)"`|Regex to detect long form registry paths
shortReg|`Regex`|`"HK(C(C$|R$|U$)|LM$|U$)"`|Detects short form registry paths
secRefNums|`Regex`|`"30(0(5|6)|2([0-9])|3(0|1))"`|Detects valid LangSecRef numbers
driveLtrs|`Regex`|`"[a-zA-z]:"`|Detects valid drive letter parameters
envVarRegex|`Regex`|`"%[A-Za-z0-9]*%"`|Detects potential %EnvironmentVariables%

## handleCmdLine

### Handles the commandline args for **WinappDebug**

```vb
Public Sub handleCmdLine()
    initDefaultSettings()
    invertSettingAndRemoveArg(SaveChanges, "-c")
    getFileAndDirParams(winappDebugFile1, New iniFile, winappDebugFile3)
    If Not cmdargs.Contains("UNIT_TESTING_HALT") Then initDebug()
End Sub
```

## initDefaultSettings

### Restore the default state of the module's parameters

```vb
Private Sub initDefaultSettings()
    winappDebugFile1.resetParams()
    winappDebugFile3.resetParams()
    ModuleSettingsChanged = False
    RepairErrsFound = True
    SaveChanges = False
    resetScanSettings()
End Sub
```

## printMenu

### Displays the WinappDebug menu to the user

```vb
Public Sub printMenu()
    printMenuTop({"Scan winapp2.ini for style and syntax errors, and attempt to repair them where possible."})
    print(1, "Run (Default)", "Run the debugger")
    print(1, "File Chooser (winapp2.ini)", "Choose a new winapp2.ini name or path", leadingBlank:=True, trailingBlank:=True)
    print(5, "Toggle Saving", "saving the file after correcting errors", enStrCond:=SaveChanges)
    print(1, "File Chooser (save)", "Save a copy of changes made instead of overwriting winapp2.ini", SaveChanges, trailingBlank:=True)
    print(1, "Toggle Scan Settings", "Enable or disable individual scan and correction routines", leadingBlank:=Not SaveChanges, trailingBlank:=True)
    print(0, $"Current winapp2.ini:  {replDir(winappDebugFile1.Path)}", closeMenu:=Not SaveChanges And Not ModuleSettingsChanged And MostRecentLintLog = "")
    print(0, $"Current save target:  {replDir(winappDebugFile3.Path)}", cond:=SaveChanges, closeMenu:=Not ModuleSettingsChanged And MostRecentLintLog = "")
    print(2, "WinappDebug", cond:=ModuleSettingsChanged, closeMenu:=MostRecentLintLog = "")
    print(1, "Log Viewer", "Show the most recent lint results", cond:=Not MostRecentLintLog = "", closeMenu:=True, leadingBlank:=True)
End Sub
```

## handleUserInput

### Handles the user's input from the menu

```vb
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
        Case Not MostRecentLintLog = "" And (input = "5" And Not ModuleSettingsChanged) Or
                                        ModuleSettingsChanged And ((input = "6" And Not SaveChanges) Or (input = "7" And SaveChanges))
            printSlice(MostRecentLintLog)
        Case Else
            setHeaderText(invInpStr, True)
    End Select
End Sub
```

|Parameter|Type|Description
|:-|:-|:-
input|`String`|The String containing the user's input

## initDebug

### Validates winapp2.ini and runs the linter from the main menu or commandline

```vb
Private Sub initDebug()
    If Not enforceFileHasContent(winappDebugFile1) Then Exit Sub
    Dim wa2 As New winapp2file(winappDebugFile1)
    clrConsole()
    print(3, "Beginning analysis of winapp2.ini", trailr:=True)
    gLog("Beginning lint", leadr:=True, ascend:=True)
    MostRecentLintLog = ""
    debug(wa2)
    gLog("", descend:=True)
    gLog("Lint complete")
    setHeaderText("Lint complete")
    print(4, "Completed analysis of winapp2.ini", conjoin:=True)
    print(0, $"{ErrorsFound} possible errors were detected.")
    print(0, $"Number of entries: {winappDebugFile1.Sections.Count}", trailingBlank:=True)
    rewriteChanges(wa2)
    print(0, anyKeyStr, closeMenu:=True)
    crk()
End Sub
```

## debug

### Performs syntax and format checking on a winapp2.ini format `iniFile`

```vb
Public Sub debug(ByRef fileToBeDebugged As winapp2file)
    ErrorsFound = 0
    allEntryNames = New strList
    gLog("", ascend:=True)
    For Each entryList In fileToBeDebugged.Winapp2entries
        If entryList.Count = 0 Then Continue For
        entryList.ForEach(Sub(entry) processEntry(entry))
    Next
    fileToBeDebugged.rebuildToIniFiles()
    alphabetizeEntries(fileToBeDebugged)
End Sub
```

## processEntry

### Audits the given `winapp2entry` for errors

```vb
Private Sub processEntry(ByRef entry As winapp2entry)
    gLog($"Processing entry {entry.Name}", buffr:=True)
    Dim hasFileExcludes = False
    Dim hasRegExcludes = False
    ' Check for duplicate names that are differently cased
    fullNameErr(allEntryNames.chkDupes(entry.Name), entry, "Duplicate entry name detected")
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
    fullNameErr(entry.LangSecRef.KeyCount <> 0 And entry.SectionKey.KeyCount <> 0 And lintSyntax.ShouldScan, entry, "Section key found alongside LangSecRef key, only one is required.")
    ' Make sure we have at least 1 valid detect key and at least one valid cleaning key
    fullNameErr(entry.DetectOS.KeyCount + entry.Detects.KeyCount + entry.SpecialDetect.KeyCount + entry.DetectFiles.KeyCount = 0, entry, "Entry has no valid detection keys (Detect, DetectFile, DetectOS, SpecialDetect)")
    fullNameErr(entry.FileKeys.KeyCount + entry.RegKeys.KeyCount = 0 And lintSyntax.ShouldScan, entry, "Entry has no valid FileKeys or RegKeys")
    ' If we don't have FileKeys or RegKeys, we shouldn't have ExcludeKeys.
    fullNameErr(entry.ExcludeKeys.KeyCount > 0 And entry.FileKeys.KeyCount + entry.RegKeys.KeyCount = 0, entry, "Entry has ExcludeKeys but no valid FileKeys or RegKeys")
    ' Make sure that if we have excludes, we also have corresponding file/reg keys
    fullNameErr(entry.FileKeys.KeyCount = 0 And hasFileExcludes, entry, "ExcludeKeys targeting filesystem locations found without any corresponding FileKeys")
    fullNameErr(entry.RegKeys.KeyCount = 0 And hasRegExcludes, entry, "ExcludeKeys targeting registry locations found without any corresponding RegKeys")
    ' Make sure we have a Default key.
    fullNameErr(entry.DefaultKey.KeyCount = 0 And lintDefaults.ShouldScan, entry, "Entry is missing a Default key")
    entry.DefaultKey.add(New iniKey("Default=False"), lintDefaults.fixFormat And entry.DefaultKey.KeyCount = 0)
    gLog($"Finished processing {entry.Name}", buffr:=True)
End Sub
```

|Parameter|Type|Description
|:-|:-|:-|
entry|`winapp2entry`|The `winapp2entry` to be audited for errors

## validateKeys

### Checks the validity of all `iniKeys` in a `winapp2entry` and removes any that are too problematic to continue with

```vb
Private Sub validateKeys(ByRef entry As winapp2entry)
    For Each lst In entry.KeyListList
        Dim brokenKeys As New keyList
        lst.Keys.ForEach(Sub(key) brokenKeys.add(key, Not cValidity(key)))
        lst.remove(brokenKeys.Keys)
        entry.ErrorKeys.remove(brokenKeys.Keys)
    Next
    ' Attempt to assign keys that had errors to their intended lists
    Dim toRemove As New keyList
    For Each key In entry.ErrorKeys.Keys
        For Each lst In entry.KeyListList
            If lst.KeyType = "Error" Then Continue For
            lst.add(key, key.typeIs(lst.KeyType))
            toRemove.add(key, key.typeIs(lst.KeyType))
        Next
    Next
    ' Remove any repaired keys
    entry.ErrorKeys.remove(toRemove.Keys)
End Sub
```

|Parameter|Type|Description
|:-|:-|:-|
entry|`winapp2entry`|A `winapp2entry` whose `iniKeys` will be audited for validity

## alphabetizeEntries

### Alphabetizes all the entries in a winapp2.ini file and observes any that were out of place

```vb
Private Sub alphabetizeEntries(ByRef winapp As winapp2file)
    For Each innerFile In winapp.EntrySections
        Dim unsortedEntryList = innerFile.namesToStrList
        Dim sortedEntryList = sortEntryNames(innerFile)
        findOutOfPlace(unsortedEntryList, sortedEntryList, "Entry", innerFile.getLineNumsFromSections)
        If lintAlpha.fixFormat Then innerFile.sortSections(sortedEntryList)
    Next
End Sub
```

|Parameter|Type|Description
|:-|:-|:-
winapp|`winapp2file`|The `winapp2file` to be linted

## rewriteChanges

### Writes any changes made during the lint process back to disk

```vb
Private Sub rewriteChanges(ByRef winapp2file As winapp2file)
    If SaveChanges Then
        print(0, "Saving changes, do not close winapp2ool or data loss may occur...", leadingBlank:=True)
        winappDebugFile3.overwriteToFile(winapp2file.winapp2string)
        print(0, "Finished saving changes.", trailingBlank:=True)
    End If
End Sub
```

## findOutOfPlace

### Assess a list and its sorted state to observe changes in neighboring strings, such as the changes made while sorting the strings alphabetically

```vb
Private Sub findOutOfPlace(ByRef someList As strList, ByRef sortedList As strList, ByVal findType As String, ByRef LineCountList As List(Of Integer), Optional ByRef oopBool As Boolean = False)
    ' Only try to find out of place keys when there's more than one
    If someList.Count > 1 Then
        Dim misplacedEntries As New strList
        ' Learn the neighbors of each string in each respective list
        Dim initialNeighbors = someList.getNeighborList
        Dim sortedNeighbors = sortedList.getNeighborList
        ' Make sure at least one of the neighbors of each string are the same in both the sorted and unsorted state, otherwise the string has moved
        For i = 0 To someList.Count - 1
            Dim sind = sortedList.Items.IndexOf(someList.Items(i))
            misplacedEntries.add(someList.Items(i), Not (initialNeighbors(i).Key = sortedNeighbors(sind).Key And initialNeighbors(i).Value = sortedNeighbors(sind).Value))
        Next
        ' Report any misplaced entries back to the user
        For Each entry In misplacedEntries.Items
            Dim recInd = someList.indexOf(entry)
            Dim sortInd = sortedList.indexOf(entry)
            Dim curLine = LineCountList(recInd)
            Dim sortLine = LineCountList(sortInd)
            If (recInd = sortInd Or curLine = sortLine) Then Continue For
            entry = If(findType = "Entry", entry, $"{findType & recInd + 1}={entry}")
            If Not oopBool Then oopBool = True
            customErr(LineCountList(recInd), $"{findType} alphabetization", {$"{entry} appears to be out of place",
                                                                                $"Current line: {curLine}",
                                                                                $"Expected line: {sortLine}"})
        Next
    End If
End Sub
```

|Parameter|Type|Description|Optional
|:-|:-|:-|:-
someList|`strList`|An unsorted list of strings|No
sortedList|`strList`|The sorted state of `someList`|No
findType|`String`|The type of neighbor checking \*|No
LineCountList|`List(Of Integer)`|The line numbers associated with the lines in `someList`|No
oopBool|`Boolean`|Indicates that there are out of place entries in the list|Yes, default: `False`

\* When checking keys, `findType` contains a `KeyType`

## processKeyList

### Hands off each `iniKey` in a winapp2.ini format `keyList` to be audited for correctness

```vb
Private Sub processKeyList(ByRef kl As keyList, processKey As Func(Of iniKey, iniKey), Optional ByRef hasF As Boolean = False, Optional ByRef hasR As Boolean = False)
    If kl.KeyCount = 0 Then Exit Sub
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
                If key.typeIs("SpecialDetect") Then chkCasing(key, {"DET_CHROME", "DET_MOZILLA", "DET_THUNDERBIRD", "DET_OPERA"}, key.Value, False)
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
    sortKeys(kl, dupes.KeyCount > 0)
    ' Run optimization checks on FileKey lists only
    If kl.typeIs("FileKey") Then cOptimization(kl)
    gLog("", descend:=True)
End Sub
```

|Parameter|Type|Description|Optional
|:-|:-|:-|:-
kl|`keyList`|A `keyList` of a particular `KeyType` to be audited|No
processKey|`Func(Of iniKey, iniKey)`|The `function` that audits keys of the particular `KeyType` provided in `kl` \*|No
hasF|`Boolean`|Indicates that the ExcludeKeys contain file system locations|Yes, default: `False`
hasR|`Boolean`|Indicates that the ExcludeKeys contain registry locations|Yes, default: False

\* `voidDelegate` is provided for the case where no audits are required beyond what's provided in `processKeyList`

## voidDelegate

### This function does nothing by design

```vb
Private Function voidDelegate(key As iniKey) As iniKey
    Return key
End Function
```

|Parameter|Type|Description|
|:-|:-|:-
key|`iniKey`|An `iniKey` to do nothing with

## cFormat

### Does some basic formatting checks that apply to all winapp2.ini format  `iniKeys`

```vb
Private Sub cFormat(ByRef key As iniKey, ByRef keyNumber As Integer, ByRef keyValues As strList, ByRef dupeList As keyList, Optional noNumbers As Boolean = False)
    ' Check for duplicates
    If keyValues.contains(key.Value, True) Then
        Dim dupeKeyStr = $"{key.KeyType}{If(Not noNumbers, (keyValues.Items.IndexOf(key.Value) + 1).ToString, "")}={key.Value}"
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
    fullKeyErr(key, "Forward slash (/) detected in lieu of backslash (\).", Not key.typeIs("Warning") And lintSlashes.ShouldScan And key.vHas("/"),
                                                                                                        lintSlashes.fixFormat, key.Value, key.Value.Replace("/", "\"))
    fullKeyErr(key, "Trailing semicolon (;).", key.toString.Last = CChar(";") And lintSemis.ShouldScan, lintSemis.fixFormat, key.Value, key.Value.TrimEnd(CChar(";")))
    ' Do some formatting checks for environment variables if needed
    If {"FileKey", "ExcludeKey", "DetectFile"}.Contains(key.KeyType) Then cEnVar(key)
    keyNumber += 1
End Sub
```

|Parameter|Type|Description|Optional
|:-|:-|:-|:-
key|`iniKey`|An `iniKey` whose format will be audited|No
keyNumber|`Integer`|The current expected key number for numbered keys|No
keyValues|`strList`|The current list of observed `iniKey` values|No
dupeList|`keyList`|A tracking list of `iniKeys` with duplicated values|No
noNumbers|`Boolean`|Indicates that the current set of keys should not be numbered|Yes, default: `False`

## cEnVar

### Audits the formatting of any %EnvironmentVariables% in a given `iniKey`

```vb
Private Sub cEnVar(key As iniKey)
    ' Valid Environmental Variables for winapp2.ini
    Dim enVars = {"AllUsersProfile", "AppData", "CommonAppData", "CommonProgramFiles",
        "Documents", "HomeDrive", "LocalAppData", "LocalLowAppData", "Music", "Pictures", "ProgramData", "ProgramFiles", "Public",
        "RootDir", "SystemDrive", "SystemRoot", "Temp", "Tmp", "UserName", "UserProfile", "Video", "WinDir"}
    fullKeyErr(key, "%EnvironmentVariables% must be surrounded on both sides by a single '%' character.", key.vHas("%") And envVarRegex.Matches(key.Value).Count = 0)
    For Each m As Match In envVarRegex.Matches(key.Value)
        Dim strippedText = m.ToString.Trim(CChar("%"))
        chkCasing(key, enVars, strippedText, False)
    Next
    fullKeyErr(key, "Missing backslash (\) after %EnvironmentVariable%.", lintSlashes.ShouldScan And key.vHas("%") And Not key.vHasAny({"%|", "%\"}))
End Sub
```

|Parameter|Type|Description
|:-|:-|:-|
key|`iniKey`|The `iniKey` to audit

## fixMissingEquals

### Attempts to insert missing equal signs (=) into `iniKeys`

Returns `True` if the repair is successful, otherwise returns `False`

```vb
Private Function fixMissingEquals(ByRef key As iniKey, cmds As String()) As Boolean
    gLog("Attempting missing equals repair", ascend:=True)
    For Each cmd In cmds
        If key.Name.ToLower.Contains(cmd.ToLower) Then
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
                        If Char.IsNumber(c) Then
                            newName += c
                        Else
                            Exit For
                        End If
                    Next
                    key.Value = key.Name.Replace(newName, "")
                    key.Name = newName
                    key.KeyType = cmd
            End Select
            gLog($"Repair complete. Result: {key.toString}", indent:=True, descend:=True)
            Return True
        End If
    Next
    ' Return false if no valid command is found
    gLog("Repair failed, key will be removed.", descend:=True)
    Return False
End Function
```

|Parameter|Type|Description
|:-|:-|:-|
key|`iniKey`|The misformatted `iniKey` to be repaired
cmds|`String()`|The array containing valid winapp2.ini `KeyTypes`

## cValidity

### Does basic syntax and formatting audits that apply across all keys, returns `False` iff a key is malformed

```vb
Private Function cValidity(key As iniKey) As Boolean
    Dim validCmds = {"Default", "DetectOS", "DetectFile", "Detect", "ExcludeKey",
                        "FileKey", "LangSecRef", "RegKey", "Section", "SpecialDetect", "Warning"}
    If key.typeIs("DeleteMe") Then
        gLog($"Broken Key Found: {key.Name}", indent:=True, ascend:=True)
        ' Try to fix broken keys
        If fixMissingEquals(key, validCmds) Then
            fullKeyErr(key, "Missing '=' detected and repaired in key.")
        Else
            ' If we didn't find a fixable situation, delete the key
            customErr(key.LineNumber, $"{key.Name} is missing a '=' or was not provided with a value. It will be deleted.", {})
            Return False
        End If
    End If
    If key.Value.Contains("\\") Then
        fullKeyErr(key, "Extraneous backslashes (\\) detected", lintSlashes.ShouldScan)
        While (key.Value.Contains("\\") And lintSlashes.fixFormat)
            key.Value = key.Value.Replace("\\", "\")
        End While
    End If
    ' Check for leading or trailing whitespace, do this always as spaces in the name interfere with proper keyType identification
    If key.Name.StartsWith(" ") Or key.Name.EndsWith(" ") Or key.Value.StartsWith(" ") Or key.Value.EndsWith(" ") Then
        fullKeyErr(key, "Detected unwanted whitespace in iniKey", True)
        fixStr(True, key.Value, key.Value.Trim)
        fixStr(True, key.Name, key.Name.Trim)
        fixStr(True, key.KeyType, key.KeyType.Trim)
    End If
    ' Make sure the keyType is valid
    chkCasing(key, validCmds, key.KeyType, True)
    Return True
End Function
```

|Parameter|Type|Description
|:-|:-|:-|
key|`iniKey`|The `iniKey` whose validity will be audited

## chkCasing

### Checks either the `Value` or the `KeyType` of an `iniKey` against a given array of expected cased values

```vb
Private Sub chkCasing(ByRef key As iniKey, casedArray As String(), ByRef strToChk As String, chkType As Boolean)
    ' Get the properly cased string
    Dim casedString = strToChk
    For Each casedText In casedArray
        If strToChk.Equals(casedText, StringComparison.InvariantCultureIgnoreCase) Then casedString = casedText
    Next
    ' Determine if there's a casing error
    Dim hasCasingErr = Not casedString.Equals(strToChk) And casedArray.Contains(casedString)
    Dim replacementText = If(chkType, key.Name.Replace(key.KeyType, casedString), key.Value.Replace(key.Value, casedString))
    Dim validData = String.Join(", ", casedArray)
    fullKeyErr(key, $"{casedString} has a casing error.", hasCasingErr And lintCasing.ShouldScan, lintCasing.fixFormat, strToChk, replacementText)
    fixStr(chkType, key.Name, replacementText)
    fullKeyErr(key, $"Invalid data provided: {strToChk} in {key.toString}{Environment.NewLine}Valid data: {validData}", Not casedArray.Contains(casedString) And lintInvalid.ShouldScan)
End Sub
```

|Parameter|Type|Description
|:-|:-|:-|
key|`iniKey`|The `iniKey` whose casing will be audited
casedArray|`String()`|The array of expected cased values
strToChk|`String`|A pointer to the value being audited
chkType|`Boolean`|`True` to check `KeyTypes`, `False` to check `Values`

## pFileKey

### Processes a FileKey format winapp2.ini `iniKey` and checks it for errors, correcting them where possible

```vb
Public Function pFileKey(key As iniKey) As iniKey
    ' Pipe symbol checks
    Dim iteratorCheckerList = Split(key.Value, "|")
    fullKeyErr(key, "Missing pipe (|) in FileKey.", Not key.vHas("|"))
    ' Captures any incident of semi colons coming before the first pipe symbol
    fullKeyErr(key, "Semicolon (;) found before pipe (|).", lintSemis.ShouldScan And key.vHas(";") And (key.Value.IndexOf(";") < key.Value.IndexOf("|")))
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
            customErr(key.LineNumber, $"{If(arg = "", "Empty", "Duplicate")} FileKey parameter found", {$"Command: {arg}"})
            dupeArgs.add(arg, lintParams.fixFormat)
        End If
    Next
    ' Remove any duplicate arguments from the key parameters and reconstruct keys we've modified above
    If lintParams.fixFormat Then
        dupeArgs.Items.ForEach(Sub(arg) keyParams.ArgsList.Remove(arg))
        keyParams.reconstructKey(key)
    End If
    Return key
End Function
```

|Parameter|Type|Description
|:-|:-|:-|
key|`iniKey`|A winapp2.ini FileKey format `iniKey` to be checked for correctness

## pDetectFile

### Processes a DetectFile format `iniKey` and checks it for errors, correcting where possible

```vb
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
```

|Parameter|Type|Description
|:-|:-|:-|
key|`iniKey`|A winapp2.ini DetectFile format `iniKey` to be checked for correctness

## chkPathFormatValidity

### Audits the syntax of file system and registry paths

```vb
Private Sub chkPathFormatValidity(key As iniKey, isRegistry As Boolean)
    If Not lintPathValidity.ShouldScan Then Exit Sub
    ' Remove the flags from ExcludeKeys if we have them before getting the first directory portion
    Dim rootStr = If(key.KeyType <> "ExcludeKey", getFirstDir(key.Value), getFirstDir(pathFromExcludeKey(key)))
    ' Ensure that registry paths have a valid hive and file paths have either a variable or a drive letter
    fullKeyErr(key, "Invalid registry path detected.", isRegistry And Not longReg.IsMatch(rootStr) And Not shortReg.IsMatch(rootStr))
    fullKeyErr(key, "Invalid file system path detected.", Not isRegistry And Not driveLtrs.IsMatch(rootStr) And Not rootStr.StartsWith("%"))
End Sub
```

|Parameter|Type|Description
|:-|:-|:-|
key|`iniKey`|An `iniKey` to be audited
isRegistry|`Boolean`|Indicates that the given key is expected to hold a registry path

## pExcludeKey

### Processes a list of ExcludeKey format `iniKeys` and checks them for errors, correcting where possible

```vb
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
End
```

|Parameter|Type|Description
|:-|:-|:-|
key|`iniKey`|A winapp2.ini ExcludeKey format `iniKey` to be checked for correctness
hasF|`Boolean`|Indicates whether the entry excludes any filesystem locations
hasR|`Boolean`|Indicates whether the entry excludes any registry locations

## sortKeys

### Sorts a `keyList` alphabetically with winapp2.ini precedence applied to the key values

```vb
Private Sub sortKeys(ByRef kl As keyList, hadDuplicatesRemoved As Boolean)
    If Not lintAlpha.ShouldScan Or kl.KeyCount <= 1 Then Exit Sub
    Dim keyValues = kl.toStrLst(True)
    Dim sortedKeyValues = replaceAndSort(keyValues, "|", " \ \")
    ' Rewrite the alphabetized keys back into the keylist (also fixes numbering)
    Dim keysOutOfPlace = False
    findOutOfPlace(keyValues, sortedKeyValues, kl.KeyType, kl.lineNums, keysOutOfPlace)
    If (keysOutOfPlace Or hadDuplicatesRemoved) And (lintAlpha.fixFormat Or lintWrongNums.fixFormat Or lintExtraNums.fixFormat) Then
        kl.renumberKeys(sortedKeyValues)
    End If
End Sub
```

|Parameter|Type|Description
|:-|:-|:-|
kl|`keyList`|The `keyList` to be sorted
hadDuplicatesRemoved|`Boolean`|Indicates that keys have been removed from `kl`

## inputMismatchErr

### Prints an error when data is received that does not match an expected value

```vb
Private Sub inputMismatchErr(linecount As Integer, err As String, received As String, expected As String, Optional cond As Boolean = True)
    If cond Then customErr(linecount, err, {$"Expected: {expected}", $"Found: {received}"})
End Sub
```

|Parameter|Type|Description|Optional
|:-|:-|:-|:-
linecount|`Integer`|The line number on which the error was detected|No
err|`String`|A description of the error as it will be displayed to the user|No
received|`String`|The erroneous input data|No
expected|`String`|The expected data|No
cond|`Boolean`|Indicates that the error condition is present|Yes, default: `True`

## fullNameErr

### Prints an error followed by the [Full Name *] of the entry to which it belongs

```vb
Private Sub fullNameErr(cond As Boolean, entry As winapp2entry, errTxt As String)
    If cond Then customErr(entry.LineNum, errTxt, {$"Entry Name: {entry.FullName}"})
End Sub
```

|Parameter|Type|Description
|:-|:-|:-
cond|`Boolean`|Indicates that the error condition is present
entry|`winapp2entry`|The `winapp2entry` containing an error
errTxt|`String`|A description of the error as it will be displayed to the user

## customErr

### Prints arbitrarily defined errors without a precondition

```vb
Private Sub customErr(lineCount As Integer, err As String, lines As String())
    gLog(err, ascend:=True)
    cwl($"Line: {lineCount} - Error: {err}")
    MostRecentLintLog += $"Line: {lineCount} - Error: {err}" & Environment.NewLine
    For Each errStr In lines
        cwl(errStr)
        gLog(errStr, indent:=True)
        MostRecentLintLog += errStr & Environment.NewLine
    Next
    gLog("", descend:=True)
    cwl()
    MostRecentLintLog += Environment.NewLine
    ErrorsFound += 1
End Sub
```

|Parameter|Type|Description
|:-|:-|:-
lintCount|`Integer`|The line number on which the error was detected
err|`String`|A description of the error as it will be displayed to the user
lines|`String()`|Any additional error information to be printed alongside the description

## fixStr

### Replace a given string with a new value if the fix condition is met

```vb
Private Sub fixStr(param As Boolean, ByRef currentValue As String, ByRef newValue As String)
    If param Then
        gLog($"Changing '{currentValue}' to '{newValue}'", ascend:=True, descend:=True, indent:=True, buffr:=True)
        currentValue = newValue
    End If
End Sub
```

|Parameter|Type|Description
|:-|:-|:-
param|`Boolean`|The condition under which the string should be replaced
currentValue|`String`|A pointer to the string to be replaced
newValue|`String`|The replacement value for `currentValue`
