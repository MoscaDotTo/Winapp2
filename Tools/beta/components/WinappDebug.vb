'    Copyright (C) 2018 Robbie Ward
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
Imports System.IO
Imports System.Text.RegularExpressions

''' <summary>
''' A program whose purpose is to observe and attempt to repair errors in winapp2.ini
''' </summary>
Module WinappDebug

    'File handlers
    Dim winappFile As iniFile = New iniFile(Environment.CurrentDirectory, "winapp2.ini")
    Dim outputFile As iniFile = New iniFile(Environment.CurrentDirectory, "winapp2.ini", "winapp2-debugged.ini")

    'Menu settings
    Dim settingsChanged As Boolean = False
    Dim scanSettingsChanged As Boolean = False

    'Module parameters
    Dim correctFormatting As Boolean = False
    Dim allEntryNames As New List(Of String)
    Dim numErrs As Integer = 0
    Dim correctSomeFormatting As Boolean = True

    'Autocorrect Parameters
    Dim correctCasing As Boolean = True
    Dim correctAlpha As Boolean = True
    Dim correctNumbers As Boolean = True
    Dim correctParameters As Boolean = True
    Dim correctFlags As Boolean = True
    Dim correctSlashes As Boolean = True
    Dim correctDefaults As Boolean = True
    Dim correctOpti As Boolean = False

    'Scan parameters
    Dim scanCasing As Boolean = True
    Dim scanAlpha As Boolean = True
    Dim scanNumbers As Boolean = True
    Dim scanParams As Boolean = True
    Dim scanFlags As Boolean = True
    Dim scanSlashes As Boolean = True
    Dim scanDefaults As Boolean = True
    Dim scanOpti As Boolean = False

    'Winapp2 Parameters
    Dim enVars As String() = {"UserProfile", "UserName", "ProgramFiles", "RootDir", "WinDir", "AppData", "SystemDrive", "SystemRoot", "Documents", "ProgramData", "AllUsersProfile", "Pictures", "Video", "CommonAppData", "LocalAppData", "CommonProgramFiles", "HomeDrive", "Music", "tmp", "Temp", "LocalLowAppData", "Public"}
    Dim validCmds As String() = {"SpecialDetect", "FileKey", "RegKey", "Detect", "LangSecRef", "Warning", "Default", "Section", "ExcludeKey", "DetectFile", "DetectOS"}
    Dim sdList As String() = {"DET_CHROME", "DET_MOZILLA", "DET_THUNDERBIRD", "DET_OPERA"}

    'Regex strings 
    ReadOnly longReg As New Regex("HKEY_(C(URRENT_(USER$|CONFIG$)|LASSES_ROOT$)|LOCAL_MACHINE$|USERS$)")
    ReadOnly shortReg As New Regex("HK(C(C$|R$|U$)|LM$|U$)")
    ReadOnly secRefNums As New Regex("30(2([0-9])|3(0|1))")
    ReadOnly driveLtrs As New Regex("[a-zA-z]:")
    ReadOnly envVarRegex As New Regex("%[A-Za-z0-9]*%")

    ''' <summary>
    ''' Resets the individual scan settings to their defaults
    ''' </summary>
    Private Sub resetScanSettings()
        correctCasing = True
        correctAlpha = True
        correctNumbers = True
        correctParameters = True
        correctFlags = True
        correctSlashes = True
        correctDefaults = True
        correctOpti = False
        scanCasing = True
        scanAlpha = True
        scanNumbers = True
        scanParams = True
        scanFlags = True
        scanSlashes = True
        scanDefaults = True
        scanOpti = False
        scanSettingsChanged = False
        correctSomeFormatting = False
    End Sub

    ''' <summary>
    ''' Initializes the default module settings and returns references to them to the calling function
    ''' </summary>
    ''' <param name="firstFile">The winapp2.ini file</param>
    ''' <param name="secondFile">The save file</param>
    ''' <param name="cf">boolean for autocorrect</param>
    Public Sub initDebugParams(ByRef firstFile As iniFile, ByRef secondFile As iniFile, cf As Boolean)
        initDefaultSettings()
        firstFile = winappFile
        secondFile = outputFile
        cf = correctFormatting
    End Sub

    ''' <summary>
    ''' Restore the default state of the module's parameters
    ''' </summary>
    Private Sub initDefaultSettings()
        winappFile.resetParams()
        outputFile.resetParams()
        settingsChanged = False
        correctFormatting = False
        resetScanSettings()
    End Sub

    ''' <summary>
    ''' Begins the debugger from outside the module
    ''' </summary>
    ''' <param name="firstFile">The winapp2.ini file</param>
    ''' <param name="secondFile">The output file</param>
    ''' <param name="cformatting">The autocorrect boolean</param>
    Public Sub remoteDebug(ByRef firstFile As iniFile, secondFile As iniFile, cformatting As Boolean)
        winappFile = firstFile
        outputFile = secondFile
        correctFormatting = cformatting
        initDebug()
    End Sub

    ''' <summary>
    ''' Matches paired boolean states for scans and their repairs
    ''' </summary>
    ''' ie. make sure that if a scan is disabled, so too is its repair, and if a repair is enabled, so too is its scan
    ''' <param name="setting"></param>
    ''' <param name="pairedSetting"></param>
    ''' <param name="type"></param>
    Private Sub toggleScanSetting(ByRef setting As Boolean, ByRef pairedSetting As Boolean, type As String)
        If type = "Scan" Then
            toggleSettingParam(setting, "Scan ", scanSettingsChanged)
            If Not (setting) And pairedSetting Then toggleSettingParam(pairedSetting, "Scan ", scanSettingsChanged)
        Else
            toggleSettingParam(setting, "Repair ", scanSettingsChanged)
            If setting And Not pairedSetting Then toggleSettingParam(pairedSetting, "Scan ", scanSettingsChanged)
        End If
    End Sub

    ''' <summary>
    ''' Prints the menu for individual scans and their repairs to the user
    ''' </summary>
    Private Sub printScansMenu()
        Console.WindowHeight = 35
        printMenuTop({"Enable or disable specific scans or repairs"})
        print(0, "Scan Options", leadingBlank:=True, trailingBlank:=True)
        print(1, "Casing", $"{enStr(scanCasing)} detecting improper CamelCasing")
        print(1, "Alphabetization", $"{enStr(scanAlpha)} detecting improper alphabetization")
        print(1, "Numbering", $"{enStr(scanNumbers)} detecting improper key numbering")
        print(1, "Parameters", $"{enStr(scanParams)} detecting improper parameterization on FileKeys and ExcludeKeys")
        print(1, "Flags", $"{enStr(scanFlags)} detecting improper RECURSE and REMOVESELF formatting")
        print(1, "Slashes", $"{enStr(scanSlashes)} detecting propblems surrounding use of slashes (\)")
        print(1, "Defaults", $"{enStr(scanDefaults)} detecting Default=True or missing Default")
        print(1, "Optimizations", $"{enStr(scanOpti)} detecting situations where keys can be merged")
        print(0, "Repair Options", leadingBlank:=True, trailingBlank:=True)
        print(1, "Casing", $"{enStr(correctCasing)} fixing improper CamelCasing")
        print(1, "Alphabetization", $"{enStr(correctAlpha)} fixing improper alphabetization")
        print(1, "Numbering", $"{enStr(correctNumbers)} fixing improper key numbering")
        print(1, "Parameters", $"{enStr(correctParameters)} fixing improper parameterization On FileKeys And ExcludeKeys")
        print(1, "Flags", $"{enStr(correctFlags)} fixing improper RECURSE And REMOVESELF formatting")
        print(1, "Slashes", $"{enStr(correctSlashes)} fixing propblems surrounding use Of slashes (\)")
        print(1, "Defaults", $"{enStr(correctDefaults)} setting Default=True To Default=False or missing Default")
        print(1, "Optimizations", $"{enStr(correctOpti)} automatic merging of keys", closeMenu:=Not scanSettingsChanged)
        print(2, "Scan And Repair", cond:=scanSettingsChanged, closeMenu:=True)
    End Sub

    ''' <summary>
    ''' Handles the user input for the scan settings menu
    ''' </summary>
    ''' <param name="input">The String containing the user's input</param>
    Private Sub handleScanInput(input As String)
        Dim originalSettingsState As Boolean = New Boolean() {scanCasing, scanAlpha, scanNumbers, scanParams, scanFlags, scanSlashes, scanDefaults,
                correctCasing, correctAlpha, correctNumbers, correctParameters, correctFlags, correctSlashes,
                correctDefaults, scanSettingsChanged = False, scanOpti = False, correctOpti = False}.All(Function(x As Boolean) x)
        settingsChanged = Not originalSettingsState
        Dim tmp3 As Boolean = New Boolean() {correctFormatting = False, correctCasing, correctAlpha, correctNumbers, correctParameters, correctFlags, correctSlashes, correctDefaults}.All(Function(x As Boolean) x)
        correctFormatting = Not tmp3
        If Not correctFormatting And tmp3 Then correctSomeFormatting = True
        Select Case input
            Case "0"
                If scanSettingsChanged Then settingsChanged = True
                exitCode = True
            Case "1"
                toggleScanSetting(scanCasing, correctCasing, "Scan")
            Case "2"
                toggleScanSetting(scanAlpha, correctAlpha, "Scan")
            Case "3"
                toggleScanSetting(scanNumbers, correctNumbers, "Scan")
            Case "4"
                toggleScanSetting(scanParams, correctParameters, "Scan")
            Case "5"
                toggleScanSetting(scanFlags, correctFlags, "Scan")
            Case "6"
                toggleScanSetting(scanSlashes, correctSlashes, "Scan")
            Case "7"
                toggleScanSetting(scanDefaults, correctDefaults, "Scan")
            Case "8"
                toggleScanSetting(scanOpti, correctOpti, "Scan")
            Case "9"
                toggleScanSetting(correctCasing, scanCasing, "Repair")
            Case "10"
                toggleScanSetting(correctAlpha, scanAlpha, "Repair")
            Case "11"
                toggleScanSetting(correctNumbers, scanNumbers, "Repair")
            Case "12"
                toggleScanSetting(correctParameters, scanParams, "Repair")
            Case "13"
                toggleScanSetting(correctFlags, scanFlags, "Repair")
            Case "14"
                toggleScanSetting(correctSlashes, scanSlashes, "Repair")
            Case "15"
                toggleScanSetting(correctDefaults, scanDefaults, "Repair")
            Case "16"
                toggleScanSetting(correctOpti, scanOpti, "Repair")
            Case "17"
                menuHeaderText = If(scanSettingsChanged, "Settings Reset", invInpStr)
                If scanSettingsChanged Then resetScanSettings()
            Case Else
                menuHeaderText = invInpStr
        End Select
    End Sub

    ''' <summary>
    ''' Prints the main menu to the user
    ''' </summary>
    Public Sub printMenu()
        printMenuTop({"Scan winapp2.ini for syntax and style errors, and attempt to repair them."})
        print(1, "Run (Default)", "Run the debugger")
        print(1, "File Chooser (winapp2.ini)", "Choose a new winapp2.ini name or path", leadingBlank:=True, trailingBlank:=True)
        print(1, "Toggle Autocorrect", $"{enStr(correctFormatting)} saving of corrected errors")
        print(1, "File Chooser (save)", "Save a copy of changes made instead of overwriting winapp2.ini", correctFormatting, trailingBlank:=True)
        print(1, "Toggle Scan Settings", "Enable Or disable individual scan and correction routines", leadingBlank:=Not correctFormatting, trailingBlank:=True)
        print(0, $"Current winapp2.ini:  {replDir(winappFile.path)}", closeMenu:=Not correctFormatting And Not settingsChanged)
        print(0, $"Current save target:  {replDir(outputFile.path)}", cond:=correctFormatting, closeMenu:=Not settingsChanged)
        print(2, "WinappDebug", cond:=settingsChanged, closeMenu:=True)
    End Sub

    ''' <summary>
    ''' Handles the user's input from the menu
    ''' </summary>
    ''' <param name="input">The String containing the user's input</param>
    Public Sub handleUserInput(input As String)
        Select Case True
            Case input = "0"
                exitModule("WinappDebug")
            Case input = "1" Or input = ""
                initDebug()
            Case input = "2"
                changeFileParams(winappFile, settingsChanged)
            Case input = "3"
                toggleSettingParam(correctFormatting, "Autocorrect ", settingsChanged)
            Case input = "4" And correctFormatting
                changeFileParams(outputFile, settingsChanged)
            Case (input = "4" And Not correctFormatting) Or (input = "5" And correctFormatting)
                initModule("Scan Settings", AddressOf printScansMenu, AddressOf handleScanInput)
            Case settingsChanged And ((input = "5" And Not correctFormatting) Or (input = "6" And correctFormatting))
                resetModuleSettings("WinappDebug", AddressOf initDefaultSettings)
            Case Else
                menuHeaderText = invInpStr
        End Select
    End Sub

    ''' <summary>
    ''' Validates and debugs the ini file, informs the user upon completion
    ''' </summary>
    Private Sub initDebug()
        winappFile.validate()
        debug(winappFile)
        menuHeaderText = "Debug Complete"
    End Sub

    ''' <summary>
    ''' Performs syntax and format checking on a winapp2.ini format ini file
    ''' </summary>
    ''' <param name="cfile">A winapp2.ini format iniFile object</param>
    Private Sub debug(cfile As iniFile)
        If exitCode Then Exit Sub
        Console.Clear()

        print(0, tmenu("Beginning analysis of winapp2.ini"), closeMenu:=True)
        cwl()
        numErrs = 0
        allEntryNames = New List(Of String)
        Dim winapp2file As New winapp2file(cfile)
        'Process the winapp2entries
        processEntries(winapp2file)
        'Rebuild our internal changes
        winapp2file.rebuildToIniFiles()
        'Check the alphabetization of entries by name
        alphabetizeEntries(winapp2file)
        'Print out the analysis results
        print(0, tmenu("Completed analysis of winapp2.ini"))
        print(0, menuStr03)
        print(0, $"{numErrs} possible errors were detected.")
        print(0, $"Number of entries: {cfile.sections.Count}", trailingBlank:=True)
        'Save changes we've made if necessary
        rewriteChanges(winapp2file)
        print(0, anyKeyStr, closeMenu:=True)
        If Not suppressOutput Then Console.ReadKey()
    End Sub

    ''' <summary>
    ''' Initiates the debugger on each entry in a given winapp2file
    ''' </summary>
    ''' <param name="winapp">The winapp2file to be debugged</param>
    Private Sub processEntries(ByRef winapp As winapp2file)
        For Each entryList In winapp.winapp2entries
            If entryList.Count = 0 Then Continue For
            For Each entry In entryList
                processEntry(entry)
            Next
        Next
    End Sub

    ''' <summary>
    ''' Alphabetizes all the entries in a winapp2.ini file and observes any that were out of place
    ''' </summary>
    ''' <param name="winapp">The object containing the winapp2.ini being operated on</param>
    Private Sub alphabetizeEntries(ByRef winapp As winapp2file)
        For Each innerFile In winapp.entrySections
            Dim unsortedEntryList As List(Of String) = innerFile.getSectionNamesAsList
            Dim sortedEntryList As List(Of String) = sortEntryNames(innerFile)
            findOutOfPlace(unsortedEntryList, sortedEntryList, "Entry Name", innerFile.getLineNumsFromSections)
            If fixFormat(correctAlpha) Then innerFile.sortSections(sortedEntryList)
        Next
    End Sub

    ''' <summary>
    ''' Overwrites the the file on disk with any changes we've made if we are saving them
    ''' </summary>
    ''' <param name="winapp2file">The object representing winapp2.ini</param>
    Private Sub rewriteChanges(ByRef winapp2file As winapp2file)
        If correctFormatting Then
            print(0, "Saving changes, do not close winapp2ool or data loss may occur...", leadingBlank:=True)
            outputFile.overwriteToFile(winapp2file.winapp2string)
            print(0, "Finished saving changes.", trailingBlank:=True)
        End If
    End Sub

    ''' <summary>
    ''' Construct a list of neighbors for strings in a list
    ''' </summary>
    ''' <param name="someList">A list of strings</param>
    ''' <param name="neighborList">The paired values of neighbors in the list of strings</param>
    Private Sub buildNeighborList(someList As List(Of String), neighborList As List(Of KeyValuePair(Of String, String)))
        neighborList.Add(New KeyValuePair(Of String, String)("first", someList(1)))
        For i As Integer = 1 To someList.Count - 2
            neighborList.Add(New KeyValuePair(Of String, String)(someList(i - 1), someList(i + 1)))
        Next
        neighborList.Add(New KeyValuePair(Of String, String)(someList(someList.Count - 2), "last"))
    End Sub

    ''' <summary>
    ''' Assess a list and its sorted state to obvserve changes in neighboring strings
    ''' </summary>
    ''' eg. changes made to string ordering through sorting
    ''' <param name="someList">A list of strings</param>
    ''' <param name="sortedList">The sorted state of someList</param>
    ''' <param name="findType">The type of neighbor checking (keyType for iniKey values)</param>
    ''' <param name="LineCountList">A list containing the line counts of the Strings in someList</param>
    Private Sub findOutOfPlace(ByRef someList As List(Of String), ByRef sortedList As List(Of String), ByVal findType As String, ByRef LineCountList As List(Of Integer), Optional ByRef oopBool As Boolean = False)
        'Only try to find out of place keys when there's more than one
        If someList.Count > 1 Then
            Dim misplacedEntries As New List(Of String)
            'Learn the neighbors of each string in each respective list
            Dim initalNeighbors As New List(Of KeyValuePair(Of String, String))
            Dim sortedNeigbors As New List(Of KeyValuePair(Of String, String))
            'Build our neighbor lists
            buildNeighborList(someList, initalNeighbors)
            buildNeighborList(sortedList, sortedNeigbors)
            'Make sure at least one of the neighbors of each string are the same in both the sorted and unsorted state, otherwise the string has moved 
            For i As Integer = 0 To someList.Count - 1
                Dim sind As Integer = sortedList.IndexOf(someList(i))
                If Not (initalNeighbors(i).Key = sortedNeigbors(sind).Key And initalNeighbors(i).Value = sortedNeigbors(sind).Value) Then misplacedEntries.Add(someList(i))
            Next
            'Report any misplaced entries back to the user
            For Each entry In misplacedEntries
                Dim recind As Integer = someList.IndexOf(entry)
                Dim sortind As Integer = sortedList.IndexOf(entry)
                Dim curpos As Integer = LineCountList(recind)
                Dim sortPos As Integer = LineCountList(sortind)
                If (recind = sortind Or curpos = sortPos) Then Continue For
                entry = If(entry = "Entry Name", entry, findType & recind + 1 & "=" & entry)
                If Not oopBool Then oopBool = True
                customErr(LineCountList(recind), $"{findType} alphabetization", {$"{entry} appears to be out of place", $"Current line: {curpos}", $"Expected line: {sortPos}"})
            Next
        End If
    End Sub

    ''' <summary>
    ''' Initiates the cFormat function
    ''' </summary>
    ''' <param name="keyList">A given list of iniKey objects to have their format audited</param>
    Private Sub initCFormat(ByRef keyList As List(Of iniKey))
        Dim curNum As Integer = 1
        Dim curStrings As New List(Of String)
        Dim dupeKeys As New List(Of iniKey)
        For Each key In keyList
            cFormat(key, curNum, curStrings, dupeKeys)
        Next
        deDupeAndSort(keyList, dupeKeys)
    End Sub

    ''' <summary>
    ''' Does some basic formatting checks that apply to most winapp2.ini format iniKey objects
    ''' </summary>
    ''' <param name="key">The current iniKey object being processed</param>
    ''' <param name="keyNumber">The current expected key number</param>
    ''' <param name="keyValues">The current list of observed inikey values</param>
    ''' <param name="dupeList">A tracking list of detected duplicate iniKeys</param>
    Private Sub cFormat(ByVal key As iniKey, ByRef keyNumber As Integer, ByRef keyValues As List(Of String), ByRef dupeList As List(Of iniKey))
        'Check for wholly duplicate keys and audit their numbering if applicable
        checkDupsAndNumbering(keyValues, key, keyNumber, dupeList)
        'Scan for and fix any use of incorrect slashes
        fullKeyErr(key, "Forward slash (/) detected in lieu of blackslash (\)", scanSlashes And key.vHas(CChar("/")), correctSlashes, key.value, key.value.Replace(CChar("/"), CChar("\")))
        'make sure we don't have any dangly bits on the end of our key
        fullKeyErr(key, "Trailing semicolon (;).", key.toString.Last = CChar(";") And scanParams, correctParameters, key.value, key.value.TrimEnd(CChar(";")))
        'Do some formatting checks for environment variables
        If {"FileKey", "ExcludeKey", "DetectFile"}.Contains(key.keyType) Then cEnVar(key)
    End Sub

    ''' <summary>
    ''' Checks a String for casing errors against a provided array of cased strings, returns the input string if no error is detected
    ''' </summary>
    ''' <param name="caseArray">The parent array of cased Strings</param>
    ''' <param name="keyValue">The string to be checked for casing errors</param>
    Private Function checkCasingError(caseArray As String(), keyValue As String) As String
        For Each casedText In caseArray
            If keyValue.ToLower = casedText.ToLower Then Return casedText
        Next
        Return keyValue
    End Function

    ''' <summary>
    ''' Assess the formatting of any %EnvironmentVariables% in a given iniKey
    ''' </summary>
    ''' <param name="key">An iniKey object to be processed</param>
    Private Sub cEnVar(key As iniKey)
        fullKeyErr(key, "%EnvironmentVariables% must be surrounded on both sides by a single '%' character.", key.vHas("%") And envVarRegex.Matches(key.value).Count = 0)
        For Each m As Match In envVarRegex.Matches(key.value)
            Dim strippedText As String = m.ToString.Trim(CChar("%"))
            Dim camelText As String = checkCasingError(enVars, strippedText)
            fullKeyErr(key, $"Invalid CamelCasing - expected %{camelText}% but found {m.ToString}", scanCasing And camelText <> strippedText, correctCasing, key.value, key.value.Replace(strippedText, camelText))
            'If we don't have a casing error and enVars doesn't contain our value, it's invalid. 
            fullKeyErr(key, $"Misformatted Or invalid environment variable:  {m.ToString}", camelText = strippedText And Not enVars.Contains(strippedText))
        Next
    End Sub

    ''' <summary>
    ''' Does basic syntax and formatting audits that apply across all keys, returns false iff a key in malformatted
    ''' </summary>
    ''' <param name="key">an iniKey object to be checked for errors</param>
    ''' <returns></returns>
    Private Function cValidity(key As iniKey) As Boolean
        'Return false immediately if we are meant to delete the current key 
        If key.keyType = "DeleteMe" Then Return False
        'Check for leading or trailing whitespace
        fullKeyErr(key, "Detected unwanted whitespace in iniKey value", key.vStartsOrEndsWith(" "), True, key.value, key.value.Trim(CChar(" ")))
        fullKeyErr(key, "Detected unwanted whitespace in iniKey name", key.nStartsOrEndsWith(" "), True, key.name, key.name.Trim(CChar(" ")))
        'Make sure the keyType is a valid winapp2.ini command
        Dim casedString As String = checkCasingError(validCmds, key.keyType)
        fullKeyErr(key, $"{casedString} has a casing error.", casedString <> key.keyType And scanCasing, correctCasing, key.name, key.name.Replace(key.keyType, casedString))
        fixStr(casedString <> key.keyType, key.keyType, casedString)
        fullKeyErr(key, "Invalid keyType detected.", Not validCmds.Contains(key.keyType))
        Return True
    End Function

    ''' <summary>
    ''' Determines whether or not a fix that sits behind an optional flag should be run
    ''' </summary>
    ''' <param name="setting">The parameter to be observed if only correcting some formatting</param>
    ''' <returns></returns>
    Private Function fixFormat(setting As Boolean) As Boolean
        Return correctFormatting Or (correctSomeFormatting And setting)
    End Function

    ''' <summary>
    ''' Processes a list of LangSecRef or Section format winapp2.ini iniKey objects and checks them for errors, correcting where possible
    ''' </summary>
    ''' <param name="entry">A winapp2entry to process</param>
    Private Sub pLangSecRef(ByRef entry As winapp2entry)
        'Make sure we only have LangSecRef if we have LangSecRef at all.
        fullKeyErr(If(entry.sectionKey.Count = 0, New iniKey("1=1", 1), entry.sectionKey.First), "Section key found alongside LangSecRef key, only one is required.", entry.langSecRef.Count <> 0 And entry.sectionKey.Count <> 0)
        'Ensure at most one key for both LangSecRef and Section
        confirmOnlyOne(entry.langSecRef)
        confirmOnlyOne(entry.sectionKey)
        'Validate LangSecRef numbers
        For Each key In entry.langSecRef
            fullKeyErr(key, "LangSecRef holds an invalid value.", Not secRefNums.IsMatch(key.value))
        Next
    End Sub

    ''' <summary>
    ''' Processes a list of FileKEy format winapp2.ini iniKey objects and checks them for errors, correcting where possible
    ''' </summary>
    ''' <param name="keyList"></param>
    Private Sub pFileKey(ByRef keyList As List(Of iniKey))
        If keyList.Count = 0 Then Exit Sub
        initCFormat(keyList)
        For Each key In keyList
            'Get the parameters given to the file key and sort them 
            Dim keyParams As New winapp2KeyParameters(key)
            Dim argsStrings As New List(Of String)
            Dim dupeArgs As New List(Of String)
            'check for duplicate args
            For Each arg In keyParams.argsList
                If chkDupes(argsStrings, arg) And scanParams Then
                    err(key.lineNumber, "Duplicate FileKey parameter found", arg)
                    dupeArgs.Add(arg)
                End If
            Next
            'Remove any duplicate arguments from the key parameters
            For Each arg In dupeArgs
                If fixFormat(correctParameters) Then keyParams.argsList.Remove(arg)
            Next
            'Reconstruct keys we've modified above
            If fixFormat(correctParameters) Then keyParams.reconstructKey(key)
            'Pipe symbol checks
            Dim iteratorCheckerList() As String = Split(key.value, "|")
            fullKeyErr(key, "Missing pipe (|) in FileKey.", Not key.vHas("|"))
            'Captures any incident of semi colons coming before the first pipe symbol
            fullKeyErr(key, "Semicolon (;) found before pipe (|).", key.vHas(";") And key.value.IndexOf(";") < key.value.IndexOf("|"))
            'Check for incorrect spellings of RECURSE or REMOVESELF
            If iteratorCheckerList.Length > 2 Then fullKeyErr(key, "RECURSE or REMOVESELF is incorrectly spelled, or there are too many pipe (|) symbols.", Not iteratorCheckerList(2).Contains("RECURSE") And Not iteratorCheckerList(2).Contains("REMOVESELF"))
            'Check for missing pipe symbol on recurse and removeself, fix them if detected
            cFlags(key, {"RECURSE", "REMOVESELF"})
            'Make sure VirtualStore folders point to the correct place
            inputMismatchErr(key.lineNumber, "Incorrect VirtualStore location.", key.value, "%LocalAppData%\VirtualStore\Program Files*\", key.vHas("\virtualStore\p", True) And Not key.vHasAny({"programdata", "program files*", "program*"}, True))
            'Backslash checks, fix if detected
            fullKeyErr(key, "Backslash (\) found before pipe (|).", scanSlashes And key.vHas("%\|"), correctSlashes, key.value, key.value.Replace("%\|", "%|"))
            fullKeyErr(key, "Missing backslash (\) after %EnvironmentVariable%.", key.vHas("%") And Not key.vHasAny({"%|", "%\"}))
        Next
        'Optimization Check
        cOptimization(keyList)
    End Sub

    ''' <summary>
    ''' Processes the Default key(s - there should only be one!) in an entry
    ''' </summary>
    ''' <param name="keyList">The list of Default keys in an winapp2entry</param>
    Private Sub pDefault(keyList As List(Of iniKey))
        'Make sure we only have one Default.
        confirmOnlyOne(keyList)
        'Make sure all entries are disabled by Default.
        For Each key In keyList
            fullKeyErr(key, "Unnecessary numbering detected.", scanNumbers And Not key.nameIs("Default"), correctNumbers, key.name, "Default")
            fullKeyErr(key, "All entries should be disabled by default (Default=False).", scanDefaults And Not key.vIs("False"), correctDefaults, key.value, "False")
        Next
    End Sub

    ''' <summary>
    ''' Prints an error for each object in a list of iniKey objects with a length greater than 1
    ''' </summary>
    ''' <param name="keylist">A list of iniKey objects</param>
    Private Sub confirmOnlyOne(keylist As List(Of iniKey))
        If keylist.Count > 1 Then
            For i As Integer = 1 To keylist.Count - 1
                fullKeyErr(keylist(i), $"Multiple {keylist(i).keyType} detected.")
            Next
        End If
    End Sub

    ''' <summary>
    ''' Checks the validity of the keys in an entry and removes any that are too problematic to continue with
    ''' </summary>
    ''' <param name="entry">A winapp2entry object to be audited</param>
    Private Sub validateKeys(ByRef entry As winapp2entry)
        For Each lst In entry.keyListList
            Dim brokenKeys As New List(Of iniKey)
            For Each key In lst
                addKeyToListIf(key, brokenKeys, Not cValidity(key))
            Next
            removeDuplicateKeys(lst, brokenKeys)
            removeDuplicateKeys(entry.errorKeys, brokenKeys)
        Next
    End Sub

    ''' <summary>
    ''' Processes a winapp2entry object (generated from a winapp2.ini format iniKey object) for errors 
    ''' </summary>
    ''' <param name="entry">The winapp2entry object to be processed</param>
    Private Sub processEntry(ByRef entry As winapp2entry)
        Dim entryLinesList As New List(Of Integer)
        Dim hasFileExcludes As Boolean = False
        Dim hasRegExcludes As Boolean = False
        'Check for duplicate names that are differently cased 
        fullNameErrIf(chkDupes(allEntryNames, entry.name), entry, "Duplicate entry name detected")
        'Check that the entry is named properly 
        fullNameErrIf(Not entry.name.EndsWith(" *"), entry, "All entries must End In ' *'")
        'Confirm the validity of keys and remove any broken ones before continuing
        validateKeys(entry)
        'Proess individual keys/keylists in winapp2.ini order
        pDetOS(entry.detectOS)
        pLangSecRef(entry)
        pSpecialDetect(entry.specialDetect)
        pDetect(entry.detects)
        pDetect(entry.detectFiles)
        pDefault(entry.defaultKey)
        pWarning(entry.warningKey)
        pFileKey(entry.fileKeys)
        pRegKey(entry.regKeys)
        pExcludeKey(entry.excludeKeys, hasFileExcludes, hasRegExcludes)
        'Make sure we have at least 1 valid detect key 
        fullNameErrIf(entry.detectOS.Count + entry.detects.Count + entry.specialDetect.Count + entry.detectFiles.Count = 0, entry, "Entry has no valid detection keys (Detect, DetectFile, DetectOS, SpecialDetect)")
        'Make sure we have at least 1 FileKey or RegKey
        fullNameErrIf(entry.fileKeys.Count + entry.regKeys.Count = 0, entry, "Entry has no valid FileKeys or RegKeys")
        'If we don't have FileKeys or RegKeys, we shouldn't have ExcludeKeys.
        fullNameErrIf(entry.excludeKeys.Count > 0 And entry.fileKeys.Count + entry.regKeys.Count = 0, entry, "Entry has ExcludeKeys but no valid FileKeys or RegKeys")
        'Make sure that if we have file Excludes, we have FileKeys
        fullNameErrIf(entry.fileKeys.Count = 0 And hasFileExcludes, entry, "ExcludeKeys targeting filesystem locations found without any corresponding FileKeys")
        'Likewise for registry
        fullNameErrIf(entry.regKeys.Count = 0 And hasRegExcludes, entry, "ExcludeKeys targeting registry locations found without any corresponding RegKeys")
        'Make sure we have a Default key.
        fullNameErrIf(entry.defaultKey.Count = 0 And scanDefaults, entry, "Entry is missing a Default key")
        addKeyToListIf(New iniKey("Default=False", 0), entry.defaultKey, fixFormat(correctDefaults) And entry.defaultKey.Count = 0)
    End Sub

    ''' <summary>
    ''' Processes a list of DetectOS format winapp2.ini iniKey objects and checks them for errors, correcting where possible
    ''' </summary>
    ''' <param name="keyList">A list of winapp2.ini DetectOS format iniKey objects to be processed</param>
    Private Sub pDetOS(ByRef keyList As List(Of iniKey))
        'Make sure we have only one DetectOS
        confirmOnlyOne(keyList)
    End Sub

    ''' <summary>
    ''' Processes a list of Warning format winapp2.ini iniKey objects and checks them for errors, correcting where possible
    ''' </summary>
    ''' <param name="keyList">A list of winapp2.ini Warning format iniKeys to be scanned</param>
    Private Sub pWarning(ByRef keyList As List(Of iniKey))
        'Make sure we have only one warning
        confirmOnlyOne(keyList)
    End Sub

    ''' <summary>
    ''' Processes a list of DetectFile format winapp2.ini iniKey objects and checks them for errors, correcting where possible
    ''' </summary>
    ''' <param name="keyList">A list of winapp2.ini DetectFile format iniKeys to be scanned</param>
    Private Sub pDetectFile(ByRef keyList As List(Of iniKey))
        For Each key In keyList
            'Trailing Backslashes
            fullKeyErr(key, "Trailing backslash (\) found in DetectFile", scanSlashes And key.value.Last = CChar("\"), correctSlashes, key.value, key.value.TrimEnd(CChar("\")))
            'Nested wildcards
            If key.vHas("*") Then
                Dim splitDir As String() = key.value.Split(CChar("\"))
                For i As Integer = 0 To splitDir.Count - 1
                    fullKeyErr(key, "Nested wildcard found in DetectFile", splitDir(i).Contains("*") And i <> splitDir.Count - 1)
                Next
            End If
            'Make sure that DetectFile paths point to a filesystem location
            chkPathFormatValidity(key, False)
        Next
    End Sub

    ''' <summary>
    ''' Audits the syntax of file system and registry paths
    ''' </summary>
    ''' <param name="key">An iniKey to be audited</param>
    ''' <param name="isRegistry">Specifies whether the given key holds a registry path</param>
    Private Sub chkPathFormatValidity(key As iniKey, isRegistry As Boolean)
        'Remove the flags from ExcludeKeys if we have them before getting the first directory portion
        Dim rootStr As String = If(key.keyType <> "ExcludeKey", getFirstDir(key.value), getFirstDir(pathFromExcludeKey(key)))
        'Ensure that registry paths have a valid hive and file paths have either a variable or a drive letter
        fullKeyErr(key, "Invalid registry path detected.", isRegistry And Not longReg.IsMatch(rootStr) And Not shortReg.IsMatch(rootStr))
        fullKeyErr(key, "Invalid file system path detected.", Not isRegistry And Not driveLtrs.IsMatch(rootStr) And Not rootStr.StartsWith("%"))
    End Sub

    ''' <summary>
    ''' Returns the first portion of a registry or filepath parameterization
    ''' </summary>
    ''' <param name="val">The directory listing to be split</param>
    ''' <returns></returns>
    Public Function getFirstDir(val As String) As String
        Return val.Split(CChar("\"))(0)
    End Function

    ''' <summary>
    '''Processes a list of Detect format winapp2.ini iniKey objects and checks them for errors, correcting where possible
    ''' </summary>
    ''' <param name="keylist">A list of iniKey objects containing RegKey format winapp2.ini keys</param>
    Private Sub pRegDetect(keylist As List(Of iniKey))
        'Ensure each Detect points to a valid registry hive
        For Each key In keylist
            chkPathFormatValidity(key, True)
        Next
    End Sub

    ''' <summary>
    ''' Checks whether the current value appears in the given list of strings (case insensitive). Returns true if there is a duplicate,
    ''' otherwise, adds the current value to the list.
    ''' </summary>
    ''' <param name="valueStrings">A list of strings holding observed values</param>
    ''' <param name="currentValue">The current value to be audited</param>
    ''' <returns></returns>
    Private Function chkDupes(ByRef valueStrings As List(Of String), currentValue As String) As Boolean
        If valueStrings.Contains(currentValue.ToLower) Then
            Return True
        Else
            valueStrings.Add(currentValue.ToLower)
            Return False
        End If
    End Function

    ''' <summary>
    ''' Audits the numbering of keys in their name from 1 to infinity and returns any duplicate keys detected back to the calling function
    ''' </summary>
    ''' <param name="keyStrings">The list of values from the iniKeys being audited</param>
    ''' <param name="key">The current key being audited</param>
    ''' <param name="keyNumber">The current expected key number</param>
    ''' <param name="dupeList">The list containing any duplicate iniKey objects found during the audit</param>
    Private Sub checkDupsAndNumbering(ByRef keyStrings As List(Of String), ByRef key As iniKey, ByRef keyNumber As Integer, ByRef dupeList As List(Of iniKey))
        'Check for duplicates
        If chkDupes(keyStrings, key.value) Then
            customErr(key.lineNumber, "Duplicate key value found", {$"Key:            {key.toString}", $"Duplicates:     {key.keyType}{keyStrings.IndexOf(key.value.ToLower) + 1}={key.value}"})
            dupeList.Add(key)
        End If
        'Make sure the current key is correctly numbered
        inputMismatchErr(key.lineNumber, $"{key.keyType} entry is incorrectly numbered.", key.name, $"{key.keyType}{keyNumber}", scanNumbers And Not key.nameIs(key.keyType & keyNumber))
        fixStr(scanNumbers And correctNumbers And Not key.nameIs(key.keyType & keyNumber), key.name, key.keyType & keyNumber)
        keyNumber += 1
    End Sub

    ''' <summary>
    ''' Processes a list of DetectFile or Detect format winapp2.ini iniKey objects and checks them for errors, correcting where possible
    ''' </summary>
    ''' <param name="keyList">The list of iniKeys in winapp2 Detect or DetectFile format</param>
    Private Sub pDetect(ByRef keyList As List(Of iniKey))
        If keyList.Count = 0 Then Exit Sub
        'Send off Detect/Files for their specific checks 
        If keyList(0).typeIs("Detect") Then
            pRegDetect(keyList)
        Else
            pDetectFile(keyList)
        End If
        'Formatting and duplicate checks
        If keyList.Count > 1 Then
            initCFormat(keyList)
        Else
            'Make sure that if we have only one detect/file, it doesn't have a number
            inputMismatchErr(keyList(0).lineNumber, "Detected unnecessary numbering.", keyList(0).name, keyList(0).keyType, keyList.Count = 1 And Not keyList(0).nameIs(keyList(0).keyType) And scanNumbers)
            fixStr(scanNumbers And correctNumbers And keyList.Count = 1 And Not keyList(0).nameIs(keyList(0).keyType), keyList(0).name, keyList(0).keyType)
            'Check the env vars iff we're operating on a DetectFile
            If keyList.First.keyType = "DetectFile" Then cEnVar(keyList.First)
        End If
    End Sub

    ''' <summary>
    ''' Sorts a list of iniKey objects alphabetically (with some changes made for winapp2.ini syntax) based on the contents of their value field
    ''' </summary>
    ''' <param name="keyList">The list of iniKey objects to be sorted</param>
    Private Sub sortKeys(ByRef keyList As List(Of iniKey))
        If (Not scanAlpha) Or keyList.Count <= 1 Then Exit Sub
        Dim keyStrings As List(Of String) = getValues(keyList)
        Dim sortedKeyList As List(Of String) = replaceAndSort(keyStrings, "|", " \ \")
        'Rewrite the alphabetized keys back into the keylist (fixes numbering silently) 
        Dim keysOutOfPlace As Boolean = False
        findOutOfPlace(keyStrings, sortedKeyList, keyList(0).keyType, getLineNumsFromKeyList(keyList), keysOutOfPlace)
        If keysOutOfPlace Then renumberKeys(keyList, sortedKeyList)
    End Sub

    ''' <summary>
    ''' Rewrites the iniKey data in a given list of iniKeys to be numerically ordered and sorted alphabetically
    ''' </summary>
    ''' <param name="keyList">The list of iniKey objects to reorder</param>
    ''' <param name="sortedKeyList">The sorted state of the key values in the list</param>
    Private Sub renumberKeys(ByRef keyList As List(Of iniKey), sortedKeyList As List(Of String))
        If fixFormat(correctAlpha) Then
            Dim i As Integer = 1
            For Each key In keyList
                key.name = key.keyType & i
                key.value = sortedKeyList(i - 1)
                i += 1
            Next
        End If
    End Sub

    ''' <summary>
    ''' Processes a list of RegKey format winapp2.ini iniKey objects and checks them for errors, correcting where possible
    ''' </summary>
    ''' <param name="keyList">A list of winapp2.ini format RegKey iniKey objects</param>
    Private Sub pRegKey(ByVal keyList As List(Of iniKey))
        'Formatting checks
        initCFormat(keyList)
        For Each key In keyList
            'Path format validation
            chkPathFormatValidity(key, True)
        Next
    End Sub

    ''' <summary>
    ''' Detects if a pipe symbol is missing before an array of given params
    ''' </summary>
    ''' <param name="key">An iniKey object to be observed</param>
    ''' <param name="flagStrs">The array of parameters which should be prceeded by a pipe symbol if they exist in the key value</param>
    Private Sub cFlags(ByRef key As iniKey, flagStrs As String())
        For Each flagstr In flagStrs
            fullKeyErr(key, $"Missing pipe (|) before {flagstr}.", scanFlags And key.vHas(flagstr) And Not key.vHas($"|{flagstr}"), correctFlags, key.value, key.value.Replace(flagstr, $"|{flagstr}"))
        Next
    End Sub

    ''' <summary>
    ''' Processes a list of SpecialDetect format winapp2.ini iniKey objects and checks them for errors, correcting where possible
    ''' </summary>
    ''' <param name="keyList">A list of SpecialDetect format winapp2.ini iniKey objects</param>
    Private Sub pSpecialDetect(ByRef keyList As List(Of iniKey))
        'Make sure we have at most 1 SpecialDetect.
        confirmOnlyOne(keyList)
        'Make sure that any SpecialDetect keys hold a valid value
        For Each key In keyList
            'Make sure SpecialDetect is not followed by a number
            fullKeyErr(key, "Unnecessary numbering detected.", scanNumbers And Not key.nameIs("SpecialDetect"), correctNumbers, key.name, "SpecialDetect")
            'Confirm that the key holds a valid value
            fullKeyErr(key, "SpecialDetect holds an invalid value.", Not sdList.Contains(key.value))
        Next
    End Sub

    ''' <summary>
    ''' Processes a list of ExcludeKey format winapp2.ini iniKey objects and checks them for errors, correcting where possible
    ''' </summary>
    ''' <param name="keyList">A list of ExcludeKey format winapp2.ini iniKey objects</param>
    ''' <param name="hasF">Indicates whether the given list excludes filesystem locations</param>
    ''' <param name="hasR">Indicates whether the given list excludes registry locations</param>
    Private Sub pExcludeKey(ByRef keyList As List(Of iniKey), ByRef hasF As Boolean, ByRef hasR As Boolean)
        If keyList.Count = 0 Then Exit Sub
        initCFormat(keyList)
        For Each key In keyList
            Select Case True
                Case key.vHasAny({"FILE|", "PATH|"})
                    hasF = True
                    chkPathFormatValidity(key, False)
                    'Make sure any filesystem exclude paths have a backslash before their pipe symbol.
                    fullKeyErr(key, "Missing backslash (\) before pipe (|) in ExcludeKey.", (key.vHas("|") And Not key.vHas("\|")))
                Case key.vHas("REG|")
                    hasR = True
                    chkPathFormatValidity(key, True)
                Case Else
                    'If we made it here, we don't have a valid flag
                    fullKeyErr(key, "No valid exclude flag (FILE, PATH, or REG) found in ExcludeKey.")
            End Select
        Next
    End Sub

    ''' <summary>
    ''' Returns the value from an ExcludeKey with the Flag parameter removed as a String
    ''' </summary>
    ''' <param name="key">An ExcludeKey iniKey</param>
    ''' <returns></returns>
    Private Function pathFromExcludeKey(key As iniKey) As String
        Dim pathFromKey As String = key.value.TrimStart(CType("FILE|", Char()))
        pathFromKey = pathFromKey.TrimStart(CType("PATH|", Char()))
        pathFromKey = pathFromKey.TrimStart(CType("REG|", Char()))
        Return pathFromKey
    End Function

    ''' <summary>
    ''' Attempts to merge FileKeys together if syntactically possible.
    ''' </summary>
    ''' <param name="keyList">A list of FileKeys</param>
    Private Sub cOptimization(ByRef keyList As List(Of iniKey))
        If keyList.Count < 2 Or Not scanOpti Then Exit Sub
        Dim dupeList As New List(Of iniKey)
        Dim flagList As New List(Of String)
        Dim paramList As New List(Of String)
        Dim newKeyList As New List(Of iniKey)
        newKeyList.AddRange(keyList)
        For i As Integer = 0 To keyList.Count - 1
            Dim key As iniKey = keyList(i)
            Dim tmpWa2 As New winapp2KeyParameters(key)
            'If we have yet to record any params, record them and move on
            If paramList.Count = 0 Then
                trackParamAndFlags(paramList, flagList, tmpWa2)
                Continue For
            End If
            'This should handle the case where for a FileKey: 
            'The folder provided has appeared in another key
            'The flagstring (RECURSE, REMOVESELF, "") for both keys matches
            'The first appearing key should have its parameters appended to and the second appearing key should be removed
            If paramList.Contains(tmpWa2.paramString) Then
                For j As Integer = 0 To paramList.Count - 1
                    If tmpWa2.paramString = paramList(j) And tmpWa2.flagString = flagList(j) Then
                        Dim keyToMergeInto As New winapp2KeyParameters(keyList(j))
                        Dim mergeKeyStr As String = ""
                        addArgs(mergeKeyStr, keyToMergeInto)
                        For Each arg In tmpWa2.argsList
                            mergeKeyStr += $";{arg}"
                        Next
                        If tmpWa2.flagString <> "None" Then mergeKeyStr += $"|{tmpWa2.flagString}"
                        dupeList.Add(keyList(i))
                        newKeyList(j) = New iniKey(mergeKeyStr, 1)
                        Exit For
                    End If
                Next
                trackParamAndFlags(paramList, flagList, tmpWa2)
            Else
                trackParamAndFlags(paramList, flagList, tmpWa2)
            End If
        Next
        If dupeList.Count > 0 Then
            removeDuplicateKeys(newKeyList, dupeList)
            For i As Integer = 0 To newKeyList.Count - 1
                newKeyList(i).name = $"FileKey{i + 1}"
            Next
            printOptiSect("Optmization opportunity detected", keyList)
            printOptiSect("The following keys can be merged into other keys:", dupeList)
            printOptiSect("The resulting keyList will be reduced to: ", newKeyList)
            If correctOpti Then keyList = newKeyList
        End If
    End Sub

    ''' <summary>
    ''' Prints output from the Optimization function
    ''' </summary>
    ''' <param name="boxStr">The text to go in the optimization section box</param>
    ''' <param name="keyList">The list of keys to be printed beneath the box</param>
    Private Sub printOptiSect(boxStr As String, keyList As List(Of iniKey))
        cwl()
        print(0, tmenu(boxStr), closeMenu:=True)
        cwl()
        For Each key In keyList
            cwl(key.toString)
        Next
        cwl()
    End Sub

    ''' <summary>
    ''' Tracks params and flags from a winapp2key
    ''' </summary>
    ''' <param name="paramList">The list of params observed</param>
    ''' <param name="flagList">The list of flags observed</param>
    ''' <param name="params">The current set of params and flags</param>
    Private Sub trackParamAndFlags(ByRef paramList As List(Of String), ByRef flagList As List(Of String), params As winapp2KeyParameters)
        paramList.Add(params.paramString)
        flagList.Add(params.flagString)
    End Sub

    ''' <summary>
    ''' Constructs a new iniKey in an attempt to merge keys together
    ''' </summary>
    ''' <param name="tmpKeyStr">The string to contain the new key text</param>
    ''' <param name="tmp2wa2">A set of parameters to append</param>
    Private Sub addArgs(ByRef tmpKeyStr As String, tmp2wa2 As winapp2KeyParameters)
        appendStrs({$"{tmp2wa2.keyType}{tmp2wa2.keyNum}=", $"{tmp2wa2.paramString}|", tmp2wa2.argsList(0)}, tmpKeyStr)
        If tmp2wa2.argsList.Count > 1 Then
            For i As Integer = 1 To tmp2wa2.argsList.Count - 1
                tmpKeyStr += $";{tmp2wa2.argsList(i)}"
            Next
        End If
    End Sub

    ''' <summary>
    ''' Removes any occurance of items in a list of iniKey objects from a given list of iniKey objects
    ''' </summary>
    ''' <param name="keylist">The list from which objects might be removed</param>
    ''' <param name="dupeList">The list of objects to remove</param>
    Private Sub removeDuplicateKeys(ByRef keylist As List(Of iniKey), ByVal dupeList As List(Of iniKey))
        For Each key In dupeList
            keylist.Remove(key)
        Next
    End Sub

    ''' <summary>
    ''' Removes duplicate keys from a list of iniKeys and sorts the result
    ''' </summary>
    ''' <param name="keyList">A list of iniKey objects to be culled and sorted</param>
    ''' <param name="dupeList">A list of duplicate iniKey objects to be culled</param>
    Private Sub deDupeAndSort(ByRef keyList As List(Of iniKey), dupeList As List(Of iniKey))
        'Remove any duplicates
        removeDuplicateKeys(keyList, dupeList)
        'Sort the keys
        sortKeys(keyList)
    End Sub

    ''' <summary>
    ''' Prints an error when data is received that does not match an expected value.
    ''' </summary>
    ''' <param name="linecount">The line number of the error</param>
    ''' <param name="err">The string containing the output error text</param>
    ''' <param name="received">The (erroneous) input data</param>
    ''' <param name="expected">The expected data</param>
    Private Sub inputMismatchErr(linecount As Integer, err As String, received As String, expected As String, Optional cond As Boolean = True)
        If cond Then customErr(linecount, err, {$"Expected: {expected}", $"Found: {received}"})
    End Sub

    ''' <summary>
    ''' Prints an error when invalid data is received.
    ''' </summary>
    ''' <param name="linecount">The line number of the error</param>
    ''' <param name="errTxt">The string containing the output error text</param>
    ''' <param name="received">The (erroneous) input data</param>
    Private Sub err(linecount As Integer, errTxt As String, received As String)
        customErr(linecount, errTxt, {$"Command: {received}"})
    End Sub

    ''' <summary>
    ''' Prints an error followed by the [Full Name *] of the entry to which it belongs
    ''' </summary>
    ''' <param name="cond">The condition under which to print</param>
    ''' <param name="entry">The entry containing an error</param>
    ''' <param name="errTxt">The String containing the text to be printed to the user</param>
    Private Sub fullNameErrIf(cond As Boolean, entry As winapp2entry, errTxt As String)
        If cond Then customErr(entry.lineNum, errTxt, {$"Entry Name: {entry.fullname}"})
    End Sub

    ''' <summary>
    ''' Prints an error whose output text contains an iniKey string
    ''' </summary>
    ''' <param name="key">The inikey to be printed</param>
    ''' <param name="err">The string containing the output error text</param>
    Private Sub fullKeyErr(key As iniKey, err As String, Optional cond As Boolean = True, Optional repCond As Boolean = False, Optional ByRef repairVal As String = "", Optional newVal As String = "")
        If cond Then customErr(key.lineNumber, err, {$"Key: {key.toString}"})
        fixStr(cond And repCond, repairVal, newVal)
    End Sub

    ''' <summary>
    ''' Prints arbitrarily defined errors without a precondition
    ''' </summary>
    ''' <param name="lineCount"></param>
    ''' <param name="err"></param>
    ''' <param name="lines"></param>
    Private Sub customErr(lineCount As Integer, err As String, lines As String())
        cwl($"Line: {lineCount} - Error: {err}")
        For Each errStr In lines
            cwl(errStr)
        Next
        cwl()
        numErrs += 1
    End Sub

    ''' <summary>
    ''' Replace a given string with a new value if the fix condition is met. 
    ''' </summary>
    ''' <param name="param">The condition under which the string should be replaced</param>
    ''' <param name="currentValue">The current value of the given string</param>
    ''' <param name="newValue">The replacement value for the given string</param>
    Private Sub fixStr(param As Boolean, ByRef currentValue As String, newValue As String)
        If param Then currentValue = newValue
    End Sub
End Module