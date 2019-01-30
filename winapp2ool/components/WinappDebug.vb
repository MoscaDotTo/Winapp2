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
Imports System.IO
Imports System.Text.RegularExpressions

''' <summary>
''' A program whose purpose is to observe, report, and attempt to repair errors in winapp2.ini
''' </summary>
Module WinappDebug
    ' File handlers
    Dim winappFile As iniFile = New iniFile(Environment.CurrentDirectory, "winapp2.ini")
    Dim outputFile As iniFile = New iniFile(Environment.CurrentDirectory, "winapp2.ini", "winapp2-debugged.ini")
    ' Menu settings
    Dim settingsChanged As Boolean = False
    Dim scanSettingsChanged As Boolean = False
    ' Module parameters
    Dim correctFormatting As Boolean = False
    Dim allEntryNames As New List(Of String)
    Dim numErrs As Integer = 0
    Dim correctSomeFormatting As Boolean = True
    ' Autocorrect Parameters
    Dim correctCasing As Boolean = True
    Dim correctAlpha As Boolean = True
    Dim correctNumbers As Boolean = True
    Dim correctParameters As Boolean = True
    Dim correctFlags As Boolean = True
    Dim correctSlashes As Boolean = True
    Dim correctDefaults As Boolean = True
    Dim correctOpti As Boolean = False
    ' Scan parameters
    Dim scanCasing As Boolean = True
    Dim scanAlpha As Boolean = True
    Dim scanNumbers As Boolean = True
    Dim scanParams As Boolean = True
    Dim scanFlags As Boolean = True
    Dim scanSlashes As Boolean = True
    Dim scanDefaults As Boolean = True
    Dim scanOpti As Boolean = False
    ' Winapp2 Parameters
    ReadOnly enVars As String() = {"UserProfile", "UserName", "ProgramFiles", "RootDir", "WinDir", "AppData", "SystemDrive", "SystemRoot", "Documents", "ProgramData", "AllUsersProfile", "Pictures", "Video", "CommonAppData", "LocalAppData", "CommonProgramFiles", "HomeDrive", "Music", "tmp", "Temp", "LocalLowAppData", "Public"}
    ReadOnly validCmds As String() = {"SpecialDetect", "FileKey", "RegKey", "Detect", "LangSecRef", "Warning", "Default", "Section", "ExcludeKey", "DetectFile", "DetectOS"}
    ReadOnly sdList As String() = {"DET_CHROME", "DET_MOZILLA", "DET_THUNDERBIRD", "DET_OPERA"}
    ' Regex strings 
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
    ''' Handles the commandline args for WinappDebug
    ''' </summary>
    ''' WinappDebug specific command line args
    ''' -c          : enable autocorrect
    Public Sub handleCmdLine()
        invertSettingAndRemoveArg(correctFormatting, "-c")
        getFileAndDirParams(winappFile, New iniFile, outputFile)
        initDebug()
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
    ''' Matches paired boolean states for scans and their repairs
    ''' </summary>
    ''' ie. make sure that if a scan is disabled, so too is its repair, and if a repair is enabled, so too is its scan
    ''' <param name="setting"></param>
    ''' <param name="pairedSetting"></param>
    ''' <param name="type"></param>
    Private Sub toggleScanSetting(ByRef setting As Boolean, ByRef pairedSetting As Boolean, type As String)
        toggleSettingParam(setting, type, scanSettingsChanged)
        If (type = "Scan" And Not setting And pairedSetting) Or (type = "Repair" And setting And Not pairedSetting) Then toggleSettingParam(pairedSetting, type, scanSettingsChanged)
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
        print(1, "Toggle Autocorrect", $"{enStr(correctFormatting)} saving the file after correcting errors")
        print(1, "File Chooser (save)", "Save a copy of changes made instead of overwriting winapp2.ini", correctFormatting, trailingBlank:=True)
        print(1, "Toggle Scan Settings", "Enable or disable individual scan and correction routines", leadingBlank:=Not correctFormatting, trailingBlank:=True)
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
                exitModule()
            Case input = "1" Or input = ""
                initDebug()
            Case input = "2"
                changeFileParams(winappFile, settingsChanged)
            Case input = "3"
                toggleSettingParam(correctFormatting, "Autocorrect", settingsChanged)
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
        If pendingExit() Then Exit Sub
        Console.Clear()
        print(0, tmenu("Beginning analysis of winapp2.ini"), closeMenu:=True)
        cwl()
        numErrs = 0
        allEntryNames = New List(Of String)
        Dim winapp2file As New winapp2file(cfile)
        processEntries(winapp2file)
        winapp2file.rebuildToIniFiles()
        alphabetizeEntries(winapp2file)
        print(0, tmenu("Completed analysis of winapp2.ini"))
        print(0, menuStr03)
        print(0, $"{numErrs} possible errors were detected.")
        print(0, $"Number of entries: {cfile.sections.Count}", trailingBlank:=True)
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
            entryList.ForEach(Sub(entry) processEntry(entry))
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
            findOutOfPlace(unsortedEntryList, sortedEntryList, "Entry", innerFile.getLineNumsFromSections)
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
    ''' <param name="oopBool">Optional Boolean that reports whether or not there are any out of place entries in the list</param>
    Private Sub findOutOfPlace(ByRef someList As List(Of String), ByRef sortedList As List(Of String), ByVal findType As String, ByRef LineCountList As List(Of Integer), Optional ByRef oopBool As Boolean = False)
        'Only try to find out of place keys when there's more than one
        If someList.Count > 1 Then
            Dim misplacedEntries As New List(Of String)
            ' Learn the neighbors of each string in each respective list
            Dim initalNeighbors As New List(Of KeyValuePair(Of String, String))
            Dim sortedNeigbors As New List(Of KeyValuePair(Of String, String))
            buildNeighborList(someList, initalNeighbors)
            buildNeighborList(sortedList, sortedNeigbors)
            ' Make sure at least one of the neighbors of each string are the same in both the sorted and unsorted state, otherwise the string has moved 
            For i As Integer = 0 To someList.Count - 1
                Dim sind As Integer = sortedList.IndexOf(someList(i))
                If Not (initalNeighbors(i).Key = sortedNeigbors(sind).Key And initalNeighbors(i).Value = sortedNeigbors(sind).Value) Then misplacedEntries.Add(someList(i))
            Next
            ' Report any misplaced entries back to the user
            For Each entry In misplacedEntries
                Dim recind As Integer = someList.IndexOf(entry)
                Dim sortind As Integer = sortedList.IndexOf(entry)
                Dim curpos As Integer = LineCountList(recind)
                Dim sortPos As Integer = LineCountList(sortind)
                If (recind = sortind Or curpos = sortPos) Then Continue For
                entry = If(findType = "Entry", entry, $"{findType & recind + 1}={entry}")
                If Not oopBool Then oopBool = True
                customErr(LineCountList(recind), $"{findType} alphabetization", {$"{entry} appears to be out of place", $"Current line: {curpos}", $"Expected line: {sortPos}"})
            Next
        End If
    End Sub

    ''' <summary>
    ''' Audits a keyList of winapp2.ini format iniKeys for errors, alerting the user and correcting where possible.
    ''' </summary>
    ''' <param name="kl">A keylist to audit</param>
    ''' <param name="processKey">The function that audits the keys of the keyType provided in the keyList</param>
    ''' <param name="hasF">Optional boolean for the ExcludeKey case</param>
    ''' <param name="hasR">Optional boolean for the ExcludeKey case</param>
    Private Sub processKeyList(ByRef kl As keyList, processKey As Func(Of iniKey, iniKey), Optional ByRef hasF As Boolean = False, Optional ByRef hasR As Boolean = False)
        If kl.keyCount = 0 Then Exit Sub
        Dim curNum As Integer = 1
        Dim curStrings As New List(Of String)
        Dim dupes As New keyList
        For Each key In kl.keys
            ' Preprocess key formats before doing key specific audits
            Select Case key.keyType
                Case "ExcludeKey"
                    cFormat(key, curNum, curStrings, dupes)
                    pExcludeKey(key, hasF, hasR)
                Case "Detect", "DetectFile"
                    If key.typeIs("Detect") Then chkPathFormatValidity(key, True)
                    cFormat(key, curNum, curStrings, dupes, kl.keyCount = 1)
                Case "RegKey"
                    cFormat(key, curNum, curStrings, dupes)
                    chkPathFormatValidity(key, True)
                Case "Warning", "DetectOS", "SpecialDetect", "LangSecRef", "Section", "Default"
                    If curNum > 1 Then
                        fullKeyErr(key, $"Multiple {key.keyType} detected.")
                        dupes.add(key)
                    End If
                    cFormat(key, curNum, curStrings, dupes, True)
                    fullKeyErr(key, "LangSecRef holds an invalid value.", key.typeIs("LangSecRef") And Not secRefNums.IsMatch(key.value))
                    fullKeyErr(key, "SpecialDetect holds an invalid value.", key.typeIs("SpecialDetect") And Not sdList.Contains(key.value))
                    fullKeyErr(key, "All entries should be disabled by default (Default=False).", scanDefaults And Not key.vIs("False") And key.typeIs("Default"), correctDefaults, key.value, "False")
                Case Else
                    cFormat(key, curNum, curStrings, dupes)
            End Select
            key = processKey(key)
        Next
        kl.remove(dupes.keys)
        sortKeys(kl, dupes.keyCount > 0)
        If kl.typeIs("FileKey") Then cOptimization(kl)
    End Sub

    ''' <summary>
    ''' Does some basic formatting checks that apply to most winapp2.ini format iniKey objects
    ''' </summary>
    ''' <param name="key">The current iniKey object being processed</param>
    ''' <param name="keyNumber">The current expected key number</param>
    ''' <param name="keyValues">The current list of observed inikey values</param>
    ''' <param name="dupeList">A tracking list of detected duplicate iniKeys</param>
    Private Sub cFormat(ByVal key As iniKey, ByRef keyNumber As Integer, ByRef keyValues As List(Of String), ByRef dupeList As keyList, Optional noNumbers As Boolean = False)
        ' Check for duplicates
        If chkDupes(keyValues, key.value) Then
            Dim duplicateKeyStr As String = $"{key.keyType}{If(Not noNumbers, (keyValues.IndexOf(key.value.ToLower) + 1).ToString, "")}={key.value}"
            customErr(key.lineNumber, "Duplicate key value found", {$"Key:            {key.toString}", $"Duplicates:     {duplicateKeyStr}"})
            dupeList.add(key)
        End If
        ' Check for numbering errors
        Dim hasNumberingError As Boolean = If(noNumbers, Not key.nameIs(key.keyType), Not key.nameIs(key.keyType & keyNumber))
        Dim numberingErrStr As String = If(noNumbers, "Detected unnecessary numbering.", $"{key.keyType} entry is incorrectly numbered.")
        Dim fixedStr As String = If(noNumbers, key.keyType, key.keyType & keyNumber)
        inputMismatchErr(key.lineNumber, numberingErrStr, key.name, fixedStr, scanNumbers And hasNumberingError)
        fixStr(correctNumbers And hasNumberingError, key.name, fixedStr)
        ' Scan for and fix any use of incorrect slashes (except in Warning keys) or trailing semicolons
        fullKeyErr(key, "Forward slash (/) detected in lieu of blackslash (\).", Not key.typeIs("Warning") And scanSlashes And key.vHas(CChar("/")), correctSlashes, key.value, key.value.Replace(CChar("/"), CChar("\")))
        fullKeyErr(key, "Trailing semicolon (;).", key.toString.Last = CChar(";") And scanParams, correctParameters, key.value, key.value.TrimEnd(CChar(";")))
        ' Do some formatting checks for environment variables if needed
        If {"FileKey", "ExcludeKey", "DetectFile"}.Contains(key.keyType) Then cEnVar(key)
        keyNumber += 1
    End Sub

    ''' <summary>
    ''' Checks a String for casing errors against a provided array of cased strings, returns the input string if no error is detected
    ''' </summary>
    ''' <param name="caseArray">The parent array of cased Strings</param>
    ''' <param name="keyValue">The String to be checked for casing errors</param>
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
            ' If we don't have a casing error and enVars doesn't contain our value, it's invalid. 
            fullKeyErr(key, $"Misformatted or invalid environment variable:  {m.ToString}", camelText = strippedText And Not enVars.Contains(strippedText))
        Next
    End Sub

    ''' <summary>
    ''' Does basic syntax and formatting audits that apply across all keys, returns false iff a key in malformatted
    ''' </summary>
    ''' <param name="key">an iniKey object to be checked for errors</param>
    ''' <returns></returns>
    Private Function cValidity(key As iniKey) As Boolean
        If key.keyType = "DeleteMe" Then Return False
        ' Check for leading or trailing whitespace
        fullKeyErr(key, "Detected unwanted whitespace in iniKey value", key.vStartsOrEndsWith(" "), True, key.value, key.value.Trim(CChar(" ")))
        fullKeyErr(key, "Detected unwanted whitespace in iniKey name", key.nStartsOrEndsWith(" "), True, key.name, key.name.Trim(CChar(" ")))
        ' Make sure the keyType is a valid winapp2.ini command
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
    ''' Processes a FileKey format winapp2.ini iniKey object and checks it for errors, correcting where possible
    ''' </summary>
    ''' <param name="key">A winap2.ini FileKey format iniKey object</param>
    Private Function pFileKey(key As iniKey) As iniKey
        ' Pipe symbol checks
        Dim iteratorCheckerList() As String = Split(key.value, "|")
        fullKeyErr(key, "Missing pipe (|) in FileKey.", Not key.vHas("|"))
        ' Captures any incident of semi colons coming before the first pipe symbol
        fullKeyErr(key, "Semicolon (;) found before pipe (|).", key.vHas(";") And key.value.IndexOf(";") < key.value.IndexOf("|"))
        ' Check for incorrect spellings of RECURSE or REMOVESELF
        If iteratorCheckerList.Length > 2 Then fullKeyErr(key, "RECURSE or REMOVESELF is incorrectly spelled, or there are too many pipe (|) symbols.", Not iteratorCheckerList(2).Contains("RECURSE") And Not iteratorCheckerList(2).Contains("REMOVESELF"))
        ' Check for missing pipe symbol on recurse and removeself, fix them if detected
        Dim flags As New List(Of String) From {"RECURSE", "REMOVESELF"}
        flags.ForEach(Sub(flagStr) fullKeyErr(key, $"Missing pipe (|) before {flagStr}.", scanFlags And key.vHas(flagStr) And Not key.vHas($"|{flagStr}"), correctFlags, key.value, key.value.Replace(flagStr, $"|{flagStr}")))
        ' Make sure VirtualStore folders point to the correct place
        inputMismatchErr(key.lineNumber, "Incorrect VirtualStore location.", key.value, "%LocalAppData%\VirtualStore\Program Files*\", key.vHas("\virtualStore\p", True) And Not key.vHasAny({"programdata", "program files*", "program*"}, True))
        ' Backslash checks, fix if detected
        fullKeyErr(key, "Backslash (\) found before pipe (|).", scanSlashes And key.vHas("%\|"), correctSlashes, key.value, key.value.Replace("%\|", "%|"))
        fullKeyErr(key, "Missing backslash (\) after %EnvironmentVariable%.", key.vHas("%") And Not key.vHasAny({"%|", "%\"}))
        ' Get the parameters given to the file key and sort them 
        Dim keyParams As New winapp2KeyParameters(key)
        Dim argsStrings As New List(Of String)
        Dim dupeArgs As New List(Of String)
        ' Check for duplicate args
        For Each arg In keyParams.argsList
            If chkDupes(argsStrings, arg) And scanParams Then
                err(key.lineNumber, "Duplicate FileKey parameter found", arg)
                dupeArgs.Add(arg)
            End If
        Next
        ' Remove any duplicate arguments from the key parameters
        dupeArgs.ForEach(Sub(arg) If fixFormat(correctParameters) Then keyParams.argsList.Remove(arg))
        ' Reconstruct keys we've modified above
        If fixFormat(correctParameters) Then keyParams.reconstructKey(key)
        Return key
    End Function

    ''' <summary>
    ''' Checks the validity of the keys in an entry and removes any that are too problematic to continue with
    ''' </summary>
    ''' <param name="entry">A winapp2entry object to be audited</param>
    Private Sub validateKeys(ByRef entry As winapp2entry)
        For Each lst In entry.keyListList
            Dim brokenKeys As New keyList
            lst.keys.ForEach(Sub(key) brokenKeys.add(key, Not cValidity(key)))
            lst.remove(brokenKeys.keys)
            entry.errorKeys.remove(brokenKeys.keys)
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
        ' Check for duplicate names that are differently cased 
        fullNameErrIf(chkDupes(allEntryNames, entry.name), entry, "Duplicate entry name detected")
        ' Check that the entry is named properly 
        fullNameErrIf(Not entry.name.EndsWith(" *"), entry, "All entries must End In ' *'")
        ' Confirm the validity of keys and remove any broken ones before continuing
        validateKeys(entry)
        ' Proess individual keys/keylists in winapp2.ini order
        processKeyList(entry.detectOS, AddressOf voidDelegate)
        processKeyList(entry.sectionKey, AddressOf voidDelegate)
        processKeyList(entry.langSecRef, AddressOf voidDelegate)
        processKeyList(entry.specialDetect, AddressOf voidDelegate)
        processKeyList(entry.detects, AddressOf voidDelegate)
        processKeyList(entry.detectFiles, AddressOf pDetectFile)
        processKeyList(entry.defaultKey, AddressOf voidDelegate)
        processKeyList(entry.warningKey, AddressOf voidDelegate)
        processKeyList(entry.fileKeys, AddressOf pFileKey)
        processKeyList(entry.regKeys, AddressOf voidDelegate)
        processKeyList(entry.excludeKeys, AddressOf voidDelegate, hasFileExcludes, hasRegExcludes)
        ' Make sure we only have LangSecRef if we have LangSecRef at all
        fullNameErrIf(entry.langSecRef.keyCount <> 0 And entry.sectionKey.keyCount <> 0, entry, "Section key found alongside LangSecRef key, only one is required.")
        ' Make sure we have at least 1 valid detect key 
        fullNameErrIf(entry.detectOS.keyCount + entry.detects.keyCount + entry.specialDetect.keyCount + entry.detectFiles.keyCount = 0, entry, "Entry has no valid detection keys (Detect, DetectFile, DetectOS, SpecialDetect)")
        ' Make sure we have at least 1 FileKey or RegKey
        fullNameErrIf(entry.fileKeys.keyCount + entry.regKeys.keyCount = 0, entry, "Entry has no valid FileKeys or RegKeys")
        ' If we don't have FileKeys or RegKeys, we shouldn't have ExcludeKeys.
        fullNameErrIf(entry.excludeKeys.keyCount > 0 And entry.fileKeys.keyCount + entry.regKeys.keyCount = 0, entry, "Entry has ExcludeKeys but no valid FileKeys or RegKeys")
        ' Make sure that if we have file Excludes, we have FileKeys
        fullNameErrIf(entry.fileKeys.keyCount = 0 And hasFileExcludes, entry, "ExcludeKeys targeting filesystem locations found without any corresponding FileKeys")
        ' Likewise for registry
        fullNameErrIf(entry.regKeys.keyCount = 0 And hasRegExcludes, entry, "ExcludeKeys targeting registry locations found without any corresponding RegKeys")
        ' Make sure we have a Default key.
        fullNameErrIf(entry.defaultKey.keyCount = 0 And scanDefaults, entry, "Entry is missing a Default key")
        entry.defaultKey.add(New iniKey("Default=False"), fixFormat(correctDefaults) And entry.defaultKey.keyCount = 0)
    End Sub

    ''' <summary>
    ''' This method does nothing by design
    ''' </summary>
    ''' <param name="key">An iniKey object to do nothing with</param>
    Private Function voidDelegate(key As iniKey) As iniKey
        Return key
    End Function

    ''' <summary>
    ''' Processes a DetectFile format winapp2.ini iniKey objects and checks it for errors, correcting where possible
    ''' </summary>
    ''' <param name="key">A winapp2.ini DetectFile format iniKey</param>
    ''' <returns></returns>
    Private Function pDetectFile(key As iniKey) As iniKey
        ' Trailing Backslashes
        fullKeyErr(key, "Trailing backslash (\) found in DetectFile", scanSlashes And key.value.Last = CChar("\"), correctSlashes, key.value, key.value.TrimEnd(CChar("\")))
        ' Nested wildcards
        If key.vHas("*") Then
            Dim splitDir As String() = key.value.Split(CChar("\"))
            For i As Integer = 0 To splitDir.Count - 1
                fullKeyErr(key, "Nested wildcard found in DetectFile", splitDir(i).Contains("*") And i <> splitDir.Count - 1)
            Next
        End If
        ' Make sure that DetectFile paths point to a filesystem location
        chkPathFormatValidity(key, False)
        Return key
    End Function

    ''' <summary>
    ''' Audits the syntax of file system and registry paths
    ''' </summary>
    ''' <param name="key">An iniKey to be audited</param>
    ''' <param name="isRegistry">Specifies whether the given key holds a registry path</param>
    Private Sub chkPathFormatValidity(key As iniKey, isRegistry As Boolean)
        ' Remove the flags from ExcludeKeys if we have them before getting the first directory portion
        Dim rootStr As String = If(key.keyType <> "ExcludeKey", getFirstDir(key.value), getFirstDir(pathFromExcludeKey(key)))
        ' Ensure that registry paths have a valid hive and file paths have either a variable or a drive letter
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
    ''' Checks whether the current value appears in the given list of strings (case insensitive). Returns true if there is a duplicate,
    ''' otherwise, adds the current value to the list.
    ''' </summary>
    ''' <param name="valueStrings">A list of strings holding observed values</param>
    ''' <param name="currentValue">The current value to be audited</param>
    ''' <returns></returns>
    Private Function chkDupes(ByRef valueStrings As List(Of String), currentValue As String) As Boolean
        If valueStrings.Contains(currentValue.ToLower) Then Return True
        valueStrings.Add(currentValue.ToLower)
        Return False
    End Function

    ''' <summary>
    ''' Sorts a keylist alphabetically with winapp2.ini precedence applied to the key values
    ''' </summary>
    ''' <param name="kl">The keylist to be sorted</param>
    ''' <param name="hadDuplicatesRemoved">The boolean indicating whether keys have been removed from this list</param>
    Private Sub sortKeys(ByRef kl As keyList, hadDuplicatesRemoved As Boolean)
        If Not scanAlpha Or kl.keyCount <= 1 Then Exit Sub
        Dim keyValues = kl.toListOfStr(True)
        Dim sortedKeyValues = replaceAndSort(keyValues, "|", " \ \")
        ' Rewrite the alphabetized keys back into the keylist (fixes numbering silently) 
        Dim keysOutOfPlace = False
        findOutOfPlace(keyValues, sortedKeyValues, kl.keyType, kl.lineNums, keysOutOfPlace)
        If (keysOutOfPlace Or hadDuplicatesRemoved) And fixFormat(correctAlpha) Then kl.renumberKeys(sortedKeyValues)
    End Sub

    ''' <summary>
    ''' Processes a list of ExcludeKey format winapp2.ini iniKey objects and checks them for errors, correcting where possible
    ''' </summary>
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
                If key.value.StartsWith("FILE") Or key.value.StartsWith("PATH") Or key.value.StartsWith("REG") Then fullKeyErr(key, "Missing pipe symbol after ExcludeKey flag)") : Exit Sub
                fullKeyErr(key, "No valid exclude flag (FILE, PATH, or REG) found in ExcludeKey.")
        End Select
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
    ''' <param name="kl">A list of winapp2.ini FileKey format iniFiles</param>
    Private Sub cOptimization(ByRef kl As keyList)
        If kl.keyCount < 2 Or Not scanOpti Then Exit Sub
        Dim dupes As New keyList
        Dim newKeys As New keyList
        Dim flagList As New List(Of String)
        Dim paramList As New List(Of String)
        newKeys.add(kl.keys)
        For i As Integer = 0 To kl.keyCount - 1
            Dim tmpWa2 As New winapp2KeyParameters(kl.keys(i))
            ' If we have yet to record any params, record them and move on
            If paramList.Count = 0 Then trackParamAndFlags(paramList, flagList, tmpWa2) : Continue For
            ' This should handle the case where for a FileKey: 
            ' The folder provided has appeared in another key
            ' The flagstring (RECURSE, REMOVESELF, "") for both keys matches
            ' The first appearing key should have its parameters appended to and the second appearing key should be removed
            If paramList.Contains(tmpWa2.pathString) Then
                For j As Integer = 0 To paramList.Count - 1
                    If tmpWa2.pathString = paramList(j) And tmpWa2.flagString = flagList(j) Then
                        Dim keyToMergeInto As New winapp2KeyParameters(kl.keys(j))
                        Dim mergeKeyStr As String = ""
                        addArgs(mergeKeyStr, keyToMergeInto)
                        tmpWa2.argsList.ForEach(Sub(arg) mergeKeyStr += $";{arg}")
                        If tmpWa2.flagString <> "None" Then mergeKeyStr += $"|{tmpWa2.flagString}"
                        dupes.add(kl.keys(i))
                        newKeys.keys(j) = New iniKey(mergeKeyStr)
                        Exit For
                    End If
                Next
                trackParamAndFlags(paramList, flagList, tmpWa2)
            Else
                trackParamAndFlags(paramList, flagList, tmpWa2)
            End If
        Next
        If dupes.keyCount > 0 Then
            newKeys.remove(dupes.keys)
            For i As Integer = 0 To newKeys.keyCount - 1
                newKeys.keys(i).name = $"FileKey{i + 1}"
            Next
            printOptiSect("Optmization opportunity detected", kl)
            printOptiSect("The following keys can be merged into other keys:", dupes)
            printOptiSect("The resulting keyList will be reduced to: ", newKeys)
            If correctOpti Then kl = newKeys
        End If
    End Sub

    ''' <summary>
    ''' Prints output from the Optimization function
    ''' </summary>
    ''' <param name="boxStr">The text to go in the optimization section box</param>
    ''' <param name="kl">The list of keys to be printed beneath the box</param>
    Private Sub printOptiSect(boxStr As String, kl As keyList)
        cwl() : print(0, tmenu(boxStr), closeMenu:=True) : cwl()
        kl.keys.ForEach(Sub(key) cwl(key.toString))
        cwl()
    End Sub

    ''' <summary>
    ''' Tracks params and flags from a winapp2key
    ''' </summary>
    ''' <param name="paramList">The list of params observed</param>
    ''' <param name="flagList">The list of flags observed</param>
    ''' <param name="params">The current set of params and flags</param>
    Private Sub trackParamAndFlags(ByRef paramList As List(Of String), ByRef flagList As List(Of String), params As winapp2KeyParameters)
        paramList.Add(params.pathString)
        flagList.Add(params.flagString)
    End Sub

    ''' <summary>
    ''' Constructs a new iniKey in an attempt to merge keys together
    ''' </summary>
    ''' <param name="tmpKeyStr">The string to contain the new key text</param>
    ''' <param name="tmp2wa2">A set of parameters to append</param>
    Private Sub addArgs(ByRef tmpKeyStr As String, tmp2wa2 As winapp2KeyParameters)
        appendStrs({$"{tmp2wa2.keyType}{tmp2wa2.keyNum}=", $"{tmp2wa2.pathString}|", tmp2wa2.argsList(0)}, tmpKeyStr)
        If tmp2wa2.argsList.Count > 1 Then
            For i As Integer = 1 To tmp2wa2.argsList.Count - 1
                tmpKeyStr += $";{tmp2wa2.argsList(i)}"
            Next
        End If
    End Sub

    ''' <summary>
    ''' Prints an error when data is received that does not match an expected value.
    ''' </summary>
    ''' <param name="linecount">The line number of the error</param>
    ''' <param name="err">The string containing the output error text</param>
    ''' <param name="received">The (erroneous) input data</param>
    ''' <param name="expected">The expected data</param>
    ''' <param name="cond">Optional condition under which to display the error</param>
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
        If cond Then customErr(entry.lineNum, errTxt, {$"Entry Name: {entry.fullName}"})
    End Sub

    ''' <summary>
    ''' Prints an error whose output text contains an iniKey string
    ''' </summary>
    ''' <param name="key">The inikey to be printed</param>
    ''' <param name="err">The string containing the output error text</param>
    ''' <param name="cond">Optional condition under which the err should be printed</param>
    ''' <param name="repCond">Optional condition under which to repair the given key</param>
    ''' <param name="newVal">The value to replace the error value</param>
    ''' <param name="repairVal">The error value</param>
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
        lines.ToList.ForEach(Sub(errStr As String) cwl(errStr))
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