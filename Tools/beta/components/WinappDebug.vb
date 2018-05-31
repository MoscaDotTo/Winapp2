Option Strict On
Imports System.IO
Imports System.Text.RegularExpressions

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
    Dim enVars As String() = {"UserProfile", "ProgramFiles", "RootDir", "WinDir", "AppData", "SystemDrive", "SystemRoot", "Documents", "ProgramData", "AllUsersProfile", "Pictures", "Video", "CommonAppData", "LocalAppData", "CommonProgramFiles", "HomeDrive", "Music", "tmp", "Temp", "LocalLowAppData", "Public"}
    Dim validCmds As String() = {"SpecialDetect", "FileKey", "RegKey", "Detect", "LangSecRef", "Warning", "Default", "Section", "ExcludeKey", "DetectFile", "DetectOS"}
    Dim sdList As String() = {"DET_CHROME", "DET_MOZILLA", "DET_THUNDERBIRD", "DET_OPERA"}

    'Registry hive regex strings 
    Dim longReg As New Regex("HKEY_(C(URRENT_(USER$|CONFIG$)|LASSES_ROOT$)|LOCAL_MACHINE$|USERS$)")
    Dim shortReg As New Regex("HK(C(C$|R$|U$)|LM$|U$)")
    Dim secRefNums As New Regex("30(2([0-9])|3(0|1))")

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

    'Return the default parameters to the commandline handler
    Public Sub initDebugParams(ByRef firstFile As iniFile, ByRef secondFile As iniFile, cf As Boolean)
        initDefaultSettings()
        firstFile = winappFile
        secondFile = outputFile
        cf = correctFormatting
    End Sub

    'Restore the default state of the module's parameters
    Private Sub initDefaultSettings()
        winappFile.resetParams()
        outputFile.resetParams()
        settingsChanged = False
        correctFormatting = False
    End Sub

    'Handle the input from the command line
    Public Sub remoteDebug(ByRef firstFile As iniFile, secondFile As iniFile, cformatting As Boolean)
        winappFile = firstFile
        outputFile = secondFile
        correctFormatting = cformatting
        initDebug()
    End Sub

    'Make sure that if a scan is disabled, so too is its repair. Likewise, if a repair is enabled, also enable the scan for it.
    Private Sub toggleScanSetting(ByRef setting As Boolean, ByRef pairedSetting As Boolean, type As String)
        If type = "Scan" Then
            toggleSettingParam(setting, "Scan ", scanSettingsChanged)
            If exitCode Then Exit Sub
            If Not (setting) And pairedSetting Then toggleSettingParam(pairedSetting, "Scan ", scanSettingsChanged)
        Else
            toggleSettingParam(setting, "Repair ", scanSettingsChanged)
            If exitCode Then Exit Sub
            If setting And Not pairedSetting Then toggleSettingParam(pairedSetting, "Scan ", scanSettingsChanged)
        End If
    End Sub

    'Restore the settings and alert the user
    Private Sub resetSettings()
        initDefaultSettings()
        resetScanSettings()
        menuTopper = "WinappDebug settings have been reset to their defaults"
    End Sub

    Private Sub printScansMenu()
        printMenuTop({"Enable or disable specific scans or repairs"}, True)
        printBlankMenuLine()
        printMenuLine("Scan Options", "l")
        printBlankMenuLine()
        printMenuOpt("Casing", enStr(scanCasing) & " detecting improper CamelCasing")
        printMenuOpt("Alphabetization", enStr(scanAlpha) & " detecting improper alphabetization")
        printMenuOpt("Numbering", enStr(scanNumbers) & " detecting improper key numbering")
        printMenuOpt("Parameters", enStr(scanParams) & " detecting improper parameterization on FileKeys and ExcludeKeys")
        printMenuOpt("Flags", enStr(scanFlags) & " detecting improper RECURSE and REMOVESELF formatting")
        printMenuOpt("Slashes", enStr(scanSlashes) & " detecting propblems surrounding use of slashes (\)")
        printMenuOpt("Defaults", enStr(scanDefaults) & " detecting Default=True or missing Default")
        printMenuOpt("Optimizations", enStr(scanOpti) & " detecting situations where keys can be merged")
        printBlankMenuLine()
        printMenuLine("Repair Options", "l")
        printBlankMenuLine()
        printMenuOpt("Casing", enStr(correctCasing) & " fixing improper CamelCasing")
        printMenuOpt("Alphabetization", enStr(correctAlpha) & " fixing improper alphabetization")
        printMenuOpt("Numbering", enStr(correctNumbers) & " fixing improper key numbering")
        printMenuOpt("Parameters", enStr(correctParameters) & " fixing improper parameterization on FileKeys and ExcludeKeys")
        printMenuOpt("Flags", enStr(correctFlags) & " fixing improper RECURSE and REMOVESELF formatting")
        printMenuOpt("Slashes", enStr(correctSlashes) & " fixing propblems surrounding use of slashes (\)")
        printMenuOpt("Defaults", enStr(correctDefaults) & " setting Default=True to Default=False or missing Default")
        printMenuOpt("Optimizations", enStr(correctOpti) & " automatic merging of keys")
        printIf(scanSettingsChanged, "reset", "Scan and Repair", "")
        printMenuLine(menuStr02)
    End Sub

    'Present the scan setting menu 
    Private Sub changeScans()
        Console.Clear()
        Console.WindowHeight = 35
        menuTopper = "Scan and Repair Settings"
        Dim input As String
        Dim iExitCode As Boolean = False
        Do Until iExitCode
            Console.Clear()
            printScansMenu()
            Console.Write("Enter a number: ")
            input = Console.ReadLine
            Dim tmp As Boolean = New Boolean() {scanCasing, scanAlpha, scanNumbers, scanParams, scanFlags, scanSlashes, scanDefaults,
                correctCasing, correctAlpha, correctNumbers, correctParameters, correctFlags, correctSlashes,
                correctDefaults, scanSettingsChanged}.All(Function(x As Boolean) x)
            Dim tmp2 As Boolean = New Boolean() {scanOpti, correctOpti}.All(Function(x As Boolean) x)

            settingsChanged = Not tmp And Not tmp2
            Dim tmp3 As Boolean = New Boolean() {correctCasing, correctAlpha, correctNumbers, correctParameters, correctFlags, correctSlashes, correctDefaults}.All(Function(x As Boolean) x)
            correctFormatting = tmp3
            correctSomeFormatting = Not tmp3

            Select Case input
                Case "0"
                    If scanSettingsChanged Then settingsChanged = True
                    iExitCode = True
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
                    menuTopper = If(scanSettingsChanged, "Settings Reset", invInpStr)
                Case Else
                    menuTopper = invInpStr
            End Select
        Loop
    End Sub

    'Present the menu to the user
    Private Sub printMenu()
        printMenuTop({"Scan winapp2.ini for syntax and style errors, and attempt to repair them."}, True)
        printMenuOpt("Run (default)", "Run the debugger")

        printBlankMenuLine()
        printMenuOpt("File Chooser (winapp2.ini)", "Choose a new winapp2.ini name or location")

        printBlankMenuLine()
        printMenuOpt("Toggle Autocorrect", enStr(correctFormatting) & " saving of corrected errors")
        printIf(correctFormatting, "opt", "File Chooser (save)", "Choose a new save name or location")

        printBlankMenuLine()
        printMenuOpt("Toggle Scan Settings", "Enable or disable individual scan routines and auto corrections")

        printBlankMenuLine()
        printMenuLine("Current winapp2.ini:  " & replDir(winappFile.path), "l")
        printIf(correctFormatting, "line", "Current save target:  " & replDir(outputFile.path), "l")
        printIf(settingsChanged, "reset", "WinappDebug", "")
        printMenuLine(menuStr02)
    End Sub

    Sub main()
        initMenu("WinappDebug", 35)

        Do While exitCode = False
            Console.Clear()
            printMenu()
            cwl()
            Console.Write(promptStr)
            Dim input As String = Console.ReadLine()

            Try
                Select Case True
                    Case input = "0"
                        cwl("Returning to menu...")
                        exitCode = True
                    Case input = "1" Or input = ""
                        initDebug()
                    Case input = "2"
                        changeFileParams(winappFile, settingsChanged)
                    Case input = "3"
                        toggleSettingParam(correctFormatting, "Autocorrect ", settingsChanged)
                    'If the input is 4, either want to change the save file or change the scan settings
                    Case input = "4" And correctFormatting
                        changeFileParams(outputFile, settingsChanged)
                    Case input = "4" And Not correctFormatting
                        changeScans()
                    'If the input is 5, we want to reset the settings iff both correctformatting and settingschanged are true, 
                    'otherwise, if correctformatting, we want to change the scan settings
                    Case input = "5" And (Not correctFormatting And settingsChanged)
                        resetSettings()
                    Case input = "5" And correctFormatting
                        changeScans()
                    Case input = "6" And (correctFormatting And settingsChanged)
                        resetSettings()
                    Case Else
                        menuTopper = invInpStr
                End Select
            Catch ex As Exception
                exc(ex)
            End Try
        Loop
        revertMenu()
    End Sub

    'Sort the entry names as we would keys
    Private Function sortEntryNames(ByVal file As iniFile) As List(Of String)
        Dim entryList As List(Of String) = file.getSectionNamesAsList
        Dim sortedEntryList As List(Of String) = replaceAndSort(entryList, "-", "  ")
        findOutOfPlace(entryList, sortedEntryList, "Entry", file.getLineNumsFromSections, False)
        Return sortedEntryList
    End Function

    Private Sub initDebug()
        winappFile.validate()
        If exitCode Then Exit Sub
        debug(winappFile)
        If exitCode Then Exit Sub
        menuTopper = "Debug Complete"
    End Sub

    Private Sub debug(cfile As iniFile)
        'don't continue if we have a pending exit
        If exitCode Then Exit Sub
        Console.Clear()

        printMenuLine(tmenu("Beginning analysis of winapp2.ini"))
        printMenuLine(menuStr02)
        cwl()

        Dim winapp2file As New winapp2file(cfile)
        numErrs = 0
        allEntryNames = New List(Of String)

        'process the chrome entries
        For Each entry In winapp2file.cEntriesW
            processEntry(entry)
        Next

        'process the firefox entries
        For Each entry In winapp2file.fxEntriesW
            processEntry(entry)
        Next

        'process the thunderbird entries
        For Each entry In winapp2file.tbEntriesW
            processEntry(entry)
        Next

        'process the rest of the entries
        For Each entry In winapp2file.mEntriesW
            processEntry(entry)
        Next

        'Sort all the entries by name
        Dim CEntryList As List(Of String) = sortEntryNames(winapp2file.cEntries)
        Dim TBentryList As List(Of String) = sortEntryNames(winapp2file.tbEntries)
        Dim fxEntryList As List(Of String) = sortEntryNames(winapp2file.fxEntries)
        Dim mEntryList As List(Of String) = sortEntryNames(winapp2file.mEntries)

        printMenuLine(tmenu("Completed analysis of winapp2.ini"))
        printMenuLine(menuStr03)
        printMenuLine(numErrs & " possible errors were detected.", "l")
        printMenuLine("Number of entries: " & cfile.sections.Count, "l")
        printBlankMenuLine()

        're-write any changes we've made back to the file
        If correctFormatting Then
            printBlankMenuLine()
            printMenuLine("Saving changes, do not close winapp2ool or data loss may occur...", "l")

            'rebuild our internal changes
            winapp2file.rebuildToIniFiles()
            sortIniFile(winapp2file.cEntries, CEntryList)
            sortIniFile(winapp2file.tbEntries, TBentryList)
            sortIniFile(winapp2file.fxEntries, fxEntryList)
            sortIniFile(winapp2file.mEntries, mEntryList)

            'save them to file
            Try
                Dim file As New StreamWriter(outputFile.path, False)
                file.Write(winapp2file.winapp2string)
                file.Close()
            Catch ex As Exception
                exc(ex)
            End Try
            printMenuLine("Finished saving changes.", "l")
            printBlankMenuLine()
        End If

        printMenuLine("Press any key to return to the menu.", "l")
        printMenuLine(menuStr02)
        If Not suppressOutput Then Console.ReadKey()
    End Sub

    'Construct a neighbor list of values for strings in a list
    Private Sub buildNeighborList(someList As List(Of String), neighborList As List(Of KeyValuePair(Of String, String)))
        neighborList.Add(New KeyValuePair(Of String, String)("first", someList(1)))
        For i As Integer = 1 To someList.Count - 2
            neighborList.Add(New KeyValuePair(Of String, String)(someList(i - 1), someList(i + 1)))
        Next
        neighborList.Add(New KeyValuePair(Of String, String)(someList(someList.Count - 2), "last"))
    End Sub


    Private Sub findOutOfPlace(ByRef someList As List(Of String), ByRef sortedList As List(Of String), ByVal findType As String, ByRef LineCountList As List(Of Integer), ByRef oopBool As Boolean)

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

                If Not initalNeighbors(i).Key.Equals(sortedNeigbors(sind).Key) And Not initalNeighbors(i).Value.Equals(sortedNeigbors(sind).Value) Then
                    misplacedEntries.Add(someList(i))
                End If
            Next

            'Report any misplaced entries back to the user
            For Each entry In misplacedEntries
                Dim recind As Integer = someList.IndexOf(entry)
                Dim sortind As Integer = sortedList.IndexOf(entry)
                Dim curpos As Integer = LineCountList(recind)
                Dim sortpos As Integer = LineCountList(sortind)

                If findType <> "Entry" Then entry = findType & recind + 1 & "=" & entry
                If Not oopBool Then oopBool = True

                customErr(LineCountList(recind), findType & " alphabetization", {entry & " appears to be out of place", "Current line: " & curpos, "Expected line: " & LineCountList(sortind)})
            Next
        End If
    End Sub

    Private Sub cFormat(ByVal key As iniKey, ByRef keyNumber As Integer, ByRef keyList As List(Of String), ByRef dupeList As List(Of iniKey))

        'check for wholly duplicate commands 
        checkDupsAndNumbering(keyList, key, keyNumber, dupeList)

        'Scan for and fix any use of incorrect slashes
        fixFullKeyErrIf(scanSlashes, key.vHas(CChar("/")), key, "Forward slash (/) detected in lieu of blackslash (\)", correctSlashes, key.value, key.value.Replace(CChar("/"), CChar("\")))

        'make sure we don't have any dangly bits on the end of our key
        fixFullKeyErrIf(scanParams, key.toString.Last = CChar(";"), key, "Trailing semicolon (;).", correctParameters, key.value, key.value.TrimEnd(CChar(";")))

        'Do some formatting checks for environment variables
        If {"FileKey", "ExcludeKey", "DetectFile"}.Contains(key.keyType) Then cEnVar(key)

    End Sub

    Private Sub cEnVar(key As iniKey)
        'Checks the validaity of any %EnvironmentVariables%
        If key.vHas("%") Then

            Dim varcheck As String() = key.value.Split(CChar("%"))
            fullKeyErrIf(varcheck.Count <> 3 And varcheck.Count <> 5, key, "%EnvironmentVariables% must be surrounded on both sides by a single '%' character.")

            If varcheck.Count = 3 And Not enVars.Contains(varcheck(1)) Then
                Dim casingerror As Boolean = False

                'if we have a casing error, fix it in memory and inform the user
                For Each var In enVars
                    If varcheck(1).ToLower = var.ToLower Then
                        casingerror = True
                        fixFullKeyErrIf(scanCasing, True, key, "Invalid CamelCasing - expected %" & var & "% but found %" & varcheck(1) & "%", correctCasing, key.value, key.value.Replace(varcheck(1), var))
                    End If
                Next

                'If we don't have a casing error and enVars doesn't contain our value, it's invalid. 
                fullKeyErrIf(Not casingerror, key, "Misformatted or invalid environment variable: %" & varcheck(1) & "%")
            End If
        End If
    End Sub

    'Do some basic syntax/formatting checks that apply across all keys and return false if a key is malformatted
    Private Function cValidity(key As iniKey) As Boolean

        'Return false immediately if we are meant to delete the current key 
        If key.keyType = "DeleteMe" Then Return False

        'Check for leading or trailing whitespace
        fixFullKeyErrIf(True, key.vStartsOrEndsWith(" "), key, "Detected unwanted whitespace in inikey value", True, key.value, key.value.Trim(CChar("\")))
        fixFullKeyErrIf(True, key.nStartsOrEndsWith(" "), key, "Detected unwanted whitespace in inikey name", True, key.name, key.name.Trim(CChar("\")))

        'make sure we have a valid command
        If validCmds.Contains(key.keyType) Then
            Return True
        Else
            'check if there's a casing error
            If scanCasing Then
                For Each cmd In validCmds
                    If key.keyType.ToLower = cmd.ToLower Then
                        fullKeyErr(key, cmd & " is formatted improperly")
                        fixStr(correctCasing, key.name, key.name.Replace(key.keyType, cmd))
                        fixStr(correctCasing, key.keyType, cmd)
                        Return True
                    End If
                Next
            End If

            'If there's no casing error, inform the user and return false
            fullKeyErr(key, "Invalid command detected.")
            Return True
        End If
    End Function

    Private Function fixFormat(setting As Boolean) As Boolean
        Return correctFormatting Or (correctSomeFormatting And setting)
    End Function

    Private Sub pLangSecRef(ByRef entry As winapp2entry)

        'make sure we only have LangSecRef if we have LangSecRef at all.
        If entry.langSecRef.Count <> 0 And entry.sectionKey.Count <> 0 Then
            For Each key In entry.sectionKey
                err(key.lineNumber, "Section detected alongside LangSecRef.", key.toString)
            Next
        End If

        'make sure we only have 1 langsecref at most.
        confirmOnlyOne(entry.langSecRef)

        'make sure we only have 1 section at most.
        confirmOnlyOne(entry.sectionKey)

        'validate the content of any LangSecRef keys.
        For Each key In entry.langSecRef
            'and make sure the LangSecRef number is a valid one.
            fullKeyErrIf(Not secRefNums.IsMatch(key.value), key, "LangSecRef holds an invalid value.")
        Next
    End Sub

    Private Sub pFileKey(ByRef keyList As List(Of iniKey))
        If keyList.Count = 0 Then Exit Sub

        Dim curFileKeyNum As Integer = 1
        Dim fileKeyStrings As New List(Of String)
        Dim dupeKeys As New List(Of iniKey)

        For Each key In keyList

            'Get the parameters given to the file key and sort them 
            Dim keyParams As New winapp2KeyParameters(key)
            keyParams.argsList.Sort()
            Dim argsStrings As New List(Of String)
            Dim trimmedArgStrings As New List(Of String)
            Dim dupeArgs As New List(Of String)

            'check for duplicate args
            For Each arg In keyParams.argsList
                If argsStrings.Contains(arg) Then
                    'Or trimmedArgStrings.Contains(arg.ToLower) Then '    <--- "temporarily" disable this
                    err(key.lineNumber, "Duplicate FileKey parameter found", arg)
                    dupeArgs.Add(arg)
                Else
                    argsStrings.Add(arg)
                    'If arg.Contains("*") Then     <---- Disable the string matching for wildcards 
                    '    Dim splitArg As String() = arg.Split(CChar("*"))
                    '    Dim trimmedArg As String = ""
                    '    For i As Integer = 0 To splitArg.Length - 1
                    '        trimmedArg += splitArg(i)
                    '    Next
                    '    If trimmedArgStrings.Contains(trimmedArg.ToLower) Then
                    '        err(key.lineNumber, "Duplicate FileKey parameter found", arg)
                    '        dupeArgs.Add(arg)
                    '    Else
                    '        trimmedArgStrings.Add(trimmedArg.ToLower)
                    '    End If
                    'Else
                    '    trimmedArgStrings.Add(arg.ToLower)
                    'End If

                End If
            Next

            'Remove any duplicate arguments from the key parameters
            For Each arg In dupeArgs
                keyParams.argsList.Remove(arg)
            Next

            If fixFormat(correctParameters) Then keyParams.reconstructKey(key)

            'check the format of the filekey
            cFormat(key, curFileKeyNum, fileKeyStrings, dupeKeys)

            'Pipe symbol checks
            Dim iteratorCheckerList() As String = Split(key.value, "|")
            fullKeyErrIf(Not key.vHas("|"), key, "Missing pipe (|) in FileKey.")

            'Captures any incident of semi colons coming before the first pipe symbol

            fullKeyErrIf(key.vHas(";") And key.value.IndexOf(";") < key.value.IndexOf("|"), key, "Semicolon (;) found before pipe (|).")

            'check for incorrect spellings of RECURSE or REMOVESELF
            If iteratorCheckerList.Length > 2 Then fullKeyErrIf(Not iteratorCheckerList(2).Contains("RECURSE") And Not iteratorCheckerList(2).Contains("REMOVESELF"),
                             key, "RECURSE or REMOVESELF is incorrectly spelled, or there are too many pipe (|) symbols.")

            'check for missing pipe symbol on recurse and removeself, fix them if detected
            cFlags(key, {"RECURSE", "REMOVESELF"})

            'make sure VirtualStore folders point to the correct place
            If key.vHas("\virtualStore\p", True) And key.vHasnt({"programdata", "program files*", "program*"}, True) Then
                err2(key.lineNumber, "Incorrect VirtualStore location.", key.value, "%LocalAppData%\VirtualStore\Program Files*\")
            End If

            'backslash checks, fix if detected
            fixFullKeyErrIf(scanSlashes, key.vHas("%\|"), key, "Backslash (\) found before pipe (|).", correctSlashes, key.value, key.value.Replace("%\|", "%|"))

            fullKeyErrIf(key.vHas("%") And key.vHasnt({"%|", "%\"}, False), key, "Missing backslash (\) after %EnvironmentVariable%.")

        Next

        If keyList.Count > 1 Then
            'Optimization Check
            If scanOpti Then cOptimization(keyList, dupeKeys)

            'remove any duplicates
            removeDuplicateKeys(keyList, dupeKeys)

            'check the alphabetization of our filekeys
            sortKeys(keyList)
        End If

    End Sub

    Private Sub pDefault(keyList As List(Of iniKey))

        'Make sure we only have one Default.
        confirmOnlyOne(keyList)

        'Make sure all entries are disabled by Default.
        For Each key In keyList
            fixFullKeyErrIf(scanNumbers, Not key.nameIs("Default"), key, "Unnecessary numbering detected", correctNumbers, key.name, "Default")
            fixFullKeyErrIf(scanDefaults, Not key.vIs("False"), key, "All entries should be disabled by default (Default= False).", correctDefaults, key.value, "False")
        Next
    End Sub

    'Print an error for each extra key in keylists longer than 1 
    Private Sub confirmOnlyOne(keylist As List(Of iniKey))
        If keylist.Count > 1 Then
            For i As Integer = 1 To keylist.Count - 1
                fullKeyErr(keylist(i), "Multiple " & keylist(i).keyType & " detected.")
            Next
        End If
    End Sub

    Private Sub pDetOS(ByRef keyList As List(Of iniKey))
        'Make sure we have only one DetectOS
        confirmOnlyOne(keyList)
    End Sub

    'this will very simply and non verbosely alert the user when entries aren't in winapp2 order.
    'this is primarily to give output where none might otherwise exist when WinappDebug reoganzes the entries 
    'to be in proper order, but isn't otherwise all that helpful or detailed.
    Private Sub validateLineNums(ByRef entryLinesList As List(Of Integer), keyList As List(Of iniKey))

        Dim newLines As List(Of Integer) = getLineNumsFromKeyList(keyList)
        newLines.Sort()

        If newLines.Count > 0 And entryLinesList.Count > 0 Then
            If newLines(0) < entryLinesList.Last Then
                If correctFormatting Then

                    cwl(keyList(0).keyType & " detected out Of place.")
                    cwl("This error will be corrected automatically.")
                    cwl()
                End If
            End If
            entryLinesList.AddRange(newLines)
        Else
            If entryLinesList.Count = 0 Then entryLinesList.AddRange(newLines)
        End If
        entryLinesList.Sort()
    End Sub

    Private Sub processEntry(entry As winapp2entry)

        'Check for duplicate names that are differently cased 
        If allEntryNames.Contains(entry.name.ToLower) Then
            err(entry.lineNum, "Duplicate entry name detected", entry.fullname)
        Else
            allEntryNames.Add(entry.name.ToLower)
        End If

        'Check that the entry is named properly 
        If Not entry.name.EndsWith(" *") Then err(entry.lineNum, "All entries must End In ' *'", entry.fullname)

        Dim entryLinesList As New List(Of Integer)

        'validate all the keys in the entry
        For Each lst In entry.keyListList
            Dim brokenKeys As New List(Of iniKey)
            For Each key In lst
                If Not cValidity(key) Then brokenKeys.Add(key)
            Next
            For Each key In brokenKeys
                lst.Remove(key)
                entry.errorKeys.Remove(key)
            Next
        Next

        'process the DetectOS key
        pDetOS(entry.detectOS)
        validateLineNums(entryLinesList, entry.detectOS)

        'process LangSecRef &/or Section keys
        pLangSecRef(entry)
        validateLineNums(entryLinesList, entry.langSecRef)
        validateLineNums(entryLinesList, entry.sectionKey)

        'process the SpecialDetect key
        pSpecialDetect(entry.specialDetect)
        validateLineNums(entryLinesList, entry.specialDetect)

        'process the registry Detects
        pDetect(entry.detects)
        validateLineNums(entryLinesList, entry.detects)

        'process the DetectFiles
        pDetect(entry.detectFiles)
        validateLineNums(entryLinesList, entry.detectFiles)

        'process the Default key
        pDefault(entry.defaultKey)
        validateLineNums(entryLinesList, entry.defaultKey)

        'process warnings
        pWarning(entry.warningKey)
        validateLineNums(entryLinesList, entry.warningKey)

        'process the FileKeys
        pFileKey(entry.fileKeys)
        validateLineNums(entryLinesList, entry.fileKeys)

        'process the RegKeys
        pRegKey(entry.regKeys)
        validateLineNums(entryLinesList, entry.regKeys)

        'process the ExcludeKeys
        pExcludeKey(entry.excludeKeys)
        validateLineNums(entryLinesList, entry.excludeKeys)

        'Make sure we have at least 1 valid detect key 
        oneOffErr(entry.detectOS.Count + entry.detects.Count + entry.specialDetect.Count + entry.detectFiles.Count = 0,
                  "No valid detection keys detected in " & entry.name & "(Line " & entry.lineNum & ")")

        'Make sure we have at least 1 FileKey or RegKey
        If entry.fileKeys.Count + entry.regKeys.Count = 0 Then
            oneOffErr(True, "No valid FileKey/RegKeys detected in " & entry.name & "(Line " & entry.lineNum & ")")

            'If we have no FileKeys or RegKeys, we shouldn't have any ExcludeKeys.
            oneOffErr(entry.excludeKeys.Count > 0, "ExcludeKeys detected without any valid FileKeys or RegKeys in " & entry.name & " (Line " & entry.lineNum & ")")
        End If

        'Make sure we have a Default key.
        If entry.defaultKey.Count = 0 Then
            oneOffErr(True, "No Default key found in " & entry.fullname)

            'We don't have a default key, so create one and insert it into the entry.
            If fixFormat(correctDefaults) Then
                entry.defaultKey.Add(New iniKey("Default=False", 0))
            End If

        End If

    End Sub

    Private Sub pWarning(ByRef keyList As List(Of iniKey))
        'Make sure we have only one warning
        confirmOnlyOne(keyList)
    End Sub

    Private Sub pDetectFile(ByRef keyList As List(Of iniKey))

        For Each key In keyList

            'Check our environment variables
            cEnVar(key)

            'backslash check
            fixFullKeyErrIf(scanSlashes, key.value.Last = CChar("\"), key, "Trailing backslash (\) found in DetectFile", correctSlashes, key.value, key.value.TrimEnd(CChar("\")))

            'check for nested wildcards
            If key.value.Contains("*") Then
                Dim splitDir As String() = key.value.Split(CChar("\"))
                For i As Integer = 0 To splitDir.Count - 1
                    If splitDir.Last = Nothing Then Continue For
                    fullKeyErrIf(splitDir(i).Contains("*") And i <> splitDir.Count - 1, key, "Nested wildcard found in DetectFile")
                Next
            End If

            'Make sure that DetectFile paths point to a filesystem location
            fullKeyErrIf(Not key.value.StartsWith("%") And Not key.value.Contains(":\"), key, "DetectFile can only be used for file system paths.")
        Next
    End Sub

    Private Sub chkRegistryValidity(ByRef key As iniKey)
        Dim hiveStr As String = key.value.Split(CChar("\"))(0)
        fullKeyErrIf(Not longReg.IsMatch(hiveStr) And Not shortReg.IsMatch(hiveStr), key, "Invalid registry path detected")
    End Sub

    Private Sub pRegDetect(ByRef keylist As List(Of iniKey))
        'Make sure that detect paths point to a registry location.
        For Each key In keylist
            chkRegistryValidity(key)
        Next
    End Sub

    'Audit numbering from 1 to infinity and return any duplicates back to the calling function to be deleted 
    Public Sub checkDupsAndNumbering(ByRef keyStrings As List(Of String), ByRef key As iniKey, ByRef keyNumber As Integer, ByRef dupeList As List(Of iniKey))

        'check for duplicates
        If keyStrings.Contains(key.value) Then
            customErr(key.lineNumber, "Duplicate key value found", {"Key:            " & key.toString, "Duplicates:     " & key.keyType & keyStrings.IndexOf(key.value) + 1 & "=" & key.value})
            dupeList.Add(key)
        Else
            keyStrings.Add(key.value)
        End If

        'make sure the current key is correctly numbered
        If scanNumbers And Not key.nameIs(key.keyType & keyNumber) Then
            err2(key.lineNumber, key.keyType & " entry is incorrectly numbered.", key.name, key.keyType & keyNumber)
            fixStr(correctNumbers, key.name, key.keyType & keyNumber)
        End If
        keyNumber += 1
    End Sub

    Private Sub pDetect(ByRef keyList As List(Of iniKey))
        If keyList.Count = 0 Then Exit Sub

        'tracking variables
        Dim currentDetectNum As Integer = 1
        Dim detectStrings As New List(Of String)

        'make sure that if we have only one detect/file, it doesn't have a number
        If keyList.Count = 1 And Not keyList(0).nameIs(keyList(0).keyType) Then
            err2(keyList(0).lineNumber, "Detected unnecessary numbering.", keyList(0).name, keyList(0).keyType)
            fixStr(correctNumbers, keyList(0).name, keyList(0).keyType)
        End If

        'send off Detect/Files for their specific checks 
        If keyList(0).typeIs("Detect") Then
            pRegDetect(keyList)
        Else
            pDetectFile(keyList)
        End If

        'build a list of any duplicate keys that may exist
        Dim dupeKeys As New List(Of iniKey)

        If keyList.Count > 1 Then

            For Each key In keyList
                cFormat(key, currentDetectNum, detectStrings, dupeKeys)
            Next

            'remove any duplicates
            removeDuplicateKeys(keyList, dupeKeys)

            'check alphabetization
            sortKeys(keyList)
        End If
    End Sub

    'Sort a list of keys alphabetically based on their values 
    Private Sub sortKeys(ByRef keyList As List(Of iniKey))
        If Not scanAlpha Then Exit Sub

        If keyList.Count > 0 Then
            Dim keyStrings As List(Of String) = getValuesFromKeyList(keyList)
            Dim sortedKeyList As List(Of String) = replaceAndSort(keyStrings, "|", " \ \")
            Dim anyOutOfPlace As Boolean = False
            findOutOfPlace(keyStrings, sortedKeyList, keyList(0).keyType, getLineNumsFromKeyList(keyList), anyOutOfPlace)

            'rewrite the alphabetized keys back into the keylist (fixes numbering silently) 
            If anyOutOfPlace Then
                If fixFormat(correctAlpha) Then
                    Dim i As Integer = 1
                    For Each key In keyList
                        key.name = key.keyType & i
                        key.value = sortedKeyList(i - 1)
                        i += 1
                    Next
                End If
            End If
        End If
    End Sub

    Private Sub pRegKey(ByVal keyList As List(Of iniKey))

        Dim curRegKey As Integer = 1
        Dim regKeyList As New List(Of String)
        Dim dupeKeys As New List(Of iniKey)

        For Each key In keyList
            'Check the formatting
            cFormat(key, curRegKey, regKeyList, dupeKeys)

            'Ensure that each RegKey points to a valid registry location
            chkRegistryValidity(key)
        Next

        'remove any duplicates
        removeDuplicateKeys(keyList, dupeKeys)

        'Check the alphabetization
        sortKeys(keyList)
    End Sub

    'Detect missing pipe symbols before flags
    Private Sub cFlags(ByRef key As iniKey, flagStrs As String())
        For Each flagstr In flagStrs
            fixFullKeyErrIf(scanFlags, key.vHas(flagstr) And Not key.vHas("|" & flagstr), key, "Missing pipe (|) before " & flagstr & ".", correctFlags, key.value, key.value.Replace(flagstr, "|" & flagstr))
        Next
    End Sub

    Private Sub pSpecialDetect(ByRef keyList As List(Of iniKey))

        'make sure we have at most 1 SpecialDetect.
        confirmOnlyOne(keyList)

        'make sure that any SpecialDetect keys hold a valid value
        For Each key In keyList

            'Make sure SpecialDetect is not followed by a number
            fixFullKeyErrIf(scanNumbers, Not key.nameIs("SpecialDetect"), key, "Unnecessary numbering detected.", correctNumbers, key.name, "SpecialDetect")

            'Confirm that the key holds a valid value
            If Not sdList.Contains(key.value) Then
                Dim casingError As Boolean = False
                For Each item In sdList

                    'Check for casing errors
                    If key.vIs(item.ToLower, True) Then
                        fixFullKeyErrIf(scanCasing, True, key, "SpecialDetect has a casing error.", correctCasing, key.value, item)
                        casingError = True
                    End If
                Next
                fullKeyErrIf(Not casingError, key, "SpecialDetect holds an invalid value.")
            End If
        Next
    End Sub

    Private Sub pExcludeKey(ByRef keyList As List(Of iniKey))

        Dim curExcludeKeyNum As Integer = 1
        Dim excludeStrings As New List(Of String)
        Dim dupeKeys As New List(Of iniKey)
        For Each key In keyList

            'Check the format.
            cFormat(key, curExcludeKeyNum, excludeStrings, dupeKeys)

            'Make sure any filesystem exclude paths have a backslash before their pipe symbol.
            If key.vHasAny({"FILE|", "PATH|"}) Then
                Dim iteratorCheckerList() As String = Split(key.value, "|")
                Dim endingslashchecker() As String = Split(key.value, "\|")
                fullKeyErrIf(Not endingslashchecker.Count = 2, key, "Missing backslash (\) before pipe (|) in ExcludeKey.")
            Else
                fullKeyErrIf(Not key.vHas("REG|"), key, "No proper exclude flag (FILE, PATH, or REG) found in ExcludeKey")
            End If
        Next

        'Observe any potential optimizations
        If keyList.Count > 1 Then
            If scanOpti Then cOptimization(keyList, dupeKeys)

            'Remove any duplicates
            removeDuplicateKeys(keyList, dupeKeys)

            'Sort the ExcludeKeys.
            sortKeys(keyList)
        End If
    End Sub

    'Observe whether or not we can attempt to merge keys together, TODO: see how much of this functionality can be moved into the keyparamaters class
    Private Sub cOptimization(ByRef keyList As List(Of iniKey), ByRef dupeList As List(Of iniKey))
        Dim flagList As New List(Of String)
        Dim paramList As New List(Of String)
        For Each key In keyList
            Dim tmpWa2 As New winapp2KeyParameters(key)
            If paramList.Contains(tmpWa2.paramString) Then
                Dim firstInd As Integer = paramList.IndexOf(tmpWa2.paramString)
                If flagList(firstInd) = tmpWa2.flagString And Not (tmpWa2.flagString = "REG" Or tmpWa2.flagString = "PATH") Then
                    Dim tmp2wa2 As New winapp2KeyParameters(keyList(firstInd))
                    Dim tmpKeyStr As String = ""
                    Select Case key.keyType
                        Case "ExcludeKey"
                            tmpKeyStr += keyList(firstInd).name & "="
                            tmpKeyStr += tmpWa2.flagString & "|"
                            tmpKeyStr += tmpWa2.paramString & "|"
                            addArgs(tmpKeyStr, tmp2wa2)
                            For Each arg In tmpWa2.argsList
                                tmpKeyStr += ";" & arg
                            Next
                        Case "FileKey"
                            tmpKeyStr += keyList(firstInd).name & "="
                            tmpKeyStr += tmpWa2.paramString & "|"
                            addArgs(tmpKeyStr, tmp2wa2)
                            For Each arg In tmpWa2.argsList
                                tmpKeyStr += ";" & arg
                            Next
                            If tmpWa2.flagString <> "" Then tmpKeyStr += "|" & tmpWa2.flagString
                    End Select
                    Dim tmpKey As New iniKey(tmpKeyStr, keyList(firstInd).lineNumber)

                    printMenuLine(tmenu("Optmization opportunity detected"))
                    printBlankMenuLine()
                    printMenuLine("It appears that these two keys can be merged:", "c")
                    printMenuLine(menuStr02)
                    cwl()
                    cwl(keyList(firstInd).toString)
                    cwl()
                    cwl(key.toString)
                    cwl()
                    cwl("The result will be:")
                    cwl(tmpKey.toString)
                    cwl()
                    printMenuLine(menuStr00)
                    printMenuLine("If multiple lines are detected for optimization, the above result may not represent the final output.", "c")
                    printMenuLine(menuStr02)
                    cwl()
                    If fixFormat(correctOpti) Then
                        keyList(firstInd).value = tmpKey.value
                        dupeList.Add(key)
                    End If
                End If
            End If
            paramList.Add(tmpWa2.paramString)
            flagList.Add(tmpWa2.flagString)
        Next
    End Sub

    Private Sub addArgs(ByRef tmpKeyStr As String, tmp2wa2 As winapp2KeyParameters)
        tmpKeyStr += tmp2wa2.argsList(0)
        If tmp2wa2.argsList.Count > 1 Then
            For i As Integer = 1 To tmp2wa2.argsList.Count - 1
                tmpKeyStr += ";" & tmp2wa2.argsList(i)
            Next
        End If
    End Sub

    'Take remove the items in one list of inikeys from another 
    Private Sub removeDuplicateKeys(ByRef keylist As List(Of iniKey), ByVal dupeList As List(Of iniKey))
        For Each key In dupeList
            keylist.Remove(key)
        Next
    End Sub

    'Print an error when we receive an input different from that which we expected
    Private Sub err2(linecount As Integer, err As String, command As String, expected As String)
        customErr(linecount, err, {"Expected: " & expected, "Found: " & command})
    End Sub

    'Print an error when we receive invalid or improper input of some sort
    Private Sub err(linecount As Integer, err As String, command As String)
        customErr(linecount, err, {"Command: " & command})
    End Sub

    'Print an error whose output contains the full text of the key
    Private Sub fullKeyErr(key As iniKey, err As String)
        customErr(key.lineNumber, err, {"Key: " & key.toString})
    End Sub

    'Print arbitrarily defined errors
    Private Sub customErr(lineCount As Integer, err As String, lines As String())
        cwl("Line: " & lineCount & " - Error: " & err)
        For Each errStr In lines
            cwl(errStr)
        Next
        cwl()
        numErrs += 1
    End Sub

    'Print a fullKeyErr given its precondition
    Private Sub fullKeyErrIf(cond As Boolean, key As iniKey, errStr As String)
        If cond Then fullKeyErr(key, errStr)
    End Sub

    'Given both a scan and a fix condition, print a fullkeyerr and attempt to fix it 
    Private Sub fixFullKeyErrIf(scanCond As Boolean, errCond As Boolean, key As iniKey, errStr As String, repCond As Boolean, ByRef repairVal As String, newVal As String)
        If scanCond Then
            If errCond Then
                fullKeyErr(key, errStr)
                fixStr(repCond, repairVal, newVal)
            End If
        End If
    End Sub

    'Print a single line error string given its precondition
    Private Sub oneOffErr(cond As Boolean, errStr As String)
        If cond Then
            cwl(errStr)
            numErrs += 1
        End If
    End Sub

    'Replace a string (from an inikey most likely) with a new value if fixFormat for the given param is true
    Private Sub fixStr(param As Boolean, ByRef currentValue As String, newValue As String)
        If fixFormat(param) Then currentValue = newValue
    End Sub

End Module