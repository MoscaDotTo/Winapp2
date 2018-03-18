Option Strict On
Imports System.IO

Module WinappDebug

    Dim numErrs As Integer
    Dim enVars As String() = {"UserProfile", "ProgramFiles", "RootDir", "WinDir", "AppData", "SystemDrive", "SystemRoot", "Documents", "ProgramData", "AllUsersProfile", "Pictures", "Video", "CommonAppData", "LocalAppData", "CommonProgramFiles", "HomeDrive", "Music", "tmp", "Temp", "LocalLowAppData", "Public"}
    Dim validCmds As String() = {"SpecialDetect", "FileKey", "RegKey", "Detect", "LangSecRef", "Warning", "Default", "Section", "ExcludeKey", "DetectFile", "DetectOS"}
    Dim sdList As String() = {"DET_CHROME", "DET_MOZILLA", "DET_THUNDERBIRD", "DET_OPERA"}
    Dim name As String = "\winapp2.ini"
    Dim path As String = Environment.CurrentDirectory

    Dim allEntryNames As List(Of String)
    Dim correctFormatting As Boolean = False
    Dim menuHasTopper As Boolean = False
    Dim exitCode As Boolean = False

    Public Sub remoteDebug(filedir As String, filename As String, cformatting As Boolean)
        path = filedir
        name = filename
        correctFormatting = cformatting
        initDebug()
    End Sub

    Private Sub printMenu()
        Dim correctStatus As String = IIf(correctFormatting, "on", "off").ToString
        If Not menuHasTopper Then
            menuHasTopper = True
            printMenuLine(tmenu("WinappDebug"))
        End If
        printMenuLine(menuStr03)
        printMenuLine("This tool will check winapp2.ini For common syntax And style errors.", "c")
        printMenuLine(menuStr04)
        printMenuLine("0. Exit                          - Return to the winapp2ool menu", "l")
        printMenuLine("1. Run (default)                 - Run with the default settings", "l")
        printMenuLine("2. Run (custom)                  - Run with an option to provide the path and filename", "l")
        printMenuLine(menuStr01)
        printMenuLine("3. Toggle Autocorrect            - Enable/Disable automatic correction of certain types of errors. (" & correctStatus & ")", "l")
        printMenuLine(menuStr02)
    End Sub

    Sub main()
        Console.Clear()
        exitCode = False
        menuHasTopper = False
        Do While exitCode = False
            printMenu()
            cwl()
            Console.Write("Enter a number, or leave blank to run the default: ")
            Dim input As String = Console.ReadLine()

            Try
                Select Case input
                    Case "0"
                        cwl("Returning to menu...")
                        exitCode = True
                    Case "1", ""
                        initDebug()
                    Case "2"
                        fChooser(path, name, exitCode, "\winapp2.ini", "")
                        initDebug()
                    Case "3"
                        correctFormatting = Not correctFormatting
                        Console.Clear()
                        printMenuLine(tmenu("Autocorrect toggled."))
                    Case Else
                        Console.Clear()
                        printMenuLine(tmenu("Invalid input. Please try again."))
                End Select
            Catch ex As Exception
                exc(ex)
            End Try
        Loop
    End Sub

    Private Function sortEntryNames(ByVal file As iniFile) As List(Of String)
        Dim entryList As List(Of String) = file.getSectionNamesAsList
        Dim sortedEntryList As List(Of String) = replaceAndSort(entryList, "-", "  ")
        findOutOfPlace(entryList, sortedEntryList, "Entry", getLineNumsFromSections(file))
        Return entryList
    End Function

    Private Sub initDebug()
        Dim winapp2 As iniFile = validate(path, name, exitCode, "\winapp2.ini", "")
        debug(winapp2)
        revertMenu(exitCode)
        If Not exitCode Then
            Console.Clear()
            printMenuLine(tmenu("WinappDebug"))
        End If
    End Sub

    Private Sub debug(cfile As iniFile)
        'don't continue if we have a pending exit
        If exitCode Then Exit Sub
        Console.Clear()

        printMenuLine(tmenu("Beginning analysis of winapp2.ini"))
        printMenuLine(menuStr02)
        cwl()

        Dim winapp2file As New winapp2file(cFile)
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

        'grab the sorted state of the entry names
        Dim CEntryList As List(Of String) = sortEntryNames(winapp2file.cEntries)
        Dim TBentryList As List(Of String) = sortEntryNames(winapp2file.tbEntries)
        Dim fxEntryList As List(Of String) = sortEntryNames(winapp2file.fxEntries)
        Dim mEntryList As List(Of String) = sortEntryNames(winapp2file.mEntries)

        printMenuLine(tmenu("Completed analysis of winapp2.ini"))
        printMenuLine(menuStr03)
        printMenuLine(numErrs & " possible errors were detected.", "l")
        printMenuLine("Number of entries: " & cFile.sections.Count, "l")
        printMenuLine(menuStr01)

        're-write any changes we've made back to the file
        If correctFormatting Then
            printMenuLine(menuStr01)
            printMenuLine("Saving changes, do not close winapp2ool or data loss may occur...", "l")

            'rebuild our internal changes
            winapp2file.rebuildToIniFiles()
            sortIniFile(winapp2file.cEntries, CEntryList)
            sortIniFile(winapp2file.tbEntries, TBentryList)
            sortIniFile(winapp2file.fxEntries, fxEntryList)
            sortIniFile(winapp2file.mEntries, mEntryList)

            'save them to file
            Try
                Dim file As New StreamWriter(cFile.dir & cFile.name, False)
                file.Write(winapp2file.winapp2string)
                file.Close()
            Catch ex As Exception
                exc(ex)
            End Try
            printMenuLine("Finished saving changes.", "l")
            printMenuLine(menuStr01)
        End If

        printMenuLine("Press any key to return to the winapp2ool menu.", "l")
        printMenuLine(menuStr02)
        If Not suppressOutput Then Console.ReadKey()
    End Sub


    Private Sub findOutOfPlace(ByRef someList As List(Of String), ByRef sortedList As List(Of String), ByVal findType As String, ByRef LineCountList As List(Of Integer))

        Dim originalPlacement As New List(Of String)
        originalPlacement.AddRange(someList.ToArray)
        Dim originalLines As New List(Of Integer)
        originalLines.AddRange(LineCountList)

        Dim moveRight As New List(Of String)
        Dim moveLeft As New List(Of String)

        Dim offset As Integer = 0
        For i As Integer = 0 To someList.Count - 1

            Dim entry As String = someList(i)
            Dim sortInd As Integer = sortedList.IndexOf(entry)

            If i <> sortInd Then

                If Not moveRight.Contains(entry) Then
                    If sortInd > i Then
                        moveRight.Add(entry)
                    Else
                        LineCountList.Insert(sortInd, LineCountList(i))
                        LineCountList.RemoveAt(i + 1)

                        moveLeft.Add(entry)

                        'move the entry to the left in our list
                        someList.Insert(sortInd, entry)
                        someList.RemoveAt(i + 1)

                        'jump back to the position from which we sorted the entry to make sure we remove any entries that were off by one because of it not having been moved yet
                        i = sortInd
                    End If
                End If
            Else
                'Don't keep tracking lines who are now correctly positioned once we've moved another line left in the list
                If moveRight.Contains(entry) Then moveRight.Remove(entry)
            End If
        Next

        Dim misplacedList As New List(Of String)
        misplacedList.AddRange(moveRight.ToArray)
        misplacedList.AddRange(moveLeft.ToArray)

        For Each entry In misplacedList
            Dim recind As Integer = originalPlacement.IndexOf(entry)
            Dim sortind As Integer = sortedList.IndexOf(entry)
            Dim curpos As Integer = originalLines(recind)
            Dim sortpos As Integer = LineCountList(sortind)

            If findType <> "Entry" Then entry = findType & recind + 1 & "=" & entry
            cwl("Line: " & originalLines(recind) & " - Error: '" & findType & "' alphabetization. ")
            cwl(entry & " appears to be out of place.")
            cwl("Current Position: Line " & curpos)
            cwl("Expected Position: Line " & originalLines(sortind))
            cwl()
            numErrs += 1
        Next

    End Sub

    Private Sub err2(linecount As Integer, err As String, command As String, expected As String)
        cwl("Line: " & linecount & " - Error: " & err)
        cwl("Expected: " & expected)
        cwl("Command: " & command)
        cwl()
        numErrs += 1
    End Sub

    Private Sub err(linecount As Integer, err As String, command As String)
        cwl("Line: " & linecount & " - Error: " & err)
        cwl("Command: " & command)
        cwl()
        numErrs += 1
    End Sub

    Private Sub fullKeyErr(key As iniKey, err As String)
        cwl("Line: " & key.lineNumber & " - Error: " & err)
        cwl("Command: " & key.toString)
        cwl()
        numErrs += 1
    End Sub

    Private Sub cFormat(ByVal key As iniKey, ByRef keyNumber As Integer, ByRef keyList As List(Of String), ByRef dupeList As List(Of iniKey))

        'check for wholly duplicate commands 
        checkDupsAndNumbering(keyList, key, keyNumber, dupeList)

        'make sure we don't have any dangly bits on the end of our key
        If key.toString.Last = CChar(";") Then
            err(key.lineNumber, "Trailing semicolon (;).", key.toString)
            key.value = key.value.TrimEnd(CChar(";"))
        End If

        'Do some formatting checks for environment variables
        If key.keyType = "FileKey" Or key.keyType = "ExcludeKey" Or key.keyType = "DetectFile" Then cEnVar(key)

    End Sub

    Private Sub cEnVar(key As iniKey)
        'Checks the validaity of any %EnvironmentVariables%

        If key.value.Contains("%") Then

            Dim varcheck As String() = key.value.Split(Convert.ToChar("%"))
            If varcheck.Count <> 3 And varcheck.Count <> 5 Then err(key.lineNumber, "%EnvironmentVariables% must be surrounded on both sides by a single '%' character.", key.value)

            If varcheck.Count = 3 Then
                If Not enVars.Contains(varcheck(1)) Then
                    Dim casingerror As Boolean = False
                    For Each var In enVars
                        If varcheck(1).ToLower = var.ToLower Then
                            casingerror = True
                            'if we have a casing error, fix it in memory and inform the user
                            err2(key.lineNumber, "Invalid CamelCasing on environment variable.", varcheck(1), var)
                            key.value = key.value.Replace(varcheck(1), var)
                        End If
                    Next

                    'If we don't have a casing error and enVars doesn't contain our value, it's invalid. 
                    If Not casingerror Then err(key.lineNumber, "Misformatted or invalid environment variable.", key.value)
                End If
            End If
        End If
    End Sub

    Private Function cValidity(key As iniKey) As Boolean
        'returns true if the key contains a valid command, otherwise returns false
        'Fixes any casing errors and trims any unwanted whitespace

        'Check for trailing whitespace
        If key.value.EndsWith(" ") Or key.value.StartsWith(" ") Or key.name.EndsWith(" ") Or key.name.StartsWith(" ") Then
            fullKeyErr(key, "Detected unwanted whitepace.")
            key.value = key.value.Trim(CChar(" "))
            key.name = key.name.Trim(CChar(" "))
        End If

        'make sure we have a valid command
        If validCmds.Contains(key.keyType) Then
            Return True
        Else
            'check if there's a casing error
            For Each cmd In validCmds
                If key.keyType.ToLower = cmd.ToLower Then
                    err2(key.lineNumber, "Command is formatted improperly.", key.keyType, cmd)
                    key.name = key.name.Replace(key.keyType, cmd)
                    key.keyType = cmd
                    Return True
                End If
            Next
            'If there's no casing error, inform the user and return false
            fullKeyErr(key, "Invalid command detected.")
            Return False
        End If
    End Function

    Private Sub pLangSecRef(ByRef entry As winapp2entry)

        'make sure we only have LangSecRef if we have LangSecRef at all.
        If entry.langSecRef.Count <> 0 And entry.sectionKey.Count <> 0 Then
            For Each key In entry.sectionKey
                err(key.lineNumber, "Section detected alongside LangSecRef.", key.toString)
            Next
        End If

        'make sure we only have 1 langsecref at most.
        If entry.langSecRef.Count > 1 Then
            For i As Integer = 1 To entry.langSecRef.Count - 1
                err(entry.langSecRef(i).lineNumber, "Muliple LangSecRef detected.", entry.langSecRef(i).toString)
            Next
        End If

        'make sure we only have 1 section at most.
        If entry.sectionKey.Count > 1 Then
            For i As Integer = 1 To entry.sectionKey.Count - 1
                err(entry.sectionKey(i).lineNumber, "Multiple Sections detected.", entry.sectionKey(i).toString)
            Next
        End If

        'validate the content of any LangSecRef keys.
        For Each key In entry.langSecRef
            If key.value <> "" Then
                Dim validSecRefs As New List(Of String)
                validSecRefs.AddRange(New String() {"3021", "3022", "3023", "3024", "3025", "3026", "3027", "3028", "3029", "3030", "3031"})
                Dim hasValidSecRef As Boolean = False

                'and make sure the LangSecRef number is a valid one.
                If Not validSecRefs.Contains(key.value) Then fullKeyErr(key, "LangSecRef holds an invalid value.")
            End If
        Next
    End Sub

    Private Sub pFileKey(ByRef keyList As List(Of iniKey))

        Dim curFileKeyNum As Integer = 1
        Dim fileKeyStrings As New List(Of String)
        Dim dupeKeys As New List(Of iniKey)

        For Each key In keyList

            'Sort the alphabetically arguements given to the filekey 
            Dim keyParams As New winapp2KeyParameters(key)
            keyParams.argsList.Sort()
            keyParams.reconstructKey(key)

            'check the format of the filekey
            cFormat(key, curFileKeyNum, fileKeyStrings, dupeKeys)

            'Pipe symbol checks
            Dim iteratorCheckerList() As String = Split(key.value, "|")
            If Not key.value.Contains("|") Then fullKeyErr(key, "Missing pipe (|) in FileKey.")
            If key.value.Contains(";") Then
                'Captures any incident of semi colons coming before the first pipe symbol
                If key.value.IndexOf(";") < key.value.IndexOf("|") Then fullKeyErr(key, "Semicolon (;) found before pipe (|).")
            End If

            'check for incorrect spellings of RECURSE or REMOVESELF
            If iteratorCheckerList.Length > 2 Then
                If Not iteratorCheckerList(2).Contains("RECURSE") And Not iteratorCheckerList(2).Contains("REMOVESELF") Then fullKeyErr(key, "'RECURSE' or 'REMOVESELF' entry is incorrectly spelled, or there are too many pipe (|) symbols.")
            End If

            'check for missing pipe symbol on recurse and removeself, fix them if detected
            If key.value.Contains("RECURSE") And Not key.value.Contains("|RECURSE") Then
                fullKeyErr(key, "Missing pipe (|) before RECURSE.")
                key.value = key.value.Replace("RECURSE", "|RECURSE")
            End If
            If key.value.Contains("REMOVESELF") And Not key.value.Contains("|REMOVESELF") Then
                fullKeyErr(key, "Missing pipe (|) before REMOVESELF.")
                key.value = key.value.Replace("REMOVESELF", "|REMOVESELF")
            End If

            'make sure VirtualStore folders point to the correct place
            If key.value.Contains("\VirtualStore\P") And (Not key.value.ToLower.Contains("programdata") And Not key.value.ToLower.Contains("program files*") And Not key.value.ToLower.Contains("program*")) Then
                err2(key.lineNumber, "Incorrect VirtualStore location.", key.value, "%LocalAppData%\VirtualStore\Program Files*\")
            End If

            'backslash checks, fix if detected
            If key.value.Contains("%\|") Then
                fullKeyErr(key, "Backslash (\) found before pipe (|).")
                key.value = key.value.Replace("%\|", "%|")
            End If
            If key.value.Contains("%") And Not key.value.Contains("%|") And Not key.value.Contains("%\") Then fullKeyErr(key, "Missing backslash (\) after %EnvironmentVariable%.")
        Next

        'remove any duplicates
        removeDuplicateKeys(keyList, dupeKeys)

        'check the alphabetization of our filekeys
        sortKeys(keyList)
    End Sub

    Private Sub pDefault(keyList As List(Of iniKey))

        'Make sure we only have one Default.
        If keyList.Count > 1 Then
            For i As Integer = 1 To keyList.Count - 1
                err(keyList(i).lineNumber, "Multiple Default keys found.", keyList(i).toString)
            Next
        End If

        'Make sure all entries are disabled by Default.
        For Each key In keyList

            If Not key.name = "Default" Then
                fullKeyErr(key, "Unnecessary numbering detected")
                key.name = "Default"
            End If

            If Not key.value.ToLower.Equals("false") Then
                fullKeyErr(key, "All entries should be disabled by default (Default=False).")
                key.value = "False"
            End If
        Next
    End Sub

    Private Sub pDetOS(ByRef keyList As List(Of iniKey))
        'Make sure we have only one DetectOS
        If keyList.Count > 1 Then
            For i As Integer = 1 To keyList.Count - 1
                err(keyList(i).lineNumber, "Multiple DetectOS detected.", keyList(i).toString)
            Next
        End If
    End Sub

    Private Sub validateLineNums(ByRef entryLinesList As List(Of Integer), keyList As List(Of iniKey))

        Dim newLines As List(Of Integer) = getLineNumsFromKeyList(keyList)
        newLines.Sort()

        If newLines.Count > 0 And entryLinesList.Count > 0 Then
            'this will very simply and non verbosely alert the user when entries aren't in winapp2 order.
            'this is primarily to give output where none might otherwise exist when WinappDebug reoganzes the entries 
            'to be in proper order, but isn't otherwise all that helpful or detailed.
            If newLines(0) < entryLinesList.Last Then
                If correctFormatting Then
                    cwl(keyList(0).keyType & " detected out of place.")
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
        If Not entry.name.EndsWith(" *") Then err(entry.lineNum, "All entries must end in ' *'", entry.fullname)

        Dim entryLinesList As New List(Of Integer)

        'validate all the keys in the entry
        For Each lst In entry.keyListList
            For Each key In lst
                cValidity(key)
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
        If entry.detectOS.Count + entry.detects.Count + entry.specialDetect.Count + entry.detectFiles.Count = 0 Then
            cwl("No valid detection keys detected in " & entry.name & "(Line " & entry.lineNum & ")")
            numErrs += 1
        End If

        'Make sure we have at least 1 FileKey or RegKey
        If entry.fileKeys.Count + entry.regKeys.Count = 0 Then
            cwl("No valid FileKey/RegKeys detected in " & entry.name & "(Line " & entry.lineNum & ")")
            numErrs += 1

            'If we have no FileKeys or RegKeys, we shouldn't have any ExcludeKeys.
            If entry.excludeKeys.Count > 0 Then
                cwl("ExcludeKeys detected without any valid FileKeys or RegKeys in " & entry.name & "(Line " & entry.lineNum & ")")
                numErrs += 1
            End If
        End If

        'Make sure we have a Default key.
        If entry.defaultKey.Count = 0 Then
            cwl("No Default key found in " & entry.fullname)

            'We don't have a default key, so create one and insert it into the entry.
            Dim defKey As New iniKey()
            defKey.name = "Default"
            defKey.value = "False"
            entry.defaultKey.Add(defKey)
            numErrs += 1
        End If

    End Sub

    Private Sub pWarning(ByRef keyList As List(Of iniKey))

        'Make sure we have only one warning
        If keyList.Count > 1 Then
            For i As Integer = 1 To keyList.Count - 1
                err(keyList(i).lineNumber, "Multiple Warning detected.", keyList(i).toString)
            Next
        End If

    End Sub

    Private Sub pDetectFile(ByRef keyList As List(Of iniKey))

        For Each key In keyList

            'Check our environment variables
            cEnVar(key)

            'backslash check
            If key.value.Last = CChar("\") Then
                fullKeyErr(key, "Trailing backslash (\) found in DetectFile")
                key.value = key.value.TrimEnd(CChar("\"))
            End If

            'check for nested wildcards
            If key.value.Contains("*") Then
                Dim splitDir As String() = key.value.Split(Convert.ToChar("\"))
                For i As Integer = 0 To splitDir.Count - 1
                    If splitDir.Last = Nothing Then Continue For
                    If splitDir(i).Contains("*") And i <> splitDir.Count - 1 Then fullKeyErr(key, "Nested wildcard found in DetectFile")
                Next
            End If

            'Make sure that DetectFile paths point to a filesystem location
            If Not key.value.StartsWith("%") And Not key.value.Contains(":\") Then fullKeyErr(key, "'DetectFile' can only be used for file system paths.")
        Next
    End Sub

    Private Sub pRegDetect(ByRef keylist As List(Of iniKey))
        For Each key In keylist
            'Make sure that detect paths point to a registry location.
            If (key.toString.Contains("=%") Or key.toString.Contains("=C:\")) Or (Not key.toString.Contains("=HKLM") And Not key.toString.Contains("=HKC") And Not key.toString.Contains("=HKU")) Then
                fullKeyErr(key, "'Detect' can only be used for registry keys paths.")
            End If
        Next
    End Sub

    Public Sub checkDupsAndNumbering(ByRef keyStrings As List(Of String), ByRef key As iniKey, ByRef keyNumber As Integer, ByRef dupeList As List(Of iniKey))
        'Audit numbering from 1 to infinity and return any duplicates back to the calling function to be deleted 

        'check for duplicates
        If keyStrings.Contains(key.value) Then
            cwl("Line: " & key.lineNumber & " - Error: Duplicate key found." & Environment.NewLine &
                                  "Command: " & key.value & Environment.NewLine &
                                  "Duplicates: " & key.keyType & keyStrings.IndexOf(key.value) + 1 & "=" & key.value.ToLower & Environment.NewLine)
            dupeList.Add(key)
        Else
            keyStrings.Add(key.value)
        End If

        'make sure the current key is correctly numbered
        If Not key.name = key.keyType & keyNumber Then
            err2(key.lineNumber, key.keyType & " entry is incorrectly numbered.", key.name, key.keyType & keyNumber)
        End If
        keyNumber += 1

    End Sub


    Private Sub pDetect(ByRef keyList As List(Of iniKey))

        'tracking variables
        Dim currentDetectNum As Integer = 1
        Dim detectStrings As New List(Of String)

        'make sure that if we have only one detect/file, it doesn't have a number
        If keyList.Count = 1 Then
            If keyList.Count = 1 And Not keyList(0).name = keyList(0).keyType Then
                err2(keyList(0).lineNumber, "Detected unnecessary numbering.", keyList(0).name, keyList(0).keyType)
                keyList(0).name = keyList(0).keyType
            End If
        End If

        'send off Detect/Files for their specific checks 
        If keyList.Count > 0 Then
            If keyList(0).keyType = "Detect" Then
                pRegDetect(keyList)
            Else
                pDetectFile(keyList)
            End If
        End If

        'build a list of any duplicate keys that may exist
        Dim dupeKeys As New List(Of iniKey)

        If keyList.Count > 1 Then

            For Each key In keyList
                'check formatting
                cFormat(key, currentDetectNum, detectStrings, dupeKeys)
            Next

            'remove any duplicates
            removeDuplicateKeys(keyList, dupeKeys)

            'check alphabetization
            sortKeys(keyList)
        End If
    End Sub

    Private Sub sortKeys(ByRef keyList As List(Of iniKey))
        If keyList.Count > 0 Then
            Dim keyStrings As List(Of String) = getValuesFromKeyList(keyList)
            Dim sortedKeyList As List(Of String) = replaceAndSort(keyStrings, "|", " \ \")
            findOutOfPlace(keyStrings, sortedKeyList, keyList(0).keyType, getLineNumsFromKeyList(keyList))

            'rewrite the alphabetized keys back into the keylist
            Dim i As Integer = 1
            For Each key In keyList
                key.name = key.keyType & i
                key.value = keyStrings(i - 1)
                i += 1
            Next
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
            If Not key.toString.Contains("=HKLM\") And Not key.toString.Contains("=HKC\") And Not key.toString.Contains("=HKCU\") And Not key.toString.Contains("=HKCR\") And Not key.toString.Contains("=HKU\") Then
                fullKeyErr(key, "RegKey can only be used for registry key paths.")
            End If
        Next

        'remove any duplicates
        removeDuplicateKeys(keyList, dupeKeys)

        'Check the alphabetization
        sortKeys(keyList)
    End Sub

    Private Sub pSpecialDetect(ByRef keyList As List(Of iniKey))

        'make sure we have at most 1 SpecialDetect.
        If keyList.Count > 1 Then
            For i As Integer = 1 To keyList.Count - 1
                err(keyList(i).lineNumber, "Multiple SpecialDetects detected", keyList(i).toString)
            Next
        End If

        'make sure that any SpecialDetect keys hold a valid value
        For Each key In keyList

            'Make sure SpecialDetect is not followed by a number
            If Not key.name = "SpecialDetect" Then
                fullKeyErr(key, "Unnecessary numbering detected.")
                key.name = "SpecialDetect"
            End If

            'Confirm that the key holds a valid value
            If Not sdList.Contains(key.value) Then
                Dim casingError As Boolean = False
                For Each item In sdList

                    'Check for casing errors
                    If key.value.ToLower = item.ToLower Then
                        fullKeyErr(key, "SpecialDetect has a casing error.")
                        casingError = True
                        key.value = item
                    End If
                Next
                If Not casingError Then fullKeyErr(key, "SpecialDetect holds an invalid value.")
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
            If key.value.Contains("FILE") Or key.value.Contains("PATH") Then
                Dim iteratorCheckerList() As String = Split(key.value, "|")
                Dim endingslashchecker() As String = Split(key.value, "\|")
                If Not endingslashchecker.Count = 2 Then fullKeyErr(key, "Missing backslash (\) before pipe (|) in ExcludeKey.")
            End If
        Next

        'Remove any duplicates
        removeDuplicateKeys(keyList, dupeKeys)

        'Sort the ExcludeKeys.
        sortKeys(keyList)

    End Sub

    Private Sub removeDuplicateKeys(ByRef keylist As List(Of iniKey), ByVal dupeList As List(Of iniKey))
        For Each key In dupeList
            keylist.Remove(key)
        Next
    End Sub
End Module