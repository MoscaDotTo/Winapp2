Option Strict On
Imports System.IO

Module WinappDebug

    Dim numErrs As Integer
    Dim enVars As String() = {"UserProfile", "ProgramFiles", "RootDir", "WinDir", "AppData", "SystemDrive", "SystemRoot", "Documents", "ProgramData", "AllUsersProfile", "Pictures", "Video", "CommonAppData", "LocalAppData", "CommonProgramFiles", "HomeDrive", "Music", "tmp", "Temp", "LocalLowAppData", "Public"}
    Dim validCmds As String() = {"SpecialDetect", "FileKey", "RegKey", "Detect", "LangSecRef", "Warning", "Default", "Section", "ExcludeKey", "DetectFile", "DetectOS"}
    Dim sdList As String() = {"DET_CHROME", "DET_MOZILLA", "DET_THUNDERBIRD", "DET_OPERA"}
    Dim name As String = "\winapp2.ini"
    Dim path As String = Environment.CurrentDirectory
    Dim savePath As String = Environment.CurrentDirectory
    Dim saveName As String = "\winapp2.ini"
    'these temporary inifiles are used to hold the entries within the special sections
    Dim chromeEntries As New iniFile
    Dim fxEntries As New iniFile
    Dim tbEntries As New iniFile
    Dim mEntries As New iniFile
    Dim cEntryLines As New List(Of String)
    Dim fEntryLines As New List(Of String)
    Dim tEntryLines As New List(Of String)

    Dim menuHasTopper As Boolean = False
    Dim exitCode As Boolean = False

    Private Sub printMenu()
        If Not menuHasTopper Then
            menuHasTopper = True
            tmenu("WinappDebug")
        End If
        menu(menuStr03)
        menu("This tool will check winapp2.ini For common syntax And style errors.", "c")
        menu(menuStr04)
        menu("0. Exit                          - Return to the winapp2ool menu", "l")
        menu("1. Run (default)                 - Run with the default settings", "l")
        menu("2. Run (custom)                  - Run with an option to provide the path and filename", "l")
        menu(menuStr02)

    End Sub

    Sub main()
        Console.Clear()
        exitCode = False
        menuHasTopper = False
        Do While exitCode = False
            printMenu()
            Console.WriteLine()
            Console.Write("Enter a number, or leave blank to run the default: ")
            Dim input As String = Console.ReadLine()

            Try
                Select Case input
                    Case "0"
                        Console.WriteLine("Returning to menu...")
                        exitCode = True
                    Case "1", ""
                        debug()
                        revertMenu(exitCode)
                        If Not exitCode Then
                            Console.Clear()
                            tmenu("WinappDebug")
                        End If
                    Case "2"
                        fChooser(path, name, exitCode, "\winapp2.ini", "")
                        debug()
                        revertMenu(exitCode)
                        If Not exitCode Then
                            Console.Clear()
                            tmenu("WinappDebug")
                        End If
                    Case Else
                        Console.Clear()
                        tmenu("Invalid input. Please try again.")
                End Select
            Catch ex As Exception
                Console.WriteLine("Error: " & ex.ToString)
                Console.WriteLine("Please report this error on GitHub")
                Console.WriteLine()
            End Try
        Loop
    End Sub

    Private Sub debug()

        'don't continue if we have a pending exit
        If exitCode Then
            Exit Sub
        End If
        Console.Clear()
        'make sure our file exists and isn't empty 
        validate(path, name, exitCode, "\winapp2.ini", "")
        If exitCode Then
            Exit Sub
        End If

        Dim cFile As New iniFile(path, name)

        tmenu("Beginning analysis of winapp2.ini")
        menu(menuStr02, "")
        Console.WriteLine()

        numErrs = 0
        Dim entryTitles As New List(Of String)
        Dim trimmedEntryTitles As New List(Of String)
        Dim entryLineCounts As New List(Of String)
        Dim TBCommentLineNum As Integer = cFile.findCommentLine("; End of Thunderbird entries.")
        Dim NonCCCommentLineNum As Integer = cFile.findCommentLine("; These entries are the exact same ones located in the Removed entries files")

        Dim hasTBcomment As Boolean = True
        Dim hasNCCComment As Boolean = False

        If TBCommentLineNum = -1 Then
            hasTBcomment = False
        End If

        If Not NonCCCommentLineNum = -1 Then
            hasNCCComment = True
        End If

        For i As Integer = 0 To cFile.sections.Count - 1

            Dim curSection As iniSection = cFile.sections.Values(i)
            pSection(curSection)

            'count the entries and check for duplicates
            If Not entryTitles.Contains(curSection.name.ToLower) Then
                entryTitles.Add(curSection.name.ToLower)
            Else
                Dim duplicateEntryIndex As Integer = entryTitles.IndexOf(curSection.name.ToLower)
                Dim duplicateEntryName As String = cFile.sections.Values(duplicateEntryIndex).name
                Dim duplicateLineNum As Integer = cFile.sections.Values(duplicateEntryIndex).startingLineNumber
                Console.WriteLine("Error: Duplicate entry name detected.")
                Console.WriteLine("Line: " & curSection.startingLineNumber & " - " & curSection.name)
                Console.WriteLine("Duplicates: " & duplicateEntryName & " on line " & duplicateLineNum)
                Console.WriteLine()
                numErrs += 1
            End If

            'check that the name is formatted correctly
            If Not curSection.name.Contains(" *") Then
                err(curSection.startingLineNumber, "All entry names should end with ' *'", curSection.name)
            End If

            'separate tracking list for the main section of the file that we want alphabetized
            If Not hasTBcomment Or (hasTBcomment And curSection.startingLineNumber > TBCommentLineNum) Then
                If Not hasNCCComment Or (hasNCCComment And curSection.startingLineNumber < NonCCCommentLineNum) Then
                    trimmedEntryTitles.Add(curSection.name)
                    entryLineCounts.Add(curSection.startingLineNumber.ToString)
                End If
            End If

        Next

        'traverse through our entry sections 

        Dim CEntryList As List(Of String) = chromeEntries.getSectionNamesAsList
        Dim sortedClist As New List(Of String)
        sortedClist.AddRange(chromeEntries.getSectionNamesAsList)

        replaceAndSort(CEntryList, sortedClist, "-", "  ")
        findOutOfPlace(CEntryList, sortedClist, "Entry", cEntryLines)

        Dim FEntryList As List(Of String) = fxEntries.getSectionNamesAsList
        Dim sortedFlist As New List(Of String)
        sortedFlist.AddRange(fxEntries.getSectionNamesAsList)
        replaceAndSort(FEntryList, sortedFlist, "-", "  ")
        findOutOfPlace(FEntryList, sortedFlist, "Entry", fEntryLines)

        Dim TEntryList As List(Of String) = tbEntries.getSectionNamesAsList
        Dim sortedTlist As New List(Of String)
        sortedTlist.AddRange(tbEntries.getSectionNamesAsList)

        replaceAndSort(TEntryList, sortedTlist, "-", "  ")
        findOutOfPlace(TEntryList, sortedTlist, "Entry", tEntryLines)

        Dim sortedEList As New List(Of String)
        sortedEList.AddRange(trimmedEntryTitles.ToArray)

        replaceAndSort(trimmedEntryTitles, sortedEList, "-", "  ")
        findOutOfPlace(trimmedEntryTitles, sortedEList, "Entry", entryLineCounts)

        tmenu("Completed analysis of winapp2.ini")
        menu(menuStr03, "")
        menu(numErrs & " possible errors were detected.", "l")
        menu("Number of entries: " & entryTitles.Count, "l")
        menu(menuStr01)
        menu("Press any key to return to the winapp2ool menu.", "l")
        menu(menuStr02, "")
        Console.ReadKey()
    End Sub

    Private Sub replaceAndSort(ByRef ListToBeSorted As List(Of String), ByRef givenSortedList As List(Of String), characterToReplace As String, ByVal replacementText As String)

        Dim receivedIndex As Integer
        Dim sortedIndex As Integer

        Dim renamedList As New List(Of String)
        Dim originalNameList As New List(Of String)

        For i As Integer = 0 To ListToBeSorted.Count - 1

            Dim entry As String = ListToBeSorted(i)

            'grab the entry at the current position
            receivedIndex = i
            sortedIndex = givenSortedList.IndexOf(entry)

            'preform our replacement
            If entry.Contains(characterToReplace) Then

                Dim newentry As String = entry.Replace(characterToReplace, replacementText)
                originalNameList.Add(entry)
                renamedList.Add(newentry)
                ListToBeSorted(receivedIndex) = newentry
                givenSortedList(sortedIndex) = newentry
                entry = newentry
            End If

            'Prefix any singular numbers with 0 to maintain their first order precedence during sorting

            Dim myChars() As Char = entry.ToCharArray()

            findAndReplNumbers(myChars, entry, originalNameList, renamedList, ListToBeSorted, givenSortedList, i, sortedIndex)
        Next

        givenSortedList.Sort()

        'undo our change made in the name of sorting such that we have the original data back
        If renamedList.Count > 0 Then
            For i As Integer = 0 To renamedList.Count - 1
                Dim newentry As String = originalNameList(i)
                Dim ren As String = renamedList(i)

                Dim recievedIndex = ListToBeSorted.IndexOf(ren)
                ListToBeSorted(recievedIndex) = newentry

                sortedIndex = givenSortedList.IndexOf(ren)
                givenSortedList(sortedIndex) = newentry
            Next
        End If
    End Sub

    Private Sub findOutOfPlace(ByRef someList As List(Of String), ByRef sortedList As List(Of String), ByVal findType As String, ByRef LineCountList As List(Of String))

        Dim originalPlacement As New List(Of String)
        originalPlacement.AddRange(someList.ToArray)
        Dim originalLines As New List(Of String)
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
                If moveRight.Contains(entry) Then
                    moveRight.Remove(entry)
                End If
            End If
        Next

        Dim misplacedList As New List(Of String)
        misplacedList.AddRange(moveRight.ToArray)
        misplacedList.AddRange(moveLeft.ToArray)

        For Each entry As String In misplacedList
            Dim recind As Integer = originalPlacement.IndexOf(entry)
            Dim sortind As Integer = sortedList.IndexOf(entry)
            Dim curpos As String = originalLines(recind)
            Dim sortpos As String = LineCountList(sortind)

            If findType <> "Entry" Then
                entry = findType & recind + 1 & "=" & entry
            End If

            Console.WriteLine("Line: " & originalLines(recind) & " - Error: '" & findType & "' alphabetization. ")
            Console.WriteLine(entry & " appears to be out of place.")
            Console.WriteLine("Current Position: Line " & curpos)
            Console.WriteLine("Expected Position: Line " & originalLines(sortind))
            Console.WriteLine()
            numErrs += 1
        Next

    End Sub

    Private Sub findAndReplNumbers(ByRef myChars() As Char, ByRef entry As String, ByRef originalNameList As List(Of String), ByRef renamedList As List(Of String), ByRef ListToBeSorted As List(Of String), ByRef givenSortedList As List(Of String), ByVal receivedIndex As Integer, ByVal sortedindex As Integer)

        Dim originalEntry As String = entry
        Dim lastCharWasNum As Boolean = False
        Dim prefixIndicies As New List(Of Integer)

        For chind As Integer = 0 To myChars.Count - 1

            Dim ch As Char = myChars(chind)
            Dim nextCharIsNum As Boolean = False

            'observe if the next character is a number, we only want to pad instances of single digit numbers
            If chind < myChars.Count - 1 Then
                Dim nextindex As Integer = chind + 1
                Dim nextChar As Char = originalEntry(nextindex)

                If Char.IsDigit(nextChar) Then
                    nextCharIsNum = True
                End If
            End If

            'observe the previous character for the same reason
            If lastCharWasNum = False Then

                If Char.IsDigit(ch) = True Then
                    lastCharWasNum = True
                    If nextCharIsNum = False Then
                        prefixIndicies.Add(chind)
                    End If
                Else
                    lastCharWasNum = False
                End If
            Else
                If Char.IsDigit(ch) = False Then
                    lastCharWasNum = False
                End If
            End If
        Next

        'prefix any numbers that we detected above
        If prefixIndicies.Count >= 1 Then
            For j As Integer = 0 To prefixIndicies.Count - 1

                Dim tmp As String = entry
                tmp = tmp.Insert(prefixIndicies(j), "0")

                'make the necessary adjustments to our tracking lists
                If renamedList.Contains(entry) = False Then
                    originalNameList.Add(entry)
                    renamedList.Add(tmp)
                    ListToBeSorted(receivedIndex) = tmp
                    givenSortedList(sortedindex) = tmp
                    entry = tmp
                Else
                    renamedList(renamedList.IndexOf(entry)) = tmp
                    ListToBeSorted(receivedIndex) = tmp
                    givenSortedList(sortedindex) = tmp
                    entry = tmp
                End If

                'each time we insert our leading zero, remember to adjust the remaining indicies by 1
                For k As Integer = j + 1 To prefixIndicies.Count - 1
                    prefixIndicies(k) = prefixIndicies(k) + 1
                Next
            Next
        End If
    End Sub

    Private Sub err2(linecount As Integer, err As String, command As String, expected As String)
        Console.WriteLine("Line: " & linecount & " - Error: " & err)
        Console.WriteLine("Expected: " & expected)
        Console.WriteLine("Command: " & command)
        Console.WriteLine()
        numErrs += 1
    End Sub

    Private Sub err(linecount As Integer, err As String, command As String)
        Console.WriteLine("Line: " & linecount & " - Error: " & err)
        Console.WriteLine("Command: " & command)
        Console.WriteLine()
        numErrs += 1
    End Sub

    Private Sub fullKeyErr(key As iniKey, err As String)
        Console.WriteLine("Line: " & key.lineNumber & " - Error: " & err)
        Console.WriteLine("Command: " & key.toString)
        Console.WriteLine()
        numErrs += 1
    End Sub

    Private Sub cFormat(ByVal key As iniKey, ByRef keyNumber As Integer, ByRef keyList As List(Of String))

        Dim command As String = key.toString
        Dim keyString As String = key.keyType

        'check for wholly duplicate commands in File/Reg/ExcludeKeys
        If keyList.Contains(key.value) Then
            Console.WriteLine("Line: " & key.lineNumber & " - Error: Duplicate key found." & Environment.NewLine &
                                  "Command: " & command & Environment.NewLine &
                                  "Duplicates: " & keyString & keyList.IndexOf(key.value) + 1 & "=" & key.value.ToLower & Environment.NewLine)
        End If

        'make sure the current key is correctly numbered
        If Not key.name.Contains(keyNumber.ToString) Then
            err2(key.lineNumber, "'" & keyString & "' entry is incorrectly numbered.", command, keyString & keyNumber)
        End If
        keyNumber += 1

        'make sure we don't have any dangly bits on the end of our key
        If command(command.Count - 1) = ";" Then
            err(key.lineNumber, "Trailing semicolon (;).", command)
        End If

        'Do some formatting checks for environment variables
        If keyString = "FileKey" Or keyString = "ExcludeKey" Then
            cEnVar(key)
        End If

    End Sub

    Private Sub cEnVar(key As iniKey)

        Dim command As String = key.value

        If command.Contains("%") Then

            Dim varcheck As String() = command.Split(Convert.ToChar("%"))
            If varcheck.Count <> 3 And varcheck.Count <> 5 Then
                err(key.lineNumber, "%EnvironmentVariables% must be surrounded on both sides by a single '%' character.", command)
            End If

            If varcheck.Count = 3 Then
                If Not enVars.Contains(varcheck(1)) Then
                    Dim casingerror As Boolean = False
                    For Each var As String In enVars
                        If varcheck(1).ToLower = var.ToLower Then
                            casingerror = True
                            err2(key.lineNumber, "Invalid CamelCasing on environment variable.", varcheck(1), var)
                        End If
                    Next

                    If casingerror = False Then
                        err(key.lineNumber, "Misformatted or invalid environment variable.", command)
                    End If
                End If
            End If
        End If
    End Sub

    Private Function cValidity(key As iniKey) As Boolean

        'Check for trailing whitespace
        If key.toString.EndsWith(" ") Then
            fullKeyErr(key, "Detected unwanted trailing whitepace.")
        End If

        'Check for ending whitespace
        If key.toString.StartsWith(" ") Then
            fullKeyErr(key, "Detected unwanted leading whitespeace.")
        End If

        'make sure we have a valid command
        If validCmds.Contains(key.keyType) Then
            Return True
        Else
            'check if there's a casing error
            For Each cmd As String In validCmds
                If key.keyType.ToLower = cmd.ToLower Then
                    err2(key.lineNumber, "Command is formatted improperly.", key.keyType, cmd)
                    key.keyType = cmd
                    Return True
                End If
            Next
            fullKeyErr(key, "Invalid command detected.")
            Return False
        End If

    End Function

    Private Sub pLangSecRef(ByVal key As iniKey)

        If key.value <> "" Then
            Dim validSecRefs As New List(Of String)
            validSecRefs.AddRange(New String() {"3021", "3022", "3023", "3024", "3025", "3026", "3027", "3028", "3029", "3030", "3031"})
            Dim hasValidSecRef As Boolean = False

            'and make sure the number that follows it is valid
            If Not validSecRefs.Contains(key.value) Then
                fullKeyErr(key, "LangSecRef holds an invalid value.")
            End If
        End If
    End Sub

    Private Sub pFileKey(ByRef keylist As List(Of iniKey))

        Dim curFileKeyNum As Integer = 1
        Dim fileKeyList As New List(Of String)

        For Each key As iniKey In keylist
            'check the format of the filekey
            cFormat(key, curFileKeyNum, fileKeyList)

            'add the filekey contents to the duplicate checking list
            Dim command As String = key.value
            fileKeyList.Add(command)

            'Pipe symbol checks
            Dim iteratorCheckerList() As String = Split(command, "|")
            If Not command.Contains("|") Then
                fullKeyErr(key, "Missing pipe (|) in 'FileKey.'")
            End If
            If command.Contains(";") Then
                If command.IndexOf(";") < command.IndexOf("|") Then
                    fullKeyErr(key, "Semicolon (;) found before pipe (|).")
                End If
            End If

            'check for incorrect spellings of RECURSE or REMOVESELF
            If iteratorCheckerList.Length > 2 Then
                If Not iteratorCheckerList(2).Contains("RECURSE") And Not iteratorCheckerList(2).Contains("REMOVESELF") Then
                    fullKeyErr(key, "'RECURSE' or 'REMOVESELF' entry is incorrectly spelled, or there are too many pipe (|) symbols.")
                End If
            End If

            'check for missing pipe symbol on recurse and removeself
            If command.Contains("RECURSE") And Not command.Contains("|RECURSE") Then
                fullKeyErr(key, "Missing pipe (|) before 'RECURSE.'")
            End If
            If command.Contains("REMOVESELF") And Not command.Contains("|REMOVESELF") Then
                fullKeyErr(key, "Missing pipe (|) before 'REMOVESELF.'")
            End If

            'make sure VirtualStore folders point to the correct place
            If command.Contains("\VirtualStore\P") And (Not command.ToLower.Contains("programdata") And Not command.ToLower.Contains("program files*") And Not command.ToLower.Contains("program*")) Then
                err2(key.lineNumber, "Incorrect VirtualStore location.", command, "%LocalAppData%\VirtualStore\Program Files*\")
            End If

            'backslash checks
            If command.Contains("%\|") Then
                fullKeyErr(key, "Backslash (\) found before pipe (|).")
            End If
            If command.Contains("%") And Not command.Contains("%|") And Not command.Contains("%\") Then
                fullKeyErr(key, "Missing backslash (\) after %EnvironmentVariable%.")
            End If
        Next
    End Sub

    Private Sub pDefault(key As iniKey)

        'make sure all entries are disabled by default
        If Not key.value.ToLower.Equals("false") Then
            fullKeyErr(key, "All entries should be disabled by default (Default=False).")
        End If
    End Sub

    Private Sub pSection(section As iniSection)

        'initialize some tracking variables
        Dim entryHasRegKeys As Boolean = False
        Dim entryKeysOutOfOrder As Boolean = False
        Dim detectKeys As New List(Of iniKey)
        Dim fileKeys As New List(Of iniKey)
        Dim regKeys As New List(Of iniKey)
        Dim excludeKeys As New List(Of iniKey)
        Dim defaultKey As New iniKey
        Dim secRefKey As New iniKey
        Dim sectionKey As New iniKey
        Dim sdKey As New iniKey

        For i As Integer = 0 To section.keys.Count - 1
            Dim key As iniKey = section.keys.Values(i)
            'make sure our key contains a valid command
            If Not cValidity(key) Then
                Continue For
            End If

            Select Case key.keyType
                Case "LangSecRef"
                    secRefKey = key
                Case "FileKey"
                    fileKeys.Add(key)

                    'Make sure that FileKeys come before RegKeys in the section (style error but not syntax)
                    If entryHasRegKeys And Not entryKeysOutOfOrder Then
                        err(key.lineNumber, "FileKeys should precede RegKeys.", key.toString)
                        entryKeysOutOfOrder = True
                    End If
                Case "Detect"
                    detectKeys.Add(key)
                Case "DetectFile"
                    detectKeys.Add(key)
                Case "ExcludeKey"
                    excludeKeys.Add(key)
                Case "Default"
                    defaultKey = key
                Case "RegKey"
                    regKeys.Add(key)
                    If Not entryHasRegKeys Then
                        entryHasRegKeys = True
                    End If
                Case "SpecialDetect"
                    sdKey = key
            End Select
        Next

        'process our filekeys
        pFileKey(fileKeys)

        Dim fkstrings As New List(Of String)
        Dim sortedFKStrings As New List(Of String)
        Dim fklinecounts As New List(Of String)
        For Each filekey As iniKey In fileKeys
            fklinecounts.Add(filekey.lineNumber.ToString)
            fkstrings.Add(filekey.value)
        Next
        sortedFKStrings.AddRange(fkstrings)
        replaceAndSort(fkstrings, sortedFKStrings, "|", " \ \")
        findOutOfPlace(fkstrings, sortedFKStrings, "FileKey", fklinecounts)

        'process our regkeys
        pRegKey(regKeys)

        Dim rkstrings As New List(Of String)
        Dim sortedRKStrings As New List(Of String)
        Dim rklinecounts As New List(Of String)
        For Each regKey As iniKey In regKeys
            rklinecounts.Add(regKey.lineNumber.ToString)
            rkstrings.Add(regKey.value)
        Next
        sortedRKStrings.AddRange(rkstrings)
        replaceAndSort(rkstrings, sortedRKStrings, "|", " \ \")
        findOutOfPlace(rkstrings, sortedRKStrings, "RegKey", rklinecounts)

        'process our excludekeys
        pExcludeKey(excludeKeys)

        Dim ekstrings As New List(Of String)
        Dim sortedEKStrings As New List(Of String)
        Dim eklinecounts As New List(Of String)
        For Each excludeKey As iniKey In excludeKeys
            eklinecounts.Add(excludeKey.lineNumber.ToString)
            ekstrings.Add(excludeKey.value)
        Next
        sortedEKStrings.AddRange(ekstrings)
        replaceAndSort(ekstrings, sortedEKStrings, "|", " \ \")
        findOutOfPlace(ekstrings, sortedEKStrings, "ExcludeKey", eklinecounts)

        'process our detect keys (if any)
        pDetect(detectKeys)

        'process our langsecref (if any)
        pLangSecRef(secRefKey)
        'track our custom sections
        Select Case secRefKey.value
            Case "3029"
                If Not chromeEntries.sections.Keys.Contains(section.name) Then
                    chromeEntries.sections.Add(section.name, section)
                    cEntryLines.Add(section.startingLineNumber.ToString)
                End If
            Case "3026"
                If Not fxEntries.sections.Keys.Contains(section.name) Then
                    fxEntries.sections.Add(section.name, section)
                    fEntryLines.Add(section.startingLineNumber.ToString)
                End If
            Case "3030"
                If Not tbEntries.sections.Keys.Contains(section.name) Then
                    tbEntries.sections.Add(section.name, section)
                    tEntryLines.Add(section.startingLineNumber.ToString)
                End If
        End Select

        'process our default key, throw an error if it's absent
        If defaultKey.value <> "" Then
            pDefault(defaultKey)
        Else
            err(section.startingLineNumber, "Default state missing", section.name)
        End If

        pSpecialDetect(sdKey)
    End Sub

    Private Sub pDetectFile(ByRef key As iniKey)

        'check our environment variables
        cEnVar(key)

        Dim command As String = key.value

        'backslash check
        If command(command.Count - 1) = "\" Then
            fullKeyErr(key, "Trailing backslash on DetectFile.")
        End If

        'check for nested wildcards
        If command.Contains("*") Then
            Dim splitDir As String() = command.Split(Convert.ToChar("\"))
            For i As Integer = 0 To splitDir.Count - 1
                If splitDir.Last = Nothing Then Continue For
                If splitDir(i).Contains("*") And i <> splitDir.Count - 1 Then
                    fullKeyErr(key, "Nested wildcard found in DetectFile")
                End If
            Next
        End If

        'Make sure that DetectFile paths point to a filesystem location
        If Not command.StartsWith("%") And Not command.Contains(":\") Then
            fullKeyErr(key, "'DetectFile' can only be used for file system paths.")
        End If
    End Sub

    Private Sub pRegDetect(ByRef key As iniKey)

        Dim command As String = key.toString
        'Make sure that detect paths point to a registry location
        If (command.Contains("=%") Or command.Contains("=C:\")) Or (Not command.Contains("=HKLM") And Not command.Contains("=HKC") And Not command.Contains("=HKU")) Then
            fullKeyErr(key, "'Detect' can only be used for registry keys paths.")
        End If
    End Sub

    Private Sub pDetect(ByRef iKeyList As List(Of iniKey))

        'tracking variables
        Dim firstNumberStatus As Boolean = False
        Dim fdnumStatus As Boolean = False
        Dim fdfnumStatus As Boolean = False
        Dim curDNum As Integer = 1
        Dim curDFNum As Integer = 1
        Dim dkeylist As New List(Of String)
        Dim dlclist As New List(Of String)
        Dim dflclist As New List(Of String)
        Dim dfkeylist As New List(Of String)
        Dim keyList As New List(Of String)
        Dim detectType As String
        Dim number As Integer = 1

        For Each key As iniKey In iKeyList

            Dim command As String = key.value

            If key.keyType = "DetectFile" Then
                detectType = "DetectFile"
                keyList = dfkeylist
                number = curDFNum
                firstNumberStatus = fdfnumStatus
                dflclist.Add(key.lineNumber.ToString)
            Else
                detectType = "Detect"
                keyList = dkeylist
                number = curDNum
                firstNumberStatus = fdnumStatus
                dlclist.Add(key.lineNumber.ToString)
            End If

            'check for dangly bits
            If command.Contains(";") Then
                fullKeyErr(key, "Semicolon (;) found in " & detectType)
            End If

            'check for duplicates
            If keyList.Contains(command) Then
                Console.WriteLine("Line: " & key.lineNumber & " - Error: Duplicate command found.")
                Console.WriteLine("Command: " & key.toString)
                Console.WriteLine("Duplicates: " & detectType & keyList.IndexOf(command) + 1 & "=" & command & Environment.NewLine)
                numErrs += 1
            Else
                keyList.Add(command)
            End If

            'Check whether our first Detect or DetectFile is trailed by a number
            If number = 1 Then
                If key.toString.StartsWith(detectType & "=") Or (Not key.toString.StartsWith(detectType & "=") And Not key.toString.StartsWith(detectType & "1")) Then
                    firstNumberStatus = False
                    If detectType.Equals("DetectFile") Then
                        fdfnumStatus = False
                    Else
                        fdnumStatus = False
                    End If
                Else
                    firstNumberStatus = True
                    If detectType.Equals("DetectFile") Then
                        fdfnumStatus = True
                    Else
                        fdnumStatus = True
                    End If
                End If
            End If

            'Make sure our first number is a 1 if there's a number
            If number = 1 And (Not key.toString.ToLower.Contains(detectType.ToLower & "1=") And Not key.toString.ToLower.Contains(detectType.ToLower & "=")) Then
                err2(key.lineNumber, detectType & " numbering.", command, detectType & " or " & detectType & "1.")
            End If

            'If we're on our second detect, make sure the first one had a 1
            If number = 2 And firstNumberStatus = False Then
                err(key.lineNumber - 1, detectType & number & " detected without preceding " & detectType & "1", key.toString)
            End If

            'if we only have one detect entry, make sure it doesnt have a number
            If iKeyList.Count = 1 Then
                For i As Integer = 0 To 9
                    If key.name.Contains(i.ToString) Then
                        fullKeyErr(key, "Detected unneeded numbering")
                        Exit For
                    End If
                Next
            End If

            'otherwise, make sure our numbers match up
            If number > 1 Then
                If key.name.ToLower.Contains(detectType.ToLower & number.ToString) = False Then
                    err2(key.lineNumber, detectType & " is incorrectly numbered.", key.toString, detectType & number)
                End If
            End If
            number += 1

            'update our tracking variables as needed
            If detectType.Equals("DetectFile") Then
                curDFNum = number
                dfkeylist = keyList
                pDetectFile(key)
            Else
                curDNum = number
                dkeylist = keyList
                pRegDetect(key)
            End If
        Next

        'check alphabetization
        Dim sortedDStrings As New List(Of String)
        Dim sortedDFStrings As New List(Of String)

        sortedDStrings.AddRange(dkeylist)
        sortedDFStrings.AddRange(dfkeylist)
        replaceAndSort(dkeylist, sortedDStrings, "|", " \ \")
        findOutOfPlace(dkeylist, sortedDStrings, "Detect", dlclist)

        replaceAndSort(dfkeylist, sortedDFStrings, "|", " \ \")
        findOutOfPlace(dfkeylist, sortedDFStrings, "DetectFile", dflclist)
    End Sub

    Private Sub pRegKey(ByVal keyList As List(Of iniKey))

        Dim curRegKey As Integer = 1
        Dim regKeyList As New List(Of String)

        'Ensure that each RegKey points to a valid registry location
        For Each key As iniKey In keyList
            cFormat(key, curRegKey, regKeyList)
            If Not key.toString.Contains("=HKLM") And Not key.toString.Contains("=HKC") And Not key.toString.Contains("=HKU") Then
                fullKeyErr(key, "'RegKey' can only be used for registry key paths.")
            End If
            regKeyList.Add(key.value)
        Next
    End Sub

    Private Sub pSpecialDetect(ByRef key As iniKey)

        'make sure that any SpecialDetect keys hold a valid value
        If key.value <> "" Then
            If Not sdList.Contains(key.value) Then
                fullKeyErr(key, "SpecialDetect holds an invalid value.")
            End If
        End If
    End Sub

    Private Sub pExcludeKey(ByRef keyList As List(Of iniKey))

        Dim curExcludeKeyNum As Integer = 1
        Dim ekKeyList As New List(Of String)
        For Each key As iniKey In keyList

            'check the format
            cFormat(key, curExcludeKeyNum, ekKeyList)
            Dim command As String = key.value
            ekKeyList.Add(command)

            'Make sure any FILE exclude paths have a backslash before their pipe symbol
            If command.Contains("FILE") Then
                Dim iteratorCheckerList() As String = Split(command, "|")
                Dim endingslashchecker() As String = Split(command, "\|")
                If endingslashchecker.Count = 1 Then
                    fullKeyErr(key, "Missing backslash (\) before pipe (|) in ExcludeKey.")
                End If
            End If
        Next
    End Sub

End Module