Option Strict On
Imports System.IO

Module Module1

    Public Sub printMenu()

        Console.WriteLine("*--------------------------------------------------------------------------------------------------*")
        Console.WriteLine("*                   Winapp2ool - A multitool for winapp2.ini and related files                     *")
        Console.WriteLine("*                                                                                                  *")
        Console.WriteLine("*                                     Menu: Enter a number to select                               *")
        Console.WriteLine("*                                                                                                  *")
        Console.WriteLine("* 0. Exit             - Exits the application                                                      *")
        Console.WriteLine("* 1. WinappDebug      - Loads the WinappDebug tool to check for errors in winapp2.ini              *")
        Console.WriteLine("* 2. ccinidebug       - Loads the ccinidebug tool to sort and trim ccleaner.ini                    *")
        Console.WriteLine("* 3. Diff             - Loads the diff tool to observe the changes between two winapp2.ini files   *")
        Console.WriteLine("* 4. Trim             - Loads the trim tool to debloat winapp2.ini for your system                 *")
        Console.WriteLine("*--------------------------------------------------------------------------------------------------*")
        Console.Write("Enter a number now: ")

    End Sub

    Sub Main()
        Dim inputTest As Boolean = False

        printMenu()
        Dim cki As String = Console.ReadLine()
        Do Until inputTest = True
            Select Case cki
                Case "0"
                    inputTest = True
                    Console.WriteLine("Exiting...")
                    Environment.Exit(1)
                Case "1"
                    WinappDebug.Main()
                    Console.WriteLine("*--------------------------------------------------------------------------------------------------*")
                    Console.WriteLine("*********************************Finished running WinappDebug***************************************")
                    printMenu()
                    cki = Console.ReadLine
                Case "2"
                    ccinidebug.Main()
                    Console.WriteLine("*--------------------------------------------------------------------------------------------------*")
                    Console.WriteLine("*********************************Finished running ccinidebug****************************************")
                    printMenu()
                    cki = Console.ReadLine
                Case "3"
                    diff.main()
                    Console.WriteLine("*--------------------------------------------------------------------------------------------------*")
                    Console.WriteLine("*************************************Finished running Diff******************************************")
                    printMenu()
                    cki = Console.ReadLine
                Case "4"
                    trim.main()
                    Console.WriteLine("*--------------------------------------------------------------------------------------------------*")
                    Console.WriteLine("************************************Finished running trim*******************************************")
                    printMenu()
                    cki = Console.ReadLine
                Case Else
                    Console.WriteLine(Environment.NewLine & "Invalid command. Please try again.")
                    printMenu()
                    cki = Console.ReadLine
            End Select
        Loop
    End Sub
End Module

Module WinappDebug

    Sub Main()

        Console.Clear()
        Console.WriteLine("*--------------------------------------------------------------------------------------------------*")
        Console.WriteLine("*                                           WinappDebug                                            *")
        Console.WriteLine("*                                                                                                  *")
        Console.WriteLine("*             This tool will check winapp.ini for common syntax and style errors.                  *")
        Console.WriteLine("*                                                                                                  *")
        Console.WriteLine("*             Make sure both winapp2.ini is in the same folder as as winapp2ool.exe                *")
        Console.WriteLine("*             if the current folder is the Program Files directory, you may need                   *")
        Console.WriteLine("*             to relaunch winapp2ool.exe as an administrator.                                      *")
        Console.WriteLine("*--------------------------------------------------------------------------------------------------*")

        'Check if the winapp2.ini file is in the current directory. End program if it isn't.
        If Not File.Exists(Environment.CurrentDirectory & "\winapp2.ini") Then
            Console.WriteLine("winapp2.ini file could not be located in the current working directory (" & Environment.CurrentDirectory & ")")
            Console.ReadKey()
            End
        End If
        Dim cfile As New iniFile("winapp2.ini")
        Dim numErrs As Integer = 0
        Dim entryTitles As New List(Of String)
        Dim trimmedEntryTitles As New List(Of String)
        Dim entryLineCounts As New List(Of String)
        Dim TBCommentLineNum As Integer = findCommentLine(cfile, "; End of Thunderbird entries.")
        Dim NonCCCommentLineNum As Integer = findCommentLine(cfile, "; These entries are the exact same ones located in the Removed entries files")

        Dim hasTBcomment As Boolean = True
        Dim hasNCCComment As Boolean = False

        If TBCommentLineNum = 0 Then
            hasTBcomment = False
        End If

        If Not NonCCCommentLineNum = 0 Then
            hasNCCComment = True
        End If

        For i As Integer = 0 To cfile.sections.Count - 1

            Dim curSection As iniSection = cfile.sections.Values(i)
            pSection(curSection, numErrs)

            'count the entries and check for duplicates
            If Not entryTitles.Contains(curSection.name) Then
                entryTitles.Add(curSection.name)
                entryLineCounts.Add(curSection.startingLineNumber.ToString)
            Else
                err(curSection.startingLineNumber, "Duplicate entry name.", curSection.name, numErrs)
            End If

            'check that the name is formatted correctly
            If Not curSection.name.Contains(" *") Then
                err(curSection.startingLineNumber, "All entry names should end with ' *'", curSection.name, numErrs)
            End If

            'separate tracking list for the main section of the file that we want alphabetized
            If Not hasTBcomment Or (hasTBcomment And curSection.startingLineNumber > TBCommentLineNum) Then
                If Not hasNCCComment Or (hasNCCComment And curSection.startingLineNumber < NonCCCommentLineNum) Then
                    trimmedEntryTitles.Add(curSection.name)
                End If
            End If
        Next

        'traverse through our trimmed entry titles list

        Dim sortedEList As New List(Of String)
        sortedEList.AddRange(trimmedEntryTitles.ToArray)

        replaceAndSort(trimmedEntryTitles, sortedEList, "-", "  ", "entries")
        findOutOfPlace(trimmedEntryTitles, sortedEList, "Entry", numErrs, entryLineCounts)

        'Stop the program from closing on completion
        Console.WriteLine("***********************************************" & Environment.NewLine & "Completed analysis of winapp2.ini. " & numErrs & " possible errors were detected. " & Environment.NewLine & "Number of entries: " & entryTitles.Count & Environment.NewLine & "Press any key to return to the menu.")
        Console.ReadKey()

    End Sub

    Public Sub replaceAndSort(ByRef ListToBeSorted As List(Of String), ByRef givenSortedList As List(Of String), characterToReplace As String, ByVal replacementText As String, ByVal sortType As String)
        If ListToBeSorted.Count > 50 Then
            Console.WriteLine()
        End If
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
            'because VB.NET is silly and thinks "10" comes before "2" 

            Dim myChars() As Char = entry.ToCharArray()

            findAndReplNumbers(myChars, entry, originalNameList, renamedList, ListToBeSorted, givenSortedList, i, sortedIndex)

        Next

        givenSortedList.Sort()

        'undo our change made in the name of sorting such that we have the original data back
        If renamedList.Count > 1 Then
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

    Public Sub findOutOfPlace(ByRef someList As List(Of String), ByRef sortedList As List(Of String), ByVal findType As String, ByRef numErrs As Integer, ByRef LineCountList As List(Of String))
        Dim originalPlacement As New List(Of String)
        originalPlacement.AddRange(someList.ToArray)
        Dim originalLines As New List(Of String)
        originalLines.AddRange(LineCountList)
        Dim misplacedEntryList As New List(Of String)
        For i As Integer = 0 To someList.Count - 1

            Dim entry As String = someList(i)

            Dim sortedIndex As Integer = sortedList.IndexOf(entry)

            If i <> sortedIndex Then

                If misplacedEntryList.Contains(entry) = False Then

                    If sortedIndex > i Then

                        misplacedEntryList.Add(entry)

                        someList.Insert(sortedIndex, entry)
                        someList.RemoveAt(i)

                        'Adjust i because the item in its position has now changed and we don't want to skip it
                        i = i - 1

                    Else

                        'some contingency for when we're tracking lines
                        If LineCountList.Count > 1 Then
                            LineCountList.Insert(sortedIndex, LineCountList(i))
                            LineCountList.RemoveAt(i + 1)

                        End If
                        Dim recInd As Integer = originalPlacement.IndexOf(entry)
                        Dim sortInd As Integer = sortedList.IndexOf(entry)
                        Dim curPos As String = originalLines(recInd)
                        Dim sortPos As String = originalLines(sortInd)

                        If findType <> "Entry" Then
                            entry = findType & recInd + 1 & "=" & entry
                        End If

                        Console.WriteLine("Line: " & originalLines(recInd) & " - Error: '" & findType & "' Alphabetization. " & entry & " appears to be out of place." & Environment.NewLine & "Current  Position: Line " & curPos & Environment.NewLine & "Expected Position: Line " & sortPos & Environment.NewLine)
                        numErrs += 1

                        'move the entry to the left in our list
                        someList.Insert(sortedIndex, entry)
                        someList.RemoveAt(i + 1)

                        'jump back to the position from which we sorted the entry to make sure we remove any entries that were off by one because of it not having been moved yet
                        i = sortedIndex + 1

                    End If

                End If
            Else
                'If we have moved backwards because we moved an element left in the list, we want to remove anything that was off-by-one because of it from the misplaced entry list
                If misplacedEntryList.Contains(entry) = True Then
                    misplacedEntryList.Remove(entry)
                End If
            End If
        Next

        For Each entry As String In misplacedEntryList

            'Any remaining entries in the misplaced list were actually misplaced and not just out of order because of a misplacement
            Dim recInd As Integer = originalPlacement.IndexOf(entry)
            Dim sortInd As Integer = sortedList.IndexOf(entry)
            Dim curPos As String = originalLines(recInd)
            Dim sortPos As String = LineCountList(sortInd)

            If findType <> "Entry" Then
                entry = findType & recInd + 1 & "=" & entry
            End If

            Console.WriteLine("Line: " & originalLines(recInd) & " - Error: '" & findType & "' Alphabetization. " & entry & " appears to be out of place, possibly as a result of another out of place line." & Environment.NewLine & "Current Position: Line " & curPos & Environment.NewLine & "Expected Position: Line " & originalLines(recInd + 1) & Environment.NewLine)
            numErrs += 1
        Next

    End Sub

    Public Sub findAndReplNumbers(ByRef myChars() As Char, ByRef entry As String, ByRef originalNameList As List(Of String), ByRef renamedList As List(Of String), ByRef ListToBeSorted As List(Of String), ByRef givenSortedList As List(Of String), ByVal receivedIndex As Integer, ByVal sortedindex As Integer)

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

    Public Sub err2(ByVal linecount As Integer, ByVal err As String, ByVal command As String, ByVal expected As String, ByRef numErrs As Integer)
        Console.WriteLine("Line: " & linecount & " - Error: " & err & Environment.NewLine & "Expected: " & expected & Environment.NewLine & "Command: " & command & Environment.NewLine)
        numErrs += 1
    End Sub

    Public Sub err(ByVal linecount As Integer, ByVal err As String, ByVal command As String, ByRef numErrs As Integer)
        Console.WriteLine("Line: " & linecount & " - Error: " & err & Environment.NewLine & "Command: " & command & Environment.NewLine)
        numErrs += 1

    End Sub

    Public Sub cFormat(ByVal key As iniKey, ByVal keyString As String, ByRef keyNumber As Integer, ByRef numErrs As Integer, ByRef keyList As List(Of String))

        Dim command As String = key.ToString

        'Check for trailing whitespace
        If command.EndsWith(" ") Then
            err(key.lineNumber, "Detected unwanted whitepace at end of line.", key.ToString, numErrs)
        End If

        'Check for ending whitespace
        If key.ToString.StartsWith(" ") Then
            err(key.lineNumber, "Detected unwanted whitepace at beginning of line.", key.ToString, numErrs)
        End If

        'check for duplicates
        For Each item As String In keyList

            If item.Contains(key.value.ToLower) Then
                Console.WriteLine("Line: " & key.lineNumber & " - Error: Duplicate key found." & Environment.NewLine &
                                  "Command: " & command & Environment.NewLine &
                                  "Duplicates: " & keyString & keyList.IndexOf(item) + 1 & "=" & key.value.ToLower & Environment.NewLine)
                Exit For
            End If
        Next


        'make sure the current key is correctly numbered
        If Not command.Contains(keyString & keyNumber) Then
            err2(key.lineNumber, "'" & keyString & "' entry is incorrectly spelled or formatted.", command, keyString & keyNumber, numErrs)
        End If
        keyNumber = keyNumber + 1

        'make sure we don't have any dangly bits on the end of our key
        If command(command.Count - 1) = ";" Then
            err(key.lineNumber, "Trailing semicolon (;).", command, numErrs)
        End If

        'Do some formatting checks for environment variables
        If keyString = "DetectFile" Or keyString = "FileKey" Or keyString = "ExcludeKey" Then

            If command.Contains("%") Then
                Dim envars As New List(Of String)
                envars.AddRange(New String() {"UserProfile", "ProgramFiles", "RootDir", "WinDir", "AppData", "SystemDrive", "Documents", "ProgramData", "AllUsersProfile",
                            "Pictures", "Video", "CommonAppData", "LocalAppData", "CommonProgramFiles", "HomeDrive", "Music", "tmp", "Temp", "LocalLowAppData", "Public"})


                Dim varcheck As String() = command.Split(System.Convert.ToChar("%"))
                If varcheck.Count <> 3 And varcheck.Count <> 5 Then
                    err(key.lineNumber, "%EnvironmentVariables% must be surrounded on both sides by a single '%' character.", command, numErrs)
                End If

                If varcheck.Count = 3 Then
                    If Not envars.Contains(varcheck(1)) Then

                        Dim casingerror As Boolean = False
                        For Each var As String In envars

                            If varcheck(1).ToLower = var.ToLower Then
                                casingerror = True
                                err2(key.lineNumber, "Invalid CamelCasing on environment variable.", varcheck(1), var, numErrs)
                            End If

                        Next

                        If casingerror = False Then
                            err(key.lineNumber, "Misformatted or invalid environment variable.", command, numErrs)
                        End If
                    End If
                End If
            End If

        End If

    End Sub

    Public Function cValidity(key As iniKey, ByVal validList As List(Of String), numErrs As Integer) As Boolean

        'make sure we have a valid command
        Dim isValidCmd As Boolean = False
        For Each cmd As String In validList
            If key.name.ToLower.StartsWith(cmd) Then
                isValidCmd = True
                Exit For
            End If
        Next

        'if the command is invalid, print an error and move on to the next command
        If isValidCmd = False Then
            err(key.lineNumber, "Invalid command type detected.", key.ToString, numErrs)
            Return False
        End If
        Return True
    End Function


    Public Sub pLangSecRef(ByVal lsrKey As iniKey, ByRef numErrs As Integer)

        If lsrKey.value <> "" Then

            Dim validSecRefs As New List(Of String)
            validSecRefs.AddRange(New String() {"3021", "3022", "3023", "3024", "3025", "3026", "3027", "3028", "3029", "3030", "3031"})
            Dim hasValidSecRef As Boolean = False

            'make sure langsecref is spelled correctly
            If Not lsrKey.name.Contains("LangSecRef") Then
                err(lsrKey.lineNumber, "LangSecRef is incorrected spelled or formatted.", lsrKey.ToString, numErrs)
            End If

            'and make sure the number that follows it is valid
            For Each secref As String In validSecRefs
                If lsrKey.value.Contains(secref) Then
                    hasValidSecRef = True
                    Exit For
                End If

            Next
            'if the langsecref number is invalid, throw an error
            If Not hasValidSecRef Then
                err(lsrKey.lineNumber, "LangSecRef holds an invalid value.", lsrKey.ToString, numErrs)
            End If
        End If
    End Sub

    Public Sub pFileKey(ByRef keylist As List(Of iniKey), ByRef numErrs As Integer)
        Dim curFileKeyNum As Integer = 1
        Dim fileKeyList As New List(Of String)
        For Each key As iniKey In keylist

            Dim command As String = key.value
            fileKeyList.Add(command)
            cFormat(key, "FileKey", curFileKeyNum, numErrs, fileKeyList)

            Dim iteratorCheckerList() As String = Split(command, "|")

            'Pipe symbol checks
            If Not command.Contains("|") Then
                err(key.lineNumber, "Missing pipe (|) in 'FileKey.'", command, numErrs)
            End If
            If command.Contains(";|") Then
                err(key.lineNumber, "Semicolon (;) found before pipe (|).", command, numErrs)
            End If

            'check for incorrect spellings of RECURSE or REMOVESELF
            If iteratorCheckerList.Length > 2 Then
                If Not iteratorCheckerList(2).Contains("RECURSE") And Not iteratorCheckerList(2).Contains("REMOVESELF") Then

                    err(key.lineNumber, "'RECURSE' or 'REMOVESELF' entry is incorrectly spelled, or there are too many pipe (|) symbols.", command, numErrs)

                End If
            End If

            'check for missing pipe symbol on recurse and removeself
            If command.Contains("RECURSE") And Not command.Contains("|RECURSE") Then
                err(key.lineNumber, "Missing pipe (|) before 'RECURSE.'", command, numErrs)
            End If
            If command.Contains("REMOVESELF") And Not command.Contains("|REMOVESELF") Then
                err(key.lineNumber, "Missing pipe (|) before 'REMOVESELF.'", command, numErrs)
            End If

            'make sure VirtualStore folders point to the correct place
            If command.Contains("\VirtualStore\P") And (Not command.ToLower.Contains("programdata") And Not command.ToLower.Contains("program files*") And Not command.ToLower.Contains("program*")) Then
                err2(key.lineNumber, "Incorrect VirtualStore location.", command, "%LocalAppData%\VirtualStore\Program Files*\", numErrs)
            End If

            'backslash checks
            If command.Contains("%\|") Then
                err(key.lineNumber, "Backslash (\) found before pipe (|).", command, numErrs)
            End If
            If command.Contains("%") And Not command.Contains("%|") And Not command.Contains("%\") Then
                err(key.lineNumber, "Missing backslash (\) after %EnvironmentVariable%.", command, numErrs)
            End If

        Next
    End Sub

    Public Sub pDefault(key As iniKey, ByRef numErrs As Integer)

        'spell check
        If Not key.name.StartsWith("Default") And Not key.name.ToLower.Contains("detecto") Then
            err2(key.lineNumber, "'Default' incorrectly spelled.", key.name, "Default", numErrs)
        End If

        'make sure all entries are disabled by default
        If command.ToLower.Contains("true") Then
            err(key.lineNumber, "All entries should be disabled by default (Default=False).", key.ToString, numErrs)
        End If

    End Sub

    Public Sub pSection(section As iniSection, ByRef numErrs As Integer)

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

        Dim validCmds As New List(Of String)
        validCmds.AddRange(New String() {"specialdetect", "filekey", "regkey", "detect", "langsecref", "warning", "[", "default", "section", "excludekey"})

        For Each key As iniKey In section.keys

            'make sure our key contains a valid command
            If Not cValidity(key, validCmds, numErrs) Then
                Continue For
            End If

            If key.name.StartsWith("L") Then
                secRefKey = key
            End If

            If key.name.StartsWith("F") Then
                fileKeys.Add(key)

                'Make sure that FileKeys come before RegKeys in the section (style error but not syntax)
                If entryHasRegKeys And Not entryKeysOutOfOrder Then
                    err(key.lineNumber, "FileKeys should precede RegKeys.", key.ToString, numErrs)
                    entryKeysOutOfOrder = True
                End If

            End If

            If key.name.StartsWith("D") Then
                'We'll make the bold assumption here that all our defaults are spelled correctly, even if not formatted correctly. 
                If Not key.name.ToLower.Contains("default") And Not key.name.ToLower.Contains("detecto") Then
                    detectKeys.Add(key)
                Else
                    defaultKey = key
                End If
            End If

            If key.name.StartsWith("E") Then
                excludeKeys.Add(key)
            End If

            If key.name.StartsWith("R") Then
                regKeys.Add(key)
                If Not entryHasRegKeys Then
                    entryHasRegKeys = True
                End If
            End If
        Next

        'process our filekeys
        pFileKey(fileKeys, numErrs)

        Dim fkstrings As New List(Of String)
        Dim sortedFKStrings As New List(Of String)
        Dim fklinecounts As New List(Of String)
        For Each filekey As iniKey In fileKeys
            fklinecounts.Add(filekey.lineNumber.ToString)
            fkstrings.Add(filekey.value)
        Next
        sortedFKStrings.AddRange(fkstrings)
        replaceAndSort(fkstrings, sortedFKStrings, "|", " \ \", "keys")
        findOutOfPlace(fkstrings, sortedFKStrings, "FileKey", numErrs, fklinecounts)

        'process our regkeys
        pRegKey(regKeys, numErrs)

        Dim rkstrings As New List(Of String)
        Dim sortedRKStrings As New List(Of String)
        Dim rklinecounts As New List(Of String)
        For Each filekey As iniKey In regKeys
            rklinecounts.Add(filekey.lineNumber.ToString)
            rkstrings.Add(filekey.value)
        Next
        sortedRKStrings.AddRange(rkstrings)
        replaceAndSort(rkstrings, sortedRKStrings, "|", " \ \", "keys")
        findOutOfPlace(rkstrings, sortedRKStrings, "RegKey", numErrs, rklinecounts)

        'process our excludekeys
        pExcludeKey(excludeKeys, numErrs)

        Dim ekstrings As New List(Of String)
        Dim sortedEKStrings As New List(Of String)
        Dim eklinecounts As New List(Of String)
        For Each excludeKey As iniKey In excludeKeys
            eklinecounts.Add(excludeKey.lineNumber.ToString)
            ekstrings.Add(excludeKey.value)
        Next
        sortedEKStrings.AddRange(ekstrings)
        replaceAndSort(ekstrings, sortedEKStrings, "|", " \ \", "keys")
        findOutOfPlace(ekstrings, sortedEKStrings, "ExcludeKey", numErrs, eklinecounts)

        'process our detect keys (if any)
        pDetect(detectKeys, numErrs)

        'process our langsecref (if any)
        pLangSecRef(secRefKey, numErrs)

        'process our default key, throw an error if it's absent 
        If defaultKey.value <> "" Then
            pDefault(defaultKey, numErrs)
        Else
            err(section.startingLineNumber, "Default state missing", section.name, numErrs)
        End If

    End Sub

    Private Sub pDetectFile(ByRef key As iniKey, ByRef numErrs As Integer)

        'check for trailing backslashes
        Dim command As String = key.value
        If command(command.Count - 1) = "\" Then
            err(key.lineNumber, "Trailing backslash on DetectFile.", key.ToString, numErrs)
        End If

        'check for nested wildcards 
        If command.Contains("*") Then
            Dim splitDir As String() = command.Split(System.Convert.ToChar("\"))
            For i As Integer = 0 To splitDir.Count - 1
                If splitDir.Last = Nothing Then Continue For
                If splitDir(i).Contains("*") And i <> splitDir.Count - 1 Then
                    err(key.lineNumber, "Nested wildcard found in DetectFile.", key.ToString, numErrs)
                End If
            Next
        End If

        'Make sure that DetectFile paths point to a filesystem location
        If Not command.StartsWith("%") And Not command.Contains(":\") Then
            err(key.lineNumber, "'DetectFile' can only be used for file system paths.", key.ToString, numErrs)
        End If

    End Sub

    Private Sub pRegDetect(ByRef key As iniKey, ByRef numErrs As Integer)

        Dim command As String = key.ToString
        'Make sure that detect paths point to a registry location
        If (command.Contains("=%") Or command.Contains("=C:\")) Or (Not command.Contains("=HKLM") And Not command.Contains("=HKC") And Not command.Contains("=HKU")) Then
            err(key.lineNumber, "'Detect' can only be used for registry keys paths.", key.ToString, numErrs)
        End If

    End Sub

    Private Sub pDetect(ByRef iKeyList As List(Of iniKey), ByRef numErrs As Integer)

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

            If key.name.Length > 8 Then
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

            'check spelling/formatting
            If Not key.name.Contains(detectType) Then
                err(key.lineNumber, "Misformatted '" & detectType & "'.", command, numErrs)
            End If

            'check for duplicates
            If keyList.Contains(command) Then
                Console.WriteLine("Line: " & key.lineNumber & " - Error: Duplicate command found. " & Environment.NewLine & "Command: " & key.ToString & Environment.NewLine & "Duplicates: " & detectType & keyList.IndexOf(command) + 1 & "=" & command & Environment.NewLine)
            Else
                keyList.Add(command)
            End If

            'Check whether our first Detect or DetectFile is trailed by a number
            If number = 1 Then
                If key.ToString.StartsWith(detectType & "=") Or (Not key.ToString.StartsWith(detectType & "=") And Not key.ToString.StartsWith(detectType & "1")) Then
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
            If number = 1 And (Not key.ToString.Contains(detectType & "1=") And Not key.ToString.Contains(detectType & "=")) Then
                err2(key.lineNumber, "'" & detectType & "' numbering.", command, "'" & detectType & "' or '" & detectType & "1.'", numErrs)
            End If

            'If we're on our second detect, make sure the first one had a 1 
            If number = 2 And firstNumberStatus = False Then
                err(key.lineNumber - 1, "'" & detectType & number & "' detected without preceding '" & detectType & number - 1 & ".'", key.ToString, numErrs)
            End If

            'if we only have one detect entry, make sure it doesnt have a number
            If iKeyList.Count = 1 Then
                For i As Integer = 0 To 9
                    If key.name.Contains(i.ToString) Then
                        err(key.lineNumber, "Detected unneeded numbering", key.ToString, numErrs)
                        Exit For
                    End If
                Next
            End If

            'otherwise, make sure our numbers match up
            If number > 1 Then
                If key.name.Contains(detectType & number) = False Then
                    err2(key.lineNumber, "'" & detectType & "' entry is incorrectly numbered.", key.ToString, "'" & detectType & number & "'", numErrs)
                End If
            End If
            number += 1

            'update our tracking variables as needed
            If detectType.Equals("DetectFile") Then
                curDFNum = number
                dfkeylist = keyList
                pDetectFile(key, numErrs)
            Else
                curDNum = number
                dkeylist = keyList
                pRegDetect(key, numErrs)
            End If
        Next

        'check alphabetization
        Dim sortedDStrings As New List(Of String)
        Dim sortedDFStrings As New List(Of String)

        sortedDStrings.AddRange(dkeylist)
        sortedDFStrings.AddRange(dfkeylist)
        replaceAndSort(dkeylist, sortedDStrings, "|", " \ \", "keys")
        findOutOfPlace(dkeylist, sortedDStrings, "Detect", numErrs, dlclist)

        replaceAndSort(dfkeylist, sortedDFStrings, "|", " \ \", "keys")
        findOutOfPlace(dfkeylist, sortedDFStrings, "DetectFile", numErrs, dflclist)

    End Sub

    Private Sub pRegKey(ByVal keyList As List(Of iniKey), ByRef numErrs As Integer)

        'Ensure that each RegKey points to a valid registry location
        For Each key As iniKey In keyList
            If Not key.ToString.Contains("=HKLM") And Not key.ToString.Contains("=HKC") And Not key.ToString.Contains("=HKU") Then
                err(key.lineNumber, "'RegKey' can only be used for registry key paths.", key.ToString, numErrs)
            End If
        Next

    End Sub

    Private Sub pExcludeKey(ByRef keyList As List(Of iniKey), ByRef numErrs As Integer)

        'Make sure any FILE exclude paths have a backslash before their pipe symbol
        For Each key As iniKey In keyList

            Dim command As String = key.value
            Dim iteratorCheckerList() As String = Split(command, "|")
            If iteratorCheckerList(0).Contains("FILE") Then
                Dim endingslashchecker() As String = Split(command, "\|")
                If endingslashchecker.Count = 1 Then
                    err(key.lineNumber, "Missing backslash (\) before pipe (|) in ExcludeKey.", command, numErrs)
                End If
            End If
        Next

    End Sub

    Public Function findCommentLine(sFile As iniFile, com As String) As Integer

        'find the line number of a particular comment by its string
        For Each comment As iniComment In sFile.comments
            If comment.comment = com Then
                Return comment.lineNumber
            End If
        Next

        Return -1

    End Function

End Module

Public Module iniFileHandler
    Class iniFile

        Public name As String
        Public sections As New Dictionary(Of String, iniSection)
        'may want to consder changing this to a dictionary too
        Public comments As New List(Of iniComment)

        Public Sub New()
            name = ""
            sections = New Dictionary(Of String, iniSection)
            comments = New List(Of iniComment)
        End Sub

        Public Sub New(name As String)
            Me.New(Environment.CurrentDirectory, name)
        End Sub

        Public Sub New(path As String, name As String)

            name = name

            Dim r As IO.StreamReader
            Try
                Dim sectionToBeBuilt As New List(Of String)
                Dim lineTrackingList As New List(Of Integer)

                r = New StreamReader(path & "\" & name)
                Dim lineCount As Integer = 1
                Do While (r.Peek() > -1)
                    Dim currentLine As String = r.ReadLine.ToString

                    If currentLine.StartsWith(";") Then
                        Dim newCom As New iniComment(currentLine, lineCount)

                        Me.comments.Add(newCom)
                    End If

                    If currentLine.Trim <> "" And Not currentLine.StartsWith(";") Then
                        sectionToBeBuilt.Add(currentLine)
                        lineTrackingList.Add(lineCount)
                    Else

                        If Not sectionToBeBuilt.Count < 2 Then

                            Dim sectionHolder As New iniSection(sectionToBeBuilt, lineTrackingList)
                            Try
                                sections.Add(sectionHolder.name, sectionHolder)

                            Catch ex As Exception
                                Console.WriteLine(ex.Message & Environment.NewLine & "Failure occurred during iniFile construction at line: " & Me.sections.Last.Value.endingLineNumber + 2)
                                Console.WriteLine("Likely a duplicate entry name")
                                Console.WriteLine()
                            End Try
                            sectionToBeBuilt.Clear()
                            lineTrackingList.Clear()
                        End If
                    End If
                    lineCount += 1
                Loop
                If sectionToBeBuilt.Count <> 0 Then
                    Dim sectionHolder As New iniSection(sectionToBeBuilt, lineTrackingList)
                    sections.Add(sectionHolder.name, sectionHolder)
                    sectionToBeBuilt.Clear()
                    lineTrackingList.Clear()
                End If
                r.Close()

            Catch ex As Exception
                Console.WriteLine(ex.Message & " failure detected during iniFile construction")
            End Try

        End Sub

        Public Sub compareTo(ByVal secondFile As iniFile)

            Dim comparedList As New List(Of String)

            For Each i As String In sections.Keys

                If secondFile.sections.ContainsKey(sections(i).name) And Not comparedList.Contains(sections(i).name) Then
                    Dim mySection As iniSection = sections(i)
                    Dim sSection As iniSection = secondFile.sections(i)
                    mySection.compareTo(sSection)
                    comparedList.Add(mySection.name)
                ElseIf Not secondFile.sections.ContainsKey(sections(i).name) Then
                    Console.WriteLine(sections(i).name & " has been removed.")
                    Console.WriteLine()
                    sections(i).ToString()
                    Console.WriteLine("----------------------------------------------------------------------------------------------------------------------")
                End If

            Next

            For Each i As String In secondFile.sections.Keys

                If Not sections.ContainsKey(i) Then
                    Console.WriteLine(secondFile.sections(i).name & " has been added.")
                    Console.WriteLine()
                    secondFile.sections(i).ToString()
                    Console.WriteLine("----------------------------------------------------------------------------------------------------------------------")

                End If
            Next
        End Sub

    End Class

    Class iniSection
        Public startingLineNumber As Integer
        Public endingLineNumber As Integer
        Public name As String
        Public keys As New List(Of iniKey)

        Public Sub New(ByVal listOfLines As List(Of String), listOfLineCounts As List(Of Integer))

            Dim tmp1 As String() = listOfLines(0).Split(System.Convert.ToChar("["))
            Dim tmp2 As String() = tmp1(1).Split(System.Convert.ToChar("]"))
            name = tmp2(0)
            startingLineNumber = listOfLineCounts(0)
            endingLineNumber = listOfLineCounts(listOfLineCounts.Count - 1)

            For i As Integer = 1 To listOfLines.Count - 1
                keys.Add(New iniKey(listOfLines(i), listOfLineCounts(i)))
            Next

        End Sub

        Public Sub diffSections(secondSection As iniSection)

            Console.WriteLine(name & " has been modified.")
            Console.WriteLine()
            secondSection.ToString()
            Console.WriteLine("----------------------------------------------------------------------------------------------------------------------")

        End Sub

        Public Sub compareTo(secondSection As iniSection)

            If keys.Count <> secondSection.keys.Count Then

                diffSections(secondSection)

            Else
                Dim isdiff As Boolean = False
                For i As Integer = 0 To keys.Count - 1

                    If Not keys(i).compareTo(secondSection.keys(i)) Then
                        isdiff = True
                    End If
                Next

                If isdiff Then
                    diffSections(secondSection)
                End If

            End If
        End Sub

        Public Overrides Function ToString() As String
            Console.WriteLine("[" & name & "]")

            For Each key As iniKey In keys
                Console.WriteLine(key.ToString())
            Next

            Return ""

        End Function

    End Class

    Class iniKey
        Public name As String
        Public value As String
        Public lineNumber As Integer

        Public Sub New()

            name = ""
            value = ""
            lineNumber = 0

        End Sub

        Public Sub New(ByVal line As String, ByVal count As Integer)

            Try
                Dim splitLine As String() = line.Split(System.Convert.ToChar("="))
                name = splitLine(0)
                value = splitLine(1)
                lineNumber = count
            Catch ex As Exception
                Console.WriteLine(ex)
            End Try

        End Sub

        Public Overrides Function ToString() As String
            Return name & "=" & value
        End Function

        Public Function lineString() As String
            Return "Line: " & lineNumber & " - " & name & "=" & value
        End Function

        Public Function compareTo(secondKey As iniKey) As Boolean

            If Not Me.ToString.Equals(secondKey.ToString) Then
                Return False
            Else
                Return True
            End If
        End Function

    End Class

    Class iniComment
        Public comment As String
        Public lineNumber As Integer

        Public Sub New(c As String, l As Integer)
            comment = c
            lineNumber = l
        End Sub

    End Class
End Module

Module diff

    Public Function fileMaker(inLine As String) As iniFile
        If inLine.Contains(",") Then
            Dim inSplit As String() = inLine.Split(System.Convert.ToChar(","))
            Dim path As String = inSplit(0)
            Dim name As String = inSplit(1)
            Return New iniFile(path, name)
        Else
            Return New iniFile(inLine)
        End If

    End Function

    Public Sub main()
        Console.Clear()
        Console.WriteLine("*--------------------------------------------------------------------------------------------------*")
        Console.WriteLine("*                                     Diff                                                         *")
        Console.WriteLine("*                                                                                                  *")
        Console.WriteLine("*                This tool will output the diff between two winapp2 files                          *")
        Console.WriteLine("*                If both are in the same folder as winapp2ool.exe, you need only                   *")
        Console.WriteLine("*                enter their filenames below. Otherwise, enter the path and the filename           *")
        Console.WriteLine("*                separated by a comma (,) (eg. C:\Program Files\CCleaner,winapp2.ini )             *")
        Console.WriteLine("*                The first file should be the older version, and the second the newer.             *")
        Console.WriteLine("*--------------------------------------------------------------------------------------------------*")

        Console.Write("Enter first file name or first path and file name: ")
        Dim inLine As String = Console.ReadLine()
        Dim firstFile As New iniFile
        Dim secondFile As New iniFile
        firstFile = fileMaker(inLine)

        Console.Write("Enter the second file name or second file path and file name: ")
        inLine = Console.ReadLine()
        secondFile = fileMaker(inLine)
        Try
            Dim fver As String = firstFile.comments(0).comment.ToString
            fver = fver.Split(System.Convert.ToChar(";"))(1)
            Dim sver As String = secondFile.comments(0).comment.ToString
            sver = sver.Split(System.Convert.ToChar(";"))(1)
            Console.WriteLine("Changes made between " & fver & " and" & sver)

            firstFile.compareTo(secondFile)
        Catch ex As Exception
            Console.WriteLine("Error: " & ex.ToString)
        End Try
        Console.WriteLine("End of diff. Press any key to return to the menu.")
        Console.ReadKey()
    End Sub

End Module

Public Module trim

    Public Sub main()

        Console.Clear()

        Console.WriteLine("*--------------------------------------------------------------------------------------------------*")
        Console.WriteLine("*                                             Trim                                                 *")
        Console.WriteLine("*                                                                                                  *")
        Console.WriteLine("*                       This tool will trim winapp2.ini down to contain only                       *")
        Console.WriteLine("*                       entries relevant to your machine, greatly reducing both                    *")
        Console.WriteLine("*                       application load time and the winapp2.ini filesize.                        *")
        Console.WriteLine("*                       You must launch winapp2ool as an administrator and have                    *")
        Console.WriteLine("*                       winapp2.ini in the same folder as winapp2ool for this feature.             *")
        Console.WriteLine("*--------------------------------------------------------------------------------------------------*")

        Console.WriteLine("Make sure winapp2.ini is in the same folder as winapp2ool.exe and enter Y to begin trim, or press any key to exit.")
        'load winapp2.ini into memory
        Dim winappfile As New iniFile("winapp2.ini")

        'create a list of sections that belong in the trimmed file
        Dim trimmedfile As New List(Of iniSection)

        'get the windows version so we can test properly our DetectOS
        Dim winveri As Double = getWinVer()

        Dim input As String = Console.ReadLine()
        Dim cEntries As New List(Of iniSection)
        Dim fxEntries As New List(Of iniSection)
        Dim tbEntries As New List(Of iniSection)

        If input.ToLower.Equals("y") Then
            Console.WriteLine("Trimming...")

            Dim detChrome As New List(Of String)
            detChrome.AddRange(New String() {"%AppData%\ChromePlus\chrome.exe", "%LocalAppData%\Chromium\Application\chrome.exe", "%LocalAppData%\Chromium\chrome.exe", "%LocalAppData%\Flock\Application\flock.exe", "%LocalAppData%\Google\Chrome SxS\Application\chrome.exe",
                               "%LocalAppData%\Google\Chrome\Application\chrome.exe", "%LocalAppData%\RockMelt\Application\rockmelt.exe", "%LocalAppData%\SRWare Iron\iron.exe", "%ProgramFiles%\Chromium\Application\chrome.exe", "%ProgramFiles%\SRWare Iron\iron.exe",
                               "%ProgramFiles%\Chromium\chrome.exe", "%ProgramFiles%\Flock\Application\flock.exe", "%ProgramFiles%\Google\Chrome SxS\Application\chrome.exe", "%ProgramFiles%\Google\Chrome\Application\chrome.exe", "%ProgramFiles%\RockMelt\Application\rockmelt.exe",
                               "Software\Chromium", "Software\SuperBird", "Software\Torch", "Software\Vivaldi"})

            For i As Integer = 0 To winappfile.sections.Count - 1

                Dim cursection As iniSection = winappfile.sections.Values(i)
                Dim hasDetOS As Boolean = False
                Dim hasMetDetOS As Boolean = True
                Dim hasDets As Boolean = False
                Dim exists As Boolean = False
                For Each key As iniKey In cursection.keys

                    If exists Then
                        Continue For
                    End If

                    If key.name.StartsWith("Detect") And Not key.name.Equals("DetectOS") Then
                        hasDets = True
                        If checkExist(key.value) Then
                            trimmedfile.Add(cursection)
                            Exit For
                        End If

                    End If

                    If key.name.Equals("DetectOS") Then

                        hasDetOS = True
                        hasMetDetOS = detOSCheck(key.value, winveri)

                        'move on if we don't meet the DetectOS criteria 
                        If Not hasMetDetOS Then
                            Exit For
                        End If
                    End If

                    If key.name.StartsWith("Special") Then
                        hasDets = True

                        'handle our SpecialDetects
                        Select Case key.value
                            Case "DET_CHROME"
                                For Each path As String In detChrome
                                    If checkExist(path) Then
                                        cEntries.Add(cursection)
                                        Exit For
                                    End If
                                Next
                                Exit For
                            Case "DET_MOZILLA"
                                If checkExist("%AppData%\Mozilla\Firefox") Then
                                    fxEntries.Add(cursection)
                                    Exit For
                                End If
                            Case "DET_THUNDERBIRD"
                                If checkExist("%AppData%\Thunderbird") Then
                                    tbEntries.Add(cursection)
                                    Exit For
                                End If
                            Case "DET_OPERA"
                                If checkExist("%AppData%\Opera Software") Then
                                    trimmedfile.Add(cursection)
                                    Exit For
                                End If
                        End Select
                    End If
                Next

                If hasDetOS Then
                    If Not hasDets And hasMetDetOS Then
                        trimmedfile.Add(cursection)
                    End If
                Else
                    If Not hasDets Then
                        trimmedfile.Add(cursection)
                    End If
                End If

            Next
            Console.WriteLine("Trimming complete.")
            Console.WriteLine("Initial number Of entries: " & winappfile.sections.Count)
            Console.WriteLine("Number of entries after trimming: " & trimmedfile.Count)
            Console.WriteLine("Enter '1' to save trimmed file as a new file, or '2' to overwrite your existing winapp2.ini file. Press any other key to exit.")
            Dim inStr As String = Console.ReadLine()
            Dim fname As String
            If inStr.Equals("1") Or inStr.Equals("2") Then
                Try
                    If inStr.Equals("1") Then
                        fname = "\winapp2-trimmed.ini"
                    Else
                        fname = "\winapp2.ini"
                    End If
                    Dim file As New System.IO.StreamWriter(Environment.CurrentDirectory & fname, False)

                    For i As Integer = 0 To 8
                        file.WriteLine(winappfile.comments(i).comment)
                    Next

                    writeCustomSectionToFile(cEntries, file)

                    file.WriteLine(winappfile.comments(9).comment)
                    file.WriteLine(winappfile.comments(10).comment)
                    file.WriteLine(winappfile.comments(11).comment)

                    writeCustomSectionToFile(fxEntries, file)

                    file.WriteLine(winappfile.comments(12).comment)
                    file.WriteLine(winappfile.comments(13).comment)
                    file.WriteLine(winappfile.comments(14).comment)

                    writeCustomSectionToFile(tbEntries, file)

                    file.WriteLine(winappfile.comments(15).comment)
                    file.WriteLine()

                    For Each section As iniSection In trimmedfile
                        file.WriteLine("[" & section.name & "]")
                        For Each key As iniKey In section.keys
                            file.WriteLine(key.ToString)
                        Next
                        If Not trimmedfile.IndexOf(section).Equals(trimmedfile.Count - 1) Then
                            file.WriteLine()
                        End If
                    Next

                    file.Close()
                    Console.WriteLine("*--------------------------------------------------------------------------------------------------*")
                    Console.WriteLine("Finished trimming winapp2.ini, press any key to exit.")
                    Console.ReadKey()
                Catch ex As Exception
                    Console.WriteLine(ex.Message)
                End Try
            End If

        End If

    End Sub

    Public Function checkExist(key As String) As Boolean

        Dim dir As String = key
        Dim isProgramFiles As Boolean = False

        'make sure we get the proper path for environment variables
        If dir.Contains("%") Then
            Dim splitDir As String() = dir.Split(System.Convert.ToChar("%"))
            Dim var As String = splitDir(1)
            Dim envDir As String = Environment.GetEnvironmentVariable(var)
            Select Case var
                Case "ProgramFiles"
                    isProgramFiles = True
                Case "Documents"
                    envDir = Environment.GetEnvironmentVariable("UserProfile")
                    envDir += "\Documents"
                Case "CommonAppData"
                    envDir = Environment.GetEnvironmentVariable("ProgramData")
            End Select
            dir = envDir + splitDir(2)
        End If

        'Observe the registry paths
        If dir.StartsWith("HK") Then
            Dim splitDir As String() = dir.Split(System.Convert.ToChar("\"))
            Try
                Select Case splitDir(0)
                    Case "HKCU"
                        dir = dir.Replace("HKCU\", "")
                        If Microsoft.Win32.Registry.CurrentUser.OpenSubKey(dir, True) IsNot Nothing Then
                            Return True
                        End If
                    Case "HKLM"
                        dir = dir.Replace("HKLM\", "")
                        If Microsoft.Win32.Registry.LocalMachine.OpenSubKey(dir, True) IsNot Nothing Then
                            Return True
                        Else
                            Dim rDir As String = dir.ToLower.Replace("software\", "Software\WOW6432Node\")
                            If Microsoft.Win32.Registry.LocalMachine.OpenSubKey(rDir, True) IsNot Nothing Then
                                Return True
                            Else
                                'Console.WriteLine()
                            End If
                        End If
                    Case "HKU"
                        dir = dir.Replace("HKU\", "")
                        If Microsoft.Win32.Registry.Users.OpenSubKey(dir, True) IsNot Nothing Then
                            Return True
                        End If
                    Case "HKCR"
                        dir = dir.Replace("HKCR\", "")
                        If Microsoft.Win32.Registry.ClassesRoot.OpenSubKey(dir, True) IsNot Nothing Then
                            Return True
                        End If
                End Select
            Catch ex As Exception
                'The most common exception here is a permissions one, so assume true if we hit 
                'because a permissions exception implies the key exists anyway.
                Return True
            End Try
        End If
        Try
            'check out those file/folder paths
            If Directory.Exists(dir) Or File.Exists(dir) Then
                Return True
            End If

            'if we didn't find it and we're looking in Program Files, check the (x86) directory
            If isProgramFiles Then

                'it's unlike that (m)any people have this configured differently, but it may be better
                'to query the env var for %ProgramFiles(x86)% for this. 
                dir = dir.Replace("Program Files", "Program Files (x86)")
                If Directory.Exists(dir) Or File.Exists(dir) Then
                    Return True
                End If
            End If

        Catch ex As Exception
            Console.WriteLine(ex.ToString)
        End Try

        Return False
    End Function

    'This function is for writing the chrome/firefox/thunderbird sections back into the file so we don't produce a poorly formatted ini
    Public Sub writeCustomSectionToFile(entryList As List(Of iniSection), file As System.IO.StreamWriter)
        file.WriteLine()

        If entryList.Count > 0 Then
            For Each section As iniSection In entryList
                file.WriteLine("[" & section.name & "]")
                For Each key As iniKey In section.keys
                    file.WriteLine(key.ToString)
                Next
                file.WriteLine()
            Next
        End If
    End Sub

    Private Function detOSCheck(value As String, winveri As Double) As Boolean
        Dim splitKey As String() = value.Split(System.Convert.ToChar("|"))

        If value.StartsWith("|") Then
            If winveri > Double.Parse(splitKey(1)) Then
                Return False
            Else
                Return True
            End If
        Else
            If winveri < Double.Parse(splitKey(0)) Then
                Return False
            Else
                Return True
            End If
        End If
    End Function

    Public Function getWinVer() As Double
        Dim winver As String = My.Computer.Info.OSFullName.ToString
        Dim winveri As Double = 0.0
        If winver.Contains("XP") Then
            winveri = 5.1
        End If
        If winver.Contains("Vista") Then
            winveri = 6.0
        End If
        If winver.Contains("7") Then
            winveri = 6.1
        End If
        If winver.Contains("8") And Not winver.Contains("8.1") Then
            winveri = 6.2
        End If
        If winver.Contains("8.1") Then
            winveri = 6.3
        End If
        If winver.Contains("10") Then
            winveri = 10.0
        End If

        Return winveri
    End Function
End Module

'here the line-by-line implementation actually works fairly well here since our process is simple,
'so this likely wont be refactored to use inifilehandler
Module ccinidebug
    Sub Main()

        Console.Clear()
        Console.WriteLine("*--------------------------------------------------------------------------------------------------*")
        Console.WriteLine("*                                           ccinidebug                                             *")
        Console.WriteLine("*                                                                                                  *")
        Console.WriteLine("*                       This tool will sort ccleaner.ini alphabetically                            *")
        Console.WriteLine("*                       And also offer To remove outdated winapp2.ini from it                      *")
        Console.WriteLine("*        make sure both winapp2.ini And ccleaner.ini are In the same folder As winapp2ool.exe      *")
        Console.WriteLine("*                       If the current folder Is the Program Files directory,                      *")
        Console.WriteLine("*                    you may need To relaunch winapp2ool.exe As an administrator                   *")
        Console.WriteLine("*--------------------------------------------------------------------------------------------------*")

        Dim lines As New ArrayList()
        Dim trimmedCCIniEntryList As New List(Of String)
        Dim trimmedWA2IniEntryList As New List(Of String)
        Dim entriesToPrune As New List(Of String)
        If File.Exists(Environment.CurrentDirectory & "\ccleaner.ini") = False Then
            Console.WriteLine("ccleaner.ini file could Not be located In the current working directory (" & Environment.CurrentDirectory & ") Then")
            Console.ReadKey()
            End
        End If

        Dim r As IO.StreamReader
        Try
            r = New IO.StreamReader(Environment.CurrentDirectory & "\ccleaner.ini")
            Do While (r.Peek() > -1)
                Dim currentLine As String = r.ReadLine.ToString
                If currentLine.Trim <> "" Then
                    lines.Add(currentLine)
                End If
                If currentLine.StartsWith("(App)") Then
                    Dim tmp1 As String() = Split(currentLine, "(App)")
                    Dim tmp2 As String() = Split(tmp1(1), "=")
                    If tmp2(0).Contains("*") Then
                        trimmedCCIniEntryList.Add(tmp2(0))
                    End If
                End If
            Loop
            r.Close()
            lines.Sort()
            lines.Remove("[Options]")
            lines.Insert(0, "[Options]")
            Console.WriteLine("Press Y To prune stale winapp2.ini entries from ccleaner.ini")
            Try
                If Console.ReadLine().ToLower = "y" Then
                    Dim w As IO.StreamReader
                    w = New IO.StreamReader(Environment.CurrentDirectory & "\winapp2.ini")
                    Do While (w.Peek > -1)
                        Dim curWA2Line As String = w.ReadLine.ToString()

                        If curWA2Line.StartsWith("[") Then
                            curWA2Line = curWA2Line.Remove(0, 1)
                            curWA2Line = curWA2Line.Remove(curWA2Line.Length - 1)
                            trimmedWA2IniEntryList.Add(curWA2Line)
                        End If
                    Loop

                    w.Close()
                    For Each entry As String In trimmedCCIniEntryList
                        If Not trimmedWA2IniEntryList.Contains(entry) Then
                            entriesToPrune.Add(entry)
                        End If
                    Next

                    Dim linesCopy As New ArrayList()
                    linesCopy.AddRange(lines)
                    Console.WriteLine("The following stale entries will be pruned: ")
                    For Each item As String In entriesToPrune
                        Console.WriteLine(item)
                        For Each entry As String In lines

                            If entry.Contains(item) Then
                                linesCopy.Remove(entry)
                            End If
                        Next
                    Next
                    Console.WriteLine("*--------------------------------------------------------------------------------------------------*")
                    lines = linesCopy
                End If
            Catch ex As Exception
                Console.WriteLine(ex.Message)
            End Try
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try
        Console.WriteLine("Press Y to sort and save ccleaner.ini, or any other key to exit.")

        If Console.ReadLine.ToString.ToLower = "y" Then
            Console.WriteLine("Modifying ccleaner.ini...")
            Try
                Dim file As New System.IO.StreamWriter(Environment.CurrentDirectory & "\ccleaner.ini", False)

                For Each line As String In lines
                    file.WriteLine(line.ToString)
                Next
                file.Close()
                Console.WriteLine("*--------------------------------------------------------------------------------------------------*")
                Console.WriteLine("Finished modifying ccleaner.ini, press any key to exit.")
                Console.ReadKey()
            Catch ex As Exception
                Console.WriteLine(ex.Message)
            End Try
        End If
    End Sub
End Module