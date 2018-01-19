Option Strict On
Imports System.IO
Imports System.Net

Module Module1
    Dim version As Double = 0.35
    Dim checkedForUpdates As Boolean = False
    Dim updateIsAvail As Boolean = False
    Public Sub printMenu()

        Console.WriteLine("*--------------------------------------------------------------------------------------------------*")
        Console.WriteLine("*                   Winapp2ool - A multitool for winapp2.ini and related files                     *")
        Console.WriteLine("*                                                                                                  *")
        Console.WriteLine("*                                     Menu: Enter a number to select                               *")
        If updateIsAvail Then
            Console.WriteLine("*                                                                                                  *")
            Console.WriteLine("*                                 An update is available for winapp2ool                            *")
            Console.WriteLine("*                                                                                                  *")
        End If
        Console.WriteLine("*                                                                                                  *")
        Console.WriteLine("* 0. Exit             - Exits the application                                                      *")
        Console.WriteLine("* 1. WinappDebug      - Loads the WinappDebug tool to check for errors in winapp2.ini              *")
        Console.WriteLine("* 2. ccinidebug       - Loads the ccinidebug tool to sort and trim ccleaner.ini                    *")
        Console.WriteLine("* 3. Diff             - Loads the diff tool to observe the changes between two winapp2.ini files   *")
        Console.WriteLine("* 4. Trim             - Loads the trim tool to debloat winapp2.ini for your system                 *")
        Console.WriteLine("* 5. Download         - Download files from the Winapp2 GitHub (including winapp2ool!)             *")
        Console.WriteLine("*--------------------------------------------------------------------------------------------------*")
        Console.Write("Enter a number: ")

    End Sub

    Public Sub checkUpdates()
        Dim address As String = "https://raw.githubusercontent.com/MoscaDotTo/Winapp2/master/Tools/beta/version.txt"
        Dim client As WebClient = New WebClient()
        Dim reader As StreamReader = New StreamReader(client.OpenRead(address))
        Dim latestVer As String = reader.ReadToEnd()
        Dim lvNum As Double = Convert.ToDouble(latestVer)
        If lvNum > version Then
            updateIsAvail = True
        End If
        checkedForUpdates = True
    End Sub

    Sub Main()
        checkUpdates()
        Dim exitCode As Boolean = False
        printMenu()
        Dim cki As String = Console.ReadLine
        Do Until exitCode = True
            Select Case cki
                Case "0"
                    exitCode = True
                    Console.WriteLine("Exiting...")
                    Environment.Exit(1)
                Case "1"
                    WinappDebug.Main()
                    Console.WriteLine("*--------------------------------------------------------------------------------------------------*")
                    Console.WriteLine("*********************************Finished running WinappDebug***************************************")
                    printMenu()
                    cki = Console.ReadLine()
                Case "2"
                    ccinidebug.Main()
                    Console.WriteLine(" * ------------------------------------------------------------------------------------------------*")
                    Console.WriteLine(" ********************************* Finished running ccinidebug**************************************")
                    printMenu()
                    cki = Console.ReadLine()
                Case "3"
                    diff.main()
                    Console.WriteLine("*--------------------------------------------------------------------------------------------------*")
                    Console.WriteLine("*************************************Finished running Diff******************************************")
                    printMenu()
                    cki = Console.ReadLine()
                Case "4"
                    trim.main()
                    Console.WriteLine("*--------------------------------------------------------------------------------------------------*")
                    Console.WriteLine("************************************Finished running trim*******************************************")
                    printMenu()
                    cki = Console.ReadLine()
                Case "5"
                    downloader.main()
                    Console.WriteLine("*--------------------------------------------------------------------------------------------------*")
                    Console.WriteLine("************************************Finished running downloader*************************************")
                    Console.WriteLine()
                    printMenu()
                    cki = Console.ReadLine()
                Case Else
                    Console.Write(Environment.NewLine & "Invalid input. Please Try again: ")
                    cki = Console.ReadLine()
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
        Console.WriteLine("*             This tool will check winapp2.ini For common syntax And style errors.                 *")
        Console.WriteLine("*                                                                                                  *")
        Console.WriteLine("*                                          Menu:                                                   *")
        Console.WriteLine("* 0. Exit                     - Return to the winapp2ool menu.                                     *")
        Console.WriteLine("* 1. Run (default)            - Run with the default settings                                      *")
        Console.WriteLine("* 2. Run (custom)             - Run with an option to provide the path and filename                *")
        Console.WriteLine("*--------------------------------------------------------------------------------------------------*")
        Console.Write("Enter a number: ")

        Dim cki As String = Console.ReadLine()
        Dim exitCode As Boolean = False
        Dim cfile As New iniFile
        Do While exitCode = False
            Try
                Select Case cki
                    Case "0"
                        Console.WriteLine("Returning to menu...")
                        exitCode = True
                    Case "1"
                        cfile = New iniFile("winapp2.ini")
                        debug(cfile)
                        exitCode = True
                    Case "2"
                        Console.Write("Enter the directory of your file, or leave blank to use default (current folder): ")
                        Dim path As String = Console.ReadLine()
                        If path.Trim = "" Then
                            path = Environment.CurrentDirectory
                        End If
                        Console.Write("Enter the name of your file, or leave blank to use default (winapp2.ini): ")
                        Dim name As String = Console.ReadLine()
                        If name.Trim = "" Then
                            name = "winapp2.ini"
                        End If
                        cfile = New iniFile(path, name)
                        debug(cfile)
                        exitCode = True
                    Case Else
                        Console.Write("Invalid input. Please try again: ")
                        cki = Console.ReadLine()
                End Select
            Catch ex As Exception
                Console.WriteLine(ex.ToString)
            End Try
        Loop
    End Sub

    Public Sub debug(cFile As iniFile)

        Dim numErrs As Integer = 0
        Dim entryTitles As New List(Of String)
        Dim trimmedEntryTitles As New List(Of String)
        Dim entryLineCounts As New List(Of String)
        Dim TBCommentLineNum As Integer = findCommentLine(cFile, "; End of Thunderbird entries.")
        Dim NonCCCommentLineNum As Integer = findCommentLine(cFile, "; These entries are the exact same ones located in the Removed entries files")

        Dim hasTBcomment As Boolean = True
        Dim hasNCCComment As Boolean = False

        If TBCommentLineNum = 0 Then
            hasTBcomment = False
        End If

        If Not NonCCCommentLineNum = 0 Then
            hasNCCComment = True
        End If

        For i As Integer = 0 To cFile.sections.Count - 1

            Dim curSection As iniSection = cFile.sections.Values(i)
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
        Console.WriteLine("***********************************************" & Environment.NewLine & "Completed analysis of winapp2.ini. " & numErrs & " possible errors were detected. " & Environment.NewLine & "Number of entries: " & entryTitles.Count & Environment.NewLine & "Press any key to return to the winapp2ool menu.")
        Console.ReadKey()
    End Sub


    Public Sub replaceAndSort(ByRef ListToBeSorted As List(Of String), ByRef givenSortedList As List(Of String), characterToReplace As String, ByVal replacementText As String, ByVal sortType As String)

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
                        i -= 1
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

    Public Sub cFormat(ByVal key As iniKey, ByRef keyNumber As Integer, ByRef numErrs As Integer, ByRef keyList As List(Of String))

        Dim command As String = key.ToString
        Dim keyString As String = key.keyType

        'check for duplicates
        If keyList.Contains(key.value) Then
            Console.WriteLine("Line: " & key.lineNumber & " - Error: Duplicate key found." & Environment.NewLine &
                                  "Command: " & command & Environment.NewLine &
                                  "Duplicates: " & keyString & keyList.IndexOf(key.value) + 1 & "=" & key.value.ToLower & Environment.NewLine)
        End If

        'make sure the current key is correctly numbered
        If Not command.Contains(keyString & keyNumber) Then
            err2(key.lineNumber, "'" & keyString & "' entry is incorrectly spelled or formatted.", command, keyString & keyNumber, numErrs)
        End If
        keyNumber += 1

        'make sure we don't have any dangly bits on the end of our key
        If command(command.Count - 1) = ";" Then
            err(key.lineNumber, "Trailing semicolon (;).", command, numErrs)
        End If

        'Do some formatting checks for environment variables
        If keyString = "FileKey" Or keyString = "ExcludeKey" Then
            enVarChecker(key, numErrs)
        End If

    End Sub

    Public Sub enVarChecker(key As iniKey, numErrs As Integer)

        If Command.Contains("%") Then
            Dim envars As New List(Of String)
            envars.AddRange(New String() {"UserProfile", "ProgramFiles", "RootDir", "WinDir", "AppData", "SystemDrive", "Documents", "ProgramData", "AllUsersProfile",
                        "Pictures", "Video", "CommonAppData", "LocalAppData", "CommonProgramFiles", "HomeDrive", "Music", "tmp", "Temp", "LocalLowAppData", "Public"})

            Dim varcheck As String() = Command.Split(Convert.ToChar("%"))
            If varcheck.Count <> 3 And varcheck.Count <> 5 Then
                err(key.lineNumber, "%EnvironmentVariables% must be surrounded on both sides by a single '%' character.", Command, numErrs)
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
                        err(key.lineNumber, "Misformatted or invalid environment variable.", Command, numErrs)
                    End If
                End If
            End If
        End If
    End Sub

    Public Function cValidity(key As iniKey, ByVal validList As List(Of String), numErrs As Integer) As Boolean

        'Check for trailing whitespace
        If key.ToString.EndsWith(" ") Then
            err(key.lineNumber, "Detected unwanted whitepace at end of line.", key.ToString, numErrs)
        End If

        'Check for ending whitespace
        If key.ToString.StartsWith(" ") Then
            err(key.lineNumber, "Detected unwanted whitepace at beginning of line.", key.ToString, numErrs)
        End If

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
            'check the format of the filekey
            cFormat(key, curFileKeyNum, numErrs, fileKeyList)

            'add the filekey contents to the duplicate checking list
            Dim command As String = key.value
            fileKeyList.Add(command)

            'Pipe symbol checks
            Dim iteratorCheckerList() As String = Split(command, "|")
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
        If Command.ToLower.Contains("true") Then
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
        Dim sdKey As New iniKey
        Dim validCmds As New List(Of String)
        validCmds.AddRange(New String() {"specialdetect", "filekey", "regkey", "detect", "langsecref", "warning", "[", "default", "section", "excludekey"})

        For i As Integer = 0 To section.keys.Count - 1
            Dim key As iniKey = section.keys.Values(i)
            'make sure our key contains a valid command
            If Not cValidity(key, validCmds, numErrs) Then
                Continue For
            End If

            Select Case key.keyType
                Case "LangSecRef"
                    secRefKey = key
                Case "FileKey"
                    fileKeys.Add(key)

                    'Make sure that FileKeys come before RegKeys in the section (style error but not syntax)
                    If entryHasRegKeys And Not entryKeysOutOfOrder Then
                        err(key.lineNumber, "FileKeys should precede RegKeys.", key.ToString, numErrs)
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

        pSpecialDetect(sdKey, numErrs)
    End Sub

    Private Sub pDetectFile(ByRef key As iniKey, ByRef numErrs As Integer)

        'check our environment variables
        enVarChecker(key, numErrs)

        'check for trailing backslashes
        Dim command As String = key.value

        'backslash check
        If command(command.Count - 1) = "\" Then
            err(key.lineNumber, "Trailing backslash on DetectFile.", key.ToString, numErrs)
        End If

        'check for nested wildcards
        If command.Contains("*") Then
            Dim splitDir As String() = command.Split(Convert.ToChar("\"))
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
                err(key.lineNumber, "Semicolon (;) found in " & detectType, key.ToString, numErrs)
            End If

            'check for duplicates
            If keyList.Contains(command) Then
                Console.WriteLine("Line: " & key.lineNumber & " - Error: Duplicate command found. " & Environment.NewLine & "Command: " & key.ToString & Environment.NewLine & "Duplicates: " & detectType & keyList.IndexOf(command) + 1 & "=" & command & Environment.NewLine)
                numErrs += 1
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

        Dim curRegKey As Integer = 1
        Dim regKeyList As New List(Of String)

        'Ensure that each RegKey points to a valid registry location
        For Each key As iniKey In keyList
            cFormat(key, curRegKey, numErrs, regKeyList)
            If Not key.ToString.Contains("=HKLM") And Not key.ToString.Contains("=HKC") And Not key.ToString.Contains("=HKU") Then
                err(key.lineNumber, "'RegKey' can only be used for registry key paths.", key.ToString, numErrs)
            End If
            regKeyList.Add(key.value)
        Next
    End Sub

    Public Sub pSpecialDetect(ByRef key As iniKey, numErrs As Integer)

        'make sure that any SpecialDetect keys hold a valid value
        If key.value <> "" Then
            Dim sdList As New List(Of String)
            sdList.AddRange(New String() {"DET_CHROME", "DET_MOZILLA", "DET_THUNDERBIRD", "DET_OPERA"})

            If Not sdList.Contains(key.value) Then
                err(key.lineNumber, "SpecialDetect holds an invalid value.", key.ToString, numErrs)
            End If
        End If
    End Sub

    Private Sub pExcludeKey(ByRef keyList As List(Of iniKey), ByRef numErrs As Integer)

        Dim curExcludeKeyNum As Integer = 1
        Dim ekKeyList As New List(Of String)
        For Each key As iniKey In keyList
            'check the format
            cFormat(key, curExcludeKeyNum, numErrs, ekKeyList)
            Dim command As String = key.value
            ekKeyList.Add(command)

            'Make sure any FILE exclude paths have a backslash before their pipe symbol
            If command.Contains("FILE") Then
                Dim iteratorCheckerList() As String = Split(command, "|")
                Dim endingslashchecker() As String = Split(command, "\|")
                If endingslashchecker.Count = 1 Then
                    err(key.lineNumber, "Missing backslash (\) before pipe (|) in ExcludeKey.", command, numErrs)
                End If
            End If
        Next
    End Sub

    Public Function findCommentLine(sFile As iniFile, com As String) As Integer

        'find the line number of a particular comment by its string
        For i As Integer = 0 To sFile.comments.Count - 1
            If sFile.comments(i).comment.Equals(com) Then
                Return sFile.comments(i).lineNumber
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
        Public comments As New Dictionary(Of Integer, iniComment)

        Public Sub New()
            name = ""
            sections = New Dictionary(Of String, iniSection)
            comments = New Dictionary(Of Integer, iniComment)
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
                        If Me.comments.Count > 0 Then
                            Me.comments.Add(comments.Count, newCom)
                        Else
                            Me.comments.Add(0, newCom)
                        End If
                    Else
                        If currentLine.Trim <> "" Then
                            sectionToBeBuilt.Add(currentLine)
                            lineTrackingList.Add(lineCount)
                        Else
                            If Not sectionToBeBuilt.Count < 2 Then
                                Dim sectionHolder As New iniSection(sectionToBeBuilt, lineTrackingList)
                                Try
                                    sections.Add(sectionHolder.name, sectionHolder)
                                Catch ex As Exception
                                    Console.WriteLine(ex.Message & Environment.NewLine & "Failure occurred during iniFile construction at line: " & Me.sections.Last.Value.endingLineNumber + 2)
                                    Console.WriteLine()
                                End Try
                                sectionToBeBuilt.Clear()
                                lineTrackingList.Clear()
                            End If
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

        Public Function getStreamOfComments(startNum As Integer, endNum As Integer) As String
            Dim out As String = ""
            While Not startNum = endNum
                out += Me.comments(startNum).comment & Environment.NewLine
                startNum += 1
            End While
            out += Me.comments(startNum).comment
            Return out
        End Function

        Public Function getDiff(section As iniSection, changeType As String) As String
            Dim out As String = ""
            out += section.name & " has been " & changeType & Environment.NewLine
            out += Environment.NewLine
            out += section.ToString & Environment.NewLine
            out += "*--------------------------------------------------------------------------------------------------*"
            out += Environment.NewLine
            Return out
        End Function

        Public Function compareTo(secondFile As iniFile) As List(Of String)

            Dim outList As New List(Of String)
            Dim comparedList As New List(Of String)

            For i As Integer = 0 To Me.sections.Count - 1
                Dim curSection As iniSection = Me.sections.Values(i)
                Dim curName As String = curSection.name
                Try
                    If secondFile.sections.Keys.Contains(curName) And Not comparedList.Contains(curName) Then
                        Dim sSection As iniSection = secondFile.sections(curName)
                        If Not curSection.compareTo(sSection) Then
                            outList.Add(getDiff(curSection, "modified."))
                        End If
                        comparedList.Add(curName)
                    ElseIf Not secondFile.sections.Keys.Contains(curName) Then
                        outList.Add(getDiff(curSection, "removed."))
                    End If
                Catch ex As Exception
                    Console.WriteLine(ex.ToString)
                End Try
            Next

            For i As Integer = 0 To secondFile.sections.Count - 1
                Dim curSection As iniSection = secondFile.sections.Values(i)
                Dim curName As String = curSection.name

                If Not Me.sections.Keys.Contains(curName) Then
                    outList.Add(getDiff(curSection, "added."))
                End If
            Next

            Return outList
        End Function
    End Class

    Class iniSection
        Public startingLineNumber As Integer
        Public endingLineNumber As Integer
        Public name As String
        Public keys As New Dictionary(Of Integer, iniKey)

        Public Function getFullName() As String
            Return "[" & Me.name & "]"
        End Function

        Public Sub New()
            startingLineNumber = 0
            endingLineNumber = 0
            name = ""
        End Sub

        Public Sub New(ByVal listOfLines As List(Of String), listOfLineCounts As List(Of Integer))

            Dim tmp1 As String() = listOfLines(0).Split(Convert.ToChar("["))
            Dim tmp2 As String() = tmp1(1).Split(Convert.ToChar("]"))
            name = tmp2(0)
            startingLineNumber = listOfLineCounts(0)
            endingLineNumber = listOfLineCounts(listOfLineCounts.Count - 1)

            If listOfLines.Count > 1 Then
                For i As Integer = 1 To listOfLines.Count - 1
                    Dim curKey As New iniKey(listOfLines(i), listOfLineCounts(i))
                    keys.Add(i - 1, curKey)
                Next
            End If
        End Sub

        'returns true if the sections are the same, else returns false
        Public Function compareTo(secondSection As iniSection) As Boolean

            If keys.Count <> secondSection.keys.Count Then
                Return False
            Else
                Dim isdiff As Boolean = False
                For i As Integer = 0 To keys.Count - 1
                    If Not keys(i).compareTo(secondSection.keys(i)) Then
                        Return False
                    End If
                Next
            End If
            Return True
        End Function

        Public Overrides Function ToString() As String

            Dim out As String = Me.getFullName

            For i As Integer = 1 To Me.keys.Count
                out += Environment.NewLine & Me.keys(i - 1).ToString
            Next
            out += Environment.NewLine
            Return out
        End Function
    End Class

    Class iniKey

        Public name As String
        Public value As String
        Public lineNumber As Integer
        Public keyType As String

        Public Sub New()

            name = ""
            value = ""
            lineNumber = 0
            keyType = ""
        End Sub

        Public Function stripNums(keyName As String) As String

            For i As Integer = 0 To 9
                keyName = keyName.Replace(i.ToString, "")
            Next
            Return keyName
        End Function

        Public Sub New(ByVal line As String, ByVal count As Integer)

            Try
                Dim splitLine As String() = line.Split(Convert.ToChar("="))
                name = splitLine(0)
                value = splitLine(1)
                keyType = stripNums(name)
                lineNumber = count
            Catch ex As Exception
                Console.WriteLine(ex)
            End Try
        End Sub

        Public Overrides Function ToString() As String
            Return Me.name & "=" & Me.value
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

    Public Sub main()

        Console.Clear()
        Console.WriteLine("*--------------------------------------------------------------------------------------------------*")
        Console.WriteLine("*                                            Diff                                                  *")
        Console.WriteLine("*                                                                                                  *")
        Console.WriteLine("*                This tool will output the diff between two winapp2 files                          *")
        Console.WriteLine("*                                                                                                  *")
        Console.WriteLine("*                                            Menu:                                                 *")
        Console.WriteLine("* 0. Exit                - Return to the winapp2ool menu                                           *")
        Console.WriteLine("* 1. Run (default)       - Run Diff on files in the current folder                                 *")
        Console.WriteLine("* 2. Run (custom)        - Run Diff on files in a different folder                                 *")
        Console.WriteLine("*--------------------------------------------------------------------------------------------------*")
        Console.Write("Enter a number: ")

        Dim input As String = Console.ReadLine()
        Dim exitCode As Boolean = False
        Dim firstFile As iniFile
        Dim secondFile As iniFile

        Do Until exitCode
            Try
                Select Case input
                    Case "0"
                        Console.WriteLine("Exiting diff...")
                        exitCode = True
                    Case "1"
                        Console.Write("Enter the name of the first older file: ")
                        Dim oName As String = Console.ReadLine()
                        Console.Write("Enter the name of the second newer file: ")
                        Dim nNname As String = Console.ReadLine()
                        firstFile = New iniFile(oName)
                        secondFile = New iniFile(nNname)
                        differ(firstFile, secondFile)
                        exitCode = True
                    Case "2"
                        Console.Write("Enter the directory of the older file: ")
                        Dim oPath As String = Console.ReadLine()
                        Console.Write("Enter the name of the older file: ")
                        Dim oName As String = Console.ReadLine()
                        Console.Write("Enter the directory of the newer file: ")
                        Dim nPath As String = Console.ReadLine()
                        Console.Write("Enter the name of the newer file: ")
                        Dim nName As String = Console.ReadLine()
                        firstFile = New iniFile(oPath, oName)
                        secondFile = New iniFile(nPath, nName)
                        exitCode = True
                    Case Else
                        Console.Write("Invalid input. Please try again: ")
                        input = Console.ReadLine()
                End Select
            Catch ex As Exception
                Console.WriteLine(ex.ToString)
            End Try
        Loop
    End Sub

    Public Sub differ(firstFile As iniFile, secondFile As iniFile)

        Try
            Dim fver As String = firstFile.comments(0).comment.ToString
            fver = fver.Split(Convert.ToChar(";"))(1)
            Dim sver As String = secondFile.comments(0).comment.ToString
            sver = sver.Split(Convert.ToChar(";"))(1)
            Console.WriteLine("Changes made between " & fver & " and" & sver)
            Console.WriteLine()
            Dim outList As List(Of String) = firstFile.compareTo(secondFile)
            Dim remCt As Integer = 0
            Dim modCt As Integer = 0
            Dim addCt As Integer = 0
            For Each change As String In outList

                If change.Contains("has been added.") Then
                    addCt += 1
                ElseIf change.Contains("has been removed") Then
                    remCt += 1
                Else
                    modCt += 1
                End If
                Console.WriteLine(change)
            Next

            Console.WriteLine("Finished diffing. Change counts: ")
            Console.WriteLine()
            Console.WriteLine("Added entries: " & addCt)
            Console.WriteLine("Modified entires: " & modCt)
            Console.WriteLine("Removed entries: " & remCt)
            Console.WriteLine()
            Console.WriteLine("*--------------------------------------------------------------------------------------------------*")

        Catch ex As Exception
            Console.WriteLine("Error: " & ex.ToString)
        End Try
        Console.WriteLine("End of diff. Press any key to return to the winapp2ool menu.")
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
        Console.WriteLine("*                                             Menu:                                                *")
        Console.WriteLine("*0. Exit                - Return to the winapp2ool menu                                            *")
        Console.WriteLine("*1. Trim (default)      - Trim winapp2.ini and overwrite the existing file                         *")
        Console.WriteLine("*2. Trim (newfile)      - Trim winapp2.ini and save the output to a new file                       *")
        Console.WriteLine("*--------------------------------------------------------------------------------------------------*")
        Console.Write("Enter a number: ")

        Dim exitCode As Boolean = False
        Dim input As String = Console.ReadLine()
        Do Until exitCode
            Select Case input
                Case "0"
                    Console.WriteLine("Returning to winapp2ool menu...")
                    exitCode = True
                Case "1"
                    trim("\winapp2.ini")
                    exitCode = True
                Case "2"
                    Console.Write("Enter the name of the new file, or press enter to use the default (winapp2-trimmed.ini): ")
                    Dim in2 As String = Console.ReadLine()
                    If in2.Trim <> Nothing Then
                        trim("\" & in2)
                    Else
                        trim("\winapp2-trimmed.ini")
                    End If
                    exitCode = True
                Case Else
                    Console.Write("Invalid input. Please try again: ")
                    input = Console.ReadLine
            End Select
        Loop
    End Sub

    Public Sub printComments(ifile As iniFile, start As Integer, endpt As Integer, file As StreamWriter)
        For i As Integer = start To endpt
            file.WriteLine(ifile.comments(i).comment)
        Next
    End Sub


    Public Sub trim(name As String)

        Console.WriteLine("Trimming...")

        'load winapp2.ini into memory
        Dim winappfile As New iniFile("winapp2.ini")

        'create a list of sections that belong in the trimmed file
        Dim trimmedfile As New List(Of iniSection)

        'get the windows version so we can test properly our DetectOS
        Dim winveri As Double = getWinVer()

        Dim cEntries As New List(Of iniSection)
        Dim fxEntries As New List(Of iniSection)
        Dim tbEntries As New List(Of iniSection)

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
            For j As Integer = 0 To cursection.keys.Count - 1
                Dim key As iniKey = cursection.keys(j)

                If exists Then
                    Exit For
                End If

                Dim type As String = key.keyType

                Select Case type
                    Case "Detect"
                        hasDets = True
                        If checkExist(key.value) Then
                            trimmedfile.Add(cursection)
                            Exit For
                        End If
                    Case "DetectFile"
                        hasDets = True
                        If checkExist(key.value) Then
                            trimmedfile.Add(cursection)
                            Exit For
                        End If
                    Case "DetectOS"
                        hasDetOS = True
                        hasMetDetOS = detOSCheck(key.value, winveri)

                        'move on if we don't meet the DetectOS criteria
                        If Not hasMetDetOS Then
                            Exit For
                        End If
                    Case "SpecialDetect"
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
                End Select
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
        Dim entrycount As Integer = trimmedfile.Count + cEntries.Count + fxEntries.Count + tbEntries.Count
        Console.WriteLine("Number of entries after trimming: " & entrycount)
        Try
            Dim file As New StreamWriter(Environment.CurrentDirectory & name, False)

            Dim comNum As Integer
            'contingency for non-cc ini
            If winappfile.comments.Count > 16 Then
                comNum = 9
            Else
                comNum = 8
            End If

            file.WriteLine(winappfile.getStreamOfComments(0, comNum))
            comNum += 1
            file.WriteLine()
            writeCustomSectionToFile(cEntries, file)

            file.WriteLine(winappfile.getStreamOfComments(comNum, comNum + 2))
            comNum += 3
            file.WriteLine()
            writeCustomSectionToFile(fxEntries, file)

            file.WriteLine(winappfile.getStreamOfComments(comNum, comNum + 2))
            file.WriteLine()
            writeCustomSectionToFile(tbEntries, file)
            comNum += 3
            file.WriteLine(winappfile.getStreamOfComments(comNum, comNum))
            file.WriteLine()

            For Each section As iniSection In trimmedfile
                file.Write(section.ToString())
                If Not trimmedfile.IndexOf(section) = trimmedfile.Count - 1 Then
                    file.WriteLine()
                End If
            Next
            file.Close()
        Catch ex As Exception
            Console.WriteLine(ex.ToString)
        End Try
        Console.WriteLine("*--------------------------------------------------------------------------------------------------*")
        Console.WriteLine("Finished trimming winapp2.ini, press any key to return to the winapp2ool menu.")
        Console.ReadKey()
    End Sub

    Public Function checkExist(key As String) As Boolean

        Dim dir As String = key
        Dim isProgramFiles As Boolean = False

        'make sure we get the proper path for environment variables
        If dir.Contains("%") Then
            Dim splitDir As String() = dir.Split(Convert.ToChar("%"))
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
            Dim splitDir As String() = dir.Split(Convert.ToChar("\"))
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
    Public Sub writeCustomSectionToFile(entryList As List(Of iniSection), file As StreamWriter)

        For i As Integer = 0 To entryList.Count - 1
            file.WriteLine(entryList(i).ToString)
        Next
    End Sub

    Private Function detOSCheck(value As String, winveri As Double) As Boolean

        Dim splitKey As String() = value.Split(Convert.ToChar("|"))

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

Module ccinidebug
    Sub Main()

        Console.Clear()
        Console.WriteLine("*--------------------------------------------------------------------------------------------------*")
        Console.WriteLine("*                                        ccinidebug                                                *")
        Console.WriteLine("*                                                                                                  *")
        Console.WriteLine("*               This tool will sort alphabetically the contents of ccleaner.ini                    *")
        Console.WriteLine("*                   and can also prune 'stale' winapp2.ini entries from it                         *")
        Console.WriteLine("*                                           Menu:                                                  *")
        Console.WriteLine("* 0. Exit                 - Return to the winapp2ool menu                                          *")
        Console.WriteLine("* 1. Run (default)        - Prune stale winapp2.ini entries from ccleaner.ini and sort it          *")
        Console.WriteLine("* 2. Run (sort only)      - Only sort ccleaner.ini                                                 *")
        Console.WriteLine("*--------------------------------------------------------------------------------------------------*")
        Console.Write("Enter a number: ")

        Dim exitCode As Boolean = False
        Dim input As String = Console.ReadLine
        Dim ccini As iniSection
        Dim winappini As iniFile

        Do Until exitCode
            Try
                Select Case input
                    Case "0"
                        Console.WriteLine("Returning to winapp2ool menu...")
                        exitCode = True
                    Case "1"
                        ccini = buildOptions()
                        winappini = New iniFile("winapp2.ini")
                        ccini = prune(ccini, winappini)
                        writeccini(ccini)
                        exitCode = True
                    Case "2"
                        ccini = buildOptions()
                        writeccini(ccini)
                    Case Else
                        Console.Write("Invalid input. Please try again: ")
                        input = Console.ReadLine
                End Select
            Catch ex As Exception
                Console.WriteLine(ex.ToString)
            End Try
        Loop
    End Sub

    Public Sub writeccini(ccini As iniSection)

        Console.WriteLine("Modifying ccleaner.ini...")
        Try
            Dim file As New StreamWriter(Environment.CurrentDirectory & "\ccleaner.ini", False)

            file.WriteLine("[Options]")
            For i As Integer = 0 To ccini.keys.Count - 1
                file.WriteLine(ccini.keys.Values(i))
            Next

            file.Close()
            Console.WriteLine("*--------------------------------------------------------------------------------------------------*")
            Console.WriteLine("Finished modifying ccleaner.ini, press any key to return to the winapp2ool menu.")
            Console.ReadKey()
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try
    End Sub

    'this function builds the single Options section found in ccleaner.ini and avoids being broken by newlines
    Public Function buildOptions() As iniSection

        Dim returnSection As New iniSection
        Dim line As Integer = 0
        Dim lineList As New List(Of String)
        Dim countList As New List(Of Integer)
        Dim r As IO.StreamReader
        Try
            r = New IO.StreamReader(Environment.CurrentDirectory & "\ccleaner.ini")
            Do While (r.Peek() > -1)
                line += 1
                Dim currentLine As String = r.ReadLine.ToString
                If currentLine.Trim <> "" Then
                    lineList.Add(currentLine)
                    countList.Add(line)
                End If
            Loop
            r.Close()
            lineList.Sort()
            lineList.Remove("[Options]")
            lineList.Insert(0, "[Options]")
            returnSection = New iniSection(lineList, countList)
            Return returnSection
        Catch ex As Exception
            Console.WriteLine(ex.ToString)
        End Try
        Return returnSection
    End Function

    Public Function prune(ccini As iniSection, winappini As iniFile) As iniSection

        'collect the keys we must remove
        Dim tbTrimmed As New List(Of Integer)

        For i As Integer = 0 To ccini.keys.Count - 1
            Dim optionStr As String = ccini.keys.Values(i).ToString

            'only operate on (app) keys
            If optionStr.StartsWith("(App)") And optionStr.Contains("*") Then
                optionStr = optionStr.Replace("(App)", "")
                optionStr = optionStr.Replace("=True", "")
                optionStr = optionStr.Replace("=False", "")
                If Not winappini.sections.ContainsKey(optionStr) Then
                    Console.WriteLine(ccini.keys.Values(i).ToString & " will be pruned")
                    tbTrimmed.Add(ccini.keys.Keys(i))
                End If
            End If
        Next

        'reverse the keys we must remove to avoid any problems with modifying the dictionary as we do so
        tbTrimmed.Reverse()
        For i As Integer = 0 To tbTrimmed.Count - 1
            ccini.keys.Remove(tbTrimmed(i))
        Next
        Console.WriteLine("Removed " & tbTrimmed.Count & " stale entries.")
        Return ccini
    End Function
End Module

Module downloader

    Public Sub download(filename As String, fileLink As String, downloadDir As String)
        If Not Directory.Exists(downloadDir) Then
            Directory.CreateDirectory(downloadDir)
        End If
        Console.WriteLine("Downloading " & filename & "...")
        Try
            Dim dl As New WebClient
            dl.DownloadFileAsync(New Uri(fileLink), (downloadDir & "\" & filename))
        Catch ex As Exception
            Console.WriteLine(ex.ToString)
        End Try
        Console.WriteLine("Downloaded " & filename & " to " & downloadDir)
    End Sub

    Public Sub main()

        Console.Clear()
        Console.WriteLine("*--------------------------------------------------------------------------------------------------*")
        Console.WriteLine("*                                             Download                                             *")
        Console.WriteLine("*                                                                                                  *")
        Console.WriteLine("*                                     Menu: Enter a number to select                               *")
        Console.WriteLine("*                                                                                                  *")
        Console.WriteLine("* 0. Exit             - Return to the winapp2ool menu                                              *")
        Console.WriteLine("* 1. winapp2.ini      - Downloads the latest winapp2.ini                                           *")
        Console.WriteLine("* 2. Non-CCleaner     - Downloads the latest winapp2.ini for non-ccleaner applications             *")
        Console.WriteLine("* 3. winapp2ool       - Downloads the latest winapp2ool.exe                                        *")
        Console.WriteLine("* 4. directory        - Change the download directory                                              *")
        Console.WriteLine("*--------------------------------------------------------------------------------------------------*")
        Console.Write("Enter a number: ")

        Dim exitCode As Boolean = False
        Dim input As String = Console.ReadLine()
        Dim downloadDir As String = Environment.CurrentDirectory & "\winapp2ool downloads"

        Dim wa2Link As String = "https://raw.githubusercontent.com/MoscaDotTo/Winapp2/master/Winapp2.ini"
        Dim nonccLink As String = "https://raw.githubusercontent.com/MoscaDotTo/Winapp2/master/Non-CCleaner/Winapp2.ini"
        Dim toolLink As String = "https://github.com/MoscaDotTo/Winapp2/raw/master/Tools/beta/winapp2ool.exe"

        Do Until exitCode
            Select Case input
                Case "0"
                    Console.WriteLine("Returning to winapp2ool menu...")
                    exitCode = True
                Case "1"
                    download("winapp2.ini", wa2Link, downloadDir)
                    Console.WriteLine("Enter a number (0 to exit): ")
                    input = Console.ReadLine()
                Case "2"
                    download("winapp2.ini", nonccLink, downloadDir)
                    Console.WriteLine("Enter a number (0 to exit): ")
                    input = Console.ReadLine()
                Case "3"
                    download("winapp2ool.exe", toolLink, downloadDir)
                    Console.WriteLine("Enter a number (0 to exit): ")
                    input = Console.ReadLine()
                Case "4"
                    Console.WriteLine("Current download directory: " & downloadDir)
                    Console.Write("Enter new directory, or just press enter to keep the current one: ")
                    Dim newDir As String = Console.ReadLine()
                    If newDir.Trim <> "" Then
                        downloadDir = newDir
                    End If
                    Console.WriteLine("Current download directory: " & downloadDir)
                    Console.Write("Enter a number: ")
                    input = Console.ReadLine()
                Case Else
                    Console.Write("Invalid input. Please try again: ")
                    input = Console.ReadLine()
            End Select
        Loop
    End Sub

End Module