Imports System.IO
Module WinappDebug
    Sub Main()

        'Check if the winapp2.ini file is in the current directory. End program if it isn't.
        If File.Exists(Environment.CurrentDirectory & "\winapp2.ini") = False Then
            Console.WriteLine("winapp2.ini file could not be located in the current working directory (" & Environment.CurrentDirectory & ")")
            Console.ReadKey()
            End
        End If

        'Create some variables for later use
        Dim linecount As Integer = 0
        Dim command As String
        Dim number_of_errors As Integer = 0
        Dim curFileKeyNumber As Integer = 1
        Dim curRegKeyNumber As Integer = 1
        Dim curExcludeKeyNumber As Integer = 1
        Dim curDetectFileNumber As Integer = 1
        Dim curDetectNumber As Integer = 1
        Dim firstDetectNumber As Boolean = False
        Dim firstDetectFileNumber As Boolean = False
        Dim havePassedTB As Boolean = False
        Dim fileKeyList As New List(Of String)
        Dim regKeyList As New List(Of String)
        Dim excludeKeyList As New List(Of String)
        Dim detectList As New List(Of String)
        Dim detectFileList As New List(Of String)
        Dim entryHasDefault As Boolean = False

        Dim entryLineCounts As New List(Of String)
        Dim excludeKeyLineCounts As New List(Of String)
        Dim fileKeyLineCounts As New List(Of String)
        Dim regKeyLineCounts As New List(Of String)
        Dim detectLineCounts As New List(Of String)
        Dim detectFileLineCounts As New List(Of String)
        'Create a list of supported environmental variables

        'Create a list of entry titles so we can look for duplicates
        Dim entry_titles As New List(Of String)
        Dim trimmed_entry_titles As New List(Of String)
        'Read the winapp2.ini file line-by-line
        Dim r As IO.StreamReader
        Try

            r = New IO.StreamReader(Environment.CurrentDirectory & "\winapp2.ini")
            Do While (r.Peek() > -1)

                'Update the line that is being tested
                command = (r.ReadLine.ToString)

                'Increment the lineocunt value
                linecount = linecount + 1
                Dim curEntry As String
                If command = "" Then

                    If entry_titles.Count = 0 Then
                        curEntry = ""
                    Else
                        curEntry = entry_titles.Last
                    End If

                    processEmptyLine(curFileKeyNumber, curRegKeyNumber, curExcludeKeyNumber, curDetectFileNumber, curDetectNumber, firstDetectFileNumber, firstDetectFileNumber,
                                     fileKeyList, regKeyList, excludeKeyList, regKeyLineCounts, fileKeyLineCounts, excludeKeyLineCounts, number_of_errors,
                                     detectFileList, detectFileLineCounts, detectList, detectLineCounts, entryHasDefault, curEntry)
                End If

                If command = "" = False And command.StartsWith(";") = False Then

                    'Check for trailing whitespace
                    If command.EndsWith(" ") Then
                        err(linecount, "Detected unwanted whitepace at end of line", command, number_of_errors)
                    End If

                    'Check for ending whitespace
                    If command.StartsWith(" ") Then
                        err(linecount, "Detected unwanted whitepace at beginning of line", command, number_of_errors)
                    End If

                Else
                    If command.Equals("; End of Thunderbird entries.") Then
                        havePassedTB = True
                    End If
                End If

                'Process entries
                If command.StartsWith("[") Then

                    processEntryName(command, entry_titles, linecount, havePassedTB, trimmed_entry_titles, number_of_errors, entryLineCounts)

                End If

                'Check for spelling errors in "LangSecRef"
                If command.StartsWith("L") Then
                    If Not command.Contains("LangSecRef=") Then
                        err(linecount, "LangSecRef is incorrected spelled or formatted.", command, number_of_errors)
                    End If
                End If

                'Process FileKey
                If command.StartsWith("F") Then

                    checkFormat(command, "FileKey", curFileKeyNumber, linecount, number_of_errors, fileKeyList, fileKeyLineCounts)
                    processFileKey(command, number_of_errors, linecount)

                End If

                'Process ExcludeKey
                If command.StartsWith("E") Then

                    checkFormat(command, "ExcludeKey", curExcludeKeyNumber, linecount, number_of_errors, excludeKeyList, excludeKeyLineCounts)
                    processExcludeKey(command, number_of_errors, linecount)

                End If

                'Process RegKey
                If command.StartsWith("R") Then

                    checkFormat(command, "RegKey", curRegKeyNumber, linecount, number_of_errors, regKeyList, regKeyLineCounts)
                    processRegKey(command, linecount, number_of_errors)
                End If

                If command.StartsWith("D") Then
                    Dim tmp As String() = Split(command, "=")

                    'We'll make the bold assumption here that all our defaults are spelled correctly, even if not formatted correctly. 
                    If Not command.ToLower.Contains("default") And Not command.ToLower.Contains("detecto") Then

                        If tmp(0).Length <= 8 Then

                            processDetect(command, "Detect", curDetectNumber, firstDetectNumber, number_of_errors, linecount, detectList, detectLineCounts)
                            processRegDetect(command, number_of_errors, linecount)

                        Else

                            processDetect(command, "DetectFile", curDetectFileNumber, firstDetectFileNumber, number_of_errors, linecount, detectFileList, detectFileLineCounts)
                            processDetectFile(command, number_of_errors, linecount)
                        End If
                    Else

                        If Not command.StartsWith("Default=") And Not command.ToLower.Contains("detecto") Then
                            err2(linecount, "'Default' incorrectly spelled", tmp.First, "Default", number_of_errors)
                        End If

                        If command.ToLower.Contains("true") Then
                            err(linecount, "All entries should be disabled by default (Default=False)", command, number_of_errors)
                        End If

                        If command.ToLower.StartsWith("default") Then
                            entryHasDefault = True
                        End If
                    End If

                End If

            Loop

            'Close unmanaged code calls
            r.Close()


        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try

        Dim sortedEList As New List(Of String)
        sortedEList.AddRange(trimmed_entry_titles.ToArray)


        'traverse through our trimmed entry titles list
        replaceAndSort(trimmed_entry_titles, sortedEList, "-", "  ", "entries")
        findOutOfPlace(trimmed_entry_titles, sortedEList, "Entry", number_of_errors, entryLineCounts)

        'Stop the program from closing on completion
        Console.WriteLine("***********************************************" & Environment.NewLine & "Completed analysis of winapp2.ini. " & number_of_errors & " possible errors were detected. " & Environment.NewLine & "Number of entries: " & entry_titles.Count & Environment.NewLine & "Press any key to close.")
        Console.ReadKey()

    End Sub

    Private Sub processDetectFile(ByRef command As String, ByRef number_of_errors As Integer, ByRef lineCount As Integer)

        If command(command.Count - 1) = "\" Then
            err(lineCount, "Trailing backslash on DetectFile", command, number_of_errors)
        End If

        Dim cDir As String = command.Split("=")(1)
        If cDir.Contains("*") Then
            Dim splitDir As String() = cDir.Split("\")
            For i As Integer = 0 To splitDir.Count - 1
                If splitDir.Last = Nothing Then Continue For
                If splitDir(i).Contains("*") And i <> splitDir.Count - 1 Then
                    err(lineCount, "Nested wildcard found in DetectFile", command, number_of_errors)
                End If

            Next

        End If


        If command.Contains("=HKLM") Or command.Contains("=HKC") Or command.Contains("=HKU") Then
            err(lineCount, "'DetectFile' can only be used for file system paths", command, number_of_errors)
        End If

    End Sub

    Private Sub processRegDetect(ByRef command As String, ByRef number_of_errors As Integer, ByRef lineCount As Integer)

        If (command.Contains("=%") Or command.Contains("=C:\")) Or (Not command.Contains("=HKLM") And Not command.Contains("=HKC") And Not command.Contains("=HKU")) Then
            err(lineCount, "'Detect' can only be used for registry keys paths.", command, number_of_errors)
        End If


    End Sub

    Private Sub processDetect(ByRef command As String, ByVal detectType As String, ByRef number As Integer, ByRef firstNumberStatus As Boolean, ByRef number_of_errors As Integer, ByRef lineCount As Integer, ByRef keyList As List(Of String), ByRef keyListLineCounts As List(Of String))

        If Not command.Contains(detectType) Then
            err(lineCount, "Misformatted '" & detectType, command, number_of_errors)

        End If

        Dim cmdList As String() = command.Split("=")
        If keyList.Contains(cmdList(1).ToLower) Then
            Console.WriteLine("Line: " & lineCount & " - Error: Duplicate command found. " & Environment.NewLine & "Command: " & command & Environment.NewLine & "Duplicates: " & detectType & keyList.IndexOf(cmdList(1).ToLower) + 1 & "=" & cmdList(1) & Environment.NewLine)
        Else
            keyList.Add(cmdList(1).ToLower)
            keyListLineCounts.Add(lineCount)

        End If

        'Check whether our first Detect or DetectFile is trailed by a number
        If number = 1 Then
            If command.StartsWith(detectType & "=") Or (Not command.StartsWith(detectType & "=") And Not command.StartsWith(detectType & "1")) Then
                firstNumberStatus = False
            Else
                firstNumberStatus = True
            End If

        End If

        'Make sure our first number is a 1 if there's a number
        If number = 1 And (Not command.Contains(detectType & "1=") And Not command.Contains(detectType & "=")) Then
            err2(lineCount, "'" & detectType & "' numbering.", command, "'" & detectType & "' or '" & detectType & "1", number_of_errors)

        End If

        'If we're on our second detect, make sure the first one had a 1 
        If number = 2 And firstNumberStatus = False Then
            err(lineCount - 1, "'" & detectType & number & "' detected without preceding '" & detectType & number - 1 & "'", command, number_of_errors)

            number_of_errors = number_of_errors + 1
        End If

        'otherwise, make sure our numbers match up
        If number > 1 Then
            If command.Contains(detectType & number) = False Then
                err2(lineCount, "'" & detectType & "' entry is incorrectly numbered", command, "'" & detectType & number & "'", number_of_errors)
            End If
        End If
        number = number + 1

    End Sub

    Private Sub processRegKey(ByVal command As String, ByVal lineCount As Integer, ByRef number_of_errors As Integer)

        If Not command.Contains("=HKLM") And Not command.Contains("=HKC") And Not command.Contains("=HKU") Then
            err(lineCount, "'RegKey' can only be used for registry key paths", command, number_of_errors)
        End If

    End Sub

    Private Sub processExcludeKey(ByRef command As String, ByRef number_of_errors As Integer, ByVal lineCount As Integer)

        'Make sure any FILE exclude paths have a backslash before their pipe symbol
        Dim iteratorCheckerList() As String = Split(command, "|")

        If iteratorCheckerList(0).Contains("FILE") Then
            Dim endingslashchecker() As String = Split(command, "\|")

            If endingslashchecker.Count = 1 Then
                err(lineCount, "Missing backslash (\) before pipe (|) in ExcludeKey", command, number_of_errors)
            End If

        End If

    End Sub

    Private Sub processFileKey(ByRef command As String, ByRef number_of_errors As Integer, ByVal lineCount As Integer)

        Dim iteratorCheckerList() As String = Split(command, "|")

        If Not command.Contains("|") Then

            err(lineCount, "Missing pipe (|) in 'FileKey'", command, number_of_errors)

        End If
        If command.Contains(";|") Then
            err(lineCount, "Semicolon (;) found before pipe (|)", command, number_of_errors)
        End If

        'check for incorrect spellings of RECURSE or REMOVESELF
        If iteratorCheckerList.Length > 2 Then
            If Not iteratorCheckerList(2).Contains("RECURSE") And Not iteratorCheckerList(2).Contains("REMOVESELF") Then

                err(lineCount, "'RECURSE' or 'REMOVESELF' entry is incorrectly spelled.", command, number_of_errors)

            End If
        End If

        'check for missing pipe symbol on recurse and removeself
        If command.Contains("RECURSE") And Not command.Contains("|RECURSE") Then
            err(lineCount, "Missing pipe (|) before 'RECURSE'", command, number_of_errors)
        End If
        If command.Contains("REMOVESELF") And Not command.Contains("|REMOVESELF") Then
            err(lineCount, "Missing pipe (|) before 'REMOVESELF'", command, number_of_errors)
        End If

        If command.Contains("\VirtualStore\P") And (Not command.ToLower.Contains("programdata") And Not command.ToLower.Contains("program files*") And Not command.ToLower.Contains("program*")) Then
            err2(lineCount, "Incorrect VirtualStore location.", command, "%LocalAppData%\VirtualStore\Program Files*\", number_of_errors)
        End If

        If command.Contains("%\|") Then
            err(lineCount, "Backslash (\) found before pipe (|)", command, number_of_errors)
        End If

        If command.Contains("%") And Not command.Contains("%|") And Not command.Contains("%\") Then
            err(lineCount, "Missing backslash (\) after %EnvironmentVariable%", command, number_of_errors)
        End If

    End Sub

    Private Sub processEntryName(ByRef command As String, ByRef entry_titles As List(Of String), ByRef linecount As Integer, ByRef havePassedTB As Boolean, ByRef trimmed_entry_titles As List(Of String), ByRef number_of_errors As Integer, ByRef entryLineCounts As List(Of String))
        'Check if it's already in the list
        If entry_titles.Contains(command) Then
            Console.WriteLine("Line: " & linecount & " Duplicate entry." & Environment.NewLine & "Command: " & command & Environment.NewLine)
            number_of_errors = number_of_errors + 1
        Else

            'Not already in the list. Add it.

            entry_titles.Add(command)

            Dim currentEntry As String = entry_titles.Last
            currentEntry = currentEntry.Remove(0, 1)
            currentEntry = currentEntry.Remove(currentEntry.Length - 2)

            'the entries in and above the thunderbird section don't need to be in order because they are grouped categorically 
            If havePassedTB Then
                trimmed_entry_titles.Add(command)
                entryLineCounts.Add(linecount)
            End If

            If Not command.Contains(" *]") Then
                err(linecount, "Improper entry name. All entries should end with a ' *'", command, number_of_errors)
            End If

        End If
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
            If sortType = "keys" Or sortType = "entries" Then
                findAndReplNumbers(myChars, entry, originalNameList, renamedList, ListToBeSorted, givenSortedList, i, sortedIndex)
            End If

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

    Public Sub findOutOfPlace(ByRef someList As List(Of String), ByRef sortedList As List(Of String), ByVal findType As String, ByRef number_of_errors As Integer, ByRef LineCountList As List(Of String))
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
                        If LineCountList.Count > 0 Then
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
                        number_of_errors = number_of_errors + 1

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

            Console.WriteLine("Line: " & originalLines(recInd) & " - Error: '" & findType & "' Alphabetization. " & entry & " appears to be out of place." & Environment.NewLine & "Current  Position: Line " & curPos & Environment.NewLine & "Expected Position: Line " & sortPos & Environment.NewLine)
            number_of_errors = number_of_errors + 1


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

    Public Sub checkFormat(ByVal command As String, ByVal keyString As String, ByRef keyNumber As Integer, ByVal lineCount As Integer, ByRef number_of_errors As Integer, ByRef keyList As List(Of String), ByRef keyLineCounts As List(Of String))

        'Split the command and save what comes after the '='
        Dim cmdList As String() = command.Split("=")

        If keyList.Contains(cmdList(1).ToLower) Then
            Console.WriteLine("Line: " & lineCount & " - Error: Duplicate command found. " & Environment.NewLine & "Command: " & command & Environment.NewLine & "Duplicates: " & keyString & keyList.IndexOf(cmdList(1).ToLower) + 1 & "=" & cmdList(1) & Environment.NewLine)
        Else
            keyList.Add(cmdList(1).ToLower)
            keyLineCounts.Add(lineCount)

        End If

        'make sure the current key is correctly numbered
        If Not command.Contains(keyString & keyNumber) Then
            err2(lineCount, "'" & keyString & "' entry is incorrectly spelled or formatted.", command, keyString & keyNumber, number_of_errors)
        End If
        keyNumber = keyNumber + 1

        If command(command.Count - 1) = ";" Then
            err(lineCount, "Trailing semicolon (;)", command, number_of_errors)
        End If


        'Do some formatting checks for environment variables
        If keyString = "DetectFile" Or keyString = "FileKey" Or keyString = "ExcludeKey" Then

            If command.Contains("%") Then
                Dim envir_vars As New List(Of String)
                envir_vars.AddRange(New String() {"UserProfile", "ProgramFiles", "RootDir", "WinDir", "AppData", "SystemDrive", "Documents", "ProgramData", "AllUsersProfile",
                            "Pictures", "Video", "CommonAppData", "LocalAppData", "CommonProgramFiles", "HomeDrive", "Music", "tmp", "Temp", "LocalLowAppData", "Public"})


                Dim varcheck As String() = command.Split("%")
                If varcheck.Count <> 3 And varcheck.Count <> 5 Then
                    err(lineCount, "%EnvironmentVariables% must be surrounded on both sides by a single % character.", command, number_of_errors)
                End If

                If varcheck.Count = 3 Then
                    If Not envir_vars.Contains(varcheck(1)) Then

                        Dim casingerror As Boolean = False
                        For Each var As String In envir_vars

                            If varcheck(1).ToLower = var.ToLower Then
                                casingerror = True
                                err2(lineCount, "Invalid CamelCasing on environment variable.", command, var, number_of_errors)

                            End If

                        Next

                        If casingerror = False Then
                            err(lineCount, "Misformatted or invalid environment variable.", command, number_of_errors)

                        End If
                    End If

                End If
            End If


        End If

    End Sub

    Public Sub err2(ByVal linecount As Int16, ByVal err As String, ByVal command As String, ByVal expected As String, ByRef number_of_errors As Integer)


        Console.WriteLine("Line: " & linecount & " - Error: " & err & Environment.NewLine & "Expected: " & expected & Environment.NewLine & "Command: " & command & Environment.NewLine)
        number_of_errors = number_of_errors + 1

    End Sub

    Public Sub err(ByVal linecount As Integer, ByVal err As String, ByVal command As String, ByRef number_of_errors As Integer)

        Console.WriteLine("Line: " & linecount & " - Error: " & err & Environment.NewLine & "Command: " & command & Environment.NewLine)
        number_of_errors = number_of_errors + 1

    End Sub

    Public Sub processEmptyLine(ByRef curFileKeyNumber As Integer, ByRef curRegKeyNumber As Integer, ByRef curExcludeKeyNumber As Integer, ByRef curDetectFileNumber As Integer, ByRef curDetectNumber As Integer,
                                ByRef firstDetectFileNumber As Boolean, ByRef firstDetectNumber As Integer, ByRef fileKeyList As List(Of String), ByRef regKeyList As List(Of String),
                                ByRef excludeKeyList As List(Of String), ByRef regKeyLineCounts As List(Of String), ByRef fileKeyLineCounts As List(Of String), ByRef excludeKeyLineCounts As List(Of String), ByRef number_of_errors As Integer,
                                ByRef detectFileList As List(Of String), ByRef detectFileLineCounts As List(Of String), ByRef detectList As List(Of String), ByRef detectLineCounts As List(Of String), ByRef entryHasDefault As Boolean, ByVal entryName As String)


        'reset our counters for the numbers next to commands when we move to the next entry
        curFileKeyNumber = 1
        curRegKeyNumber = 1
        curDetectFileNumber = 1
        curDetectNumber = 1
        curExcludeKeyNumber = 1
        firstDetectFileNumber = False
        firstDetectNumber = False

        'Observe whether or not we detected a default state (and ignore it when we're looking at sequential newline characters) 
        If Not entryHasDefault And Not entryName.Equals("") And (fileKeyLineCounts.Count > 0 Or regKeyLineCounts.Count > 0) Then
            Console.WriteLine("Error: " & entryName & " appears to be missing its default state. All entries should contain 'Default=False'" & Environment.NewLine)
            number_of_errors = number_of_errors + 1
        End If

        entryHasDefault = False

        'create the (soon to be) sorted versions of our lists
        Dim sortedFKList As New List(Of String)
        sortedFKList.AddRange(fileKeyList)

        Dim sortedRKList As New List(Of String)
        sortedRKList.AddRange(regKeyList)

        Dim sortedEKList As New List(Of String)
        sortedEKList.AddRange(excludeKeyList)

        Dim sortedDList As New List(Of String)
        sortedDList.AddRange(detectList)

        Dim sortedDFList As New List(Of String)
        sortedDFList.AddRange(detectFileList)

        'Assess our Keys and cleverly sort their stringy selves
        If fileKeyList.Count > 1 Then

            replaceAndSort(fileKeyList, sortedFKList, "|", " \ \", "keys")
            findOutOfPlace(fileKeyList, sortedFKList, "FileKey", number_of_errors, fileKeyLineCounts)

        End If
        fileKeyList.Clear()
        fileKeyLineCounts.Clear()

        If regKeyLineCounts.Count > 1 Then

            replaceAndSort(regKeyList, sortedRKList, "|", " \ \", "keys")
            findOutOfPlace(regKeyList, sortedRKList, "RegKey", number_of_errors, regKeyLineCounts)

        End If
        regKeyList.Clear()
        regKeyLineCounts.Clear()

        If excludeKeyLineCounts.Count > 1 Then

            replaceAndSort(excludeKeyList, sortedEKList, "|", " \ \", "keys")
            findOutOfPlace(excludeKeyList, sortedEKList, "ExcludeKey", number_of_errors, excludeKeyLineCounts)

        End If
        excludeKeyList.Clear()
        excludeKeyLineCounts.Clear()

        If detectFileLineCounts.Count > 1 Then

            replaceAndSort(detectFileList, sortedDFList, "|", " \ \", "keys")
            findOutOfPlace(detectFileList, sortedDFList, "DetectFile", number_of_errors, detectFileLineCounts)

        End If
        detectFileList.Clear()
        detectFileLineCounts.Clear()

        If detectLineCounts.Count > 1 Then

            replaceAndSort(detectList, sortedDList, "|", " \ \", "keys")
            findOutOfPlace(detectList, sortedDList, "Detect", number_of_errors, detectLineCounts)

        End If
        detectList.Clear()
        detectLineCounts.Clear()

    End Sub
End Module