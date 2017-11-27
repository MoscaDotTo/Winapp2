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

        Dim excludeKeyLineCounts As New List(Of String)
        Dim fileKeyLineCounts As New List(Of String)
        Dim regKeyLineCounts As New List(Of String)
        'Create a list of supported environmental variables
        Dim envir_vars As New List(Of String)
        envir_vars.AddRange(New String() {"%userprofile%", "%ProgramFiles%", "%rootdir%", "%windir%", "%appdata%", "%systemdrive%", "%Documents%",
                            "%pictures%", "%video%", "%CommonAppData%", "%LocalAppData%", "%CommonProgramFiles%", "%homedrive%", "%music%", "%tmp%", "%temp%"})

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

                If command = "" Then

                    processEmptyLine(curFileKeyNumber, curRegKeyNumber, curExcludeKeyNumber, curDetectFileNumber, curDetectNumber, firstDetectFileNumber, firstDetectFileNumber,
                                     fileKeyList, regKeyList, excludeKeyList, regKeyLineCounts, fileKeyLineCounts, excludeKeyLineCounts, number_of_errors)
                End If

                If command = "" = False And command.StartsWith(";") = False Then

                    'Check for trailing whitespace
                    If command.EndsWith(" ") Then
                        Console.WriteLine("Line: " & linecount & " Error: Detected unwanted whitespace at end of line." & Environment.NewLine & "Command: " & command & Environment.NewLine)
                        number_of_errors = number_of_errors + 1
                    End If

                    'Check for ending whitespace
                    If command.StartsWith(" ") Then
                        Console.WriteLine("Line: " & linecount & " Error: Detected unwanted whitespace at beginning of line." & Environment.NewLine & "Command: " & command & Environment.NewLine)
                        number_of_errors = number_of_errors + 1
                    End If

                Else
                    If command.Equals("; End of Thunderbird entries.") Then
                        havePassedTB = True
                    End If
                End If

                'Process entries
                If command.StartsWith("[") Then

                    processEntryName(command, entry_titles, linecount, havePassedTB, trimmed_entry_titles, number_of_errors)
                    ' Console.WriteLine("Finished checking entry name")
                End If

                'Check for spelling errors in "LangSecRef"
                If command.StartsWith("L") Then
                    If Not command.Contains("LangSecRef=") Then
                        Console.WriteLine("Line:  " & linecount & " Error: 'LangSecRef' entry is incorrectly spelled or formatted." & Environment.NewLine & "Command: " & command & Environment.NewLine)
                        number_of_errors = number_of_errors + 1
                    End If
                End If

                'Check for environmental variable spacing errors
                If command.Contains("%Program Files%") Then
                    Console.WriteLine("Line: " & linecount & " Error: '%ProgramFiles%' variable should not have spacing." & Environment.NewLine & "Command: " & command & Environment.NewLine)
                    number_of_errors = number_of_errors + 1
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

                End If

                If command.StartsWith("D") Then
                    Dim tmp As String() = Split(command, "=")

                    'We'll make the bold assumption here that all our defaults are spelled correctly, even if not formatted correctly. 
                    If Not command.ToLower.Contains("default") And Not command.ToLower.Contains("detecto") Then

                        If tmp(0).Length <= 8 Then

                            processDetect(command, "Detect", curDetectNumber, firstDetectNumber, number_of_errors, linecount)
                            processRegDetect(command, number_of_errors, linecount)

                        Else

                            processDetect(command, "DetectFile", curDetectFileNumber, firstDetectFileNumber, number_of_errors, linecount)
                            processDetectFile(command, number_of_errors, linecount)
                        End If
                    Else

                        If Not command.Contains("Default") And Not command.ToLower.Contains("detecto") Then
                            Console.WriteLine("Line: " & linecount & " Error: Expected 'Default', found: " & tmp.First)
                            number_of_errors = number_of_errors + 1
                        End If

                        If command.ToLower.Contains("true") Then
                            Console.WriteLine("Line: " & linecount & " Error: All entries should be disabled by default." & Environment.NewLine & "Command: " & command & Environment.NewLine)
                            number_of_errors = number_of_errors + 1
                        End If
                    End If

                End If

                'Check for missing backslashes on environmental variables
                For Each var As String In envir_vars
                    If command.Contains(var) = True Then
                        If command.Contains(var & "\") = False Then
                            Console.WriteLine("Line: " & linecount & " Error: Environmental variables must have a trailing backslash." & Environment.NewLine & "Command: " & command & Environment.NewLine)
                            number_of_errors = number_of_errors + 1
                        End If
                    End If
                Next


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
        findOutOfPlace(trimmed_entry_titles, sortedEList, "Error: Entry Alphabetization. Command: [", "*] may be out of place. It should follow: [", "*]", "Follows: [", "*]", number_of_errors, New List(Of String))

        'Stop the program from closing on completion
        Console.WriteLine("***********************************************" & Environment.NewLine & "Completed analysis of winapp2.ini. " & number_of_errors & " possible errors were detected. " & Environment.NewLine & "Number of entries: " & entry_titles.Count & Environment.NewLine & "Press any key to close.")
        Console.ReadKey()

    End Sub

    Private Sub processDetectFile(ByRef command As String, ByRef number_of_errors As Integer, ByRef lineCount As Integer)

        If command.Contains("=HKLM") Or command.Contains("=HKC") Or command.Contains("=HKU") Then
            Console.WriteLine("Line: " & lineCount & " Error: 'DetectFile' can only be used for filesystem paths." & Environment.NewLine & "Command: " & command & Environment.NewLine)
            number_of_errors = number_of_errors + 1
        End If

    End Sub

    Private Sub processRegDetect(ByRef command As String, ByRef number_of_errors As Integer, ByRef lineCount As Integer)

        If command.Contains("=%") Or command.Contains("=C:\") Then
            Console.WriteLine("Line: " & lineCount & " Error: 'Detect' can only be used for registry key paths." & Environment.NewLine & "Command: " & command & Environment.NewLine)
            number_of_errors = number_of_errors + 1
        End If

    End Sub

    Private Sub processDetect(ByRef command As String, ByVal detectType As String, ByRef number As Integer, ByRef firstNumberStatus As Boolean, ByRef number_of_errors As Integer, ByRef lineCount As Integer)

        If Not command.Contains(detectType) Then
            Console.WriteLine("Line: " & lineCount & " Error: Misformatted " & detectType & " found: " & Environment.NewLine & "Command: " & command)

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
            Console.WriteLine("Line: " & lineCount & " Error: '" & detectType & "' numbering. Expected '" & detectType & "' or '" & detectType & "1, found: " & Environment.NewLine & "Command: " & command & Environment.NewLine)
            number_of_errors = number_of_errors + 1
        End If

        'If we're on our second detect, make sure the first one had a 1 
        If number = 2 And firstNumberStatus = False Then
            Console.WriteLine("Line:  " & lineCount - 1 & " Error: '" & detectType & number & "' detected without preceding '" & detectType & number - 1 & "'" & Environment.NewLine)
            number_of_errors = number_of_errors + 1
        End If

        'otherwise, make sure our numbers match up
        If number > 1 Then
            If command.Contains(detectType & number) = False Then
                Console.WriteLine("Line: " & lineCount & " Error: '" & detectType & "' entry is incorrectly numbered: Expected '" & detectType & number & "'  found " & Environment.NewLine & "Command: " & command & Environment.NewLine)
                number_of_errors = number_of_errors + 1
            End If
        End If
        number = number + 1


    End Sub

    Private Sub processRegKey()
        'We don't do much here... Yet? Stub method in case further checks are needed for RegKeys 

    End Sub

    Private Sub processExcludeKey(ByRef command As String, ByRef number_of_errors As Integer, ByVal lineCount As Integer)

        'Make sure any FILE exclude paths have a backslash before their pipe symbol
        Dim iteratorCheckerList() As String = Split(command, "|")
        If iteratorCheckerList(0).Contains("FILE") Then
            Dim endingslashchecker() As String = Split(command, "\|")
            If endingslashchecker.Count = 1 Then
                Console.WriteLine("Line: " & lineCount & " Error: Missing backslash before pipe symbol." & Environment.NewLine & "Command: " & command & Environment.NewLine)
                number_of_errors = number_of_errors + 1
            End If
        End If

    End Sub

    Private Sub processFileKey(ByRef command As String, ByRef number_of_errors As Integer, ByVal lineCount As Integer)

        Dim iteratorCheckerList() As String = Split(command, "|")

        'check for incorrect spellings of RECURSE or REMOVESELF
        If iteratorCheckerList.Length > 2 Then
            If Not iteratorCheckerList(2).Contains("RECURSE") And Not iteratorCheckerList(2).Contains("REMOVESELF") Then

                Console.WriteLine("Line: " & lineCount & " Error: 'RECURSE' or 'REMOVESELF' entry is incorrectly spelled, found " & Environment.NewLine & "Command:  " & command & Environment.NewLine)
                number_of_errors = number_of_errors + 1

            End If
        End If

        'check for missing pipe symbol on recurse and removeself
        If command.Contains("RECURSE") And Not command.Contains("|RECURSE") Then
            Console.WriteLine("Line: " & lineCount & " Error: Missing pipe symbol | before RECURSE" & Environment.NewLine & "Command:  " & command & Environment.NewLine)
            number_of_errors = number_of_errors + 1
        End If
        If command.Contains("REMOVESELF") And Not command.Contains("|REMOVESELF") Then
            Console.WriteLine("Line: " & lineCount & " Error: Missing pipe symbol | before REMOVESELF" & Environment.NewLine & "Command:  " & command & Environment.NewLine)
            number_of_errors = number_of_errors + 1
        End If

    End Sub

    Private Sub processEntryName(ByRef command As String, ByRef entry_titles As List(Of String), ByRef linecount As Integer, ByRef havePassedTB As Boolean, ByRef trimmed_entry_titles As List(Of String), ByRef number_of_errors As Integer)
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
                trimmed_entry_titles.Add(currentEntry)

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

    Public Sub findOutOfPlace(ByRef someList As List(Of String), ByRef sortedList As List(Of String), ByVal Err1 As String, ByVal Err2 As String, ByVal Err3 As String, ByVal err4 As String, ByVal err5 As String, ByRef number_of_errors As Integer, ByRef LineCountList As List(Of String))
        Dim originalPlacement As New List(Of String)
        originalPlacement.AddRange(someList.ToArray)

        Dim misplacedEntryList As New List(Of String)
        For i As Integer = 0 To someList.Count - 1

            Dim entry As String = someList(i)

            Dim receivedIndex As Integer = i
            Dim sortedIndex As Integer = sortedList.IndexOf(entry)

            If receivedIndex <> sortedIndex Then

                If misplacedEntryList.Contains(entry) = False Then

                    If sortedIndex > receivedIndex Then

                        misplacedEntryList.Add(entry)
                        number_of_errors = number_of_errors + 1

                        someList.Insert(sortedIndex, entry)
                        someList.RemoveAt(receivedIndex)

                        'Adjust i because the item in its position has now changed and we don't want to skip it
                        i = i - 1

                    Else
                        Dim tmpErr1 As String = Err1

                        'some contingency for when we're tracking lines (TODO: Add line tracking for entries)
                        If LineCountList.Count > 0 Then
                            LineCountList.Insert(sortedIndex, LineCountList(i))
                            LineCountList.RemoveAt(i + 1)
                            Err1 = "Line: " & LineCountList(i) & Err1
                        End If
                        Dim shouldBe As String
                        If (sortedIndex = 0) Then
                            shouldBe = " Should be first"
                        Else
                            shouldBe = sortedList(sortedIndex - 1)
                        End If

                        Dim recInd As Integer = originalPlacement.IndexOf(entry)
                        Dim follows As String
                        If recInd = 0 Then
                            follows = " Is first"
                        Else
                            follows = originalPlacement(recInd - 1)
                        End If
                        Console.WriteLine(Err1 & entry & Err2 & shouldBe & Err3 & Environment.NewLine & err4 & follows & err5 & Environment.NewLine)
                        number_of_errors = number_of_errors + 1
                        Err1 = tmpErr1

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
                    number_of_errors = number_of_errors - 1
                    misplacedEntryList.Remove(entry)
                End If
            End If
        Next

        For Each entry As String In misplacedEntryList
            Dim tmperr As String = Err1
            If LineCountList.Count > 0 Then
                Err1 = "Line: " & LineCountList(someList.IndexOf(entry)) & " "
            End If
            Dim sortedIndex As Integer = sortedList.IndexOf(entry)
            Dim shouldBe As String
            Dim recInd As Integer = originalPlacement.IndexOf(entry)
            Dim follows As String

            If (sortedIndex = 0) Then
                shouldBe = " Should be first"
            Else
                shouldBe = sortedList(sortedIndex - 1)
            End If

            If recInd = 0 Then
                follows = " Is first"
            Else
                follows = originalPlacement(recInd - 1)
            End If

            Console.WriteLine(Err1 & entry & Err2 & shouldBe & Err3 & Environment.NewLine & err4 & follows & err5 & Environment.NewLine)

            Err1 = tmperr
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

        'prefix any numbers that we detected above (except the potential last which has already been replaced)
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
        keyList.Add(cmdList(1))
        keyLineCounts.Add(lineCount)

        'make sure the current key is correctly numbered
        If Not command.Contains(keyString & keyNumber) Then
            Console.WriteLine("Line: " & lineCount & " Error: '" & keyString & "' entry is incorrectly spelled or formatted. Expected: " & keyString & keyNumber & "', Found: " & Environment.NewLine & "Command:  " & command & Environment.NewLine)
            number_of_errors = number_of_errors + 1
        End If
        keyNumber = keyNumber + 1

    End Sub

    Public Sub processEmptyLine(ByRef curFileKeyNumber As Integer, ByRef curRegKeyNumber As Integer, ByRef curExcludeKeyNumber As Integer, ByRef curDetectFileNumber As Integer, ByRef curDetectNumber As Integer,
                                ByRef firstDetectFileNumber As Boolean, ByRef firstDetectNumber As Integer, ByRef fileKeyList As List(Of String), ByRef regKeyList As List(Of String),
                                ByRef excludeKeyList As List(Of String), ByRef regKeyLineCounts As List(Of String), ByRef fileKeyLineCounts As List(Of String), ByRef excludeKeyLineCounts As List(Of String), ByRef number_of_errors As Integer)

        'reset our counters for the numbers next to commands when we move to the next entry
        curFileKeyNumber = 1
        curRegKeyNumber = 1
        curDetectFileNumber = 1
        curDetectNumber = 1
        curExcludeKeyNumber = 1
        firstDetectFileNumber = False
        firstDetectNumber = False

        'create the (soon to be) sorted versions of our lists
        Dim sortedFKList As New List(Of String)
        sortedFKList.AddRange(fileKeyList)

        Dim sortedRKList As New List(Of String)
        sortedRKList.AddRange(regKeyList)

        Dim sortedEKList As New List(Of String)
        sortedEKList.AddRange(excludeKeyList)

        'Assess our Keys and cleverly sort their stringy selves
        If fileKeyList.Count > 1 Then

            replaceAndSort(fileKeyList, sortedFKList, "|", " \ \", "keys")
            findOutOfPlace(fileKeyList, sortedFKList, " Error: FileKey Alphabetization: ", " appears to be out of place, it should follow: ", "", "Follows: ", "", number_of_errors, fileKeyLineCounts)

        End If
        fileKeyList.Clear()
        fileKeyLineCounts.Clear()

        If regKeyLineCounts.Count > 1 Then

            replaceAndSort(regKeyList, sortedRKList, "|", " \ \", "keys")
            findOutOfPlace(regKeyList, sortedRKList, " Error: RegKey Alphabetization: ", " appears to be out of place, it should follow: ", "", "Follows: ", "", number_of_errors, regKeyLineCounts)

        End If
        regKeyList.Clear()
        regKeyLineCounts.Clear()

        If excludeKeyLineCounts.Count > 1 Then

            replaceAndSort(excludeKeyList, sortedEKList, "|", " \ \", "keys")
            findOutOfPlace(excludeKeyList, sortedEKList, " Error: ExcludeKey Alphabetization: ", " appears to be out of place, it should follow: ", "", "Follows: ", "", number_of_errors, excludeKeyLineCounts)

        End If
        excludeKeyList.Clear()
        excludeKeyLineCounts.Clear()
    End Sub

End Module