Imports System.IO
Module Module1

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
        Dim curDetectFileNumber As Integer = 1
        Dim curDetectNumber As Integer = 1
        Dim firstDetectNumber As Boolean = False
        Dim firstDetecFileNumber As Boolean = False
        Dim havePassedTB As Boolean = False
        Dim FileKeyList As New List(Of String)
        Dim RegKeyList As New List(Of String)

        Dim FileKeyLineCounts As New List(Of String)
        Dim RegKeyLineCounts As New List(Of String)
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

                'Whitespace checks

                If command = "" Then
                    'reset our counters for the numbers next to commands when we move to the next entry
                    curFileKeyNumber = 1
                    curRegKeyNumber = 1
                    curDetectFileNumber = 1
                    curDetectNumber = 1
                    firstDetecFileNumber = False
                    firstDetectNumber = False

                    'Assess the sorted state of our filekeys and regkeys
                    Dim SortedFKList As New List(Of String)
                    SortedFKList.AddRange(FileKeyList)

                    Dim SortedRKList As New List(Of String)
                    SortedRKList.AddRange(RegKeyList)

                    Dim FileKeyRenamedList As New List(Of String)
                    Dim FileKeyOriginalNames As New List(Of String)

                    Dim RegKeyRenamedList As New List(Of String)
                    Dim RegKeyOriginalNames As New List(Of String)

                    'Assess our FileKeys
                    'do a necessary replacement
                    If FileKeyList.Count > 1 Then

                        replaceAndSort(FileKeyList, SortedFKList, "|", " \ \", "keys")
                        findOutOfPlace(FileKeyList, SortedFKList, " Error: FileKey Alphabetization: ", " appears to be out of place, it should follow: ", "", "Follows: ", "", number_of_errors, FileKeyLineCounts)

                    End If
                    FileKeyList.Clear()
                    FileKeyLineCounts.Clear()

                    'now do all that again but for regkeys
                    If RegKeyLineCounts.Count > 1 Then

                        replaceAndSort(RegKeyList, SortedRKList, "|", " \ \", "keys")
                        findOutOfPlace(RegKeyList, SortedRKList, " Error: RegKey Alphabetization: ", " appears to be out of place, it should follow: ", "", "Follows: ", "", number_of_errors, RegKeyLineCounts)

                    End If
                    RegKeyList.Clear()
                    RegKeyLineCounts.Clear()

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


                'Check for duplicate titles
                If command.StartsWith("[") Then


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

                End If

                'Check for spelling errors in "LangSecRef"
                If command.StartsWith("Lan") Then
                    If command.Contains("LangSecRef=") = False Then
                        Console.WriteLine("Line:  " & linecount & " Error: 'LangSecRef' entry is incorrectly spelled or formatted." & Environment.NewLine & "Command: " & command & Environment.NewLine)
                        number_of_errors = number_of_errors + 1
                    End If
                End If

                If command = "Default=True" Or command = "default=true" Then
                    Console.WriteLine("Line: " & linecount & " Error: All entries should be disabled by default." & Environment.NewLine & "Command: " & command & Environment.NewLine)
                    number_of_errors = number_of_errors + 1
                End If

                'Check for environmental variable spacing errors
                If command.Contains("%Program Files%") Then
                    Console.WriteLine("Line: " & linecount & " Error: '%ProgramFiles%' variable should not have spacing." & Environment.NewLine & "Command: " & command & Environment.NewLine)
                    number_of_errors = number_of_errors + 1
                End If

                'Check for cleaning command spelling errors (files)
                If command.StartsWith("Fi") Then

                    Dim cmdList As String() = command.Split("=")
                    FileKeyList.Add(cmdList(1))
                    FileKeyLineCounts.Add(linecount)
                    If command.Contains("FileKey") = False Then
                        Console.WriteLine("Line: " & linecount & " Error: 'FileKey' entry is incorrectly spelled or formatted. Spelling should be CamelCase." & Environment.NewLine & "Command: " & command & Environment.NewLine)
                        number_of_errors = number_of_errors + 1
                    End If
                    If command.Contains("FileKey" & curFileKeyNumber) = False Then
                        Console.WriteLine("Line: " & linecount & " Error: 'FileKey' entry is incorrectly numbered: Expected FileKey" & curFileKeyNumber & " found " & Environment.NewLine & "Command: " & command & Environment.NewLine)
                        number_of_errors = number_of_errors + 1
                    End If
                    Dim iteratorCheckerList() As String = Split(command, "|")

                    'check for incorrect spellings of RECURSE or REMOVESELF
                    If iteratorCheckerList.Length > 2 Then
                        If iteratorCheckerList(2).Contains("RECURSE") = False And iteratorCheckerList(2).Contains("REMOVESELF") = False Then

                            Console.WriteLine("Line: " & linecount & " Error: 'RECURSE' or 'REMOVESELF' entry is incorrectly spelled, found " & Environment.NewLine & "Command:  " & command & Environment.NewLine)
                            number_of_errors = number_of_errors + 1

                        End If
                    End If

                    'check for missing pipe symbol on recurse and removeself
                    If command.Contains("RECURSE") = True And Not command.Contains("|RECURSE") Then
                        Console.WriteLine("Line: " & linecount & " Error: Missing pipe symbol | before RECURSE" & Environment.NewLine & "Command:  " & command & Environment.NewLine)
                        number_of_errors = number_of_errors + 1
                    End If
                    If command.Contains("REMOVESELF") = True And Not command.Contains("|REMOVESELF") Then
                        Console.WriteLine("Line: " & linecount & " Error: Missing pipe symbol | before REMOVESELF" & Environment.NewLine & "Command:  " & command & Environment.NewLine)
                        number_of_errors = number_of_errors + 1
                    End If
                    curFileKeyNumber = curFileKeyNumber + 1
                End If

                'Check for cleaning command spelling errors (registry keys)
                If command.StartsWith("Re") Then
                    Dim cmdList2 As String() = command.Split("=")
                    RegKeyList.Add(cmdList2(1))
                    RegKeyLineCounts.Add(linecount)

                    If command.Contains("RegKey") = False Then
                        Console.WriteLine("Line: " & linecount & " Error: 'RegKey' entry is incorrectly spelled or formatted. Spelling should be CamelCase." & Environment.NewLine & "Command: " & command & Environment.NewLine)
                        number_of_errors = number_of_errors + 1
                    End If
                    If command.Contains("RegKey" & curRegKeyNumber) = False Then
                        Console.WriteLine("Line: " & linecount & " Error: 'RegKey' entry is incorrectly numbered: Expected RegKey" & curRegKeyNumber & " found " & Environment.NewLine & "Command: " & command & Environment.NewLine)
                    End If
                    curRegKeyNumber = curRegKeyNumber + 1
                End If

                'Check for missing numbers next to cleaning commands
                If command.StartsWith("FileKey=") Or command.StartsWith("RegKey=") Then
                    Console.WriteLine("Line: " & linecount & " Error: Cleaning path entry needs to have a trailing number." & Environment.NewLine & "Command: " & command & Environment.NewLine)
                    number_of_errors = number_of_errors + 1
                End If

                'Check for Detect that contains a filepath instead of a registry path
                If command.StartsWith("Detect=%") Or command.StartsWith("Detect=C:\") Then
                    Console.WriteLine("Line: " & linecount & " Error: 'Detect' can only be used for registry key paths." & Environment.NewLine & "Command: " & command & Environment.NewLine)
                    number_of_errors = number_of_errors + 1
                End If

                'Check to make sure Detects are properly numbered
                If command.StartsWith("DetectF") = False And command.StartsWith("DetectO") = False And command.StartsWith("Detect") Then


                    'make sure we notice if there are multiple detects but the first is missing a number

                    If curDetectNumber = 1 Then
                        If command.StartsWith("Detect=") Or (command.StartsWith("Detect=") = False And command.StartsWith("Detect1=") = False) Then
                            firstDetectNumber = False
                        Else
                            firstDetectNumber = True
                        End If

                    End If


                    If curDetectNumber = 2 And firstDetectNumber = False Then
                        Console.WriteLine("Line: " & linecount - 1 & " Error: 'Detect" & curDetectNumber & "' detected without preceding 'Detect" & curDetectNumber - 1 & "'" & Environment.NewLine)
                        number_of_errors = number_of_errors + 1
                    End If
                    If curDetectNumber > 1 Then
                        If command.Contains("Detect" & curDetectNumber) = False Then
                            Console.WriteLine("Line: " & linecount & " Error: 'Detect' entry is incorrectly numbered: Expected Detect" & curDetectNumber & " found " & Environment.NewLine & "Command: " & command & Environment.NewLine)
                            number_of_errors = number_of_errors + 1
                        End If
                    End If
                    curDetectNumber = curDetectNumber + 1
                End If

                'Check for detectfile that contains a registry path
                If command.StartsWith("DetectFile=HKLM") Or command.StartsWith("DetectFile=HKCU") Or command.StartsWith("DetectFile=HKC") Or command.StartsWith("DetectFile=HKCR") Then
                    Console.WriteLine("Line: " & linecount & " Error: 'DetectFile' can only be used for filesystem paths." & Environment.NewLine & "Command: " & command & Environment.NewLine)
                    number_of_errors = number_of_errors + 1
                End If



                'make sure we notice if there are multiple detectfiles but the first is missing a number
                If command.StartsWith("DetectFile") Then

                    If curDetectFileNumber = 1 Then

                        If command.StartsWith("DetectFile=") Or (command.StartsWith("DetectFile=") = False And command.StartsWith("DetectFile1=") = False) Then
                            firstDetecFileNumber = False
                        Else
                            firstDetecFileNumber = True
                        End If

                    End If

                    If curDetectFileNumber = 2 And firstDetecFileNumber = False Then
                        Console.WriteLine("Line: " & linecount - 1 & " Error: 'DetectFile" & curDetectFileNumber & "' detected without preceding 'DetectFile" & curDetectFileNumber - 1 & "'" & Environment.NewLine)
                        number_of_errors = number_of_errors + 1
                    End If

                    If curDetectFileNumber > 1 Then
                        If command.Contains("DetectFile" & curDetectFileNumber) = False Then
                            Console.WriteLine("Line: " & linecount & " Error: 'DetectFile' entry is incorrectly numbered: Expected DetectFile" & curDetectFileNumber & " found " & Environment.NewLine & "Command: " & command & Environment.NewLine)
                            number_of_errors = number_of_errors + 1
                        End If

                    End If
                    curDetectFileNumber = curDetectFileNumber + 1

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

            'replace any instance of "-" with "  " for the purposes of sorting 
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

End Module