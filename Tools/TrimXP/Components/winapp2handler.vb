Option Strict On
Module winapp2handler

    'Take in a list of strings (key values) to be sorted, replace any characters that need replacing (also pad single digit numbers), sort the list, 
    'And restore the original state of those characters, returning the sorted list
    Public Function replaceAndSort(ListToBeSorted As List(Of String), characterToReplace As String, replacementText As String) As List(Of String)

        Dim sortedEntryList As New List(Of String)

        'track any renames we must make using these 1:1 lists 
        Dim renamedList As New List(Of String)
        Dim originalNamesList As New List(Of String)

        For i As Integer = 0 To ListToBeSorted.Count - 1

            Dim item As String = ListToBeSorted(i)

            If item.Contains(characterToReplace) Then

                originalNamesList.Add(item)
                Dim renamedItem As String = item.Replace(characterToReplace, replacementText)
                renamedList.Add(renamedItem)

                'replace the item in the given list
                ListToBeSorted(i) = renamedItem
                item = renamedItem

            End If

            'Prefix any singular numbers with 0 to maintain their first order precedence during sorting
            findAndReplaceNumbers(item, originalNamesList, renamedList, ListToBeSorted)

        Next

        'copy the modified list to be sorted and sort it
        sortedEntryList.AddRange(ListToBeSorted)
        sortedEntryList.Sort()

        'Undo any changes we've made to both lists for the purposes of sorting 
        If renamedList.Count > 0 Then
            For i As Integer = 0 To renamedList.Count - 1
                sortedEntryList(sortedEntryList.IndexOf(renamedList(i))) = originalNamesList(i)
                ListToBeSorted(ListToBeSorted.IndexOf(renamedList(i))) = originalNamesList(i)
            Next
        End If

        Return sortedEntryList
    End Function

    'Pad any instances of single-digit numbers with a 0 and track the changes, with a mindfulness for previous changes that may have been made
    Private Sub findAndReplaceNumbers(ByRef item As String, ByRef originalNameList As List(Of String), ByRef renamedList As List(Of String), ByRef listToBeSorted As List(Of String))

        Dim myChars As Char() = item.ToCharArray

        Dim lastCharWasNum As Boolean = False
        Dim prefixIndicies As New List(Of Integer)
        Dim nextCharIsNum As Boolean = False

        For chind As Integer = 0 To myChars.Count - 1

            Dim chIsDig As Boolean = Char.IsDigit(myChars(chind))

            If Not chIsDig Then lastCharWasNum = False

            'observe if the next character is a number, we only want to pad instances of single digit numbers
            nextCharIsNum = If(chind < myChars.Count - 1, Char.IsDigit(item(chind + 1)), False)

            'At this point, if we're not a digit, we can move on to the next step in the loop
            If Not chIsDig Then
                'If the next character isn't a digit either, we can skip it
                If Not nextCharIsNum Then
                    chind += 1
                End If
                Continue For
            End If

            'observe the previous character to see if it was a digit (since at this point, we are a digit)
            If Not lastCharWasNum Then
                lastCharWasNum = True

                'If the next character isn't a number and the last character wasn't a number and we are, we've found an instance of a single digit number
                If Not nextCharIsNum Then
                    prefixIndicies.Add(chind)
                    chind += 1
                End If
            End If
        Next

        'prefix any numbers that we detected above
        If prefixIndicies.Count >= 1 Then

            Dim tmp As String = item

            'Reverse the indicies so we can insert them without adjustment
            prefixIndicies.Reverse()
            For j As Integer = 0 To prefixIndicies.Count - 1
                tmp = tmp.Insert(prefixIndicies(j), "0")
            Next

            'Keep track of any naming changes we've done 
            If Not renamedList.Contains(item) Then
                originalNameList.Add(item)
                renamedList.Add(tmp)
            Else
                renamedList(renamedList.IndexOf(item)) = tmp
            End If

            'Send the rename back 
            listToBeSorted(listToBeSorted.IndexOf(item)) = tmp

        End If
    End Sub

    Public Sub removeEntries(ByRef sectionList As List(Of winapp2entry), ByRef removalList As List(Of winapp2entry))
        For Each item In removalList
            sectionList.Remove(item)
        Next
        removalList.Clear()
    End Sub

    Public Class winapp2file

        Public entryList As List(Of String)

        Public cEntries As iniFile
        Public fxEntries As iniFile
        Public tbEntries As iniFile
        Public mEntries As iniFile

        Public cEntryLines As List(Of Integer)
        Public fxEntryLines As List(Of Integer)
        Public tbEntryLines As List(Of Integer)
        Public mEntryLines As List(Of Integer)

        Public cEntriesW As List(Of winapp2entry)
        Public fxEntriesW As List(Of winapp2entry)
        Public tbEntriesW As List(Of winapp2entry)
        Public mEntriesW As List(Of winapp2entry)

        Public isNCC As Boolean

        Public dir As String
        Public name As String
        Dim version As String

        Public Sub New(ByVal file As iniFile)
            'retain the sections as separate inifiles comprised of inisections
            cEntries = New iniFile
            fxEntries = New iniFile
            tbEntries = New iniFile
            mEntries = New iniFile

            'Determine if we're the Non-CCleaner variant of the ini
            isNCC = Not file.findCommentLine("; This is the non-CCleaner version of Winapp2 that contains extra entries that were removed due to them being added to CCleaner.") = -1

            If file.comments.Count = 0 Then
                version = "; version 000000"
            Else
                version = If(Not file.comments.Values(0).comment.ToLower.Contains("version"), "; version 000000", file.comments.Values(0).comment)
            End If

            'record the full list of all the entry names
            entryList = New List(Of String)

            'track the starting lines of each entry for sorting purposes
            cEntryLines = New List(Of Integer)
            fxEntryLines = New List(Of Integer)
            tbEntryLines = New List(Of Integer)
            mEntryLines = New List(Of Integer)

            'produce lists of winapp2entries 
            cEntriesW = New List(Of winapp2entry)
            fxEntriesW = New List(Of winapp2entry)
            tbEntriesW = New List(Of winapp2entry)
            mEntriesW = New List(Of winapp2entry)

            For Each section In file.sections.Values
                entryList.Add(section.name)
                Dim tmpwa2entry As New winapp2entry(section)
                If tmpwa2entry.langSecRef.Count > 0 Then
                    Select Case tmpwa2entry.langSecRef(0).value
                        Case "3029"
                            addToInnerFile(cEntries, cEntriesW, cEntryLines, tmpwa2entry, section)
                        Case "3026"
                            addToInnerFile(fxEntries, fxEntriesW, fxEntryLines, tmpwa2entry, section)
                        Case "3030"
                            addToInnerFile(tbEntries, tbEntriesW, tbEntryLines, tmpwa2entry, section)
                        Case Else
                            addToInnerFile(mEntries, mEntriesW, mEntryLines, tmpwa2entry, section)
                    End Select
                Else
                    addToInnerFile(mEntries, mEntriesW, mEntryLines, tmpwa2entry, section)
                End If
            Next
        End Sub

        'Add an inisection to the inner tracking inifile for its section type and index the winapp2entry form in the internal tracking list
        Private Sub addToInnerFile(ByRef innerFile As iniFile, ByRef innerFileW As List(Of winapp2entry), ByRef innerLines As List(Of Integer), entry As winapp2entry, section As iniSection)
            If Not innerFile.sections.Keys.Contains(section.name) Then
                innerFile.sections.Add(section.name, section)
                innerLines.Add(section.startingLineNumber)
                innerFileW.Add(entry)
            End If
        End Sub

        'return the total count of entries stored in the internal inifiles
        Public Function count() As Integer
            Return cEntries.sections.Count + tbEntries.sections.Count + fxEntries.sections.Count + mEntries.sections.Count
        End Function

        'sort the internal inifiles in winapp2 entry precendence order
        Public Sub sortInneriniFiles()
            sortIniFile(cEntries, replaceAndSort(cEntries.getSectionNamesAsList, "-", "  "))
            sortIniFile(fxEntries, replaceAndSort(fxEntries.getSectionNamesAsList, "-", "  "))
            sortIniFile(tbEntries, replaceAndSort(tbEntries.getSectionNamesAsList, "-", "  "))
            sortIniFile(mEntries, replaceAndSort(mEntries.getSectionNamesAsList, "-", "  "))
        End Sub

        'Rebuilds a list of winapp2entry objects back into inifile section objects and returns the inifile to the caller
        Private Function rebuildInnerIni(ByRef entryList As List(Of winapp2entry)) As iniFile
            Dim tmpini As New iniFile
            For Each entry In entryList
                tmpini.sections.Add(entry.name, New iniSection(entry.dumpToListOfStrings))
            Next
            Return tmpini
        End Function

        'Update the internal inifile objects
        Public Sub rebuildToIniFiles()
            cEntries = rebuildInnerIni(cEntriesW)
            fxEntries = rebuildInnerIni(fxEntriesW)
            tbEntries = rebuildInnerIni(tbEntriesW)
            mEntries = rebuildInnerIni(mEntriesW)
        End Sub

        'Build the winapp2.ini text for writing back to a file 
        Public Function winapp2string() As String

            Dim fileName As String = If(isNCC, "Winapp2 (Non-CCleaner version)", "Winapp2")
            Dim licLink As String = If(isNCC, "https://github.com/MoscaDotTo/Winapp2/blob/master/Non-CCleaner/License.md", "https://github.com/MoscaDotTo/Winapp2/blob/master/License.md") & Environment.NewLine
            'Version string (YYMMDD format) & entry count 
            Dim out As String = version & Environment.NewLine
            out += "; # of entries: " & count.ToString("#,###") & Environment.NewLine

            out += ";" & Environment.NewLine
            out += "; " & fileName & " is fully licensed under the CC-BY-SA-4.0 license agreement. Please refer to our license agreement before using Winapp2: " & licLink
            out += "; If you plan on modifying, distributing, and/or hosting " & fileName & " for your own program or website, please ask first." & Environment.NewLine
            out += ";" & Environment.NewLine

            If isNCC Then
                out += "; This Is the non-CCleaner version of Winapp2 that contains extra entries that were removed due to them being added to CCleaner." & Environment.NewLine
                out += "; DO Not use this file for CCleaner as the extra cleaners may cause conflicts with CCleaner." & Environment.NewLine
            End If
            out += "; You can get the latest Winapp2 here: https://github.com/MoscaDotTo/Winapp2" & Environment.NewLine
            out += "; Any contributions are appreciated. Please refer to our readme to learn to make your own entries here: https://github.com/MoscaDotTo/Winapp2/blob/master/README.md" & Environment.NewLine
            If Not isNCC Then
                out += "; Is CCleaner taking too long to load with Winapp2? Then download our Winapp2ool and use the trim function to remove unneeded entries in Winapp2. https://github.com/MoscaDotTo/Winapp2/blob/master/Tools/beta/winapp2ool.exe" & Environment.NewLine
            Else
                out += "; Check out our Winapp2ool to download and trim Winapp2 if needed. https://github.com/MoscaDotTo/Winapp2/blob/master/Tools/beta/winapp2ool.exe" & Environment.NewLine
            End If

            out += getEntrySectionString(cEntries, "Chrome/Chromium based browsers")
            out += getEntrySectionString(fxEntries, "Firefox/Mozilla based browsers")
            out += getEntrySectionString(tbEntries, "Thunderbird entries")

            out += Environment.NewLine
            out += mEntries.toString
            Return out
        End Function
    End Class

    Private Function getEntrySectionString(entrySection As iniFile, secName As String) As String
        Dim out As String = ""
        If entrySection.sections.Count > 0 Then
            out += ";" & Environment.NewLine
            out += "; " & secName & Environment.NewLine & Environment.NewLine
            out += entrySection.toString
            out += Environment.NewLine & "; End of " & secName & Environment.NewLine
        End If
        Return out
    End Function

    Public Class winapp2entry
        Public name As String
        Public fullname As String
        Public detectOS As New List(Of iniKey)
        Public langSecRef As New List(Of iniKey)
        Public sectionKey As New List(Of iniKey)
        Public specialDetect As New List(Of iniKey)
        Public detects As New List(Of iniKey)
        Public detectFiles As New List(Of iniKey)
        Public warningKey As New List(Of iniKey)
        Public defaultKey As New List(Of iniKey)
        Public fileKeys As New List(Of iniKey)
        Public regKeys As New List(Of iniKey)
        Public excludeKeys As New List(Of iniKey)
        Public errorKeys As New List(Of iniKey)
        Public keyListList As New List(Of List(Of iniKey))
        Public lineNum As New Integer

        Public Sub New(ByVal section As iniSection)
            fullname = section.getFullName
            name = section.name
            keyListList.Add(detectOS)
            keyListList.Add(langSecRef)
            keyListList.Add(sectionKey)
            keyListList.Add(specialDetect)
            keyListList.Add(detects)
            keyListList.Add(detectFiles)
            keyListList.Add(defaultKey)
            keyListList.Add(warningKey)
            keyListList.Add(fileKeys)
            keyListList.Add(regKeys)
            keyListList.Add(excludeKeys)
            keyListList.Add(errorKeys)
            lineNum = section.startingLineNumber

            'construct the keylists
            Dim keylist As New List(Of String)
            keylist.AddRange(New String() {"detectos", "langsecref", "section", "specialdetect", "detect", "detectfile", "default", "warning", "filekey", "regkey", "excludekey"})
            section.constructKeyLists(keylist, keyListList)
        End Sub

        'dump the keys in the keylist back to a list of strings in proper winapp2.ini order 
        Public Function dumpToListOfStrings() As List(Of String)
            Dim outList As New List(Of String)
            outList.Add(fullname)
            For Each lst In keyListList
                For Each key In lst
                    outList.Add(key.toString)
                Next
            Next
            Return outList
        End Function
    End Class

    Public Class winapp2KeyParameters
        Public paramString As String = ""
        Public argsList As New List(Of String)
        Public flagString As String = ""
        Public keyType As String = ""

        'This method will parameterize a key's arguments into a small object. 
        Public Sub New(key As iniKey)
            keyType = key.keyType

            Select Case key.keyType
                Case "FileKey"
                    Dim splitKey As String() = key.value.Split(CChar("|"))
                    If splitKey.Count > 1 Then
                        paramString = splitKey(0)
                        argsList.AddRange(splitKey(1).Split(CChar(";")))
                        If splitKey.Count >= 3 Then flagString = splitKey(2)
                    End If

                Case "ExcludeKey"
                    Dim splitKey As String() = key.value.Split(CChar("|"))
                    Select Case splitKey.Count
                        Case 2
                            paramString = splitKey(1)
                            flagString = splitKey(0)
                        Case 3
                            paramString = splitKey(1)
                            argsList.AddRange(splitKey(2).Split(CChar(";")))
                            flagString = splitKey(0)
                    End Select
                Case "RegKey"
                    Dim splitKey As String() = key.value.Split(CChar("|"))
                    paramString = splitKey(0)
                    If splitKey.Count > 1 Then argsList.Add(splitKey(1))
            End Select
        End Sub

        'Reconstruct a filekey's format in the form of path|file;file;file..|FLAG
        'This also trims any empty comments in a non transparent way
        Public Sub reconstructKey(ByRef key As iniKey)

            Dim out As String = ""
            out += paramString & "|"
            If argsList.Count > 1 Then
                For i As Integer = 0 To argsList.Count - 2
                    If Not argsList(i) = "" Then out += argsList(i) & ";"
                Next
            End If
            out += argsList.Last
            If Not flagString = "" Then out += "|" & flagString
            key.value = out

        End Sub

    End Class

End Module