Module winapp2handler

    Public Function replaceAndSort(ByRef ListToBeSorted As List(Of String), characterToReplace As String, ByVal replacementText As String) As List(Of String)
        'Take in a list of strings (key values) to be sorted, replace any characters that need replacing (also pad single digit numbers), sort the list, and restore the original state of those characters
        'Return the sorted version of the list

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

    Private Sub findAndReplaceNumbers(ByRef item As String, ByRef originalNameList As List(Of String), ByRef renamedList As List(Of String), ByRef listToBeSorted As List(Of String))
        'Pad any instances of single-digit numbers with a 0 and track the changes, with a mindfulness for previous changes that may have been made

        Dim myChars As Char() = item.ToCharArray

        Dim lastCharWasNum As Boolean = False
        Dim prefixIndicies As New List(Of Integer)
        Dim nextCharIsNum As Boolean = False

        For chind As Integer = 0 To myChars.Count - 1

            Dim chIsDig As Boolean = Char.IsDigit(myChars(chind))

            'observe if the next character is a number, we only want to pad instances of single digit numbers
            If chind < myChars.Count - 1 Then
                nextCharIsNum = Char.IsDigit(item(chind + 1))
            Else
                nextCharIsNum = False
            End If

            'observe the previous character for the same reason
            If Not lastCharWasNum Then
                If chIsDig Then
                    lastCharWasNum = True
                    If Not nextCharIsNum Then
                        prefixIndicies.Add(chind)
                    End If
                Else
                    lastCharWasNum = False
                End If
            Else
                If Not chIsDig Then
                    lastCharWasNum = False
                End If
            End If
        Next

        'prefix any numbers that we detected above
        If prefixIndicies.Count >= 1 Then

            Dim tmp As String = item

            For j As Integer = 0 To prefixIndicies.Count - 1

                tmp = tmp.Insert(prefixIndicies(j), "0")

                'each time we insert our leading zero, remember to adjust the remaining indicies by 1
                For k As Integer = j + 1 To prefixIndicies.Count - 1
                    prefixIndicies(k) += 1
                Next
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
        Public duplicateList

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
            version = file.comments.Values(0).comment
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
                            If Not cEntries.sections.Keys.Contains(section.name) Then
                                cEntries.sections.Add(section.name, section)
                                cEntryLines.Add(section.startingLineNumber.ToString)
                                cEntriesW.Add(tmpwa2entry)
                            End If
                        Case "3026"
                            If Not fxEntries.sections.Keys.Contains(section.name) Then
                                fxEntries.sections.Add(section.name, section)
                                fxEntryLines.Add(section.startingLineNumber.ToString)
                                fxEntriesW.Add(tmpwa2entry)
                            End If
                        Case "3030"
                            If Not tbEntries.sections.Keys.Contains(section.name) Then
                                tbEntries.sections.Add(section.name, section)
                                tbEntryLines.Add(section.startingLineNumber.ToString)
                                tbEntriesW.Add(tmpwa2entry)
                            End If
                        Case Else
                            If Not mEntries.sections.Keys.Contains(section.name) Then
                                mEntries.sections.Add(section.name, section)
                                mEntryLines.Add(section.startingLineNumber.ToString)
                                mEntriesW.Add(tmpwa2entry)
                            End If
                    End Select
                Else
                    mEntries.sections.Add(section.name, section)
                    mEntryLines.Add(section.startingLineNumber.ToString)
                    mEntriesW.Add(tmpwa2entry)
                End If

            Next
        End Sub

        Public Function count() As Integer
            Return cEntries.sections.Count + tbEntries.sections.Count + fxEntries.sections.Count + mEntries.sections.Count
        End Function

        Public Sub sortInneriniFiles()
            sortIniFile(cEntries, replaceAndSort(cEntries.getSectionNamesAsList, "-", "  "))
            sortIniFile(fxEntries, replaceAndSort(fxEntries.getSectionNamesAsList, "-", "  "))
            sortIniFile(tbEntries, replaceAndSort(tbEntries.getSectionNamesAsList, "-", "  "))
            sortIniFile(mEntries, replaceAndSort(mEntries.getSectionNamesAsList, "-", "  "))
        End Sub

        Public Sub rebuildToIniFiles()

            Dim newCFile As New iniFile
            Dim newFFile As New iniFile
            Dim newTFile As New iniFile
            Dim newMfile As New iniFile

            For Each entry In cEntriesW
                newCFile.sections.Add(entry.name, New iniSection(entry.dumpToListOfStrings))
            Next

            For Each entry In fxEntriesW
                newFFile.sections.Add(entry.name, New iniSection(entry.dumpToListOfStrings))
            Next

            For Each entry In tbEntriesW
                newTFile.sections.Add(entry.name, New iniSection(entry.dumpToListOfStrings))
            Next

            For Each entry In mEntriesW
                newMfile.sections.Add(entry.name, New iniSection(entry.dumpToListOfStrings))
            Next

            cEntries = newCFile
            fxEntries = newFFile
            tbEntries = newTFile
            mEntries = newMfile

        End Sub

        Public Function winapp2string() As String
            Dim totalEntryCount As Integer = cEntries.sections.Count + tbEntries.sections.Count + fxEntries.sections.Count + mEntries.sections.Count
            Dim out As String = version & Environment.NewLine
            out += "; # of entries: " & totalEntryCount.ToString("#,###") & Environment.NewLine
            If isNCC Then
                out += "; This is the non-CCleaner version of Winapp2 that contains extra entries that were removed due to them being added to CCleaner." & Environment.NewLine
                out += "; DO NOT use this file for CCleaner as the extra cleaners may cause conflicts with CCleaner." & Environment.NewLine
            End If
            out += "; You can get the latest Winapp2 here: https://github.com/MoscaDotTo/Winapp2" & Environment.NewLine
            out += "; Any contributions are appreciated. Please send them to the link above." & Environment.NewLine
            If Not isNCC Then
                out += "; Is CCleaner taking too long to load with Winapp2? Please head to this link and follow the instructions: http://www.winapp2.com/howto.html" & Environment.NewLine
            End If
            out += "; Valid commands can be found on the first post here: https://forum.piriform.com/index.php?showtopic=32310" & Environment.NewLine
            out += "; Please do not host this file anywhere without permission. This is to facilitate proper distribution of the latest version. Thanks." & Environment.NewLine
            out += ";" & Environment.NewLine
            out += "; Chrome/Chromium based browsers." & Environment.NewLine & Environment.NewLine
            out += cEntries.toString
            out += Environment.NewLine & "; End of Chrome/Chromium based browsers." & Environment.NewLine
            out += ";" & Environment.NewLine
            out += "; Firefox/Mozilla based browsers." & Environment.NewLine & Environment.NewLine
            out += fxEntries.toString
            out += Environment.NewLine & "; End of Firefox/Mozilla based browsers." & Environment.NewLine
            out += ";" & Environment.NewLine
            out += "; Thunderbird entries." & Environment.NewLine & Environment.NewLine
            out += tbEntries.toString
            out += Environment.NewLine & "; End of Thunderbird entries." & Environment.NewLine & Environment.NewLine
            out += mEntries.toString
            Return out
        End Function


    End Class

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

End Module
