'    Copyright (C) 2018-2025 Hazel Ward
' 
'    This file is a part of Winapp2ool
' 
'    Winapp2ool is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    Winapp2ool is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with Winapp2ool.  If not, see <http://www.gnu.org/licenses/>.
Option Strict On
''' <summary> This module holds any scans/repairs for <c> WinappDebug </c> that are be disabled by default due either 
''' to incompleteness or by virtue of being out of scope of the normal linting process </summary>
''' Docs last updated: 2021-11-13 | Code last updated: 2021-11-13
Module experimentalScans

    ''' <summary> Holds the entire text of every RegKey observed during duplicate checks between all entries </summary>
    ''' Docs last updated: 2022-07-14 | Code last updated: 2022-07-14
    Private Property regKeyTracker As New HashSet(Of String)

    ''' <summary> Holds the path of every FileKey observed during duplicate checks between all entries </summary>
    ''' Docs last updated: 2022-07-14 | Code last updated: 2022-07-14
    Private Property fileKeyTracker As New HashSet(Of String)

    ''' <summary> Holds the path of every Detect observed during duplicate checks between all entries </summary>
    ''' Docs last updated: 2022-07-14 | Code last updated: 2022-07-14

    Private Property detectTracker As New HashSet(Of String)

    ''' <summary> Holds the path of every DetectFile observed during duplicate checks between all entries </summary>
    ''' Docs last updated: 2022-07-14 | Code last updated: 2022-07-14
    Private Property detectFileTracker As New HashSet(Of String)

    ''' <summary> Empties the key value trackers of their contents </summary>
    ''' Docs last updated: 2022-07-14 | Code last updated: 2022-07-14
    Public Sub resetKeyTrackers()

        regKeyTracker.Clear()
        fileKeyTracker.Clear()
        detectFileTracker.Clear()
        detectTracker.Clear()

    End Sub

    ''' <summary> Attempts to merge FileKeys together if syntactically possible </summary>
    ''' <param name="kl"> A <c> keyList </c> of FileKey format <c> iniKeys </c> which will be assessed for 
    ''' potential to merge multiple keys into a single key </param>
    ''' Docs last updated: 2022-07-14 | Code last updated: 2022-07-14
    Public Sub cOptimization(ByRef kl As keyList)

        ' No need to check for duplicates if there can't possibly be any 
        If kl.KeyCount < 2 Then Return

        Dim dupes As New keyList
        Dim newKeys As New keyList
        Dim flagList As New strList
        Dim paramList As New strList
        newKeys.add(kl.Keys)

        For i = 0 To kl.KeyCount - 1

            Dim tmpWa2 As New winapp2KeyParameters(kl.Keys(i))

            ' If we have yet to record any params, record them and move on
            If paramList.Count = 0 Then tmpWa2.trackParamAndFlags(paramList, flagList) : Continue For

            ' FileKey Case: 
            ' The folder provided has appeared in another key
            ' The flagstring (RECURSE, REMOVESELF, "") for both keys matches
            ' The first key to appear will have the parameters from the second key appended into its own 
            ' Then, the second key is removed 
            If paramList.contains(tmpWa2.PathString) Then

                gLog($"{kl.Keys(i)} has a path that matches another key")

                For j = 0 To paramList.Count - 1

                    ' If the current processing key's path has already appeared, create a new temporary winapp2entry at the index of the first 
                    ' item in the paramlist whose path and flag matches the current key 
                    If tmpWa2.PathString = paramList.Items(j) And tmpWa2.FlagString = flagList.Items(j) Then

                        gLog($"Matching key has index {j} in the unique path list")

                        Dim keyToMergeInto As New winapp2KeyParameters(newKeys.Keys(j))
                        Dim mergeKeyStr = ""
                        keyToMergeInto.addArgs(mergeKeyStr)
                        tmpWa2.ArgsList.ForEach(Sub(arg) mergeKeyStr += $";{arg}")

                        If tmpWa2.FlagString <> "None" Then mergeKeyStr += $"|{tmpWa2.FlagString}"
                        dupes.add(kl.Keys(i))
                        gLog($"Key will be merged and have the new value: {mergeKeyStr}")

                        ' Overwrite the key with the same index in the unique path list with the new parameters list
                        newKeys.Keys(j) = New iniKey(mergeKeyStr)
                        Exit For

                    End If
                Next

                tmpWa2.trackParamAndFlags(paramList, flagList)

            Else

                tmpWa2.trackParamAndFlags(paramList, flagList)

            End If
        Next

        ' Print out any observations to the user 
        If dupes.KeyCount > 0 Then

            newKeys.remove(dupes.Keys)

            For i = 0 To newKeys.KeyCount - 1

                newKeys.Keys(i).Name = $"FileKey{i + 1}"

            Next

            printOptiSect("Optimization opportunity detected", kl)
            printOptiSect("The following keys can be merged into other keys:", dupes)
            printOptiSect("The resulting keyList will be reduced to: ", newKeys)

            If lintOpti.ShouldRepair Then kl.Keys = newKeys.Keys

        End If

    End Sub

    ''' <summary> Prints output from the Optimization function </summary>
    ''' <param name="boxStr"> The text to be printed in the optimization section box </param>
    ''' <param name="kl"> The list of <c> iniKeys </c>to be printed beneath the box </param>
    ''' Docs last updated: 2022-07-14 | Code last updated: 2022-07-14
    Private Sub printOptiSect(boxStr As String, kl As keyList)

        print(3, boxStr, buffr:=True, trailr:=True)
        kl.Keys.ForEach(Sub(key) cwl(key.toString))
        cwl()

    End Sub

    ''' <summary> Sets up the duplicate key text checker with the proper tracker </summary>
    ''' <param name="key"> An <c> iniKey </c> to have its value audited against the duplicate list </param>
    ''' Docs last updated: 2022-07-14 | Code last updated: 2022-07-14
    Public Sub cDuplicateKeysBetweenEntries(key As iniKey)

        Select Case key.KeyType

            Case "RegKey"

                auditDupe(regKeyTracker, key)

            Case "FileKey"

                auditDupe(fileKeyTracker, key)

            Case "Detect"

                auditDupe(detectTracker, key)

            Case "DetectFile"

                auditDupe(detectFileTracker, key)

        End Select

    End Sub

    ''' <summary>
    ''' Tracks whether or not the value of a key has been obsered multiple times during this lint session. FileKeys are considered
    ''' to be potential duplicates if they have the same path parameter but different file parameters, All other keys are only considered potential 
    ''' duplicates if their entire parameterization is identical. 
    ''' </summary>
    ''' <param name="tracker"> The set of all values of keys of the type given by <c> <paramref name="key"/> </c> to have been observed during this lint session </param>
    ''' <param name="key"> A particular iniKey to check against the set of observed values </param>
    ''' Docs last updated: 2022-07-14 | Code last updated: 2022-07-14
    Private Sub auditDupe(ByRef tracker As HashSet(Of String), key As iniKey)

        Dim tmpKey As New winapp2KeyParameters(key)
        Dim UpperKeyText As String
        Dim RawKeyText As String

        Select Case True

            Case key.KeyType = "FileKey"

                ' For FileKeys we are interested in the case where paths collide, but perhaps have different file parameters
                UpperKeyText = tmpKey.PathString.ToUpperInvariant
                RawKeyText = tmpKey.PathString

            Case Else

                ' For other keys, we're interested in the case where the entire value matches 
                UpperKeyText = key.Value.ToUpperInvariant
                RawKeyText = key.Value

        End Select

        ' If the cased key text is in the tracking set, inform the user (this will be NOISY) 
        print(3, $"{RawKeyText} exists in multiple places", cond:=tracker.Contains(UpperKeyText))

        ' Add the current text to the tracker if it hasn't been already 
        tracker.Add(UpperKeyText)

    End Sub

End Module