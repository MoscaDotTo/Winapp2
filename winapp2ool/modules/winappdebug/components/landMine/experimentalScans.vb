'    Copyright (C) 2018-2021 Hazel Ward
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
    ''' Docs last updated: 2021-11-13 | Code last updated: 2021-11-13
    Private Property masterRegKeyList As New strList
    ''' <summary> Holds the path of every FileKey observed during duplicate checks between all entries </summary>
    ''' Docs last updated: 2021-11-13 | Code last updated: 2021-11-13
    Private Property masterFileKeyList As New strList

    ''' <summary>
    ''' Empties the contents of the master key lists used for duplicate checking between entries 
    ''' </summary>
    ''' Docs last updated: 2021-11-13 | Code last updated: 2021-11-13
    Public Sub resetMasterKeyLists()
        masterRegKeyList.Items.Clear()
        masterFileKeyList.Items.Clear()
    End Sub

    ''' <summary> Attempts to merge FileKeys together if syntactically possible </summary>
    ''' <param name="kl"> A <c> keyList </c> of FileKey format <c> iniKeys </c> which will be assessed for 
    ''' potential to merge multiple keys into a single key </param>
    ''' Docs last updated: 2021-11-13 | Code last updated: 2021-11-13
    Public Sub cOptimization(ByRef kl As keyList)
        If kl.KeyCount < 2 Then Exit Sub
        Dim dupes As New keyList
        Dim newKeys As New keyList
        Dim flagList As New strList
        Dim paramList As New strList
        newKeys.add(kl.Keys)
        For i = 0 To kl.KeyCount - 1
            Dim tmpWa2 As New winapp2KeyParameters(kl.Keys(i))
            ' If we have yet to record any params, record them and move on
            If paramList.Count = 0 Then tmpWa2.trackParamAndFlags(paramList, flagList) : Continue For
            ' This should handle the case where for a FileKey: 
            ' The folder provided has appeared in another key
            ' The flagstring (RECURSE, REMOVESELF, "") for both keys matches
            ' The first appearing key should have its parameters appended to and the second appearing key should be removed
            If paramList.contains(tmpWa2.PathString) Then
                For j = 0 To paramList.Count - 1
                    If tmpWa2.PathString = paramList.Items(j) And tmpWa2.FlagString = flagList.Items(j) Then
                        Dim keyToMergeInto As New winapp2KeyParameters(kl.Keys(j))
                        Dim mergeKeyStr = ""
                        keyToMergeInto.addArgs(mergeKeyStr)
                        tmpWa2.ArgsList.ForEach(Sub(arg) mergeKeyStr += $";{arg}")
                        If tmpWa2.FlagString <> "None" Then mergeKeyStr += $"|{tmpWa2.FlagString}"
                        dupes.add(kl.Keys(i))
                        newKeys.Keys(j) = New iniKey(mergeKeyStr)
                        Exit For
                    End If
                Next
                tmpWa2.trackParamAndFlags(paramList, flagList)
            Else
                tmpWa2.trackParamAndFlags(paramList, flagList)
            End If
        Next
        If dupes.KeyCount > 0 Then
            newKeys.remove(dupes.Keys)
            For i = 0 To newKeys.KeyCount - 1
                newKeys.Keys(i).Name = $"FileKey{i + 1}"
            Next
            printOptiSect("Optimization opportunity detected", kl)
            printOptiSect("The following keys can be merged into other keys:", dupes)
            printOptiSect("The resulting keyList will be reduced to: ", newKeys)
            If Rules.Last.fixFormat Then kl = newKeys
        End If
    End Sub

    ''' <summary> Prints output from the Optimization function </summary>
    ''' <param name="boxStr"> The text to be printed in the optimization section box </param>
    ''' <param name="kl"> The list of <c> iniKeys </c>to be printed beneath the box </param>
    ''' Docs last updated: 2021-11-13 | Code last updated: 2021-11-13
    Private Sub printOptiSect(boxStr As String, kl As keyList)
        print(3, boxStr, buffr:=True, trailr:=True)
        kl.Keys.ForEach(Sub(key) cwl(key.toString))
        cwl()
    End Sub

    ''' <summary>
    ''' Sets up the duplicate key text checker with the proper master list 
    ''' </summary>
    ''' <param name="key"> An <c> iniKey </c> to have its value audited against the duplicate list </param>
    ''' Docs last updated: 2021-11-13 | Code last updated: 2021-11-13
    Public Sub cDuplicateKeysBetweenEntries(key As iniKey)
        Select Case key.KeyType
            Case "RegKey"
                auditDupe(masterRegKeyList, key)
            Case "FileKey"
                auditDupe(masterFileKeyList, key)
        End Select
    End Sub

    ''' <summary>
    ''' Observes whether or not the text of a FileKey or RegKey has previously been seen during this lint session. FileKeys are considered
    ''' to be potential duplicates if they have the same path parameter but different file parameters, RegKeys are only considered potential 
    ''' duplicates if their entire parameterization is identical. 
    ''' </summary>
    ''' <param name="masterList"> The list of all keys of this type to have been observed during this lint session </param>
    ''' <param name="key"> A particular iniKey to check against the list </param>
    ''' Docs last updated: 2021-11-13 | Code last updated: 2021-11-13
    Private Sub auditDupe(masterList As strList, key As iniKey)
        Dim tmpKey As New winapp2KeyParameters(key)
        If masterList.Count = 0 Then
            masterList.add(If(key.KeyType = "FileKey", tmpKey.PathString.ToUpperInvariant, key.Value.ToUpperInvariant))
        Else
            Dim UpperKeyText As String
            Dim RawKeyText As String
            ' We may be interested in FileKeys who point to the same folder but have different file parameters
            If key.KeyType = "FileKey" Then
                UpperKeyText = tmpKey.PathString.ToUpperInvariant
                RawKeyText = tmpKey.PathString
            Else
                ' We're interested in RegKeys where the entire key matches 
                UpperKeyText = key.Value.ToUpperInvariant
                RawKeyText = key.Value
            End If
            If masterList.contains(UpperKeyText, True) Then
                print(3, $"{RawKeyText} may exist in multiple places")
            Else
                masterList.add(UpperKeyText)
            End If
        End If
    End Sub
End Module