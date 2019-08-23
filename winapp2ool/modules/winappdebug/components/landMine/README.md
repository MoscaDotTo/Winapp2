
# experimentalScans

This module holds any scans/repairs for **WinappDebug** that are be disabled by default due to incompleteness

# Optimization

Optimization is an instance where FileKeys hold the same path as each other but seek different files, using the same (or no) mask. This generally means that the two (or more) keys can be merged into each other to produce the same effect with less syntax. There is generally no advantage to doing this outside of minifying the sometimes very verbose winapp2.ini code

## cOptimization

### Attempts to merge FileKeys together if syntactically possible

```vb
Public Sub cOptimization(ByRef kl As keyList)
    ' Rules1.Last here is lintOpti
    If kl.KeyCount < 2 Or Not Rules.Last.ShouldScan Then Exit Sub
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
```

|Parameter|Type|Description|
|:-|:-|:-
kl|`keyList`|A `keyList` of FileKey format `iniKeys`

## printOptiSect

### Prints output from the Optimization function

```vb
Private Sub printOptiSect(boxStr As String, kl As keyList)
    print(3, boxStr, buffr:=True, trailr:=True)
    kl.Keys.ForEach(Sub(key) cwl(key.toString))
    cwl()
End Sub
```

|Parameter|Type|Description|
|:-|:-|:-
boxStr|`String`|The text to be printed in the optimization section box
kl|`keyList`|The list of `iniKeys` to be printed beneath the box

