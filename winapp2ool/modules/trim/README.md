# Trim

Trim is designed to do as its name implies: trim winapp2.ini. It does this by processing the detection criteria provided in each entry and confirming their existence on the user's machine. Any entries whose detection criteria are invalid for the current system are pruned from the file, resulting in a much smaller file filled only with entries relevant to the current machine. The performance impact of this is most notable on CCleaner which sees a significant performance increase on loading the trimmed file vs. an untrimmed file

## Menu Options

Option|Effect|Notes
:-|:-|:-
Exit|Returns you to the winapp2ool menu|
Run|Trims winapp2.ini using your current settings|Default option
Toggle Download|Enables/Disables downloading of the latest winapp2.ini from GitHub to use as input for the trimmer|Unavailable in offline mode
File Chooser (winapp2.ini)|Opens the interface to choose a new local file path for winapp2.ini|Only shown when downloading is disabled
File Chooser (save)|Opens the interface to select a new save location for the trimmed file|Suggests `"winapp2-trimmed.ini"` as the default renamed file
Reset Settings|Restores the module's settings to their default state, undoing any changes the user has made|Only shown when settings have been modified

## Command-line args

|Arg|Effect|
|:-|:-
`-1f` or `-1d`|Defines the path for winapp2.ini
`-3f` or `-3d`|Defines the path for the save file
`-d`|Enables downloading

### Code Documentation

Below is documentation for how Trim works under the hood

## Properties

*italicized* properties are *private*

|Name|Type|Default Value|Description|
|:-|:-|:-|:-
TrimFile1|`iniFile`|`winapp2.ini` in the current directory|The winapp2.ini that will be trimmed
|TrimFile3|`iniFile`|`winapp2.ini` in the current directory|Holds the path where the output file will be saved to disk. Overwrites the input file by default
*winVer*|`Double`|`Nothing`|The major/minor version number on the current system
*ModuleSettingsChanged*|`Boolean`|`False`|Indicates that the module settings have been modified from their defaults
*DownloadFileToTrim*|`Boolean`|`False`|Indicates that we are downloading a winapp2.ini from GitHub

## handleCmdLine

### Handles the commandline args for Trim

```vb
Public Sub handleCmdLine()
    initDefaultSettings()
    handleDownloadBools(DownloadFileToTrim)
    getFileAndDirParams(TrimFile1, New iniFile, TrimFile3)
    initTrim()
End Sub
```

## initDefaultSettings

### Restores the default state of the module's parameters

```vb
Private Sub initDefaultSettings()
    TrimFile1.resetParams()
    TrimFile3.resetParams()
    DownloadFileToTrim = False
    ModuleSettingsChanged = False
End Sub
```

## remoteTrim

### Trims an `iniFile` from outside the module

```vb
Public Sub remoteTrim(firstFile As iniFile, thirdFile As iniFile, d As Boolean)
    TrimFile1 = firstFile
    TrimFile3 = thirdFile
    DownloadFileToTrim = d
    initTrim()
End Sub
```

|Parameter|Type|Description
:-|:-|:-
firstFile|`iniFile`|The winapp2.ini file to be trimmed
thirdFile|`iniFile`|The `iniFile` containing the path on disk to which the trimmed file will be saved
d|`Boolean`|Indicates that the input winapp2.ini should be downloaded from GitHub

## printMenu

### Prints the `Trim` menu to the user

```vb
Public Sub printMenu()
    printMenuTop({"Trim winapp2.ini such that it contains only entries relevant to your machine,", "greatly reducing both application load time and the winapp2.ini file size."})
    print(1, "Run (default)", "Trim winapp2.ini")
    print(5, "Toggle Download", "using the latest winapp2.ini from GitHub as the input file", Not isOffline, True, enStrCond:=DownloadFileToTrim, trailingBlank:=True)
    print(1, "File Chooser (winapp2.ini)", "Change the winapp2.ini name or location", Not DownloadFileToTrim, isOffline, True)
    print(1, "File Chooser (save)", "Change the save file name or location", trailingBlank:=True)
    print(0, $"Current winapp2.ini location: {If(DownloadFileToTrim, GetNameFromDL(DownloadFileToTrim), replDir(TrimFile1.Path))}")
    print(0, $"Current save location: {replDir(TrimFile3.Path)}", closeMenu:=Not ModuleSettingsChanged)
    print(2, "Trim", cond:=ModuleSettingsChanged, closeMenu:=True)
End Sub
```

## handleUserInput

### Handles the user input from the menu

```vb
Public Sub handleUserInput(input As String)
    Select Case True
        Case input = "0"
            exitModule()
        Case (input = "1" Or input = "")
            initTrim()
        Case input = "2" And Not isOffline
            toggleDownload(DownloadFileToTrim, ModuleSettingsChanged)
        Case (input = "3" And Not DownloadFileToTrim And Not isOffline) Or (input = "2" And isOffline)
            changeFileParams(TrimFile1, ModuleSettingsChanged)
        Case (input = "4" And Not DownloadFileToTrim And Not isOffline) Or (input = "3" And (isOffline Or DownloadFileToTrim))
            changeFileParams(TrimFile3, ModuleSettingsChanged)
        Case ModuleSettingsChanged And ((input = "5" And Not DownloadFileToTrim) Or (input = "4" And (isOffline Or DownloadFileToTrim)))
            resetModuleSettings("Trim", AddressOf initDefaultSettings)
        Case Else
            setHeaderText(invInpStr, True)
    End Select
End Sub
```

|Parameter|Type|Description
:-|:-|:-
input|`String`|The `String` containing the user's input

## initTrim

### Initiates the `Trim` process from the main menu or commandline

```vb
Private Sub initTrim()
    If Not DownloadFileToTrim Then If Not enforceFileHasContent(TrimFile1) Then Exit Sub
    Dim winapp2 = If(Not DownloadFileToTrim, New winapp2file(TrimFile1), New winapp2file(getRemoteIniFile(winapp2link)))
    clrConsole()
    print(3, "Trimming... Please wait, this may take a moment...")
    Dim entryCountBeforeTrim = winapp2.count
    trim(winapp2)
    clrConsole()
    print(3, "Finished!")
    clrConsole()
    print(4, "Trim Complete", conjoin:=True)
    print(0, "Entry Count", isCentered:=True, trailingBlank:=True)
    print(0, $"Initial: {entryCountBeforeTrim}")
    print(0, $"Trimmed: {winapp2.count}")
    Dim difference = entryCountBeforeTrim - winapp2.count
    print(0, $"{difference} entries trimmed from winapp2.ini ({Math.Round((difference / entryCountBeforeTrim) * 100)}%)")
    print(0, anyKeyStr, leadingBlank:=True, closeMenu:=True)
    gLog($"{difference} entries trimmed from winapp2.ini ({Math.Round((difference / entryCountBeforeTrim) * 100)}%)")
    gLog($"{winapp2.count} entries remain.")
    TrimFile3.overwriteToFile(winapp2.winapp2string)
    setHeaderText($"{TrimFile3.Name} saved")
    crk()
End Sub
```

## trim

### Trims a `winapp2file`, removing entries not relevant to the current system

```vb
Public Sub trim(winapp2 As winapp2file)
    For Each entryList In winapp2.Winapp2entries
        processEntryList(entryList)
    Next
    winapp2.rebuildToIniFiles()
    winapp2.sortInneriniFiles()
End Sub
```

|Parameter|Type|Description
:-|:-|:-
winapp2|`winapp2file`|The `winapp2file` to be trimmed to fit the current system

## checkExistence

### Evaluates a `keyList` to observe whether they exist on the current machine

```vb
Private Function checkExistence(ByRef kl As keyList, chkExist As Func(Of String, Boolean)) As Boolean
    If kl.KeyCount = 0 Then Return False
    For Each key In kl.Keys
        If chkExist(key.Value) Then
            gLog($"{key.Value} matched a path on the system", Not kl.KeyType = "DetectOS", descend:=True, indent:=True)
            Return True
        End If
    Next
    Return False
End Function
```

|Parameter|Type|Description
:-|:-|:-
kl|`keyList`|The `keyList` containing detection criteria to be evaluated
chkExist|`Func(Of String, Boolean)`|The `function` that evaluates the detection criteria in `kl`

## processEntryExistence

Returns `True` if any detection criteria are met, `False` otherwise

### Audits the detection criteria in a given `winapp2entry` against the current system

```vb
Private Function processEntryExistence(ByRef entry As winapp2entry) As Boolean
    gLog($"Processing entry: {entry.Name}", ascend:=True)
    Dim hasMetDetOS = False
    ' Process the DetectOS if we have one, take note if we meet the criteria, otherwise return false
    If Not entry.DetectOS.KeyCount = 0 Then
        If winVer = Nothing Then winVer = getWinVer()
        hasMetDetOS = checkExistence(entry.DetectOS, AddressOf checkDetOS)
        gLog($"Met DetectOS criteria. {winVer} satisfies {entry.DetectOS.Keys.First.Value}", hasMetDetOS, indent:=True)
        gLog($"Did not meet DetectOS criteria. {winVer} does not satisfy {entry.DetectOS.Keys.First.Value}", Not hasMetDetOS, descend:=True, indent:=True)
        If Not hasMetDetOS Then Return False
    End If
    ' Process any other Detect criteria we have
    If checkExistence(entry.SpecialDetect, AddressOf checkSpecialDetects) Then Return True
    If checkExistence(entry.Detects, AddressOf checkRegExist) Then Return True
    If checkExistence(entry.DetectFiles, AddressOf checkPathExist) Then Return True
    ' Return true for the case where we have only a DetectOS and we meet its criteria
    Dim onlyHasDetOS = entry.SpecialDetect.KeyCount + entry.DetectFiles.KeyCount + entry.Detects.KeyCount = 0
    gLog("No other detection keys found than DetectOS", onlyHasDetOS And hasMetDetOS, descend:=True)
    If hasMetDetOS And onlyHasDetOS Then Return True
    ' Return true for the case where we have no valid detect criteria
    Dim hasNoDetectKeys = entry.DetectOS.KeyCount + entry.DetectFiles.KeyCount + entry.Detects.KeyCount + entry.SpecialDetect.KeyCount = 0
    gLog("No detect keys found, entry will be retained.", hasNoDetectKeys, descend:=True)
    If hasNoDetectKeys Then Return True
    gLog("", descend:=True)
    Return False
End Function
```

|Parameter|Type|Description
:-|:-|:-
entry|`winapp2entry`|A `winapp2entry` to whose detection criteria will be audited

## virtualStoreChecker

### Audits the given entry for legacy codepaths in the machine's VirtualStore

```vb
Private Sub virtualStoreChecker(ByRef entry As winapp2entry)
    vsKeyChecker(entry.FileKeys)
    vsKeyChecker(entry.RegKeys)
    vsKeyChecker(entry.ExcludeKeys)
End Sub
```

|Parameter|Type|Description
:-|:-|:-
entry|`winapp2entry`|The `winapp2entry` to audit

## vsKeyChecker

### Generates keys for VirtualStore locations that exist on the current system and inserts them into the given list

```vb
Private Sub vsKeyChecker(ByRef kl As keyList)
    If kl.KeyCount = 0 Then Exit Sub
    Dim starterCount = kl.KeyCount
    Select Case kl.KeyType
        Case "FileKey", "ExcludeKey"
            mkVsKeys({"%ProgramFiles%", "%CommonAppData%", "%CommonProgramFiles%", "HKLM\Software"}, {"%LocalAppData%\VirtualStore\Program Files*", "%LocalAppData%\VirtualStore\ProgramData", "%LocalAppData%\VirtualStore\Program Files*\Common Files", "HKCU\Software\Classes\VirtualStore\MACHINE\SOFTWARE"}, kl)
        Case "RegKey"
            mkVsKeys({"HKLM\Software"}, {"HKCU\Software\Classes\VirtualStore\MACHINE\SOFTWARE"}, kl)
    End Select
    If Not starterCount = kl.KeyCount Then kl.renumberKeys(replaceAndSort(kl.toStrLst(True), "|", " \ \"))
End Sub
```

|Parameter|Type|Description
:-|:-|:-
kl|`keyList`|The `keyList` of FileKey, RegKey, or ExcludeKeys to be checked against the VirtualStore

## mkVsKeys

### Creates `iniKeys` to handle VirtualStore locations that correspond to paths given in `kl`

```vb
Private Sub mkVsKeys(findStrs As String(), replStrs As String(), ByRef kl As keyList)
    Dim initVals = kl.toStrLst(True)
    Dim keysToAdd As New keyList(kl.KeyType)
    For Each key In kl.Keys
        If Not key.vHasAny(findStrs, True) Then Continue For
        For i = 0 To findStrs.Count - 1
            Dim keyToAdd = createVSKey(findStrs(i), replStrs(i), key)
            ' Don't recreate keys that already exist
            If initVals.contains(keyToAdd.Value) Then Continue For
            keysToAdd.add(keyToAdd, Not key.Value = keyToAdd.Value)
        Next
    Next
    Dim kl2 = kl
    keysToAdd.Keys.ForEach(Sub(key) kl2.add(key, checkExist(New winapp2KeyParameters(key).PathString)))
    kl = kl2
End Sub
```

|Parameter|Type|Description
:-|:-|:-
findStrs|`String()`|An array of Strings to seek for in the key value
replStrs|`String()`|An array of strings to replace the sought after key values
kl|`keyList`|The `keylist` to be processed

## createVSKey

### Creates the VirtualStore version of a given `iniKey`

```vb
Private Function createVSKey(findStr As String, replStr As String, key As iniKey) As iniKey
    Return New iniKey($"{key.Name}={key.Value.Replace(findStr, replStr)}")
End Function
```

|Parameter|Type|Description
:-|:-|:-
findStr|`String`|The normal filesystem path
replStr|`String`|The VirtualStore path
key|`iniKey`|The `iniKey` to processed into a VirtualStore key

## processEntryList

### Processes a list of `winapp2entries` and removes any from the list that wouldn't be detected by CCleaner

```vb
Private Sub processEntryList(ByRef entryList As List(Of winapp2entry))
    Dim sectionsToBePruned As New List(Of winapp2entry)
    ' If the entry's Detect criteria doesn't return true, prune it
    entryList.ForEach(Sub(entry) If Not processEntryExistence(entry) Then sectionsToBePruned.Add(entry) Else virtualStoreChecker(entry))
    removeEntries(entryList, sectionsToBePruned)
End Sub
```

|Parameter|Type|Description
:-|:-|:-
entryList|`List(Of winapp2entry)`|The list of `winapp2entries` who detection criteria will be audited

## checkSpecialDetects

### Returns `True` if a SpecialDetect location exists, `False` otherwise

```vb
Private Function checkSpecialDetects(ByVal key As String) As Boolean
    Select Case key
        Case "DET_CHROME"
            Dim detChrome As New List(Of String) _
                        From {"%AppData%\ChromePlus\chrome.exe", "%LocalAppData%\Chromium\Application\chrome.exe", "%LocalAppData%\Chromium\chrome.exe",
                        "%LocalAppData%\Flock\Application\flock.exe", "%LocalAppData%\Google\Chrome SxS\Application\chrome.exe", "%LocalAppData%\Google\Chrome\Application\chrome.exe",
                        "%LocalAppData%\RockMelt\Application\rockmelt.exe", "%LocalAppData%\SRWare Iron\iron.exe", "%ProgramFiles%\Chromium\Application\chrome.exe",
                        "%ProgramFiles%\SRWare Iron\iron.exe", "%ProgramFiles%\Chromium\chrome.exe", "%ProgramFiles%\Flock\Application\flock.exe",
                        "%ProgramFiles%\Google\Chrome SxS\Application\chrome.exe", "%ProgramFiles%\Google\Chrome\Application\chrome.exe", "%ProgramFiles%\RockMelt\Application\rockmelt.exe",
                        "HKCU\Software\Chromium", "HKCU\Software\SuperBird", "HKCU\Software\Torch", "HKCU\Software\Vivaldi"}
            For Each path As String In detChrome
                If checkExist(path) Then Return True
            Next
        Case "DET_MOZILLA"
            Return checkPathExist("%AppData%\Mozilla\Firefox")
        Case "DET_THUNDERBIRD"
            Return checkPathExist("%AppData%\Thunderbird")
        Case "DET_OPERA"
            Return checkPathExist("%AppData%\Opera Software")
    End Select
    ' If we didn't return above, SpecialDetect definitely doesn't exist
    Return False
End Function
```

|Parameter|Type|Description
:-|:-|:-
key|`iniKey`|A SpecialDetect format `iniKey`

## checkExist

### Handles passing off checks from sources that may vary between file system and registry

```vb
Private Function checkExist(path As String) As Boolean
    Return If(path.StartsWith("HK"), checkRegExist(path), checkPathExist(path))
End Function
```

|Parameter|Type|Description
:-|:-|:-
path|`String`|A filesystem or registry path whose existence will be audited

## checkRegExist

### Returns `True` if a given key exists in the Windows Registry, `False` otherwise

```vb
Private Function checkRegExist(path As String) As Boolean
    Dim dir = path
    Dim root = getFirstDir(path)
    dir = dir.Replace(root & "\", "")
    Dim exists = getRegExists(root, dir)
    ' If we didn't return anything above, registry location probably doesn't exist
    Return exists
End Function
```

|Parameter|Type|Description
:-|:-|:-
root|`String`|The registry hive that contains the key whose existence will be audited
dir|`String`|The path of the key whose existence will be audited

## processEnvDirs

### Handles some CCleaner variables and logs if the current variable is ProgramFiles so the 32bit location can be checked later

```vb
Private Sub processEnvDirs(ByRef dir As String, ByRef isProgramFiles As Boolean)
    If dir.Contains("%") Then
        Dim splitDir = dir.Split(CChar("%"))
        Dim var = splitDir(1)
        Dim envDir = Environment.GetEnvironmentVariable(var)
        Select Case var
            Case "ProgramFiles"
                isProgramFiles = True
            Case "Documents"
                envDir = $"{Environment.GetEnvironmentVariable("UserProfile")}\{If(winVer = 5.1, "My ", "")}Documents"
            Case "CommonAppData"
                envDir = Environment.GetEnvironmentVariable("ProgramData")
        End Select
        dir = envDir + splitDir(2)
    End If
End Sub
```

|Parameter|Type|Description
:-|:-|:-
dir|`String`|A filesystem path to process for environment variables
isProgramFiles|`Boolean`|Indicates that the %ProgramFiles% variable has been seen

## checkPathExist

### Returns `True` if a path exists on the file system, `False` otherwise

```vb
Private Function checkPathExist(key As String) As Boolean
    Dim isProgramFiles = False
    Dim dir = key
    ' Make sure we get the proper path for environment variables
    processEnvDirs(dir, isProgramFiles)
    Try
        ' Process wildcards appropriately if we have them
        If dir.Contains("*") Then
            Dim exists = expandWildcard(dir, True)
            ' Small contingency for the isProgramFiles case
            If Not exists And isProgramFiles Then
                swapDir(dir, key)
                exists = expandWildcard(dir, True)
            End If
            Return exists
        End If
        ' Check out those file/folder paths
        If Directory.Exists(dir) Or File.Exists(dir) Then Return True
        ' If we didn't find it and we're looking in Program Files, check the (x86) directory
        If isProgramFiles Then
            swapDir(dir, key)
            Dim exists = Directory.Exists(dir) Or File.Exists(dir)
            Return exists
        End If
    Catch ex As Exception
        exc(ex)
        Return True
    End Try
    Return False
End Function
```

|Parameter|Type|Description
:-|:-|:-
key|`String`|A filesystem path

## swapDir

### Swaps out a directory with the ProgramFiles parameterization on 64bit computers

```vb
Private Sub swapDir(ByRef dir As String, key As String)
    Dim envDir = Environment.GetEnvironmentVariable("ProgramFiles(x86)")
    dir = envDir & key.Split(CChar("%"))(2)
End Sub
```

|Parameter|Type|Description
:-|:-|:-
dir|`String`|The file system path to be modified
key|`String`|The original state of the path

## expandWildcard

### Interprets parameterized wildcards for the current system

```vb
Private Function expandWildcard(dir As String, isFileSystem As Boolean) As Boolean
    ' This should handle wildcards anywhere in a path even though CCleaner only supports them at the end for DetectFiles
    Dim possibleDirs As New strList
    Dim currentPaths As New strList
    currentPaths.add("")
    ' Split the given string into sections by directory
    Dim splitDir = dir.Split(CChar("\"))
    For Each pathPart In splitDir
        ' If this directory parameterization includes a wildcard, expand it appropriately
        ' This probably wont work if a string for some reason starts with a *
        If pathPart.Contains("*") Then
            For Each currentPath In currentPaths.Items
                Try
                    ' Query the existence of child paths for each current path we hold
                    If isFileSystem Then
                        Dim possibilities = Directory.GetDirectories(currentPath, pathPart)
                        ' If there are any, add them to our possibility list
                        possibleDirs.add(possibilities, Not possibilities.Count = 0)
                    Else
                        ' Registry Query here
                    End If
                Catch
                    ' The exception we encounter here is going to be the result of directories not existing.
                    ' The exception will be thrown from the GetDirectories call and will prevent us from attempting to add new
                    ' items to the possibility list. In this instance, we can silently fail (here).
                End Try
            Next
            ' If no possibilities remain, the wildcard parameterization hasn't left us with any real paths on the system, so we may return false.
            If possibleDirs.Count = 0 Then Return False
            ' Otherwise, clear the current paths and repopulate them with the possible paths
            currentPaths.clear()
            currentPaths.add(possibleDirs)
            possibleDirs.clear()
        Else
            If currentPaths.Count = 0 Then
                currentPaths.add($"{pathPart}")
            Else
                Dim newCurPaths As New strList
                For Each path In currentPaths.Items
                    If Not path.EndsWith("\") And path <> "" Then path += "\"
                    newCurPaths.add($"{path}{pathPart}\", Directory.Exists($"{path}{pathPart}\"))
                Next
                currentPaths = newCurPaths
            End If
        End If
    Next
    ' If any file/path exists, return true
    For Each currDir In currentPaths.Items
        If Directory.Exists(currDir) Or File.Exists(currDir) Then Return True
    Next
    Return False
End Function
```

|Parameter|Type|Description
:-|:-|:-
dir|`String`|A path containing a wildcard

## checkDetOS

### Returns `True` if the system satisfies the DetectOS citeria, `False` otherwise
