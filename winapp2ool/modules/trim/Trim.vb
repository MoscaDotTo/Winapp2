'    Copyright (C) 2018-2019 Robbie Ward
' 
'    This file is a part of Winapp2ool
' 
'    Winapp2ool is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    Winap2ool is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with Winapp2ool.  If not, see <http://www.gnu.org/licenses/>.
Option Strict On
Imports System.IO
'''    <summary>
'''    This module parses a winapp2.ini file and checks each entry therein 
'''    removing any whose detection parameters do not exist on the current system
'''    and outputting a "trimmed" file containing only entries that exist on the system
'''    to the user.
'''   </summary>
Public Module Trim
    '''<summary>Indicates the major/minor version number on the current system</summary>
    Private Property winVer As Double
    '''<summary>Hardcoded list of SpecialDetect evaluations</summary>
    Private ReadOnly Property detChrome As New List(Of String) _
        From {"%AppData%\ChromePlus\chrome.exe", "%LocalAppData%\Chromium\Application\chrome.exe", "%LocalAppData%\Chromium\chrome.exe", "%LocalAppData%\Flock\Application\flock.exe", "%LocalAppData%\Google\Chrome SxS\Application\chrome.exe",
                           "%LocalAppData%\Google\Chrome\Application\chrome.exe", "%LocalAppData%\RockMelt\Application\rockmelt.exe", "%LocalAppData%\SRWare Iron\iron.exe", "%ProgramFiles%\Chromium\Application\chrome.exe", "%ProgramFiles%\SRWare Iron\iron.exe",
                           "%ProgramFiles%\Chromium\chrome.exe", "%ProgramFiles%\Flock\Application\flock.exe", "%ProgramFiles%\Google\Chrome SxS\Application\chrome.exe", "%ProgramFiles%\Google\Chrome\Application\chrome.exe", "%ProgramFiles%\RockMelt\Application\rockmelt.exe",
                           "HKCU\Software\Chromium", "HKCU\Software\SuperBird", "HKCU\Software\Torch", "HKCU\Software\Vivaldi"}
    '''<summary>Indicates module settings have changed and can be reset</summary>
    Private Property ModuleSettingsChanged As Boolean = False
    '''<summary>Indicates that we are downloading a winapp2.ini from GitHub</summary>
    Private Property DownloadFileToTrim As Boolean = False
    '''<summary>The winapp2.ini file that will be trimmed</summary>
    Public Property TrimFile1 As New iniFile(Environment.CurrentDirectory, "winapp2.ini", mExist:=True)
    '''<summary>Holds the path where the output file will be saved to disk. Overwrites the input file by default.</summary>
    Public Property TrimFile3 As New iniFile(Environment.CurrentDirectory, "winapp2.ini", "winapp2-trimmed.ini")

    ''' <summary>Handles the commandline args for Trim</summary>
    ''' Trim args:
    ''' -d          : download the latest winapp2.ini
    Public Sub handleCmdLine()
        initDefaultSettings()
        handleDownloadBools(DownloadFileToTrim)
        getFileAndDirParams(TrimFile1, TrimFile3, New iniFile)
        initTrim()
    End Sub

    ''' <summary>Restores the default state of the module's parameters</summary>
    Private Sub initDefaultSettings()
        TrimFile1.resetParams()
        TrimFile3.resetParams()
        DownloadFileToTrim = False
        ModuleSettingsChanged = False
    End Sub

    ''' <summary>Runs the trimmer from outside the module</summary>
    ''' <param name="firstFile">The winapp2.ini file</param>
    ''' <param name="secondFile">The output file</param>
    ''' <param name="d">Download boolean</param>
    Public Sub remoteTrim(firstFile As iniFile, secondFile As iniFile, d As Boolean)
        TrimFile1 = firstFile
        TrimFile3 = secondFile
        DownloadFileToTrim = d
        initTrim()
    End Sub

    ''' <summary>Prints the main menu to the user</summary>
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

    ''' <summary>Handles the user input from the menu</summary>
    ''' <param name="input">The String containing the user's input</param>
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

    ''' <summary>Initiates the trim after validating our ini files</summary>
    Private Sub initTrim()
        If Not DownloadFileToTrim Then
            TrimFile1.validate()
            If Not enforceFileHasContent(TrimFile1) Then Exit Sub
        End If
        Dim winapp2 = If(Not DownloadFileToTrim, New winapp2file(TrimFile1), New winapp2file(getRemoteIniFile(winapp2link)))
        clrConsole()
        print(3, "Trimming... Please wait, this may take a moment...")
        Dim entryCountBeforeTrim = winapp2.count
        trim(winapp2)
        clrConsole()
        print(0, tmenu("Finished!"), closeMenu:=True)
        clrConsole()
        print(0, tmenu("Trim Complete"))
        print(0, menuStr03)
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
        If Not SuppressOutput Then Console.ReadKey()
    End Sub

    ''' <summary>Performs the trim</summary>
    ''' <param name="winapp2">A winapp2.ini file</param>
    Public Sub trim(winapp2 As winapp2file)
        For Each entryList In winapp2.Winapp2entries
            processEntryList(entryList)
        Next
        winapp2.rebuildToIniFiles()
        winapp2.sortInneriniFiles()
    End Sub

    ''' <summary>Evaluates a list of keys to observe whether they exist on the current machine</summary>
    ''' <param name="kl">The list of iniKeys to query</param>
    ''' <param name="chkExist">The function that evaluates that keyType's parameters</param>
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

    ''' <summary>Returns true if an entry's detection criteria is matched by the system, false otherwise.</summary>
    ''' <param name="entry">A winapp2.ini entry</param>
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

    ''' <summary>Audits the given entry for legacy codepaths in the machine's VirtualStore locations</summary>
    ''' <param name="entry"></param>
    Private Sub virtualStoreChecker(ByRef entry As winapp2entry)
        vsKeyChecker(entry.FileKeys)
        vsKeyChecker(entry.RegKeys)
        vsKeyChecker(entry.ExcludeKeys)
    End Sub

    ''' <summary>Generates keys for VirtualStore locations that exist on the current system and inserts them into the given list</summary>
    ''' <param name="kl">The keylist of FileKey, RegKey, or ExcludeKeys to be checked against the VirtualStore</param>
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

    ''' <summary>Audits the existence of VirtualStore locations for an iniKey and if they exist, adds them to the list</summary>
    ''' <param name="findStrs">A list of strings to seek for in the key value</param>
    ''' <param name="replStrs">A list of strings to replace the sought after key values</param>
    ''' <param name="kl">The keylist to be processed</param>
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

    ''' <summary>Creates the VirtualStore version of a given iniKey</summary>
    ''' <param name="findStr">The normal filesystem path</param>
    ''' <param name="replStr">The VirtualStore location path</param>
    ''' <param name="key">The key to processed</param>
    ''' <returns></returns>
    Private Function createVSKey(findStr As String, replStr As String, key As iniKey) As iniKey
        Return New iniKey($"{key.Name}={key.Value.Replace(findStr, replStr)}")
    End Function

    ''' <summary>Processess a list of winapp2.ini entries and removes any from the list that wouldn't be detected by CCleaner</summary>
    ''' <param name="entryList">The list of winapp2entry objects to check existence for</param>
    Private Sub processEntryList(ByRef entryList As List(Of winapp2entry))
        Dim sectionsToBePruned As New List(Of winapp2entry)
        ' If the entry's Detect criteria doesn't return true, prune it 
        entryList.ForEach(Sub(entry) If Not processEntryExistence(entry) Then sectionsToBePruned.Add(entry) Else virtualStoreChecker(entry))
        removeEntries(entryList, sectionsToBePruned)
    End Sub

    ''' <summary>Returns true if a SpecialDetect location exist, false otherwise</summary>
    ''' <param name="key">A SpecialDetect format iniKey</param>
    Private Function checkSpecialDetects(ByVal key As String) As Boolean
        Select Case key
            Case "DET_CHROME"
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

    ''' <summary>Handles passing off checks from sources that may vary between file system and registry</summary>
    ''' <param name="path">A filesystem or registry path to be audited for existence</param>
    Private Function checkExist(path As String) As Boolean
        Return If(path.StartsWith("HK"), checkRegExist(path), checkPathExist(path))
    End Function

    ''' <summary>Returns True if a given Detect path exists in the system registry, false otherwise.</summary>
    ''' <param name="path">A registry path from a Detect key</param>
    Private Function checkRegExist(path As String) As Boolean
        Dim dir = path
        Dim root = getFirstDir(path)
        dir = dir.Replace(root & "\", "")
        Dim exists = getRegExists(root, dir)
        ' If we didn't return anything above, registry location probably doesn't exist
        Return exists
    End Function

    Private Function getRegExists(root As String, dir As String) As Boolean
        Try
            Select Case root
                Case "HKCU"
                    Return getCUKey(dir) IsNot Nothing
                Case "HKLM"
                    If getLMKey(dir) IsNot Nothing Then Return True
                    ' Support checking for 32bit applications on Win64
                    dir = dir.ToLower.Replace("software\", "Software\WOW6432Node\")
                    Return getLMKey(dir) IsNot Nothing
                Case "HKU"
                    Return getUserKey(dir) IsNot Nothing
                Case "HKCR"
                    Return getCRKey(dir) IsNot Nothing
            End Select
        Catch ex As Exception
            ' The most common (only?) exception here is a permissions one, so assume true if we hit because a permissions exception implies the key exists anyway.
            gLog(ex.Message)
            Return True
        End Try
        Return True
    End Function

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

    ''' <summary>Returns True if a path exists on the system, false otherwise.</summary>
    ''' <param name="key">A filesystem path</param>
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

    ''' <summary>Swaps out a directory with the ProgramFiles parameterization on 64bit computers</summary>
    ''' <param name="dir">The text to be modified</param>
    ''' <param name="key">The original state of the text</param>
    Private Sub swapDir(ByRef dir As String, key As String)
        Dim envDir = Environment.GetEnvironmentVariable("ProgramFiles(x86)")
        dir = envDir & key.Split(CChar("%"))(2)
    End Sub

    ''' <summary>Interpret parameterized wildcards for the current system</summary>
    ''' <param name="dir">A path containing a wildcard</param>
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
                            If Not possibilities.Count = 0 Then possibleDirs.add(possibilities)
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

    ''' <summary>Returns true if we satisfy the DetectOS citeria, false otherwise</summary>
    ''' <param name="value">The DetectOS criteria to be checked</param>
    Private Function checkDetOS(value As String) As Boolean
        Dim splitKey = value.Split(CChar("|"))
        Select Case True
            Case value.StartsWith("|")
                Return Not winVer > Val(splitKey(1))
            Case value.EndsWith("|")
                Return Not winVer < Val(splitKey(0))
            Case Else
                Return winVer >= Val(splitKey(0)) And winVer <= Val(splitKey(1))
        End Select
    End Function
End Module