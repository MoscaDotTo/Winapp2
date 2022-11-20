'    Copyright (C) 2018-2022 Hazel Ward
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
Imports System.IO
'''    <summary>
'''    This module parses a winapp2.ini file and checks each entry therein
'''    removing any whose detection parameters do not exist on the current system
'''    and outputting a "trimmed" file containing only entries that exist on the system
'''    to the user.
'''   </summary>
Public Module Trim
    ''' <summary> The winapp2.ini file that will be trimmed </summary>
    Public Property TrimFile1 As New iniFile(Environment.CurrentDirectory, "winapp2.ini", mExist:=True)
    ''' <summary> Holds the path of an iniFile containing the names of Sections who should never be trimmed </summary>
    Public Property TrimFile2 As New iniFile(Environment.CurrentDirectory, "includes.ini")
    ''' <summary> Holds the path where the output file will be saved to disk. Overwrites the input file by default </summary>
    Public Property TrimFile3 As New iniFile(Environment.CurrentDirectory, "winapp2.ini", "winapp2-trimmed.ini")
    ''' <summary> Holds the path of an iniFile containing the names of Sections who should always be trimmed </summary>
    Public Property TrimFile4 As New iniFile(Environment.CurrentDirectory, "excludes.ini")
    ''' <summary> The major/minor version number on the current system </summary>
    Private Property winVer As Double
    ''' <summary> Indicates that the module settings have been modified from their defaults </summary>
    Private Property ModuleSettingsChanged As Boolean = False
    ''' <summary> Indicates that we are downloading a winapp2.ini from GitHub </summary>
    Private Property DownloadFileToTrim As Boolean = False
    ''' <summary> Indicates that the includes should be consulted while trimming </summary>
    Private Property UseIncludes As Boolean = False
    ''' <summary> Indicates that the excludes should be consulted while trimming </summary>
    Private Property useExcludes As Boolean = False

    ''' <summary> Handles the commandline args for Trim </summary>
    ''' Trim args:
    ''' -d          : download the latest winapp2.ini
    Public Sub handleCmdLine()

        initDefaultSettings()
        handleDownloadBools(DownloadFileToTrim)
        getFileAndDirParams(TrimFile1, New iniFile, TrimFile3)
        initTrim()

    End Sub

    ''' <summary> Restores the default state of the module's parameters </summary>
    Private Sub initDefaultSettings()

        TrimFile1.resetParams()
        TrimFile2.resetParams()
        TrimFile3.resetParams()
        TrimFile4.resetParams()
        DownloadFileToTrim = False
        ModuleSettingsChanged = False
        UseIncludes = False
        useExcludes = False
        restoreDefaultSettings(NameOf(Trim), AddressOf createTrimSettingsSection)

    End Sub

    ''' <summary> Loads values from disk into memory for the Trim module settings </summary>
    Public Sub getSerializedTrimSettings()

        For Each kvp In settingsDict(NameOf(Trim))

            Select Case kvp.Key

                Case NameOf(TrimFile1) & "_Name"

                    TrimFile1.Name = kvp.Value

                Case NameOf(TrimFile1) & "_Dir"

                    TrimFile1.Dir = kvp.Value

                Case NameOf(TrimFile2) & "_Name"

                    TrimFile2.Name = kvp.Value

                Case NameOf(TrimFile2) & "_Dir"

                    TrimFile2.Dir = kvp.Value

                Case NameOf(TrimFile3) & "_Name"

                    TrimFile3.Name = kvp.Value

                Case NameOf(TrimFile3) & "_Dir"

                    TrimFile3.Dir = kvp.Value

                Case NameOf(TrimFile4) & "_Name"

                    TrimFile4.Name = kvp.Value

                Case NameOf(TrimFile4) & "_Dir"

                    TrimFile4.Dir = kvp.Value

                Case NameOf(DownloadFileToTrim)

                    DownloadFileToTrim = CBool(kvp.Value)

                Case NameOf(ModuleSettingsChanged)

                    ModuleSettingsChanged = CBool(kvp.Value)

                Case NameOf(UseIncludes)

                    UseIncludes = CBool(kvp.Value)

                Case NameOf(useExcludes)

                    useExcludes = CBool(kvp.Value)

            End Select

        Next

    End Sub

    ''' <summary> Adds the current (typically default) state of the module's settings into the disk-writable settings representation </summary>
    Public Sub createTrimSettingsSection()

        Dim trimSettingsTuple As New List(Of String) From {NameOf(DownloadFileToTrim), tsInvariant(DownloadFileToTrim), NameOf(UseIncludes), tsInvariant(UseIncludes),
        NameOf(useExcludes), tsInvariant(useExcludes), NameOf(ModuleSettingsChanged), tsInvariant(ModuleSettingsChanged), NameOf(TrimFile1), TrimFile1.Name, TrimFile1.Dir,
        NameOf(TrimFile2), TrimFile2.Name, TrimFile2.Dir, NameOf(TrimFile3), TrimFile3.Name, TrimFile3.Dir, NameOf(TrimFile4), TrimFile4.Name, TrimFile4.Dir}
        createModuleSettingsSection(NameOf(Trim), trimSettingsTuple, 4, 4)

    End Sub

    ''' <summary> Trims an <c> iniFile </c> from outside the module </summary>
    ''' <param name="firstFile"> The winapp2.ini file to be trimmed </param>
    ''' <param name="thirdFile"> <c> iniFile </c> containing the path on disk to which the trimmed file will be saved </param>
    ''' <param name="d"> Indicates that the input winapp2.ini should be downloaded from GitHub </param>
    Public Sub remoteTrim(firstFile As iniFile, thirdFile As iniFile, d As Boolean)

        TrimFile1 = firstFile
        TrimFile3 = thirdFile
        DownloadFileToTrim = d
        initTrim()

    End Sub

    ''' <summary> Prints the <c> Trim </c> menu to the user </summary>
    Public Sub printMenu()

        If isOffline Then DownloadFileToTrim = False

        printMenuTop({"Trim winapp2.ini such that it contains only entries relevant to your machine,", "greatly reducing both application load time and the winapp2.ini file size."})
        print(1, "Run (default)", "Trim winapp2.ini")
        print(5, "Toggle Download", "using the latest winapp2.ini from GitHub as the input file", Not isOffline, True, enStrCond:=DownloadFileToTrim, trailingBlank:=True)
        print(1, "File Chooser (winapp2.ini)", "Configure the path to winapp2.ini ", Not DownloadFileToTrim, isOffline, True)
        print(1, "File Chooser (save)", "Cofigure the path to which the trimmed winapp2.ini will be saved", trailingBlank:=True)
        print(5, "Toggle Includes", "always keeping certain entries", enStrCond:=UseIncludes, trailingBlank:=Not UseIncludes)
        print(1, "File Chooser (Includes)", "Configure the path to the includes file", cond:=UseIncludes, trailingBlank:=True)
        print(5, "Toggle Excludes", "always discarding certain entries", enStrCond:=useExcludes, trailingBlank:=Not useExcludes)
        print(1, "File Chooser (Excludes)", "Configure the path to the excludes file", cond:=useExcludes, trailingBlank:=True)
        print(0, $"Current winapp2.ini path: {If(DownloadFileToTrim, GetNameFromDL(DownloadFileToTrim), replDir(TrimFile1.Path))}")
        print(0, $"Current save path: {replDir(TrimFile3.Path)}", closeMenu:=Not (UseIncludes Or useExcludes Or ModuleSettingsChanged))
        print(0, $"Current includes path: {replDir(TrimFile2.Path)}", cond:=UseIncludes, closeMenu:=Not (useExcludes Or ModuleSettingsChanged))
        print(0, $"Current excludes path: {replDir(TrimFile4.Path)}", cond:=useExcludes, closeMenu:=Not ModuleSettingsChanged)
        print(2, NameOf(Trim), cond:=ModuleSettingsChanged, closeMenu:=True)

    End Sub

    ''' <summary> Handles the user input from the menu </summary>
    ''' <param name="input"> The String containing the user's input </param>
    Public Sub handleUserInput(input As String)

        If input Is Nothing Then argIsNull(NameOf(input)) : Return

        ' At least one Include/Exclude option is toggled
        Dim IncOrExcl = UseIncludes Or useExcludes
        ' Both the Include and the Exclude options are toggled, applies +2 to all settings numbers following the excludes option in the menu 
        Dim IncAndExcl = UseIncludes And useExcludes
        ' One and only one of the Include/Exclude options are toggled, applies +1 to all settings after the enabled option 
        Dim IncXorExcl = UseIncludes Xor useExcludes

        Select Case True

            ' Option Name:                                 Exit 
            ' Option States:
            ' Default                                      -> 0 (Default) 
            Case input = "0"

                exitModule()

            ' Option Name:                                 Run (default) 
            ' Option States:
            ' Default                                      -> 1 (Default)
            Case (input = "1" OrElse input.Length = 0)

                initTrim()

            ' Option Name:                                  Toggle Download 
            ' Option States: 
            ' Offline                                      -> Unavailable (not displayed) (offsets all following menu options by -1 from their defaults)
            ' Online                                       -> 2 (default) 
            Case input = "2" AndAlso Not isOffline

                If Not denySettingOffline() Then toggleSettingParam(DownloadFileToTrim, "Downloading", ModuleSettingsChanged, NameOf(Trim), NameOf(DownloadFileToTrim), NameOf(ModuleSettingsChanged))

            ' Option Name:                                 File Chooser (winapp2.ini) 
            ' Option states: 
            ' Downloading                                  -> Unavailable (not displayed) (offsets all following menu options by -1 from their default ) 
            ' Offline (-1)                                 -> 2 (default offline)
            ' Online (-1)                                  -> 3 (default online) 
            Case (
                    (input = "3" AndAlso Not DownloadFileToTrim AndAlso Not isOffline) OrElse
                    (input = "2" AndAlso isOffline))

                changeFileParams(TrimFile1, ModuleSettingsChanged, NameOf(Trim), NameOf(TrimFile1), NameOf(ModuleSettingsChanged))

            ' Option Name:                                 File Chooser (save) 
            ' Option states:
            ' Offline (-1) or Downloading (-1)             -> 3 (default offline) [these two settings are mutually exclusive]
            ' Online, Not downloading                      -> 4 (default online) 
            Case (Not isOffline AndAlso (
                    (input = "4" AndAlso Not DownloadFileToTrim)) OrElse
                    (input = "3" AndAlso (DownloadFileToTrim OrElse isOffline)))

                changeFileParams(TrimFile3, ModuleSettingsChanged, NameOf(Trim), NameOf(TrimFile3), NameOf(ModuleSettingsChanged))

            ' Reset Settings menu option (online case) 
            ' Option States: 
            ' ModuleSettingsChanged = False                 -> Unavailable (not displayed) 
            ' Offline (-1), IncOrExcl = False               -> 6 (default offline)
            ' Downloading (-1), and IncOrExcl = False       -> 6 
            ' Not Downloading, and IncOrExcl = False        -> 7 (default online) 
            ' Downloading (-1), and IncXorExcl (+1) = True  -> 7
            ' Offline (-1), IncXorExclude = True            -> 7 
            ' Not Downloading, and IncXorExcl (+1) = True   -> 8 
            ' Downloading (-1), and IncAndExcl = True (+2)  -> 8
            ' Offline (-1), IncAndExcl (+2) = True          -> 8               
            ' Not Downloading, and IncAndExcl (+2) = True   -> 9 
            Case ModuleSettingsChanged AndAlso (
                    (Not isOffline AndAlso (
                    (input = "6" AndAlso DownloadFileToTrim And Not IncOrExcl) OrElse
                    (input = "7" AndAlso ((Not DownloadFileToTrim AndAlso Not IncOrExcl) OrElse (DownloadFileToTrim AndAlso IncXorExcl))) OrElse
                    (input = "8" AndAlso ((Not DownloadFileToTrim AndAlso IncXorExcl) OrElse (DownloadFileToTrim AndAlso IncAndExcl))) OrElse
                    (input = "9" AndAlso Not DownloadFileToTrim AndAlso IncAndExcl))) _
                    Or (isOffline AndAlso (
                    (input = "6" AndAlso Not IncOrExcl) OrElse
                    (input = "7" AndAlso IncXorExcl) OrElse
                    (input = "8" AndAlso IncAndExcl))))

                resetModuleSettings(NameOf(Trim), AddressOf initDefaultSettings)

            ' Option Name:                                 Toggle Includes 
            ' Option states: 
            ' Downloading                                  -> 4
            ' Offline (-1) or Downloading (-1)             -> 4 [these two settings are mutually exclusive]
            ' Online, Not Downloading                      -> 5 (default) 
            Case Not isOffline AndAlso (
                    (input = "4" AndAlso DownloadFileToTrim) OrElse
                    (input = "5" AndAlso Not DownloadFileToTrim)) OrElse
                    (input = "4" AndAlso isOffline)

                toggleSettingParam(UseIncludes, "Includes", ModuleSettingsChanged, NameOf(Trim), NameOf(UseIncludes), NameOf(ModuleSettingsChanged))

            ' Option Name:                                 File Chooser (Includes) 
            ' Option states: 
            ' Downloading (-1)                             -> 5 
            ' Online, Not Downloading                      -> 6 (default) 
            ' Offline                                      -> 5 (default offline) 
            Case UseIncludes AndAlso Not isOffline AndAlso (
                    (input = "5" AndAlso DownloadFileToTrim) OrElse
                    (input = "6" AndAlso Not DownloadFileToTrim)) OrElse
                    (input = "5" AndAlso isOffline)

                changeFileParams(TrimFile2, ModuleSettingsChanged, NameOf(Trim), NameOf(TrimFile1), NameOf(ModuleSettingsChanged))

            ' Option Name:                                 Toggle Excludes 
            ' Option states: 
            ' Offline (-1), Not Including                  -> 5 (default offline)
            ' Online, Downloading (-1), Not including      -> 5
            ' Offline (-1), Including (+1)                 -> 6 
            ' Online, Not Downloading, Not Inlcuding       -> 6 (default online)  
            ' Online, Downloading (-1), Including (+1)     -> 6 
            ' Online, Not Downloading, Including           -> 7 
            Case (Not isOffline AndAlso (
                    (input = "7" AndAlso Not DownloadFileToTrim AndAlso UseIncludes) OrElse
                    (input = "6" AndAlso ((Not DownloadFileToTrim AndAlso Not UseIncludes) OrElse DownloadFileToTrim AndAlso UseIncludes) Or
                    (input = "5" AndAlso DownloadFileToTrim AndAlso Not UseIncludes))) OrElse
                    (isOffline AndAlso (
                    (input = "5" AndAlso Not UseIncludes) OrElse
                    (input = "6" AndAlso UseIncludes))))

                toggleSettingParam(useExcludes, "Excludes", ModuleSettingsChanged, NameOf(Trim), NameOf(useExcludes), NameOf(ModuleSettingsChanged))

            ' Option Name:                                 File Chooser (Exclude) 
            ' Option States:
            ' Online, Downloading (-1), Not Including      -> 6
            ' Offline (-1), Not including                  -> 6 (default offline)
            ' Online, Not Downloading, Not Including       -> 7 (default online)
            ' Online, Downloading (-1), Including (+1)     -> 7
            ' Offline (-1), Including (+1)                 -> 7
            ' Online, Not Downloading, Including (+1)      -> 8
            Case useExcludes AndAlso ((Not isOffline AndAlso (
                    (input = "6" AndAlso DownloadFileToTrim AndAlso Not UseIncludes) OrElse
                    (input = "7" AndAlso ((Not DownloadFileToTrim AndAlso Not UseIncludes) OrElse (DownloadFileToTrim AndAlso UseIncludes))) OrElse
                    (input = "8" AndAlso Not DownloadFileToTrim AndAlso UseIncludes))) OrElse
                    (isOffline AndAlso
                    (input = "6" AndAlso Not UseIncludes) OrElse
                    (input = "7" AndAlso UseIncludes)))

                changeFileParams(TrimFile4, ModuleSettingsChanged, NameOf(Trim), NameOf(TrimFile4), NameOf(ModuleSettingsChanged))

            Case Else

                setHeaderText(invInpStr, True)

        End Select

    End Sub

    ''' <summary> Initiates the <c> Trim </c> process from the main menu or commandline </summary>
    Private Sub initTrim()

        ' Don't try to trim an empty file 
        If Not DownloadFileToTrim Then If Not enforceFileHasContent(TrimFile1) Then Return

        ' Ensure we have an online connection before continuing if necessary 
        If DownloadFileToTrim Then If Not checkOnline() Then setHeaderText("Internet connection lost! Please check your network connection and try again", True) : Return

        Dim winapp2 = If(DownloadFileToTrim, New winapp2file(getRemoteIniFile(winapp2link)), New winapp2file(TrimFile1))

        ' Spin up our include/excludes 
        If UseIncludes Then TrimFile2.validate()
        If useExcludes Then TrimFile4.validate()

        clrConsole()
        print(3, "Trimming... Please wait, this may take a moment...")

        Dim entryCountBeforeTrim = winapp2.count

        ' Perform the trim 
        trimFile(winapp2)

        ' Print trim summary to the user 
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

        ' Save the trimmed file back to disk 
        TrimFile3.overwriteToFile(winapp2.winapp2string)

        ' If we downloaded the latest file, then we probably can mark winapp2 as having been updated 
        If DownloadFileToTrim Then waUpdateIsAvail = False
        setHeaderText($"{TrimFile3.Name} saved")

        crk()

    End Sub

    ''' <summary> Trims a <c> winapp2file </c>, removing entries not relevant to the current system </summary>
    ''' <param name="winapp2"> A <c> winapp2file </c> to be trimmed to fit the current system </param>
    Public Sub trimFile(winapp2 As winapp2file)

        If winapp2 Is Nothing Then argIsNull(NameOf(winapp2)) : Return

        For i = 0 To winapp2.Winapp2entries.Count - 1
            Dim entryList = winapp2.Winapp2entries(i)
            processEntryList(entryList)
        Next

        winapp2.rebuildToIniFiles()
        winapp2.sortInneriniFiles()

    End Sub

    ''' <summary> Evaluates a <c> keyList </c> to observe whether they exist on the current machine </summary>
    ''' <param name="kl"> The <c> keyList </c> containing detection criteria to be evaluated </param>
    ''' <param name="chkExist"> The <c> function </c> that evaluates the detection criteria in <c> <paramref name="kl"/> </c> </param>
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

    ''' <summary> Audits the detection criteria in a given <c> winapp2entry </c> against the current system <br/> <br/>
    ''' Returns <c> True </c> if the detection criteria are met, <c> False </c> otherwise </summary>
    ''' <param name="entry"> A <c> winapp2entry </c> to whose detection criteria will be audited </param>
    Private Function processEntryExistence(ByRef entry As winapp2entry) As Boolean

        gLog($"Processing entry: {entry.Name}", ascend:=True)

        ' Respect the include/excludes 
        If UseIncludes And TrimFile2.hasSection(entry.Name) Then Return True
        If useExcludes And TrimFile4.hasSection(entry.Name) Then Return False


        ' Process the DetectOS if we have one, take note if we meet the criteria, otherwise return false
        Dim hasMetDetOS = False

        If Not entry.DetectOS.KeyCount = 0 Then

            If winVer = Nothing Then winVer = getWinVer()

            hasMetDetOS = checkExistence(entry.DetectOS, AddressOf checkDetOS)
            gLog($"Met DetectOS criteria. {winVer} satisfies {entry.DetectOS.Keys.First.Value}", hasMetDetOS, indent:=True)
            gLog($"Did not meet DetectOS criteria. {winVer} does not satisfy {entry.DetectOS.Keys.First.Value}", Not hasMetDetOS, descend:=True, indent:=True)

            If Not hasMetDetOS Then Return False

        End If

        ' Process any other Detect criteria we have
        If checkExistence(entry.Detects, AddressOf checkRegExist) Then Return True
        If checkExistence(entry.DetectFiles, AddressOf checkPathExist) Then Return True
        If checkExistence(entry.SpecialDetect, AddressOf checkSpecialDetects) Then Return True

        ' Return true for the case where we have only a DetectOS and we meet its criteria
        Dim onlyHasDetOS = entry.SpecialDetect.KeyCount + entry.DetectFiles.KeyCount + entry.Detects.KeyCount = 0
        gLog("No other detection keys found than DetectOS", onlyHasDetOS And hasMetDetOS, descend:=True)
        If hasMetDetOS And onlyHasDetOS Then Return True

        ' Return true for the case where we have no valid detect criteria
        Dim hasNoDetectKeys = entry.DetectOS.KeyCount + entry.DetectFiles.KeyCount + entry.Detects.KeyCount + entry.SpecialDetect.KeyCount = 0
        gLog("No detect keys found, entry will be retained.", hasNoDetectKeys, descend:=True)
        If hasNoDetectKeys Then Return True

        gLog(descend:=True)

        Return False

    End Function

    ''' <summary> Audits the given entry for legacy codepaths in the machine's VirtualStore </summary>
    ''' <param name="entry"> The <c> winapp2entry </c> to audit </param>
    Private Sub virtualStoreChecker(ByRef entry As winapp2entry)

        vsKeyChecker(entry.FileKeys)
        vsKeyChecker(entry.RegKeys)
        vsKeyChecker(entry.ExcludeKeys)

    End Sub

    ''' <summary> Generates keys for VirtualStore locations that exist on the current system and inserts them into the given list </summary>
    ''' <param name="kl"> The <c> keyList </c> of FileKey, RegKey, or ExcludeKeys to be checked against the VirtualStore </param>
    Private Sub vsKeyChecker(ByRef kl As keyList)

        If kl.KeyCount = 0 Then Return

        Dim starterCount = kl.KeyCount

        Select Case kl.KeyType

            Case "FileKey", "ExcludeKey"

                mkVsKeys({"%ProgramFiles%", "%CommonAppData%", "%CommonProgramFiles%", "HKLM\Software"}, {"%LocalAppData%\VirtualStore\Program Files*", "%LocalAppData%\VirtualStore\ProgramData", "%LocalAppData%\VirtualStore\Program Files*\Common Files", "HKCU\Software\Classes\VirtualStore\MACHINE\SOFTWARE"}, kl)

            Case "RegKey"

                mkVsKeys({"HKLM\Software"}, {"HKCU\Software\Classes\VirtualStore\MACHINE\SOFTWARE"}, kl)

        End Select

        If Not starterCount = kl.KeyCount Then kl.renumberKeys(replaceAndSort(kl.toStrLst(True), "|", " \ \"))

    End Sub

    ''' <summary> Creates <c> iniKeys </c> to handle VirtualStore locations that correspond to paths given in <c> <paramref name="kl"/> </c> </summary>
    ''' <param name="findStrs"> An array of Strings to seek for in the key value </param>
    ''' <param name="replStrs"> An array of strings to replace the sought after key values </param>
    ''' <param name="kl"> The <c> keylist </c> to be processed </param>
    Private Sub mkVsKeys(findStrs As String(), replStrs As String(), ByRef kl As keyList)

        Dim initVals = kl.toStrLst(True)
        Dim keysToAdd As New keyList(kl.KeyType)

        For Each key In kl.Keys

            If Not key.vHasAny(findStrs, True) Then Continue For

            For i = 0 To findStrs.Length - 1

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

    ''' <summary> Creates the VirtualStore version of a given <c> iniKey </c> </summary>
    ''' <param name="findStr"> The normal filesystem path </param>
    ''' <param name="replStr"> The VirtualStore path </param>
    ''' <param name="key"> The <c> iniKey </c> to processed into a VirtualStore key </param>
    Private Function createVSKey(findStr As String, replStr As String, key As iniKey) As iniKey

        Return New iniKey($"{key.Name}={key.Value.Replace(findStr, replStr)}")

    End Function

    ''' <summary> Processes a list of <c> winapp2entries </c> and removes any from the list that wouldn't be detected by CCleaner </summary>
    ''' <param name="entryList"> The list of <c> winapp2entries </c> who detection criteria will be audited </param>
    Private Sub processEntryList(ByRef entryList As List(Of winapp2entry))

        ' If the entry's Detect criteria doesn't return true, prune it
        Dim sectionsToBePruned As New List(Of winapp2entry)
        entryList.ForEach(Sub(entry) If Not processEntryExistence(entry) Then sectionsToBePruned.Add(entry) Else virtualStoreChecker(entry))
        removeEntries(entryList, sectionsToBePruned)

    End Sub

    ''' <summary> Returns <c> True </c> if a SpecialDetect location exists, <c> False </c> otherwise </summary>
    ''' <param name="key"> A SpecialDetect format <c> iniKey </c> </param>
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

                For Each path In detChrome

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

    ''' <summary> Handles passing off checks from sources that may vary between file system and registry </summary>
    ''' <param name="path"> A filesystem or registry path whose existence will be audited </param>
    Private Function checkExist(path As String) As Boolean

        Return If(path.StartsWith("HK", StringComparison.InvariantCulture), checkRegExist(path), checkPathExist(path))

    End Function

    ''' <summary> Returns <c> True </c> if a given key exists in the Windows Registry, <c> False </c> otherwise </summary>
    ''' <param name="path"> A registry path to be audited for existence </param>
    Private Function checkRegExist(path As String) As Boolean

        Dim dir = path
        Dim root = getFirstDir(path)
        dir = dir.Replace(root & "\", "")
        Dim exists = getRegExists(root, dir)
        gLog($"{root}\{dir} exists", exists, indent:=True)
        ' If we didn't return anything above, registry location probably doesn't exist
        Return exists

    End Function

    ''' <summary> Returns <c> True </c> if a given key exists in the registry, <c> False </c> otherwise </summary>
    ''' <param name="root"> The registry hive that contains the key whose existence will be audited </param>
    ''' <param name="dir"> The path of the key whose existence will be audited </param>
    Private Function getRegExists(root As String, dir As String) As Boolean

        Try

            Select Case root

                Case "HKCU"

                    Return getCUKey(dir) IsNot Nothing

                Case "HKLM"

                    If getLMKey(dir) IsNot Nothing Then Return True

                    ' Support checking for 32bit applications on Win64
                    dir = root + "\" + dir
                    dir = dir.ToUpperInvariant.Replace("HKLM\SOFTWARE", "SOFTWARE\WOW6432Node")
                    Return getLMKey(dir) IsNot Nothing

                Case "HKU"

                    Return getUserKey(dir) IsNot Nothing

                Case "HKCR"

                    Return getCRKey(dir) IsNot Nothing

                Case Else

                    ' Reject malformated keys
                    gLog($"Your key seems to be malformatted (bad root? - root: {root} - expected 'HKCU','HKLM','HKU' or 'HKCR')", indent:=True)
                    Return False

            End Select

        Catch ex As UnauthorizedAccessException

            ' The most common (only?) exception here is a permissions one, so assume true if we hit because a permissions exception implies the key exists anyway.
            Return True

        End Try

        Return True

    End Function

    ''' <summary> Handles some CCleaner variables and logs if the current variable is ProgramFiles so the 32bit location can be checked later </summary>
    ''' <param name="dir"> A filesystem path to process for environment variables </param>
    ''' <param name="isProgramFiles"> Indicates that the %ProgramFiles% variable has been seen </param>
    Private Sub processEnvDirs(ByRef dir As String, ByRef isProgramFiles As Boolean)

        If dir.Contains("%") Then

            Dim splitDir = dir.Split(CChar("%"))
            Dim var = splitDir(1)
            Dim envDir = Environment.GetEnvironmentVariable(var)
            Dim userProfileDir = Environment.GetEnvironmentVariable("UserProfile")
            Select Case var

                ' %ProgramFiles% in CCleaner points to both C:\Program Files and C:\Program Files (x86)
                ' This particular case is handled later in the trim process, we simply note it here for that purpose 
                Case "ProgramFiles"

                    isProgramFiles = True

                ' %Documents% is a CCleaner-only variable and points to two paths depending on system 
                ' Windows XP:       %UserProfile%\My Documents
                ' Windows Vista+:   %UserProfile%\Documents	
                Case "Documents"

                    envDir = $"{userProfileDir}\{If(winVer = 5.1 Or winVer = 5.2, "My ", "")}Documents"

                ' %CommonAppData% is a CCleaner-only variable which creates parity between the all users profile in windows xp and the programdata folder in vista+ 
                ' Windows XP:       %AllUsersProfile%\Application Data
                ' Windows Vista+    %AllUsersProfile%\
                Case "CommonAppData"

                    envDir = $"{Environment.GetEnvironmentVariable("AllUsersProfile")}\{If(winVer = 5.1 Or winVer = 5.2, "Application Data\", "")}"

                ' %LocalLowAppData% is a CCleaner-only variable which points to %UserProfile%\AppData\LocalLow
                Case "LocalLowAppData"

                    envDir = $"{Environment.GetEnvironmentVariable("LocalAppData").Replace("Local", "LocalLow")}"

                ' %Pictures% is a CCleaner-only variable which points to two paths depending on system 
                ' Windows XP:       %UserProfile%\My Documents\My Pictures
                ' Windows Vista+:   %UserProfile%\Pictures
                Case "%Pictures%"

                    envDir = $"{userProfileDir}\{If(winVer = 5.1 Or winVer = 5.2, "My Documents\My ", "")}Pictures"

                ' %Music% is a CCleaner-only variable which points to two paths depending on system 
                ' Windows XP:       %UserProfile%\My Documents\My Music
                ' Windows Vista+:   %UserProfile%\Music
                Case "%Music%"

                    envDir = $"{userProfileDir}\{If(winVer = 5.1 Or winVer = 5.2, "My Documents\My ", "")}Music"

                ' %Video% is a CCleaner-only variable which points to two paths depending on system 
                ' Windows XP:       %UserProfile%\My Documents\My Videos
                ' Windows Vista+:   %UserProfile%\Videos
                Case "%Video%"

                    envDir = $"{userProfileDir}\{If(winVer = 5.1 Or winVer = 5.2, "My Documents\My ", "")}Videos"

            End Select

            dir = envDir + splitDir(2)

        End If

    End Sub

    ''' <summary> Returns <c> True </c> if a path exists on the file system, <c> False </c> otherwise </summary>
    ''' <param name="key"> A filesystem path </param>
    Private Function checkPathExist(key As String) As Boolean

        ' Make sure we get the proper path for environment variables
        Dim isProgramFiles = False
        Dim dir = key
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

        Catch ex As UnauthorizedAccessException

            Return True

        End Try

        Return False

    End Function

    ''' <summary> Swaps out a directory with the ProgramFiles parameterization on 64bit computers </summary>
    ''' <param name="dir"> The file system path to be modified </param>
    ''' <param name="key"> The original state of the path </param>
    Private Sub swapDir(ByRef dir As String, key As String)

        Dim envDir = Environment.GetEnvironmentVariable("ProgramFiles(x86)")
        dir = envDir & key.Split(CChar("%"))(2)

    End Sub

    ''' <summary> Interprets parameterized wildcards for the current system </summary>
    ''' <param name="dir"> A path containing a wildcard </param>
    Private Function expandWildcard(dir As String, isFileSystem As Boolean) As Boolean

        gLog("Expanding Wildcard: " & dir, ascend:=True)

        ' This will handle wildcards anywhere in a path even though CCleaner only supports them at the end for DetectFiles
        Dim possibleDirs As New strList
        Dim currentPaths As New strList

        ' Split the given string into sections by directory
        Dim splitDir = dir.Split(CChar("\"))
        For Each pathPart In splitDir

            ' If this directory parameterization includes a wildcard, expand it appropriately
            ' This probably wont work if a string for some reason starts with a *
            If pathPart.Contains("*") Then

                For Each currentPath In currentPaths.Items

                    If currentPath.Length = 0 Then gLog(NameOf(currentPath) & " is empty, aborting wildcard expansion", descend:=True) : Return False

                    ' Query the existence of child paths for each current path we hold
                    If isFileSystem Then

                        gLog("Investigating: " & pathPart & " as a subdir of " & currentPath, indent:=True)

                        Try

                            ' If there are any possibilities, add them to our possibility list
                            Dim possibilities = Directory.GetDirectories(currentPath, pathPart)
                            possibleDirs.add(possibilities, possibilities.Any)

                        Catch ex As ArgumentException

                            ' These are thrown by currentPaths containing illegal characters, we'll assume this means the target doesn't exist
                            Return False

                        Catch ex As UnauthorizedAccessException

                            ' Assume that if there's some directory we don't have access to, the target exists and we just can't see it 
                            Return True

                        End Try

                    Else

                        ' Registry Query here

                    End If

                Next

                ' If no possibilities remain, the wildcard parameterization hasn't left us with any real paths on the system, so we may return false.

                If possibleDirs.Count = 0 Then gLog("Wildcard parameterization did not return any valid paths", descend:=True) : Return False

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

                        If Not path.EndsWith("\", StringComparison.InvariantCulture) And Not path.Length = 0 Then path += "\"
                        newCurPaths.add($"{path}{pathPart}\", Directory.Exists($"{path}{pathPart}\"))

                    Next

                    currentPaths = newCurPaths
                    If currentPaths.Count = 0 Then gLog("Wildcard parameterization did not return any valid paths", descend:=True) : Return False

                End If

            End If

        Next

        ' If any file/path exists, return true
        For Each currDir In currentPaths.Items

            If Directory.Exists(currDir) Or File.Exists(currDir) Then gLog("Wildcard parameterization did not return any valid paths", descend:=True) : Return True

        Next

        gLog(descend:=True)
        Return False

    End Function

    ''' <summary> Returns <c> True </c> if the system satisfies the DetectOS citeria, <c> False </c> otherwise </summary>
    ''' <param name="value"> The DetectOS criteria to be checked </param>
    Private Function checkDetOS(value As String) As Boolean

        Dim splitKey = value.Split(CChar("|"))

        Select Case True

            Case value.StartsWith("|", StringComparison.InvariantCultureIgnoreCase)

                Return Not winVer > Val(splitKey(1))

            Case value.EndsWith("|", StringComparison.InvariantCultureIgnoreCase)

                Return Not winVer < Val(splitKey(0))

            Case Else

                Return winVer >= Val(splitKey(0)) And winVer <= Val(splitKey(1))

        End Select

    End Function

End Module