'    Copyright (C) 2018-2026 Hazel Ward
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

''' <summary>
''' <c> Trim </c> is a winapp2ool module that evaluates the detection criteria of each entry in
''' a winapp2.ini file against the current machine and removes any entries whose detection criteria
''' are not satisfied, producing a system-specific subset of the full database. <br /><br />
'''
''' The full winapp2.ini database contains entries for thousands of applications, the vast majority
''' of which will not be installed on any given machine. Trimming eliminates those irrelevant entries,
''' reducing load time for cleaning software (particularly CCleaner) and avoiding spurious scans. <br /><br />
'''
''' Detection criteria and evaluation order: <br /><br />
'''
''' <list type="table">
'''
''' <item>
''' <term> DetectOS </term>
''' <description>
''' Evaluated before the other detection keys. If present and not satisfied by the current
''' Windows version, the entry is immediately discarded without checking any other criteria. <br />
''' If satisfied and no other detection keys are present, the entry is retained. <br />
''' DetectOS values take the form <c> VERSION| </c> (minimum), <c> |VERSION </c> (maximum),
''' or <c> VERSION1|VERSION2 </c> (range), where version numbers are major.minor doubles
''' (e.g. <c> 6.1 </c> for Windows 7, <c> 10.0 </c> for Windows 10).
''' </description>
''' </item>
'''
''' <item>
''' <term> Detect </term>
''' <description>
''' Registry paths checked for existence. If any Detect key matches a registry key
''' present on the current system, the entry is retained. For a key under
''' <c> HKLM\Software </c> we also check <c> HKLM\SOFTWARE\WOW6432Node </c>.
''' </description>
''' </item>
'''
''' <item>
''' <term> DetectFile </term>
''' <description>
''' Filesystem paths (supporting wildcards) checked for existence. If any DetectFile key
''' matches a file or directory present on the current system, the entry is retained.
''' A wildcard only matches directory names, so a wildcard in a file name never matches, and
''' a wildcard search that reaches a folder we aren't allowed to read counts as a match.
''' </description>
''' </item>
'''
''' <item>
''' <term> SpecialDetect </term>
''' <description>
''' Deprecated CCleaner variable checked against a hardcoded list of well-known browser/application
''' install locations. Supported values: <c> DET_CHROME </c>, <c> DET_MOZILLA </c>,
''' <c> DET_THUNDERBIRD </c>, <c> DET_OPERA </c>. If the corresponding application is
''' detected, the entry is retained. Parsing for this is left only to support very old 
''' versions of winapp2.ini. 
''' </description>
''' </item>
'''
''' </list>
'''
''' <br />
''' Entries with no detection keys of any kind are retained unless the excludes file names them. <br /><br />
'''
''' Include and exclude overrides: <br /><br />
'''
''' When enabled, the includes file (<c> TrimFile2 </c>) and excludes file (<c> TrimFile4 </c>)
''' each contain a list of entry names (one per ini section header) that unconditionally override
''' the normal detection evaluation. An entry whose name appears in the includes file is always
''' retained regardless of whether its detection criteria are satisfied. An entry whose name
''' appears in the excludes file is always removed regardless of whether its detection criteria
''' are satisfied. Include and exclude checks run before any detection evaluation, includes
''' first, so an entry named in both files is retained. <br /><br />
'''
''' Environment variable expansion: <br /><br />
'''
''' In addition to standard Windows environment variables, Trim resolves several CCleaner-specific
''' variables that do not exist natively in the Windows environment. We expand only the first
''' variable in a filesystem path and assume the path starts with it:
'''
''' <list type="table">
'''
''' <item>
''' <term> %Documents% </term>
''' <description> <c> %UserProfile%\My Documents </c> on XP; <c> %UserProfile%\Documents </c> on Vista+ </description>
''' </item>
'''
''' <item>
''' <term> %CommonAppData% </term>
''' <description> <c> %AllUsersProfile%\Application Data </c> on XP; <c> %AllUsersProfile%\ </c> on Vista+ </description>
''' </item>
'''
''' <item>
''' <term> %LocalLowAppData% </term>
''' <description> <c> %UserProfile%\AppData\LocalLow </c> </description>
''' </item>
'''
''' <item>
''' <term> %Pictures% </term>
''' <description> <c> %UserProfile%\My Documents\My Pictures </c> on XP; <c> %UserProfile%\Pictures </c> on Vista+ </description>
''' </item>
'''
''' <item>
''' <term> %Music% </term>
''' <description> <c> %UserProfile%\My Documents\My Music </c> on XP; <c> %UserProfile%\Music </c> on Vista+ </description>
''' </item>
'''
''' <item>
''' <term> %Video% </term>
''' <description> <c> %UserProfile%\My Documents\My Videos </c> on XP; <c> %UserProfile%\Videos </c> on Vista+ </description>
''' </item>
'''
''' </list>
'''
''' <br />
''' %ProgramFiles% receives special handling: paths under %ProgramFiles% are checked against both
''' the native-bitness Program Files directory and the 32-bit Program Files (x86) directory on
''' 64-bit systems, so entries covering 32-bit applications installed on 64-bit Windows are
''' retained correctly. <br /><br />
'''
''' VirtualStore augmentation: <br /><br />
'''
''' For every entry it keeps, including those kept by the includes file or for having no
''' detection keys, Trim inspects the entry's FileKeys, RegKeys, and ExcludeKeys and generates
''' additional keys covering VirtualStore locations that correspond to paths found under
''' <c> %ProgramFiles% </c>, <c> %CommonAppData% </c>, <c> %CommonProgramFiles% </c>, and
''' <c> HKLM\Software </c> (only <c> HKLM\Software </c> for RegKeys). The match is case-sensitive.
''' VirtualStore keys are only appended when the corresponding VirtualStore path actually exists
''' on the current system. When we add any, we re-sort and renumber the entry's FileKeys, RegKeys and ExcludeKeys.
''' </summary>
Public Module Trim

    ''' <summary>
    ''' The major/minor version number on the current system
    ''' </summary>
    Private Property winVer As Double

    Private _includes As iniFile = Nothing
    Private _excludes As iniFile = Nothing

    ''' <summary>
    ''' Handles command-line arguments for <c> Trim </c>, starting from the default settings
    ''' rather than any saved ones, then runs the trim
    ''' </summary>
    '''
    ''' <remarks>
    ''' File arguments: <c> -1d </c>/<c> -1f </c> set the winapp2.ini to trim, <c> -2d </c>/<c> -2f </c>
    ''' the includes file, <c> -3d </c>/<c> -3f </c> the output file, and <c> -4d </c>/<c> -4f </c>
    ''' the excludes file.
    ''' Flags:
    ''' <list type="bullet">
    ''' <item><c> -d </c>: download the winapp2.ini for <see cref="CurrentWinappFlavor"/> and trim
    ''' that instead of <c> TrimFile1 </c></item>
    ''' <item><c> -includes </c>: enable the includes file (entries listed within are never trimmed)</item>
    ''' <item><c> -excludes </c>: enable the excludes file (entries listed within are always trimmed)</item>
    ''' </list>
    ''' </remarks>
    Public Sub handleCmdLine()

        InitDefaultTrimSettings()

        Dim spec As New CliArgSpec(NameOf(Trim))
        spec.WithFile(1, TrimFile1).WithFile(2, TrimFile2).WithFile(3, TrimFile3).WithFile(4, TrimFile4) _
            .WithFlag("-includes", Sub() UseTrimIncludes = Not UseTrimIncludes) _
            .WithFlag("-excludes", Sub() UseTrimExcludes = Not UseTrimExcludes) _
            .WithDownload(Sub() DownloadFileToTrim = Not DownloadFileToTrim, Function() DownloadFileToTrim) _
            .Parse()

        initTrim()

    End Sub

    ''' <summary>
    ''' Trims a winapp2.ini from outside the module. This overwrites the module's
    ''' <c> TrimFile1 </c>, <c> TrimFile3 </c> and <c> DownloadFileToTrim </c> settings for the
    ''' rest of the session, and the include and exclude settings still apply.
    ''' </summary>
    '''
    ''' <param name="firstFile">
    ''' The winapp2.ini file to be trimmed
    ''' </param>
    '''
    ''' <param name="thirdFile">
    ''' The path on disk to which the trimmed file will be saved
    ''' </param>
    '''
    ''' <param name="d">
    ''' Indicates whether to download the input winapp2.ini from GitHub instead of reading
    ''' <paramref name="firstFile"/>
    ''' </param>
    Public Sub remoteTrim(firstFile As iniFileChooser,
                          thirdFile As iniFileChooser,
                          d As Boolean)

        TrimFile1 = firstFile
        TrimFile3 = thirdFile
        DownloadFileToTrim = d
        initTrim()

    End Sub

    ''' <summary>
    ''' Loads or downloads winapp2.ini, trims it, prints a summary, and writes the result to
    ''' <c> TrimFile3 </c>. We return without trimming if the local input is empty or missing, or if we
    ''' need to download and a live connection check fails. That check ignores <c> isOffline </c>. A download goes through <see cref="getWinappLink"/>, so
    ''' it fetches the current flavor.
    ''' </summary>
    Public Sub initTrim()

        ' Don't try to trim an empty file
        If Not DownloadFileToTrim Then

            Dim winapp = TrimFile1.Load(TrimModuleSettingsChanged, NameOf(Trim), NameOf(TrimFile1), NameOf(TrimModuleSettingsChanged))
            If Not enforceFileHasContent(winapp) Then Return

        End If

        ' Ensure we have an online connection before continuing if necessary
        Dim noNtwk = "Internet connection lost! Please check your network connection and try again"
        If denyActionWithHeader(DownloadFileToTrim AndAlso Not checkOnline(), noNtwk) Then Return

        Dim winapp2 As New winapp2file(If(DownloadFileToTrim, getRemoteIniFile(getWinappLink), TrimFile1.Load(TrimModuleSettingsChanged, NameOf(Trim), NameOf(TrimFile1), NameOf(TrimModuleSettingsChanged))))

        clrConsole()
        Dim progress As New MenuSection
        progress.AddTopBorder() _
                .AddColoredLine("Trimming... Please wait, this may take a moment...", ConsoleColor.DarkCyan, centered:=True) _
                .AddBottomBorder()
        progress.Print()

        Dim entryCountBeforeTrim = winapp2.Count

        ' Perform the trim
        trimFile(winapp2)

        Dim difference = entryCountBeforeTrim - winapp2.Count
        Dim pct = Math.Round((difference / entryCountBeforeTrim) * 100)

        gLog($"{difference} entries trimmed from winapp2.ini from a total of {entryCountBeforeTrim} ({pct}%)")
        gLog($"{winapp2.Count} entries remain.")

        ' Print trim summary to the user
        clrConsole()
        Dim out As New MenuSection
        out.AddTopBorder() _
           .AddColoredLine("Trim Complete", ConsoleColor.DarkCyan, centered:=True) _
           .AddDivider() _
           .AddLine($"Initial entry count: {entryCountBeforeTrim}") _
           .AddLine($"Trimmed entry count: {winapp2.Count}") _
           .AddLine($"{difference} entries trimmed from winapp2.ini ({pct}%)") _
           .AddDivider() _
           .AddLine(anyKeyStr, centered:=True) _
           .AddBottomBorder()
        out.Print()

        ' Save the trimmed file back to disk
        Dim saved = iniFile.Empty(TrimFile3.Dir, TrimFile3.Name).OverwriteToFile(winapp2.ToWinapp2String())

        ' If we downloaded the latest file, then we probably can mark winapp2 as having been updated
        If DownloadFileToTrim AndAlso saved Then waUpdateIsAvail = False
        setNextMenuHeaderText($"{TrimFile3.Name} saved", saved)
        setNextMenuHeaderText($"{TrimFile3.Name} was not saved", Not saved, ConsoleColor.Red)

        crk()

    End Sub

    ''' <summary>
    ''' Trims a <c> winapp2file </c> in place, removing entries not relevant to the current system,
    ''' adding VirtualStore keys to the ones we keep, and then sorting the entries. We read the
    ''' includes and excludes files from <c> TrimFile2 </c> and <c> TrimFile4 </c> here, when
    ''' <c> UseTrimIncludes </c> and <c> UseTrimExcludes </c> are on.
    ''' </summary>
    '''
    ''' <param name="winapp2">
    ''' A <c> winapp2file </c> to be trimmed to fit the current system
    ''' </param>
    Public Sub trimFile(winapp2 As winapp2file)

        If winapp2 Is Nothing Then argIsNull(NameOf(winapp2)) : Return

        _includes = If(UseTrimIncludes, iniFile.FromFile(TrimFile2.Path()), Nothing)
        _excludes = If(UseTrimExcludes, iniFile.FromFile(TrimFile4.Path()), Nothing)

        If winVer = Nothing Then winVer = getWinVer()

        Dim toRemove As New List(Of winapp2entry)

        Dim results = winapp2.Entries.AsParallel().AsOrdered().Select(
            Function(entry)
                Using cap = gLogCapture()
                    Dim retained = processEntryExistence(entry)
                    If retained Then virtualStoreChecker(entry)
                    Return New With {.Entry = entry, .Retained = retained, .Log = cap.Lines}
                End Using
            End Function).ToList()

        For Each r In results
            If r.Retained Then EmitCaptured(r.Log) Else toRemove.Add(r.Entry)
        Next

        For Each entry In toRemove : winapp2.RemoveEntry(entry) : Next

        winapp2.SortEntries()

    End Sub

    ''' <summary>
    ''' Evaluates a list of detection keys to observe whether they exist on the current machine
    ''' </summary>
    '''
    ''' <param name="keys">
    ''' The detection keys to be evaluated
    ''' </param>
    '''
    ''' <param name="chkExist">
    ''' The <c> function </c> that evaluates each key value
    ''' </param>
    '''
    Private Function checkExistence(keys As IReadOnlyList(Of iniKey),
                                    chkExist As Func(Of String, Boolean)) As Boolean

        If keys.Count = 0 Then Return False

        For Each key In keys

            If Not chkExist(key.Value) Then Continue For

            gLog($"{key.Value} matched a path on the system", Not key.KeyType = "DetectOS", buffr:=True)
            Return True

        Next

        Return False

    End Function

    ''' <summary>
    ''' Returns whether we keep <paramref name="entry"/>: the includes and excludes files decide
    ''' first, then the detection criteria are audited against the current system
    ''' </summary>
    '''
    ''' <param name="entry">
    ''' A <c> winapp2entry </c> whose detection criteria will be audited
    ''' </param>
    '''
    Private Function processEntryExistence(entry As winapp2entry) As Boolean

        gLog("", leadr:=True)

        Using gLogScope($"Processing entry: {entry.Name}")

            ' Respect the include/excludes
            Dim IsInIncludes = UseTrimIncludes AndAlso _includes IsNot Nothing AndAlso _includes.Contains(entry.Name)
            If IsInIncludes Then gLog("Retaining entry: " & entry.Name, leadr:=True, buffr:=True) : Return True
            Dim isInExcludes = UseTrimExcludes AndAlso _excludes IsNot Nothing AndAlso _excludes.Contains(entry.Name)
            If isInExcludes Then gLog("Discarding entry: " & entry.Name, leadr:=True, buffr:=True) : Return False

            ' Process the DetectOS if we have one, take note if we meet the criteria, otherwise return false
            Dim hasMetDetOS = False

            If Not entry.DetectOS.Count = 0 Then

                If winVer = Nothing Then winVer = getWinVer()

                hasMetDetOS = checkExistence(entry.DetectOS, AddressOf checkDetOS)
                gLog($"Met DetectOS criteria. {winVer} satisfies {entry.DetectOS(0).Value}", hasMetDetOS)
                gLog($"Did not meet DetectOS criteria. {winVer} does not satisfy {entry.DetectOS(0).Value}", Not hasMetDetOS)

                If Not hasMetDetOS Then Return False

            End If

            ' Process any other Detect criteria we have
            Dim DetectExists = checkExistence(entry.Detects, AddressOf checkRegExist)
            Dim DetectFileExists = DetectExists OrElse checkExistence(entry.DetectFiles, AddressOf checkPathExist)
            Dim Detected = DetectFileExists OrElse checkExistence(entry.SpecialDetect, AddressOf checkSpecialDetects)

            If Detected Then
                gLog("Retaining entry: " & entry.Name, cond:=Detected, leadr:=True, buffr:=True)
                Return True
            End If

            ' Return true for the case where we have only a DetectOS and we meet its criteria
            gLog("No other detection keys found than DetectOS", entry.HasOnlyDetectOS AndAlso hasMetDetOS)
            If entry.HasOnlyDetectOS AndAlso hasMetDetOS Then gLog("Retaining entry: " & entry.Name, leadr:=True, buffr:=True) : Return True

            ' Return true for the case where we have no valid detect criteria
            gLog("No detect keys found, entry will be retained.", Not entry.HasDetectionKey)
            If Not entry.HasDetectionKey Then gLog("Retaining entry: " & entry.Name, leadr:=True, buffr:=True) : Return True

            gLog("Discarding entry: " & entry.Name, leadr:=True, buffr:=True)
            Return False

        End Using

    End Function

    ''' <summary>
    ''' Adds VirtualStore counterparts of the entry's FileKeys, RegKeys and ExcludeKeys whose
    ''' VirtualStore path exists on this machine. When we add any, we re-sort and renumber the
    ''' entry's FileKeys, RegKeys and ExcludeKeys.
    ''' </summary>
    '''
    ''' <param name="entry">
    ''' The <c> winapp2entry </c> to audit
    ''' </param>
    '''
    Private Sub virtualStoreChecker(entry As winapp2entry)

        Using gLogScope("Attempting to generate any neccessary VirtualStore keys for " & entry.Name)

            Dim newKeys As New List(Of iniKey)
            collectVsKeys(entry.FileKeys, newKeys)
            collectVsKeys(entry.RegKeys, newKeys)
            collectVsKeys(entry.ExcludeKeys, newKeys)

            For Each key In newKeys : entry.AddKey(key) : Next

            If newKeys.Count > 0 Then entry.RenumberKeys()

        End Using

    End Sub

    ''' <summary>
    ''' Collects new VirtualStore counterpart keys for the given key list and appends them to
    ''' <paramref name="newKeys"/>. Only keys whose corresponding VirtualStore path exists on
    ''' the current system are included, and we skip a counterpart whose value is already in
    ''' <paramref name="keys"/>. Each new key keeps its source key's name until the entry is
    ''' renumbered.
    ''' </summary>
    '''
    ''' <param name="keys">
    ''' The FileKey, RegKey, or ExcludeKey collection to scan. We read the key type from the
    ''' first key and apply it to all of them.
    ''' </param>
    '''
    ''' <param name="newKeys">
    ''' New VirtualStore keys are appended here
    ''' </param>
    '''
    Private Sub collectVsKeys(keys As IReadOnlyList(Of iniKey),
                              newKeys As List(Of iniKey))

        If keys.Count = 0 Then Return

        Dim findStrs() As String
        Dim replStrs() As String

        Select Case keys(0).KeyType

            Case "FileKey", "ExcludeKey"
                findStrs = {"%ProgramFiles%", "%CommonAppData%", "%CommonProgramFiles%", "HKLM\Software"}
                replStrs = {"%LocalAppData%\VirtualStore\Program Files*",
                            "%LocalAppData%\VirtualStore\ProgramData",
                            "%LocalAppData%\VirtualStore\Program Files*\Common Files",
                            "HKCU\Software\Classes\VirtualStore\MACHINE\SOFTWARE"}

            Case "RegKey"
                findStrs = {"HKLM\Software"}
                replStrs = {"HKCU\Software\Classes\VirtualStore\MACHINE\SOFTWARE"}

            Case Else
                Return

        End Select

        Dim initVals = keys.Select(Function(k) k.Value).ToList()
        Dim keysToAdd As New List(Of iniKey)

        ' Pass 1: collect candidate VS keys
        For Each key In keys

            If Not key.vHasAny(findStrs, True) Then Continue For

            For i = 0 To findStrs.Length - 1
                Dim newVal = key.Value.Replace(findStrs(i), replStrs(i))
                If initVals.Contains(newVal) Then Continue For
                If key.Value = newVal Then Continue For
                keysToAdd.Add(New iniKey($"{key.Name}={newVal}"))
            Next

        Next

        ' Pass 2: filter by system existence
        For Each key In keysToAdd
            If checkExist(getPathFromValue(key.Value, key.KeyType)) Then newKeys.Add(key)
        Next

    End Sub

    ''' <summary>
    ''' Extracts the filesystem or registry path from a key value string for use in an existence
    ''' check. A RegKey value comes back whole, including any <c> |ValueName </c> part.
    ''' </summary>
    '''
    ''' <param name="value">
    ''' The raw value of a FileKey, RegKey, or ExcludeKey
    ''' </param>
    '''
    ''' <param name="keyType">
    ''' The key type string ("FileKey", "RegKey", or "ExcludeKey")
    ''' </param>
    '''
    Private Function getPathFromValue(value As String, keyType As String) As String

        Select Case keyType
            Case "FileKey" : Return New fileKeyParams(value).Path
            Case "ExcludeKey" : Return New excludeKeyParams(value).Path
            Case Else : Return value
        End Select

    End Function

    ''' <summary>
    ''' Returns whether any location for the SpecialDetect value exists. An unrecognized value
    ''' returns <c> False </c>.
    ''' </summary>
    ''' 
    ''' <param name="key"> 
    ''' A SpecialDetect value, such as <c> DET_CHROME </c>
    ''' </param>
    ''' 
    Private Function checkSpecialDetects(ByVal key As String) As Boolean

        Select Case key

            Case "DET_CHROME"

                Dim detChrome As New List(Of String) _
                        From {"%AppData%\ChromePlus\chrome.exe",
                              "%LocalAppData%\Chromium\Application\chrome.exe",
                              "%LocalAppData%\Chromium\chrome.exe",
                              "%LocalAppData%\Flock\Application\flock.exe",
                              "%LocalAppData%\Google\Chrome SxS\Application\chrome.exe",
                              "%LocalAppData%\Google\Chrome\Application\chrome.exe",
                              "%LocalAppData%\RockMelt\Application\rockmelt.exe",
                              "%LocalAppData%\SRWare Iron\iron.exe",
                              "%ProgramFiles%\Chromium\Application\chrome.exe",
                              "%ProgramFiles%\SRWare Iron\iron.exe",
                              "%ProgramFiles%\Chromium\chrome.exe",
                              "%ProgramFiles%\Flock\Application\flock.exe",
                              "%ProgramFiles%\Google\Chrome SxS\Application\chrome.exe",
                              "%ProgramFiles%\Google\Chrome\Application\chrome.exe",
                              "%ProgramFiles%\RockMelt\Application\rockmelt.exe",
                              "HKCU\Software\Chromium",
                              "HKCU\Software\SuperBird",
                              "HKCU\Software\Torch",
                              "HKCU\Software\Vivaldi",
                              "HKCU\Software\CentBrowser",
                              "HKCU\Software\Comodo\Dragon",
                              "HKCU\Software\CocCoc\Browser",
                              "HKCU\Software\Epic Privacy Browser",
                              "HKCU\Software\Yandex\YandexBrowser",
                              "HKCU\Software\Slimjet",
                              "HKCU\Software\Iridium"
                            }

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

    ''' <summary>
    ''' Returns whether <paramref name="path"/> exists, checking the registry when it starts with
    ''' <c> HK </c> and the file system otherwise
    ''' </summary>
    ''' 
    ''' <param name="path">
    ''' A filesystem or registry path whose existence will be audited 
    ''' </param>
    ''' 
    Private Function checkExist(path As String) As Boolean

        Return If(path.StartsWith("HK", StringComparison.InvariantCulture), checkRegExist(path), checkPathExist(path))

    End Function

    ''' <summary>
    ''' Returns whether a registry key exists, splitting the hive off the front of
    ''' <paramref name="path"/>. Wildcards aren't expanded.
    ''' </summary>
    ''' 
    ''' <param name="path">
    ''' A registry path to be audited for existence 
    ''' </param>
    ''' 
    Private Function checkRegExist(path As String) As Boolean

        Dim dir = path
        Dim root = getFirstDir(path)
        dir = dir.Replace(root & "\", "")
        Dim exists = getRegExists(root, dir)
        gLog($"{root}\{dir} exists", exists, buffr:=True)
        ' If we didn't return anything above, registry location probably doesn't exist
        Return exists

    End Function

    ''' <summary>
    ''' Returns whether a key exists in the registry. Under <c> HKLM </c> we fall back to the
    ''' <c> SOFTWARE\WOW6432Node </c> view for 32-bit applications. An unrecognized hive logs a
    ''' message and returns <c> False </c>, and an <c> UnauthorizedAccessException </c> returns
    ''' <c> True </c>.
    ''' </summary>
    '''
    ''' <param name="root">
    ''' The registry hive that contains the key whose existence will be audited 
    ''' </param>
    ''' 
    ''' <param name="dir"> 
    ''' The path of the key whose existence will be audited
    ''' </param>
    ''' 
    Private Function getRegExists(root As String,
                                  dir As String) As Boolean

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
                    gLog($"Your key seems to be malformatted (bad root? - root: {root} - expected 'HKCU','HKLM','HKU' or 'HKCR')")
                    Return False

            End Select

        Catch ex As UnauthorizedAccessException

            ' The most common (only?) exception here is a permissions one, so assume true if we hit because a permissions exception implies the key exists anyway.
            Return True

        End Try

        Return True

    End Function

    ''' <summary>
    ''' Expands the first environment variable in <paramref name="dir"/>, including the
    ''' CCleaner-only ones, and flags <c> %ProgramFiles% </c> so the caller can retry under
    ''' Program Files (x86). We keep only the text between that variable and the next
    ''' <c> % </c>, so anything before the variable is dropped.
    ''' </summary>
    '''
    ''' <param name="dir">
    ''' A filesystem path to process for environment variables. We replace it with the expanded path.
    ''' </param>
    '''
    ''' <param name="isProgramFiles">
    ''' Set to <c> True </c> if the variable is <c> %ProgramFiles% </c>. Never set back to <c> False </c>.
    ''' </param>
    '''
    ''' <returns>
    ''' <c> True </c> if the path has an opening <c> % </c> with no closing one, <br />
    ''' <c> False </c> otherwise
    ''' </returns>
    ''' 
    Private Function processEnvDirs(ByRef dir As String,
                                    ByRef isProgramFiles As Boolean) As Boolean

        Dim errDetected = False

        If dir.Contains("%") Then

            Dim splitDir = dir.Split(CChar("%"))
            Dim var = splitDir(1)
            Dim envDir = Environment.GetEnvironmentVariable(var)
            Dim userProfileDir = Environment.GetEnvironmentVariable("UserProfile")
            Dim isWinXP = winVer = 5.1 OrElse winVer = 5.2
            Select Case var

                ' %ProgramFiles% in CCleaner points to both C:\Program Files and C:\Program Files (x86)
                ' This particular case is handled later in the trim process, we simply note it here for that purpose 
                Case "ProgramFiles"

                    isProgramFiles = True

                ' %Documents% is a CCleaner-only variable and points to two paths depending on system 
                ' Windows XP:       %UserProfile%\My Documents
                ' Windows Vista+:   %UserProfile%\Documents	
                Case "Documents"

                    envDir = $"{userProfileDir}\{If(isWinXP, "My ", "")}Documents"

                ' %CommonAppData% is a CCleaner-only variable which creates parity between the all users profile in windows xp and the programdata folder in vista+ 
                ' Windows XP:       %AllUsersProfile%\Application Data
                ' Windows Vista+    %AllUsersProfile%\
                Case "CommonAppData"

                    envDir = $"{Environment.GetEnvironmentVariable("AllUsersProfile")}\{If(isWinXP, "Application Data\", "")}"

                ' %LocalLowAppData% is a CCleaner-only variable which points to %UserProfile%\AppData\LocalLow
                Case "LocalLowAppData"

                    envDir = $"{Environment.GetEnvironmentVariable("LocalAppData").Replace("Local", "LocalLow")}"

                ' %Pictures% is a CCleaner-only variable which points to two paths depending on system 
                ' Windows XP:       %UserProfile%\My Documents\My Pictures
                ' Windows Vista+:   %UserProfile%\Pictures
                Case "Pictures"

                    envDir = $"{userProfileDir}\{If(isWinXP, "My Documents\My ", "")}Pictures"

                ' %Music% is a CCleaner-only variable which points to two paths depending on system 
                ' Windows XP:       %UserProfile%\My Documents\My Music
                ' Windows Vista+:   %UserProfile%\Music
                Case "Music"

                    envDir = $"{userProfileDir}\{If(isWinXP, "My Documents\My ", "")}Music"

                ' %Video% is a CCleaner-only variable which points to two paths depending on system 
                ' Windows XP:       %UserProfile%\My Documents\My Videos
                ' Windows Vista+:   %UserProfile%\Videos
                Case "Video"

                    envDir = $"{userProfileDir}\{If(isWinXP, "My Documents\My ", "")}Videos"

            End Select

            Try

                dir = envDir + splitDir(2)

            Catch ex As IndexOutOfRangeException

                errDetected = True

            End Try

        End If

        Return errDetected

    End Function

    ''' <summary>
    ''' Returns whether a path exists on the file system, after expanding its environment
    ''' variable and any wildcards. A <c> %ProgramFiles% </c> path that isn't found is retried
    ''' under Program Files (x86). A malformed variable counts as present: outside <c> -s </c> we print an error to
    ''' the console and wait for a key press, from inside a parallel worker, before returning <c> True </c>.
    ''' </summary>
    '''
    ''' <param name="key">
    ''' A filesystem path, possibly starting with an environment variable
    ''' </param>
    ''' 
    Private Function checkPathExist(key As String) As Boolean

        ' Make sure we get the proper path for environment variables
        Dim isProgramFiles = False
        Dim dir = key

        If processEnvDirs(dir, isProgramFiles) Then

            cwl("Error: " & key & " contains a malformatted environment variable and has been ignored")
            cwl("The associated entry will be retained in the final output file")
            cwl("Press any key to continue")
            crk()
            Return True

        End If

        Try

            ' Process wildcards appropriately if we have them
            If dir.Contains("*") Then

                Dim exists = expandWildcard(dir, True)

                ' Small contingency for the isProgramFiles case

                If Not exists AndAlso isProgramFiles Then

                    swapDir(dir, key)
                    exists = expandWildcard(dir, True)

                End If

                Return exists

            End If

            ' Check out those file/folder paths
            If Directory.Exists(dir) OrElse File.Exists(dir) Then Return True

            ' If we didn't find it and we're looking in Program Files, check the (x86) directory
            If isProgramFiles Then

                swapDir(dir, key)
                Dim exists = Directory.Exists(dir) OrElse File.Exists(dir)
                Return exists

            End If

        Catch ex As UnauthorizedAccessException

            Return True

        End Try

        Return False

    End Function

    ''' <summary>
    ''' Sets <paramref name="dir"/> to <paramref name="key"/> with its leading
    ''' <c> %ProgramFiles% </c> replaced by the Program Files (x86) directory
    ''' </summary>
    '''
    ''' <param name="dir">
    ''' The file system path to be replaced
    ''' </param>
    '''
    ''' <param name="key">
    ''' The unexpanded path, starting with <c> %ProgramFiles% </c>
    ''' </param>
    ''' 
    Private Sub swapDir(ByRef dir As String,
                        key As String)

        Dim envDir = Environment.GetEnvironmentVariable("ProgramFiles(x86)")
        dir = envDir & key.Split(CChar("%"))(2)

    End Sub

    ''' <summary>
    ''' Returns whether any path matching the wildcards in <paramref name="dir"/> exists. We expand
    ''' each wildcard segment against directory names only, wherever it falls in the path. A
    ''' folder we aren't allowed to list counts as a match, and a path with illegal characters
    ''' counts as no match.
    ''' </summary>
    '''
    ''' <param name="dir">
    ''' A path containing a wildcard
    ''' </param>
    '''
    ''' <param name="isFileSystem">
    ''' Indicates whether <paramref name="dir"/> is a filesystem path. The registry branch is
    ''' empty, so <c> False </c> always returns <c> False </c>.
    ''' </param>
    ''' 
    Private Function expandWildcard(dir As String,
                                    isFileSystem As Boolean) As Boolean

        Using gLogScope("Expanding Wildcard: " & dir)

            ' This will handle wildcards anywhere in a path even though CCleaner only supports them at the end for DetectFiles
            Dim possibleDirs As New List(Of String)
            Dim currentPaths As New List(Of String)

            ' Split the given string into sections by directory
            Dim splitDir = dir.Split(CChar("\"))
            For Each pathPart In splitDir

                ' If this directory parameterization includes a wildcard, expand it appropriately
                ' This probably wont work if a string for some reason starts with a *
                If pathPart.Contains("*") Then

                    For Each currentPath In currentPaths

                        If currentPath.Length = 0 Then gLog(NameOf(currentPath) & " is empty, aborting wildcard expansion") : Return False

                        ' Query the existence of child paths for each current path we hold
                        If isFileSystem Then

                            gLog("Investigating: " & pathPart & " as a subdir of " & currentPath)

                            Try

                                ' If there are any possibilities, add them to our possibility list
                                Dim possibilities = Directory.GetDirectories(currentPath, pathPart)
                                possibleDirs.AddRange(possibilities)

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

                    If possibleDirs.Count = 0 Then gLog("Wildcard parameterization did not return any valid paths", buffr:=True) : Return False

                    ' Otherwise, clear the current paths and repopulate them with the possible paths
                    currentPaths.Clear()
                    currentPaths.AddRange(possibleDirs)
                    possibleDirs.Clear()

                Else

                    If currentPaths.Count = 0 Then

                        currentPaths.Add($"{pathPart}")
                        Continue For

                    End If

                    Dim newCurPaths As New List(Of String)

                    For Each path In currentPaths

                        If Not path.EndsWith("\", StringComparison.InvariantCulture) AndAlso Not path.Length = 0 Then path += "\"
                        If Directory.Exists($"{path}{pathPart}\") Then
                            newCurPaths.Add($"{path}{pathPart}\")
                        End If

                    Next

                    currentPaths = newCurPaths
                    If currentPaths.Count = 0 Then gLog("Wildcard parameterization did not return any valid paths") : Return False

                End If

            Next

            ' If any file/path exists, return true
            For Each currDir In currentPaths

                If Directory.Exists(currDir) OrElse File.Exists(currDir) Then gLog($"Wildcard parameterization returned a valid path: {currDir}", buffr:=True) : Return True

            Next

            ' If we make it this far, the path does not exist
            Return False

        End Using

    End Function

    ''' <summary>
    ''' Returns whether the system's Windows version satisfies a DetectOS value. Both bounds
    ''' are inclusive.
    ''' </summary>
    ''' 
    ''' <param name="value"> 
    ''' The DetectOS criteria to be checked 
    ''' </param>
    ''' 
    Private Function checkDetOS(value As String) As Boolean

        Dim splitKey = value.Split(CChar("|"))

        ' There's three cases here:
        ' |VERSION                  -> winVer <= VERSION (maximum)
        ' VERSION|                  -> winVer >= VERSION (minimum)
        ' VERSION1|VERSION2         -> VERSION1 <= winVer <= VERSION2
        Select Case True

            Case value.StartsWith("|", StringComparison.InvariantCultureIgnoreCase)

                Return Not winVer > Val(splitKey(1))

            Case value.EndsWith("|", StringComparison.InvariantCultureIgnoreCase)

                Return Not winVer < Val(splitKey(0))

            Case Else

                Return winVer >= Val(splitKey(0)) AndAlso winVer <= Val(splitKey(1))

        End Select

    End Function

End Module