'    Copyright (C) 2018-2024 Hazel Ward
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
Imports System.Globalization
Imports System.Text.RegularExpressions

''' <summary> 
''' Compares two winapp2.ini format <c> iniFile </c>s and summarizes the changes to the user 
''' <br /> 
''' <br /> NOTE: to "exist" here means for an entry of the same exact name (case sensitive) to exist 
''' <br /> Changes fall three major categories:
''' <list type="table">
''' <item> 
''' <term> Added entries </term>
''' <description> exist in the new file and not in the old file </description>
''' </item>
''' <item>
''' <term> Modified entries </term>
''' <description> exist in both the new file and the old file and have been changed in some way  </description>
''' </item>
''' <item>
''' <term> Removed entries </term>
''' <description> exist in the old file but not in the new file  </description>
''' </item>
''' </list> 
''' <br />
''' <br /> Additionally, Removed entries have three sub categories: 
''' <list type="table">
''' <item> 
''' <term> Renamed entries </term>
''' <description> do not exist in the new file, but their content exists in some other entry in the file,
''' mostly unchanged from the old version (may contain minor changes) </description>
''' </item>
''' <item>
''' <term> Merged entries </term>
''' <description> do not exist in the new file, but their content exists in some other entry in the new file 
''' which is substantially different from the old version </description>
''' </item>
''' <item>
''' <term> Removed without replacement </term>
''' <description> do not exist in the new file and their content was not found in some other entry in the new file  </description>
''' </item>
''' </list>
''' <br /> 
''' <br /> Likewise, Merged entries themselves are broken into two categories 
''' <list type="table">
''' <item>
''' <term> Modified </term>
''' <description> Entries that existed in the old file which have been modified to contain content from entries which have been removed  </description>
''' </item>
''' <item>
''' <term> Added </term>
''' <description> Entries which did not exist in the old file but who contain content from entries which have been removed  </description>
''' </item>
''' </list>
''' </summary>
''' 
''' Docs last updated: 2023-06-10
Module Diff

    ''' <summary> 
    ''' The "old" version of winapp2.ini, against which <c> DiffFile2 </c> will be compared 
    ''' <br /> When downloading, this is the local version of winapp2.ini
    ''' </summary>
    ''' 
    ''' Docs last updated: 2022-06-07 | Code last updated: 2022-07-14
    Public Property DiffFile1 As iniFile = New iniFile(Environment.CurrentDirectory, "winapp2.ini", mExist:=True)

    ''' <summary> 
    ''' The "new" version of winapp2.ini against which <c> DiffFile1 </c> will be compared 
    ''' <br /> When downloading, this is the remote version of winapp2.ini
    ''' </summary>
    ''' 
    ''' Docs last updated: 2022-06-07 | Code last updated: 2022-07-14
    Public Property DiffFile2 As iniFile = New iniFile(Environment.CurrentDirectory, "", mExist:=True)

    ''' <summary> 
    ''' The path for the Diff log 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2022-06-07 | Code last updated: 2022-07-14
    Public Property DiffFile3 As iniFile = New iniFile(Environment.CurrentDirectory, "diff.txt")

    ''' <summary> 
    ''' Indicates that a remote winapp2.ini should be downloaded to use as <c> DiffFile2 </c> 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2022-06-07 | Code last updated: 2022-07-14
    Public Property DownloadDiffFile As Boolean = Not isOffline

    ''' <summary> 
    ''' Indicates that the diff output from winapp2ool's global log should be saved to disk 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2022-06-07 | Code last updated: 2022-07-14
    Public Property SaveDiffLog As Boolean = False

    ''' <summary> 
    ''' Indicates that the module settings have been modified from their defaults 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2022-06-07 | Code last updated: 2022-07-14
    Public Property DiffModuleSettingsChanged As Boolean = False

    ''' <summary> 
    ''' Indicates that the remote (online) should be trimmed for the local system before beginning the Diff 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2022-06-07 | Code last updated: 2022-07-14
    Public Property TrimRemoteFile As Boolean = Not isOffline

    ''' <summary> 
    ''' The number of removed entries whose content was detected in an added entry
    ''' </summary>
    ''' 
    ''' Docs last updated: 2022-06-07 | Code last updated: 2022-07-14
    Public Property RemovedByAdditionCount As Integer = 0

    ''' <summary> 
    ''' The total number of removed entries whose content was detected in added or modified entry 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2022-06-07 | Code last updated: 2022-07-14
    Public Property MergedEntryCount As Integer = 0

    ''' <summary> 
    ''' The number of entries that both been added and also contain keys from entries that were removed 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2022-06-07 | Code last updated: 2022-07-14
    Public Property AddedEntryWithMergerCount As Integer = 0

    ''' <summary> 
    ''' The total number of keys that were added to modified entries 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2022-06-07 | Code last updated: 2022-07-14
    Private Property ModEntriesAddedKeyTotal As Integer = 0

    ''' <summary> 
    ''' The total number of keys that were removed from modified entries 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2022-06-07 | Code last updated: 2022-07-14
    Private Property ModEntriesRemovedKeyTotal As Integer = 0

    ''' <summary>
    ''' The total number of keys that were updated in modified entries 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2023-06-07 | Code last updated: 2022-07-14
    Private Property ModEntriesUpdatedKeyTotal As Integer = 0

    ''' <summary>
    ''' The number of keys determined to have been replaced by updated keys in modified entries 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2023-06-07 | Code last updated: 2022-06-07
    Private Property ModEntriesReplacedByUpdateTotal As Integer = 0

    ''' <summary> 
    ''' Indicates that full entries should be printed in the Diff output. <br/> <br/> Called "verbose mode" in the menu 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2022-07-14 | Code last updated: 2022-07-14
    Public Property ShowFullEntries As Boolean = False

    ''' <summary> 
    ''' Holds the slice of the winapp2ool global log containing the most recent Diff results 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2022-06-08 | Code last updated: 2022-07-14
    Public MostRecentDiffLog As String = ""

    ''' <summary>
    ''' A list of Regex characters that need to be escaped when using Regex.Match
    ''' </summary>
    ''' 
    ''' Docs last updated: 2022-06-08 | Code last updated: 2022-06-08
    Private Property RegexCharsIn As String() = {"*", "+", "{", "}", "[", "]", "$", "(", ")"}

    ''' <summary>
    ''' A list of replacement characters for the RegexCharsIn list
    ''' </summary>
    ''' 
    ''' Docs last updated: 2022-06-08 | Code last updated: 2022-06-08
    Private Property RegexCharsOut As String() = {".*", "\+", "\{", "\}", "\[", "\]", "\$", "\(", "\)"}

    ''' <summary>
    ''' LangSecRef values associated with browsers 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2022-06-08 | Code last updated: 2022-06-08
    Private Property BrowserSecRefs As String() = {"3029", "3006", "3033", "3034", "3027", "3026", "3030", "3001"}

    ''' <summary>
    ''' The names of removed entries determined to have been renamed 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2022-06-08 | Code last updated: 2022-06-08
    Private Property RenamedEntryTracker As New HashSet(Of String)

    ''' <summary>
    ''' The names of removed entries determined to have been merged into new or modified entries
    ''' </summary>
    ''' 
    ''' Docs last updated: 2022-06-08 | Code last updated: 2022-06-08
    Private Property MergedEntryTracker As New HashSet(Of String)

    ''' <summary>
    ''' The names of entries determined to have been modified between versions
    ''' </summary>
    ''' 
    ''' Docs last updated: 2022-06-08 | Code last updated: 2022-06-08
    Private Property ModifiedEntryTracker As New HashSet(Of String)

    ''' <summary>
    ''' The names of entries which appear in the new file but not the old file 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2022-06-08 | Code last updated: 2022-06-08
    Private Property AddedEntryTracker As New HashSet(Of String)

    ''' <summary>
    ''' The names of entries which appear in the old file but not the new file
    ''' </summary>
    ''' 
    ''' Docs last updated: 2022-06-08 | Code last updated: 2022-06-08
    Private Property RemovedEntryTracker As New HashSet(Of String)

    ''' <summary>
    ''' A paired list of old and new entry names determined to have been renamed. 
    ''' <br /> Even indices are new names, the next following odd indices are the associated old name 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2022-06-08 | Code last updated: 2022-06-08
    Private Property RenamedEntryPairs As New List(Of String)

    ''' <summary>
    ''' Associates entries containing merged content with the entries from which the content was merged  
    ''' </summary>
    ''' 
    ''' Docs last updated: 2022-06-08 | Code last updated: 2022-06-08
    Private Property MergeDict As New Dictionary(Of String, List(Of String))

    ''' <summary>
    ''' The values of all modified keys for each entry detected as containing modified keys
    ''' </summary>
    ''' 
    ''' Docs last updated: 2022-06-08 | Code last updated: 2022-06-08
    Private Property ModifiedKeyTracker As New Dictionary(Of String, Dictionary(Of iniKey, keyList))

    ''' <summary>
    ''' The values of all removed keys for each entry detected as having had keys removed 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2022-06-08 | Code last updated: 2022-06-08
    Private Property RemovedKeyTracker As New Dictionary(Of String, keyList)

    ''' <summary>
    ''' The values of all added keys for each entry detected as having had keys added
    ''' </summary>
    ''' 
    ''' Docs last updated: 2022-06-08 | Code last updated: 2022-06-08
    Private Property AddedKeyTracker As New Dictionary(Of String, keyList)

    ''' <summary>
    ''' The list of all possible <c> iniSections </c> that could be a match for any given entry 
    ''' <br /> ie. the list of all entries in the new version determined to have been added or modified 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2022-06-08 | Code last updated: 2022-06-08
    Private Property PotentialMatches As New List(Of iniSection)

    ''' <summary>
    ''' Replacement paths for key values containing items in <c> OldPaths </c>
    ''' </summary>
    ''' 
    ''' Docs last updated: 2022-06-08 | Code last updated: 2022-06-08
    Private Property NewPaths As String() = {"%ProgramData%", "%UserProfile%\AppData\LocalLow", "*",
                                            "%UserProfile%\Pictures", "%UserProfile%\Videos",
                                            "%UserProfile%\Documents", "%UserProfile%\Music"}

    ''' <summary>
    ''' Deprecated values that are no longer used or being phased out of winapp2.ini. 
    ''' <br /> These values will be replaced with the corresponding <c> NewPath </c> value being used instead for the purposes of not triggering a 
    ''' "false positive" when diffing, particularly in the case of V22XXXX to V23XXXX or newer 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2022-06-08 | code last updated: 2022-06-08
    Private Property OldPaths As String() = {"%CommonAppData%", "%LocalLowAppData%", "*.*", "%Pictures%", "%Video%", "%Documents%", "%Music%"}

    ''' <summary>
    ''' File system and registry locations which are considered too vague to be used to establish matching key content across entries on their own
    ''' </summary>
    ''' 
    ''' Docs last updated: 2022-06-08 | code last updated: 2022-06-08
    Private Property Disallowed As New HashSet(Of String) From {"%Documents%\Add-in Express", "%UserProfile%\Desktop", "%LocalAppData%", "%WinDir%\System32",
                                                              "%SystemDrive%", "%WinDir%", "%UserProfile%", "%Documents%", "%CommonAppData%", "%AppData%",
                                                              "%Pictures%", "%Public%", "%Music%", "%Video%", "HKCU\Software\Microsoft\Windows",
                                                              "HKLM\Software\Microsoft\Windows", "HKCU\Software\Microsoft\VisualStudio", "%LocalAppData%\Microsoft\Edge*",
                                                              "HKCU\Software\Opera Software", "HKCU\Software\Vivaldi", "HKCU\Software\BraveSoftware"}

    ''' <summary> 
    ''' Runs a diff using command line arguments, allowing Diff to be called programmatically 
    ''' <list type="table"> 
    ''' <item> -d </item>  
    ''' 
    ''' </list>
    ''' 
    ''' <br /> Valid Diff args: 
    ''' <br /> -d           : download the latest winapp2.ini 
    ''' <br /> -ncc         : download the latest non-ccleaner winapp2.ini (implies -d)
    ''' <br /> -savelog     : save diff.txt to disk on exit 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2022-06-10 | Code last updated: 2022-06-08
    Public Sub HandleCmdLine()

        InitDefaultDiffSettings()
        handleDownloadBools(DownloadDiffFile)

        If DownloadDiffFile Then DiffFile2.Name = "Online winapp2.ini"

        invertSettingAndRemoveArg(SaveDiffLog, "-savelog")
        getFileAndDirParams(DiffFile1, DiffFile2, DiffFile3)

        If DiffFile2.Name.Length <> 0 Then conductDiff()

    End Sub

    ''' <summary> 
    ''' Runs a Diff between a local file and the most recent winapp2.ini from GitHub 
    ''' </summary>
    ''' 
    ''' <param name="firstFile">
    ''' The local winapp2.ini file to diff against the master GitHub copy
    ''' </param>
    ''' 
    ''' Docs last updated: 2022-06-08 | Code last updated: 2022-06-08
    Public Sub DiffRemoteFile(firstFile As iniFile)

        DiffFile1 = firstFile
        DownloadDiffFile = True
        conductDiff()

    End Sub

    ''' <summary> 
    ''' Ensures both files have content before kicking off the Diff and then summarizes the output from the Diff 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2023-05-31 | Code last updated: 2023-05-31
    Public Sub ConductDiff()

        If Not enforceFileHasContent(DiffFile1) Then Return

        If DownloadDiffFile Then

            Dim downloadedIniFile = getRemoteIniFile(winapp2link)
            DiffFile2.Sections = downloadedIniFile.Sections

        End If

        If Not enforceFileHasContent(DiffFile2) Then Return

        If TrimRemoteFile AndAlso DownloadDiffFile Then

            Dim tmp As New winapp2file(DiffFile2)
            Trim.trimFile(tmp)
            DiffFile2.Sections = tmp.toIni.Sections

        End If

        clrConsole()
        LogAndPrint(6, $"Changes between{GetVer(DiffFile1)} and{GetVer(DiffFile2)}", enStrCond:=False, ascend:=True)

        CompareFiles()

        LogPostDiff()

        print(3, pressEnterStr)

        crl()

        MostRecentDiffLog = If(SaveDiffLog, getLogSliceFromGlobal("Beginning diff", "Diff complete"), "hasBeenRun")
        DiffFile3.overwriteToFile(MostRecentDiffLog, SaveDiffLog)
        setHeaderText(If(SaveDiffLog, DiffFile3.Name & " saved", "Diff complete"))

    End Sub

    ''' <summary> 
    ''' Gets the version from winapp2.ini 
    ''' </summary>
    ''' 
    ''' <param name="someFile">
    ''' A winapp2.ini format <c> iniFile </c>
    ''' </param>
    ''' 
    ''' <returns>
    ''' The current version if available
    ''' <br /> "version not given" otherwise
    ''' </returns>
    ''' 
    ''' Docs last updated: 2023-06-08 | Code last updated: 2023-06-08
    Private Function GetVer(someFile As iniFile) As String

        Dim ver = If(someFile.Comments.Count > 0, someFile.Comments(0).Comment.ToString(CultureInfo.InvariantCulture).ToUpperInvariant, "000000")
        Return If(ver.Contains("VERSION"), ver.TrimStart(CChar(";")).Replace("VERSION:", "version"), " version not given")

    End Function

    ''' <summary> Records the summary of the diff results and reports them to the user <br />
    ''' Summary: <br />
    ''' Net entry count change: <br />
    ''' Modified Entries: total <br />
    '''  added keys (if applicable) <br />
    '''  removed keys (if applicable) <br />
    '''  updated keys (if applicable) <br />
    ''' Removed entries: total <br />
    '''  Merged entries: total <br />
    '''  Removed without replacement: total <br />
    ''' Added entries: total <br />
    '''  Added entries with no merged content: total <br />
    '''  Added entries with merged content: total <br />
    '''  Added entries determined renamed: total <br />
    ''' </summary>
    ''' 
    ''' Docs last updated: 2023-05-29 | Code last updated: 2023-05-29
    Private Sub LogPostDiff()

        Dim netDiff = DiffFile2.Sections.Count - DiffFile1.Sections.Count
        Dim totalOldEntriesMerged = MergedEntryCount + RenamedEntryTracker.Count
        Dim totalNewEntriesWithMergedKeys = RenamedEntryTracker.Count + MergedEntryTracker.Count
        Dim newVersionsWithMergeKeys = MergedEntryTracker.Count - AddedEntryWithMergerCount
        Dim oldRemovedNoRepl = RemovedEntryTracker.Count - MergedEntryCount - RenamedEntryTracker.Count
        Dim newNoMerged = AddedEntryTracker.Count - RenamedEntryTracker.Count - AddedEntryWithMergerCount

        Dim netChange = $"Net entry count change: {netDiff}"
        Dim modifiedSummaryOpener = $"Modified entries: {ModifiedEntryTracker.Count}"
        Dim modifiedAdded = $" + {ModEntriesAddedKeyTotal} added keys across {AddedKeyTracker.Count} entries "
        Dim modifiedRemoved = $" - {ModEntriesRemovedKeyTotal} removed keys across {RemovedKeyTracker.Count} entries "
        Dim modifiedUpdated = $" ~ {ModEntriesUpdatedKeyTotal} updated keys replacing {ModEntriesReplacedByUpdateTotal} keys across {ModifiedKeyTracker.Count} entries "
        Dim removedSummary = $"Removed entries: {RemovedEntryTracker.Count}"
        Dim removedMergedAdded = $" @ {MergedEntryCount - RemovedByAdditionCount} removed entries have been merged into {MergedEntryTracker.Count - AddedEntryWithMergerCount} modified entries "
        Dim removedReadded = $" + {RemovedByAdditionCount} removed entries have been merged into {AddedEntryWithMergerCount} added entries"
        Dim removedRenamed = $" & {RenamedEntryTracker.Count } removed entries have been renamed and may contain other minor changes "
        Dim removedNoReplacement = $" - {oldRemovedNoRepl} entries have been removed without replacement"
        Dim added = $"Added entries: {AddedEntryTracker.Count}"
        Dim addedNew = $" + {newNoMerged} new entries without merged content"
        Dim addedMerged = $" @ {AddedEntryWithMergerCount} added entries contain keys merged from {RemovedByAdditionCount} removed entries"
        Dim addedRenamed = $" & {RenamedEntryTracker.Count} added entries are renamed versions of removed entries and may contain other minor changes"

        Dim modifiedEntriesHaveAdditions = ModEntriesAddedKeyTotal > 0
        Dim modEntriesHaveRemovals = ModEntriesRemovedKeyTotal > 0
        Dim modEntriesHaveUpdates = ModEntriesUpdatedKeyTotal > 0
        Dim hasRenames = RenamedEntryTracker.Count > 0
        Dim hasAddedMergers = AddedEntryWithMergerCount > 0
        Dim addedKeysHaveMergedContent = hasRenames OrElse hasAddedMergers

        print(0, Nothing, closeMenu:=True)
        LogAndPrint(4, "Summary", conjoin:=True, ascend:=True, leadr:=True)
        LogAndPrint(0, netChange, colorLine:=True, useArbitraryColor:=True, arbitraryColor:=ConsoleColor.White, indent:=True)
        LogAndPrint(0, modifiedSummaryOpener, colorLine:=True, useArbitraryColor:=True, arbitraryColor:=ConsoleColor.Yellow, indent:=True)
        LogAndPrint(0, modifiedAdded, colorLine:=True, enStrCond:=True, cond:=modifiedEntriesHaveAdditions, indent:=True)
        LogAndPrint(0, modifiedRemoved, colorLine:=True, enStrCond:=False, cond:=modEntriesHaveRemovals, indent:=True)
        LogAndPrint(0, modifiedUpdated, colorLine:=True, useArbitraryColor:=True, arbitraryColor:=ConsoleColor.Yellow, cond:=modEntriesHaveUpdates, indent:=True)
        LogAndPrint(0, removedSummary, colorLine:=True, indent:=True)
        LogAndPrint(0, removedMergedAdded, colorLine:=True, useArbitraryColor:=True, arbitraryColor:=ConsoleColor.Cyan, indent:=True, cond:=Not totalOldEntriesMerged = 0)
        LogAndPrint(0, removedReadded, colorLine:=True, enStrCond:=True, indent:=True, cond:=Not RemovedByAdditionCount = 0)
        LogAndPrint(0, removedRenamed, colorLine:=True, useArbitraryColor:=True, arbitraryColor:=ConsoleColor.Magenta, indent:=True, cond:=hasRenames)
        LogAndPrint(0, removedNoReplacement, colorLine:=True, enStrCond:=False, indent:=True)
        LogAndPrint(0, added, colorLine:=True, enStrCond:=True, indent:=True)
        LogAndPrint(0, addedNew, colorLine:=True, enStrCond:=True, closeMenu:=Not addedKeysHaveMergedContent, cond:=newNoMerged > 0, indent:=True)
        LogAndPrint(0, addedMerged, cond:=hasRenames, colorLine:=True, useArbitraryColor:=True, arbitraryColor:=ConsoleColor.DarkCyan, closeMenu:=Not hasRenames, indent:=True)
        LogAndPrint(0, addedRenamed, cond:=hasAddedMergers, colorLine:=True, useArbitraryColor:=True, arbitraryColor:=ConsoleColor.Magenta, closeMenu:=True, indent:=True)
        gLog("", descend:=True)
        gLog("Diff complete", descend:=True)

    End Sub

    ''' <summary>
    ''' Checks the equivalence of two strings using regex. Regex matches must capture the 0th character in the old string to be considered equivalent. 
    ''' </summary>
    ''' 
    ''' <param name="newVal">
    ''' The new value, assessed to see if it is identical to or captures via regex <c> <paramref name="oldVal"/> </c>
    ''' </param>
    ''' 
    ''' <param name="oldVal">
    ''' The old value, either matching directly or captured via regex by <c> <paramref name="newVal"/> </c>
    ''' </param>
    ''' 
    ''' <returns>
    ''' <c> True </c> if the values are equivalent
    ''' <br/> <c> False </c> otherwise
    ''' </returns>
    ''' 
    ''' Docs last updated: 2023-05-29 | Code last updated: 2023-05-29
    Private Function CompareValues(newVal As String, oldVal As String) As Boolean

        Dim newValHasWildcard = newVal.Contains("*")
        Dim newValIsOnlyWildcard = newVal.Equals(".*", StringComparison.InvariantCultureIgnoreCase)

        If newValIsOnlyWildcard Then Return True

        Dim matched = String.Equals(newVal, oldVal, StringComparison.InvariantCultureIgnoreCase)

        If matched OrElse Not newValHasWildcard Then Return matched

        matched = Regex.IsMatch(oldVal, newVal, RegexOptions.IgnoreCase)

        If Not matched Then Return False

        Dim firstMatch = Regex.Match(oldVal, newVal)

        ' There might be a better way to do this but due to the nature of regex matches, 
        ' we can match a sub string where we might not actually want to.
        ' Eg. a file parameter such as "Media History" would be captured "wrongly" 
        ' by a parameter such as "History*". Let's make sure we captured the starting substring 
        ' with this match to have some assurance that this is really a match 
        matched = oldVal.StartsWith(firstMatch.ToString, StringComparison.InvariantCultureIgnoreCase)

        Return matched

    End Function

    ''' <summary>
    ''' Escapes any regex special characters that may exist within the value of an <c> iniKey </c> object 
    ''' </summary>
    ''' 
    ''' <param name="splitPath"> 
    ''' A <c> DetectFile </c> or <c> FileKey </c> value, split at the backslash character, to be sanitized for regex queries
    ''' </param>
    ''' 
    ''' Docs last updated: 2023-05-29 | Code last updated: 2023-05-29
    Private Sub SanitizeRegex(ByRef splitPath As String())

        For k = 0 To splitPath.Length - 1

            For j = 0 To RegexCharsIn.Length - 1

                If Not splitPath(k).Contains(RegexCharsIn(j)) Then Continue For

                splitPath(k) = splitPath(k).Replace(RegexCharsIn(j), RegexCharsOut(j))

            Next

        Next

    End Sub

    ''' <summary>
    ''' Escapes regex characters within FileKey or DetectFile key values so that they can be compared using regex
    ''' </summary>
    ''' 
    ''' <param name="newkey">
    ''' The new <c> iniKey </c> being assessed
    ''' </param>
    ''' 
    ''' <param name="newKeySplit">
    ''' The value of <c> <paramref name="newkey"/> </c>, split at the backslash ( \ ) character
    ''' </param>
    ''' 
    ''' <param name="oldKeySplit">
    ''' The value of the old <c> iniKey </c> against which <c> <paramref name="newkey"/> </c> is being compared
    ''' </param>
    ''' 
    ''' Docs last updated: 2023-05-31 | Code last updated: 2023-06-06
    Private Sub Sanitize(newkey As iniKey,
                         ByRef newKeySplit As String(),
                         ByRef oldKeySplit As String())


        If Not (newkey.typeIs("DetectFile") OrElse newkey.typeIs("FileKey")) Then Return

        Dim firstWildcardIndex = newkey.Value.IndexOf("*", StringComparison.InvariantCultureIgnoreCase)
        If firstWildcardIndex = -1 Then Return

        Dim pipeIndex = newkey.Value.IndexOf("|", StringComparison.InvariantCultureIgnoreCase)
        Dim wildcardIsBeforePipe = firstWildcardIndex < pipeIndex

        If wildcardIsBeforePipe OrElse pipeIndex = -1 Then

            SanitizeRegex(newKeySplit)
            SanitizeRegex(oldKeySplit)
            Return

        End If

        Dim newFlags = {newKeySplit.Last, oldKeySplit.Last}
        SanitizeRegex(newFlags)
        newKeySplit(newKeySplit.Length - 1) = newFlags(0)
        oldKeySplit(oldKeySplit.Length - 1) = newFlags(1)

    End Sub

    ''' <summary>
    ''' Checks the key value equivalence of two <c> iniKey </c> objects 
    ''' <br /> For most keys, this is a simple string equivalence check 
    ''' <br /> For FileKeys, DetectFiles, and Detects, wildcards and recursive flags are evaluated and equivalence is 
    ''' established if <c> newKey's </c> value captures <c> oldKey's </c>
    ''' </summary>
    ''' 
    ''' <param name="newKey">
    ''' A key from the new version of winapp2.ini 
    ''' </param>
    ''' 
    ''' <param name="oldKey">
    ''' A key from the old version of winapp2.ini 
    ''' </param>
    ''' 
    ''' <param name="matchedFileKeyHasMoreParams">
    ''' Tracking variable indicating that the new version of a FileKey contains a larger number of parameters than the old version
    ''' <br /> If this flag is set to <c> True, </c> this suggests an entry merger has occurred 
    ''' </param>
    ''' 
    ''' <param name="possibleWildCardReduction"> 
    ''' Tracking variable indicating that the new version of a matched FileKey contains a wildcard 
    ''' <br /> If this flag is set to <c> True, </c> this suggests that an entry merger has occurred
    ''' </param>
    ''' 
    ''' <returns>
    ''' <c> True </c> if the key values have equivalence, 
    ''' <br /> <c> False </c> otherwise
    ''' </returns>
    ''' 
    ''' <remarks> 
    ''' FileKeys and DetectFiles support wildcards, so we'll do some sanitization so as to be able to evaluate and match them 
    ''' We'll only do this after we've matched the first path piece (usually an environment variable) to reduce the amount of times we do it 
    ''' </remarks>
    ''' 
    ''' Docs last updated: 2023-06-10 | Code last updated: 2023-06-10
    Private Function CheckKeyValueEquivalence(newKey As iniKey,
                                              oldKey As iniKey,
                                              Optional ByRef matchedFileKeyHasMoreParams As Boolean = False,
                                              Optional ByRef possibleWildCardReduction As Boolean = False) As Boolean

        If Not newKey.compareTypes(oldKey) Then Return False

        If String.Equals(newKey.Value, oldKey.Value, StringComparison.InvariantCultureIgnoreCase) Then Return True

        Dim isFileKey = newKey.typeIs("FileKey")
        Dim isDetectFile = newKey.typeIs("DetectFile")
        Dim isDetect = newKey.typeIs("Detect")

        If Not (isFileKey OrElse isDetect OrElse isDetectFile) Then Return False

        Dim newKeySplit = newKey.Value.Split(CChar("\"))
        Dim oldKeySplit = oldKey.Value.Split(CChar("\"))

        If newKeySplit.Length > oldKeySplit.Length Then Return False

        Dim isRecurse = isFileKey AndAlso (newKey.Value.Contains("RECURSE") OrElse newKey.Value.Contains("REMOVESELF"))
        Dim isSanitized = False

        If isFileKey AndAlso Not isRecurse AndAlso newKeySplit.Length < oldKeySplit.Length Then Return False

        For i = 0 To newKeySplit.Length - 1

            If Not isSanitized AndAlso Not isDetect AndAlso (i >= 1 OrElse newKeySplit.Length - 1 = 0) Then

                Sanitize(newKey, newKeySplit, oldKeySplit)
                isSanitized = True

            End If

            Dim newVal = newKeySplit(i)
            Dim oldVal = oldKeySplit(i)
            Dim isLastPiece = i = newKeySplit.Length - 1

            If isLastPiece AndAlso (isDetect OrElse isDetectFile) Then Return CompareValues(newVal, oldVal)

            If isLastPiece AndAlso isFileKey Then Return finalizeFileKeyEquivalence(oldVal, newVal, oldKeySplit, newKeySplit, matchedFileKeyHasMoreParams, possibleWildCardReduction)

            If Not CompareValues(newVal, oldVal) Then Return False

        Next

        Return False

    End Function

    ''' <summary>
    ''' Compares the final parameter components for a pair of FileKeys whose paths have been matched
    ''' </summary>
    ''' 
    ''' <param name="oldVal">
    ''' The value of an old FileKey, possibly captured by <c> <paramref name="newVal"/> </c>
    ''' </param>
    ''' 
    ''' <param name="newVal"> 
    ''' The value of a new FileKey, possibly capturing <c> <paramref name="oldVal"/> </c> 
    ''' </param>
    ''' 
    ''' <param name="oldKeySplit">
    ''' <c> <paramref name="oldVal"/> </c> split at the path delimiter ( \ ) 
    ''' </param>
    ''' 
    ''' <param name="newKeySplit">
    ''' <c> <paramref name="newVal"/> </c> split at the path delimiter ( \ ) 
    ''' </param>
    ''' 
    ''' <param name="matchedFileKeyHasMoreParams">
    ''' Tracking variable indicating that the number of parameters in the new key is larger than the number of parameters in the old key
    ''' </param> 
    ''' 
    ''' <param name="possibleWildCardReduction">
    ''' Tracking variable indicating that the number of parameters may have been reduced through use of a wildcard 
    ''' </param> 
    ''' 
    ''' <returns>
    ''' <c> True </c> if the paths match and at least one of the parameters in <c> <paramref name="newVal"/> </c> captures a parameter from <c> <paramref name="oldVal"/> </c>
    ''' <br /> <c> False </c> otherwise 
    ''' </returns>
    ''' 
    ''' Docs last updated: 2023-06-10 | Code last updated: 2023-06-10
    Private Function FinalizeFileKeyEquivalence(oldVal As String,
                                                newVal As String,
                                                oldKeySplit As String(),
                                                newKeySplit As String(),
                                                ByRef matchedFileKeyHasMoreParams As Boolean,
                                                ByRef possibleWildCardReduction As Boolean) As Boolean

        Dim pipe = CChar("|")
        Dim semi = CChar(";")
        Dim oldValFinal = oldVal.Split(pipe)(0)
        Dim newValFinal = newVal.Split(pipe)(0)
        Dim flags = newKeySplit.Last.Split(pipe)(1)
        Dim oldFlags = oldKeySplit.Last.Split(pipe)(1)

        If Not CompareValues(newValFinal, oldValFinal) Then Return False

        If CompareValues(flags, oldFlags) Then Return True

        If Not (flags.Contains(semi) OrElse oldFlags.Contains(semi)) Then Return False

        Return MatchParameters(flags, oldFlags, matchedFileKeyHasMoreParams, possibleWildCardReduction)

    End Function

    ''' <summary>
    ''' Confirms that the any two parameters for a pair of FileKeys match
    ''' </summary>
    ''' 
    ''' <param name="flags"> 
    ''' The parameters (everything after the pipe symbol ( | ) for a FileKey in the new file 
    ''' </param>
    ''' 
    ''' <param name="oldFlags">
    ''' The parameters (everything after the pipe symbol ( | ) for a FileKey in the old file  
    ''' </param>
    ''' 
    ''' <param name="matchedFileKeyHasMoreParams"> 
    ''' Tracking variable indicating that the number of parameters in the new key is larger than the number of parameters in the old key 
    ''' </param>
    ''' 
    ''' <param name="possibleWildCardReduction"> 
    ''' Tracking variable indicating that the number of parameters may have been reduced through use of a wildcard 
    ''' </param>
    ''' 
    ''' <returns> 
    ''' <c> True </c> if any of the items in <c> <paramref name="flags"/> </c> capture an item in <c> <paramref name="oldFlags"/> ></c> 
    ''' <br /> <c> False </c> otherwise 
    ''' </returns>
    ''' 
    ''' <remarks> 
    ''' <c> <paramref name="matchedFileKeyHasMoreParams"/> </c> and <c> <paramref name="possibleWildCardReduction"/> </c>
    ''' are used to take note of how the number of parameters may have changed. This is to help match renames vs. mergers. 
    ''' <br /> Likely only particularly useful for V22XXXX -> V23XXXX+ Diffs for now 
    ''' </remarks>
    ''' 
    ''' Docs last updated: 2023-06-10 | Code last updated: 2023-06-10
    Public Function MatchParameters(flags As String,
                                    oldFlags As String,
                                    ByRef matchedFileKeyHasMoreParams As Boolean,
                                    ByRef possibleWildCardReduction As Boolean) As Boolean

        Dim delimiter = CChar(";")
        Dim splitParams = flags.Split(delimiter)
        Dim oldSplitParams = oldFlags.Split(delimiter)

        For Each param In splitParams

            For Each oldParam In oldFlags.Split(delimiter)

                If Not CompareValues(param, oldParam) Then Continue For

                matchedFileKeyHasMoreParams = splitParams.Length > oldSplitParams.Length
                possibleWildCardReduction = splitParams.Length < oldSplitParams.Length AndAlso flags.Contains("*") AndAlso Not oldFlags.Contains("*")
                Return True

            Next

        Next

        Return False

    End Function

    ''' <summary>
    ''' Places an entry into the provided tracker and adds it to the potential matches list
    ''' </summary>
    ''' 
    ''' <param name="tracker">
    ''' A tracking list for a particular category of change (added or modified)
    ''' </param>
    ''' 
    ''' <param name="entry">
    ''' A winapp2.ini entry which has been either added or modified between versions 
    ''' </param>
    ''' 
    ''' Docs last updated: 2023-06-10 | Code last updated: 2023-06-10
    Private Sub TrackEntry(ByRef tracker As HashSet(Of String),
                           entry As String)

        tracker.Add(entry)
        PotentialMatches.Add(DiffFile2.Sections(entry))

    End Sub

    ''' <summary> 
    ''' Compares two winapp2.ini format <c> iniFiles </c>, itemizes, and summarizes the differences to the user 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2023-06-10 | Code last updated: 2023-06-10
    Private Sub CompareFiles()

        ModEntriesAddedKeyTotal = 0
        ModEntriesRemovedKeyTotal = 0
        ModEntriesUpdatedKeyTotal = 0
        ModEntriesReplacedByUpdateTotal = 0
        MergedEntryCount = 0
        AddedEntryWithMergerCount = 0
        RemovedByAdditionCount = 0
        RenamedEntryTracker.Clear()
        MergedEntryTracker.Clear()
        ModifiedEntryTracker.Clear()
        AddedEntryTracker.Clear()
        RemovedEntryTracker.Clear()
        MergeDict.Clear()
        ModifiedKeyTracker.Clear()
        RemovedKeyTracker.Clear()
        AddedKeyTracker.Clear()
        PotentialMatches.Clear()

        SnuffNoisyChanges(DiffFile1)
        SnuffNoisyChanges(DiffFile2)

        Dim start = Now

        ProcessNewEntries()

        ProcessOldEntries()

        processRemovals()

        SummarizeRenames()

        SummarizeMergers()

        itemizeMergers()

        itemizeModifications()

        itemizeAdditions()

        Dim ender = Now

        Dim timeSpan = ender - start

        gLog(timeSpan.ToString)

    End Sub

    ''' <summary>
    ''' Records each removed entry from the old version which has been merged into an entry in the new version 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2023-06-05 | Code last updated: 2023-06-05
    Private Sub SummarizeMergers()

        For Each entry In MergeDict

            For Each oldEntry In entry.Value

                MergedEntryCount += 1
                MakeDiff(DiffFile1.Sections(oldEntry), 4, DiffFile2.Sections(entry.Key))

            Next

        Next

    End Sub

    ''' <summary>
    ''' Records each removed entry from the old version which has been given a new name in the new version
    ''' </summary>
    ''' 
    ''' Docs last updated: 2023-06-05 | Code last updated: 2023-06-05
    Private Sub SummarizeRenames()

        For Each entry In RenamedEntryTracker

            Dim ind = RenamedEntryPairs.IndexOf(entry)
            Dim oldName = RenamedEntryPairs(ind + 1)
            MakeDiff(DiffFile1.Sections(oldName), 3, DiffFile2.Sections(entry))

        Next

    End Sub

    ''' <summary>
    ''' Processes the entries in the new file and records any which don't appear in the old file 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2023-06-05 | Code last updated: 2023-06-05
    Private Sub ProcessNewEntries()

        For Each section In DiffFile2.Sections.Values

            If DiffFile1.Sections.ContainsKey(section.Name) Then Continue For

            trackEntry(AddedEntryTracker, section.Name)

        Next

    End Sub

    ''' <summary>
    ''' Processes the entries in the old file. If an entry is not present in the new file, it is recorded as removed. 
    ''' If an entry is present in both files, it is compared against the new version for modifications.
    ''' </summary>
    ''' 
    ''' Docs last updated: 2023-06-05 | Code last updated: 2023-06-05
    Private Sub ProcessOldEntries()

        For Each section In DiffFile1.Sections.Values

            If Not DiffFile2.Sections.ContainsKey(section.Name) Then

                RemovedEntryTracker.Add(section.Name)
                Continue For

            End If

            findModifications(section, DiffFile2.Sections(section.Name))

        Next

    End Sub

    ''' <summary>
    ''' Adds any unique values from an entry in the old file which has been merged into a new file into the <c> <paramref name="entryBuilder"/> </c> 
    ''' used to construct an entry representing the combination of multiple merged entries which will be diffed against the new version 
    ''' </summary>
    ''' 
    ''' <param name="uniqueKeyValues">
    ''' The key values already observed as existing between some set of entries
    ''' </param>
    ''' 
    ''' <param name="entryBuilder">
    ''' A combined entry representing the combination of at least 2 keys, built out of only the unique keys from each 
    ''' </param>
    ''' 
    ''' <param name="oldEntry"> 
    ''' An entry which has been merged and will have its unique keys diffed against the new entry 
    ''' </param>
    ''' 
    ''' Docs last updated: 2023-06-05 | Code last updated: 2023-06-05
    Private Sub BuildMergedOldEntry(ByRef uniqueKeyValues As HashSet(Of String),
                                  ByRef entryBuilder As List(Of String),
                                  oldEntry As iniSection)

        For Each key In oldEntry.Keys.Keys

            If uniqueKeyValues.Contains(key.Value) Then Continue For

            entryBuilder.Add(key.toString)
            uniqueKeyValues.Add(key.Value)

        Next

    End Sub

    ''' <summary>
    ''' Conducts a Diff of each entry detected as containing merged content <br />
    ''' The old entries are collated into a single entry, and then compared against the new entry for modifications
    ''' </summary>
    ''' 
    ''' Docs last updated: 2023-06-05 | Code last updated: 2023-06-05
    Private Sub ItemizeMergers()

        For Each entry In MergedEntryTracker

            Dim uniqueKeyValues = New HashSet(Of String)
            Dim entryBuilder As New List(Of String) From {$"[{entry}]"}
            If DiffFile1.Sections.ContainsKey(entry) Then MergeDict(entry).Add(entry)

            For i = 0 To MergeDict(entry).Count - 1

                Dim oldEnt = DiffFile1.Sections(MergeDict(entry)(i))

                BuildMergedOldEntry(uniqueKeyValues, entryBuilder, oldEnt)

            Next

            Dim mergedOldEntries = New iniSection(entryBuilder)
            findModifications(mergedOldEntries, DiffFile2.Sections(entry))

        Next

        ItemizeModifications(True)

    End Sub

    ''' <summary>
    ''' Processes the entries determined to have been "Removed" and categorizes them into 3 bins: <br />
    ''' Entries which have been renamed: all FileKeys / RegKeys match, but there may be minor changes <br />
    ''' Entries which have been merged: all or most FileKeys / RegKeys / Detects / DetectFiles match, but there may be major changes <br />
    ''' Entries which have actually been removed: key values not found in any new entries <br />
    ''' </summary>
    ''' 
    ''' Docs last updated: 2023-06-05 | Code last updated: 2023-06-05
    Private Sub ProcessRemovals()

        For Each entry In RemovedEntryTracker

            Dim oldSectionVersion = DiffFile1.getSection(entry)

            Dim probableMatches = FindProbableMatches(entry.Split(CChar(" ")), entry)

            Dim changesRecorded = AssessRenamesAndMergers(probableMatches, oldSectionVersion)

            If changesRecorded Then Continue For

            changesRecorded = AssessRenamesAndMergers(PotentialMatches, oldSectionVersion)

            If Not changesRecorded Then MakeDiff(oldSectionVersion, 1)

        Next

    End Sub

    ''' <summary>
    ''' Outputs each added entry and any entries which have been merged into it 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2023-06-05 | Code last updated: 2023-06-05
    Private Sub ItemizeAdditions()

        Dim lastEntryHadMergers = False

        For Each entry In AddedEntryTracker.ToList

            If RenamedEntryTracker.Contains(entry) Then Continue For

            Dim section = DiffFile2.Sections(entry)

            print(0, "", cond:=Not lastEntryHadMergers AndAlso MergeDict.ContainsKey(section.Name))

            MakeDiff(section, 0)

            lastEntryHadMergers = MergeDict.ContainsKey(section.Name)

            If Not lastEntryHadMergers Then Continue For

            AddedEntryWithMergerCount += 1
            Dim out = "This entry contains keys merged from the following removed entries:"
            LogAndPrint(0, out, isCentered:=True, colorLine:=True, useArbitraryColor:=True, arbitraryColor:=ConsoleColor.Yellow, indent:=True)

            For Each mergedEntry In MergeDict(section.Name)

                RemovedByAdditionCount += 1
                LogAndPrint(0, mergedEntry, isCentered:=True, colorLine:=True, useArbitraryColor:=True, arbitraryColor:=ConsoleColor.DarkCyan, indent:=True, indAmt:=4)

            Next

            print(0, "", fillBorder:=False)

        Next

    End Sub

    ''' <summary>
    ''' Itemizes the ways in which a given entry has been modified and outputs them to the user
    ''' </summary>
    ''' 
    ''' <param name="isMerger"> 
    ''' Indicates that the current set of entries which have been modified are the product of merging multiple entries together 
    ''' </param>
    ''' 
    ''' Docs last updated: 2023-06-10 | Code last updated: 2023-06-10
    Private Sub ItemizeModifications(Optional isMerger As Boolean = False)

        For Each entry In ModifiedEntryTracker.ToList

            If Not isMerger AndAlso MergedEntryTracker.Contains(entry) Then Continue For
            If isMerger AndAlso Not MergedEntryTracker.Contains(entry) Then Continue For

            Dim modCounts, addCounts, remCounts As New List(Of Integer)
            Dim modKeyTypes, addKeyTypes, remKeyTypes As New List(Of String)
            Dim newSectionVer = DiffFile2.Sections(entry)
            Dim addedKeys = If(AddedKeyTracker.ContainsKey(entry), AddedKeyTracker(entry), New keyList)
            Dim removedKeys = If(RemovedKeyTracker.ContainsKey(entry), RemovedKeyTracker(entry), New keyList)

            Dim updatedKeysDict = If(ModifiedKeyTracker.ContainsKey(entry), ModifiedKeyTracker(entry), New Dictionary(Of iniKey, keyList))

            If removedKeys.KeyCount + addedKeys.KeyCount + updatedKeysDict.Count = 0 Then Continue For

            MakeDiff(newSectionVer, 2)
            ItemizeChangesFromList(addedKeys, True, addKeyTypes, addCounts)
            ItemizeChangesFromList(removedKeys, False, remKeyTypes, remCounts)

            ItemizeUpdatedKeys(updatedKeysDict, addedKeys, removedKeys, modKeyTypes, modCounts)

            ModEntriesAddedKeyTotal += addCounts.Sum
            ModEntriesRemovedKeyTotal += remCounts.Sum
            ModEntriesUpdatedKeyTotal += modCounts.Sum

            ItemizeMergedEntries(entry, isMerger)

        Next

    End Sub

    ''' <summary>
    ''' Outputs the changes made to the keys within an entry to the user 
    ''' </summary>
    ''' 
    ''' <param name="updatedKeysDict">
    ''' The set of updated keys for a particular entry, where the key is the new version of an updated iniKey and the value is the list of old keys it captures 
    ''' </param>
    ''' 
    ''' <param name="addedKeys"> 
    ''' The set of keys newly added to the entry 
    ''' </param>
    ''' 
    ''' <param name="removedKeys">
    ''' The set of keys which have been removed from the entry 
    ''' </param>
    ''' 
    ''' <param name="modKeyTypes">
    ''' Tracking variable recording which KeyTypes have been modified for the purposes of the user summary
    ''' </param>
    ''' 
    ''' <param name="modCounts">
    ''' Tracking variable recording the number of each type of modification detected for the purposes of the user summary
    ''' </param>
    '''
    ''' Docs last updated: 2023-06-10 | Code last updated: 2023-06-10
    Private Sub ItemizeUpdatedKeys(updatedKeysDict As Dictionary(Of iniKey, keyList),
                                   addedKeys As keyList,
                                   removedKeys As keyList,
                                   modKeyTypes As List(Of String),
                                   modCounts As List(Of Integer))

        If updatedKeysDict.Count = 0 Then Return

        gLog("", ascend:=True, cond:=addedKeys.KeyCount + removedKeys.KeyCount = 0)

        For Each changeList In updatedKeysDict.Values

            recordModification(modKeyTypes, modCounts, changeList.Keys(0).KeyType)

        Next

        summarizeEntryUpdate(modKeyTypes, modCounts, "Modified")

        For i = 0 To updatedKeysDict.Count - 1

            Dim newKey = updatedKeysDict.Keys(i)
            Dim oldKeys = updatedKeysDict.Values(i).Keys
            Dim isRename = newKey.typeIs("Name")
            Dim count = updatedKeysDict.Values(i).Keys.Count

            Dim outTxt1 = $"{If(isRename, "Entry Name", newKey.Name)} has been modified{If(Not isRename, $", replacing {count} old key{If(count > 1, "s", "")}", "")}"
            LogAndPrint(0, outTxt1, colorLine:=True, useArbitraryColor:=True, arbitraryColor:=ConsoleColor.Yellow, indent:=True, indAmt:=1, leadr:=i = 0)

            Dim outTxt2 = $"New: {If(isRename, newKey.Name, newKey.toString)}"
            LogAndPrint(0, outTxt2, colorLine:=True, enStrCond:=True, indent:=True, indAmt:=4)

            For Each oldKey In oldKeys

                Dim old = $"Old: {If(isRename, oldKey.toString.TrimEnd(CChar("=")), oldKey.toString)}"
                LogAndPrint(0, old, colorLine:=True, indent:=True, indAmt:=4)

            Next

            LogAndPrint(0, Nothing, logCond:=False, fillBorder:=False)

        Next

        gLog(descend:=True, cond:=addedKeys.KeyCount + removedKeys.KeyCount = 0)

    End Sub

    ''' <summary>
    ''' Itemizes the names of any removed entries that were merged into <c> <paramref name="entry"/> </c>
    ''' </summary>
    ''' 
    ''' <param name="entry">
    ''' The name of an entry which contains content 
    ''' </param>
    ''' 
    ''' <param name="isMerger">
    ''' Indicates that the current Diff is operating on an entry which is "Merged" and not "New" 
    ''' </param>
    ''' 
    ''' Docs last updated: 2023-06-10 | Code last updated: 2023-06-10
    Private Sub ItemizeMergedEntries(entry As String,
                                       isMerger As Boolean)

        If Not MergeDict.ContainsKey(entry) Then Return

        Dim outTxt = If(Not isMerger, "This entry contains keys merged from the following removed entries", "The above changes are measured against the following removed/old entries")
        LogAndPrint(0, outTxt, isCentered:=True, colorLine:=True, useArbitraryColor:=True, arbitraryColor:=ConsoleColor.Yellow, indAmt:=1, indent:=True)

        For Each mergedEntry In MergeDict(entry)

            LogAndPrint(0, mergedEntry, isCentered:=True, colorLine:=True, useArbitraryColor:=True, arbitraryColor:=ConsoleColor.DarkCyan, indent:=True)

        Next

        print(0, Nothing, fillBorder:=False)

    End Sub

    ''' <summary>
    ''' In the rare event that an entry has its classification changed from "renamed" to "merged" through the Diff process, this function handles rolling back the 
    ''' record keeping associated with the first classification 
    ''' </summary>
    ''' 
    ''' <param name="sectionName"> 
    ''' The new of an entry whose classification is being updated from "renamed" to "merged"
    ''' </param>
    ''' 
    ''' Docs last updated: 2023-06-10 | Code last updated: 2023-06-10
    Private Sub RollBackPreviouslyObservedChanges(sectionName As String)

        If AddedKeyTracker.ContainsKey(sectionName) Then ModEntriesAddedKeyTotal -= AddedKeyTracker(sectionName).KeyCount
        If RemovedKeyTracker.ContainsKey(sectionName) Then ModEntriesRemovedKeyTotal -= RemovedKeyTracker(sectionName).KeyCount
        If ModifiedKeyTracker.ContainsKey(sectionName) Then ModEntriesUpdatedKeyTotal -= ModifiedKeyTracker(sectionName).Count

        If MergeDict.ContainsKey(sectionName) Then ModEntriesReplacedByUpdateTotal -= MergeDict(sectionName).Count

        ModifiedEntryTracker.Remove(sectionName)
        AddedKeyTracker.Remove(sectionName)
        RemovedKeyTracker.Remove(sectionName)
        ModifiedKeyTracker.Remove(sectionName)

    End Sub


    ''' <summary>
    ''' Determines the changes made to the <c> iniKey </c> values in an <c> iniSection </c> that has been updated between versions
    ''' </summary>
    ''' 
    ''' <param name="newSection"> 
    ''' An entry from the new version of winapp2.ini 
    ''' </param>
    ''' 
    ''' <param name="oldSection"> 
    ''' One or more entries from the old version of winapp2.ini as a single unit 
    ''' </param>
    ''' 
    ''' <remarks> 
    ''' In the case of tracking multiple mergers, <c> <paramref name="oldSection"/> </c> isn't a single no-longer-extant entry, but instead 
    ''' is a combination of all the keys from all the merged entries. 
    ''' </remarks>
    ''' 
    ''' Docs last updated: 2023-06-13 | Code last updated: 2023-06-13
    Private Sub FindModifications(oldSection As iniSection,
                                  newSection As iniSection)

        Dim addedKeys, removedKeys As New keyList

        If oldSection.compareTo(newSection, removedKeys, addedKeys) Then Return

        If ModifiedEntryTracker.Contains(newSection.Name) Then RollBackPreviouslyObservedChanges(newSection.Name)

        Dim updatedKeys = DetermineModifiedKeys(removedKeys, addedKeys)
        If removedKeys.KeyCount + addedKeys.KeyCount + updatedKeys.Count = 0 Then Return

        If removedKeys.KeyCount > 0 Then RemovedKeyTracker.Add(newSection.Name, removedKeys)
        If addedKeys.KeyCount > 0 Then AddedKeyTracker.Add(newSection.Name, addedKeys)

        TrackEntry(ModifiedEntryTracker, newSection.Name)

        If Not oldSection.Name.Equals(newSection.Name, StringComparison.InvariantCultureIgnoreCase) Then

            Dim oldName = New iniKey(oldSection.Name) With {.KeyType = "Name"}
            Dim newName = New iniKey(newSection.Name) With {.KeyType = "Name"}
            updatedKeys.Add(New KeyValuePair(Of iniKey, iniKey)(newName, oldName))

        End If

        If updatedKeys.Count > 0 Then ModifiedKeyTracker(newSection.Name) = BuildModifications(updatedKeys)

    End Sub

    ''' <summary>
    ''' Builds a dictionary which contains as its keys each updated key from the new version of winapp2.ini and as the values a list of the old 
    ''' keys which were replaced by the new key
    ''' </summary>
    ''' 
    ''' <param name="updatedKeys">
    ''' A list of updated keys and the old keys determined to have been supereceded by them 
    ''' </param>
    ''' 
    ''' <returns> 
    ''' A <c> Dictionary (Of iniKey, KeyList) </c> where each <c> iniKey </c>is a key in the new version of the file and 
    ''' the <c> keyList </c> contains all of the keys from the old version which have been determined to be captured by the new version 
    ''' </returns>
    ''' 
    ''' Docs last updated: 2023-06-13 | Code last updated: 2023-06-13
    Private Function BuildModifications(ByRef updatedKeys As List(Of KeyValuePair(Of iniKey, iniKey))) As Dictionary(Of iniKey, keyList)

        Dim modifications As New Dictionary(Of iniKey, keyList)

        For Each kvpair In updatedKeys

            If Not modifications.ContainsKey(kvpair.Key) Then modifications.Add(kvpair.Key, New keyList)

            modifications(kvpair.Key).add(kvpair.Value)

            ModEntriesReplacedByUpdateTotal += 1

        Next

        Return modifications

    End Function

    ''' <summary>
    ''' Produces a list of iniSections who may potentially be merger/rename candidates based on traits such as section and name similarities 
    ''' </summary>
    ''' 
    ''' <param name="oldNameBroken"> 
    ''' The old name of the entry broken into pieces about the space character, ie. each "word" in the name 
    ''' </param>
    ''' 
    ''' <param name="entry">
    ''' The name of the new entry 
    ''' </param>
    ''' 
    ''' <returns> 
    ''' A <c> List(Of iniSection) </c> either falling into the same <c> LangSecRef </c> or matching a component of the name of the "old" iniSection 
    ''' </returns>
    ''' 
    ''' Docs last updated: 2023-06-13 | Code last updated: 2023-06-13
    Private Function FindProbableMatches(oldNameBroken As String(),
                                         entry As String) As List(Of iniSection)

        Dim out = New List(Of iniSection)

        For Each newName In PotentialMatches

            Dim upperNewName = newName.Name.ToUpperInvariant()

            Dim matched = False
            Dim newVer = newName.ToString.ToUpperInvariant

            For Each browser In BrowserSecRefs

                If Not newVer.Contains(browser) Then Continue For

                Dim oldVer = DiffFile1.Sections(entry).ToString
                If Not oldVer.Contains(browser) Then Continue For

                out.Add(newName)
                matched = True
                Exit For

            Next

            If matched Then Continue For

            For Each oldNamePiece In oldNameBroken

                ' This'll be the last entry in any given list here 
                If String.Equals(oldNamePiece, "*", StringComparison.InvariantCultureIgnoreCase) Then Exit For

                If Not upperNewName.Contains($"{oldNamePiece.ToUpperInvariant} ") Then Continue For
                out.Add(newName)
                Exit For

            Next

        Next

        Return out

    End Function

    ''' <summary>
    ''' Records the candidate for a merger or rename based on the number of keys in the entry and the number of keys 
    ''' which match between the old and new versions of the entry
    ''' </summary>
    ''' 
    ''' <param name="count">
    ''' The number of matching keys between two entries 
    ''' </param>
    ''' 
    ''' <param name="newSectionName">
    ''' An entry who is a potential candidate for having been the product of a merger 
    ''' </param>
    ''' 
    ''' <param name="highestCount">
    ''' The highest currently observed number of matching keys between two entries
    ''' </param>
    ''' 
    ''' <param name="newNameCandidate">
    ''' The entry name associated with the <c> <paramref name="highestCount"/> </c>
    ''' </param>
    ''' 
    ''' <param name="mergeTracker">
    ''' Tracking variable indicating that a merger has been determined as having likely taken place 
    ''' </param>
    ''' 
    ''' Docs last updated: 2023-06-13 | Code last updated: 2023-06-13
    Private Sub SetMergeCandidate(count As Integer,
                                  newSectionName As String,
                                  ByRef highestCount As Integer,
                                  ByRef newNameCandidate As String,
                                  ByRef mergeTracker As Boolean)

        If count <= highestCount Then Return

        newNameCandidate = newSectionName
        highestCount = count
        mergeTracker = True

    End Sub

    ''' <summary>
    ''' Determines into which extant entry a removed entry may have been merged into. If all the deletion and detection criteria match 
    ''' without a substantial change to the number or parameters of the keys, the entry is considered "renamed." otherwise, if the keys 
    ''' match but there is some other mismatch in the number of keys or parameters, the entry is considered "merged" 
    ''' </summary>
    ''' 
    ''' <param name="entriesAddedOrModified">
    ''' The set of all entries of interest that have been added or changed in the new version of winapp2.ini 
    ''' </param>
    ''' 
    ''' <param name="oldSectionVersion"> 
    ''' An old (removed) winapp2.ini entry
    ''' </param>
    ''' 
    ''' Docs last updated: 2023-05-30 | Code last updated: 2023-05-30

    Private Function AssessRenamesAndMergers(ByRef entriesAddedOrModified As List(Of iniSection),
                                             ByRef oldSectionVersion As iniSection) As Boolean

        Dim highestMatchCount = 0
        Dim newMergedOrRenamedName = ""
        Dim entryWasRenamedOrMerged = False

        For Each section In entriesAddedOrModified

            If Not AddedEntryTracker.Contains(section.Name) AndAlso Not ModifiedEntryTracker.Contains(section.Name) Then Continue For

            Dim allFileKeysMatched = False
            Dim allRegKeysMatched = False
            Dim regKeyCountsMatch = False
            Dim fileKeyCountsMatch = False
            Dim newWa2sSection As New winapp2entry(section)
            Dim oldWa2Section As New winapp2entry(oldSectionVersion)
            Dim fileKeyMatches = 0
            Dim regKeyMatches = 0
            Dim matchHadMoreParams = False
            Dim possibleWildCardReduction = False
            Dim newSectionName = section.Name

            assessKeyMatches(oldWa2Section.FileKeys, newWa2sSection.FileKeys, fileKeyCountsMatch,
                              allFileKeysMatched, fileKeyMatches, matchedFileKeyHasMoreParams:=matchHadMoreParams, possibleWildCardReduction:=possibleWildCardReduction)

            assessKeyMatches(oldWa2Section.RegKeys, newWa2sSection.RegKeys, regKeyCountsMatch,
                             allRegKeysMatched, regKeyMatches, Disallowed)

            setMergeCandidate(fileKeyMatches + regKeyMatches, newSectionName, highestMatchCount, newMergedOrRenamedName, entryWasRenamedOrMerged)

            If fileKeyMatches = 0 AndAlso regKeyMatches = 0 Then Continue For

            Dim hasNotBeenObserved = Not (RenamedEntryTracker.Contains(newMergedOrRenamedName) OrElse MergedEntryTracker.Contains(newMergedOrRenamedName))
            Dim countsMatched = fileKeyCountsMatch AndAlso regKeyCountsMatch
            Dim allKeysMatched = allFileKeysMatched AndAlso allRegKeysMatched
            Dim paramsUnchanged = Not (matchHadMoreParams OrElse possibleWildCardReduction)
            Dim entryWasRenamed = hasNotBeenObserved AndAlso countsMatched AndAlso paramsUnchanged

            Dim changeWasRecorded = allKeysMatched AndAlso ConfirmRenameOrMerger(newMergedOrRenamedName, entryWasRenamed, oldSectionVersion)
            If changeWasRecorded Then Return True

            ' Just skip this next check on the browser sections, they all share the same detections 
            If oldWa2Section.LangSecRef.KeyCount > 0 AndAlso BrowserSecRefs.Contains(oldWa2Section.LangSecRef.Keys(0).Value) Then Continue For

            Dim detectMatches = 0
            Dim detectFileMatches = 0

            assessKeyMatches(oldWa2Section.Detects, newWa2sSection.Detects, True, False, detectMatches, Disallowed)
            assessKeyMatches(oldWa2Section.DetectFiles, newWa2sSection.DetectFiles, True, False, detectFileMatches, Disallowed)
            setMergeCandidate(fileKeyMatches + regKeyMatches + detectFileMatches + detectMatches, newSectionName, highestMatchCount, newMergedOrRenamedName, entryWasRenamedOrMerged)

        Next

        If Not entryWasRenamedOrMerged Then Return False

        trackMerger(oldSectionVersion, DiffFile2.Sections(newMergedOrRenamedName))
        Return True

    End Function


    ''' <summary>
    ''' Tracks a merger or finds modifications between two entries as appropriate 
    ''' </summary>
    ''' 
    ''' <param name="newMergedOrRenamedName">
    ''' The name of an entry in the new version of winapp2.ini which has been determined to be the product of a merger or rename
    ''' </param>
    ''' 
    ''' <param name="entryWasRenamed">
    ''' Tracking variable indicating that Diff has determined <c> <paramref name="oldSectionVersion"/> </c> to have been 
    ''' renamed to <c> <paramref name="newMergedOrRenamedName"/> </c>
    ''' </param>
    ''' 
    ''' <param name="oldSectionVersion">
    ''' 
    ''' </param>
    ''' 
    ''' <returns>
    ''' <c> True, </c> always
    ''' </returns>
    ''' 
    ''' Docs last updated: 2023-06-13 | Code last updated: 2023-06-13
    Private Function ConfirmRenameOrMerger(newMergedOrRenamedName As String,
                                           entryWasRenamed As Boolean,
                                           oldSectionVersion As iniSection) As Boolean

        Dim section = DiffFile2.getSection(newMergedOrRenamedName)

        If Not entryWasRenamed Then trackMerger(oldSectionVersion, section) : Return True

        RenamedEntryTracker.Add(newMergedOrRenamedName)
        RenamedEntryPairs.Add(newMergedOrRenamedName)
        RenamedEntryPairs.Add(oldSectionVersion.Name)
        FindModifications(oldSectionVersion, section)

        Return True

    End Function

    ''' <summary>
    ''' Tracks which entries have been merged into which other entries. If an entry previously determined to be renamed
    ''' contains merged content, it is removed from the renamed list and considered instead a merged entry. 
    ''' </summary>
    ''' 
    ''' <param name="oldSectionVersion">
    ''' The version of an entry which has been merged into <c> <paramref name="newIniSectionVersion"/> </c>
    ''' </param>
    ''' 
    ''' <param name="newIniSectionVersion"> 
    ''' An <c> iniSection </c> from the new version of the file containing keys merged from <c> <paramref name="oldSectionVersion"/> </c>
    ''' </param>
    ''' 
    ''' Docs last updated: 2023-05-30 | Code last updated: 2023-05-30
    Private Sub trackMerger(oldSectionVersion As iniSection,
                          newIniSectionVersion As iniSection)

        Dim mergeName = newIniSectionVersion.Name
        MergedEntryTracker.Add(mergeName)

        If Not MergeDict.ContainsKey(mergeName) Then MergeDict.Add(mergeName, New List(Of String))

        MergeDict(mergeName).Add(oldSectionVersion.Name)

        If Not RenamedEntryTracker.Contains(mergeName) Then Return

        Dim ind = RenamedEntryPairs.IndexOf(mergeName)
        Dim renameHolder = RenamedEntryPairs(ind + 1)
        MergeDict(mergeName).Add(renameHolder)
        RenamedEntryPairs.RemoveAt(ind + 1)
        RenamedEntryPairs.RemoveAt(ind)
        RenamedEntryTracker.Remove(mergeName)

    End Sub

    ''' <summary> 
    ''' Counts the number of matching key contents between two ini keyLists 
    ''' </summary>
    ''' 
    ''' <param name="oldKeyList"> 
    ''' The list of keys from the "old" version of the entry 
    ''' </param>
    ''' 
    ''' <param name="newKeyList"> 
    ''' The list of keys from the "new" version of the entry
    ''' </param>
    ''' 
    ''' <param name="countTracker"> 
    ''' Tracks the number of matches observed between the two given keyLists 
    ''' </param>
    ''' 
    ''' <param name="allKeysMatchedTracker"> 
    ''' Tracks whether or not all the keys have matched
    ''' </param>
    ''' 
    ''' <param name="MatchCount"> 
    ''' The number of matches recorded 
    ''' </param>
    ''' 
    ''' <param name="disallowedValues"> 
    ''' Any values whose matching should be ignored <br /> Optional, Default: <c> Nothing </c> 
    ''' </param>
    ''' 
    ''' <param name="matchedFileKeyHasMoreParams">
    ''' Indicates that a passed FileKey that matched had more parameters in the new version than the old
    ''' <br /> Optional, Default: <c> False </c> 
    ''' </param>
    ''' 
    ''' Docs last updated: 2023-04-20 | Code last updated: 2023-04-20

    Private Sub assessKeyMatches(oldKeyList As keyList,
                                newKeyList As keyList,
                                ByRef countTracker As Boolean,
                                ByRef allKeysMatchedTracker As Boolean,
                                ByRef MatchCount As Integer,
                                Optional disallowedValues As HashSet(Of String) = Nothing,
                                Optional ByRef matchedFileKeyHasMoreParams As Boolean = False,
                                Optional ByRef possibleWildCardReduction As Boolean = False)


        For Each key In oldKeyList.Keys

            For Each newKey In newKeyList.Keys

                Dim matched = String.Equals(newKey.Value, key.Value, StringComparison.InvariantCultureIgnoreCase)

                If matched AndAlso disallowedValues IsNot Nothing AndAlso disallowedValues.Contains(key.Value) Then Exit For

                If Not matched Then matched = CheckKeyValueEquivalence(newKey, key, matchedFileKeyHasMoreParams, possibleWildCardReduction)

                If matched Then MatchCount += 1

            Next

        Next

        allKeysMatchedTracker = MatchCount = oldKeyList.KeyCount

        ' If the new key list is smaller than the old one but all of the keys in the old one exist in the new one, 
        ' We have a case where an entry was merged into another entry and also had a wildcard key reduction. in this case
        ' we'll consider the key count to be equivalent for the purposes of tracking renames 
        countTracker = allKeysMatchedTracker AndAlso Not newKeyList.KeyCount > MatchCount

    End Sub

    ''' <summary> 
    ''' Records the number of changes made in a modified entry
    ''' </summary>
    ''' 
    ''' <param name="ktList">
    ''' The KeyTypes for the type of change being observed 
    ''' </param>
    ''' 
    ''' <param name="countsList">
    ''' The counts of the changed KeyTypes 
    ''' </param>
    ''' 
    ''' <param name="keyType"> 
    ''' A KeyType from a key that has been changed and whose change will be recorded 
    ''' </param>
    ''' 
    ''' Docs last updated: 2023-06-05 | Code last updated: 2023-06-05
    Private Sub recordModification(ByRef ktList As List(Of String),
                                   ByRef countsList As List(Of Integer),
                                   ByRef keyType As String)

        If Not ktList.Contains(keyType) Then

            ktList.Add(keyType)
            countsList.Add(0)

        End If

        countsList(ktList.IndexOf(keyType)) += 1

    End Sub

    ''' <summary> 
    ''' Creates and outputs summaries for entry modifications <br/>
    ''' eg. "Added 2 FileKeys" or "Removed 1 Detect" 
    ''' </summary>
    ''' 
    ''' <param name="keyTypeList"> 
    ''' The KeyTypes that have been updated 
    ''' </param>
    ''' 
    ''' <param name="countList"> 
    ''' The quantity of keys by KeyType who have been modified 
    ''' </param>
    ''' 
    ''' <param name="changeType">
    ''' The type of change being summarized 
    ''' </param>
    ''' 
    ''' Docs last updated: 2023-05-30 | Code last updated: 2023-05-30
    Private Sub summarizeEntryUpdate(keyTypeList As List(Of String),
                                     countList As List(Of Integer),
                                     changeType As String)

        For i = 0 To keyTypeList.Count - 1

            Dim out = $"{changeType} {countList(i)} {keyTypeList(i)}{If(countList(i) > 1, "s", "")}"
            LogAndPrint(0, out, colorLine:=True, useArbitraryColor:=True, arbitraryColor:=ConsoleColor.Yellow, isCentered:=True, trailingBlank:=i = keyTypeList.Count - 1)

        Next

    End Sub

    ''' <summary> 
    ''' Prints any added or removed keys from an updated entry to the user 
    ''' </summary>
    ''' 
    ''' <param name="kl"> 
    ''' The iniKeys that have been added/removed from an entry 
    ''' </param>
    ''' 
    ''' <param name="wasAdded"> 
    ''' <c> True </c> if keys in <c> <paramref name="kl"/> </c> were added, <c> False </c> otherwise 
    ''' </param>
    ''' 
    ''' <param name="ktList"> 
    ''' The KeyTypes of modified keys
    ''' </param>
    ''' 
    ''' <param name="countList"> 
    ''' The counts of the KeyTypes for modified keys 
    ''' </param>
    ''' 
    ''' Docs last updated: 2023-06-05 | Code last updated: 2023-06-05
    Private Sub ItemizeChangesFromList(kl As keyList,
                                   wasAdded As Boolean,
                                   ByRef ktList As List(Of String),
                                   ByRef countList As List(Of Integer))

        If kl.KeyCount = 0 Then Return

        gLog(ascend:=True)

        Dim changeTxt = If(wasAdded, "Added", "Removed")
        Dim tmpKtList = ktList
        Dim tmpCountList = countList

        kl.Keys.ForEach(Sub(key) recordModification(tmpKtList, tmpCountList, key.KeyType))

        ktList = tmpKtList
        countList = tmpCountList

        summarizeEntryUpdate(ktList, countList, changeTxt)

        For i = 0 To kl.KeyCount - 1

            Dim key = kl.Keys(i).toString
            LogAndPrint(0, key, colorLine:=True, enStrCond:=wasAdded, indent:=True, indAmt:=4)

        Next

        gLog(descend:=True)

    End Sub

    ''' <summary>
    ''' Creates parity between certain key values that have been changed was part of the winapp2.ini v23XXXX changes <br />
    ''' Used to ignore broad cases of text conversion in Diff, such as environmental variables changes introduced as part of 
    ''' the effort to go Non-CCleaner as default winapp2.ini 
    ''' </summary>
    ''' 
    ''' <param name="winapp"> 
    ''' A winapp2.ini format <c> iniFile </c> 
    ''' </param>
    ''' 
    ''' <remarks> 
    ''' These changes overwrite the way the output will be generated and will produce slightly misleading 
    ''' albeit syntactically identical output regarding the contents of the old file <br />
    ''' ie. This will cause Diff to suggest that the new values exist in the old file if the old values did when that's 
    ''' not actually true (but it is what we want because we want to "ignore" these changes)
    ''' </remarks>
    ''' 
    ''' Docs last updated: 2023-06-13 | Code last updated: 2023-06-08
    Private Sub SnuffNoisyChanges(ByRef winapp As iniFile)

        For Each section In winapp.Sections.Values

            For Each key In section.Keys.Keys

                CleanKeyValue(key)

            Next

        Next

    End Sub

    ''' <summary>
    ''' Replaces values in a given <c> iniKey </c> that contain any of the <c> OldPaths </c> with the corresponding <c> NewPaths </c>
    ''' </summary>
    ''' 
    ''' <param name="key">
    ''' An <c> iniKey </c> to be sanitized 
    ''' </param>
    ''' 
    ''' Docs last updated: 2023-06-08 | Code last updated: 2023-06-08
    Private Sub CleanKeyValue(ByRef key As iniKey)

        For k = 0 To NewPaths.Length - 1

            If Not key.Value.Contains(OldPaths(k)) Then Continue For

            key.Value = key.Value.Replace(OldPaths(k), NewPaths(k))

        Next

    End Sub

    ''' <summary>
    ''' Resolves collisions between added and removed keys so as to identify which added keys are modified versions of removed keys, 
    ''' updating the given <c> keyLists </c> accordingly
    ''' </summary>
    ''' 
    ''' <param name="removedKeys"> 
    ''' <c> iniKeys </c> determined to have been removed from the newer version of the <c> iniSection </c> 
    ''' </param>
    ''' 
    ''' <param name="addedKeys"> 
    ''' <c> iniKeys </c> determined to have been added to the newer version of the <c> iniSection </c> 
    ''' </param>
    ''' 
    ''' <returns> 
    ''' A list of <c> iniKeys </c> and their matched partners (in [new, old] KeyValuePairs) for the purpose of identifying "modified" keys 
    ''' </returns>
    ''' Docs last updated: 2023-06-12 | Code last updated: 2023-06-12

    Private Function DetermineModifiedKeys(ByRef removedKeys As keyList,
                                           ByRef addedKeys As keyList) As List(Of KeyValuePair(Of iniKey, iniKey))

        Dim updatedKeys As New List(Of KeyValuePair(Of iniKey, iniKey))
        Dim classifiers = {"LangSecRef", "Section"}
        Dim defunctSingletonKeys = {"Warning", "DetectOS", "SpecialDetect"}

        For Each key In addedKeys.Keys

            Dim newKeyType = key.KeyType

            For Each sKey In removedKeys.Keys

                Dim oldKeyType = sKey.KeyType

                Dim shouldExistOnce = classifiers.Contains(newKeyType) AndAlso classifiers.Contains(oldKeyType) OrElse
                            defunctSingletonKeys.Contains(newKeyType) AndAlso newKeyType = oldKeyType

                If Not (shouldExistOnce OrElse CheckKeyValueEquivalence(key, sKey)) Then Continue For

                updatedKeys.Add(New KeyValuePair(Of iniKey, iniKey)(key, sKey))

            Next

        Next

        For Each pair In updatedKeys

            addedKeys.remove(pair.Key)
            removedKeys.remove(pair.Value)

        Next

        Return updatedKeys

    End Function

    ''' <summary> 
    ''' Outputs the details of a modified entry's changes to the user 
    ''' </summary>
    ''' 
    ''' <param name="section">
    ''' The modified entry whose changes are being observed 
    ''' </param>
    ''' 
    ''' <param name="changeType"> 
    ''' The type of change to observe (as it will be described to the user) <br />
    ''' <list type="table">
    ''' <item> <term> 0 </term> <description> "added"        </description> </item>
    ''' <item> <term> 1 </term> <description> "removed"      </description> </item>
    ''' <item> <term> 2 </term> <description> "modified"     </description> </item>
    ''' <item> <term> 3 </term> <description> "renamed to "  </description> </item>
    ''' <item> <term> 4 </term> <description> "merged into " </description> </item>
    ''' </list> 
    ''' </param>
    ''' 
    ''' <param name="newSection"> 
    ''' The new version of the modified entry or the entry into which another has been merged 
    ''' </param>
    ''' 
    ''' Docs last updated: 2022-06-10 | Code last updated: 2022-06-10
    Private Sub MakeDiff(section As iniSection,
                        changeType As Integer,
                        Optional newSection As iniSection = Nothing)

        Dim printColor As ConsoleColor = ConsoleColor.Cyan
        If changeType = 2 OrElse changeType = 3 Then printColor = If(changeType = 2, ConsoleColor.Yellow, ConsoleColor.Magenta)

        Dim renamedOrMergedEntryName = If(newSection IsNot Nothing, newSection.Name, "")
        Dim changeTypeStrs = {"added", "removed", "modified", "renamed to ", "merged into "}
        Dim changeStr = $"{section.Name} has been {changeTypeStrs(changeType)}{renamedOrMergedEntryName}"
        LogAndPrint(0, changeStr, indent:=True, leadr:=True, isCentered:=True, fillBorder:=False, colorLine:=True, useArbitraryColor:=changeType >= 2, enStrCond:=changeType < 1, arbitraryColor:=printColor)

        If Not ShowFullEntries Then Return

        Dim isMergeOrRenamed = changeType >= 3 AndAlso changeType < 5

        LogAndPrint(0, "Old entry: ", cond:=isMergeOrRenamed, leadingBlank:=True, logCond:=isMergeOrRenamed, leadr:=True)

        printEntry(section.ToString)

        If Not isMergeOrRenamed Then Return

        Dim out = If(changeType = 3, "Renamed entry: ", "Merged entry: ")

        LogAndPrint(0, out, leadingBlank:=True)

        printEntry(newSection.ToString)

    End Sub

    ''' <summary>
    ''' When <c> ShowFullEntries </c> is <c> True </c>, prints the given <c> entry </c> to the user
    ''' </summary>
    ''' 
    ''' <param name="entry">
    ''' 
    ''' </param>
    ''' 
    ''' <remarks> 
    ''' Used by <c> MakeDiff </c> for printing old/new entry pairs for "Verbose Mode" only 
    ''' </remarks>
    ''' 
    ''' Docs last updated: 2022-06-12 | Code last updated: 2022-06-12
    Private Sub printEntry(entry As String)

        Dim splitEntry = entry.Split(CChar(vbCrLf))

        For i = 0 To splitEntry.Length - 1

            Dim out = splitEntry(i).Replace(vbLf, "")
            Dim isLastLine = i = splitEntry.Length - 1

            LogAndPrint(0, out, indent:=True, indAmt:=4, trailr:=isLastLine, trailingBlank:=isLastLine)

        Next

    End Sub

End Module