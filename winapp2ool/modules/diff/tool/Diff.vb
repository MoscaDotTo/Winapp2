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

Imports System.Globalization

'''<summary>
'''
''' Compares two winapp2.ini format <c> iniFile </c>s and summarizes the changes to the user
''' <br />
''' <br /> NOTE: to "exist" here means for an entry of the same exact name (case sensitive) to exist
''' <br /> Changes fall three major categories:
'''
''' <list type="table">
'''
''' <item>
''' <term> Added entries </term>
''' <description> exist in the new file and not in the old file </description>
''' </item>
'''
''' <item>
''' <term> Modified entries </term>
''' <description> exist in both the new file and the old file and have been changed in some way  </description>
''' </item>
'''
''' <item>
''' <term> Removed entries </term>
''' <description> exist in the old file but not in the new file  </description>
''' </item>
'''
''' </list>
'''
''' <br />
''' <br /> Additionally, Removed entries have three sub categories:
'''
''' <list type="table">
'''
''' <item>
''' <term> Renamed entries </term>
''' <description> do not exist in the new file, but their content exists in some other entry in the file,
''' mostly unchanged from the old version (may contain minor changes) </description>
''' </item>
'''
''' <item>
''' <term> Merged entries </term>
''' <description> do not exist in the new file, but their content exists in some other entry in the new file
''' which is substantially different from the old version </description>
''' </item>
'''
''' <item>
''' <term> Removed without replacement </term>
''' <description> do not exist in the new file and their content was not found in some other entry in the new file  </description>
''' </item>
'''
''' </list>
'''
''' <br />
''' <br /> Likewise, Merged entries themselves are broken into two categories
''' <list type="table">
'''
''' <item>
''' <term> Modified </term>
''' <description> Entries that existed in the old file which have been modified to contain content from entries which have been removed  </description>
''' </item>
'''
''' <item>
''' <term> Added </term>
''' <description> Entries which did not exist in the old file but who contain content from entries which have been removed  </description>
''' </item>
'''
''' </list>
'''
''' </summary>
Module Diff

    Private ReadOnly State As New DiffState()

    ''' <summary>
    ''' Holds the slice of the winapp2ool global log containing the most recent Diff results
    ''' </summary>
    Public Property MostRecentDiffLog As String = ""

    Public Property DiffLogStartPhrase As String = "Beginning Diff"

    Public Property DiffLogEndPhrase As String = "Diff complete"

    Private Const Diff2StartPhrase As String = "Beginning Diff2"
    Private Const Diff2EndPhrase As String = "Diff2 complete"

    Private Property mergeDetector As MergeDetector
    Private Property keyAnalyzer As KeyModificationAnalyzer
    Private Property renderer As DiffOutputRenderer

    ''' <summary>
    ''' Runs a diff using command line arguments, allowing Diff to be called programmatically
    ''' <list type="table">
    ''' <item> -d </item>
    ''' </list>
    '''
    ''' <br /> Valid Diff args:
    ''' <br /> -d           : download the latest winapp2.ini
    ''' <br /> -donttrim    : download the latest non-ccleaner winapp2.ini (implies -d)
    ''' <br /> -savelog     : save diff.txt to disk on exit
    ''' </summary>
    Public Sub HandleCmdLine()

        InitDefaultDiffSettings()

        ' Downloading the remote file is the default behavior, providing -d disables it
        handleDownloadBools(DownloadDiffFile)

        If DownloadDiffFile Then DiffFile2.Name = "Online winapp2.ini"

        invertSettingAndRemoveArg(TrimRemoteFile, "-donttrim")
        invertSettingAndRemoveArg(SaveDiffLog, "-savelog")
        getFileAndDirParams({DiffFile1, DiffFile2, DiffFile3})

        If DiffFile2.Name.Length <> 0 Then ConductDiff()

    End Sub

    ''' <summary>
    ''' Runs a Diff between a local file and the most recent winapp2.ini from GitHub
    ''' </summary>
    '''
    ''' <param name="firstFile">
    ''' The local winapp2.ini file to diff against the master GitHub copy
    ''' </param>
    Public Sub DiffRemoteFile(firstFile As iniFile)

        DiffFile1 = firstFile
        DownloadDiffFile = True
        ConductDiff()

    End Sub

    ''' <summary>
    ''' Ensures both files have content before kicking off the Diff and then summarizes the output from the Diff
    ''' </summary>
    Public Sub ConductDiff()

        If Not enforceFileHasContent(DiffFile1) Then Return

        If DownloadDiffFile Then

            Dim downloadedIniFile = getRemoteIniFile(getWinappLink)
            DiffFile2.Sections = downloadedIniFile.Sections
            DiffFile2.Comments = downloadedIniFile.Comments

        Else

            If Not enforceFileHasContent(DiffFile2) Then Return

        End If

        If TrimRemoteFile AndAlso DownloadDiffFile Then

            Dim tmp As New winapp2file(DiffFile2)
            Trim.trimFile(tmp)
            DiffFile2.Sections = tmp.toIni.Sections

        End If

        clrConsole()

        gLog(DiffLogStartPhrase)

        Dim diffOutput As New List(Of MenuSection)

        Dim out = New MenuSection
        Dim headerText = $"Diff: {GetVer(DiffFile1)} -> {GetVer(DiffFile2)}"
        out.AddTopBorder().AddColoredLine(headerText, color:=ConsoleColor.DarkGreen, centered:=True)
        out.AddDivider()
        gLog(headerText, ascend:=True)
        diffOutput.Add(out)

        diffOutput.AddRange(CompareFiles())

        Dim out2 As New MenuSection
        out2.AddBottomBorder()
        diffOutput.Add(out2)

        diffOutput.Add(renderer.LogPostDiff())

        gLog(DiffLogEndPhrase)

        Dim out3 As New MenuSection
        out3.AddBoxWithText(pressEnterStr)

        diffOutput.Add(out3)
        diffOutput.ForEach(Sub(section) section.Print())

        MostRecentDiffLog = getLogSliceFromGlobal(DiffLogStartPhrase, DiffLogEndPhrase)
        DiffFile3.overwriteToFile(MostRecentDiffLog, SaveDiffLog)

        ' Run comparison pass with new iniFile2 implementation
        Dim file1As2 = DiffFileBridge.ToIniFile2(DiffFile1)
        Dim file2As2 = DiffFileBridge.ToIniFile2(DiffFile2)

        gLog(Diff2StartPhrase)
        gLog("", ascend:=True)
        CompareFiles2(file1As2, file2As2)
        gLog(Diff2EndPhrase)

        Dim newLog = getLogSliceFromGlobal(Diff2StartPhrase, Diff2EndPhrase)

        DiffFile3.Name = "Diff-inifile2.txt"
        DiffFile3.overwriteToFile(newLog, SaveDiffLog)

        setNextMenuHeaderText(If(SaveDiffLog, DiffFile3.Name & " saved", "Diff complete"))

    End Sub

    ''' <summary>
    ''' Gets the version from winapp2.ini
    ''' </summary>
    '''
    ''' <param name="someFile">
    ''' A winapp2.ini format <c>iniFile</c>
    ''' </param>
    '''
    ''' <returns>
    ''' The current version if available
    ''' <br /> "version not given" otherwise
    ''' </returns>
    Private Function GetVer(someFile As iniFile) As String

        Dim ver = If(someFile.Comments.Count > 0, someFile.Comments(0).Comment.ToString(CultureInfo.InvariantCulture).ToUpperInvariant, "000000")
        Return If(ver.Contains("VERSION"), ver.TrimStart(CChar(";")).Replace("VERSION:", "version"), " version not given")

    End Function

    ''' <summary>
    ''' Runs the diff pipeline using the new <c>iniFile2</c>-based core classes.
    ''' Writes output to gLog for comparison against the legacy path. Results are not printed.
    ''' </summary>
    '''
    ''' <param name="file1As2">The already-snuffed old version of winapp2.ini as an <c>iniFile2</c></param>
    ''' <param name="file2As2">The already-snuffed new version of winapp2.ini as an <c>iniFile2</c></param>
    Private Sub CompareFiles2(file1As2 As iniFile2, file2As2 As iniFile2)

        Dim state2 As New DiffState()
        state2.Clear()
        Dim start = Now

        Dim keyAnalyzer2 = New KeyModificationAnalyzer2(state2)
        Dim mergeDetector2 = New MergeDetector2(state2, file2As2, AddressOf keyAnalyzer2.FindModifications)

        ' Renderer and stats calc still need legacy iniFile objects
        Dim file1AsOld = DiffFileBridge.ToIniFile(file1As2)
        Dim file2AsOld = DiffFileBridge.ToIniFile(file2As2)

        ' Separate KeyModificationAnalyzer (legacy) for renderer merger analysis — shares state2
        Dim keyAnalyzerForRenderer = New KeyModificationAnalyzer(state2)
        Dim renderer2 = New DiffOutputRenderer(state2, file1AsOld, file2AsOld, keyAnalyzerForRenderer)
        Dim detector2 = New EntryChangeDetector2(state2, file1As2, file2As2, mergeDetector2, keyAnalyzer2, renderer2)
        Dim statsCalc2 = New DiffStatisticsCalculator(state2, file1AsOld, file2AsOld)

        ' Phase 1: Gather raw changes
        detector2.ProcessNewEntries()
        detector2.ProcessOldEntries()
        detector2.ProcessRemovals()
        statsCalc2.CalculateInitialStatistics()

        ' Phase 2: Detect cross-entry movements
        statsCalc2.DetectCrossEntryMovements()

        ' Phase 3: Generate output (writes to gLog for comparison)
        renderer2.SummarizeRenames()
        renderer2.SummarizeMergers()
        renderer2.ItemizeMergers()
        renderer2.ItemizeKeyMovements()
        renderer2.ItemizeModifications()
        renderer2.ItemizeAddedEntriesWithMergers()
        renderer2.ItemizeAdditions()

        statsCalc2.CalculateAddedWithMergersStatistics()

        Dim ender = Now
        Dim timeSpan = ender - start
        gLog(timeSpan.ToString)

        renderer2.LogPostDiff()

    End Sub

    ''' <summary>
    ''' Compares two winapp2.ini format <c>iniFiles</c>, itemizes, and summarizes the differences to the user
    ''' </summary>
    Private Function CompareFiles() As List(Of MenuSection)

        State.Clear()

        keyAnalyzer = New KeyModificationAnalyzer(State)
        mergeDetector = New MergeDetector(State, DiffFile2, AddressOf keyAnalyzer.FindModifications)
        renderer = New DiffOutputRenderer(State, DiffFile1, DiffFile2, keyAnalyzer)

        Dim detector = New EntryChangeDetector(State, DiffFile1, DiffFile2, mergeDetector, keyAnalyzer, renderer)
        Dim statsCalc = New DiffStatisticsCalculator(State, DiffFile1, DiffFile2)

        detector.SnuffNoisyChanges(DiffFile1)
        detector.SnuffNoisyChanges(DiffFile2)

        Dim out = New List(Of MenuSection)
        Dim start = Now

        ' Phase 1: Gather raw changes
        detector.ProcessNewEntries()
        detector.ProcessOldEntries()
        out.AddRange(detector.ProcessRemovals())
        statsCalc.CalculateInitialStatistics()

        ' Phase 2: Detect cross-entry movements
        statsCalc.DetectCrossEntryMovements()

        ' Phase 3: Generate output from cleaned data
        out.AddRange(renderer.SummarizeRenames())
        out.AddRange(renderer.SummarizeMergers())
        out.AddRange(renderer.ItemizeMergers())
        out.AddRange(renderer.ItemizeKeyMovements())
        out.AddRange(renderer.ItemizeModifications())
        out.AddRange(renderer.ItemizeAddedEntriesWithMergers())
        out.AddRange(renderer.ItemizeAdditions())

        statsCalc.CalculateAddedWithMergersStatistics()

        Dim ender = Now
        Dim timeSpan = ender - start
        gLog(timeSpan.ToString)

        Return out

    End Function

End Module
