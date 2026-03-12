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

    ''' <summary>
    ''' Holds the slice of the winapp2ool global log containing the most recent Diff results
    ''' </summary>
    Public Property MostRecentDiffLog As String = ""

    ''' <summary>
    ''' Phrase written to the global log to mark the beginning of a Diff run,
    ''' used to slice the relevant portion of the log afterwards
    ''' </summary>
    Public Property DiffLogStartPhrase As String = "Beginning Diff"

    ''' <summary>
    ''' Phrase written to the global log to mark the end of a Diff run,
    ''' used to slice the relevant portion of the log afterwards
    ''' </summary>
    Public Property DiffLogEndPhrase As String = "Diff complete"

    Private _spinIdx As Integer = 0
    Private ReadOnly _spinChars As Char() = {"|"c, "/"c, "-"c, "\"c}

    ''' <summary>
    ''' Overwrites the current console line with a spinner and step label.
    ''' No-ops in silent mode (<see cref="SuppressOutput"/>).
    ''' </summary>
    ''' 
    ''' <param name="curStep">
    ''' Human-readable label for the current pipeline step
    ''' </param>
    Private Sub Diff2Progress(curStep As String)

        If SuppressOutput Then Return
        Dim spin = _spinChars(_spinIdx Mod 4)
        _spinIdx += 1
        Console.Write(($"{vbCr}[Diff] {spin} {curStep}"))

    End Sub

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

        DiffFile1.Dir = firstFile.Dir
        DiffFile1.Name = firstFile.Name
        DownloadDiffFile = True
        ConductDiff()

    End Sub

    ''' <summary>
    ''' Ensures both files have content before kicking off the Diff
    ''' and then summarizes the output from the Diff
    ''' </summary>
    Public Sub ConductDiff()

        Dim file1As2 As iniFile2
        Dim file2As2 As iniFile2

        file1As2 = DiffFile1.Load()
        If Not enforceFileHasContent(file1As2) Then Return

        If DownloadDiffFile Then

            file2As2 = getRemoteIniFile2(getWinappLink)
            If file2As2 Is Nothing Then Return

        Else

            file2As2 = DiffFile2.Load()
            If Not enforceFileHasContent(file2As2) Then Return

        End If

        If TrimRemoteFile AndAlso DownloadDiffFile Then

            ' Trim still uses the old winapp2file model - bridge until Trim is migrated
            Dim tmp As New winapp2file(getRemoteIniFile(getWinappLink))
            Trim.trimFile(tmp)
            file2As2 = IniFileBridge.ToIniFile2(tmp.toIni)

        End If

        If Not enforceFileHasContent(file1As2) Then Return
        If Not enforceFileHasContent(file2As2) Then Return

        clrConsole()

        gLog(DiffLogStartPhrase)

        Dim diffOutput As New List(Of MenuSection)

        Dim out = New MenuSection
        Dim headerText = $"Diff: {GetVer(file1As2)} -> {GetVer(file2As2)}"
        out.AddTopBorder().AddColoredLine(headerText, color:=ConsoleColor.DarkGreen, centered:=True).AddDivider()

        gLog(headerText, ascend:=True)
        diffOutput.Add(out)

        diffOutput.AddRange(CompareFiles2(file1As2, file2As2))

        gLog(DiffLogEndPhrase)

        Dim out3 As New MenuSection
        out3.AddBoxWithText(pressEnterStr)

        diffOutput.Add(out3)
        diffOutput.ForEach(Sub(section) section.Print())

        MostRecentDiffLog = getLogSliceFromGlobal(DiffLogStartPhrase, DiffLogEndPhrase)

        Dim logFile = iniFile2.Empty(DiffFile3.Dir, DiffFile3.Name)
        logFile.OverwriteToFile(MostRecentDiffLog, SaveDiffLog)

        setNextMenuHeaderText(If(SaveDiffLog, DiffFile3.Name & " saved", "Diff complete"))

        crl()

    End Sub

    ''' <summary>
    ''' Gets the version string from the first comment of a winapp2.ini file
    ''' </summary>
    ''' 
    ''' <param name="someFile">The <c>iniFile2</c> whose first comment is inspected for a version tag</param>
    ''' <returns>
    ''' A human-readable version string, or <c>" version not given"</c> if no version comment is present
    ''' </returns>
    Private Function GetVer(someFile As iniFile2) As String

        Dim ver = If(someFile.Comments.Count > 0, someFile.Comments(0).Text.ToUpperInvariant(), "000000")
        Return If(ver.Contains("VERSION"), ver.TrimStart(CChar(";")).Replace("VERSION:", "version"), " version not given")

    End Function

    ''' <summary>
    ''' Runs the diff pipeline using the <c>iniFile2</c>-based core classes.
    ''' Returns all output sections for display and logging.
    ''' </summary>
    '''
    ''' <param name="file1As2">
    ''' The old version of winapp2.ini as an <c>iniFile2</c>
    ''' </param>
    '''
    ''' <param name="file2As2">
    ''' The new version of winapp2.ini as an <c>iniFile2</c>
    ''' </param>
    '''
    ''' <returns>
    ''' All <c>MenuSection</c>s produced by the diff pipeline, in display order
    ''' </returns>
    Private Function CompareFiles2(file1As2 As iniFile2,
                                   file2As2 As iniFile2) As List(Of MenuSection)

        Dim out As New List(Of MenuSection)

        Dim state2 As New DiffState()
        state2.Clear()

        Dim keyAnalyzer2 = New KeyModificationAnalyzer2(state2)
        Dim mergeDetector2 = New MergeDetector2(state2, file2As2, AddressOf keyAnalyzer2.FindModifications)
        Dim renderer2 = New DiffOutputRenderer2(state2, file1As2, file2As2, keyAnalyzer2)
        Dim detector2 = New EntryChangeDetector2(state2, file1As2, file2As2, mergeDetector2, keyAnalyzer2, renderer2)
        Dim statsCalc2 = New DiffStatisticsCalculator2(state2, file1As2, file2As2)

        detector2.SnuffNoisyChanges(file1As2)
        detector2.SnuffNoisyChanges(file2As2)

        Dim stepNum = 0
        Const totalSteps = 16

        Dim doStep = Sub(label As String, action As Action)
                         stepNum += 1
                         Diff2Progress($"{label} (step {stepNum}/{totalSteps})")
                         action()
                     End Sub

        Dim collectStep = Sub(label As String, fn As Func(Of IEnumerable(Of MenuSection)))
                              stepNum += 1
                              Diff2Progress($"{label} (step {stepNum}/{totalSteps})")
                              out.AddRange(fn())
                          End Sub

        Dim start = Now

        ' Phase 1: Gather raw changes
        doStep("Phase 1 · processing new entries", Sub() detector2.ProcessNewEntries())
        doStep("Phase 1 · processing old entries", Sub() detector2.ProcessOldEntries())
        collectStep($"Phase 1 · processing {state2.ModifiedEntries.RemovedEntryNames.Count} removals", Function() detector2.ProcessRemovals())
        doStep("Phase 1 · calculating initial statistics", Sub() statsCalc2.CalculateInitialStatistics())

        ' Phase 2: Detect cross-entry movements
        doStep("Phase 2 · detecting cross-entry key movements", Sub() statsCalc2.DetectCrossEntryMovements())
        doStep("Phase 2 · calculating rename statistics", Sub() statsCalc2.CalculateRenameStatistics())

        ' Phase 3: Generate output
        collectStep("Phase 3 · processing renames", Function() renderer2.SummarizeRenames())
        collectStep("Phase 3 · processing mergers", Function() renderer2.SummarizeMergers())
        collectStep("Phase 3 · itemizing rename key changes", Function() renderer2.ItemizeRenameChanges())
        collectStep("Phase 3 · itemizing mergers into existing entries", Function() renderer2.ItemizeMergers())
        collectStep("Phase 3 · itemizing cross-entry key movements", Function() renderer2.ItemizeKeyMovements())
        collectStep("Phase 3 · itemizing entry modifications", Function() renderer2.ItemizeModifications())
        collectStep("Phase 3 · itemizing mergers into newly added entries", Function() renderer2.ItemizeAddedEntriesWithMergers())
        collectStep("Phase 3 · itemizing additions", Function() renderer2.ItemizeAdditions())
        doStep("Phase 3 · calculating added-with-mergers statistics", Sub() statsCalc2.CalculateAddedWithMergersStatistics())

        Dim timeSpan = Now - start
        gLog($"Total diff time: {timeSpan}")

        doStep("Phase 3 · calculating summary statistics", Sub() out.Add(renderer2.LogPostDiff()))

        Return out

    End Function

End Module
