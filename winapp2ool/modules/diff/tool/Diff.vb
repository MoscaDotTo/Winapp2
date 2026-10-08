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
''' <br /> To "exist" here means for an entry of the same name (compared case-insensitively) to exist
''' <br /> Changes fall into three major categories:
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
    ''' Holds the counts from the most recent Diff, or <c> Nothing </c> if no Diff has run
    ''' </summary>
    Public Property MostRecentDiffOutcome As DiffOutcome = Nothing

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
    Private Sub DiffProgress(curStep As String)

        If SuppressOutput Then Return
        Dim spin = _spinChars(_spinIdx Mod 4)
        _spinIdx += 1
        Console.Write(($"{vbCr}[Diff] {spin} {curStep}"))

    End Sub

    ''' <summary>
    ''' Resets the Diff settings to their defaults, applies the command line arguments, and runs a
    ''' diff if there is a newer file to compare against: the download when downloading is on,
    ''' otherwise <c> -2f </c>. <c> -d </c> without <c> -2f </c> leaves nothing to compare and the
    ''' run does nothing.
    '''
    ''' <br /> Valid Diff args:
    ''' <br /> -d           : turn downloading off and compare two local files. Downloading is on unless offline
    ''' <br /> -donttrim    : turn off trimming the downloaded file before diffing, also on unless offline
    ''' <br /> -savelog     : toggle saving the diff output to disk
    ''' <br /> -verbose     : toggle printing the full text of changed entries in the diff output
    ''' <br /> -1f/-2f/-3f  : the old file, the new file, and the log file
    ''' <br /> -4f/-summaryf: write the machine-readable outcome summary to this file
    ''' </summary>
    Public Sub HandleCmdLine()

        InitDefaultDiffSettings()

        Dim spec As New CliArgSpec("diff")
        spec.WithFile(1, DiffFile1, "old") _
            .WithFile(2, DiffFile2, "new") _
            .WithFile(3, DiffFile3, "log") _
            .WithFile(4, DiffFile4, "summary") _
            .WithDownload(Sub() DownloadDiffFile = False) _
            .WithFlag("-donttrim", Sub() TrimRemoteFile = False) _
            .WithFlag("-savelog", Sub() SaveDiffLog = Not SaveDiffLog) _
            .WithFlag("-verbose", Sub() ShowFullEntries = Not ShowFullEntries) _
            .Parse()

        If DownloadDiffFile Then DiffFile2.Name = "Online winapp2.ini"

        If DiffFile2.Name.Length <> 0 Then ConductDiff()

    End Sub


    ''' <summary>
    ''' Diffs <paramref name="firstFile"/> against the remote winapp2.ini for the current flavor.
    ''' We restore the download and trim settings afterward, but leave <c> DiffFile1 </c> pointing
    ''' at <paramref name="firstFile"/>.
    ''' </summary>
    '''
    ''' <param name="firstFile">The local file to use as the old version</param>
    '''
    ''' <param name="trimFile">
    ''' Indicates whether to trim the downloaded file for the current system before diffing <br /><br />
    ''' Optional, Default: <c> False </c>
    ''' </param>
    Public Sub DiffRemoteFile(firstFile As iniFileChooser,
                     Optional trimFile As Boolean = False)

        DiffFile1.Dir = firstFile.Dir
        DiffFile1.Name = firstFile.Name

        Dim initDDF = DownloadDiffFile
        Dim initTrimRemote = TrimRemoteFile

        DownloadDiffFile = True
        TrimRemoteFile = trimFile

        ConductDiff()

        DownloadDiffFile = initDDF
        TrimRemoteFile = initTrimRemote

    End Sub

    ''' <summary>
    ''' Loads both files, trimming the downloaded one when <c> TrimRemoteFile </c> is set, and
    ''' returns early with a header message if either has no sections. Otherwise runs the diff,
    ''' prints the output unless output is suppressed, saves the log when <c> SaveDiffLog </c> is
    ''' set, writes the outcome summary, and sets the next menu header. The trim path downloads
    ''' the remote file a second time.
    ''' </summary>
    Public Sub ConductDiff()

        Dim oldFile As iniFile
        Dim newFile As iniFile

        oldFile = DiffFile1.Load(DiffModuleSettingsChanged, NameOf(Diff), NameOf(DiffFile1), NameOf(DiffModuleSettingsChanged))
        If Not enforceFileHasContent(oldFile) Then Return

        newFile = If(DownloadDiffFile, getRemoteIniFile(getWinappLink), DiffFile2.Load(DiffModuleSettingsChanged, NameOf(Diff), NameOf(DiffFile2), NameOf(DiffModuleSettingsChanged)))
        If Not enforceFileHasContent(newFile) Then Return

        If TrimRemoteFile AndAlso DownloadDiffFile Then

            Dim tmp As New winapp2file(getRemoteIniFile(getWinappLink))

            Trim.trimFile(tmp)
            newFile = tmp.ToIni()
            If Not enforceFileHasContent(newFile) Then Return

        End If

        clrConsole()

        gLog(DiffLogStartPhrase)

        Dim diffOutput As New List(Of MenuSection)

        Dim out = New MenuSection
        Dim headerText = $"Diff: {GetVer(oldFile)} -> {GetVer(newFile)}"
        out.AddTopBorder().AddColoredLine(headerText, color:=ConsoleColor.DarkGreen, centered:=True).AddDivider()

        diffOutput.Add(out)

        Using gLogScope(headerText)

            diffOutput.AddRange(CompareFiles(oldFile, newFile))

        End Using

        gLog(DiffLogEndPhrase)

        Dim out3 As New MenuSection
        out3.AddBoxWithText(pressEnterStr)

        diffOutput.Add(out3)

        clrConsole()

        If Not SuppressOutput Then diffOutput.ForEach(Sub(section) section.Print())

        MostRecentDiffLog = getLogSliceFromGlobal(DiffLogStartPhrase, DiffLogEndPhrase)

        Dim logFile = iniFile.Empty(DiffFile3.Dir, DiffFile3.Name)
        Dim logSaved = logFile.OverwriteToFile(MostRecentDiffLog, SaveDiffLog)

        WriteOutcomeSummary()

        setNextMenuHeaderText(If(logSaved, DiffFile3.Name & " saved", "Diff complete"), Not SaveDiffLog OrElse logSaved)
        setNextMenuHeaderText($"Diff complete, but {DiffFile3.Name} was not saved", SaveDiffLog AndAlso Not logSaved, ConsoleColor.Red)

        crl()

    End Sub

    ''' <summary>
    ''' Writes the most recent Diff's outcome to <c> DiffFile4 </c> as parseable
    ''' <c> key=value </c> lines, so a calling script can act on what the Diff found without
    ''' parsing the prose log. No-ops when no summary path was given or no Diff has completed.
    ''' </summary>
    Private Sub WriteOutcomeSummary()

        If DiffFile4.Name.Length = 0 OrElse MostRecentDiffOutcome Is Nothing Then Return

        gLog($"Diff outcome: {MostRecentDiffOutcome}")

        Dim summaryFile = iniFile.Empty(DiffFile4.Dir, DiffFile4.Name)
        summaryFile.OverwriteToFile(MostRecentDiffOutcome.ToSummaryText())

    End Sub

    ''' <summary>
    ''' Returns the version string from the first comment of a winapp2.ini file: the comment
    ''' uppercased, with its leading semicolons trimmed and <c> VERSION: </c> replaced by
    ''' <c> version </c>
    ''' </summary>
    ''' 
    ''' <param name="someFile">
    ''' The <c> iniFile </c> whose first comment is inspected for a version tag
    ''' </param>
    ''' 
    ''' <returns>
    ''' A human-readable version string, or <c> " version not given" </c> if no version comment is present
    ''' </returns>
    Private Function GetVer(someFile As iniFile) As String

        Dim ver = If(someFile.Comments.Count > 0, someFile.Comments(0).Text.ToUpperInvariant(), "000000")
        Return If(ver.Contains("VERSION"), ver.TrimStart(CChar(";")).Replace("VERSION:", "version"), " version not given")

    End Function

    ''' <summary>
    ''' Runs the diff pipeline over the two files and stores the result in
    ''' <see cref="MostRecentDiffOutcome"/>. We normalize deprecated paths in both files in place
    ''' first, so the caller's <c> iniFile </c> objects are changed.
    ''' </summary>
    '''
    ''' <param name="oldFile">
    ''' The old version of winapp2.ini as an <c> iniFile </c>
    ''' </param>
    '''
    ''' <param name="newFile">
    ''' The new version of winapp2.ini as an <c> iniFile </c>
    ''' </param>
    '''
    ''' <returns>
    ''' All <c> MenuSection </c>s produced by the diff pipeline, in display order
    ''' </returns>
    Friend Function CompareFiles(oldFile As iniFile,
                                   newFile As iniFile) As List(Of MenuSection)

        Dim out As New List(Of MenuSection)

        Dim state As New DiffState()
        state.Clear()

        Dim keyAnalyzer = New KeyModificationAnalyzer(state)
        Dim mergeDetector = New MergeDetector(state, newFile, AddressOf keyAnalyzer.FindModifications)
        Dim renderer = New DiffOutputRenderer(state, oldFile, newFile, keyAnalyzer)
        Dim detector = New EntryChangeDetector(state, oldFile, newFile, mergeDetector, keyAnalyzer, renderer)
        Dim statsCalc = New DiffStatisticsCalculator(state, oldFile, newFile)

        detector.SnuffNoisyChanges(oldFile)
        detector.SnuffNoisyChanges(newFile)

        Dim stepNum = 0
        Const totalSteps = 19

        Dim doStep = Sub(label As String, action As Action)
                         stepNum += 1
                         DiffProgress($"{label} (step {stepNum}/{totalSteps})")
                         action()
                     End Sub

        ' A step that found nothing contributes nothing, including its trailing divider. Adding
        ' the divider unconditionally left a diff with few changes (and a diff with none at all)
        ' rendering a stack of empty bordered rows with no content between them
        Dim collectStep = Sub(label As String, fn As Func(Of IEnumerable(Of MenuSection)))
                              stepNum += 1
                              DiffProgress($"{label} (step {stepNum}/{totalSteps})")
                              Dim produced = fn().Where(Function(section) Not section.IsEmpty).ToList()
                              If produced.Count = 0 Then Return
                              out.AddRange(produced)
                              out.Add(New MenuSection().AddDivider(solid:=False))
                          End Sub

        Dim start = Now

        doStep("· processing new entries ", Sub() detector.ProcessNewEntries())
        doStep("· processing old entries ", Sub() detector.ProcessOldEntries())
        doStep("· detecting browser changes ", Sub() statsCalc.DetectNewBrowserSupport())
        collectStep("· itemizing new browsers    ", Function() renderer.ItemizeNewBrowsers())
        collectStep("· itemizing removed browsers", Function() renderer.ItemizeRemovedBrowsers())
        collectStep($"· itemizing removals            ", Function() detector.ProcessRemovals())
        doStep("· tracking keys across entries   ", Sub() statsCalc.DetectCrossEntryMovements())
        doStep("· calculating initial statistics ", Sub() statsCalc.CalculateInitialStatistics())
        doStep("· calculating rename statistics  ", Sub() statsCalc.CalculateRenameStatistics())
        collectStep("· tracking renamed entries      ", Function() renderer.SummarizeRenames())
        collectStep("· tracking splits and mergers   ", Function() renderer.SummarizeMergers())
        collectStep("· diffing renamed entries       ", Function() renderer.ItemizeRenameChanges())
        collectStep("· diffing merged entries        ", Function() renderer.ItemizeMergers())
        collectStep("· itemizing key movement info   ", Function() renderer.ItemizeKeyMovements())
        collectStep("· diffing modified entries      ", Function() renderer.ItemizeModifications())
        collectStep("· itemizing added-with-mergers  ", Function() renderer.ItemizeAddedEntriesWithMergers())
        collectStep("· itemizing novel entries       ", Function() renderer.ItemizeAdditions())
        doStep("· calculating final statistics  ", Sub() statsCalc.CalculateAddedWithMergersStatistics())

        Dim timeSpan = Now - start
        gLog($"Total diff time: {timeSpan}")

        out.Add(New MenuSection().AddBottomBorder)

        doStep("· calculating summary statistics ", Sub() out.Add(renderer.LogPostDiff()))

        MostRecentDiffOutcome = renderer.BuildOutcome()

        Return out

    End Function

End Module
