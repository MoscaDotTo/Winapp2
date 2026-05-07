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

''' <summary>
''' CCiniDebug is a winapp2ool module which performs housekeeping on the CCleaner Classic (v1-6)
''' configuration file ccleaner.ini to clean up leftovers from winapp2.ini
''' <br /><br />
''' As entries are removed or renamed in winapp2.ini over time, stale configuration keys are
''' leftover in ccleaner.ini. CCiniDebug processes a ccleaner.ini for its entry configuration keys
''' and checks them each against the names of entries in winapp2.ini. Any configuration keys found
''' with the winapp2.ini indicator (*) in the name which do not have a corresponding entry by the
''' same name in the most recent winapp2.ini can then be easily and automatically removed
''' </summary>
Module CCiniDebug

    ''' <summary>
    ''' Handles the commandline args for CCiniDebug
    ''' </summary>
    '''
    ''' <remarks>
    ''' CCiniDebug setting toggles:
    ''' -noprune    : disable pruning of stale winapp2.ini entries
    ''' -nosort     : disable sorting ccleaner.ini alphabetically
    ''' -nosave     : disable saving the modified ccleaner.ini back to file
    ''' </remarks>
    Public Sub handleCmdlineArgs()

        initDefaultCCDBSettings()

        Dim spec As New CliArgSpec("ccinidebug")
        spec.WithFlag("-noprune", Sub() PruneStaleEntries = Not PruneStaleEntries) _
            .WithFlag("-nosort", Sub() SortFileForOutput = Not SortFileForOutput) _
            .WithFlag("-nosave", Sub() SaveDebuggedFile = Not SaveDebuggedFile) _
            .WithFile(1, CCDebugFile1) _
            .WithFile(2, CCDebugFile2) _
            .WithFile(3, CCDebugFile3) _
            .Parse()

        initCCDebug()

    End Sub

    ''' <summary>
    ''' Loads and analyzes <c> ccleaner.ini </c>, pruning orphaned entries and sorting if enabled,
    ''' then displays the results.
    ''' </summary>
    Public Sub initCCDebug()

        Dim ccIni = CCDebugFile2.Load
        If Not enforceFileHasContent(ccIni) Then Return

        Dim winapp2 As iniFile2 = Nothing

        If PruneStaleEntries Then

            winapp2 = CCDebugFile1.Load
            If Not enforceFileHasContent(winapp2) Then Return

        End If

        Dim orphans As List(Of String)

        Using gLogScope($"Analyzing {CCDebugFile2.Name}")
            orphans = ccDebug(winapp2, ccIni)
        End Using

        Dim completionText = $"Analysis complete{If(SaveDebuggedFile, $" — {CCDebugFile3.Name} saved", "")}."
        gLog(completionText)

        Dim out As New MenuSection
        out.AddTopBorder() _
           .AddColoredLine($"Analyzing {CCDebugFile2.Name}", ConsoleColor.DarkGreen, centered:=True) _
           .AddDivider()

        If orphans IsNot Nothing Then

            If orphans.Count > 0 Then orphans.ForEach(Sub(entry) out.AddColoredLine($"Orphaned: {entry}", ConsoleColor.Yellow))
            out.AddLine($"{orphans.Count} orphaned setting{If(orphans.Count > 1, "s", "")} removed")

        End If

        out.AddColoredLine(completionText, ConsoleColor.Green, centered:=True) _
           .AddDivider() _
           .AddLine(anyKeyStr, centered:=True) _
           .AddBottomBorder()

        clrConsole()
        out.Print()
        crk()

    End Sub

    ''' <summary>
    ''' Prunes, sorts, and saves ccleaner.ini using the given loaded files.
    ''' </summary>
    '''
    ''' <param name="wa2file">
    ''' The loaded winapp2.ini used for pruning, or <c> Nothing </c> if pruning is disabled
    ''' </param>
    '''
    ''' <param name="ccinifile">
    ''' The loaded ccleaner.ini to process
    ''' </param>
    '''
    ''' <returns>
    ''' A <c> List(Of String) </c> of orphaned entry names removed during pruning,
    ''' or <c> Nothing </c> if pruning was not performed.
    ''' </returns>
    Private Function ccDebug(wa2file As iniFile2,
                             ccinifile As iniFile2) As List(Of String)

        Dim orphans As List(Of String) = Nothing

        If PruneStaleEntries Then

            Dim optSec = ccinifile.GetSection("Options")
            If optSec IsNot Nothing Then orphans = prune(optSec, wa2file)

        End If

        If SortFileForOutput Then sortCC(ccinifile)

        Dim outFile = iniFile2.Empty(CCDebugFile3.Dir, CCDebugFile3.Name)
        outFile.OverwriteToFile(ccinifile.ToString(), SaveDebuggedFile)

        Return orphans

    End Function

    ''' <summary>
    ''' Scans for and removes stale winapp2.ini entry settings from the Options section of ccleaner.ini
    ''' </summary>
    '''
    ''' <param name="optionsSec">
    ''' The Options section from ccleaner.ini
    ''' </param>
    '''
    ''' <param name="wa2file">
    ''' The winapp2.ini file against which to check for orphaned entries in ccleaner.ini
    ''' </param>
    '''
    ''' <returns>
    ''' A <c> List(Of String) </c> of orphaned entry names that were removed
    ''' </returns>
    Private Function prune(optionsSec As iniSection2,
                           wa2file As iniFile2) As List(Of String)

        Dim tbTrimmed As New List(Of iniKey2)
        Dim orphanNames As New List(Of String)

        Using gLogScope($"Scanning {CCDebugFile2.Name} for settings left over from removed winapp2.ini entries")

            For Each key In optionsSec.Keys

                Dim isValid = key.Name.StartsWith("(App)", StringComparison.InvariantCulture) AndAlso key.Name.Contains("*")
                If Not isValid Then Continue For

                Dim entryName = key.Name.Replace("(App)", "").Trim()

                If wa2file.Contains(entryName) Then Continue For

                tbTrimmed.Add(key)
                orphanNames.Add(entryName)
                gLog($"Orphaned entry detected: {entryName}")

            Next

            For Each k In tbTrimmed : optionsSec.Keys.Remove(k) : Next

        End Using

        gLog($"  {tbTrimmed.Count} orphaned settings detected")

        Return orphanNames

    End Function

    ''' <summary>
    ''' Sorts the keys in the Options section of <c> f2 </c> alphabetically
    ''' </summary>
    '''
    ''' <param name="ccinifile">
    ''' The loaded ccleaner.ini whose Options section keys will be sorted
    ''' </param>
    Private Sub sortCC(ccinifile As iniFile2)

        Using gLogScope($"Sorting {CCDebugFile2.Name}")

            Dim optSec = ccinifile.GetSection("Options")
            If optSec Is Nothing Then Return

            Dim sorted = optSec.Keys.ToList()
            sorted.Sort(Function(a, b) String.Compare(a.Name, b.Name, StringComparison.OrdinalIgnoreCase))

            Dim all = optSec.Keys.ToList()
            For Each k In all : optSec.Keys.Remove(k) : Next
            For Each k In sorted : optSec.Keys.Add(k) : Next

        End Using

    End Sub

End Module
