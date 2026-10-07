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
''' CC7Patcher is a winapp2ool module that patches the CCleaner 7 ccleaner.ini
''' with CCleaner 7 syntax winapp2.ini entries to enable CCleaner 7 compatibility
''' <br /><br />
''' By default the module downloads the CCleaner 7 flavor of winapp2.ini from GitHub, ignoring
''' whatever <see cref="CurrentWinappFlavor"/> says. It can instead read a local file, which it
''' uses as is without converting it. It can trim the input before applying the patches to
''' ccleaner.ini. Patching first removes the entries a previous patch added, so running it
''' again replaces them rather than duplicating them.
''' </summary>
Public Module CC7Patcher

    ''' <summary>
    ''' The <c> Author </c> key value stamped onto every entry by the CCleaner7 flavor
    ''' (<c> cc7_additions.ini </c>). It identifies winapp2-authored sections in a
    ''' previously-patched <c> ccleaner.ini </c>, so they can be pruned before re-patching to
    ''' keep patching idempotent. It must match the value in <c> cc7_additions.ini </c>, and
    ''' nothing checks that it does.
    ''' </summary>
    Private Const CC7AuthorStamp As String = "Winapp2.ini Project"

    ''' <summary>
    ''' Handles the command line arguments for CC7Patcher, starting from the default settings
    ''' rather than any saved ones, then runs the patcher
    ''' </summary>
    '''
    ''' <remarks>
    ''' File arguments: <c> -1d </c>/<c> -1f </c> (or <c> -winapp2d </c>/<c> -winapp2f </c>) set
    ''' the local winapp2.ini, <c> -2d </c>/<c> -2f </c> (or <c> -ccleanerd </c>/<c> -ccleanerf </c>)
    ''' the ccleaner.ini to patch, and <c> -3d </c>/<c> -3f </c> the output file.
    ''' Flags:
    ''' <list type="bullet">
    ''' <item><c> -nodownload </c>: read the local winapp2.ini instead of downloading (download is enabled by default)</item>
    ''' <item><c> -trim </c>: Trim winapp2.ini before patching</item>
    ''' </list>
    ''' </remarks>
    Public Sub handleCmdLine()

        InitDefaultCC7PatcherSettings()

        Dim spec As New CliArgSpec(NameOf(CC7Patcher))
        spec.WithFlag("-nodownload", Sub() DownloadWinapp2 = Not DownloadWinapp2) _
            .WithFlag("-trim", Sub() TrimBeforePatching = Not TrimBeforePatching) _
            .WithFile(1, CC7PatcherFile1, "winapp2") _
            .WithFile(2, CC7PatcherFile2, "ccleaner") _
            .WithFile(3, CC7PatcherFile3) _
            .Parse()

        initCC7Patcher()

    End Sub

    ''' <summary>
    ''' Gets the winapp2.ini input, trims it if <c> TrimBeforePatching </c> is on, patches
    ''' ccleaner.ini with it, and displays the results. If ccleaner.ini or the local winapp2.ini
    ''' is missing, or the download can't happen, we set an error header for the next menu and
    ''' return. A missing file or being offline doesn't set a failing exit code. Trimming uses
    ''' the Trim module's include and exclude settings.
    ''' </summary>
    Public Sub initCC7Patcher()

        clrConsole()

        If Not CC7PatcherFile2.Exists() Then

            setNextMenuHeaderText("ccleaner.ini not found. Please select a valid file.", printColor:=ConsoleColor.Red)
            Return

        End If

        Dim winapp2Input As iniFile

        If DownloadWinapp2 Then

            If Not checkOnline() Then

                setNextMenuHeaderText("Internet connection required to download winapp2.ini. Please check your connection.", printColor:=ConsoleColor.Red)
                Return

            End If

            gLog("Downloading CCleaner7 winapp2.ini from GitHub")
            winapp2Input = getRemoteIniFile(cc7FlavorLink)

            If winapp2Input Is Nothing Then

                setNextMenuHeaderText("Failed to download winapp2.ini", printColor:=ConsoleColor.Red)
                Return

            End If

        Else

            If Not CC7PatcherFile1.Exists() Then

                setNextMenuHeaderText("winapp2.ini not found. Please select a valid file.", printColor:=ConsoleColor.Red)
                Return

            End If

            Dim loaded = CC7PatcherFile1.Load(CC7PatcherModuleSettingsChanged, NameOf(CC7Patcher), NameOf(CC7PatcherFile1), NameOf(CC7PatcherModuleSettingsChanged))
            If loaded Is Nothing Then Return
            winapp2Input = loaded

        End If

        Dim menuOutput As New MenuSection
        Dim headerMsg = "CCleaner 7 Patcher"
        menuOutput.AddBoxWithText(headerMsg)

        Using gLogScope(headerMsg)

            If TrimBeforePatching Then

                Dim trimMsg = "Trimming winapp2.ini..."
                menuOutput.AddColoredLine(trimMsg, ConsoleColor.Cyan)
                gLog(trimMsg)

                Dim wa2file As New winapp2file(winapp2Input)
                Trim.trimFile(wa2file)
                winapp2Input = wa2file.ToIni()

                Dim trimCompleteMsg = $"Trimming complete: {wa2file.Count} entries remain"
                menuOutput.AddColoredLine(trimCompleteMsg, ConsoleColor.Green)
                gLog(trimCompleteMsg)

            End If

            patchCCleaner(winapp2Input, menuOutput)

            Dim completeMsg = "CCleaner 7 patching complete"
            menuOutput.AddBoxWithText(completeMsg)
            gLog(completeMsg)

        End Using

        menuOutput.AddAnyKeyPrompt()

        If Not SuppressOutput Then menuOutput.Print()
        crk()

    End Sub

    ''' <summary>
    ''' Reloads ccleaner.ini from <c> CC7PatcherFile2 </c>, prunes the sections a previous patch
    ''' added, and adds the entries from <paramref name="winapp2Input"/> using Transmute's Add
    ''' mode as plain ini rather than winapp2.ini. The result goes to <c> CC7PatcherFile3 </c>.
    ''' </summary>
    '''
    ''' <param name="winapp2Input">
    ''' The winapp2.ini file to use as the source for patching
    ''' </param>
    '''
    ''' <param name="menuOutput">
    ''' The <c> MenuSection </c> containing output to be displayed to the user
    ''' </param>
    Private Sub patchCCleaner(winapp2Input As iniFile,
                              ByRef menuOutput As MenuSection)

        Using gLogScope("Beginning ccleaner.ini patching process")

            Dim baseFile = CC7PatcherFile2.Load(CC7PatcherModuleSettingsChanged, NameOf(CC7Patcher), NameOf(CC7PatcherFile2), NameOf(CC7PatcherModuleSettingsChanged))
            If baseFile Is Nothing Then Return

            pruneWinapp2Sections(baseFile, menuOutput)

            Dim outputFile = iniFile.Empty(CC7PatcherFile3.Dir, CC7PatcherFile3.Name)

            Dim patchMsg = $"Patching {CC7PatcherFile2.Name} with entries from winapp2.ini"
            menuOutput.AddColoredLine(patchMsg, ConsoleColor.Yellow)
            gLog(patchMsg)

            If Not Transmute.RemoteTransmute(baseFile, winapp2Input, outputFile, False, menuOutput, Transmute.TransmuteMode.Add) Then

                Dim failedMsg = $"Patched file was not saved to {CC7PatcherFile3.Path()}"
                menuOutput.AddWarning(failedMsg)
                gLog(failedMsg)
                Return

            End If

            Dim savedMsg = $"Patched file saved to {CC7PatcherFile3.Path()}"
            menuOutput.AddColoredLine(savedMsg, ConsoleColor.Green)
            gLog(savedMsg)

        End Using

    End Sub

    ''' <summary>
    ''' Removes every winapp2-authored section from <paramref name="baseFile"/> in place,
    ''' so that the subsequent Transmute Add starts from a clean slate rather than duplicating keys
    ''' into sections a previous patch already created. This makes patching idempotent and lets a
    ''' dirty <c> ccleaner.ini </c> be updated to a newer winapp2.ini, including dropping entries
    ''' that no longer exist upstream. <br /> <br />
    '''
    ''' winapp2 sections are identified by the <c> Author=Winapp2.ini Project </c> stamp the
    ''' CCleaner7 flavor applies to every entry (see <see cref="CC7AuthorStamp"/>), compared
    ''' ignoring case. Sections without it (CCleaner's own settings, user-authored customs) are
    ''' left untouched, and any section that carries it is removed, whoever wrote it.
    ''' </summary>
    '''
    ''' <param name="baseFile">
    ''' The loaded <c> ccleaner.ini </c> to prune winapp2 sections from
    ''' </param>
    '''
    ''' <param name="menuOutput">
    ''' The <c> MenuSection </c> containing output to be displayed to the user
    ''' </param>
    '''
    ''' <returns>
    ''' The number of sections pruned
    ''' </returns>
    Private Function pruneWinapp2Sections(ByRef baseFile As iniFile,
                                          ByRef menuOutput As MenuSection) As Integer

        Using gLogScope("Pruning existing winapp2 entries from ccleaner.ini")

            Dim toRemove As New List(Of String)

            For Each section In baseFile

                Dim authorKey = section.Keys.GetKey("Author")

                If authorKey IsNot Nothing AndAlso authorKey.Value.Equals(CC7AuthorStamp, StringComparison.OrdinalIgnoreCase) Then

                    toRemove.Add(section.Name)

                End If

            Next

            For Each name In toRemove

                baseFile.RemoveSection(name)

            Next

            Dim entryWord = If(toRemove.Count = 1, "entry", "entries")
            Dim prunedMsg = $"Pruned {toRemove.Count} existing winapp2 {entryWord} before patching"
            menuOutput.AddColoredLine(prunedMsg, ConsoleColor.Yellow)
            gLog(prunedMsg)

            Return toRemove.Count

        End Using

    End Function

End Module
