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
''' The module can optionally download the latest winapp2.ini from GitHub and trim it
''' before applying the patches to ccleaner.ini
''' </summary>
Public Module CC7Patcher

    ''' <summary>
    ''' Handles the command line arguments for CC7Patcher
    ''' </summary>
    '''
    ''' <remarks>
    ''' CC7Patcher args:
    ''' -nodownload     : Disable downloading winapp2.ini (download is enabled by default)
    ''' -trim           : Trim winapp2.ini before patching
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
    ''' Initializes the CC7Patcher process
    ''' </summary>
    Public Sub initCC7Patcher()

        clrConsole()

        ' Validate ccleaner.ini exists before proceeding
        If Not CC7PatcherFile2.Exists() Then

            setNextMenuHeaderText("ccleaner.ini not found. Please select a valid file.", printColor:=ConsoleColor.Red)
            Return

        End If

        ' Handle winapp2.ini acquisition
        Dim winapp2Input As iniFile2

        If DownloadWinapp2 Then

            If Not checkOnline() Then

                setNextMenuHeaderText("Internet connection required to download winapp2.ini. Please check your connection.", printColor:=ConsoleColor.Red)
                Return

            End If

            gLog("Downloading winapp2.ini from GitHub")
            winapp2Input = getRemoteIniFile2(getWinappLink())

            If winapp2Input Is Nothing Then

                setNextMenuHeaderText("Failed to download winapp2.ini", printColor:=ConsoleColor.Red)
                Return

            End If

        Else

            If Not CC7PatcherFile1.Exists() Then

                setNextMenuHeaderText("winapp2.ini not found. Please select a valid file.", printColor:=ConsoleColor.Red)
                Return

            End If

            Dim loaded = CC7PatcherFile1.Load()
            If loaded Is Nothing Then Return
            winapp2Input = loaded

        End If

        ' Initialize menu output
        Dim menuOutput As New MenuSection
        Dim headerMsg = "CCleaner 7 Patcher"
        menuOutput.AddBoxWithText(headerMsg)
        gLog(headerMsg, buffr:=True, ascend:=True)

        ' Optionally trim winapp2.ini
        If TrimBeforePatching Then

            Dim trimMsg = "Trimming winapp2.ini..."
            menuOutput.AddColoredLine(trimMsg, ConsoleColor.Cyan)
            gLog(trimMsg)

            Dim wa2file As New winapp2file2(winapp2Input)
            Trim.trimFile(wa2file)
            winapp2Input = wa2file.ToIni()

            Dim trimCompleteMsg = $"Trimming complete: {wa2file.Count} entries remain"
            menuOutput.AddColoredLine(trimCompleteMsg, ConsoleColor.Green)
            gLog(trimCompleteMsg)

        End If

        ' Perform the patching
        patchCCleaner(winapp2Input, menuOutput)

        Dim completeMsg = "CCleaner 7 patching complete"
        menuOutput.AddBoxWithText(completeMsg)
        gLog(completeMsg, descend:=True, buffr:=True)

        menuOutput.AddAnyKeyPrompt()

        If Not SuppressOutput Then menuOutput.Print()
        crk()

    End Sub

    ''' <summary>
    ''' Patches ccleaner.ini with entries from winapp2.ini using Transmute's Add mode
    ''' </summary>
    '''
    ''' <param name="winapp2Input">
    ''' The winapp2.ini file to use as the source for patching
    ''' </param>
    '''
    ''' <param name="menuOutput">
    ''' The menu output section for logging
    ''' </param>
    Private Sub patchCCleaner(winapp2Input As iniFile2,
                              ByRef menuOutput As MenuSection)

        gLog("Beginning ccleaner.ini patching process", ascend:=True)

        Dim baseFile = CC7PatcherFile2.Load()
        If baseFile Is Nothing Then Return

        Dim outputFile = iniFile2.Empty(CC7PatcherFile3.Dir, CC7PatcherFile3.Name)

        Dim patchMsg = $"Patching {CC7PatcherFile2.Name} with entries from winapp2.ini"
        menuOutput.AddColoredLine(patchMsg, ConsoleColor.Yellow)
        gLog(patchMsg)

        Transmute.RemoteTransmute(baseFile, winapp2Input, outputFile, False, menuOutput, Transmute.TransmuteMode.Add)

        gLog("Patching process complete", descend:=True)

        Dim savedMsg = $"Patched file saved to {CC7PatcherFile3.Path()}"
        menuOutput.AddColoredLine(savedMsg, ConsoleColor.Green)
        gLog(savedMsg)

    End Sub

End Module
