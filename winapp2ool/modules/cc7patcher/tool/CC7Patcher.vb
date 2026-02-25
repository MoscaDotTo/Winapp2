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

''' <summary>
''' CC7Patcher is a winapp2ool module that patches ccleaner.ini with entries from winapp2.ini
''' to enable CCleaner 7 compatibility
''' <br /><br />
''' The module can optionally download the latest winapp2.ini from GitHub and trim it
''' before applying the patches to ccleaner.ini
''' </summary>
''' 
''' Docs last updated: 2025-11-19 | Code last updated: 2025-11-19
Public Module CC7Patcher

    ''' <summary>
    ''' Handles the command line arguments for CC7Patcher
    ''' </summary>
    ''' 
    ''' <remarks>
    ''' CC7Patcher args:
    ''' -d              : Download the latest winapp2.ini from GitHub (enabled by default)
    ''' -nodownload     : Disable downloading winapp2.ini
    ''' -trim           : Trim winapp2.ini before patching
    ''' </remarks>
    ''' 
    ''' Docs last updated: 2025-11-19 | Code last updated: 2025-11-19
    Public Sub handleCmdLine()

        initDefaultCC7PatcherSettings()

        ' Handle download flag (default is True, so -nodownload disables it)
        invertSettingAndRemoveArg(DownloadWinapp2, "-nodownload")

        ' Handle trim flag
        invertSettingAndRemoveArg(TrimBeforePatching, "-trim")

        ' Get file parameters
        getFileAndDirParams({CC7PatcherFile1, CC7PatcherFile2, CC7PatcherFile3})

        ' Initialize the patching process
        initCC7Patcher()

    End Sub

    ''' <summary>
    ''' Initializes the CC7Patcher process
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-11-19 | Code last updated: 2025-11-19
    Public Sub initCC7Patcher()

        clrConsole()

        ' Validate ccleaner.ini exists
        If Not enforceFileHasContent(CC7PatcherFile2) Then Return

        ' Handle winapp2.ini acquisition
        Dim winapp2Input As iniFile

        If DownloadWinapp2 Then

            ' Check for online connectivity
            If Not checkOnline() Then
                setHeaderText("Internet connection required to download winapp2.ini. Please check your connection.", True)
                Return
            End If

            gLog("Downloading winapp2.ini from GitHub")
            winapp2Input = getRemoteIniFile(getWinappLink())

            If winapp2Input Is Nothing Then
                setHeaderText("Failed to download winapp2.ini", True)
                Return
            End If

        Else

            ' Use local file
            If Not enforceFileHasContent(CC7PatcherFile1) Then Return
            winapp2Input = CC7PatcherFile1

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

            Dim wa2file As New winapp2file(winapp2Input)
            Trim.trimFile(wa2file)

            ' Create a new iniFile with the trimmed content
            winapp2Input = wa2file.toIni

            Dim trimCompleteMsg = $"Trimming complete: {wa2file.count} entries remain"
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
    ''' 
    ''' Docs last updated: 2025-11-19 | Code last updated: 2025-11-19
    Private Sub patchCCleaner(winapp2Input As iniFile, ByRef menuOutput As MenuSection)

        gLog("Beginning ccleaner.ini patching process", ascend:=True)

        ' Load ccleaner.ini
        CC7PatcherFile2.init()

        ' Log patch details
        Dim patchMsg = $"Patching {CC7PatcherFile2.Name} with entries from winapp2.ini"
        menuOutput.AddColoredLine(patchMsg, ConsoleColor.Yellow)
        gLog(patchMsg)

        ' Use Transmute's RemoteTransmute with Add mode to patch ccleaner.ini
        Transmute.RemoteTransmute(
            CC7PatcherFile2,
            winapp2Input,
            CC7PatcherFile3,
            False, ' Not a winapp2.ini file
            menuOutput,
            Transmute.TransmuteMode.Add
        )

        gLog("Patching process complete", descend:=True)

        ' Save the patched file
        Dim savedMsg = $"Patched file saved to {CC7PatcherFile3.Path}"
        menuOutput.AddColoredLine(savedMsg, ConsoleColor.Green)
        gLog(savedMsg)

    End Sub

    ''' <summary>
    ''' Facilitates patching ccleaner.ini from outside the module
    ''' </summary>
    ''' 
    ''' <param name="winapp2File">
    ''' The winapp2.ini file to use as source
    ''' </param>
    ''' 
    ''' <param name="ccleanerFile">
    ''' The ccleaner.ini file to patch
    ''' </param>
    ''' 
    ''' <param name="outputFile">
    ''' The output file path
    ''' </param>
    ''' 
    ''' <param name="shouldDownload">
    ''' Whether to download winapp2.ini
    ''' </param>
    ''' 
    ''' <param name="shouldTrim">
    ''' Whether to trim winapp2.ini before patching
    ''' </param>
    ''' 
    ''' Docs last updated: 2025-11-19 | Code last updated: 2025-11-19
    Public Sub remotePatch(winapp2File As iniFile,
                          ccleanerFile As iniFile,
                          outputFile As iniFile,
                          shouldDownload As Boolean,
                          shouldTrim As Boolean)

        CC7PatcherFile1 = winapp2File
        CC7PatcherFile2 = ccleanerFile
        CC7PatcherFile3 = outputFile
        DownloadWinapp2 = shouldDownload
        TrimBeforePatching = shouldTrim

        initCC7Patcher()

    End Sub

End Module