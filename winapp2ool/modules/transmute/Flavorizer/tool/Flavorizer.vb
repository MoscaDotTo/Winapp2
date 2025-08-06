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
''' The Flavorizer module provides a user interface for applying "flavors" to ini files.
''' A flavor is a set of modifications that adapt a base ini file for specific use cases.
''' 
''' The flavorization process applies modifications in this specific order:
''' 1. Section Removal -> Remove entire sections
''' 2. Key Name Removal -> Remove keys by name matching
''' 3. Key Value Removal -> Remove keys by value and keytype matching  
''' 4. Section Replacement -> Replace entire sections
''' 5. Key Replacement -> Replace individual key values
''' 6. Section and Key Additions -> Add new sections and keys
''' 
''' If a modification file isn't present, that modification will be skipped 
''' 
''' This module wraps Transmute's Flavorize function with an intuitive UI for
''' managing the multiple correction files used in the flavorization process 
''' and also provides the ability to detect a group of flavor files within a 
''' particular directory for ease of configuration
''' 
''' </summary>
''' 
''' Docs last updated: 2025-08-05 | Code last updated: 2025-08-05
Public Module Flavorizer

    ''' <summary>
    ''' Handles command line arguments for the Flavorizer module <br />
    ''' Flavorizer args:
    ''' -nowinapp         : Disable processing as winapp2.ini format (default: true)
    ''' -autodetect       : Automatically detect a group of Flavor files in the target directory
    ''' </summary>
    ''' 
    ''' <remarks>
    ''' To refer -autodetect to a different directory than the current one, also provide 
    ''' -9d with the appropriate directory. <c> FlavorizerFile9 </c> holds the target directory 
    ''' for the auto detect function within its Dir property 
    ''' </remarks>
    ''' 
    ''' Docs last updated: 2025-08-01 | Code last updated: 2025-08-01
    Public Sub handleCmdLine()

        initDefaultFlavorizerSettings()

        invertSettingAndRemoveArg(FlavorizeAsWinapp, "-nowinapp")

        Dim autoDetect = False
        invertSettingAndRemoveArg(autoDetect, "-autodetect")

        If autoDetect Then

            ' If we're auto-detecting, we need to know the base file, the save file, and the 
            ' target directory for flavor files. we can fill in the rest from there 
            getFileAndDirParams({FlavorizerFile1, FlavorizerFile2, New iniFile,
                                New iniFile, New iniFile, New iniFile,
                                New iniFile, New iniFile, FlavorizerFile9})

            DetectFlavorFiles(FlavorizerFile9.Dir)

        Else

            getFileAndDirParams({FlavorizerFile1, FlavorizerFile2, FlavorizerFile3,
                                FlavorizerFile4, FlavorizerFile5, FlavorizerFile6,
                                FlavorizerFile7, FlavorizerFile8, FlavorizerFile9})

        End If

        If FlavorizerFile1.Name.Length > 0 Then initFlavorizer()

    End Sub

    ''' <summary>
    ''' Initializes the Flavorizer process and validates required files
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-08-01 | Code last updated: 2025-08-01
    Public Sub initFlavorizer()

        clrConsole()

        If Not enforceFileHasContent(FlavorizerFile1) Then Return

        LogAndPrint(4, $"Applying flavor to {FlavorizerFile1.Name}", trailr:=True, leadr:=True, ascend:=True, closeMenu:=True, arbitraryColor:=ConsoleColor.Yellow)

        Dim correctionFiles As New List(Of iniFile) From {FlavorizerFile3, FlavorizerFile4, FlavorizerFile5, FlavorizerFile6, FlavorizerFile7, FlavorizerFile8}
        Dim validFiles = correctionFiles.Where(Function(f) f.exists).Count()

        Dim hasValidFiles = Not validFiles = 0

        LogAndPrint(0, "No correction files specified - output will be identical to input", cond:=Not hasValidFiles, logCond:=Not hasValidFiles, colorLine:=True, useArbitraryColor:=True, arbitraryColor:=ConsoleColor.Yellow)
        LogAndPrint(0, $"Applying {validFiles} correction file(s)", colorLine:=True, useArbitraryColor:=True, arbitraryColor:=ConsoleColor.Cyan)

        performFlavorization()

        LogAndPrint(4, "Flavorization completed successfully", buffr:=True, descend:=True, conjoin:=True, arbitraryColor:=ConsoleColor.Yellow)
        print(0, anyKeyStr, closeMenu:=True)
        crk()

    End Sub

    ''' <summary>
    ''' Performs the actual flavorization using the Transmute.Flavorize function
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-08-01 | Code last updated: 2025-08-01
    Private Sub performFlavorization()

        gLog("Starting flavorization process", ascend:=True, buffr:=True)

        Dim baseFile = If(FlavorizerFile1.exists, FlavorizerFile1, Nothing)
        Dim saveFile = FlavorizerFile2
        Dim additionsFile = If(FlavorizerFile8.exists, FlavorizerFile8, Nothing)
        Dim sectionRemovalFile = If(FlavorizerFile3.exists, FlavorizerFile3, Nothing)
        Dim keyNameRemovalFile = If(FlavorizerFile4.exists, FlavorizerFile4, Nothing)
        Dim keyValueRemovalFile = If(FlavorizerFile5.exists, FlavorizerFile5, Nothing)
        Dim sectionReplacementFile = If(FlavorizerFile6.exists, FlavorizerFile6, Nothing)
        Dim keyReplacementFile = If(FlavorizerFile7.exists, FlavorizerFile7, Nothing)

        Flavorize(baseFile, saveFile,
                  additionsFile,
                  sectionRemovalFile, keyNameRemovalFile, keyValueRemovalFile,
                  sectionReplacementFile, keyReplacementFile,
                  FlavorizeAsWinapp, SuppressOutput)

        gLog("Flavorization process completed", descend:=True)

    End Sub

    ''' <summary>
    ''' Automatically detects and assigns flavor files based on standard naming conventions
    ''' </summary>
    ''' 
    ''' <param name="targetDirectory">
    ''' The directory to search for flavor files. If empty, uses the current directory.
    ''' </param>
    ''' 
    ''' <remarks>
    ''' - section_removals.ini -> FlavorizerFile3 (Section removal file)
    ''' - name_removals.ini -> FlavorizerFile4 (Key name removal file)  
    ''' - value_removals.ini -> FlavorizerFile5 (Key value removal file)
    ''' - section_replacements.ini -> FlavorizerFile6 (Section replacement file)
    ''' - key_replacements.ini -> FlavorizerFile7 (Key replacement file)
    ''' - additions.ini -> FlavorizerFile8 (Additions file)
    ''' </remarks>
    ''' 
    ''' Docs last updated: 2025-08-05 | Code last updated: 2025-08-05
    Public Sub DetectFlavorFiles(Optional targetDirectory As String = "")

        gLog("Starting automatic flavor file detection", ascend:=True)

        ' Use current directory if none specified
        If String.IsNullOrWhiteSpace(targetDirectory) Then targetDirectory = Environment.CurrentDirectory

        gLog($"Searching for flavor files in: {targetDirectory}")

        Dim flavorFiles As New Dictionary(Of String, iniFile) From {
            {"section_removals.ini", FlavorizerFile3},
            {"name_removals.ini", FlavorizerFile4},
            {"value_removals.ini", FlavorizerFile5},
            {"section_replacements.ini", FlavorizerFile6},
            {"key_replacements.ini", FlavorizerFile7},
            {"additions.ini", FlavorizerFile8}
        }

        Dim filesInTargetDir = My.Computer.FileSystem.GetFiles(targetDirectory)

        For Each kvp In flavorFiles

            For Each file In filesInTargetDir

                If Not file.Contains(kvp.Key) Then Continue For

                kvp.Value.Dir = targetDirectory
                kvp.Value.Name = file.Replace(targetDirectory & "\", "")
                Exit For

            Next

        Next

        updateFlavorFileSettings()

    End Sub

    ''' <summary>
    ''' Updates the settings for all Flavorizer files that have been assigned
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-08-05 | Code last updated: 2025-08-05
    Private Sub updateFlavorFileSettings()

        Dim flavorFiles As New Dictionary(Of String, iniFile) From {
            {NameOf(FlavorizerFile3), FlavorizerFile3},
            {NameOf(FlavorizerFile4), FlavorizerFile4},
            {NameOf(FlavorizerFile5), FlavorizerFile5},
            {NameOf(FlavorizerFile6), FlavorizerFile6},
            {NameOf(FlavorizerFile7), FlavorizerFile7},
            {NameOf(FlavorizerFile8), FlavorizerFile8}
        }

        For Each kvp In flavorFiles

            If kvp.Value.Name.Length = 0 Then Continue For

            updateSettings(NameOf(Flavorizer), $"{kvp.Key}_Name", kvp.Value.Name)
            updateSettings(NameOf(Flavorizer), $"{kvp.Key}_Dir", kvp.Value.Dir)
            updateSettings(NameOf(Flavorizer), NameOf(FlavorizerModuleSettingsChanged), tsInvariant(FlavorizerModuleSettingsChanged))
            FlavorizerModuleSettingsChanged = True

        Next

        settingsFile.overwriteToFile(settingsFile.toString, Not IsCommandLineMode AndAlso saveSettingsToDisk)

    End Sub

End Module
