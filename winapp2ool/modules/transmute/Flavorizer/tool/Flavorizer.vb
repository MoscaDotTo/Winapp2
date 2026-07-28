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
''' The Flavorizer module provides a user interface for applying "flavors" to ini files.
''' A flavor is a set of modifications that adapt a base ini file for specific use cases. <br /><br />
'''
''' The flavorization process applies modifications in this specific order:
'''
''' <list type="number">
''' <item> Section Removal — Remove entire sections </item>
''' <item> Key Name Removal — Remove keys by name matching </item>
''' <item> Key Value Removal — Remove keys by value and keytype matching </item>
''' <item> Section Replacement — Replace entire sections </item>
''' <item> Key Replacement — Replace individual key values </item>
''' <item> Section and Key Additions — Add new sections and keys </item>
''' </list>
'''
''' If a modification file is not present, that modification step will be skipped. <br /><br />
'''
''' This module wraps <c> Transmute.Flavorize </c> with an intuitive UI for managing the multiple
''' correction files used in the flavorization process, and also provides the ability to
''' auto-detect a group of flavor files within a target directory.
''' </summary>
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
    Public Sub handleCmdLine()

        initDefaultFlavorizerSettings()

        ' Detect mode flags first — they affect which file slots are bound
        Dim autoDetect = False

        If cmdargs.Contains("-autodetect") Then
            autoDetect = True
            cmdargs.Remove("-autodetect")
        End If

        Dim spec As New CliArgSpec("flavorize")
        spec.WithFlag("-nowinapp", Sub() FlavorizeAsWinapp = Not FlavorizeAsWinapp)

        If autoDetect Then
            spec.WithFile(1, FlavorizerFile1).WithFile(2, FlavorizerFile2).WithFile(9, FlavorizerFile9)
        Else
            spec.WithFile(1, FlavorizerFile1).WithFile(2, FlavorizerFile2).WithFile(3, FlavorizerFile3) _
                .WithFile(4, FlavorizerFile4).WithFile(5, FlavorizerFile5).WithFile(6, FlavorizerFile6) _
                .WithFile(7, FlavorizerFile7).WithFile(8, FlavorizerFile8).WithFile(9, FlavorizerFile9)
        End If

        spec.Parse()

        If autoDetect Then DetectFlavorFiles(FlavorizerFile9.Dir)

        If FlavorizerFile1.Name.Length > 0 Then initFlavorizer()

    End Sub

    ''' <summary>
    ''' Initializes the Flavorizer process and validates required files
    ''' </summary>
    Public Sub initFlavorizer()

        clrConsole()

        Dim baseFile = FlavorizerFile1.Load(FlavorizerModuleSettingsChanged, NameOf(Flavorizer), NameOf(FlavorizerFile1), NameOf(FlavorizerModuleSettingsChanged))

        ' Ordinarily, we would gate this but it's actually probably fine if the base file is empty 
        ' If Not enforceFileHasContent(baseFile) Then Return

        Dim applyingTxt = $"Applying flavor to {FlavorizerFile1.Name}"
        Dim output As New MenuSection
        output.AddBoxWithText(applyingTxt)

        Using gLogScope(applyingTxt)

            Dim correctionFiles As New List(Of iniFileChooser) From {FlavorizerFile3, FlavorizerFile4, FlavorizerFile5, FlavorizerFile6, FlavorizerFile7, FlavorizerFile8}
            Dim validFiles = correctionFiles.Where(Function(f) f.Name.Length > 0 AndAlso f.Exists()).Count()

            Dim hasValidFiles = Not validFiles = 0

            Dim noCorrectionsMsg = "No correction files specified - output will be identical to input"
            output.AddWarning(noCorrectionsMsg, Not hasValidFiles)
            gLog(noCorrectionsMsg, , Not hasValidFiles)

            Dim numFilesApplyingMsg = $"Applying {validFiles} correction file(s)"
            output.AddColoredLine(numFilesApplyingMsg, ConsoleColor.Cyan)
            gLog(numFilesApplyingMsg)

            performFlavorization(output)

            Dim finishedMsg = "Flavorization completed successfully"
            output.AddBoxWithText(finishedMsg)
            output.AddAnyKeyPrompt()
            gLog(finishedMsg)

        End Using

        If SuppressOutput Then Return

        output.Print()
        crk()

    End Sub

    ''' <summary>
    ''' Performs the actual flavorization using the Transmute.Flavorize function
    ''' </summary>
    '''
    ''' <param name="menuOutput">The <c> MenuSection </c> to which flavorization output lines are appended</param>
    Private Sub performFlavorization(ByRef menuOutput As MenuSection)

        Dim baseFile = iniFile2.FromFile(FlavorizerFile1.Path())
        Dim saveFile = iniFile2.Empty(FlavorizerFile2.Dir, FlavorizerFile2.Name)

        Dim additionsFile = If(FlavorizerFile8.Name.Length > 0 AndAlso FlavorizerFile8.Exists(), iniFile2.FromFile(FlavorizerFile8.Path()), Nothing)
        Dim sectionRemovalFile = If(FlavorizerFile3.Name.Length > 0 AndAlso FlavorizerFile3.Exists(), iniFile2.FromFile(FlavorizerFile3.Path()), Nothing)
        Dim keyNameRemovalFile = If(FlavorizerFile4.Name.Length > 0 AndAlso FlavorizerFile4.Exists(), iniFile2.FromFile(FlavorizerFile4.Path()), Nothing)
        Dim keyValueRemovalFile = If(FlavorizerFile5.Name.Length > 0 AndAlso FlavorizerFile5.Exists(), iniFile2.FromFile(FlavorizerFile5.Path()), Nothing)
        Dim sectionReplacementFile = If(FlavorizerFile6.Name.Length > 0 AndAlso FlavorizerFile6.Exists(), iniFile2.FromFile(FlavorizerFile6.Path()), Nothing)
        Dim keyReplacementFile = If(FlavorizerFile7.Name.Length > 0 AndAlso FlavorizerFile7.Exists(), iniFile2.FromFile(FlavorizerFile7.Path()), Nothing)

        Using gLogScope("Flavorizing")

            Flavorize(baseFile, saveFile, menuOutput,
                      additionsFile,
                      sectionRemovalFile, keyNameRemovalFile, keyValueRemovalFile,
                      sectionReplacementFile, keyReplacementFile,
                      FlavorizeAsWinapp)

        End Using

        gLog("Flavorization complete!")

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
    Public Sub DetectFlavorFiles(Optional targetDirectory As String = "")

        If String.IsNullOrWhiteSpace(targetDirectory) Then targetDirectory = Environment.CurrentDirectory

        Using gLogScope("Starting automatic flavor file detection")

            gLog($"Searching for flavor files in: {targetDirectory}")

            Dim flavorFiles As New Dictionary(Of String, iniFileChooser) From {
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

            FlavorizerModuleSettingsChanged = True
            SaveModule2(NameOf(Flavorizer), GetType(FlavorizerSettings))

        End Using

    End Sub

End Module
