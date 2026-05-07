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

Imports System.IO

''' <summary>
''' Combine is a winapp2ool module that takes all files with the ini extension within a target
''' directory (including its subdirectories) and combines them into a single ini file. When
''' duplicate section names are encountered, their unique keys are merged together into the output.
''' <br />
''' Files that cannot be parsed or have no sections are ignored.
''' <br /><br />
''' If the final combined output contains no sections, it will not be saved to disk
''' </summary>
Public Module Combine

    ''' <summary>
    ''' Stores the most recent Combine operation log for display in the log viewer
    ''' </summary>
    Public Property MostRecentCombineLog As String = ""

    ''' <summary>
    ''' The phrase that marks the beginning of a Combine operation in the global log
    ''' </summary>
    Public Const CombineLogStartPhrase As String = "Combining files from"

    ''' <summary>
    ''' The phrase that marks the end of a Combine operation in the global log
    ''' </summary>
    Public Const CombineLogEndPhrase As String = "Combination complete!"

    ''' <summary>
    ''' Handles command line arguments for the Combine module
    ''' </summary>
    '''
    ''' <remarks>
    ''' File arguments:
    ''' <list type="bullet">
    ''' <item><c> -1d path </c> — Set the target directory</item>
    ''' <item><c> -3d path </c> — Set the output directory</item>
    ''' <item><c> -3f name </c> — Set the output file name</item>
    ''' </list>
    ''' </remarks>
    Public Sub handleCmdLine()

        InitDefaultCombineSettings()

        Dim spec As New CliArgSpec(NameOf(Combine))
        spec.WithFile(1, CombineFile1, "targetdir") _
            .WithFile(3, CombineFile3, "output") _
            .Parse()

        initCombine(CombineFile1.Dir, CombineFile3)

    End Sub

    ''' <summary>
    ''' Initializes the combine process, validates the target directory, and displays the results
    ''' </summary>
    '''
    ''' <param name="targetDir">
    ''' The directory to scan for <c> .ini </c> files
    ''' </param>
    '''
    ''' <param name="outputFile">
    ''' The <c> iniFileChooser </c> describing where the combined output will be saved
    ''' </param>
    Public Sub initCombine(targetDir As String,
                           outputFile As iniFileChooser)

        clrConsole()

        If Not Directory.Exists(targetDir) Then

            setNextMenuHeaderText("Target directory not found. Please select a valid directory.", printColor:=ConsoleColor.Red)
            Return

        End If

        Dim CombineUserOutput As New MenuSection
        Dim combinedOutput As iniFile2 = iniFile2.Empty(outputFile.Dir, outputFile.Name)

        processCombine(CombineUserOutput, targetDir, combinedOutput)

        CombineUserOutput.AddAnyKeyPrompt()

        clrConsole()
        CombineUserOutput.Print()

        crk()

    End Sub

    ''' <summary>
    ''' Processes all files in the target directory and combines them into a single ini file
    ''' </summary>
    '''
    ''' <param name="outputMenu">
    ''' The <c> MenuSection </c> containing the module's output to be displayed to the user
    ''' </param>
    '''
    ''' <param name="targetDir">
    ''' The parent directory potentially containing the ini files to combine
    ''' </param>
    '''
    ''' <param name="combinedOutput">
    ''' The <c> iniFile2 </c> into which all other ini files found in
    ''' <paramref name="targetDir"/> will be merged
    ''' </param>
    Private Sub processCombine(outputMenu As MenuSection,
                                targetDir As String,
                          ByRef combinedOutput As iniFile2)

        Dim allINIFiles = Directory.GetFiles(targetDir, "*.ini", SearchOption.AllDirectories).ToList()
        allINIFiles.Sort()

        Dim outputHeader = $"{CombineLogStartPhrase} {targetDir}"
        Using gLogScope(outputHeader)

            outputMenu.AddTopBorder()
            outputMenu.AddLine(outputHeader, centered:=True)
            outputMenu.AddDivider()

            Dim foundMsg = $"Found {allINIFiles.Count} files with ini extension in {targetDir}"
            Using gLogScope(foundMsg)
                outputMenu.AddLine(foundMsg)

                Dim processedCount = 0
                Dim validFileCount = 0

                For Each filePath In allINIFiles

                    updateProgress(processedCount, allINIFiles.Count)

                    ' Don't process the output file if it's in the target directory
                    If filePath = combinedOutput.Path() Then gLog($"Output file found in target directory, skipping: {filePath}") : Continue For

                    Try

                        attemptCombine(filePath, combinedOutput, processedCount, validFileCount, outputMenu)

                    Catch ex As Exception

                        handleCombineException(filePath, outputMenu, ex)

                    End Try

                    processedCount += 1

                Next

                gLog($"Processed {processedCount} files, {validFileCount} contained combinable sections")

                Dim outputIsEmpty = combinedOutput.Count = 0

                Dim emptyOutputMsg = $"No valid sections found to combine - {combinedOutput.Name} will not be saved"
                gLog(emptyOutputMsg, cond:=outputIsEmpty)
                outputMenu.AddWarning(emptyOutputMsg, condition:=outputIsEmpty)

                combinedOutput.OverwriteToFile(combinedOutput.ToString(), Not outputIsEmpty)

                Dim combinedCountMsg = $"Combined {validFileCount} files into {combinedOutput.Name} with {combinedOutput.Count} sections"
                gLog(combinedCountMsg, cond:=Not outputIsEmpty)
                outputMenu.AddBlank()
                outputMenu.AddColoredLine(combinedCountMsg, ConsoleColor.Green, centered:=True, condition:=Not outputIsEmpty)
                outputMenu.AddBottomBorder()

            End Using

            outputMenu.AddNewLine()
            outputMenu.AddBoxWithText(CombineLogEndPhrase)
            gLog(CombineLogEndPhrase)

        End Using

        MostRecentCombineLog = getLogSliceFromGlobal(CombineLogStartPhrase, CombineLogEndPhrase)

    End Sub

    ''' <summary>
    ''' Handles logging of exceptions thrown during the Combine process. <br />
    ''' Uses a broad <c> Exception </c> catch intentionally — each file is processed independently
    ''' and a parse failure on one file should not abort the remaining files.
    ''' </summary>
    '''
    ''' <param name="filepath">
    ''' The path of the file during whose processing <paramref name="ex"/> was thrown
    ''' </param>
    '''
    ''' <param name="outputMenu">
    ''' The <c> MenuSection </c> containing the Combine module's output as it will be displayed
    ''' to the user
    ''' </param>
    '''
    ''' <param name="ex">
    ''' The exception thrown while Combine was processing <paramref name="filepath"/>
    ''' </param>
    Private Sub handleCombineException(filepath As String,
                                 ByRef outputMenu As MenuSection,
                                       ex As Exception)

        Dim errMsg = $"Error processing file: {filepath}"

        gLog($"{errMsg}: {ex.Message}")
        outputMenu.AddWarning(errMsg)
        outputMenu.AddWarning($"Check the winapp2ool log for more information: {GlobalLogFile.Path()}")

        saveGlobalLog()

    End Sub

    ''' <summary>
    ''' Tries to combine a single ini file into the combined output, logging the success of this operation
    ''' </summary>
    '''
    ''' <param name="filepath">
    ''' The path of a particular ini file to be combined
    ''' </param>
    '''
    ''' <param name="combinedOutput">
    ''' The output file into which <paramref name="filepath"/> will be combined
    ''' </param>
    '''
    ''' <param name="processedCount">
    ''' The number of files that have been processed so far
    ''' </param>
    '''
    ''' <param name="validFileCount">
    ''' The number of files that have been successfully combined so far
    ''' </param>
    '''
    ''' <param name="outputMenu">
    ''' The <c> MenuSection </c> containing the Combine module's output as it will be displayed
    ''' to the user
    ''' </param>
    Private Sub attemptCombine(filepath As String,
                         ByRef combinedOutput As iniFile2,
                         ByRef processedCount As Integer,
                         ByRef validFileCount As Integer,
                         ByRef outputMenu As MenuSection)

        Dim currentFile As iniFile2 = iniFile2.FromFile(filepath)

        If currentFile.Count = 0 Then

            gLog($"Skipping file with no sections: {Path.GetFileName(filepath)}", buffr:=True)
            processedCount += 1
            Return

        End If

        Dim processingMsg = $"Processing: {Path.GetFileName(filepath)} ({currentFile.Count} sections)"

        Using gLogScope(processingMsg)

            mergeFileIntoOutput(currentFile, combinedOutput)

            validFileCount += 1


        End Using

        Dim processedMsg = $"Processed: {Path.GetFileName(filepath)} ({currentFile.Count} sections)"
        gLog(processedMsg, buffr:=True)
        outputMenu.AddLine(processedMsg)

    End Sub

    ''' <summary>
    ''' Updates the console with the current progress of the combination process while it runs
    ''' </summary>
    '''
    ''' <param name="processedCount">
    ''' The number of files that have been processed so far
    ''' </param>
    '''
    ''' <param name="totalCount">
    ''' The total number of files to be processed
    ''' </param>
    Private Sub updateProgress(processedCount As Integer,
                               totalCount As Integer)

        Console.SetCursorPosition(0, 0)
        cwl($"Combining files... ({processedCount}/{totalCount})")

    End Sub

    ''' <summary>
    ''' Merges the sections from a source file into the combined output,
    ''' merging keys when sections with the same name already exist
    ''' </summary>
    '''
    ''' <param name="sourceFile">
    ''' The source file whose sections will be merged into the output
    ''' </param>
    '''
    ''' <param name="combinedOutput">
    ''' The combined output file that will receive the merged sections
    ''' </param>
    Private Sub mergeFileIntoOutput(sourceFile As iniFile2, ByRef combinedOutput As iniFile2)

        For Each sourceSection In sourceFile

            If combinedOutput.Contains(sourceSection.Name) Then AddUniqueKeys(sourceSection, combinedOutput, sourceSection.Name) : Continue For

            combinedOutput.AddSection(sourceSection)
            gLog($"Added new section: [{sourceSection.Name}] ({sourceSection.Keys.Count} keys)")

        Next

    End Sub

    ''' <summary>
    ''' Merges keys from a source section into an existing section in the output, preventing
    ''' any keys with duplicate names and values from being added. <br />
    ''' Note: Matching values with unlike names will still be added
    ''' </summary>
    '''
    ''' <param name="sourceSection">
    ''' The <c> iniSection2 </c> whose contents will be merged into the output
    ''' </param>
    '''
    ''' <param name="combinedOutput">
    ''' The combined output into which keys from <paramref name="sourceSection"/> will be merged
    ''' </param>
    '''
    ''' <param name="sectionName">
    ''' The name of the current section being processed
    ''' </param>
    Private Sub AddUniqueKeys(sourceSection As iniSection2,
                               combinedOutput As iniFile2,
                               sectionName As String)

        Dim existingSection = combinedOutput.GetSection(sectionName)

        Dim extantKeys As New HashSet(Of String)(existingSection.Keys.Select(Function(k) $"{k.Name.ToLowerInvariant()}={k.Value.ToLowerInvariant()}"))
        Dim addedKeyCount = 0
        Dim skippedKeyCount = 0

        For Each sourceKey In sourceSection.Keys

            Dim keyExists = extantKeys.Contains($"{sourceKey.Name.ToLowerInvariant()}={sourceKey.Value.ToLowerInvariant()}")

            If Not keyExists Then

                existingSection.AddKey(sourceKey)
                addedKeyCount += 1
                gLog($"Added {sourceKey.Name} to {existingSection.GetFullName}")

            Else

                skippedKeyCount += 1
                gLog($"Skipped duplicate key in {sourceSection.GetFullName}: {sourceKey.Name}")

            End If

        Next

    End Sub

    ''' <summary>
    ''' Facilitates combining files from outside the module's UI.
    ''' Returns the combined <c> iniFile2 </c> after processing; the caller may inspect the
    ''' result but does not need to save it — <c> processCombine </c> writes to disk automatically.
    ''' </summary>
    '''
    ''' <param name="targetDirectory">
    ''' The directory containing files to be combined
    ''' </param>
    '''
    ''' <param name="outputDir">
    ''' The directory component of the output file path
    ''' </param>
    '''
    ''' <param name="outputName">
    ''' The filename component of the output file path
    ''' </param>
    '''
    ''' <returns>
    ''' The resulting combined <c> iniFile2 </c> if successful, or an empty <c> iniFile2 </c>
    ''' if the target directory does not exist or otherwise lacks valid ini files
    ''' </returns>
    Public Function RemoteCombine(targetDirectory As String,
                                   outputDir As String,
                                   outputName As String) As iniFile2

        If Not Directory.Exists(targetDirectory) Then

            gLog($"Target directory not found: {targetDirectory}")
            Return iniFile2.Empty(outputDir, outputName)

        End If

        Dim combinedOutput As iniFile2 = iniFile2.Empty(outputDir, outputName)
        processCombine(New MenuSection, targetDirectory, combinedOutput)

        Return combinedOutput

    End Function

End Module
