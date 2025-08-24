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

Imports System.IO

''' <summary>
''' Combine is a winapp2ool module that takes all files with the ini extension within a target 
''' directory (including its subdirectories) and combines them into a single ini file. When 
''' duplicate section names are encountered, their unique keys are merged together into the output.
''' <br/> 
''' Files that cannot be parsed or have no sections are ignored.
''' <br/> <br/>
''' If the final combined output contains no sections, it will not be saved to disk
''' </summary>
''' 
''' Docs last updated: 2025-08-23 
Public Module Combine

    ''' <summary>
    ''' Stores the most recent Combine operation log for display in the log viewer
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-08-23 | Code last updated: 2025-08-23
    Public Property MostRecentCombineLog As String = ""

    ''' <summary>
    ''' The phrase that marks the beginning of a Combine operation in the global log
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-08-23 | Code last updated: 2025-08-23
    Public Const CombineLogStartPhrase As String = "Combining files from"

    ''' <summary>
    ''' The phrase that marks the end of a Combine operation in the global log
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-08-23 | Code last updated: 2025-08-23
    Public Const CombineLogEndPhrase As String = "Combination complete!"


    ''' <summary>
    ''' Handles command line arguments for the Combine module
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-08-20 | Code last updated: 2025-08-20
    Public Sub handleCmdLine()

        initDefaultCombineSettings()
        getFileAndDirParams({CombineFile1, New iniFile, CombineFile3})
        If CombineFile1.Dir.Length > 0 Then initCombine()

    End Sub

    ''' <summary>
    ''' Initializes the combine process and validates the target directory
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-08-20 | Code last updated: 2025-08-20
    Public Sub initCombine()

        clrConsole()


        If Not Directory.Exists(CombineFile1.Dir) Then

            setHeaderText($"Target directory not found: {CombineFile1.Dir}", True)
            Return

        End If

        Dim outputFile As New iniFile With {.Dir = CombineFile3.Dir, .Name = CombineFile3.Name}
        Dim CombineUserOutput As New MenuSection

        Dim outputHeader = $"{CombineLogStartPhrase} {CombineFile1.Dir}"
        gLog(outputHeader, ascend:=True, leadr:=True)
        CombineUserOutput.AddTopBorder()
        CombineUserOutput.AddLine(outputHeader, centered:=True)
        CombineUserOutput.AddDivider()

        processCombine(CombineUserOutput, CombineFile1.Dir, outputFile)

        CombineUserOutput.AddAnyKeyPrompt()

        clrConsole()
        CombineUserOutput.Print()

        crk()

    End Sub

    ''' <summary>
    ''' Processes all files in the target directory and combines them into a single INI file
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
    ''' <param name="outputFile">
    ''' The <c> iniFile </c> into which all other ini files found in 
    ''' <c> <paramref name="targetDir"/> </c> will be merged 
    ''' </param>
    ''' 
    ''' Docs last updated: 2025-08-20 | Code last updated: 2025-08-20
    Private Sub processCombine(outputMenu As MenuSection,
                               targetDir As String,
                         ByRef outputFile As iniFile)

        Dim allINIFiles = Directory.GetFiles(targetDir, "*.ini", SearchOption.AllDirectories).ToList()
        allINIFiles.Sort()

        Dim foundMsg = $"Found {allINIFiles.Count} files with ini extension in {targetDir}"
        gLog(foundMsg, indent:=True)
        outputMenu.AddLine(foundMsg)

        Dim processedCount = 0
        Dim validFileCount = 0

        For Each filePath In allINIFiles

            updateProgress(processedCount, allINIFiles.Count)

            ' Don't process the output file if it's in the target directory
            If filePath = outputFile.Path Then gLog($"Output file found in target directory, skipping: {filePath}", indent:=True) : Continue For

            Try

                attemptCombine(filePath, outputFile, processedCount, validFileCount, outputMenu)

            Catch ex As Exception

                handleCombineException(filePath, outputMenu, ex)

            End Try

            processedCount += 1

        Next

        gLog($"Processed {processedCount} files, {validFileCount} contained combinable sections", indent:=True)

        Dim outputIsEmpty = outputFile.Sections.Count = 0

        Dim emptyOutputMsg = $"No valid sections found to combine - {outputFile.Name} will not be saved"
        gLog(emptyOutputMsg, indent:=True, cond:=outputIsEmpty)
        outputMenu.AddWarning(emptyOutputMsg, condition:=outputIsEmpty)

        outputFile.overwriteToFile(outputFile.toString, Not outputIsEmpty)

        Dim combinedCountMsg = $"Combined {validFileCount} files into {outputFile.Name} with {outputFile.Sections.Count} sections"
        gLog(combinedCountMsg, indent:=True, cond:=Not outputIsEmpty)
        outputMenu.AddBlank()
        outputMenu.AddColoredLine(combinedCountMsg, ConsoleColor.Green, centered:=True, condition:=Not outputIsEmpty)
        outputMenu.AddBottomBorder()

        outputMenu.AddNewLine()
        outputMenu.AddBoxWithText(CombineLogEndPhrase)
        gLog(CombineLogEndPhrase, descend:=True)

        MostRecentCombineLog = getLogSliceFromGlobal(CombineLogStartPhrase, CombineLogEndPhrase)

    End Sub

    ''' <summary>
    ''' Handles logging of exceptions thrown during the Combine process
    ''' </summary>
    ''' 
    ''' <param name="filepath">
    ''' The path of the file during whose processing <c> <paramref name="ex"/> </c> was thrown
    ''' </param>
    ''' 
    ''' <param name="outputMenu">
    ''' The <c> MenuSection </c> containing the Combine module's output as it will be displayed
    ''' to the user
    ''' </param>
    ''' 
    ''' <param name="ex">The exception thrown while Combine was processing 
    ''' <c> <paramref name="filepath"/> </c>
    ''' </param>
    ''' 
    ''' Docs last updated: 2025-08-23 | Code last updated: 2025-08-23
    Private Sub handleCombineException(filepath As String,
                                 ByRef outputMenu As MenuSection,
                                       ex As Exception)

        Dim errMsg = $"Error processing file: {filepath}"

        gLog($"{errMsg}: {ex.Message}", indent:=True)
        outputMenu.AddWarning(errMsg)
        outputMenu.AddWarning($"Check the winapp2ool log for more information: {GlobalLogFile.Path}")

        saveGlobalLog()

    End Sub

    ''' <summary>
    ''' Tries to combine a single ini file into the output file, logging the success of this operation
    ''' </summary>
    ''' 
    ''' <param name="filepath">
    ''' The path of a particular ini file to be combined 
    ''' </param>
    ''' 
    ''' <param name="outputFile">
    ''' The output file into which <c> <paramref name="filepath"/> </c> will be combined
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
    ''' 
    ''' Docs last updated: 2025-08-23 | Code last updated: 2025-08-23
    Private Sub attemptCombine(filepath As String,
                         ByRef outputFile As iniFile,
                         ByRef processedCount As Integer,
                         ByRef validFileCount As Integer,
                         ByRef outputMenu As MenuSection)

        Dim currentFile As New iniFile(Path.GetDirectoryName(filepath), Path.GetFileName(filepath))
        currentFile.init()

        If currentFile.Sections.Count = 0 Then

            gLog($"Skipping file with no sections: {Path.GetFileName(filepath)}", indent:=True)
            processedCount += 1
            Return

        End If

        mergeFileIntoOutput(currentFile, outputFile)

        validFileCount += 1

        Dim processedMsg = $"Processed: {Path.GetFileName(filepath)} ({currentFile.Sections.Count} sections)"
        gLog(processedMsg, indent:=True)
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
    ''' 
    ''' Docs last updated: 2025-08-23 | Code last updated: 2025-08-23
    Private Sub updateProgress(processedCount As Integer,
                               totalCount As Integer)

        Console.SetCursorPosition(0, 0)
        cwl($"Combining files... ({processedCount}/{totalCount})")

    End Sub

    ''' <summary>
    ''' Merges the sections from a source file into the combined output file,
    ''' merging keys when sections with the same name already exist
    ''' </summary>
    ''' 
    ''' <param name="sourceFile">
    ''' The source file whose sections will be merged into the output
    ''' </param>
    ''' 
    ''' <param name="outputFile">
    ''' The combined output file that will receive the merged sections
    ''' </param>
    ''' 
    ''' Docs last updated: 2025-08-20 | Code last updated: 2025-08-20
    Private Sub mergeFileIntoOutput(sourceFile As iniFile, ByRef outputFile As iniFile)

        For Each sectionName In sourceFile.Sections.Keys

            Dim sourceSection = sourceFile.Sections(sectionName)

            If outputFile.Sections.ContainsKey(sectionName) Then AddUniqueKeys(sourceSection, outputFile, sectionName) : Continue For

            outputFile.Sections.Add(sectionName, sourceSection)
            gLog($"Added new section: [{sectionName}] ({sourceSection.Keys.KeyCount} keys)", indent:=True, indAmt:=4)

        Next

    End Sub

    ''' <summary>
    ''' Merges keys from a source section into an existing section in the output file, preventing 
    ''' any keys with duplicate names and values from being added
    ''' <br />
    ''' Note: Matching values with unlike names will still be added
    ''' </summary>
    ''' 
    ''' <param name="sourceSection">
    ''' The <c> iniSection </c> whose contents will be merged into the output file
    ''' </param>
    ''' 
    ''' <param name="outputFile">
    ''' The combined output file into which keys from <c> sourceSection </c> will be merged
    ''' </param>
    ''' 
    ''' <param name="sectionName">
    ''' The name of the current section being processed
    ''' </param>
    ''' 
    ''' Docs last updated: 2025-08-23 | Code last updated: 2025-08-23
    Private Sub AddUniqueKeys(sourceSection As iniSection,
                              outputFile As iniFile,
                              sectionName As String)

        Dim existingSection = outputFile.Sections(sectionName)

        Dim extantKeys As New HashSet(Of String)(existingSection.Keys.Keys.Select(Function(k) $"{k.Name.ToLowerInvariant()}={k.Value.ToLowerInvariant()}"))
        Dim addedKeyCount = 0
        Dim skippedKeyCount = 0

        For Each sourceKey In sourceSection.Keys.Keys

            Dim keyExists = extantKeys.Contains($"{sourceKey.Name.ToLowerInvariant()}={sourceKey.Value.ToLowerInvariant()}")

            existingSection.Keys.add(sourceKey, Not keyExists)

            If keyExists Then

                skippedKeyCount += 1

            Else

                addedKeyCount += 1

            End If

            gLog($"Skipped duplicate key: {sourceKey.Name}", indent:=True, indAmt:=6, cond:=keyExists)

        Next

    End Sub

    ''' <summary>
    ''' Facilitates combining files from outside the module's UI
    ''' </summary>
    ''' 
    ''' <param name="targetDirectory">
    ''' The directory containing files to be combined
    ''' </param>
    ''' 
    ''' <param name="outputFile">
    ''' The file location where the combined output will be saved
    ''' </param>
    ''' 
    ''' <returns>
    ''' The resulting combined output file if successful, or 
    ''' Empty <c> iniFile </c> if the target directory does not exist or otherwise lacks valid ini files
    ''' </returns>
    ''' 
    ''' Docs last updated: 2025-08-20 | Code last updated: 2025-08-20
    Public Function RemoteCombine(targetDirectory As String,
                                  outputFile As iniFile) As iniFile

        If Not Directory.Exists(targetDirectory) Then

            gLog($"Target directory not found: {targetDirectory}")
            Return New iniFile

        End If

        processCombine(New MenuSection, targetDirectory, outputFile)

        Return outputFile

    End Function

End Module
