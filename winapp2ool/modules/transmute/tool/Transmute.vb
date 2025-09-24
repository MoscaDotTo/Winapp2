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
''' <c> Transmute </c> (formerly <c> Merge </c>) is a winapp2ool module which provides the ability
''' to modify an <c> iniFile </c> object using the contents of a separate <c> iniFile </c> object
''' with conflict resolution at different levels of granularity. <br /><br />
'''
''' In the parlance of Transmute, there are two files of interest: <br />
''' The 'base' file and the 'source' file <br /> <br/>
''' The 'base' file is the one whose content will be modified by the operation. That is, this file
''' will be either added to, removed from, or have some or all of its content modified <br />
'''
''' Likewise, the 'source' file is the one from which the modifications will be provided. That is,
''' this file will have its content added to, removed from, or replace content in the 'base' file
'''
''' <br /><br />
'''
''' Transmute modes: <br />
'''
''' <list>
'''
''' <item>
''' <b> Add </b>(Default transmute mode)
'''
''' <description>
''' Adds sections to the base file from the source file <br />
''' If a section from the source file exists in the base file, the individual keys will be added
''' </description>
''' </item>
'''
''' <item>
''' <b> Replace </b>
'''
''' <description>
''' Replace has two sub modes: BySection and ByKey <br/><br/>
'''
''' ByKey (Default replace mode): Replaces the value of keys in the base file with values from the
''' source file based on their Name <br/><br/>
'''
''' BySection: Replaces entire sections in the base file with the section of the same name 
''' in the source file. <br/><br/>
'''
''' Note: Replace does nothing with sections from the source file not found in the base file
''' </description>
''' </item>
'''
''' <item>
''' Remove 
'''
''' Remove has two sub modes: BySection and ByKey <br/><br/>
'''
''' ByKey (Default remove mode): Removes keys from individual sections within the base file <br />
''' Keys must be provided within a section in the source file and will be removed from the 
''' corresponding section of the same name in the base file <br/> 
'''
''' Remove Key Modes: <br/><br/>
''' <b> ByName </b>: Removes keys based on matching key names only <br />
''' Removing a key with this mode requires knowledge of its Name. This works best for unnumbered
''' keys (in the context of winapp2.ini, keys like Section, or unnumbered Detection keys)
''' but also works for numbered keys if you know the number. For numbered keys it is generally
''' best to try using the ByValue remove mode <br />
''' 
''' <b> ByValue </b>: Removes keys based on matching key values. The KeyType (key name 
''' without numbers) must also match. Works best for numbered keys whose numbers may change 
''' or be unclear when invoking Transmute
''' </item>
''' </list>
''' 
''' <br /><br />
''' Additionally, Transmute provides an important function called Flavorize. Flavorize is used to 
''' apply "flavors" to an ini file. In the context on winapp2.ini, there has historically existed 
''' at least one flavor: the "non-ccleaner" file. Flavorize automates and also democratizes what has 
''' historically been a very labor intensive process: Maintaining a concise set of transformations
''' on the base winapp2.ini that can be used to adapt it to different purposes. The flavorization 
''' process always takes in a base ini file and then applies operations in the following order: <br />
''' <br />
''' Section Removal -> Key Name Removal -> Key Value Removal -> Section Replacement ->
''' Key Replacement -> Section and Key Additions
''' <br/>
''' </summary>
''' <remarks> 
''' <b> Remarks: </b> <br />
''' Transmute replaces the Merge module and has has three modes which are wholly discrete from 
''' one another. This is a breaking change from old versions of Merge prior to 2025, which 
''' always add entries which don't exist and perform conflict resolution as part of that process.
''' <br /><br />
''' As the functionality of Merge evolved, it no longer felt appropriate to refer to its output as
''' the result of a "merger" necessarily. We now consider the resulting output a Transmutation. 
''' Nevertheless, this is still spiritually the Merge module 
''' </remarks>
''' 
''' Docs last updated: 2025-07-25 | Code last updated: 2025-07-25
Public Module Transmute

    ''' <summary>
    ''' Enum representing the different primary modes of modifying the base file
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-07-25 | Code last updated: 2025-07-25
    Public Enum TransmuteMode

        ''' <summary>
        ''' Add the content of the source file into the base file <br />
        ''' For sections in the source file not found in the base file, they will be added as is <br /> 
        ''' sections already existing in the base file will have the keys from the source file added
        ''' </summary>
        Add = 0

        ''' <summary>
        ''' Overwrite the content of individual sections or keys in the base file with 
        ''' content from the source file. <br/>
        ''' Whether replacements are done by section or by key is is controlled by a separate
        ''' enum <c> ReplaceMode </c>
        ''' </summary>
        Replace = 1

        ''' <summary>
        ''' Remove from the base file any keys or sections found in the source file. <br />
        ''' Whether replacements are done by section or by key is controlled by a separate enum
        ''' called <c> RemoveMode </c> <br />
        ''' Whether key removals are done by Name or by Value is controlled by a separate enum 
        ''' called <c> RemoveKeyMode </c>
        ''' </summary>
        Remove = 2

    End Enum

    ''' <summary>
    ''' Enum representing the granularity of replace operations
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-07-25 | Code last updated: 2025-07-25
    Public Enum ReplaceMode

        ''' <summary>
        ''' Replaces entire sections when collisions occur <br />
        ''' 
        ''' This means that if a section exists in the base file, it will be replaced entirely
        ''' with the section from the source file. Section names must match exactly (case sensitive)
        ''' </summary>
        BySection = 0

        ''' <summary>
        ''' Replace individual keys when collisions occur <br />
        ''' 
        ''' This means that if a key exists in the base file, its value will be replaced with the
        ''' value from the source file. The key must exist in both files and have the same Name 
        ''' </summary>
        ByKey = 1

    End Enum

    ''' <summary>
    ''' Enum representing the granularity of remove operations
    ''' </summary>
    Public Enum RemoveMode

        ''' <summary>
        ''' Remove entire sections when collisions occur <br />
        ''' 
        ''' This means that if a section exists in the base file, it will be removed entirely
        ''' if it also exists in the source file. Section names must match exactly (case sensitive)
        ''' </summary>
        BySection = 0

        ''' <summary>
        ''' Remove individual keys when collisions occur <br />
        ''' 
        ''' This means that if a key exists in the base file, it will be removed if it also exists
        ''' in the source file. The key must exist in both files and have the same KeyType (key name
        ''' without numbers). Section names must match exactly (case sensitive)
        ''' </summary>
        ByKey = 1

    End Enum

    ''' <summary>
    ''' Enum representing the criteria for removing keys
    ''' </summary>
    Public Enum RemoveKeyMode

        ''' <summary>
        ''' Remove keys if they have the same Name (value to the left of the equals sign)
        ''' </summary>
        ByName = 0

        ''' <summary>
        ''' Remove keys if they have the same KeyType (Name with numbers removed) and Value 
        ''' </summary>
        ByValue = 1

    End Enum

    ''' <summary> 
    ''' Handles the commandline args for Transmute 
    ''' </summary>
    ''' 
    ''' <remarks> 
    ''' Transmute args: <br /> 
    ''' Primary modes: <br />
    ''' 
    ''' -add            : Add mode (default) <br />
    ''' -replace        : Replace mode <br />
    ''' -remove         : Remove mode <br />
    ''' 
    ''' Sub-mode: shared by remove/replace since only one can be done at a time <br />
    '''
    ''' -bysection      : Replace/Remove by section <br />
    ''' -bykey          : Replace/Remove by key (default) <br />
    ''' 
    ''' Remove key criteria: <br />
    ''' -byname         : Remove keys by name (default) <br />
    ''' -byvalue        : Remove keys by value <br />
    ''' 
    ''' Winapp2.ini syntax correction <br/>
    ''' -dontlint       : do not save output with winapp2.ini formatting 
    ''' 
    ''' Preset base file choices <br />
    ''' 
    ''' -r              : removed entries.ini  <br />
    ''' -c              : custom.ini  <br />
    ''' -w              : winapp3.ini <br />
    ''' -a              : archived entries.ini <br /> 
    ''' -b              : browsers.ini
    ''' </remarks>
    Public Sub handleCmdLine()

        initDefaultTransmuteSettings()

        ' Primary mode (Default: Add)
        Dim mode As TransmuteMode = TransmuteMode.Add
        ' Sub mode (Default: By Key)
        Dim byKeyMode As Boolean = True
        ' Key removal mode (Default: By Name)
        Dim RemoveByName = True
        ' Winapp2.ini style linting of output (Default: True)
        Dim isWinapp = True

        ' Primary mode
        Dim Modes = New Dictionary(Of String, TransmuteMode) From {
            {"add", TransmuteMode.Add},
            {"replace", TransmuteMode.Replace},
            {"remove", TransmuteMode.Remove}
        }

        ' Preset source file choices
        Dim PresetFileNames = New Dictionary(Of String, String) From {
            {"r", "Removed Entries.ini"},
            {"c", "Custom.ini"},
            {"w", "winapp3.ini"},
            {"a", "Archived Entries.ini"},
            {"b", "browsers.ini"}
        }

        ' These flags modify the default behavior 
        Dim Flags = New Dictionary(Of String, Boolean) From {
            {"bysection", byKeyMode},
            {"byvalue", RemoveByName},
            {"dontlint", isWinapp}
        }

        For Each arg In cmdargs.ToList()

            Dim trimmedArg = arg.ToLowerInvariant().TrimStart("-"c)

            Select Case True

                Case Modes.ContainsKey(trimmedArg)

                    mode = Modes(trimmedArg)
                    cmdargs.Remove(arg)

                Case trimmedArg = "bykey" OrElse trimmedArg = "byname"

                    ' This is the default behavior so we don't have to invert anything,
                    ' We do have to remove the args though 
                    invertSettingAndRemoveArg(False, arg)

                Case Flags.ContainsKey(trimmedArg)

                    invertSettingAndRemoveArg(Flags(trimmedArg), arg)

                Case PresetFileNames.ContainsKey(trimmedArg)

                    invertSettingAndRemoveArg(False, arg, TransmuteFile2.Name, PresetFileNames(trimmedArg))

            End Select

        Next

        Transmutator = mode
        TransmuteReplaceMode = If(byKeyMode, ReplaceMode.ByKey, ReplaceMode.BySection)
        TransmuteRemoveMode = If(byKeyMode, RemoveMode.ByKey, RemoveMode.BySection)
        TransmuteRemoveKeyMode = If(RemoveByName, RemoveKeyMode.ByName, RemoveKeyMode.ByValue)
        UseWinapp2Syntax = isWinapp

        getFileAndDirParams({TransmuteFile1, TransmuteFile2, TransmuteFile3})
        If TransmuteFile2.Name.Length <> 0 Then initTransmute()

    End Sub

    ''' <summary> 
    ''' Validates the <c> iniFiles </c> and kicks off the merging process 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-07-15 | Code last updated: 2025-07-15
    Public Sub initTransmute()

        clrConsole()

        If Not (enforceFileHasContent(TransmuteFile1) AndAlso enforceFileHasContent(TransmuteFile2)) Then Return

        Dim menuOutput As New MenuSection
        Dim applyingChangesStr = $"Applying changes to {TransmuteFile1.Name}"
        menuOutput.AddBoxWithText(applyingChangesStr)
        gLog(applyingChangesStr)

        Dim remModeStr = $"{TransmuteRemoveMode}{If(TransmuteRemoveMode = RemoveMode.ByKey, $" - {TransmuteRemoveKeyMode}", "")}"
        Dim replaceModeStr = $"{TransmuteReplaceMode}"
        Dim xmuteModeStr As String = $"Transmutator: {Transmutator}{If(Transmutator = TransmuteMode.Add, If(Transmutator = TransmuteMode.Replace, replaceModeStr, remModeStr), "")}"

        Dim color = If(Transmutator = TransmuteMode.Add, ConsoleColor.Green, If(Transmutator = TransmuteMode.Remove, ConsoleColor.Red, ConsoleColor.Yellow))

        menuOutput.AddColoredLine(xmuteModeStr, color)
        gLog(xmuteModeStr)

        transmute(TransmuteFile1, TransmuteFile2, TransmuteFile3, menuOutput, UseWinapp2Syntax)

        menuOutput.AddLine("") _
                  .AddBottomBorder() _
                  .AddAnyKeyPrompt()

        menuOutput.Print()

        crk()

    End Sub

    ''' <summary> 
    ''' Conducts the transmutation on the base file using the source file and saves the 
    ''' output to the location provided within saveFile. <br /> Typically by default, the output is
    ''' saved to the base file, overwriting the original content. <br />
    ''' If working on a winapp2.ini file, the output can be formatted accordingly 
    ''' </summary>
    ''' 
    ''' <param name="baseFile">
    ''' The <c> iniFile </c> whose content will be modified by the transmutation process
    ''' </param>
    ''' 
    ''' <param name="sourceFile">
    ''' The <c> iniFile </c> providing the transmutation data 
    ''' </param>
    ''' 
    ''' <param name="saveFile">
    ''' Contains the path to which the transmuted output will be written
    ''' </param>
    ''' 
    ''' <param name="menuOutput">
    ''' A <c> MenuSection </c> containing the Transmute output to be displayed to the user
    ''' </param>
    ''' 
    ''' <param name="isWinapp2">
    ''' Indicates that the <c> saveFile </c> should be formatted as a winapp2.ini file
    ''' </param>
    Private Sub transmute(ByRef baseFile As iniFile,
                          ByRef sourceFile As iniFile,
                          ByRef saveFile As iniFile,
                          ByRef menuOutput As MenuSection,
                       Optional isWinapp2 As Boolean = True)

        resolveConflicts(baseFile, sourceFile, menuOutput)

        If isWinapp2 Then

            Dim tmp As New winapp2file(baseFile)
            tmp.sortInneriniFiles()
            baseFile.Sections = tmp.toIni.Sections
            saveFile.overwriteToFile(tmp.winapp2string)

        Else

            baseFile.sortSections(baseFile.namesToStrList)
            saveFile.overwriteToFile(baseFile.toString)

        End If

    End Sub

    ''' <summary> 
    ''' Facilitates transmuting an <c> iniFile </c> from outside the module's UI
    ''' </summary>
    ''' 
    ''' <param name="baseFile">
    ''' An <c> iniFile </c> whose content will be modified by the transmutation process
    ''' </param>
    ''' 
    ''' <param name="sourceFile">
    ''' An <c> iniFile </c> whose content will be used to modify to <c> <paramref name="baseFile"/> </c> 
    ''' </param>
    ''' 
    ''' <param name="outputFile">
    ''' An <c> iniFile </c> which will be written to disk with the result of the transmutation process
    ''' </param>
    ''' 
    ''' <param name="isWinapp"> 
    ''' Indicates that the <c> iniFile </c>s being worked with contain winapp2.ini syntax 
    ''' </param>
    ''' 
    ''' <param name="menuOutput">
    ''' The <c> MenuSection </c> containing output to be displayed to the user 
    ''' </param>
    ''' 
    ''' <param name="transmuteMode"> 
    ''' Sets the primary <c> Transmutator </c> <br />
    ''' Optional, Default: <c> TransmuteMode.Add </c>
    ''' </param>
    ''' 
    ''' <param name="replaceMode">
    ''' Sets the sub mode for the <c> Replace </c> Transmutator  <br />
    ''' Optional, Default: <c> ReplaceMode.ByKey </c>
    ''' </param>
    '''  
    ''' <param name="removeMode">
    ''' Sets the sub mode for the <c> Remove </c> Transmutator <br />
    ''' Optional, Default: <c> RemoveMode.ByKey </c>
    ''' </param>
    ''' 
    ''' <param name="removeKeyMode">
    ''' Sets the sub mode for key removal operations when <c> removeMode </c> is 
    ''' <c> RemoveMode.ByKey </c> <br />
    ''' Optional, Default: <c> RemoveKeyMode.ByName </c> 
    ''' </param>
    ''' 
    ''' Docs last updated: 2025-08-27 | Code last updated: 2025-08-27
    Public Sub RemoteTransmute(ByRef baseFile As iniFile,
                               ByRef sourceFile As iniFile,
                               ByRef outputFile As iniFile,
                                     isWinapp As Boolean,
                               ByRef menuOutput As MenuSection,
                            Optional transmuteMode As TransmuteMode = TransmuteMode.Add,
                            Optional replaceMode As ReplaceMode = ReplaceMode.ByKey,
                            Optional removeMode As RemoveMode = RemoveMode.ByKey,
                            Optional removeKeyMode As RemoveKeyMode = RemoveKeyMode.ByName)

        ' We're going to assume here that not every invocation of RemoteTransmute specifically
        ' through Flavorize() will necessarily include a defined source file. 
        If sourceFile Is Nothing Then gLog($"Source file not provided, skipping!") : Return

        If baseFile.Sections.Count = 0 Then baseFile.init()

        If sourceFile.Sections.Count = 0 Then sourceFile.init()
        gLog($"{sourceFile.Name} is empty!", sourceFile.Sections.Count = 0)

        Dim initTransmutator = Transmutator
        Dim initReplMode = TransmuteReplaceMode
        Dim initRemMode = TransmuteRemoveMode
        Dim initRemKeyMode = TransmuteRemoveKeyMode

        Transmutator = transmuteMode
        TransmuteReplaceMode = replaceMode
        TransmuteRemoveMode = removeMode
        TransmuteRemoveKeyMode = removeKeyMode

        transmute(baseFile, sourceFile, outputFile, menuOutput, isWinapp)

        Transmutator = initTransmutator
        TransmuteReplaceMode = initReplMode
        TransmuteRemoveMode = initRemMode
        TransmuteRemoveKeyMode = initRemKeyMode

    End Sub

    ''' <summary> 
    ''' Steps through the sections in the <c> <paramref name="sourceFile"/> </c> and applies the 
    ''' transmutation to each. Sections found in the <c> <paramref name="sourceFile"/> </c> not 
    ''' found in the <c> <paramref name="baseFile"/> </c> when the transmutator is not <c> Add </c>
    ''' will be ignored with an error message 
    ''' </summary>
    ''' 
    ''' <param name="baseFile"> 
    ''' The <c> iniFile </c> whose content is being modified by the transmutation process
    ''' </param>
    ''' 
    ''' <param name="sourceFile"> 
    '''  The <c> iniFile </c> providing the content modification criteria for <c> <paramref name="baseFile"/> </c>
    ''' </param>
    ''' 
    ''' <param name="menuOutput">
    ''' The <c> MenuSection </c> containing output to be displayed to the user 
    ''' </param>
    ''' 
    ''' Docs last updated: 2025-08-27 | Code last updated: 2025-08-27
    Private Sub resolveConflicts(ByRef baseFile As iniFile,
                                 ByRef sourceFile As iniFile,
                                 ByRef menuOutput As MenuSection)

        For Each sectionName In sourceFile.Sections.Keys

            Dim baseFileHasSection = baseFile.Sections.Keys.Contains(sectionName)

            ' If we're not adding and the base file doesn't have a section of the same name, there's no conflicts to resolve 
            If Not baseFileHasSection AndAlso Not Transmutator = TransmuteMode.Add Then

                Dim notFoundMsg = $"Target section not found in base file: [{sectionName}] - no changes applied"
                menuOutput.AddWarning(notFoundMsg)
                gLog(notFoundMsg)

                Continue For

            End If

            Dim baseSection = If(baseFileHasSection, baseFile.Sections.Item(sectionName), New iniSection)

            processTransmutator(baseFile, sourceFile, sectionName, baseSection, menuOutput)

        Next

    End Sub

    ''' <summary>
    ''' Hands off the processing of a section to the appropriate handler based on the 
    ''' current <c> Transmutator </c> setting. <br />
    ''' </summary>
    ''' 
    ''' <param name="baseFile">
    ''' The <c> iniFile </c> which will be modified by the transmutation process
    ''' </param>
    ''' 
    ''' <param name="sourceFile">
    ''' The <c> iniFile </c> whose content will be used to modify <c> <paramref name="baseFile"/> </c>
    ''' </param>
    ''' 
    ''' <param name="sectionName">
    ''' The name of the section currently being processed. This is used to identify the section
    ''' in both the <c> baseFile </c> and <c> sourceFile </c> <br />
    ''' 
    ''' The section name must match exactly (case sensitive) between the two files for the 
    ''' transmutation to occur
    ''' </param>
    ''' 
    ''' <param name="baseSection">
    ''' The <c> iniSection </c> which will be modified by the transmutation process
    ''' </param>
    ''' 
    ''' <param name="menuOutput">
    ''' The <c> MenuSection </c> containing output to be displayed to the user 
    ''' </param>
    ''' 
    ''' Docs last updated: 2025-08-27 | Code last updated: 2025-08-27
    Private Sub processTransmutator(ByRef baseFile As iniFile,
                                    ByRef sourceFile As iniFile,
                                          sectionName As String,
                                    ByRef baseSection As iniSection,
                                    ByRef menuOutput As MenuSection)

        Dim sourceSection = sourceFile.Sections.Item(sectionName)

        Select Case Transmutator

            Case TransmuteMode.Add

                handleAddMode(baseSection, sourceSection, baseFile, sectionName, menuOutput)

            Case TransmuteMode.Replace

                handleReplaceMode(baseSection, sourceSection, baseFile, sectionName, menuOutput)

            Case TransmuteMode.Remove

                handleRemoveMode(baseSection, sourceSection, baseFile, sectionName, menuOutput)

        End Select

    End Sub

    ''' <summary>
    ''' Handles the <c> Add Transmutator </c>, either adding the <c> <paramref name="sourceSection"/> </c>
    ''' to the <c> <paramref name="baseFile"/> </c> or adding the keys from the 
    ''' <c> <paramref name="sourceSection"/> </c> to the <c> <paramref name="baseSection"/> </c>
    ''' </summary>
    ''' 
    ''' <param name="baseSection">
    ''' The <c> iniSection </c> from the <c> baseFile </c>, if it exists, which will have keys
    ''' from the <c> <paramref name="sourceSection"/> </c> added to it 
    ''' </param>
    ''' 
    ''' <param name="sourceSection">
    ''' The <c> iniSection </c> to either be added to <c> <paramref name="baseFile"/> </c> or 
    ''' whose keys will be added to the <c> <paramref name="baseSection"/> </c> <br />
    ''' 
    ''' If the section exists in the base file, its keys will be added to the base section.
    ''' If it does not exist, the section will be added as a new section in the base file
    ''' </param>
    ''' 
    ''' <param name="baseFile">
    ''' The <c> iniFile </c> whose content will be modified by the transmutation process
    ''' </param>
    ''' 
    ''' <param name="sectionName">
    ''' The name of the current section undergoing the Add operation 
    ''' </param>
    ''' 
    ''' <param name="menuOutput">
    ''' The <c> MenuSection </c> containing output to be displayed to the user 
    ''' </param>
    ''' 
    ''' Docs last updated: 2025-08-27 | Code last updated: 2025-08-27
    Public Sub handleAddMode(ByRef baseSection As iniSection,
                             ByRef sourceSection As iniSection,
                             ByRef baseFile As iniFile,
                                   sectionName As String,
                             ByRef menuOutput As MenuSection)

        If baseFile.Sections.ContainsKey(sectionName) Then addKeysToBase(baseSection, sourceSection, menuOutput) : Return

        baseFile.Sections.Add(sectionName, sourceSection)

        Dim newSectionMsg = $"+§ Added new section: {sectionName}"
        menuOutput.AddColoredLine(newSectionMsg, ConsoleColor.Green)
        gLog(newSectionMsg)

    End Sub

    ''' <summary>
    ''' Adds keys from <c> <paramref name="sourceSection"/> </c> to 
    ''' <c> <paramref name="baseSection"/> </c>
    ''' </summary>
    ''' 
    ''' <param name="baseSection">
    ''' The <c> iniSection </c> to which keys will be added from 
    ''' <c> <paramref name="sourceSection"/> </c>
    ''' </param>
    ''' 
    ''' <param name="sourceSection">
    ''' The <c> iniSection </c> providing the keys to be added to  
    ''' <c> <paramref name="baseSection"/> </c>
    ''' </param>
    ''' 
    ''' <param name="menuOutput">
    ''' The <c> MenuSection </c> containing output to be displayed to the user 
    ''' </param>
    ''' 
    ''' Docs last updated: 2025-08-27 | Code last updated: 2025-08-27
    Private Sub addKeysToBase(ByRef baseSection As iniSection,
                              ByRef sourceSection As iniSection,
                              ByRef menuOutput As MenuSection)

        Dim addKeysMsg = $"Adding keys to {sourceSection.Name}"
        menuOutput.AddColoredLine(addKeysMsg, ConsoleColor.Cyan)
        gLog(addKeysMsg)

        For Each sourceKey In sourceSection.Keys.Keys

            baseSection.Keys.add(sourceKey)

            Dim addedKeyMsg = $"  += Added key: {sourceKey.Name}={sourceKey.Value}"
            menuOutput.AddColoredLine(addedKeyMsg, ConsoleColor.Green)
            gLog(addedKeyMsg)

        Next

    End Sub

    ''' <summary>
    ''' Handles the <c> Replace Transmutator </c>, replacing sections or keys based on the 
    ''' current replaceMode setting
    ''' </summary>
    ''' 
    ''' <param name="baseSection">
    ''' The <c> iniSection </c> which will have its content mutated based on 
    ''' <c> <paramref name="sourceSection"/> </c>
    ''' </param>
    ''' 
    ''' <param name="sourceSection">
    ''' The <c> iniSection </c> providing the replacement values for 
    ''' <c> <paramref name="baseSection"/> </c>
    ''' </param>
    ''' 
    ''' <param name="baseFile">
    ''' The <c> iniFile </c> containing <c> <paramref name="baseSection"/> </c>
    ''' </param>
    ''' 
    ''' <param name="sectionName">
    ''' The name on disk of both <c> <paramref name="baseSection"/> </c>
    ''' and also <c> <paramref name="sourceSection"/> </c>
    ''' </param>
    ''' 
    ''' <param name="menuOutput">
    ''' The <c> MenuSection </c> containing output to be displayed to the user 
    ''' </param>
    ''' 
    ''' Docs last updated: 2025-08-27 | Code last updated: 2025-08-27
    Private Sub handleReplaceMode(ByRef baseSection As iniSection,
                                  ByRef sourceSection As iniSection,
                                  ByRef baseFile As iniFile,
                                        sectionName As String,
                                  ByRef menuOutput As MenuSection)

        Dim ResolvingMsg = $"Resolving collisons in {sectionName} ({Transmutator} - {TransmuteReplaceMode} mode)"
        menuOutput.AddColoredLine(ResolvingMsg, ConsoleColor.Cyan)
        gLog(ResolvingMsg)

        ' if we're replacing by key, we can return immediately after
        If TransmuteReplaceMode = ReplaceMode.ByKey Then replaceKeysInBase(baseSection, sourceSection, menuOutput) : Return

        baseFile.Sections(sectionName) = sourceSection
        Dim replMsg = $"  *§ Replaced entire section: {sectionName}"
        menuOutput.AddColoredLine(replMsg, ConsoleColor.Yellow)

    End Sub

    ''' <summary>
    ''' Handles the Replace by Key mode for the <c> Replace Transmutator </c>, replacing the values
    ''' of keys in <c> <paramref name="baseSection"/> </c> with the ones provided in 
    ''' <c> <paramref name="sourceSection"/> </c> iff they have the same Name 
    ''' </summary>
    ''' 
    ''' <param name="baseSection">
    ''' The <c> iniSection </c> which will have its key values will be replaced with values
    ''' provided in <c> <paramref name="sourceSection"/> </c>
    ''' </param>
    ''' 
    ''' <param name="sourceSection">
    ''' The <c> iniSection </c> providing the replacement values for matching keys found within
    ''' <c> <paramref name="baseSection"/> </c>
    ''' </param>
    ''' 
    ''' <param name="menuOutput">
    ''' The <c> MenuSection </c> containing output to be displayed to the user 
    ''' </param>
    ''' 
    ''' Docs last updated: 2025-08-27 | Code last updated: 2025-08-27
    Private Sub replaceKeysInBase(ByRef baseSection As iniSection,
                                  ByRef sourceSection As iniSection,
                                  ByRef menuOutput As MenuSection)

        ' create lookup of source keys by name
        Dim sourceKeys As New Dictionary(Of String, iniKey)(StringComparer.InvariantCultureIgnoreCase)
        For Each sourceKey In sourceSection.Keys.Keys

            sourceKeys(sourceKey.Name) = sourceKey

        Next

        ' replace matches
        For i As Integer = 0 To baseSection.Keys.KeyCount - 1

            Dim baseKey = baseSection.Keys.Keys(i)

            If Not sourceKeys.ContainsKey(baseKey.Name) Then Continue For

            baseSection.Keys.Keys(i) = sourceKeys(baseKey.Name)

            Dim replKeyMsg = $"  * Replaced key: {baseKey.Name}"
            menuOutput.AddColoredLine(replKeyMsg, ConsoleColor.Yellow)
            gLog(replKeyMsg)

            sourceKeys.Remove(baseKey.Name)

        Next

        ' report any missed replacement targets
        For Each key In sourceKeys.Values

            Dim errMsg = $"Replacement target not found: {key.Name} not found in {baseSection.Name}"
            menuOutput.AddWarning(errMsg)
            gLog(errMsg)

        Next

    End Sub

    ''' <summary>
    ''' Handles the <c> Remove Transmutator </c>, removing sections or keys based on the 
    ''' current removeMode setting
    ''' </summary>
    ''' 
    ''' <param name="baseSection">
    ''' The <c> iniSection </c> which will be either removed or have keys removed from it
    ''' </param>
    ''' 
    ''' <param name="sourceSection">
    ''' The <c> iniSection </c> providing the removal parameters 
    ''' </param>
    ''' 
    ''' <param name="baseFile">
    ''' The <c> iniFile </c> which will be modified by the Transmutation process
    ''' </param>
    ''' 
    ''' <param name="sectionName">
    ''' The name on disk of both <c> <paramref name="baseSection"/> </c>
    ''' and also <c> <paramref name="sourceSection"/> </c>
    ''' </param>
    ''' 
    ''' <param name="menuOutput">
    ''' The <c> MenuSection </c> containing output to be displayed to the user 
    ''' </param>
    ''' 
    ''' Docs last updated: 2025-08-27 | Code last updated: 2025-08-27
    Private Sub handleRemoveMode(ByRef baseSection As iniSection,
                                 ByRef sourceSection As iniSection,
                                 ByRef baseFile As iniFile,
                                       sectionName As String,
                                 ByRef menuOutput As MenuSection)

        Dim isKeyMode = TransmuteRemoveMode = RemoveMode.ByKey

        Dim conflictsStr = $"Resolving conflicts in {sectionName} ({Transmutator} - {TransmuteRemoveMode} - {If(isKeyMode, TransmuteRemoveKeyMode.ToString, "")})"
        menuOutput.AddColoredLine(conflictsStr, ConsoleColor.Cyan)
        gLog(conflictsStr)

        ' if we're removing keys we can immediately return when we're done 
        If isKeyMode Then remKeys(baseSection, sourceSection, menuOutput) : Return

        baseFile.Sections.Remove(sectionName)
        Dim remMsg = $"  -§ Removed entire section: {sectionName}"
        menuOutput.AddColoredLine(remMsg, ConsoleColor.Red)
        gLog(remMsg)

    End Sub

    ''' <summary>
    ''' Removes individual keys from the <c> <paramref name="baseSection"/> </c>, obeying the 
    ''' current <c> TransmuteRemoveMode </c> and <c> TransmuteRemoveKeyMode </c> settings. <br />
    ''' </summary>
    ''' 
    ''' <param name="baseSection">
    ''' The <c> iniSection </c> from which keys will be removed 
    ''' </param>
    ''' 
    ''' <param name="sourceSection">
    ''' The <c> iniSection </c> providing the keys to be removed from 
    ''' <c> <paramref name="baseSection"/> </c> <br />
    ''' </param>
    ''' 
    ''' <param name="menuOutput">
    ''' The <c> MenuSection </c> containing output to be displayed to the user 
    ''' </param>
    ''' 
    ''' Docs last updated: 2025-08-27 | Code last updated: 2025-08-27
    Private Sub remKeys(ByRef baseSection As iniSection,
                              sourceSection As iniSection,
                        ByRef menuOutput As MenuSection)

        Dim sourceData As New HashSet(Of String)(StringComparer.InvariantCultureIgnoreCase)
        Dim isByName = TransmuteRemoveKeyMode = RemoveKeyMode.ByName

        ' build a lookup of source keys by either name or value 
        For Each sourceKey In sourceSection.Keys.Keys

            sourceData.Add(If(isByName, sourceKey.Name, $"{sourceKey.KeyType}={sourceKey.Value}"))

        Next

        ' remove the matches 
        For i As Integer = baseSection.Keys.KeyCount - 1 To 0 Step -1

            Dim baseKey = baseSection.Keys.Keys(i)
            Dim matchStr = If(isByName, baseKey.Name, $"{baseKey.KeyType}={baseKey.Value}")

            If Not sourceData.Contains(matchStr) Then Continue For

            baseSection.Keys.Keys.RemoveAt(i)
            Dim remKeyMsg = $"  -= Removed key by {If(isByName, "name", "value")}: {baseKey.toString}"
            menuOutput.AddColoredLine(remKeyMsg, ConsoleColor.Red)
            gLog(remKeyMsg)

            sourceData.Remove(matchStr)

        Next

        ' report any missed removal targets 
        For Each key In sourceData

            LogAndPrint(7, $"Removal target not found: {key} not found in {baseSection.Name}")

        Next

    End Sub

    ''' <summary>
    ''' Applies a 'flavor' to an <c> iniFile </c>. A flavor is like a compatibility layer that
    ''' allows a base file to be adjusted in slight ways to make it more suitable to a specific use
    ''' <br /><br/> 
    ''' In the context of winapp2.ini, we can use this to create different versions of winapp2.ini 
    ''' for different use cases (eg. a CCleaner or BleachBit specific versions) or else to 
    ''' correct the output of generative components of winapp2ool <br /><br />
    ''' 
    ''' Always applies Flavorings in the following order: <br /><br />
    ''' Section Removal -> Key Name Removal -> Key Value Removal -> Section Replacement ->
    ''' Key Replacement -> Section and Key Additions
    ''' </summary>
    ''' 
    ''' <param name="baseFile">
    ''' The <c> iniFile </c> to whom a particular flavor will be applied
    ''' </param>
    ''' 
    ''' <param name="outputFile">
    ''' The location on disk to which the flavorized <c> baseFile </c> will be saved 
    ''' </param>
    ''' 
    ''' <param name="menuOutput">
    ''' The <c> MenuSection </c> containing output to be displayed to the user 
    ''' </param>
    ''' 
    ''' <param name="additionsFile">
    ''' The <c> iniFile </c> containing the set of sections and individual keys within sections
    ''' which should be added to <c> <paramref name="baseFile"/> </c> to create the flavor <br />
    ''' Optional, Default: <c> Nothing </c>
    ''' </param>
    ''' 
    ''' <param name="sectionRemovalFile">
    ''' The <c> iniFile </c> containing the set of sections to be removed from the 
    ''' <c> <paramref name="baseFile"/> </c> to create the flavor <br /> 
    ''' Sections will be removed regardless of whether or not keys are provided <br />
    ''' Optional, Default: <c> Nothing </c>
    ''' </param>
    ''' 
    ''' <param name="keyNameRemovalFile">
    ''' The <c> iniFile </c> containing the set of individual keys to be removed 
    ''' from the <c> <paramref name="baseFile"/> </c> when matched by their Name parameter 
    ''' to create the flavor <br />
    ''' The values provided for keys in this file do not matter and will not be used for matching <br />
    ''' Optional, Default: <c> Nothing </c>
    ''' </param>
    ''' 
    ''' <param name="keyValueRemovalFile">
    ''' The <c> iniFile </c> containing the set of individual keys to be removed from the 
    ''' <c> <paramref name="baseFile"/> </c> when matched by their KeyName and Value pairs 
    ''' to create the flavor <br />
    ''' Numbers in key names will be ignored in this file and can be omitted. Numberless name
    ''' and value pairs will be used for matching. <br />
    ''' Optional, Default: <c> Nothing </c>
    ''' </param>
    ''' 
    ''' <param name="sectionReplacementFile">
    ''' The <c> iniFile </c> containing the set of sections to replace entire sections of an
    ''' exactly matching (case sensitive) name in the <c> <paramref name="baseFile"/> </c> 
    ''' to create the flavor <br />
    ''' This will not preserve non-overlapping content from the base key <br />
    ''' Optional, Default: <c> Nothing </c>
    ''' </param>
    ''' 
    ''' <param name="keyReplacementFile">
    ''' The <c> iniFile </c> containing the set of individual keys to replace within a matching
    ''' section within <c> <paramref name="baseFile"/> </c> to create the flavor <br />
    ''' Keys in this file will only replace keys in the base file if both their Seciton and Key 
    ''' names match exactly (case sensitive) <br />
    ''' Optional, Default: <c> Nothing </c>
    ''' 
    ''' </param>
    ''' 
    ''' <param name="isWinapp">
    ''' Indicates that the <c> iniFile </c> being flavorized has winapp2.ini syntax <br />
    ''' Optional, Default: <c> True </c>
    ''' 
    ''' </param>
    ''' 
    ''' Docs last updated: 2025-08-27 | Code last updated: 2025-08-27
    Public Sub Flavorize(ByRef baseFile As iniFile,
                         ByRef outputFile As iniFile,
                         ByRef menuOutput As MenuSection,
                      Optional additionsFile As iniFile = Nothing,
                      Optional sectionRemovalFile As iniFile = Nothing,
                      Optional keyNameRemovalFile As iniFile = Nothing,
                      Optional keyValueRemovalFile As iniFile = Nothing,
                      Optional sectionReplacementFile As iniFile = Nothing,
                      Optional keyReplacementFile As iniFile = Nothing,
                      Optional isWinapp As Boolean = True)

        Dim flavorizingMsg = $"Flavorizing {baseFile.Name}"
        menuOutput.AddColoredLine(flavorizingMsg, ConsoleColor.Magenta)
        gLog(flavorizingMsg)

        Dim flavorOperations As New Dictionary(Of String, Object()) From {
            {"Removing sections", {sectionRemovalFile, TransmuteMode.Remove, ReplaceMode.ByKey, RemoveMode.BySection, RemoveKeyMode.ByName}},
            {"Removing keys by name", {keyNameRemovalFile, TransmuteMode.Remove, ReplaceMode.ByKey, RemoveMode.ByKey, RemoveKeyMode.ByName}},
            {"Removing keys by value", {keyValueRemovalFile, TransmuteMode.Remove, ReplaceMode.ByKey, RemoveMode.ByKey, RemoveKeyMode.ByValue}},
            {"Replacing sections", {sectionReplacementFile, TransmuteMode.Replace, ReplaceMode.BySection, RemoveMode.ByKey, RemoveKeyMode.ByName}},
            {"Replacing keys by name", {keyReplacementFile, TransmuteMode.Replace, ReplaceMode.ByKey, RemoveMode.ByKey, RemoveKeyMode.ByName}},
            {"Adding keys and sections", {additionsFile, TransmuteMode.Add, ReplaceMode.ByKey, RemoveMode.ByKey, RemoveKeyMode.ByName}}
        }

        For Each operation In flavorOperations

            Dim description = operation.Key
            Dim config = operation.Value
            Dim flavorFile = DirectCast(config(0), iniFile)
            Dim curMode = DirectCast(config(1), TransmuteMode)
            Dim curReplMode = DirectCast(config(2), ReplaceMode)
            Dim curRemMode = DirectCast(config(3), RemoveMode)
            Dim curRemKMode = DirectCast(config(4), RemoveKeyMode)

            menuOutput.AddColoredLine(description, ConsoleColor.Cyan)
            gLog(description)

            RemoteTransmute(baseFile, flavorFile, outputFile, isWinapp, menuOutput, curMode, curReplMode, curRemMode, curRemKMode)

        Next

        Dim flavorizedMsg = $"{baseFile.Name} Flavorized"
        menuOutput.AddColoredLine(flavorizedMsg, ConsoleColor.Magenta)
        gLog(flavorizedMsg)

    End Sub

End Module
