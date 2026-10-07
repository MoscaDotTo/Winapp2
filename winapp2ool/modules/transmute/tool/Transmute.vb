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
''' Modifies a base <see cref="iniFile"/> using the content of a separate source
''' <see cref="iniFile"/>, with conflict resolution at different levels of granularity.
''' Transmute was formerly the Merge module. <br /><br />
'''
''' In the parlance of Transmute, there are two files of interest: <br />
''' The 'base' file and the 'source' file <br /> <br />
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
''' If a section from the source file exists in the base file, its keys are appended to the base
''' section, even when the base section already has a key with the same Name
''' </description>
''' </item>
'''
''' <item>
''' <b> Replace </b>
'''
''' <description>
''' Replace has two sub modes: BySection and ByKey <br /><br />
'''
''' ByKey (Default replace mode): Replaces the value of keys in the base file with values from the
''' source file based on their Name <br /><br />
'''
''' BySection: Replaces entire sections in the base file with the section of the same name
''' in the source file. <br /><br />
'''
''' Replace skips source sections the base file doesn't have, with a warning
''' </description>
''' </item>
'''
''' <item>
''' <b> Remove </b>
'''
''' <description>
''' Remove has two sub modes: BySection and ByKey <br /><br />
'''
''' ByKey (Default remove mode): Removes keys from individual sections within the base file <br />
''' Keys must be provided within a section in the source file and will be removed from the
''' corresponding section of the same name in the base file <br />
'''
''' Remove Key Modes: <br /><br />
''' <b> ByName </b>: Removes keys based on matching key names only <br />
''' Removing a key with this mode requires knowledge of its Name. This works best for unnumbered
''' keys (in the context of winapp2.ini, keys like Section, or unnumbered Detection keys)
''' but also works for numbered keys if you know the number. For numbered keys it is generally
''' best to try using the ByValue remove mode <br />
'''
''' <b> ByValue </b>: Removes keys based on matching key values. The KeyType (key name
''' without numbers) must also match. Works best for numbered keys whose numbers may change
''' or be unclear when invoking Transmute
''' </description>
''' </item>
''' </list>
'''
''' The three modes are wholly discrete: Add never replaces or removes anything, and Replace and
''' Remove never add a section the base file lacks. <br /><br />
'''
''' Source files may also hold global sections (<c> [*] </c>, <c> [*Map: label] </c> and
''' <c> [*Name: scaffold] </c>) that apply across the base file instead of to one section of
''' the same name. <see cref="RecognizeGlobalSections"/> turns them on. <br /><br />
'''
''' <see cref="Flavorize"/> applies a "flavor" to an ini file: a set of transformations that adapts
''' the base winapp2.ini to a different purpose, such as the format a particular cleaner expects.
''' It always applies its operations in the following order: <br />
''' <br />
''' Section Removal -> Key Name Removal -> Key Value Removal -> Section Replacement ->
''' Key Replacement -> Section and Key Additions
''' <br />
''' </summary>
Public Module Transmute

    ''' <summary>
    ''' The token in a <c> [*] </c> or <c> [*Name:] </c> key value which is replaced with each
    ''' receiving section's name as the key is applied (case-insensitive). Named-section
    ''' operations leave the token literal
    ''' </summary>
    Public Const EntryNameToken As String = "%EntryName%"

    ''' <summary>
    ''' Enum representing the different primary modes of modifying the base file
    ''' </summary>
    '''
    Public Enum TransmuteMode

        ''' <summary>
        ''' Add the content of the source file into the base file <br />
        ''' For sections in the source file not found in the base file, they will be added as is <br /> 
        ''' sections already existing in the base file will have the keys from the source file added
        ''' </summary>
        Add = 0

        ''' <summary>
        ''' Overwrite the content of individual sections or keys in the base file with 
        ''' content from the source file. <br />
        ''' Whether replacements are done by section or by key is controlled by a separate
        ''' enum <see cref="ReplaceMode"/>
        ''' </summary>
        Replace = 1

        ''' <summary>
        ''' Remove from the base file any keys or sections found in the source file. <br />
        ''' Whether removals are done by section or by key is controlled by a separate enum
        ''' called <see cref="RemoveMode"/> <br />
        ''' Whether key removals are done by Name or by Value is controlled by a separate enum
        ''' called <see cref="RemoveKeyMode"/>
        ''' </summary>
        Remove = 2

    End Enum

    ''' <summary>
    ''' Enum representing the granularity of replace operations
    ''' </summary>
    '''
    Public Enum ReplaceMode

        ''' <summary>
        ''' Replaces entire sections when collisions occur <br />
        '''
        ''' This means that if a section exists in the base file, it will be replaced entirely
        ''' with the section from the source file. Section names are matched case-insensitively
        ''' </summary>
        BySection = 0

        ''' <summary>
        ''' Replace individual keys when collisions occur <br />
        '''
        ''' This means that if a key exists in the base file, its value will be replaced with the
        ''' value from the source file. The key must exist in both files and have the same Name
        ''' (case-insensitive). Only the Value is replaced. The base key's Name is untouched
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
        ''' if it also exists in the source file. Section names are matched case-insensitively
        ''' </summary>
        BySection = 0

        ''' <summary>
        ''' Remove individual keys when collisions occur <br />
        '''
        ''' This means that if a key exists in the base file, it will be removed if it also exists
        ''' in the source file. <see cref="RemoveKeyMode"/> decides what counts as the same key:
        ''' the same Name, or the same KeyType (key name without numbers) and Value. Section names
        ''' and keys are matched case-insensitively
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
    ''' Handles the commandline args for Transmute. Resets the module settings first, so a
    ''' command line run never uses saved settings. Without a source file (<c> -2f </c> or a
    ''' preset flag) we log a message and return without transmuting.
    ''' </summary>
    '''
    ''' <remarks>
    ''' Transmute args: <br />
    ''' Primary modes (the last one given wins): <br />
    '''
    ''' -add            : Add mode (default) <br />
    ''' -replace        : Replace mode <br />
    ''' -remove         : Remove mode <br />
    '''
    ''' Sub-mode: shared by remove/replace since only one can be done at a time <br />
    '''
    ''' -bysection      : Replace/Remove by section <br />
    ''' -bykey          : Replace/Remove by key (default, accepted and ignored) <br />
    '''
    ''' Remove key criteria: <br />
    ''' -byname         : Remove keys by name (default, accepted and ignored) <br />
    ''' -byvalue        : Remove keys by value <br />
    '''
    ''' Winapp2.ini syntax correction <br />
    ''' -dontlint       : save the output as a plain ini file with alphabetized sections, instead
    ''' of sorting and formatting it as winapp2.ini
    '''
    ''' Global section handling <br />
    ''' -noglobal       : treat [*], [*Map:] and [*Name:] source sections as ordinary section names
    '''
    ''' Preset source file choices <br />
    '''
    ''' -r              : removed entries.ini  <br />
    ''' -c              : custom.ini  <br />
    ''' -w              : winapp3.ini <br />
    ''' -a              : archived entries.ini <br />
    ''' -b              : browsers.ini <br />
    ''' -u              : uwp.ini
    ''' </remarks>
    Public Sub handleCmdLine()

        initDefaultTransmuteSettings()

        ' Primary mode
        Dim mode As TransmuteMode = TransmuteMode.Add
        ' Sub mode
        Dim byKeyMode As Boolean = True
        ' Key removal mode
        Dim RemoveByName As Boolean = True
        ' Winapp2.ini style linting of output
        Dim isWinapp As Boolean = True
        ' Global mappings
        Dim recognizeGlobals As Boolean = True

        ' Mode selection must precede spec construction (not boolean toggles)
        For Each arg In cmdargs.ToList()

            Dim t = arg.ToLowerInvariant().TrimStart("-"c)
            Select Case t

                Case "add" : mode = TransmuteMode.Add : cmdargs.Remove(arg)
                Case "replace" : mode = TransmuteMode.Replace : cmdargs.Remove(arg)
                Case "remove" : mode = TransmuteMode.Remove : cmdargs.Remove(arg)
                Case "bykey", "byname" : cmdargs.Remove(arg)

            End Select

        Next

        Dim spec As New CliArgSpec("transmute")
        spec.WithFile(1, TransmuteFile1) _
            .WithFile(2, TransmuteFile2) _
            .WithFile(3, TransmuteFile3) _
            .WithFlag("-bysection", Sub() byKeyMode = Not byKeyMode) _
            .WithFlag("-byvalue", Sub() RemoveByName = Not RemoveByName) _
            .WithFlag("-dontlint", Sub() isWinapp = Not isWinapp) _
            .WithFlag("-noglobal", Sub() recognizeGlobals = Not recognizeGlobals) _
            .WithFileAlias("-r", TransmuteFile2, "Removed Entries.ini") _
            .WithFileAlias("-c", TransmuteFile2, "custom.ini") _
            .WithFileAlias("-w", TransmuteFile2, "winapp3.ini") _
            .WithFileAlias("-a", TransmuteFile2, "Archived Entries.ini") _
            .WithFileAlias("-b", TransmuteFile2, "browsers.ini") _
            .WithFileAlias("-u", TransmuteFile2, "uwp.ini") _
        .Parse()

        Transmutator = mode
        TransmuteReplaceMode = If(byKeyMode, ReplaceMode.ByKey, ReplaceMode.BySection)
        TransmuteRemoveMode = If(byKeyMode, RemoveMode.ByKey, RemoveMode.BySection)
        TransmuteRemoveKeyMode = If(RemoveByName, RemoveKeyMode.ByName, RemoveKeyMode.ByValue)
        UseWinapp2Syntax = isWinapp
        RecognizeGlobalSections = recognizeGlobals

        If TransmuteFile2.Name.Length = 0 Then

            Dim noSourceMsg = "Transmute requires a source file. Provide one with -2f <name> or a preset flag (-r, -c, -w, -a, -b, -u)"
            gLog(noSourceMsg)
            cwl(noSourceMsg)
            Return

        End If

        initTransmute()

    End Sub

    ''' <summary>
    ''' Loads the base and source files, transmutes the base file with the current settings,
    ''' writes the result to <see cref="TransmuteFile3"/> and displays the output. Returns
    ''' without transmuting when either file is missing or empty.
    ''' </summary>
    Public Sub initTransmute()

        clrConsole()

        Dim baseFile = TransmuteFile1.Load(TransmuteModuleSettingsChanged, NameOf(Transmute), NameOf(TransmuteFile1), NameOf(TransmuteModuleSettingsChanged))
        Dim sourceFile = TransmuteFile2.Load(TransmuteModuleSettingsChanged, NameOf(Transmute), NameOf(TransmuteFile2), NameOf(TransmuteModuleSettingsChanged))

        If Not (enforceFileHasContent(baseFile) AndAlso enforceFileHasContent(sourceFile)) Then Return

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

        Dim saveFile = iniFile.Empty(TransmuteFile3.Dir, TransmuteFile3.Name)

        If Not transmute(baseFile, sourceFile, saveFile, menuOutput, UseWinapp2Syntax) Then menuOutput.AddWarning($"{saveFile.Name} was not saved")

        menuOutput.AddLine("") _
                  .AddBottomBorder() _
                  .AddAnyKeyPrompt()

        menuOutput.Print()

        crk()

    End Sub

    ''' <summary>
    ''' Transmutes <paramref name="baseFile"/> in memory using <paramref name="sourceFile"/> and
    ''' writes the result to the path held by <paramref name="saveFile"/>. The base file on disk
    ''' is only overwritten when <paramref name="saveFile"/> points at it.
    ''' </summary>
    '''
    ''' <param name="baseFile">
    ''' The <c> iniFile </c> whose content will be modified by the transmutation process. When
    ''' we save as winapp2.ini, we replace it with the sorted, reformatted result.
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
    ''' Indicates whether to sort and format the output as a winapp2.ini file. When
    ''' <c> False </c>, we write a plain ini file with its sections in alphabetical order. <br /><br />
    ''' Optional, Default: <c> True </c>
    ''' </param>
    '''
    ''' <param name="skipFormat">
    ''' Indicates whether to stop after modifying <paramref name="baseFile"/>, without
    ''' formatting or writing anything <br /><br />
    ''' Optional, Default: <c> False </c>
    ''' </param>
    '''
    ''' <returns>
    ''' <c> True </c> if the output was saved, <br />
    ''' <c> False </c> if saving was skipped or failed
    ''' </returns>
    Private Function transmute(ByRef baseFile As iniFile,
                          ByRef sourceFile As iniFile,
                                saveFile As iniFile,
                          ByRef menuOutput As MenuSection,
                       Optional isWinapp2 As Boolean = True,
                       Optional skipFormat As Boolean = False) As Boolean

        resolveConflicts(baseFile, sourceFile, menuOutput)

        If skipFormat Then Return False

        If isWinapp2 Then

            Dim wf As New winapp2file(baseFile)
            wf.SortEntries()
            Dim saved = saveFile.OverwriteToFile(wf.ToWinapp2String())
            baseFile = wf.ToIni()
            Return saved

        End If

        Return saveFile.OverwriteToFile(baseFile.ToString(IniFileWriteFormat.Alphabetical))

    End Function

    ''' <summary>
    ''' Transmutes an <c> iniFile </c> from outside the module's UI. We set the module's mode
    ''' settings to the given modes for the duration of the call and restore them afterward.
    ''' <see cref="RecognizeGlobalSections"/> isn't a parameter, so its current value applies.
    ''' </summary>
    '''
    ''' <param name="baseFile">
    ''' An <c> iniFile </c> whose content will be modified by the transmutation process. When
    ''' the output is saved as winapp2.ini, we replace it with the sorted, reformatted result.
    ''' </param>
    '''
    ''' <param name="sourceFile">
    ''' An <c> iniFile </c> whose content will be used to modify <paramref name="baseFile"/>
    ''' </param>
    '''
    ''' <param name="outputFile">
    ''' An <c> iniFile </c> which will be written to disk with the result of the transmutation process
    ''' </param>
    '''
    ''' <param name="isWinapp">
    ''' Indicates whether to sort and format the output as a winapp2.ini file
    ''' </param>
    ''' 
    ''' <param name="menuOutput">
    ''' The <c> MenuSection </c> containing output to be displayed to the user 
    ''' </param>
    ''' 
    ''' <param name="transmuteMode">
    ''' Sets the primary <c> Transmutator </c> <br /><br />
    ''' Optional, Default: <c> TransmuteMode.Add </c>
    ''' </param>
    '''
    ''' <param name="replaceMode">
    ''' Sets the sub mode for the <c> Replace </c> Transmutator <br /><br />
    ''' Optional, Default: <c> ReplaceMode.ByKey </c>
    ''' </param>
    '''
    ''' <param name="removeMode">
    ''' Sets the sub mode for the <c> Remove </c> Transmutator <br /><br />
    ''' Optional, Default: <c> RemoveMode.ByKey </c>
    ''' </param>
    '''
    ''' <param name="removeKeyMode">
    ''' Sets the sub mode for key removal operations when <paramref name="removeMode"/> is
    ''' <c> RemoveMode.ByKey </c> <br /><br />
    ''' Optional, Default: <c> RemoveKeyMode.ByName </c>
    ''' </param>
    '''
    ''' <param name="skipFormat">
    ''' Indicates whether to skip the post-transmutation sort, format, and disk write. <br />
    ''' Use this when chaining multiple transmutations against the same base file so that
    ''' only the final step pays the cost of constructing a <c> winapp2file </c> and writing to disk. <br /><br />
    ''' Optional, Default: <c> False </c>
    ''' </param>
    '''
    ''' <returns>
    ''' <c> True </c> if the output was saved, <br />
    ''' <c> False </c> if there was nothing to transmute, <paramref name="skipFormat"/> skipped the write, or the save failed
    ''' </returns>
    Public Function RemoteTransmute(ByRef baseFile As iniFile,
                               ByRef sourceFile As iniFile,
                                     outputFile As iniFile,
                                     isWinapp As Boolean,
                               ByRef menuOutput As MenuSection,
                            Optional transmuteMode As TransmuteMode = TransmuteMode.Add,
                            Optional replaceMode As ReplaceMode = ReplaceMode.ByKey,
                            Optional removeMode As RemoveMode = RemoveMode.ByKey,
                            Optional removeKeyMode As RemoveKeyMode = RemoveKeyMode.ByName,
                            Optional skipFormat As Boolean = False) As Boolean

        If sourceFile Is Nothing Then gLog("Source file not provided, skipping!") : Return False

        If baseFile Is Nothing OrElse baseFile.Count = 0 Then
            gLog("Base file is empty or not provided, skipping!")
            Return False
        End If

        If sourceFile.Count = 0 Then gLog($"{sourceFile.Name} is empty!") : Return False

        Dim initTransmutator = Transmutator
        Dim initReplMode = TransmuteReplaceMode
        Dim initRemMode = TransmuteRemoveMode
        Dim initRemKeyMode = TransmuteRemoveKeyMode

        Transmutator = transmuteMode
        TransmuteReplaceMode = replaceMode
        TransmuteRemoveMode = removeMode
        TransmuteRemoveKeyMode = removeKeyMode

        Dim saved = transmute(baseFile, sourceFile, outputFile, menuOutput, isWinapp, skipFormat)

        Transmutator = initTransmutator
        TransmuteReplaceMode = initReplMode
        TransmuteRemoveMode = initRemMode
        TransmuteRemoveKeyMode = initRemKeyMode

        Return saved

    End Function

    ''' <summary>
    ''' Steps through the sections in the <paramref name="sourceFile"/> and applies the
    ''' transmutation to each. Sections found in the <paramref name="sourceFile"/> but not
    ''' in the <paramref name="baseFile"/> when the transmutator is not <c> Add </c>
    ''' are skipped with a warning <br /> <br />
    '''
    ''' When <see cref="RecognizeGlobalSections"/> is enabled, three kinds of sentinel section are
    ''' processed before any named sections, in this order, so that specific per-section
    ''' operations can refine the result of global ones: <br />
    ''' <c> [*Map: label] </c> sections define key mapping rules (Replace ByKey mode only) <br />
    ''' A <c> [*] </c> section applies its keys to every section of the base file <br />
    ''' <c> [*Name: scaffold] </c> sections apply their payload keys to every base section whose
    ''' name ends with <c> " scaffold *" </c> and which satisfies the section's <c> Match= </c>
    ''' predicates (see <see cref="applyNameFilterSections"/>)
    ''' </summary>
    '''
    ''' <param name="baseFile">
    ''' The <c> iniFile </c> whose content is being modified by the transmutation process
    ''' </param>
    '''
    ''' <param name="sourceFile">
    '''  The <c> iniFile </c> providing the content modification criteria for <paramref name="baseFile"/>
    ''' </param>
    '''
    ''' <param name="menuOutput">
    ''' The <c> MenuSection </c> containing output to be displayed to the user
    ''' </param>
    '''
    Private Sub resolveConflicts(ByRef baseFile As iniFile,
                                       sourceFile As iniFile,
                                 ByRef menuOutput As MenuSection)

        Dim globalSection As iniSection = Nothing
        Dim mapSections As New List(Of iniSection)
        Dim nameSections As New List(Of iniSection)
        Dim namedSections As New List(Of iniSection)

        For Each sourceSection In sourceFile

            If RecognizeGlobalSections AndAlso sourceSection.Name = "*" Then

                globalSection = sourceSection

            ElseIf RecognizeGlobalSections AndAlso sourceSection.Name.StartsWith(MapSectionPrefix, StringComparison.OrdinalIgnoreCase) Then

                mapSections.Add(sourceSection)

            ElseIf RecognizeGlobalSections AndAlso sourceSection.Name.StartsWith(NameSectionPrefix, StringComparison.OrdinalIgnoreCase) Then

                nameSections.Add(sourceSection)

            Else

                namedSections.Add(sourceSection)

            End If

        Next

        applyKeyMapRules(baseFile, mapSections, menuOutput)

        If globalSection IsNot Nothing Then applyGlobalSection(baseFile, globalSection, menuOutput)

        transmutenamefilter.applyNameFilterSections(baseFile, nameSections, menuOutput)

        For Each sourceSection In namedSections

            Dim baseFileHasSection = baseFile.Contains(sourceSection.Name)

            If Not baseFileHasSection AndAlso Not Transmutator = TransmuteMode.Add Then

                Dim notFoundMsg = $"Target section not found in base file: [{sourceSection.Name}] - no changes applied"
                menuOutput.AddWarning(notFoundMsg)
                gLog(notFoundMsg)

                Continue For

            End If

            Dim baseSection = baseFile.GetSection(sourceSection.Name)

            processTransmutator(baseFile, sourceSection, baseSection, menuOutput)

        Next

    End Sub

    ''' <summary>
    ''' Applies the keys of a <c> [*] </c> sentinel section to every section in the
    ''' <paramref name="baseFile"/> under the current transmute mode. <br /> <br />
    '''
    ''' Section-level modes are refused with a warning: Remove BySection would empty the file
    ''' and Replace BySection is incoherent as a global operation. Numbered keys are refused
    ''' in Add mode because adding them to every section creates instant duplicates. An
    ''' unnumbered Add key skips any section that already has a key of that Name, as
    ''' <c> [*Name:] </c> does. <br /> <br />
    '''
    ''' Key values containing the <c> %EntryName% </c> token (case-insensitive) have the token
    ''' replaced with each receiving section's name as the key is applied, enabling per-section
    ''' values from a single global key (eg. <c> ID=%EntryName% </c>). Named-section operations
    ''' leave the token literal. <br /> <br />
    '''
    ''' Per-section misses are silent: the menu receives one summary line per source key
    ''' and per-hit detail is written to the log
    ''' </summary>
    '''
    ''' <param name="baseFile">
    ''' The <c> iniFile </c> whose sections will all receive the global operation
    ''' </param>
    '''
    ''' <param name="sourceSection">
    ''' The <c> [*] </c> section providing the keys to apply globally
    ''' </param>
    '''
    ''' <param name="menuOutput">
    ''' The <c> MenuSection </c> containing output to be displayed to the user
    ''' </param>
    '''
    Private Sub applyGlobalSection(ByRef baseFile As iniFile,
                                         sourceSection As iniSection,
                                   ByRef menuOutput As MenuSection)

        If Transmutator = TransmuteMode.Remove AndAlso TransmuteRemoveMode = RemoveMode.BySection Then

            Dim refuseRemMsg = "[*] cannot be used with Remove BySection (this would remove every section) - skipping"
            menuOutput.AddWarning(refuseRemMsg)
            gLog(refuseRemMsg)
            Return

        End If

        If Transmutator = TransmuteMode.Replace AndAlso TransmuteReplaceMode = ReplaceMode.BySection Then

            Dim refuseReplMsg = "[*] cannot be used with Replace BySection (a global section replacement is incoherent) - skipping"
            menuOutput.AddWarning(refuseReplMsg)
            gLog(refuseReplMsg)
            Return

        End If

        Dim totalSections = baseFile.Count

        Using gLogScope($"Applying global section [*] ({Transmutator}) to {totalSections} sections")

            For Each sourceKey In sourceSection.Keys

                If Transmutator = TransmuteMode.Add AndAlso Not sourceKey.Name.Equals(sourceKey.KeyType, StringComparison.OrdinalIgnoreCase) Then

                    Dim refuseKeyMsg = $"[*] Refusing to add numbered key {sourceKey.Name} to every section - global adds must be unnumbered"
                    menuOutput.AddWarning(refuseKeyMsg)
                    gLog(refuseKeyMsg)
                    Continue For

                End If

                Dim hasEntryNameToken = sourceKey.Value.IndexOf(EntryNameToken, StringComparison.OrdinalIgnoreCase) >= 0

                Dim singleKeySource As New iniSection(sourceSection.Name)
                If Not hasEntryNameToken Then singleKeySource.AddKey(sourceKey)

                Dim sectionsHit = 0

                For Each baseSection In baseFile

                    Dim appliedKey = sourceKey
                    Dim appliedSource = singleKeySource

                    If hasEntryNameToken Then

                        appliedKey = New iniKey($"{sourceKey.Name}={expandEntryNameToken(sourceKey.Value, baseSection.Name)}")
                        appliedSource = New iniSection(sourceSection.Name)
                        appliedSource.AddKey(appliedKey)

                    End If

                    Dim hits = 0

                    Select Case Transmutator

                        Case TransmuteMode.Add

                            If baseSection.HasKey(appliedKey.Name) Then Continue For
                            hits = addKeysToBase(baseSection, appliedSource, menuOutput, quiet:=True)

                        Case TransmuteMode.Replace

                            hits = replaceKeysInBase(baseSection, appliedSource, menuOutput, quiet:=True)

                        Case TransmuteMode.Remove

                            hits = remKeys(baseSection, appliedSource, menuOutput, quiet:=True)

                    End Select

                    If hits = 0 Then Continue For

                    sectionsHit += 1
                    gLog($"{baseSection.Name}: {appliedKey}{If(hits > 1, $" (x{hits})", "")}")

                Next

                Dim summary As String
                Dim summaryColor As ConsoleColor

                Select Case Transmutator

                    Case TransmuteMode.Add

                        summary = $"+[*] Added {sourceKey} to {sectionsHit} of {totalSections} sections"
                        summaryColor = ConsoleColor.Green

                    Case TransmuteMode.Replace

                        summary = $"*[*] Replaced {sourceKey.Name} in {sectionsHit} of {totalSections} sections"
                        summaryColor = ConsoleColor.Yellow

                    Case Else

                        Dim keyDesc = If(TransmuteRemoveKeyMode = RemoveKeyMode.ByName, sourceKey.Name, $"{sourceKey.KeyType}={sourceKey.Value}")
                        summary = $"-[*] Removed {keyDesc} from {sectionsHit} of {totalSections} sections"
                        summaryColor = ConsoleColor.Red

                End Select

                menuOutput.AddColoredLine(summary, summaryColor)
                gLog(summary)

            Next

        End Using

    End Sub

    ''' <summary>
    ''' Replaces every occurrence of <see cref="EntryNameToken"/> in
    ''' <paramref name="value"/> with <paramref name="sectionName"/>,
    ''' matching the token case-insensitively
    ''' </summary>
    '''
    ''' <param name="value">
    ''' The global section key value containing the token
    ''' </param>
    '''
    ''' <param name="sectionName">
    ''' The name of the section receiving the key
    ''' </param>
    '''
    ''' <returns>
    ''' <paramref name="value"/> with all token occurrences replaced
    ''' </returns>
    Friend Function expandEntryNameToken(value As String, sectionName As String) As String

        Dim result = ""
        Dim pos = 0

        Do
            Dim tokenPos = value.IndexOf(EntryNameToken, pos, StringComparison.OrdinalIgnoreCase)
            If tokenPos = -1 Then Exit Do

            result &= value.Substring(pos, tokenPos - pos) & sectionName
            pos = tokenPos + EntryNameToken.Length
        Loop

        Return result & value.Substring(pos)

    End Function

    ''' <summary>
    ''' Hands off the processing of a section to the appropriate handler based on the
    ''' current <c> Transmutator </c> setting. <br />
    ''' </summary>
    ''' 
    ''' <param name="baseFile">
    ''' The <c> iniFile </c> which will be modified by the transmutation process
    ''' </param>
    ''' 
    ''' <param name="sourceSection">
    ''' The <c> iniSection </c> from the source file whose content will be used to modify
    ''' <paramref name="baseFile"/>
    ''' </param>
    '''
    ''' <param name="baseSection">
    ''' The <c> iniSection </c> from <paramref name="baseFile"/> which will be modified,
    ''' or <c> Nothing </c> if the section does not exist in the base file (possible only in Add mode)
    ''' </param>
    ''' 
    ''' <param name="menuOutput">
    ''' The <c> MenuSection </c> containing output to be displayed to the user 
    ''' </param>
    '''
    Private Sub processTransmutator(ByRef baseFile As iniFile,
                                          sourceSection As iniSection,
                                          baseSection As iniSection,
                                    ByRef menuOutput As MenuSection)

        Select Case Transmutator

            Case TransmuteMode.Add

                handleAddMode(baseSection, sourceSection, baseFile, sourceSection.Name, menuOutput)

            Case TransmuteMode.Replace

                handleReplaceMode(baseSection, sourceSection, baseFile, sourceSection.Name, menuOutput)

            Case TransmuteMode.Remove

                handleRemoveMode(baseSection, sourceSection, baseFile, sourceSection.Name, menuOutput)

        End Select

    End Sub

    ''' <summary>
    ''' Handles the <c> Add Transmutator </c>, either adding the <paramref name="sourceSection"/>
    ''' to the <paramref name="baseFile"/> or adding the keys from the
    ''' <paramref name="sourceSection"/> to the <paramref name="baseSection"/>. A new section
    ''' goes in as the source's own object, not a copy.
    ''' </summary>
    ''' 
    ''' <param name="baseSection">
    ''' The <c> iniSection </c> from the <c> baseFile </c>, if it exists, which will have keys
    ''' from the <paramref name="sourceSection"/> added to it 
    ''' </param>
    ''' 
    ''' <param name="sourceSection">
    ''' The <c> iniSection </c> to either be added to <paramref name="baseFile"/> or 
    ''' whose keys will be added to the <paramref name="baseSection"/> <br />
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
    Private Sub handleAddMode(baseSection As iniSection,
                                    sourceSection As iniSection,
                              ByRef baseFile As iniFile,
                                    sectionName As String,
                              ByRef menuOutput As MenuSection)

        If baseFile.Contains(sectionName) Then addKeysToBase(baseSection, sourceSection, menuOutput) : Return

        baseFile.AddSection(sourceSection)

        Dim newSectionMsg = $"+§ Added new section: {sectionName}"
        menuOutput.AddColoredLine(newSectionMsg, ConsoleColor.Green)
        gLog(newSectionMsg)

    End Sub

    ''' <summary>
    ''' Appends a copy of every key in <paramref name="sourceSection"/> to
    ''' <paramref name="baseSection"/>, including keys whose Name the base section already has
    ''' </summary>
    '''
    ''' <param name="baseSection">
    ''' The <c> iniSection </c> to which keys will be added from
    ''' <paramref name="sourceSection"/>
    ''' </param>
    '''
    ''' <param name="sourceSection">
    ''' The <c> iniSection </c> providing the keys to be added to
    ''' <paramref name="baseSection"/>
    ''' </param>
    '''
    ''' <param name="menuOutput">
    ''' The <c> MenuSection </c> containing output to be displayed to the user
    ''' </param>
    '''
    ''' <param name="quiet">
    ''' Indicates whether to suppress all menu and log output, leaving reporting to the caller.
    ''' Global operations set it, since they aggregate per-section results into summary lines <br /><br />
    ''' Optional, Default: <c> False </c>
    ''' </param>
    '''
    ''' <returns>
    ''' The number of keys added to <paramref name="baseSection"/>, which is always the number
    ''' of keys in <paramref name="sourceSection"/>
    ''' </returns>
    Private Function addKeysToBase(baseSection As iniSection,
                                         sourceSection As iniSection,
                                   ByRef menuOutput As MenuSection,
                                Optional quiet As Boolean = False) As Integer

        If Not quiet Then

            Dim addKeysMsg = $"Adding keys to {baseSection.Name}"
            menuOutput.AddColoredLine(addKeysMsg, ConsoleColor.Cyan)
            gLog(addKeysMsg)

        End If

        Dim hits = 0

        For Each sourceKey In sourceSection.Keys

            baseSection.AddKey(New iniKey(sourceKey.ToString()))
            hits += 1

            If quiet Then Continue For

            Dim addedKeyMsg = $"  += Added key: {sourceKey.Name}={sourceKey.Value}"
            menuOutput.AddColoredLine(addedKeyMsg, ConsoleColor.Green)
            gLog(addedKeyMsg)

        Next

        Return hits

    End Function

    ''' <summary>
    ''' Handles the <c> Replace Transmutator </c>, replacing sections or keys based on
    ''' <see cref="TransmuteReplaceMode"/>. A replaced section is removed and the source's own
    ''' section object is added in its place, at the end of the file's section order.
    ''' </summary>
    ''' 
    ''' <param name="baseSection">
    ''' The <c> iniSection </c> which will have its content mutated based on 
    ''' <paramref name="sourceSection"/>
    ''' </param>
    ''' 
    ''' <param name="sourceSection">
    ''' The <c> iniSection </c> providing the replacement values for 
    ''' <paramref name="baseSection"/>
    ''' </param>
    ''' 
    ''' <param name="baseFile">
    ''' The <c> iniFile </c> containing <paramref name="baseSection"/>
    ''' </param>
    ''' 
    ''' <param name="sectionName">
    ''' The name on disk of both <paramref name="baseSection"/>
    ''' and also <paramref name="sourceSection"/>
    ''' </param>
    ''' 
    ''' <param name="menuOutput">
    ''' The <c> MenuSection </c> containing output to be displayed to the user 
    ''' </param>
    '''
    Private Sub handleReplaceMode(baseSection As iniSection,
                                        sourceSection As iniSection,
                                  ByRef baseFile As iniFile,
                                        sectionName As String,
                                  ByRef menuOutput As MenuSection)

        Dim ResolvingMsg = $"Resolving collisions in {sectionName} ({Transmutator} - {TransmuteReplaceMode} mode)"
        menuOutput.AddColoredLine(ResolvingMsg, ConsoleColor.Cyan)
        gLog(ResolvingMsg)

        If TransmuteReplaceMode = ReplaceMode.ByKey Then replaceKeysInBase(baseSection, sourceSection, menuOutput) : Return

        baseFile.RemoveSection(sectionName)
        baseFile.AddSection(sourceSection)
        Dim replMsg = $"  *§ Replaced entire section: {sectionName}"
        menuOutput.AddColoredLine(replMsg, ConsoleColor.Yellow)
        gLog(replMsg)

    End Sub

    ''' <summary>
    ''' Handles the Replace by Key mode for the <c> Replace Transmutator </c>, replacing the values
    ''' of keys in <paramref name="baseSection"/> with the ones provided in
    ''' <paramref name="sourceSection"/> iff they have the same Name (case-insensitive) <br />
    ''' Every base key sharing a source key's Name receives the replacement value, so duplicate
    ''' key names in the base section are all updated. When the source repeats a Name, its
    ''' last value wins.
    ''' </summary>
    '''
    ''' <param name="baseSection">
    ''' The <c> iniSection </c> whose key values will be replaced with values
    ''' provided in <paramref name="sourceSection"/>
    ''' </param>
    ''' 
    ''' <param name="sourceSection">
    ''' The <c> iniSection </c> providing the replacement values for matching keys found within
    ''' <paramref name="baseSection"/>
    ''' </param>
    ''' 
    ''' <param name="menuOutput">
    ''' The <c> MenuSection </c> containing output to be displayed to the user
    ''' </param>
    '''
    ''' <param name="quiet">
    ''' Indicates whether to suppress all menu and log output, including the
    ''' "replacement target not found" warnings, leaving reporting to the caller.
    ''' Global operations set it, since they expect most sections not to match <br /><br />
    ''' Optional, Default: <c> False </c>
    ''' </param>
    '''
    ''' <returns>
    ''' The number of keys in <paramref name="baseSection"/> whose values were replaced
    ''' </returns>
    Friend Function replaceKeysInBase(baseSection As iniSection,
                                       sourceSection As iniSection,
                                 ByRef menuOutput As MenuSection,
                              Optional quiet As Boolean = False) As Integer

        Dim sourceKeys As New Dictionary(Of String, iniKey)(StringComparer.OrdinalIgnoreCase)

        For Each sourceKey In sourceSection.Keys : sourceKeys(sourceKey.Name) = sourceKey : Next

        Dim matched As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Dim hits = 0

        For Each baseKey In baseSection.Keys

            If Not sourceKeys.ContainsKey(baseKey.Name) Then Continue For

            baseKey.Value = sourceKeys(baseKey.Name).Value
            hits += 1
            matched.Add(baseKey.Name)

            If quiet Then Continue For

            Dim replKeyMsg = $"  * Replaced key: {baseKey.Name}"
            menuOutput.AddColoredLine(replKeyMsg, ConsoleColor.Yellow)
            gLog(replKeyMsg)

        Next

        If quiet Then Return hits

        For Each key In sourceKeys.Values

            If matched.Contains(key.Name) Then Continue For

            Dim errMsg = $"Replacement target not found: {key.Name} not found in {baseSection.Name}"
            menuOutput.AddWarning(errMsg)
            gLog(errMsg)

        Next

        Return hits

    End Function

    ''' <summary>
    ''' Handles the <c> Remove Transmutator </c>, removing sections or keys based on
    ''' <see cref="TransmuteRemoveMode"/>
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
    ''' The name on disk of both <paramref name="baseSection"/>
    ''' and also <paramref name="sourceSection"/>
    ''' </param>
    ''' 
    ''' <param name="menuOutput">
    ''' The <c> MenuSection </c> containing output to be displayed to the user 
    ''' </param>
    '''
    Private Sub handleRemoveMode(baseSection As iniSection,
                                 sourceSection As iniSection,
                           ByRef baseFile As iniFile,
                                 sectionName As String,
                           ByRef menuOutput As MenuSection)

        Dim isKeyMode = TransmuteRemoveMode = RemoveMode.ByKey

        Dim conflictsStr = $"Resolving conflicts in {sectionName} ({Transmutator} - {TransmuteRemoveMode} - {If(isKeyMode, TransmuteRemoveKeyMode.ToString(), "")})"
        menuOutput.AddColoredLine(conflictsStr, ConsoleColor.Cyan)
        gLog(conflictsStr)

        If isKeyMode Then remKeys(baseSection, sourceSection, menuOutput) : Return

        baseFile.RemoveSection(sectionName)
        Dim remMsg = $"  -§ Removed entire section: {sectionName}"
        menuOutput.AddColoredLine(remMsg, ConsoleColor.Red)
        gLog(remMsg)

    End Sub

    ''' <summary>
    ''' Removes individual keys from the <paramref name="baseSection"/>, matching by Name or by
    ''' KeyType and Value according to <see cref="TransmuteRemoveKeyMode"/>, case-insensitively. <br />
    ''' Every base key matching a removal criterion is removed, so duplicate key names
    ''' (ByName) or duplicate KeyType/Value pairs (ByValue) in the base section are all removed
    ''' </summary>
    '''
    ''' <param name="baseSection">
    ''' The <c> iniSection </c> from which keys will be removed
    ''' </param>
    '''
    ''' <param name="sourceSection">
    ''' The <c> iniSection </c> providing the keys to be removed from
    ''' <paramref name="baseSection"/> <br />
    ''' </param>
    '''
    ''' <param name="menuOutput">
    ''' The <c> MenuSection </c> containing output to be displayed to the user
    ''' </param>
    '''
    ''' <param name="quiet">
    ''' Indicates whether to suppress all menu and log output, including the
    ''' "removal target not found" messages, leaving reporting to the caller.
    ''' Global operations set it, since they expect most sections not to match <br /><br />
    ''' Optional, Default: <c> False </c>
    ''' </param>
    '''
    ''' <returns>
    ''' The number of keys removed from <paramref name="baseSection"/>
    ''' </returns>
    Friend Function remKeys(baseSection As iniSection,
                                   sourceSection As iniSection,
                             ByRef menuOutput As MenuSection,
                          Optional quiet As Boolean = False) As Integer

        Dim sourceData As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Dim isByName = TransmuteRemoveKeyMode = RemoveKeyMode.ByName

        For Each key In sourceSection.Keys : sourceData.Add(If(isByName, key.Name, $"{key.KeyType}={key.Value}")) : Next

        Dim toRemove As New List(Of iniKey)
        Dim matched As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        For Each baseKey In baseSection.Keys

            Dim matchStr = If(isByName, baseKey.Name, $"{baseKey.KeyType}={baseKey.Value}")
            If Not sourceData.Contains(matchStr) Then Continue For
            toRemove.Add(baseKey)
            matched.Add(matchStr)

        Next

        For Each key In toRemove

            baseSection.Keys.Remove(key)

            If quiet Then Continue For

            Dim remKeyMsg = $"  -= Removed key by {If(isByName, "name", "value")}: {key.ToString()}"
            menuOutput.AddColoredLine(remKeyMsg, ConsoleColor.Red)
            gLog(remKeyMsg)

        Next

        If quiet Then Return toRemove.Count

        For Each key In sourceData

            If matched.Contains(key) Then Continue For

            Dim err = $"Removal target not found: {key} not found in {baseSection.Name}"
            menuOutput.AddColoredLine(err, ConsoleColor.Red)
            gLog(err)

        Next

        Return toRemove.Count

    End Function

    ''' <summary>
    ''' Applies a 'flavor' to an <c> iniFile </c>. A flavor is like a compatibility layer that
    ''' allows a base file to be adjusted in slight ways to make it more suitable to a specific use
    ''' <br /><br /> 
    ''' In the context of winapp2.ini, we can use this to create different versions of winapp2.ini 
    ''' for different use cases (eg. a CCleaner or BleachBit specific versions) or else to 
    ''' correct the output of generative components of winapp2ool <br /><br />
    ''' 
    ''' Always applies Flavorings in the following order: <br /><br />
    ''' Section Removal -> Key Name Removal -> Key Value Removal -> Section Replacement ->
    ''' Key Replacement -> Section and Key Additions <br /><br />
    '''
    ''' Each stage is a <see cref="RemoteTransmute"/> pass in that stage's mode, so global
    ''' sections in a flavor file act under that mode too. <c> [*Map:] </c> rules therefore only
    ''' apply from <paramref name="keyReplacementFile"/>. A stage whose file is <c> Nothing </c>
    ''' or empty is skipped. We write the output even when every stage was skipped.
    ''' </summary>
    '''
    ''' <param name="baseFile">
    ''' The <c> iniFile </c> to whom a particular flavor will be applied. When
    ''' <paramref name="isWinapp"/> is <c> True </c>, we replace it with the sorted,
    ''' reformatted result.
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
    ''' which should be added to <paramref name="baseFile"/> to create the flavor <br /><br />
    ''' Optional, Default: <c> Nothing </c>
    ''' </param>
    ''' 
    ''' <param name="sectionRemovalFile">
    ''' The <c> iniFile </c> containing the set of sections to be removed from the 
    ''' <paramref name="baseFile"/> to create the flavor <br /> 
    ''' Sections will be removed regardless of whether or not keys are provided <br /><br />
    ''' Optional, Default: <c> Nothing </c>
    ''' </param>
    ''' 
    ''' <param name="keyNameRemovalFile">
    ''' The <c> iniFile </c> containing the set of individual keys to be removed 
    ''' from the <paramref name="baseFile"/> when matched by their Name parameter 
    ''' to create the flavor <br />
    ''' The values provided for keys in this file do not matter and will not be used for matching <br /><br />
    ''' Optional, Default: <c> Nothing </c>
    ''' </param>
    ''' 
    ''' <param name="keyValueRemovalFile">
    ''' The <c> iniFile </c> containing the set of individual keys to be removed from the 
    ''' <paramref name="baseFile"/> when matched by their KeyType and Value pairs
    ''' to create the flavor <br />
    ''' Numbers in key names will be ignored in this file and can be omitted. Numberless name
    ''' and value pairs will be used for matching. <br /><br />
    ''' Optional, Default: <c> Nothing </c>
    ''' </param>
    ''' 
    ''' <param name="sectionReplacementFile">
    ''' The <c> iniFile </c> containing the set of sections to replace entire sections of a
    ''' matching (case-insensitive) name in the <paramref name="baseFile"/>
    ''' to create the flavor <br />
    ''' This will not preserve non-overlapping content from the base section <br /><br />
    ''' Optional, Default: <c> Nothing </c>
    ''' </param>
    '''
    ''' <param name="keyReplacementFile">
    ''' The <c> iniFile </c> containing the set of individual keys to replace within a matching
    ''' section within <paramref name="baseFile"/> to create the flavor <br />
    ''' Keys in this file will only replace keys in the base file if both their Section and Key
    ''' names match (case-insensitive) <br /><br />
    ''' Optional, Default: <c> Nothing </c>
    ''' </param>
    ''' 
    ''' <param name="isWinapp">
    ''' Indicates whether to sort and format the output as a winapp2.ini file. When
    ''' <c> False </c>, we write a plain ini file with its sections in alphabetical order. <br /><br />
    ''' Optional, Default: <c> True </c>
    ''' </param>
    '''
    Public Sub Flavorize(ByRef baseFile As iniFile,
                               outputFile As iniFile,
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

            RemoteTransmute(baseFile, flavorFile, outputFile, isWinapp, menuOutput, curMode, curReplMode, curRemMode, curRemKMode, skipFormat:=True)

        Next

        Dim saved As Boolean

        If isWinapp Then

            Dim wf As New winapp2file(baseFile)
            wf.SortEntries()
            saved = outputFile.OverwriteToFile(wf.ToWinapp2String())
            baseFile = wf.ToIni()

        Else

            saved = outputFile.OverwriteToFile(baseFile.ToString(IniFileWriteFormat.Alphabetical))

        End If

        If Not saved Then

            menuOutput.AddWarning($"{outputFile.Name} was not saved")
            Return

        End If

        Dim flavorizedMsg = $"{baseFile.Name} Flavorized"
        menuOutput.AddColoredLine(flavorizedMsg, ConsoleColor.Magenta)
        gLog(flavorizedMsg)

    End Sub

End Module
