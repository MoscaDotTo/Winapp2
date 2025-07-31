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
''' source file based on their KeyType (KeyName with numbers removed) <br/><br/>
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
''' 
''' </list>
''' 
''' </summary>
''' <remarks> 
''' Transmute replaces the Merge module and has has three modes which are wholly discrete from 
''' one another. This is a breaking change from old versions of Merge prior to 2025, which 
''' always add entries which don't exist and perform conflict resolution as part of that process.
''' 
''' As the functionality of Merge evolved, it no longer felt appropriate to refer to its output as
''' the result of a "merger" necessarily. We now consider the resulting output a Transmutation. 
''' Nevertheless, this is still spiritually the Merge module
''' </remarks>
''' 
''' Docs last updated: 2025-07-25 | Code last updated: 2025-07-25
Module Transmute

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
        ''' value from the source file. The key must exist in both files and have the same KeyType
        ''' (key name without numbers). Section names must match exactly (case sensitive)
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
    ''' The primary transmutation mode for the Transmute module <br />
    ''' Default: <c> <b> Add </b> </c>
    ''' 
    ''' <list>
    ''' 
    ''' <item>
    ''' <c> <b> Add </b> </c>
    ''' <description> 
    ''' Adds sections from the source file to the base file. If a section exists already in the base
    ''' file, the keys from the source section will be added to the base section. Section names 
    ''' must match exactly (case sensitive)
    ''' </description>
    ''' </item>
    ''' 
    ''' <item>
    ''' <c> <b> Replace </b></c>
    ''' <description>
    ''' Contains two sub modes. Replaces sections or individual keys in the base file with content 
    ''' from the source file. Section names must match exactly (case sensitive)
    ''' </description>
    ''' </item>
    ''' 
    ''' <item>
    ''' <c> <b> Remove </b> </c>
    ''' <description>
    ''' Contains two sub modes one of which itself contains two sub modes. Removes sections or 
    ''' individual keys from the base file. Key removal can be performed by value 
    ''' (ignores key numbering but requires a key type) or by key name (works best for unnumbered
    ''' keys) 
    ''' </description>
    ''' </item>
    ''' </list>
    '''
    ''' </summary>
    Public Property Transmutator As TransmuteMode = TransmuteMode.Add

    ''' <summary>
    ''' The granularity level for the <c> Replace Transmutator </c>, has two sub modes <br />
    ''' Default: ByKey 
    ''' <list> 
    ''' 
    ''' <item>
    ''' <c> <b> BySection </b></c>
    ''' <description>
    ''' Replaces entire sections when collisions occur
    ''' </description>
    ''' </item>
    ''' 
    ''' <item>
    ''' <c> <b> ByKey </b> </c>
    ''' <description>
    ''' Replaces individual keys when collisions occur 
    ''' </description>
    ''' </item>
    ''' </list>
    ''' 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-07-30 | Code last updated: 2025-07-30
    Public Property TransmuteReplaceMode As ReplaceMode = ReplaceMode.ByKey

    ''' <summary>
    ''' The granularity level for the <c> Remove Transmutator </c>, has two sub modes <br />
    ''' Default: ByKey 
    ''' <list> 
    ''' 
    ''' <item>
    ''' <c> <b> BySection </b></c>
    ''' <description>
    ''' Removes entire sections when collisions occur
    ''' </description>
    ''' </item>
    ''' 
    ''' <item>
    ''' <c> <b> ByKey </b> </c>
    ''' <description>
    ''' Removes individual keys when collisions occur 
    ''' </description>
    ''' </item>
    ''' </list>
    ''' 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-07-30 | Code last updated: 2025-07-30
    Public Property TransmuteRemoveMode As RemoveMode = RemoveMode.ByKey

    ''' <summary>
    ''' The granularity level for the <c> Remove by Key Transmutator </c>, has two sub modes <br />
    ''' Default: ByName
    ''' <list> 
    ''' 
    ''' <item>
    ''' <c> <b> ByName </b></c>
    ''' <description>
    ''' Removes keys from the base section if they have the same name as a key in the source section <br />
    ''' Ignores provided values
    ''' </description>
    ''' </item>
    ''' 
    ''' <item>
    ''' <c> <b> ByValue </b> </c>
    ''' <description>
    ''' Removes keys from the base section if they have the same KeyType and Value <br />
    ''' Ignores numbers in the Name of the <c> iniKey </c>
    ''' </description>
    ''' </item>
    ''' </list>
    ''' 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-07-30 | Code last updated: 2025-07-30
    Public Property TransmuteRemoveKeyMode As RemoveKeyMode = RemoveKeyMode.ByName

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
    ''' Preset base file choices <br />
    ''' 
    ''' -r              : removed entries.ini  <br />
    ''' -c              : custom.ini  <br />
    ''' -w              : winapp3.ini <br />
    ''' -a              : archived entries.ini <br /> 
    ''' 
    ''' </remarks>
    ''' 
    ''' Docs last updated: 2025-07-15 | Code last updated: 2025-07-15
    Public Sub handleCmdLine()

        initDefaultTransmuteSettings()

        ' primary transmute mode 
        Dim isAdd = False
        Dim isReplace = False
        Dim isRemove = False
        invertSettingAndRemoveArg(isAdd, "-add")
        invertSettingAndRemoveArg(isReplace, "-replace")
        invertSettingAndRemoveArg(isRemove, "-remove")
        Transmutator = If(isAdd, TransmuteMode.Add, If(isReplace, TransmuteMode.Replace, TransmuteMode.Remove))

        ' sub mode (shared remove/replace) 
        Dim bySection = False
        Dim byKey = True
        invertSettingAndRemoveArg(bySection, "-bysection")
        invertSettingAndRemoveArg(byKey, "-bykey")
        TransmuteReplaceMode = If(byKey, ReplaceMode.ByKey, ReplaceMode.BySection)
        TransmuteRemoveMode = If(byKey, RemoveMode.ByKey, RemoveMode.BySection)

        ' Remove key mode 
        Dim byName = True
        Dim byValue = False
        invertSettingAndRemoveArg(byName, "-byname")
        invertSettingAndRemoveArg(byValue, "-byvalue")
        TransmuteRemoveKeyMode = If(byName, RemoveKeyMode.ByName, RemoveKeyMode.ByValue)

        ' Preset file choices
        invertSettingAndRemoveArg(False, "-r", TransmuteFile2.Name, "Removed Entries.ini")
        invertSettingAndRemoveArg(False, "-c", TransmuteFile2.Name, "Custom.ini")
        invertSettingAndRemoveArg(False, "-w", TransmuteFile2.Name, "winapp3.ini")
        invertSettingAndRemoveArg(False, "-a", TransmuteFile2.Name, "Archived Entries.ini")

        getFileAndDirParams(TransmuteFile1, TransmuteFile2, TransmuteFile3)
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
        LogAndPrint(4, $"Applying changes to {TransmuteFile1.Name}")

        Dim remModeStr = $"{TransmuteRemoveMode}{If(TransmuteRemoveMode = RemoveMode.ByKey, $" - {TransmuteRemoveKeyMode.ToString}", "")}"
        Dim replaceModeStr = $"{TransmuteReplaceMode}"
        Dim xmuteModeStr As String = $"Transmutator: {Transmutator}{If(TransmuteModeIsAdd, If(Transmutator = TransmuteMode.Replace, replaceModeStr, remModeStr), "")}"

        Dim color = If(TransmuteModeIsAdd, ConsoleColor.Green, If(Transmutator = TransmuteMode.Remove, ConsoleColor.Red, ConsoleColor.Yellow))

        LogAndPrint(0, xmuteModeStr, colorLine:=True, useArbitraryColor:=True, arbitraryColor:=color)

        transmute(TransmuteFile1, TransmuteFile2, TransmuteFile3, True)

        print(0, "", closeMenu:=True)
        print(3, $"Changes applied. {anyKeyStr}")
        crk()

    End Sub

    ''' <summary> 
    ''' Conducts the transmutation on the base file using the source file and saves the 
    ''' output to the location provided within saveFile. <br /> Typically by default, the output is
    ''' saved to the base file, overwriting the original content. <br />
    ''' If working on a winapp2.ini file, the output can be formatted accordingly 
    ''' 
    ''' </summary>
    ''' 
    ''' <param name="isWinapp2"> 
    ''' Indicates that base file is a winapp2.ini syntax file 
    ''' </param>
    ''' 
    ''' Docs last updated: 2025-07-15 | Code last updated: 2025-07-15
    Private Sub transmute(ByRef baseFile As iniFile,
                          ByRef sourceFile As iniFile,
                          ByRef saveFile As iniFile,
                       Optional isWinapp2 As Boolean = True,
                       Optional sortBeforeSave As Boolean = True)

        ' Transmute the base file 
        resolveConflicts(baseFile, sourceFile)

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
    ''' Docs last updated: 2025-07-15 | Code last updated: 2025-07-15
    Public Sub RemoteTransmute(ByRef baseFile As iniFile,
                               ByRef sourceFile As iniFile,
                               ByRef outputFile As iniFile,
                                     isWinapp As Boolean,
                            Optional transmuteMode As TransmuteMode = TransmuteMode.Add,
                            Optional replaceMode As ReplaceMode = ReplaceMode.ByKey,
                            Optional removeMode As RemoveMode = RemoveMode.ByKey,
                            Optional removeKeyMode As RemoveKeyMode = RemoveKeyMode.ByName,
                            Optional quiet As Boolean = True)

        If baseFile.Sections.Count = 0 Then baseFile.init()

        If sourceFile.Sections.Count = 0 Then sourceFile.init()
        gLog("Source file is empty!", sourceFile.Sections.Count = 0)

        SuppressOutput = False

        Dim initTransmutator = Transmutator
        Dim initReplMode = TransmuteReplaceMode
        Dim initRemMode = TransmuteRemoveMode
        Dim initRemKeyMode = TransmuteRemoveKeyMode

        Transmutator = transmuteMode
        TransmuteReplaceMode = replaceMode
        TransmuteRemoveMode = removeMode
        TransmuteRemoveKeyMode = removeKeyMode

        transmute(baseFile, sourceFile, outputFile, isWinapp)

        Transmutator = initTransmutator
        TransmuteReplaceMode = initReplMode
        TransmuteRemoveMode = initRemMode
        TransmuteRemoveKeyMode = initRemKeyMode

        SuppressOutput = False

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
    ''' Docs last updated: 2025-07-15 | Code last updated: 2025-07-15
    Private Sub resolveConflicts(ByRef baseFile As iniFile,
                                 ByRef sourceFile As iniFile)

        For Each sectionName In sourceFile.Sections.Keys

            Dim baseFileHasSection = baseFile.Sections.Keys.Contains(sectionName)

            ' If we're not adding and the base file doesn't have a section of the same name, there's no conflicts to resolve 
            If Not baseFileHasSection AndAlso Not Transmutator = TransmuteMode.Add Then

                LogAndPrint(0, $"/!\ Target section not found in base file: [{sectionName}] - no changes applied /!\", colorLine:=True, useArbitraryColor:=True, arbitraryColor:=ConsoleColor.Yellow)
                Continue For

            End If

            Dim baseSection = If(baseFileHasSection, baseFile.Sections.Item(sectionName), New iniSection)

            processTransmutator(baseFile, sourceFile, sectionName, baseSection)

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
    ''' Docs last updated: 2025-07-23 | Code last updated: 2025-07-23
    Private Sub processTransmutator(ByRef baseFile As iniFile,
                                    ByRef sourceFile As iniFile,
                                          sectionName As String,
                                    ByRef baseSection As iniSection)

        Dim sourceSection = sourceFile.Sections.Item(sectionName)

        Select Case Transmutator

            Case TransmuteMode.Add

                handleAddMode(baseSection, sourceSection, baseFile, sectionName)

            Case TransmuteMode.Replace

                handleReplaceMode(baseSection, sourceSection, baseFile, sectionName)

            Case TransmuteMode.Remove

                handleRemoveMode(baseSection, sourceSection, baseFile, sectionName)

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
    ''' Docs last updated: 2025-07-15 | Code last updated: 2025-07-15
    Public Sub handleAddMode(ByRef baseSection As iniSection,
                             ByRef sourceSection As iniSection,
                             ByRef baseFile As iniFile,
                                   sectionName As String)

        If baseFile.Sections.ContainsKey(sectionName) Then addKeysToBase(baseSection, sourceSection) : Return

        baseFile.Sections.Add(sectionName, sourceSection)
        LogAndPrint(0, $"+§ Added new section: {sectionName}", colorLine:=True, useArbitraryColor:=True, arbitraryColor:=ConsoleColor.Magenta)

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
    ''' Docs last updated: 2025-07-15 | Code last updated: 2025-07-15
    Private Sub addKeysToBase(ByRef baseSection As iniSection,
                              ByRef sourceSection As iniSection)

        LogAndPrint(0, $"Adding keys to {sourceSection.Name}", colorLine:=True, useArbitraryColor:=True, arbitraryColor:=ConsoleColor.Cyan)

        For Each sourceKey In sourceSection.Keys.Keys

            baseSection.Keys.add(sourceKey)
            LogAndPrint(0, $"  += Added key: {sourceKey.Name}={sourceKey.Value}", colorLine:=True, enStrCond:=True)

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
    ''' Docs last updated: 2025-07-15 | Code last updated: 2025-07-15
    Private Sub handleReplaceMode(ByRef baseSection As iniSection,
                                  ByRef sourceSection As iniSection,
                                  ByRef baseFile As iniFile,
                                        sectionName As String)

        Dim logstr = $"Resolving collisons in {sectionName} ({Transmutator} - {TransmuteReplaceMode} mode)"
        LogAndPrint(0, logstr, colorLine:=True, useArbitraryColor:=True, arbitraryColor:=ConsoleColor.Cyan)

        ' if we're replacing by key, we can return immediately after
        If TransmuteReplaceMode = ReplaceMode.ByKey Then replaceKeysInBase(baseSection, sourceSection) : Return

        baseFile.Sections(sectionName) = sourceSection
        LogAndPrint(0, $"  *§ Replaced entire section: {sectionName}", colorLine:=True, enStrCond:=False)

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
    ''' Docs last updated: 2025-07-15 | Code last updated: 2025-07-15
    Private Sub replaceKeysInBase(ByRef baseSection As iniSection,
                                  ByRef sourceSection As iniSection)

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
            LogAndPrint(0, $"  * Replaced key: {baseKey.Name}", colorLine:=True, useArbitraryColor:=True, arbitraryColor:=ConsoleColor.Yellow)
            sourceKeys.Remove(baseKey.Name)

        Next

        ' report any missed replacement targets
        For Each key In sourceKeys.Values

            LogAndPrint(0, $"  /!\ Replacement target not found: {key.Name} not found in {baseSection.Name} /!\", colorLine:=True, useArbitraryColor:=True, arbitraryColor:=ConsoleColor.Yellow)

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
    ''' Docs last updated: 2025-07-15 | Code last updated: 2025-07-15
    Private Sub handleRemoveMode(ByRef baseSection As iniSection,
                                 ByRef sourceSection As iniSection,
                                 ByRef baseFile As iniFile,
                                       sectionName As String)

        Dim isKeyMode = TransmuteRemoveMode = RemoveMode.ByKey

        Dim logstr = $"Resolving conflicts in {sectionName} ({Transmutator} - {TransmuteRemoveMode} - {If(isKeyMode, TransmuteRemoveKeyMode.ToString, "")})"
        LogAndPrint(0, logstr, colorLine:=True, useArbitraryColor:=True, arbitraryColor:=ConsoleColor.Cyan)

        ' if we're removing keys we can immediately return when we're done 
        If isKeyMode Then remKeys(baseSection, sourceSection) : Return

        baseFile.Sections.Remove(sectionName)
        LogAndPrint(0, $"  -§ Removed entire section: {sectionName}", colorLine:=True, enStrCond:=False)

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
    ''' Docs last updated: 2025-07-30 | Code last updated: 2025-07-30
    Private Sub remKeys(ByRef baseSection As iniSection,
                              sourceSection As iniSection)

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
            LogAndPrint(0, $"  -= Removed key by {If(isByName, "name", "value")}: {baseKey.toString}", colorLine:=True, enStrCond:=False)
            sourceData.Remove(matchStr)

        Next

        ' report any missed removal targets 
        For Each key In sourceData

            LogAndPrint(0, $"  /!\ Removal target not found: {key} not found in {baseSection.Name} /!\", colorLine:=True, useArbitraryColor:=True, arbitraryColor:=ConsoleColor.Yellow)

        Next

    End Sub

End Module