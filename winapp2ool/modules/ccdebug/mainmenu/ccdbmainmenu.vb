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
''' Display's the CCiniDebug's main menu and handles their input accordingly 
''' </summary>
''' 
''' Docs last updated: 2025-08-21
Module ccdbmainmenu

    ''' <summary> 
    ''' Prints the CCiniDebug menu to the user 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-08-21 | Code last updated: 2025-08-21
    Public Sub printCCDBMainMenu()

        Dim menuDescriptionLines As String() = {"Sort alphabetically the contents of ccleaner.ini and prune stale winapp2.ini settings"}

        Dim menu = MenuSection.CreateCompleteMenu(NameOf(CCiniDebug), menuDescriptionLines, ConsoleColor.Red)

        menu.AddColoredOption("Run (default)", "Debug ccleaner.ini", GetRedGreen(Not (PruneStaleEntries OrElse SaveDebuggedFile OrElse SortFileForOutput))).AddBlank _
        .AddToggle("Toggle pruning", "removal of orphaned winapp2.ini settings", isEnabled:=PruneStaleEntries) _
        .AddToggle("Toggle saving", "automatic saving of changes made by CCiniDebug", isEnabled:=SaveDebuggedFile) _
        .AddToggle("Toggle sorting", "alphabetical sorting of the contents of ccleaner.ini", isEnabled:=SortFileForOutput).AddBlank() _
        .AddOption("Choose winapp2.ini", "Select a new supplemental winapp2.ini file", condition:=PruneStaleEntries) _
        .AddOption("Choose ccleaner.ini", "Select a new ccleaner.ini file for debugging") _
        .AddOption("Choose save target", "Select a new save target for the debugged ccleaner.ini", condition:=SaveDebuggedFile).AddBlank() _
        .AddLine($"Current winapp2.ini:   {replDir(CCDebugFile1.Path)}", condition:=PruneStaleEntries) _
        .AddLine($"Current ccleaner.ini:  {replDir(CCDebugFile2.Path)}") _
        .AddLine($"Current save target:   {replDir(CCDebugFile3.Path)}", condition:=SaveDebuggedFile) _
        .AddBlank(CCDBSettingsChanged) _
        .AddResetOpt(NameOf(CCiniDebug), CCDBSettingsChanged)

        menu.Print()

    End Sub

    ''' <summary> 
    ''' Handles the user's input from the CCiniDebug main menu 
    ''' </summary>
    ''' 
    ''' <param name="input"> 
    ''' The user's input 
    ''' </param>
    ''' 
    ''' Docs last updated: 2020-07-18 | Code last updated: 2025-08-21
    Public Sub handleCCDBMUserInput(input As String)

        Dim toggles = getToggleOpts()
        Dim toggleNums = getMenuNumbering(toggles, 2)

        Dim fileOpts = getFileOpts()
        Dim fileNums = getMenuNumbering(fileOpts, 2 + toggles.Count)

        Dim resetNum = CType(2 + toggles.Count + fileOpts.Count, String)

        Select Case True

            ' Exit
            ' Notes: Always "0"
            Case input = "0"

                exitModule()

            ' Run (default)
            ' Notes: Always "1", also triggered by no input if run conditions are otherwise satisfied
            Case (input = "1" OrElse input.Length = 0)

                Dim noOptionsSelected = Not (PruneStaleEntries OrElse SaveDebuggedFile OrElse SortFileForOutput)
                If Not denyActionWithHeader(noOptionsSelected, "Please enable at least one options") Then initCCDebug()

            ' Toggles
            ' Pruning 
            ' Saving 
            ' Sorting 
            Case toggleNums.Contains(input)

                Dim i = CType(input, Integer) - 2

                Dim toggleMenuText = toggles.Keys(i)
                Dim toggleName = toggles(toggleMenuText)

                toggleModuleSetting(toggleMenuText, NameOf(CCiniDebug), GetType(ccdebugsettings),
                                    toggleName, NameOf(CCDBSettingsChanged))

            ' File selectors 
            ' winapp2.ini (unavailable when not pruning)
            ' ccleaner.ini
            ' Save target (unavailable when not saving)
            Case fileNums.Contains(input)

                Dim i = CType(input, Integer) - 2 - toggles.Count

                Dim fileName = fileOpts.Keys(i)
                Dim fileObj = fileOpts(fileName)

                changeFileParams(fileObj, CCDBSettingsChanged, NameOf(CCiniDebug), fileName, NameOf(CCDBSettingsChanged))

            ' Reset Settings
            ' Notes: Only available after a setting has been changed, always comes last in the option list
            Case CCDBSettingsChanged AndAlso input = resetNum

                resetModuleSettings("CCiniDebug", AddressOf initDefaultCCDBSettings)

            Case Else

                setHeaderText(invInpStr, True)

        End Select

    End Sub

    ''' <summary>
    ''' Determines the current set of toggles displayed on the menu and returns a Dictionary 
    ''' of those options and their respective toggle names <br />
    ''' <br />
    ''' The set of possible toggles includes:
    ''' <list type="bullet">
    '''     
    '''     <item>
    '''     Pruning 
    '''     </item>
    '''     
    '''     <item>
    '''     Saving
    '''     </item>
    '''     
    '''     <item>
    '''     Sorting
    '''     </item>
    '''     
    ''' </list>
    '''  
    ''' </summary>
    ''' 
    ''' <returns>
    ''' The set of available toggles for the CCiniDebug module, with their names on the 
    ''' menu as keys and the respective property names as values
    ''' </returns>
    ''' 
    ''' Docs last updated: 2025-08-20 | Code last updated: 2025-08-20
    Private Function getToggleOpts() As Dictionary(Of String, String)

        Dim toggles As New Dictionary(Of String, String)

        toggles.Add("Pruning", NameOf(PruneStaleEntries))
        toggles.Add("Saving", NameOf(SaveDebuggedFile))
        toggles.Add("Sorting", NameOf(SortFileForOutput))

        Return toggles

    End Function

    ''' <summary>
    ''' Determines the current set of file selectors displayed on the menu and returns a Dictionary 
    ''' of those options and their respective files <br />
    ''' <br />
    ''' The set of possible files includes:
    ''' <list type="bullet">
    '''     
    '''     <item>
    '''     winapp2.ini (unavailable when not pruning)
    '''     </item>
    '''     
    '''     <item>
    '''     ccleaner.ini 
    '''     </item>
    '''     
    '''     <item>
    '''     Save target (unavailable when not saving)
    '''     </item>
    '''     
    ''' </list>
    ''' 
    ''' </summary>
    ''' 
    ''' <returns> 
    ''' The set of <c> iniFile </c> properties for an object currently displayed on the menu
    ''' </returns>
    ''' 
    ''' Docs last updated: 2025-08-20 | Code last updated: 2025-08-20
    Private Function getFileOpts() As Dictionary(Of String, iniFile)

        Dim selectors As New Dictionary(Of String, iniFile)

        If PruneStaleEntries Then selectors.Add(NameOf(CCDebugFile1), CCDebugFile1)

        selectors.Add(NameOf(CCDebugFile2), CCDebugFile2)

        If SaveDebuggedFile Then selectors.Add(NameOf(CCDebugFile3), CCDebugFile3)

        Return selectors

    End Function

End Module
