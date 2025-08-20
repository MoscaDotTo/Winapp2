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
''' Displays the Trim module's main menu to the user and handles their input accordingly
''' </summary>
''' 
''' Docs last updated: 2025-08-12 
Module trimmainmenu

    ''' <summary> 
    ''' Prints the <c> Trim </c> menu to the user 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-08-11 | Code last updated: 2025-08-11
    Public Sub printTrimMenu()

        If isOffline Then DownloadFileToTrim = False

        Dim menuDescriptionLines = {"Trim winapp2.ini such that it only contains entries relevant to this machine"}

        Dim menu As New MenuSection
        menu = MenuSection.CreateCompleteMenu(NameOf(Trim), menuDescriptionLines, ConsoleColor.DarkCyan)

        menu.AddOption("Run (default)", "Optimize winapp2.ini for the current system").AddBlank() _
            .AddToggle($"Toggle downloading", "using the latest winapp2.ini from GitHub as the input file", isEnabled:=DownloadFileToTrim, condition:=Not isOffline) _
            .AddToggle($"Toggle using include list", "never trimming certain entries", isEnabled:=UseTrimIncludes) _
            .AddToggle($"Toggle using exclude list", "always trimming certain entries", isEnabled:=UseTrimExcludes).AddBlank() _
            .AddOption("Choose winapp2.ini", "Select a new winapp2.ini file for optimization", Not DownloadFileToTrim) _
            .AddOption("Choose save target", "Select a save target for the optimized winapp2.ini file") _
            .AddOption("Choose includes file", "Select a file containing entry names which should never be trimmed", UseTrimIncludes) _
            .AddOption("Choose excludes file", "Select a file containing entry names which should always be trimmed", UseTrimExcludes).AddBlank() _
            .AddLine($"Current winapp2.ini path: {If(DownloadFileToTrim, GetNameFromDL(DownloadFileToTrim), replDir(TrimFile1.Path))}") _
            .AddLine($"Current save path: {replDir(TrimFile3.Path)}") _
            .AddLine($"Current includes path: {replDir(TrimFile2.Path)}", condition:=UseTrimIncludes) _
            .AddLine($"Current excludes path: {replDir(TrimFile4.Path)}", condition:=UseTrimExcludes) _
            .AddBlank(TrimModuleSettingsChanged) _
            .AddResetOpt(NameOf(Trim), TrimModuleSettingsChanged) _
            .Print()

    End Sub

    ''' <summary> 
    ''' Handles the user input from the menu 
    ''' </summary>
    ''' 
    ''' <param name="input">
    ''' The user's input 
    ''' </param>
    ''' 
    ''' Docs last updated: 2025-08-12 | Code last updated: 2025-08-12
    Public Sub handleTrimUserInput(input As String)

        If input Is Nothing Then argIsNull(NameOf(input)) : Return

        ' Determine set of available toggles
        'Dim toggleOpts = If(isOffline, {"2", "3"}, {"2", "3", "4"})
        Dim baseToggles = getToggleOpts()
        Dim toggleOpts = getMenuNumbering(baseToggles, 2)

        ' Determine set of available file selectors
        Dim baseFileOpts = getFileOpts()
        Dim fileOpts = getMenuNumbering(baseFileOpts, If(isOffline, 4, 5))

        Select Case True

            ' Exit 
            ' Default -> 0 (Default)
            Case input = "0"

                exitModule()

            ' Run (default)
            ' Default -> 1 (Default) | No input 
            Case (input = "1" OrElse input.Length = 0)

                initTrim()

            ' Toggles
            ' 
            ' Download Toggle
            ' Offline -> Unavailable | Online -> 2 (default) 
            '
            ' Includes Toggle
            ' Offline -> 2 | Online -> 3 (default)
            ' 
            ' Excludes Toggle 
            ' Offline -> 3 | Online -> 4 (default)
            Case toggleOpts.Contains(input)

                Dim i = CType(input, Integer) - 2

                Dim toggleMenuText = baseToggles.Keys(i)
                Dim toggleName = baseToggles(toggleMenuText)

                toggleModuleSetting(toggleMenuText, NameOf(Trim), GetType(trimsettings),
                                    toggleName, NameOf(TrimModuleSettingsChanged))

            ' File Selectors
            ' Note: Downloading implies not offline
            ' Winapp2.ini                                 
            ' Downloading -> Unavailable | Offline -> 4 | Not downloading -> 5 (default)
            '
            ' Save Target
            ' Offline XOR Downloading -> 5 | Not downloading -> 6 (default)
            '
            ' Includes list   
            ' Not including -> Unavailable | Offline XOR Downloading -> 6 | Not downloading -> 7 (default)
            '
            ' Excludes list
            ' Not excluding -> Unavailable | Downloading XOR Offline, not including  -> 6
            ' Downloading XOR offline, including | not downloading, not including    -> 7 (default)
            ' Not downloading, including                                             -> 8
            Case fileOpts.Contains(input)

                Dim i = CType(input, Integer) - 2 - toggleOpts.Count

                If i > baseFileOpts.Count - 1 Then setHeaderText(invInpStr, True) : Return

                Dim fileName = baseFileOpts.Keys(i)
                Dim fileObj = baseFileOpts(fileName)

                changeFileParams(fileObj, TrimModuleSettingsChanged, NameOf(Trim), fileName, NameOf(TrimModuleSettingsChanged))

            ' Reset Settings 
            ' 
            ' Not ModuleSettingsChanged                                            -> Unavailable 
            ' Offline XOR Downloading, Not IncOrExcl                               -> 6  
            ' Not Downloading, Not IncOrExcl | Downloading XOR Offline, IncXorExcl -> 7  
            ' Not Downloading, IncXorExcl | Offline XOR Downloading, IncAndExcl    -> 8
            ' Not Downloading, IncAndExcl                                          -> 9 
            Case TrimModuleSettingsChanged AndAlso input = computeMenuNumber(7, {isOffline, DownloadFileToTrim, UseTrimIncludes, UseTrimExcludes}, {-1, -1, +1, +1})

                resetModuleSettings(NameOf(Trim), AddressOf initDefaultTrimSettings)

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
    '''     Downloading (not available when offline)
    '''     </item>
    '''     
    '''     <item>
    '''     Includes
    '''     </item>
    '''     
    '''     <item>
    '''     Excludes
    '''     </item>
    '''     
    ''' </list>
    '''  
    ''' </summary>
    ''' 
    ''' <returns>
    ''' The set of available toggles for the Trim module, with their names on the 
    ''' menu as keys and the respective property names as values
    ''' </returns>
    ''' 
    ''' Docs last updated: 2025-08-12 | Code last updated: 2025-08-12
    Private Function getToggleOpts() As Dictionary(Of String, String)

        Dim baseToggles As New Dictionary(Of String, String)

        If Not isOffline Then baseToggles.Add("Downloading", NameOf(DownloadFileToTrim))

        baseToggles.Add("Includes", NameOf(UseTrimIncludes))
        baseToggles.Add("Excludes", NameOf(UseTrimExcludes))

        Return baseToggles

    End Function

    ''' <summary>
    ''' Determines the current set of file selectors displayed on the menu and returns a Dictionary 
    ''' of those options and their respective files <br />
    ''' <br />
    ''' The set of possible files includes:
    ''' <list type="bullet">
    '''     
    '''     <item>
    '''     winapp2.ini (not available when downloading)
    '''     </item>
    '''     
    '''     <item>
    '''     Save target
    '''     </item>
    '''     
    '''     <item>
    '''     Includes file (not available when not using includes)
    '''     </item>
    '''     
    '''     <item>
    '''     Excludes file (not available when not using excludes)
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
    ''' Docs last updated: 2028-08-12 | Code last updated: 2025-08-12
    Private Function getFileOpts() As Dictionary(Of String, iniFile)

        Dim baseFileOpts As New Dictionary(Of String, iniFile)

        If Not DownloadFileToTrim Then baseFileOpts.Add(NameOf(TrimFile1), TrimFile1)

        baseFileOpts.Add(NameOf(TrimFile3), TrimFile3)

        If UseTrimIncludes Then baseFileOpts.Add(NameOf(TrimFile2), TrimFile2)
        If UseTrimExcludes Then baseFileOpts.Add(NameOf(TrimFile4), TrimFile4)

        Return baseFileOpts

    End Function

End Module
