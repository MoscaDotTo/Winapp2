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
Module trimmainmenu

    ''' <summary> 
    ''' Prints the <c> Trim </c> menu to the user 
    ''' </summary>
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
            .AddColoredLine($"winapp2.ini:   {If(DownloadFileToTrim, GetNameFromDL(DownloadFileToTrim), replDir(TrimFile1.Path))}", ConsoleColor.Magenta) _
            .AddColoredFileInfo($"save target:   ", TrimFile3.Path, ConsoleColor.Yellow) _
            .AddFileInfo($"Includes path: ", TrimFile2.Path, condition:=UseTrimIncludes) _
            .AddFileInfo($"Excludes path: ", TrimFile4.Path, condition:=UseTrimExcludes) _
            .AddBlank(TrimModuleSettingsChanged) _
            .AddResetOpt(NameOf(Trim), TrimModuleSettingsChanged)

        menu.Print()

    End Sub

    ''' <summary> 
    ''' Handles the user input from the menu 
    ''' </summary>
    ''' 
    ''' <param name="input">
    ''' The user's input 
    ''' </param>
    Public Sub handleTrimUserInput(input As String)

        If input Is Nothing Then argIsNull(NameOf(input)) : Return

        Dim toggles = getToggleOpts()
        Dim toggleNums = getMenuNumbering(toggles, 2)

        Dim fileOpts = getFileOpts()
        Dim fileNums = getMenuNumbering(fileOpts, If(isOffline, 4, 5))

        Select Case True

            ' Exit 
            ' Notes: Always "0"
            Case input = "0"

                exitModule()

            ' Run (default)
            ' Notes: Always "1", also triggered by no input if run conditions are otherwise satisfied
            Case (input = "1" OrElse input.Length = 0)

                initTrim()

            ' Toggles
            ' Downloading (unavailable when offline)
            ' Includes 
            ' Excludes 
            Case toggleNums.Contains(input)

                Dim i = CType(input, Integer) - 2

                Dim toggleMenuText = toggles.Keys(i)
                Dim toggleName = toggles(toggleMenuText)

                toggleModuleSetting(toggleMenuText, NameOf(Trim), GetType(trimsettings),
                                    toggleName, NameOf(TrimModuleSettingsChanged))

            ' File Selectors
            ' Notes: Downloading implies not offline
            ' Winapp2.ini (unavailable when downloading)
            ' Save Target
            ' Includes (unavailable when not including)
            ' Excludes (unavailable when not excluding)
            Case fileNums.Contains(input)

                Dim i = CType(input, Integer) - 2 - toggles.Count

                Dim fileName = fileOpts.Keys(i)
                Dim fileObj = fileOpts(fileName)

                changeFileParams(fileObj, TrimModuleSettingsChanged, NameOf(Trim), fileName, NameOf(TrimModuleSettingsChanged))

            ' Reset Settings 
            ' Notes: Only available after a setting has been changed, always comes last in the menu
            Case TrimModuleSettingsChanged AndAlso CInt(input) = 2 + fileOpts.Count + fileOpts.Count + 1

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
    Private Function getToggleOpts() As Dictionary(Of String, String)

        Dim toggles As New Dictionary(Of String, String)

        If Not isOffline Then toggles.Add("Downloading", NameOf(DownloadFileToTrim))

        toggles.Add("Includes", NameOf(UseTrimIncludes))
        toggles.Add("Excludes", NameOf(UseTrimExcludes))

        Return toggles

    End Function

    ''' <summary>
    ''' Determines the current set of file selectors displayed on the menu and returns a Dictionary 
    ''' of those options and their respective files <br />
    ''' <br />
    ''' The set of possible files includes:
    ''' <list type="number">
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
    Private Function getFileOpts() As Dictionary(Of String, iniFile)

        Dim selectors As New Dictionary(Of String, iniFile)

        If Not DownloadFileToTrim Then selectors.Add(NameOf(TrimFile1), TrimFile1)

        selectors.Add(NameOf(TrimFile3), TrimFile3)

        If UseTrimIncludes Then selectors.Add(NameOf(TrimFile2), TrimFile2)
        If UseTrimExcludes Then selectors.Add(NameOf(TrimFile4), TrimFile4)

        Return selectors

    End Function

End Module
