'    Copyright (C) 2018-2022 Hazel Ward
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
''' <summary> Displays the trim main menu to the user and handles their input </summary>
''' Docs last updated: 2020-11-24
Module trimmainmenu
    ''' <summary> Prints the <c> Trim </c> menu to the user </summary>
    ''' Docs last updated: 2022-11-21 | Code last updated: 2022-11-21
    Public Sub printTrimMenu()

        If isOffline Then DownloadFileToTrim = False

        printMenuTop({"Trim winapp2.ini such that it contains only entries relevant to your machine,", "greatly reducing both application load time and the winapp2.ini file size."})
        print(1, "Run (default)", "Trim winapp2.ini")
        print(5, "Toggle Download", "using the latest winapp2.ini from GitHub as the input file", Not isOffline, True, enStrCond:=DownloadFileToTrim, trailingBlank:=True)
        print(1, "File Chooser (winapp2.ini)", "Configure the path to winapp2.ini ", Not DownloadFileToTrim, isOffline, True)
        print(1, "File Chooser (save)", "Cofigure the path to which the trimmed winapp2.ini will be saved", trailingBlank:=True)
        print(5, "Toggle Includes", "always keeping certain entries", enStrCond:=UseTrimIncludes, trailingBlank:=Not UseTrimIncludes)
        print(1, "File Chooser (Includes)", "Configure the path to the includes file", cond:=UseTrimIncludes, trailingBlank:=True)
        print(5, "Toggle Excludes", "always discarding certain entries", enStrCond:=UseTrimExcludes, trailingBlank:=Not UseTrimExcludes)
        print(1, "File Chooser (Excludes)", "Configure the path to the excludes file", cond:=UseTrimExcludes, trailingBlank:=True)
        print(0, $"Current winapp2.ini path: {If(DownloadFileToTrim, GetNameFromDL(DownloadFileToTrim), replDir(TrimFile1.Path))}")
        print(0, $"Current save path: {replDir(TrimFile3.Path)}", closeMenu:=Not (UseTrimIncludes OrElse UseTrimExcludes OrElse ModuleSettingsChanged))
        print(0, $"Current includes path: {replDir(TrimFile2.Path)}", cond:=UseTrimIncludes, closeMenu:=Not (UseTrimExcludes OrElse ModuleSettingsChanged))
        print(0, $"Current excludes path: {replDir(TrimFile4.Path)}", cond:=UseTrimExcludes, closeMenu:=Not ModuleSettingsChanged)
        print(2, NameOf(Trim), cond:=ModuleSettingsChanged, closeMenu:=True)

    End Sub

    ''' <summary> Handles the user input from the menu </summary>
    ''' <param name="input"> The String containing the user's input </param>
    ''' Docs last updated: 2022-11-21 | Code last updated: 2022-11-21
    Public Sub handleTrimUserInput(input As String)

        If input Is Nothing Then argIsNull(NameOf(input)) : Return

        ' Both the Include and the Exclude options are toggled, applies +2 to all settings numbers following the excludes option in the menu 
        Dim IncAndExcl = UseTrimIncludes AndAlso UseTrimExcludes
        ' One and only one of the Include/Exclude options are toggled, applies +1 to all settings after the enabled option 
        Dim IncXorExcl = UseTrimIncludes Xor UseTrimExcludes

        Select Case True

            ' Option Name:                                 Exit 
            ' Option States:
            ' Default                                      -> 0 (Default) 
            Case input = "0"

                exitModule()

            ' Option Name:                                 Run (default) 
            ' Option States:
            ' Default                                      -> 1 (Default)
            Case (input = "1" OrElse input.Length = 0)

                initTrim()

            ' Option Name:                                  Toggle Download 
            ' Option States: 
            ' Offline                                      -> Unavailable (not displayed) (offsets all following menu options by -1 from their defaults)
            ' Online                                       -> 2 (default) 
            Case input = "2" AndAlso Not isOffline

                If Not denySettingOffline() Then toggleSettingParam(DownloadFileToTrim, "Downloading", ModuleSettingsChanged, NameOf(Trim), NameOf(DownloadFileToTrim), NameOf(ModuleSettingsChanged))

            ' Option Name:                                 File Chooser (winapp2.ini) 
            ' Option states: 
            ' Downloading                                  -> Unavailable (not displayed) (offsets all following menu options by -1 from their default ) 
            ' Offline (-1)                                 -> 2 
            ' Online                                       -> 3 (default) 
            Case Not DownloadFileToTrim AndAlso input = computeMenuNumber(3, {isOffline}, {-1})

                changeFileParams(TrimFile1, ModuleSettingsChanged, NameOf(Trim), NameOf(TrimFile1), NameOf(ModuleSettingsChanged))

            ' Option Name:                                 File Chooser (save) 
            ' Option states:
            ' Offline (-1) or Downloading (-1)             -> 3 [these two settings are mutually exclusive]
            ' Online, Not downloading                      -> 4 (default)
            ' 
            Case input = computeMenuNumber(4, {isOffline, DownloadFileToTrim}, {-1, -1})

                changeFileParams(TrimFile3, ModuleSettingsChanged, NameOf(Trim), NameOf(TrimFile3), NameOf(ModuleSettingsChanged))

            ' Option Name:                                  Reset Settings 
            ' Option States: 
            ' ModuleSettingsChanged = False                 -> Unavailable (not displayed) 
            ' Offline (-1), IncOrExcl = False               -> 6 (Available only if the user has toggled one of these two settings on and then off again) 
            ' Downloading (-1), and IncOrExcl = False       -> 6 
            ' Not Downloading, and IncOrExcl = False        -> 7 (default) 
            ' Downloading (-1), and IncXorExcl (+1) = True  -> 7
            ' Offline (-1), IncXorExclude = True            -> 7 
            ' Not Downloading, and IncXorExcl (+1) = True   -> 8 
            ' Downloading (-1), and IncAndExcl = True (+2)  -> 8
            ' Offline (-1), IncAndExcl (+2) = True          -> 8               
            ' Not Downloading, and IncAndExcl (+2) = True   -> 9 
            Case ModuleSettingsChanged AndAlso input = computeMenuNumber(7, {isOffline, DownloadFileToTrim, IncXorExcl, IncAndExcl}, {-1, -1, +1, +2})

                resetModuleSettings(NameOf(Trim), AddressOf initDefaultTrimSettings)

            ' Option Name:                                 Toggle Includes 
            ' Option states: 
            ' Downloading                                  -> 4
            ' Offline (-1) or Downloading (-1)             -> 4 [these two settings are mutually exclusive]
            ' Online, Not Downloading                      -> 5 (default)
            ' 
            Case Not isOffline AndAlso input = computeMenuNumber(5, {isOffline, DownloadFileToTrim}, {-1, -1})

                toggleSettingParam(UseTrimIncludes, "Includes", ModuleSettingsChanged, NameOf(Trim), NameOf(UseTrimIncludes), NameOf(ModuleSettingsChanged))

            ' Option Name:                                 File Chooser (Includes) 
            ' Option states: 
            ' Downloading (-1)                             -> 5 
            ' Online, Not Downloading                      -> 6 (default) 
            ' Offline                                      -> 5 (default offline) 

            Case UseTrimIncludes AndAlso input = computeMenuNumber(6, {isOffline, DownloadFileToTrim}, {-1, -1})

                changeFileParams(TrimFile2, ModuleSettingsChanged, NameOf(Trim), NameOf(TrimFile1), NameOf(ModuleSettingsChanged))

            ' Option Name:                                 Toggle Excludes 
            ' Option states: 
            ' Offline (-1), Not Including                  -> 5 (default offline)
            ' Online, Downloading (-1), Not including      -> 5
            ' Offline (-1), Including (+1)                 -> 6 
            ' Online, Not Downloading, Not Inlcuding       -> 6 (default online)  
            ' Online, Downloading (-1), Including (+1)     -> 6 
            ' Online, Not Downloading, Including           -> 7
            ' 
            Case input = computeMenuNumber(6, {isOffline, DownloadFileToTrim, UseTrimIncludes}, {-1, -1, 1})

                toggleSettingParam(UseTrimExcludes, "Excludes", ModuleSettingsChanged, NameOf(Trim), NameOf(UseTrimExcludes), NameOf(ModuleSettingsChanged))

            ' Option Name:                                 File Chooser (Exclude) 
            ' Option States:
            ' Online, Downloading (-1), Not Including      -> 6
            ' Offline (-1), Not including                  -> 6 (default offline)
            ' Online, Not Downloading, Not Including       -> 7 (default online)
            ' Online, Downloading (-1), Including (+1)     -> 7
            ' Offline (-1), Including (+1)                 -> 7
            ' Online, Not Downloading, Including (+1)      -> 8
            Case UseTrimExcludes AndAlso input = computeMenuNumber(7, {isOffline, DownloadFileToTrim, UseTrimIncludes}, {-1, -1, 1})

                changeFileParams(TrimFile4, ModuleSettingsChanged, NameOf(Trim), NameOf(TrimFile4), NameOf(ModuleSettingsChanged))

            Case Else

                setHeaderText(invInpStr, True)

        End Select

    End Sub

End Module
