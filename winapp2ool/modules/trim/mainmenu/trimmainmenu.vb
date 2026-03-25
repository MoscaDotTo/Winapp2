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
''' Displays the <c> Trim </c> main menu and handles user input
''' </summary>
Module trimmainmenu

    ''' <summary>
    ''' Builds the <c> Trim </c> main menu with all options and their dispatch handlers registered inline.
    ''' Called by both <c> printTrimMenu </c> and <c> handleTrimUserInput </c> so the
    ''' displayed option numbers and the dispatch table are always in sync.
    ''' </summary>
    '''
    ''' <returns>
    ''' A fully configured <c> MenuSection </c> ready to print or dispatch
    ''' </returns>
    Private Function buildTrimMenu() As MenuSection

        If isOffline Then DownloadFileToTrim = False

        Dim menuDesc = {"Trim winapp2.ini such that it only contains entries relevant to this machine"}

        Return MenuSection.CreateCompleteMenu(NameOf(Trim), menuDesc, ConsoleColor.DarkCyan) _
            .AddDispatchedOption("Run (default)", "Optimize winapp2.ini for the current system", Sub() initTrim()) _
            .AddBlank() _
            .AddDispatchedToggle("downloading", "using the latest winapp2.ini from GitHub as the input file", DownloadFileToTrim,
                Sub() toggleModuleSetting("Downloading", NameOf(Trim), GetType(trimsettings), NameOf(DownloadFileToTrim), NameOf(TrimModuleSettingsChanged)),
                Not isOffline) _
            .AddDispatchedToggle("include list", "never trimming certain entries", UseTrimIncludes,
                Sub() toggleModuleSetting("Include list", NameOf(Trim), GetType(trimsettings), NameOf(UseTrimIncludes), NameOf(TrimModuleSettingsChanged))) _
            .AddDispatchedToggle("exclude list", "always trimming certain entries", UseTrimExcludes,
                Sub() toggleModuleSetting("Exclude list", NameOf(Trim), GetType(trimsettings), NameOf(UseTrimExcludes), NameOf(TrimModuleSettingsChanged))) _
            .AddBlank() _
            .AddDispatchedOption("Choose winapp2.ini", "Select a new winapp2.ini file for optimization",
                Sub() changeFile2Params(TrimFile1, TrimModuleSettingsChanged, NameOf(Trim), NameOf(TrimFile1), NameOf(TrimModuleSettingsChanged)),
                Not DownloadFileToTrim) _
            .AddDispatchedOption("Choose save target", "Select a save target for the optimized winapp2.ini file",
                Sub() changeFile2Params(TrimFile3, TrimModuleSettingsChanged, NameOf(Trim), NameOf(TrimFile3), NameOf(TrimModuleSettingsChanged))) _
            .AddDispatchedOption("Choose includes file", "Select a file containing entry names which should never be trimmed",
                Sub() changeFile2Params(TrimFile2, TrimModuleSettingsChanged, NameOf(Trim), NameOf(TrimFile2), NameOf(TrimModuleSettingsChanged)),
                UseTrimIncludes) _
            .AddDispatchedOption("Choose excludes file", "Select a file containing entry names which should always be trimmed",
                Sub() changeFile2Params(TrimFile4, TrimModuleSettingsChanged, NameOf(Trim), NameOf(TrimFile4), NameOf(TrimModuleSettingsChanged)),
                UseTrimExcludes) _
            .AddBlank() _
            .AddColoredLine($"winapp2.ini:   {If(DownloadFileToTrim, GetNameFromDL(DownloadFileToTrim), replDir(TrimFile1.Path()))}", ConsoleColor.Magenta) _
            .AddColoredFileInfo("save target:   ", TrimFile3.Path(), ConsoleColor.Yellow) _
            .AddFileInfo("Includes path: ", TrimFile2.Path(), condition:=UseTrimIncludes) _
            .AddFileInfo("Excludes path: ", TrimFile4.Path(), condition:=UseTrimExcludes) _
            .AddBlank(TrimModuleSettingsChanged) _
            .AddDispatchedResetOpt(NameOf(Trim), TrimModuleSettingsChanged, Sub() resetModuleSettings(NameOf(Trim), AddressOf InitDefaultTrimSettings))

    End Function

    ''' <summary>
    ''' Prints the <c> Trim </c> main menu to the user
    ''' </summary>
    Public Sub printTrimMenu()

        buildTrimMenu().Print()

    End Sub

    ''' <summary>
    ''' Handles user input from the <c> Trim </c> main menu
    ''' </summary>
    '''
    ''' <param name="input">
    ''' The user's input
    ''' </param>
    Public Sub handleTrimUserInput(input As String)

        Dim intInput As Integer

        If Not Integer.TryParse(input, intInput) Then

            If input.Length = 0 Then initTrim() : Return

            setNextMenuHeaderText(invInpStr, printColor:=ConsoleColor.Red)
            Return

        End If

        If intInput = 0 Then exitModule() : Return

        If Not buildTrimMenu().Dispatch(intInput) Then setNextMenuHeaderText(invInpStr, printColor:=ConsoleColor.Red)

    End Sub

End Module
