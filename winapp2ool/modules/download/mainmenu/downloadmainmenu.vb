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
''' Displays the Downloader module main menu and handles user input.
''' Called by both <c> printDownloadMainMenu </c> (to render) and <c> handleDownloadUserInput </c>
''' (to dispatch), so the displayed option numbers and the dispatch table are always in sync.
''' </summary>
Module downloadmainmenu

    ''' <summary>
    ''' Builds the Downloader main menu with all options and their dispatch handlers registered inline
    ''' </summary>
    Private Function buildDownloadMenu() As MenuSection

        Dim menuDesc = {"Download files from the winapp2 GitHub"}

        Return MenuSection.CreateCompleteMenu(NameOf(Downloader), menuDesc, ConsoleColor.DarkGreen) _
            .AddBlank() _
            .AddDispatchedColoredOption("Winapp2.ini", "Download the latest base winapp2.ini", ConsoleColor.Cyan,
                Sub()
                    downloadFile.Name = "winapp2.ini"
                    download(downloadFile, baseFlavorLink)
                    checkedForUpdates = False
                End Sub) _
            .AddDispatchedColoredOption("CCleaner Winapp2.ini", "Download the latest CCleaner Flavor of winapp2.ini", ConsoleColor.DarkRed,
                Sub()
                    downloadFile.Name = "winapp2.ini"
                    download(downloadFile, ccFlavorLink)
                    checkedForUpdates = False
                End Sub) _
            .AddDispatchedColoredOption("BleachBit Winapp2.ini", "Download the latest BleachBit Flavor of winapp2.ini", ConsoleColor.DarkCyan,
                Sub()
                    downloadFile.Name = "winapp2.rules"
                    download(downloadFile, bbFlavorLink)
                    checkedForUpdates = False
                End Sub) _
            .AddDispatchedColoredOption("System Ninja Winapp2.rules", "Download the latest System Ninja Flavor of winapp2.ini", ConsoleColor.Blue,
                Sub()
                    downloadFile.Name = "winapp2.ini"
                    download(downloadFile, snFlavorLink)
                    checkedForUpdates = False
                End Sub) _
            .AddDispatchedColoredOption("Tron", "Download the latest Tron flavor of winapp2.ini", ConsoleColor.Red,
                Sub()
                    downloadFile.Name = "winapp2.ini"
                    download(downloadFile, tronFlavorLink)
                    checkedForUpdates = False
                End Sub) _
            .AddBlank() _
            .AddDispatchedOption("Winapp2ool", "Download the latest winapp2ool.exe",
                Sub()
                    If denyActionWithHeader(DotNetFrameworkOutOfDate, "This option requires a newer version of the .NET Framework") Then Return
                    If denyActionWithHeader(cantDownloadExecutable And downloadFile.Dir = Environment.CurrentDirectory, "Unable to download winapp2ool to the current directory, choose another directory before trying again") Then Return
                    If downloadFile.Dir = Environment.CurrentDirectory Then
                        autoUpdate()
                    Else
                        downloadFile.Name = "winapp2ool.exe"
                        download(downloadFile, toolExeLink())
                    End If
                End Sub) _
            .AddBlank() _
            .AddDispatchedOption("ReadMe", "Download the top-level winapp2ool readme",
                Sub()
                    downloadFile.Name = "readme.md"
                    download(downloadFile, readMeLink)
                End Sub) _
            .AddBlank() _
            .AddDispatchedOption("Advanced", "Additional downloads for power users",
                Sub() initModule("Advanced Downloads", AddressOf printAdvMenu, AddressOf handleAdvInput)) _
            .AddBlank() _
            .AddDispatchedColoredOption("Change Save Directory", "Select a new target directory for Downloader", ConsoleColor.DarkYellow,
                Sub() changeFile2Params(downloadFile, DownloadModuleSettingsChanged, NameOf(Downloader), NameOf(downloadFile), NameOf(DownloadModuleSettingsChanged))) _
            .AddColoredFileInfo("Save Directory: ", downloadFile.Dir, ConsoleColor.DarkYellow) _
            .AddBlank(DownloadModuleSettingsChanged) _
            .AddDispatchedResetOpt(NameOf(Downloader), DownloadModuleSettingsChanged,
                Sub() resetModuleSettings(NameOf(Downloader), AddressOf InitDefaultDownloadSettings))

    End Function

    ''' <summary>
    ''' Prints the Downloader menu to the user
    ''' </summary>
    Public Sub printDownloadMainMenu()

        buildDownloadMenu().Print()

    End Sub

    ''' <summary>
    ''' Handles user input for the Downloader menu
    ''' </summary>
    '''
    ''' <param name="input">
    ''' The user's input
    ''' </param>
    Public Sub handleDownloadUserInput(input As String)

        Dim intInput As Integer

        If Not Integer.TryParse(input, intInput) Then

            setNextMenuHeaderText(invInpStr, printColor:=ConsoleColor.Red)
            Return

        End If

        If intInput = 0 Then exitModule() : Return

        If Not buildDownloadMenu().Dispatch(intInput) Then setNextMenuHeaderText(invInpStr, printColor:=ConsoleColor.Red)

    End Sub

End Module
