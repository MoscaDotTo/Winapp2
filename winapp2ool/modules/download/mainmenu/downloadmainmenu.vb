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
''' Displays the Downloader module main menu and handles user input
''' </summary>
Module downloadmainmenu

    ''' <summary>
    ''' Builds the Downloader main menu. Each winapp2.ini option downloads a fixed flavor and
    ''' ignores <see cref="CurrentWinappFlavor"/>.
    ''' </summary>
    Friend Function buildDownloadMenu() As MenuSection

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
            .AddDispatchedColoredOption("CCleaner 7 Winapp2.ini", "Download the latest CCleaner 7 Flavor of winapp2.ini", ConsoleColor.Magenta,
                Sub()
                    downloadFile.Name = "winapp2.ini"
                    download(downloadFile, cc7FlavorLink)
                    checkedForUpdates = False
                End Sub) _
            .AddDispatchedColoredOption("BleachBit Winapp2.ini", "Download the latest BleachBit Flavor of winapp2.ini", ConsoleColor.DarkCyan,
                Sub()
                    downloadFile.Name = "winapp2.ini"
                    download(downloadFile, bbFlavorLink)
                    checkedForUpdates = False
                End Sub) _
            .AddDispatchedColoredOption("FluentCleaner Winapp2.ini", "Download the latest FluentCleaner Flavor of winapp2.ini", ConsoleColor.Green,
                Sub()
                    downloadFile.Name = "winapp2.ini"
                    download(downloadFile, fcFlavorLink)
                    checkedForUpdates = False
                End Sub) _
            .AddDispatchedColoredOption("System Ninja Winapp2.rules", "Download the latest System Ninja Flavor of winapp2.ini", ConsoleColor.Blue,
                Sub()
                    downloadFile.Name = "winapp2.rules"
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
                    If denyActionWithHeader(cantDownloadExecutable AndAlso isRunningExeDir(downloadFile.Dir), "Unable to download winapp2ool to the current directory, choose another directory before trying again") Then Return
                    If isRunningExeDir(downloadFile.Dir) Then
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
                Sub() initModule("Advanced Downloads", AddressOf buildAdvMenu)) _
            .AddBlank() _
            .AddDispatchedColoredOption("Change Save Directory", "Select a new target directory for Downloader", ConsoleColor.DarkYellow,
                Sub() changeFileParams(downloadFile, DownloadModuleSettingsChanged, NameOf(Downloader), NameOf(downloadFile), NameOf(DownloadModuleSettingsChanged))) _
            .AddColoredFileInfo("Save Directory: ", downloadFile.Dir, ConsoleColor.DarkYellow) _
            .AddBlank(DownloadModuleSettingsChanged) _
            .AddDispatchedResetOpt(NameOf(Downloader), DownloadModuleSettingsChanged,
                Sub() resetModuleSettings(NameOf(Downloader), AddressOf InitDefaultDownloadSettings))

    End Function
End Module
