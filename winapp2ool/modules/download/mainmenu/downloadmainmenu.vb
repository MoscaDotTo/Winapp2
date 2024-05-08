'    Copyright (C) 2018-2024 Hazel Ward
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
''' <summary> Displays the Downloader Module main menu to the user and handles their input </summary>
''' Docs last updated: 2020-09-14
Module downloadmainmenu

    ''' <summary> Restores the default state of the module's parameters </summary>
    ''' Docs last updated: 2020-09-14 | Code last updated: 2020-09-14
    Private Sub initDefaultSettings()
        downloadFile.resetParams()
        DownloadModuleSettingsChanged = False
        restoreDefaultSettings(NameOf(Downloader), AddressOf createDownloadSettingsSection)
    End Sub

    ''' <summary> Prints the download menu to the user </summary>
    ''' Docs last updated: 2020-09-14 | Code last updated: 2021-11-16
    Public Sub printDownloadMainMenu()
        printMenuTop({"Download files from the winapp2 GitHub"})
        print(1, "Winapp2.ini", "Download the latest winapp2.ini")
        print(1, "Non-CCleaner", "Download the latest non-ccleaner winapp2.ini")
        print(1, "Winapp2ool", "Download the latest winapp2ool.exe")
        print(1, "Removed Entries.ini", "Download only entries used to create the non-ccleaner winapp2.ini", leadingBlank:=True)
        print(1, "Directory", "Change the save directory", trailingBlank:=True)
        print(1, "Advanced", "Additional downloads for power users")
        print(1, "ReadMe", "The winapp2ool ReadMe")
        print(0, $"Save directory: {replDir(downloadFile.Dir)}", leadingBlank:=True, closeMenu:=Not DownloadModuleSettingsChanged)
        print(2, NameOf(Downloader), cond:=DownloadModuleSettingsChanged, closeMenu:=True)
    End Sub

    ''' <summary> Handles user input for the Downloader main menu </summary>
    ''' <param name="input"> The user's input </param>
    ''' Docs last updated: 2020-09-14 | Code last updated: 2020-09-14
    Public Sub handleDownloadUserInput(input As String)
        Select Case input
            Case "0"
                exitModule()
            Case "1", "2"
                downloadFile.Name = "winapp2.ini"
                Dim link = If(input = "1", wa2Link, nonccLink)
                download(downloadFile, link)
                If downloadFile.Dir = Environment.CurrentDirectory Then checkedForUpdates = False
            Case "3"
                ' Feature gate downloading the executable behind .NET 4.6+
                If Not denyActionWithHeader(DotNetFrameworkOutOfDate, "This option requires a newer version of the .NET Framework") Then
                    If Not denyActionWithHeader(cantDownloadExecutable And downloadFile.Dir = Environment.CurrentDirectory,
                                                "Unable to download winapp2ool to the current directory, choose another directory before trying again") Then
                        If downloadFile.Dir = Environment.CurrentDirectory Then
                            autoUpdate()
                        Else
                            downloadFile.Name = "winapp2ool.exe"
                            download(downloadFile, toolExeLink)
                        End If
                    End If
                End If
            Case "4"
                downloadFile.Name = "Removed entries.ini"
                download(downloadFile, removedLink)
            Case "5"
                Dim tmp = downloadFile.Dir
                initModule("Directory Chooser", AddressOf downloadFile.printDirChooserMenu, AddressOf downloadFile.handleDirChooserInput)
                Dim headerTxt = "Directory change aborted"
                If Not tmp = downloadFile.Dir Then
                    headerTxt = "Save directory changed"
                    updateSettings(NameOf(Downloader), NameOf(downloadFile) & "_Dir", downloadFile.Dir)
                    DownloadModuleSettingsChanged = True
                    updateSettings(NameOf(Downloader), NameOf(DownloadModuleSettingsChanged), DownloadModuleSettingsChanged.ToString(System.Globalization.CultureInfo.InvariantCulture))
                    saveSettingsFile()
                End If
                setHeaderText(headerTxt)
            Case "6"
                initModule("Advanced Downloads", AddressOf printAdvMenu, AddressOf handleAdvInput)
            Case "7"
                ' It's actually a .md but the user doesn't need to know that  
                downloadFile.Name = "Readme.txt"
                download(downloadFile, readMeLink)
            Case "8"
                If DownloadModuleSettingsChanged Then initDefaultSettings()
            Case Else
                setHeaderText(invInpStr, True)
        End Select
    End Sub

End Module
