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
''' Displays the Downloader Module main menu to the user and handles their input 
''' </summary>
Module downloadmainmenu

    ''' <summary> 
    ''' Restores the default state of the module's parameters 
    ''' </summary>
    Private Sub initDefaultSettings()

        downloadFile.resetParams()
        DownloadModuleSettingsChanged = False
        restoreDefaultSettings(NameOf(Downloader), AddressOf createDownloadSettingsSection)

    End Sub

    ''' <summary> 
    ''' Prints the download menu to the user 
    ''' </summary>
    Public Sub printDownloadMainMenu()

        Dim MenuDesc = {"Download files from the winapp2 GitHub"}

        Dim menu = MenuSection.CreateCompleteMenu(NameOf(Downloader), MenuDesc, ConsoleColor.DarkGreen)

        menu.AddBlank _
            .AddOption("Winapp2.ini", "Download the latest CCleaner-Flavor winapp2.ini") _
            .AddOption("Non-CCleaner", "Download the latest winapp2.ini").AddBlank _
            .AddOption("Winapp2ool", "Download the latest winapp2ool.exe").AddBlank _
            .AddOption("Directory", "Change the save directory").AddBlank _
            .AddOption("Advanced", "Additional downloads for power users").AddBlank _
            .AddOption("ReadMe", "Download the top-level winapp2ool readme").AddBlank _
            .AddColoredFileInfo("Save Directory: ", downloadFile.Dir, ConsoleColor.DarkYellow) _
            .AddResetOpt(NameOf(Downloader), DownloadModuleSettingsChanged)

        menu.Print()

    End Sub

    ''' <summary> 
    ''' Handles user input for the Downloader main menu 
    ''' </summary>
    ''' 
    ''' <param name="input"> 
    ''' The user's input 
    ''' </param>
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
                If denyActionWithHeader(DotNetFrameworkOutOfDate, "This option requires a newer version of the .NET Framework") Then Return
                If denyActionWithHeader(cantDownloadExecutable And downloadFile.Dir = Environment.CurrentDirectory, "Unable to download winapp2ool to the current directory, choose another directory before trying again") Then Return

                If downloadFile.Dir = Environment.CurrentDirectory Then

                    autoUpdate()

                Else

                    downloadFile.Name = "winapp2ool.exe"
                    download(downloadFile, toolExeLink)

                End If

            Case "4"

                Dim tmp = downloadFile.Dir

                initModule("Directory Chooser", AddressOf downloadFile.printDirChooserMenu, AddressOf downloadFile.handleDirChooserInput)

                Dim headerTxt = "Directory change aborted"
                setNextMenuHeaderText(headerTxt, printColor:=ConsoleColor.Red)

                If tmp = downloadFile.Dir Then Return

                headerTxt = "Save directory changed"
                    setNextMenuHeaderText(headerTxt, printColor:=ConsoleColor.Green)

                    DownloadModuleSettingsChanged = True
                    updateSettings(NameOf(Downloader), NameOf(downloadFile) & "_Dir", downloadFile.Dir)
                    updateSettings(NameOf(Downloader), NameOf(DownloadModuleSettingsChanged), DownloadModuleSettingsChanged.ToString(System.Globalization.CultureInfo.InvariantCulture))

            Case "5"

                initModule("Advanced Downloads", AddressOf printAdvMenu, AddressOf handleAdvInput)

            Case "6"

                downloadFile.Name = "readme.md"
                download(downloadFile, readMeLink)

            Case "7"

                If DownloadModuleSettingsChanged Then initDefaultSettings()

            Case Else

                setNextMenuHeaderText(invInpStr, printColor:=ConsoleColor.Red)

        End Select

    End Sub

End Module
