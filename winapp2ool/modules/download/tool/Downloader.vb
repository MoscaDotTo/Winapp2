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
''' Downloader is a winapp2ool module which provides a very simple interface for downloading individual
''' files from the winapp2 GitHub in a way that can be automated through scripting
''' </summary>
Module Downloader

    ''' <summary> 
    ''' The web address of the winapp2.ini GitHub 
    ''' </summary>
    Public ReadOnly Property gitLink As String = "https://github.com/MoscaDotTo/Winapp2/"

    ''' <summary>
    ''' The web address of the base version winapp2.ini 
    ''' </summary>
    Public ReadOnly Property baseFlavorLink As String = "https://raw.githubusercontent.com/MoscaDotTo/Winapp2/master/Non-CCleaner/Winapp2.ini"

    ''' <summary>
    ''' The web address of the CCleaner version of winapp2.ini 
    ''' </summary>
    Public ReadOnly Property ccFlavorLink As String = "https://raw.githubusercontent.com/MoscaDotTo/Winapp2/master/Winapp2.ini"

    ''' <summary>
    ''' The web address of the System Ninja version of winapp2.ini 
    ''' </summary>
    Public ReadOnly Property snFlavorLink As String = "https://raw.githubusercontent.com/MoscaDotTo/Winapp2/master/Non-CCleaner/SystemNinja/Winapp2.rules"

    ''' <summary>
    ''' The web address of the BleachBit version of winapp2.ini 
    ''' </summary>
    Public ReadOnly Property bbFlavorLink As String = "https://raw.githubusercontent.com/MoscaDotTo/Winapp2/master/Non-CCleaner/BleachBit/Winapp2.ini"

    ''' <summary>
    ''' The web address of the Tron version of winapp2.ini 
    ''' </summary>
    Public ReadOnly Property tronFlavorLink As String = "https://raw.githubusercontent.com/MoscaDotTo/Winapp2/master/Non-CCleaner/Tron/Winapp2.ini"

    ''' <summary>
    ''' The web address of the CCleaner 7 version of winapp2.ini
    ''' </summary>
    Public ReadOnly Property cc7FlavorLink As String = "https://raw.githubusercontent.com/MoscaDotTo/Winapp2/master/Non-CCleaner/CCleaner7/Winapp2.ini"

    ''' <summary> 
    ''' The web address of winapp2ool.exe 
    ''' </summary>
    Public ReadOnly Property toolLink As String = "https://github.com/MoscaDotTo/Winapp2/raw/master/winapp2ool/bin/Release/winapp2ool.exe"

    ''' <summary> 
    ''' The web address of the beta build of winapp2ool.exe 
    ''' </summary>
    Public ReadOnly Property betaToolLink As String = "https://github.com/MoscaDotTo/Winapp2/raw/Branch1/winapp2ool/bin/Release/winapp2ool.exe"

    ''' <summary> 
    ''' The web address of version.txt (winapp2ool's public version identifer) 
    ''' </summary>
    Public ReadOnly Property toolVerLink As String = "https://raw.githubusercontent.com/MoscaDotTo/Winapp2/master/winapp2ool/version.txt"

    ''' <summary> 
    ''' The web address of winapp3.ini 
    ''' </summary>
    Public ReadOnly Property wa3link As String = "https://raw.githubusercontent.com/MoscaDotTo/Winapp2/master/Winapp3/Winapp3.ini"

    ''' <summary>
    ''' The web address of archived entries.ini 
    ''' </summary>
    Public ReadOnly Property archivedLink As String = "https://raw.githubusercontent.com/MoscaDotTo/Winapp2/master/Winapp3/Archived%20entries.ini"

    ''' <summary> 
    ''' The web address of the winapp2ool ReadMe file 
    ''' </summary>
    Public ReadOnly Property readMeLink As String = "https://raw.githubusercontent.com/MoscaDotTo/Winapp2/master/winapp2ool/Readme.md"

    ''' <summary>
    ''' The web address of the winapp2ool ReadMe for opening in web browsers
    ''' </summary>
    Public ReadOnly Property readMeUrl As String = "https://github.com/MoscaDotTo/Winapp2/blob/master/winapp2ool/Readme.md"

    ''' <summary>
    ''' Handles command line arguments for the Downloader module
    ''' </summary>
    '''
    ''' <remarks>
    ''' File arguments:
    ''' <list type="bullet">
    ''' <item><c> -1d path </c> — Set the save directory</item>
    ''' <item><c> -1f name </c> — Set the save filename</item>
    ''' </list>
    '''
    ''' Positional arguments (select the file to download):
    ''' <list type="bullet">
    ''' <item><c> 1 </c> or <c> winapp2 </c> — winapp2.ini (flavor-aware)</item>
    ''' <item><c> 2 </c> or <c> winapp2ool </c> — winapp2ool.exe</item>
    ''' <item><c> 3 </c> or <c> readme </c> — readme.txt</item>
    ''' <item><c> 4 </c> or <c> winapp3 </c> — winapp3.ini</item>
    ''' <item><c> 5 </c> or <c> archived </c> — Archived entries.ini</item>
    ''' </list>
    ''' </remarks>
    Public Sub handleCmdLine()

        Dim fileLink = ""

        Dim spec As New CliArgSpec(NameOf(Downloader))
        spec.WithFile(1, downloadFile) _
            .WithPositional(
                Sub(arg)
                    Select Case arg.ToUpperInvariant
                        Case "1", "WINAPP2"
                            fileLink = getWinappLink()
                            downloadFile.Name = If(fileLink.Contains("SystemNinja"), "winapp2.rules", "winapp2.ini")
                        Case "2", "WINAPP2OOL"
                            fileLink = toolExeLink()
                            downloadFile.Name = "winapp2ool.exe"
                        Case "3", "README"
                            fileLink = readMeLink
                            downloadFile.Name = "readme.txt"
                        Case "4", "WINAPP3"
                            fileLink = wa3link
                            downloadFile.Name = "winapp3.ini"
                        Case "5", "ARCHIVED"
                            fileLink = archivedLink
                            downloadFile.Name = "Archived entries.ini"
                        Case Else
                            cwl($"Unknown argument: {arg}", True)
                            cwl("Valid arguments are: winapp2, winapp2ool, readme, winapp3, archived", True)
                            Environment.Exit(1)
                    End Select
                End Sub) _
            .Parse()

        If downloadFile.Name = "winapp2ool.exe" AndAlso downloadFile.Dir = Environment.CurrentDirectory Then autoUpdate() : Return

        download(downloadFile, fileLink)

    End Sub

    ''' <summary>
    ''' Returns the link to winapp2ool.exe on the appropriate branch for the current tool configuration 
    ''' </summary>
    Public Function toolExeLink() As String

        Return If(isBeta, betaToolLink, toolLink)

    End Function

    ''' <summary>
    ''' Returns the link to the appropriate flavor of winapp2.ini based on the current Flavor configuration
    ''' </summary>
    Public Function getWinappLink() As String

        Select Case CurrentWinappFlavor

            Case WinappFlavor.NonCCleaner

                Return baseFlavorLink

            Case WinappFlavor.CCleaner

                Return ccFlavorLink

            Case WinappFlavor.BleachBit

                Return bbFlavorLink

            Case WinappFlavor.SystemNinja

                Return snFlavorLink

            Case WinappFlavor.Tron

                Return tronFlavorLink

            Case WinappFlavor.CCleaner7

                Return cc7FlavorLink

        End Select

        Return baseFlavorLink

    End Function

    ''' <summary> 
    ''' Returns the online download status (name) of winapp2.ini as a <c> String </c>, empty string if not downloading 
    ''' </summary>
    ''' 
    ''' <param name="shouldDownload"> 
    ''' Indicates that a module is configured to download a remote winapp2.ini 
    ''' </param>
    Public Function GetNameFromDL(shouldDownload As Boolean) As String

        Return If(shouldDownload, $"Online ({CurrentWinappFlavor.ToString} flavor)", "")

    End Function

End Module