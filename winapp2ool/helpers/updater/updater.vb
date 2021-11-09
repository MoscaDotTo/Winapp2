'    Copyright (C) 2018-2021 Hazel Ward
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
Imports System.IO
Imports System.Net
Imports System.Security
''' <summary> Holds functions used for checking for and updating winapp2.ini and winapp2ool.exe </summary>
Public Module updater

    ''' <summary> The latest available verson of winapp2ool from GitHub </summary>
    Public Property latestVersion As String = ""
    ''' <summary> The latest available version of winapp2.ini from GitHub </summary>
    Public Property latestWa2Ver As String = ""
    ''' <summary> The local version of winapp2.ini (if available) </summary>
    Public Property localWa2Ver As String = "000000"
    ''' <summary> Indicates that a winapp2ool update is available from GitHub </summary>
    Public Property updateIsAvail As Boolean = False
    ''' <summary> Indicates that a winapp2.ini update is available from GitHub </summary>
    Public Property waUpdateIsAvail As Boolean = False
    ''' <summary> The local version of winapp2ool </summary>
    Public Property currentVersion As String = ""
    ''' <summary> Indicates that an update check has been performed </summary>
    Public Property checkedForUpdates As Boolean = False

    ''' <summary> Pads the seconds portion of the version number, ensuring that it always have a length of 5 </summary>
    ''' <param name="version"> A version number to pad </param>
    Private Sub padVersionNum(ByRef version As String)
        Dim tmp = version.Split(CChar("."))
        Dim tmp1 = tmp.Last
        While tmp1.Length < 5
            tmp1 = "0" & tmp1
        End While
        version = version.Replace(tmp.Last, tmp1)
    End Sub

    Public Function getRemoteVersion(remotelink As String) As String
        If remotelink Is Nothing Then argIsNull(NameOf(remotelink)) : Return ""
        Dim tmpPath = setDownloadedFileStage(remotelink)
        Return getVersionFromLocalFile(tmpPath)
    End Function

    ''' <summary> Checks the versions of winapp2ool, .NET, and winapp2.ini and notes which, if any, are out of date </summary>
    ''' <param name="cond"> Indicates that the update check should be performed <br /> Optional, Default: <c> False </c> </param>
    Public Sub checkUpdates(Optional cond As Boolean = False)
        If checkedForUpdates Or Not cond Then Return
        gLog("Checking for updates")
        ' Query the latest winapp2ool.exe and winapp2.ini versions 
        toolVersionCheck()
        ' If winapp2.ini doesn't exist, an update is necessarily available. Avoid downloading in this case 
        ' anti virus vendors don't seem to like the fact that winapp2ool downloads a configuration file, particularly one containing 
        ' commands pertaining to yet more anti virus. If we can avoid doing this by default, we may be able to more easily fly under the radar 
        latestWa2Ver = getRemoteVersion(winapp2link)
        ' This should only be true if a user somehow has internet but cannot otherwise connect to the GitHub resources used to check for updates
        ' In this instance we should consider the update check to have failed and put the application into offline mode
        If latestVersion.Length = 0 Or latestWa2Ver.Length = 0 Then updateCheckFailed("online", True) : Return
        ' Observe whether or not updates are available, using val to avoid conversion mistakes
        updateIsAvail = Val(latestVersion.Replace(".", "")) > Val(currentVersion.Replace(".", ""))
        localWa2Ver = getVersionFromLocalFile()
        waUpdateIsAvail = Val(latestWa2Ver) > Val(localWa2Ver)
        checkedForUpdates = True
        gLog("Update check complete:")
        gLog($"Winapp2ool:")
        gLog("Local: " & currentVersion, indent:=True)
        gLog("Remote: " & latestVersion, indent:=True)
        gLog("Winapp2.ini:")
        gLog("Local: " & localWa2Ver, indent:=True)
        gLog("Remote: " & latestWa2Ver, indent:=True)
        Dim bothUpdatesAreAvail = waUpdateIsAvail And updateIsAvail
        Dim updHeader = $"Update{If(bothUpdatesAreAvail, "s", "")} available for {If(updateIsAvail, "winapp2ool ", "")}{If(bothUpdatesAreAvail, "and ", "")}{If(waUpdateIsAvail, "winapp2.ini", "")}"
        setHeaderText(updHeader, True, waUpdateIsAvail Or updateIsAvail, ConsoleColor.Green)
    End Sub

    '''<summary> Performs the version checking for winapp2ool.exe </summary>
    Private Sub toolVersionCheck()
        ' Let's just assume winapp2ool didn't update after we've checked for updates
        If Not latestVersion.Length = 0 Then Return
        If Not isBeta Then
            ' We use the txt file method for release builds to maintain support for update notifications on platforms that can't download executables
            latestVersion = getRemoteVersion(toolVerLink)
        Else
            If cantDownloadExecutable Then latestVersion = "000000 (update check disabled)" : Return
            If Not alreadyDownloadedExecutable Then
                Dim tmpPath = setDownloadedFileStage(betaToolLink)
                alreadyDownloadedExecutable = True
                ' This places a lock on winapp2ool.exe in the tmp folder that will remain until we close the application
                latestVersion = FileVersionInfo.GetVersionInfo(tmpPath).FileVersion
                ' If the build time is earlier than 2:46am (10000 seconds), the last part of the version number will be one or more digits short 
                ' Pad it with 0s when this is the case to avoid telling users there's an update available when there is not 
                padVersionNum(latestVersion)
                padVersionNum(currentVersion)
            End If
        End If
    End Sub

    ''' <summary> Handles the case where the update check has failed </summary>
    ''' <param name="name"> The name of the component whose update check failed </param>
    ''' <param name="chkOnline"> A flag specifying that the internet connection should be retested </param>
    Private Sub updateCheckFailed(name As String, Optional chkOnline As Boolean = False)
        setHeaderText($"/!\ {name} update check failed. /!\", True)
        localWa2Ver = "000000"
        If chkOnline Then chkOfflineMode()
    End Sub

    ''' <summary> Attempts to return the version number from a file found on disk, returns <c> "000000" </c> if it's unable to do so </summary>
    ''' <param name="path"> The path of the file whose version number will be queried </param>
    Private Function getVersionFromLocalFile(Optional path As String = "") As String
        ' Handle a special version.txt edge case
        If path.EndsWith("version.txt", StringComparison.InvariantCultureIgnoreCase) Then Return getFileDataAtLineNum(path)
        If path.Length = 0 Then path = Environment.CurrentDirectory & "\winapp2.ini"
        If Not File.Exists(path) Then Return "000000 (file not found)"
        Dim versionString = getFileDataAtLineNum(path)
        Return If(versionString.ToUpperInvariant.Contains("VERSION"), versionString.Split(CChar(" "))(2), "000000 (version not found)")
    End Function

    ''' <summary> Updates the offline status of winapp2ool </summary>
    Public Sub chkOfflineMode()
        gLog("Checking online status")
        isOffline = Not checkOnline()
    End Sub

    ''' <summary> Informs the user when an update is available </summary>
    ''' <param name="cond"> The update condition </param>
    ''' <param name="updName"> The item (winapp2.ini or winapp2ool) for which there is a pending update </param>
    ''' <param name="oldVer"> The old (currently in use) version </param>
    ''' <param name="newVer"> The updated version pending download </param>
    Public Sub printUpdNotif(cond As Boolean, updName As String, oldVer As String, newVer As String)
        If Not cond Then Return
        Dim tmpNewVer = newVer
        'If tmpNewVer = "999999" Then tmpNewVer = "20XXXX (latest online version)"
        gLog($"Update available for {updName} from {oldVer} to {tmpNewVer}")
        print(0, $"A new version of {updName} is available!", isCentered:=True, colorLine:=True, enStrCond:=True)
        print(0, $"Current: v{oldVer}", isCentered:=True, colorLine:=True, enStrCond:=True)
        print(0, $"Available: v{tmpNewVer}", trailingBlank:=True, isCentered:=True, colorLine:=True, enStrCond:=True)
    End Sub

    ''' <summary> Replaces the currently running executable with the latest from GitHub before launching that new executable and closing the current one,
    ''' ensures that this change can be undone by backing up the current version before replacing it </summary>
    Public Sub autoUpdate()
        gLog("Starting auto update process")
        Dim backupName = $"winapp2ool v{currentVersion}.exe.bak"
        Dim w2lName = "winapp2ool.exe"
        Try
            ' Ensure we always have the latest version
            Dim tmpToolPath = setDownloadedFileStage(toolExeLink)
            If Not File.Exists(tmpToolPath) Then Throw New WebException
            ' Replace any existing backups of this version before backing it up
            fDelete($"{Environment.CurrentDirectory}\{backupName}")
            File.Move(Environment.GetCommandLineArgs(0), backupName)
            ' Ensure that we don't have lingering winapp2ool.exes 
            fDelete(w2lName)
            ' Move the latest version to the current directory and launch it
            File.Move(tmpToolPath, $"{Environment.CurrentDirectory}\{w2lName}")
            Dim args = ""
            ' Pass any args that were used to start this instance of winapp2ool over to the next instance 
            If cmdargs.Count > 1 Then cmdargs.ForEach(Sub(arg) args += arg & ", ")
            ' Remove the trailing comma 
            If Not args.Length = 0 Then args = args.Remove(args.Length - 2)
            Process.Start(w2lName, args)
            Environment.Exit(0)
        Catch ex As IOException
            handleIOException(ex)
            File.Move(backupName, w2lName)
        Catch ex As WebException
            handleWebException(ex)
            File.Move(backupName, w2lName)
        End Try
    End Sub

    '''<summary> Deletes a file from the disk if it exists </summary>
    Public Sub fDelete(path As String)
        Try
            If File.Exists(path) Then File.Delete(path)
        Catch ex As IOException
            gLog("Failed to delete file")
            handleIOException(ex)
        End Try
    End Sub
End Module