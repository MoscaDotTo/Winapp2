'    Copyright (C) 2018-2019 Robbie Ward
' 
'    This file is a part of Winapp2ool
' 
'    Winapp2ool is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    Winap2ool is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with Winapp2ool.  If not, see <http://www.gnu.org/licenses/>.
Option Strict On
Imports System.IO
''' <summary>Holds functions used for checking for and updating winapp2.ini and winapp2ool.exe</summary>
Public Module updater
    ''' <summary>The latest available verson of winapp2ool from GitHub</summary>
    Public Property latestVersion As String = ""
    ''' <summary>The latest available version of winapp2.ini from GitHub</summary>
    Public Property latestWa2Ver As String = ""
    ''' <summary>The local version of winapp2.ini (if available)</summary>
    Public Property localWa2Ver As String = "000000"
    ''' <summary>Indicates that a winapp2ool update is available from GitHub</summary>
    Public Property updateIsAvail As Boolean = False
    ''' <summary>Indicates that a winapp2.ini update is available from GitHub</summary>
    Public Property waUpdateIsAvail As Boolean = False
    ''' <summary>The local version of winapp2ool</summary>
    Public Property currentVersion As String = Reflection.Assembly.GetExecutingAssembly.FullName.Split(CChar(","))(1).Substring(9)
    ''' <summary>Indicates that an update check has been performed</summary>
    Public Property checkedForUpdates As Boolean = False

    ''' <summary>Pads the seconds portion of the version number to always be length of 5</summary>
    ''' <param name="version">A version number to pad</param>
    Private Sub padVersionNum(ByRef version As String)
        Dim tmp = version.Split(CChar("."))
        Dim tmp1 = tmp.Last
        While tmp1.Length < 5
            tmp1 = "0" & tmp1
        End While
        version.Replace(tmp.Last, tmp1)
    End Sub

    ''' <summary>Checks the versions of winapp2ool, .NET, and winapp2.ini and records if any are outdated.</summary>
    ''' <param name="cond">Indicates that the update check should be performed. <br />Optional, Default: False</param>
    Public Sub checkUpdates(Optional cond As Boolean = False)
        If checkedForUpdates Or Not cond Then Exit Sub
        gLog("Checking for updates")
        ' Query the latest winapp2ool.exe and winapp2.ini versions 
        toolVersionCheck()
        latestWa2Ver = getFileDataAtLineNum(winapp2link).Split(CChar(" "))(2)
        ' This should only be true if a user somehow has internet but cannot otherwise connect to the GitHub resources used to check for updates
        ' In this instance we should consider the update check to have failed and put the application into offline mode
        If latestVersion = "" Or latestWa2Ver = "" Then updateCheckFailed("online", True) : Exit Sub
        ' Observe whether or not updates are available, using val to avoid conversion mistakes
        updateIsAvail = Val(latestVersion.Replace(".", "")) > Val(currentVersion.Replace(".", ""))
        getLocalWinapp2Version()
        waUpdateIsAvail = Val(latestWa2Ver) > Val(localWa2Ver)
        checkedForUpdates = True
        gLog("Update check complete:")
        gLog($"Winapp2ool:")
        gLog("Local: " & currentVersion, indent:=True)
        gLog("Remote: " & latestVersion, indent:=True)
        gLog("Winapp2.ini:")
        gLog("Local:" & localWa2Ver, indent:=True)
        gLog("Remote: " & latestWa2Ver, indent:=True)
    End Sub

    '''<summary>Performs the version chcking for winapp2ool.exe</summary>
    Private Sub toolVersionCheck()
        If Not isBeta Then
            latestVersion = getFileDataAtLineNum(toolVerLink)
        Else
            Dim tmpDir = Environment.GetEnvironmentVariable("temp")
            Dim tmpPath = $"{tmpDir}\winapp2ool.exe"
            fDelete(tmpPath)
            download(New iniFile($"{tmpDir}\", "winapp2ool.exe"), toolExeLink, False, True)
            latestVersion = Reflection.Assembly.LoadFile(tmpPath).FullName.Split(CChar(","))(1).Substring(9)
            ' If the build time is earlier than 2:46am (10000 seconds), the last part of the version number will be one or more digits short 
            ' Pad it with 0s when this is the case to avoid telling users there's an update available when there is not 
            padVersionNum(latestVersion)
            padVersionNum(currentVersion)
        End If
    End Sub

    ''' <summary>Handles the case where the update check has failed</summary>
    ''' <param name="name">The name of the component whose update check failed</param>
    ''' <param name="chkOnline">A flag specifying that the internet connection should be retested</param>
    Private Sub updateCheckFailed(name As String, Optional chkOnline As Boolean = False)
        setHeaderText($"/!\ {name} update check failed. /!\", True)
        localWa2Ver = "000000"
        If chkOnline Then chkOfflineMode()
    End Sub

    ''' <summary>Attempts to return the version number from the first line of winapp2.ini, returns "000000" if it can't</summary>
    Private Sub getLocalWinapp2Version()
        If Not File.Exists(Environment.CurrentDirectory & "\winapp2.ini") Then localWa2Ver = "000000 (File not found)" : Exit Sub
        Dim localStr = getFileDataAtLineNum(Environment.CurrentDirectory & "\winapp2.ini", remote:=False).ToLower
        If localStr.Contains("version") Then localWa2Ver = localStr.Split(CChar(" "))(2)
    End Sub

    ''' <summary>Updates the offline status of winapp2ool</summary>
    Public Sub chkOfflineMode()
        gLog("Checking online status")
        isOffline = Not checkOnline()
    End Sub

    ''' <summary>Informs the user when an update is available</summary>
    ''' <param name="cond">The update condition</param>
    ''' <param name="updName">The item (winapp2.ini or winapp2ool) for which there is a pending update</param>
    ''' <param name="oldVer">The old (currently in use) version</param>
    ''' <param name="newVer">The updated version pending download</param>
    Public Sub printUpdNotif(cond As Boolean, updName As String, oldVer As String, newVer As String)
        If Not cond Then Exit Sub
        gLog($"Update available for {updName} from {oldVer} to {newVer}")
        print(0, $"A new version of {updName} is available!", isCentered:=True, colorLine:=True, enStrCond:=True)
        print(0, $"Current  : v{oldVer}", isCentered:=True, colorLine:=True, enStrCond:=True)
        print(0, $"Available: v{newVer}", trailingBlank:=True, isCentered:=True, colorLine:=True, enStrCond:=True)
    End Sub

    ''' <summary>Downloads the latest version of winapp2ool.exe and replaces the currently running executable with it before launching that new executable and closing the program.</summary>
    Public Sub autoUpdate()
        gLog("Starting auto update process")
        Dim newTool As New iniFile(Environment.CurrentDirectory, "winapp2ool updated.exe")
        Dim backupName = $"winapp2ool v{currentVersion}.exe.bak"
        Try
            ' Replace any existing backups of this version
            fDelete($"{Environment.CurrentDirectory}\{backupName}")
            File.Move("winapp2ool.exe", backupName)
            ' Remove any old update files that didn't get renamed for whatever reason
            fDelete(newTool.Path)
            download(newTool, If(isBeta, betaToolLink, toolLink), False)
            ' Rename the executables and launch the new one
            File.Move("winapp2ool updated.exe", "winapp2ool.exe")
            System.Diagnostics.Process.Start($"{Environment.CurrentDirectory}\winapp2ool.exe")
            Environment.Exit(0)
        Catch ex As Exception
            exc(ex)
            File.Move(backupName, "winapp2ool.exe")
        End Try
    End Sub

    '''<summary>Deletes a file from the disk if it exists</summary>
    Public Sub fDelete(path As String)
        If File.Exists(path) Then File.Delete(path)
    End Sub
End Module