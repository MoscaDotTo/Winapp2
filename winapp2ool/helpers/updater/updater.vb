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
Imports System.IO
Imports System.Net
Imports System.Reflection
Imports System.Security.Cryptography
Imports System.Text
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
        latestWa2Ver = getRemoteVersion(getWinappLink)
        ' This should only be true if a user somehow has internet but cannot otherwise connect to the GitHub resources used to check for updates
        ' In this instance we should consider the update check to have failed and put the application into offline mode
        If latestVersion.Length = 0 Or latestWa2Ver.Length = 0 Then updateCheckFailed("online", True) : Return

        ' The launcher only records currentVersion after the command line is processed, so -autoupdate arrives here without it
        If currentVersion.Length = 0 Then currentVersion = runningToolVersion()

        If parseToolVersion(latestVersion) Is Nothing OrElse parseToolVersion(currentVersion) Is Nothing Then

            gLog($"Unable to compare winapp2ool versions. Local: {currentVersion} Remote: {latestVersion}")
            updateCheckFailed("Winapp2ool")
            Return

        End If

        updateIsAvail = IsNewerToolVersion(latestVersion, currentVersion)
        localWa2Ver = getVersionFromLocalFile()
        waUpdateIsAvail = Val(latestWa2Ver) > Val(localWa2Ver)
        checkedForUpdates = True
        gLog("Update check complete:")
        gLog($"Winapp2ool:")
        gLog("  Local: " & currentVersion)
        gLog("  Remote: " & latestVersion)
        gLog("Winapp2.ini:")
        gLog("  Local: " & localWa2Ver)
        gLog("  Remote: " & latestWa2Ver)
        Dim bothUpdatesAreAvail = waUpdateIsAvail And updateIsAvail
        Dim updHeader = $"Update{If(bothUpdatesAreAvail, "s", "")} available for {If(updateIsAvail, "winapp2ool ", "")}{If(bothUpdatesAreAvail, "and ", "")}{If(waUpdateIsAvail, "winapp2.ini", "")}"
        setNextMenuHeaderText(updHeader, waUpdateIsAvail Or updateIsAvail, ConsoleColor.Green)

    End Sub

    '''<summary> Performs the version checking for winapp2ool.exe </summary>
    Private Sub toolVersionCheck()
        ' Let's just assume winapp2ool didn't update after we've checked for updates
        If Not latestVersion.Length = 0 Then Return
        latestVersion = getRemoteToolVersion(toolVersionLink())
    End Sub

    ''' <summary>
    ''' Downloads a winapp2ool <c> version.txt </c> into memory and returns its first line
    ''' </summary>
    '''
    ''' <param name="link">
    ''' The URL of the <c> version.txt </c> to read
    ''' </param>
    '''
    ''' <returns>
    ''' The trimmed first line of the file, <br />
    ''' an empty string if the download fails
    ''' </returns>
    Private Function getRemoteToolVersion(link As String) As String

        Try

            Using client As New WebClient

                Using reader As New StringReader(client.DownloadString(link))

                    Return If(reader.ReadLine(), "").Trim()

                End Using

            End Using

        Catch ex As WebException

            handleWebException(ex)
            Return ""

        End Try

    End Function

    ''' <summary>
    ''' Parses a winapp2ool version number such as <c> 1.7.9767.27788 </c>
    ''' </summary>
    '''
    ''' <param name="text">
    ''' The version text, surrounding whitespace allowed
    ''' </param>
    '''
    ''' <returns>
    ''' The parsed version, <br />
    ''' <c> Nothing </c> if <paramref name="text"/> is missing or isn't a version number
    ''' </returns>
    Friend Function parseToolVersion(text As String) As Version

        Dim parsed As Version = Nothing
        If text Is Nothing OrElse Not Version.TryParse(text.Trim(), parsed) Then Return Nothing

        Return parsed

    End Function

    ''' <summary>
    ''' Reports whether <paramref name="candidate"/> is a strictly newer winapp2ool version than <paramref name="running"/>
    ''' </summary>
    '''
    ''' <param name="candidate">
    ''' The version on offer
    ''' </param>
    '''
    ''' <param name="running">
    ''' The version currently running
    ''' </param>
    '''
    ''' <returns>
    ''' <c> True </c> if both parse and <paramref name="candidate"/> is greater, <br />
    ''' <c> False </c> otherwise
    ''' </returns>
    Friend Function IsNewerToolVersion(candidate As String,
                                       running As String) As Boolean

        Dim candidateVersion = parseToolVersion(candidate)
        Dim runningVersion = parseToolVersion(running)
        If candidateVersion Is Nothing OrElse runningVersion Is Nothing Then Return False

        Return candidateVersion > runningVersion

    End Function

    ''' <summary>
    ''' Returns the absolute path of the running winapp2ool executable
    ''' </summary>
    Friend Function runningExePath() As String

        Return Path.GetFullPath(Assembly.GetEntryAssembly().Location)

    End Function

    ''' <summary>
    ''' Returns the file version of the running winapp2ool executable
    ''' </summary>
    Private Function runningToolVersion() As String

        Return FileVersionInfo.GetVersionInfo(runningExePath()).FileVersion

    End Function

    ''' <summary>
    ''' Reports whether <paramref name="dir"/> is the folder that holds the running winapp2ool executable
    ''' </summary>
    '''
    ''' <param name="dir">
    ''' A folder path, absolute or relative to the current directory
    ''' </param>
    '''
    ''' <returns>
    ''' <c> True </c> if <paramref name="dir"/> resolves to the executable's folder, <br />
    ''' <c> False </c> otherwise, including when <paramref name="dir"/> isn't a usable path
    ''' </returns>
    Public Function isRunningExeDir(dir As String) As Boolean

        If String.IsNullOrWhiteSpace(dir) Then Return False

        Dim exeDir = Path.GetDirectoryName(runningExePath()).TrimEnd(Path.DirectorySeparatorChar)

        Try

            Return String.Equals(Path.GetFullPath(dir).TrimEnd(Path.DirectorySeparatorChar), exeDir, StringComparison.OrdinalIgnoreCase)

        Catch ex As ArgumentException

            Return False

        Catch ex As NotSupportedException

            Return False

        Catch ex As PathTooLongException

            Return False

        End Try

    End Function

    ''' <summary> Handles the case where the update check has failed </summary>
    ''' <param name="name"> The name of the component whose update check failed </param>
    ''' <param name="chkOnline"> A flag specifying that the internet connection should be retested </param>
    Private Sub updateCheckFailed(name As String, Optional chkOnline As Boolean = False)
        setNextMenuHeaderText($"/!\ {name} update check failed. /!\", printColor:=ConsoleColor.Red)
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

    ''' <summary>
    ''' Replaces the running winapp2ool executable with the latest signed build from GitHub, then relaunches it
    ''' with the same arguments and exits. The running version is kept beside the new one as
    ''' <c> winapp2ool v&lt;version&gt;.exe.bak </c>
    ''' <br /><br />
    '''
    ''' The download stays in memory until its signature checks out against <see cref="TrustedUpdateKeys"/>.
    ''' We then stage it next to the executable, confirm the staged copy is the bytes we verified, that it is
    ''' winapp2ool, and that it is a newer version, since an old build carries a valid signature too.
    ''' Any failure leaves the running executable in place and tells the user why
    ''' </summary>
    Public Sub autoUpdate()

        Using gLogScope("Starting auto update process")

            If TrustedUpdateKeys.Count = 0 Then

                updateFailed("This build of winapp2ool has no trusted update signing keys, so it can't verify an update. Download the latest winapp2ool.exe from GitHub instead")
                Return

            End If

            Dim exePath = runningExePath()
            Dim exeDir = Path.GetDirectoryName(exePath)
            Dim stagedPath = Path.Combine(exeDir, "winapp2ool.exe.new")
            Dim runningVersion = FileVersionInfo.GetVersionInfo(exePath).FileVersion
            Dim backupPath = Path.Combine(exeDir, $"winapp2ool v{runningVersion}.exe.bak")
            Dim installed = False

            Try

                installed = installVerifiedUpdate(exePath, stagedPath, backupPath, runningVersion)

            Catch ex As WebException

                handleWebException(ex)
                updateFailed("Winapp2ool was unable to download the update")

            Catch ex As UnauthorizedAccessException

                handleUnauthorizedAccessException(ex)
                updateFailed($"Winapp2ool doesn't have permission to replace itself in {exeDir}")

            Catch ex As IOException

                handleIOException(ex)
                updateFailed($"Winapp2ool was unable to replace itself in {exeDir}")

            Finally

                If Not installed Then discardStagedUpdate(stagedPath)

            End Try

            If installed Then relaunch(exePath)

        End Using

    End Sub

    ''' <summary>
    ''' Downloads, verifies, stages, and swaps in the latest winapp2ool executable
    ''' </summary>
    '''
    ''' <param name="exePath">
    ''' The absolute path of the running executable
    ''' </param>
    '''
    ''' <param name="stagedPath">
    ''' Where to write the new executable before it replaces the running one
    ''' </param>
    '''
    ''' <param name="backupPath">
    ''' Where to move the running executable
    ''' </param>
    '''
    ''' <param name="runningVersion">
    ''' The file version of the running executable
    ''' </param>
    '''
    ''' <returns>
    ''' <c> True </c> if the new executable now sits at <paramref name="exePath"/>, <br />
    ''' <c> False </c> if a check refused it, in which case the user has already been told why
    ''' </returns>
    Private Function installVerifiedUpdate(exePath As String,
                                           stagedPath As String,
                                           backupPath As String,
                                           runningVersion As String) As Boolean

        Dim newExe As Byte()
        Dim signatureText As String

        Using client As New WebClient

            newExe = client.DownloadData(toolExeLink())
            signatureText = client.DownloadString(toolExeSigLink())

        End Using

        If Not VerifyUpdateSignature(newExe, signatureText, TrustedUpdateKeys) Then

            updateFailed("The downloaded winapp2ool.exe doesn't carry a valid signature, so it was discarded")
            Return False

        End If

        gLog("Update signature verified")

        ' CreateNew refuses anything that reappears at the staged path after the delete, including a planted link
        File.Delete(stagedPath)

        Using staged As New FileStream(stagedPath, FileMode.CreateNew, FileAccess.Write, FileShare.None)

            staged.Write(newExe, 0, newExe.Length)

        End Using

        Dim problem = checkStagedUpdate(stagedPath, newExe, runningVersion)

        If problem.Length > 0 Then

            updateFailed(problem)
            Return False

        End If

        swapInStagedUpdate(exePath, stagedPath, backupPath)
        gLog($"Installed the update at {exePath}")

        Return True

    End Function

    ''' <summary>
    ''' Checks that the staged executable is the verified download, is winapp2ool, and is newer than the running version
    ''' </summary>
    '''
    ''' <param name="stagedPath">
    ''' The path of the staged executable
    ''' </param>
    '''
    ''' <param name="verifiedBytes">
    ''' The bytes whose signature was verified
    ''' </param>
    '''
    ''' <param name="runningVersion">
    ''' The file version of the running executable
    ''' </param>
    '''
    ''' <returns>
    ''' An empty string if the staged file passes, <br />
    ''' otherwise a message for the user saying why it was refused
    ''' </returns>
    Private Function checkStagedUpdate(stagedPath As String,
                                       verifiedBytes As Byte(),
                                       runningVersion As String) As String

        Using sha = SHA256.Create()

            If Not sha.ComputeHash(File.ReadAllBytes(stagedPath)).SequenceEqual(sha.ComputeHash(verifiedBytes)) Then Return "The update written to disk doesn't match the verified download"

        End Using

        Try

            If Not AssemblyName.GetAssemblyName(stagedPath).Name.Equals("winapp2ool", StringComparison.Ordinal) Then Return "The downloaded file isn't winapp2ool"

        Catch ex As BadImageFormatException

            Return "The downloaded file isn't a winapp2ool executable"

        End Try

        Dim stagedVersion = FileVersionInfo.GetVersionInfo(stagedPath).FileVersion

        If Not IsNewerToolVersion(stagedVersion, runningVersion) Then Return $"The downloaded winapp2ool (v{stagedVersion}) isn't newer than this one (v{runningVersion}), so it was not installed"

        Return ""

    End Function

    ''' <summary>
    ''' Moves the running executable aside as a backup and puts the staged one in its place,
    ''' restoring the backup if the second move fails
    ''' </summary>
    '''
    ''' <param name="exePath">
    ''' The absolute path of the running executable
    ''' </param>
    '''
    ''' <param name="stagedPath">
    ''' The path of the verified, staged executable
    ''' </param>
    '''
    ''' <param name="backupPath">
    ''' Where to move the running executable
    ''' </param>
    Private Sub swapInStagedUpdate(exePath As String,
                                   stagedPath As String,
                                   backupPath As String)

        File.Delete(backupPath)
        File.Move(exePath, backupPath)

        Dim swapped = False

        Try

            File.Move(stagedPath, exePath)
            swapped = True

        Finally

            If Not swapped AndAlso File.Exists(backupPath) AndAlso Not File.Exists(exePath) Then File.Move(backupPath, exePath)

        End Try

    End Sub

    ''' <summary>
    ''' Deletes a staged update that was not installed, logging rather than throwing if it can't
    ''' </summary>
    '''
    ''' <param name="stagedPath">
    ''' The path of the staged executable
    ''' </param>
    Private Sub discardStagedUpdate(stagedPath As String)

        Try

            File.Delete(stagedPath)

        Catch ex As IOException

            gLog($"Unable to delete the staged update at {stagedPath}: {ex.Message}")

        Catch ex As UnauthorizedAccessException

            gLog($"Unable to delete the staged update at {stagedPath}: {ex.Message}")

        End Try

    End Sub

    ''' <summary>
    ''' Starts the freshly installed executable with this process's original arguments and working directory,
    ''' then exits
    ''' </summary>
    '''
    ''' <param name="exePath">
    ''' The absolute path of the new executable
    ''' </param>
    Private Sub relaunch(exePath As String)

        Dim startInfo As New ProcessStartInfo(exePath) With {
            .UseShellExecute = False,
            .WorkingDirectory = Environment.CurrentDirectory,
            .Arguments = BuildArgumentString(Environment.GetCommandLineArgs().Skip(1))
        }

        gLog($"Relaunching {exePath}")

        Try

            Using Process.Start(startInfo)
            End Using

        Catch ex As ComponentModel.Win32Exception

            printAndLogExceptionForUser(ex.ToString, ex.GetType.ToString)
            updateFailed("Winapp2ool updated itself but couldn't restart. Please start it again")
            Return

        End Try

        Environment.Exit(0)

    End Sub

    ''' <summary>
    ''' Tells the user why an update did not happen and records it in the log. Silent and command line
    ''' runs also get a nonzero exit code
    ''' </summary>
    '''
    ''' <param name="reason">
    ''' The message for the user
    ''' </param>
    Private Sub updateFailed(reason As String)

        ' -autoupdate runs before the command line handler has consumed -s, so SuppressOutput may not be set yet
        Dim silent = SuppressOutput OrElse Environment.GetCommandLineArgs().Skip(1).Any(Function(a) a.Equals("-s", StringComparison.OrdinalIgnoreCase))

        gLog($"Update aborted: {reason}")
        cwl(reason, Not silent)
        setNextMenuHeaderText(reason, printColor:=ConsoleColor.Red)

        If silent OrElse IsCommandLineMode Then Environment.ExitCode = 1

    End Sub

    ''' <summary>
    ''' Quotes a single command line argument so that <c> CommandLineToArgvW </c> and the .NET runtime
    ''' read it back unchanged
    ''' </summary>
    '''
    ''' <param name="arg">
    ''' The argument to quote
    ''' </param>
    '''
    ''' <returns>
    ''' <paramref name="arg"/> as is when it needs no quoting, <br />
    ''' otherwise <paramref name="arg"/> in double quotes with its quotes and the backslashes before them escaped
    ''' </returns>
    Friend Function QuoteArgument(arg As String) As String

        If arg Is Nothing Then arg = ""

        If arg.Length > 0 AndAlso arg.IndexOfAny({" "c, ControlChars.Tab, ControlChars.Lf, ControlChars.VerticalTab, """"c}) < 0 Then Return arg

        Dim quoted As New StringBuilder("""")
        Dim backslashes = 0

        ' Backslashes are literal unless a quote follows them, in which case each one must be doubled
        For Each ch In arg

            If ch = "\"c Then

                backslashes += 1

            ElseIf ch = """"c Then

                quoted.Append("\"c, backslashes * 2 + 1).Append(""""c)
                backslashes = 0

            Else

                quoted.Append("\"c, backslashes).Append(ch)
                backslashes = 0

            End If

        Next

        quoted.Append("\"c, backslashes * 2).Append(""""c)

        Return quoted.ToString()

    End Function

    ''' <summary>
    ''' Joins arguments into a single command line string, quoting each with <see cref="QuoteArgument"/>
    ''' </summary>
    '''
    ''' <param name="args">
    ''' The arguments to join
    ''' </param>
    Friend Function BuildArgumentString(args As IEnumerable(Of String)) As String

        Return String.Join(" ", args.Select(AddressOf QuoteArgument))

    End Function

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