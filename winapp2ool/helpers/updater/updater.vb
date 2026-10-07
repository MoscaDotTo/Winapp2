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

''' <summary>
''' Checks GitHub for newer versions of winapp2ool and winapp2.ini, and replaces the running
''' winapp2ool.exe with a signed update
''' </summary>
Public Module updater

    ''' <summary>
    ''' The winapp2ool version in <c> version.txt </c> on the branch the Beta setting selects, or an
    ''' empty string until a check reads it. Once read, we don't read it again this session.
    ''' </summary>
    Public Property latestVersion As String = ""
    ''' <summary>The version of the current flavor's winapp2.ini on GitHub, or an empty string until a check reads it</summary>
    Public Property latestWa2Ver As String = ""
    ''' <summary>
    ''' The version of the winapp2.ini in the working folder, or a string starting <c> 000000 </c> when
    ''' there's no file, no version line, or no completed check
    ''' </summary>
    Public Property localWa2Ver As String = "000000"
    ''' <summary>
    ''' Indicates whether <c> version.txt </c> names a newer winapp2ool than the one running. We set it from
    ''' <c> version.txt </c> alone: <see cref="autoUpdate"/> checks the signature and the .NET Framework only
    ''' when it installs the update.
    ''' </summary>
    Public Property updateIsAvail As Boolean = False
    ''' <summary>Indicates whether the winapp2.ini on GitHub has a higher version number than the local one</summary>
    Public Property waUpdateIsAvail As Boolean = False
    ''' <summary>The file version of the running winapp2ool, or an empty string until the launcher or an update check reads it</summary>
    Public Property currentVersion As String = ""
    ''' <summary>
    ''' Indicates whether an update check has completed this session. A failed check leaves it
    ''' <c> False </c>, so a later call to <see cref="checkUpdates"/> can run the check again.
    ''' </summary>
    Public Property checkedForUpdates As Boolean = False

    ''' <summary>
    ''' Returns the version of a remote winapp2.ini, reading only its first line
    ''' </summary>
    '''
    ''' <param name="remotelink">
    ''' The URL of the winapp2.ini to check
    ''' </param>
    '''
    ''' <returns>
    ''' The version number, <br />
    ''' <c> "000000 (version not found)" </c> if the file has no version line, <br />
    ''' an empty string if the download fails
    ''' </returns>
    Public Function getRemoteVersion(remotelink As String) As String

        If remotelink Is Nothing Then argIsNull(NameOf(remotelink)) : Return ""

        Dim firstLine = getRemoteFirstLine(remotelink)
        If firstLine Is Nothing Then Return ""

        Return versionFromHeader(firstLine)

    End Function

    ''' <summary>
    ''' Returns the version number from a winapp2.ini header line such as <c> ; Version: 260930 </c>
    ''' </summary>
    '''
    ''' <param name="header">
    ''' The first line of a winapp2.ini
    ''' </param>
    '''
    ''' <returns>
    ''' The version number, or <c> "000000 (version not found)" </c> if the line doesn't carry one
    ''' </returns>
    Private Function versionFromHeader(header As String) As String

        Dim parts = header.Split(" "c)
        Return If(header.ToUpperInvariant.Contains("VERSION") AndAlso parts.Length > 2, parts(2), "000000 (version not found)")

    End Function

    ''' <summary>
    ''' Compares the winapp2ool and winapp2.ini versions on GitHub with the local ones, sets
    ''' <see cref="updateIsAvail"/> and <see cref="waUpdateIsAvail"/>, and puts a header naming any updates on
    ''' the next menu. Does nothing when <paramref name="cond"/> is <c> False </c>, when a check has already
    ''' completed, or when winapp2ool is offline.
    ''' <br /><br />
    '''
    ''' If either remote version can't be read, we show a failure header and retest the connection, which can
    ''' put winapp2ool into offline mode. A winapp2ool version that won't parse fails the check without the retest.
    ''' </summary>
    '''
    ''' <param name="cond">
    ''' Indicates whether to run the check <br /><br />
    ''' Optional, Default: <c> False </c>
    ''' </param>
    Public Sub checkUpdates(Optional cond As Boolean = False)
        If checkedForUpdates Or Not cond Then Return

        If isOffline Then
            gLog("Skipping the update check because winapp2ool is offline")
            Return
        End If

        gLog("Checking for updates")
        ' Query the latest winapp2ool.exe and winapp2.ini versions
        toolVersionCheck()
        ' Only the first line is read, so the check doesn't download the whole winapp2.ini
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

    ''' <summary>Reads <see cref="latestVersion"/> from <c> version.txt </c>, unless an earlier check already read it</summary>
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
    ''' Returns whether <paramref name="candidate"/> is a strictly newer winapp2ool version than <paramref name="running"/>
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
    ''' Returns whether <paramref name="dir"/> is the folder that holds the running winapp2ool executable
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

    ''' <summary>
    ''' Puts a red "update check failed" header on the next menu and resets <see cref="localWa2Ver"/>
    ''' to <c> 000000 </c>
    ''' </summary>
    '''
    ''' <param name="name">The word shown before "update check failed" in the header</param>
    '''
    ''' <param name="chkOnline">
    ''' Indicates whether to retest the connection, which can put winapp2ool into offline mode <br /><br />
    ''' Optional, Default: <c> False </c>
    ''' </param>
    Private Sub updateCheckFailed(name As String, Optional chkOnline As Boolean = False)
        setNextMenuHeaderText($"/!\ {name} update check failed. /!\", printColor:=ConsoleColor.Red)
        localWa2Ver = "000000"
        If chkOnline Then chkOfflineMode()
    End Sub

    ''' <summary>Returns the version number from the first line of a winapp2.ini on disk</summary>
    '''
    ''' <param name="path">
    ''' The path of the file to read <br /><br />
    ''' Optional, Default: <c> "" </c> <br />
    ''' which reads winapp2.ini in the working folder
    ''' </param>
    '''
    ''' <returns>
    ''' The version number, <br />
    ''' <c> "000000 (file not found)" </c> if the file doesn't exist, <br />
    ''' <c> "000000 (version not found)" </c> if its first line doesn't carry one
    ''' </returns>
    Private Function getVersionFromLocalFile(Optional path As String = "") As String
        ' The main menu's winapp2.ini update downloads into the working folder, so that's the copy we compare
        If path.Length = 0 Then path = Environment.CurrentDirectory & "\winapp2.ini"
        If Not File.Exists(path) Then Return "000000 (file not found)"
        Return versionFromHeader(getFileDataAtLineNum(path))
    End Function

    ''' <summary>Tests the connection to GitHub and sets <see cref="isOffline"/> from the result, so it can also bring winapp2ool back online</summary>
    Public Sub chkOfflineMode()
        gLog("Checking online status")
        isOffline = Not checkOnline()
    End Sub

    ''' <summary>
    ''' Replaces the running winapp2ool executable with the build GitHub publishes on the branch the Beta setting
    ''' selects, then relaunches it with the same arguments and exits. The running version is kept beside the
    ''' new one as <c> winapp2ool v&lt;version&gt;.exe.bak </c>. A build with no
    ''' <see cref="TrustedUpdateKeys"/> refuses before downloading anything.
    ''' <br /><br />
    '''
    ''' We download the exe and its <c> .sig </c> into memory, and install nothing unless the signature verifies
    ''' against <see cref="TrustedUpdateKeys"/>. The signature covers only the exe's bytes, so we then check, in
    ''' order, that this PC has the .NET Framework the new build targets (a target we don't recognize counts as
    ''' missing, and so does a build already inspected earlier in this process, so retrying the same update in
    ''' one session is refused), that the copy we stage next to the executable hashes the same as the verified bytes, that its
    ''' assembly name is winapp2ool, and that its file version is newer than ours, since an old build carries a
    ''' valid signature too.
    ''' <br /><br />
    '''
    ''' A refused check, a download error or a file error stops the update, discards the staged copy and reports
    ''' why. A missing <c> .sig </c> is reported as a failed download. The running executable
    ''' stays in place unless putting it back after a failed swap also fails. We don't offer to restart elevated
    ''' when the executable's folder refuses the write.
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
    ''' Downloads the latest winapp2ool executable and its signature, verifies the signature, checks the .NET
    ''' Framework it needs, then stages it, checks the staged copy, and swaps it in. Download and file errors
    ''' go to the caller.
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

        ' Installing a build this PC can't run would leave the user with an exe that won't start
        Dim targetFramework = targetFrameworkOf(newExe)

        If Not frameworkRequirementMet(targetFramework).GetValueOrDefault(False) Then

            updateFailed($"The new winapp2ool needs {describeFramework(targetFramework)}, which isn't installed on this PC. Install it from Microsoft, then update again")
            Return False

        End If

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
    ''' Checks that the staged executable has the same SHA-256 hash as the verified download, has the assembly
    ''' name <c> winapp2ool </c>, and has a file version newer than the running one
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
    ''' Moves the running executable aside as a backup, replacing any file already at
    ''' <paramref name="backupPath"/>, and puts the staged one in its place. If the second move fails, we
    ''' move the backup back when nothing has taken its place, and let the error through to the caller.
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
    ''' then exits with code 0 without waiting for it. The new process gets this process's rights, so no UAC
    ''' prompt appears. If it won't start, we tell the user and return instead of exiting.
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
    ''' Tells the user why an update did not happen and records it in the log. Sets exit code 1 when
    ''' <c> -s </c> is on the command line or a command line module has run this session (the flag stays
    ''' set if the menu opens afterward). A non-silent
    ''' <c> -autoupdate </c> run updates before any module starts, so this doesn't set its exit code.
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
    ''' Returns the arguments joined with spaces into a single command line string, each quoted with
    ''' <see cref="QuoteArgument"/>
    ''' </summary>
    '''
    ''' <param name="args">
    ''' The arguments to join
    ''' </param>
    Friend Function BuildArgumentString(args As IEnumerable(Of String)) As String

        Return String.Join(" ", args.Select(AddressOf QuoteArgument))

    End Function

    ''' <summary>
    ''' Deletes a file from the disk if it exists. An <see cref="IOException"/> goes to
    ''' <see cref="handleIOException"/>, which marks the run failed. Other errors, such as a refused
    ''' deletion, go to the caller.
    ''' </summary>
    '''
    ''' <param name="path">The path of the file to delete</param>
    Public Sub fDelete(path As String)
        Try
            If File.Exists(path) Then File.Delete(path)
        Catch ex As IOException
            gLog("Failed to delete file")
            handleIOException(ex)
        End Try
    End Sub
End Module