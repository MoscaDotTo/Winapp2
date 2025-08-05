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
''' commandLineHandler is a winapp2ool module which handles the command line arguments passed to 
''' winapp2ool by the user. This allows winapp2ool to be called from scripting environments and 
''' enables the use of winapp2ool without having to interact with the UI
''' </summary>
''' 
''' Docs last updated: 2025-08-05 | Code last updated: 2025-08-05
Public Module commandLineHandler

    ''' <summary>
    ''' The current list of the command line args (mutable)
    ''' </summary>
    ''' 
    ''' Docs last updated: 2023-07-20 | Code last updated: 2023-07-20
    Public Property cmdargs As List(Of String)

    ''' <summary>
    ''' Indicates whether Winapp2ool was launched from the command line with module arguments
    ''' When True, prevents automatic saving of settings to preserve user's saved configuration
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-08-05 | Code last updated: 2025-08-05
    Public Property IsCommandLineMode As Boolean = False

    ''' <summary>
    ''' Configuration for module file requirements and handlers
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-08-05 | Code last updated: 2025-08-05
    Private ReadOnly ModuleConfigs As Dictionary(Of String, ModuleConfig) = CreateModuleConfigs()

    ''' <summary>
    ''' Creates the module configuration dictionary with all aliases
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-08-05 | Code last updated: 2025-08-05
    Private Function CreateModuleConfigs() As Dictionary(Of String, ModuleConfig)

        Dim configs As New Dictionary(Of String, ModuleConfig)

        ' Helper method to add a module configuration
        Dim addModule = Sub(number As String,
                            name As String,
                            fileCount As Integer,
                            handler As Action)

                            Dim config As New ModuleConfig(fileCount, handler)
                            configs.Add(number, config)
                            configs.Add($"-{number}", config)
                            configs.Add(name, config)
                            configs.Add($"-{name}", config)

                        End Sub

        addModule("1", "debug", 3, AddressOf WinappDebug.HandleLintCmdLine)
        addModule("2", "trim", 3, AddressOf Trim.handleCmdLine)
        addModule("3", "transmute", 3, AddressOf Transmute.handleCmdLine)
        addModule("4", "diff", 3, AddressOf Diff.HandleCmdLine)
        addModule("5", "ccdebug", 3, AddressOf CCiniDebug.handleCmdlineArgs)
        addModule("6", "download", 3, AddressOf Downloader.handleCmdLine)
        addModule("7", "browserbuilder", 9, AddressOf BrowserBuilder.handleCmdLine)
        addModule("8", "flavorizer", 8, AddressOf Flavorizer.handleCmdLine)

        Return configs

    End Function

    ''' <summary>
    ''' Configuration data for a module's command line handling
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-08-05 | Code last updated: 2025-08-05
    Private Structure ModuleConfig

        Public ReadOnly FileCount As Integer
        Public ReadOnly Handler As Action

        Public Sub New(fileCount As Integer,
                       handler As Action)

            Me.FileCount = fileCount
            Me.Handler = handler

        End Sub

    End Structure

    ''' <summary>
    ''' Flips a boolean setting and removes its associated argument from the args list
    ''' </summary>
    ''' 
    ''' <param name="setting">
    ''' A boolean module setting whose state will be inverted 
    ''' </param>
    ''' 
    ''' <param name="arg">
    ''' A commandline argument targeting 
    ''' <c> <paramref name="setting"/> </c>
    ''' </param>
    ''' 
    ''' <param name="name">
    ''' A File name to be modified if the arg is found
    ''' <br /> Optional, default: <c> "" </c>
    ''' </param>
    ''' 
    ''' <param name="newname">
    ''' The new name with which <c> <paramref name="name"/> </c>
    ''' will be replaced if the arg is found 
    ''' <br /> Optional, default: <c> "" </c>
    ''' </param>
    ''' 
    ''' Docs last updated: 2023-07-20 | Code last updated: 2023-07-20
    Public Sub invertSettingAndRemoveArg(ByRef setting As Boolean,
                                               arg As String,
                                Optional ByRef name As String = "",
                                Optional ByRef newname As String = "")

        If Not cmdargs.Contains(arg) Then Return

        gLog($"Found argument: {arg}")
        setting = Not setting
        cmdargs.Remove(arg)
        name = newname

    End Sub

    ''' <summary>
    ''' Processes whether to download and which file to download
    ''' </summary>
    ''' 
    ''' <param name="download">
    ''' Indicates that winapp2.ini should be downloaded from GitHub 
    ''' </param>
    ''' 
    ''' Docs last updated: 2023-07-20 | Code last updated: 2023-07-20
    Public Sub handleDownloadBools(ByRef download As Boolean)

        invertSettingAndRemoveArg(download, "-d")

        Dim networkErr = "Winapp2ool is currently in offline mode, but you have issued commands that require a network connection. Please try again with a network connection."
        If download And isOffline Then printErrExit(networkErr)

    End Sub

    ''' <summary>
    ''' Initializes the processing of the commandline args and hands the 
    ''' remaining arguments off to the respective module's handler
    ''' </summary>
    ''' 
    ''' Docs last updated: 2025-08-05 | Code last updated: 2025-08-05
    Public Sub processCommandLineArgs()

        cmdargs = Environment.GetCommandLineArgs.ToList
        Dim argStr = String.Join(",", cmdargs)
        gLog($"Found commandline args: {argStr}")

        ' 0th index holds the executable name, we don't need it. 
        cmdargs.RemoveAt(0)

        checkUpdates(cmdargs.Contains("-autoupdate"))
        If updateIsAvail Then autoUpdate()

        ' Process global flags
        invertSettingAndRemoveArg(SuppressOutput, "-s")
        invertSettingAndRemoveArg(RemoteWinappIsNonCC, "-ncc")

        If cmdargs.Count = 0 Then Return

        Dim firstArg = cmdargs(0)
        cmdargs.RemoveAt(0)

        If ModuleConfigs.ContainsKey(firstArg) Then

            Dim config = ModuleConfigs(firstArg)
            validateArgs(config.FileCount)
            IsCommandLineMode = True
            gLog("Command line mode detected - settings saving disabled to preserve user configuration")
            config.Handler()

        Else

            gLog($"Invalid command line argument provided: {firstArg}")
            printErrExit($"Unknown module identifier: {firstArg}")

        End If

    End Sub

    ''' <summary>
    ''' Gets the directory and name info for the <c> iniFile </c> properties of a module 
    ''' </summary>
    ''' 
    ''' <param name="files">
    ''' The set of <c> iniFile </c> properties belonging to a particular module 
    ''' </param>
    ''' 
    ''' Docs last updated: 2025-08-05 | Code last updated: 2025-08-05
    Public Sub getFileAndDirParams(ByRef files() As iniFile)

        If files Is Nothing Then argIsNull(NameOf(files)) : Return

        For i = 0 To files.Length - 1

            getParams(i + 1, files(i))

        Next

    End Sub

    ''' <summary>
    ''' Renames or modifies the filename of an <c> iniFile </c> via the commandline 
    ''' </summary>
    ''' 
    ''' <param name="arg">
    ''' Commandline arg pointing to some particular file in a module 
    ''' <br /> eg. -2f or -1d 
    ''' </param>
    ''' 
    ''' <param name="givenFile">
    ''' An <c> iniFile </c> module property whose path will be modified 
    ''' </param>
    ''' 
    ''' <remarks> 
    ''' Supports appending child folders to the current directory 
    ''' <br /> eg. -2f "\folder1\folder2\file.ini"
    ''' </remarks>
    ''' 
    ''' Docs last updated: 2025-08-05 | Code last updated: 2025-08-05
    Private Sub getFileName(arg As String,
                      ByRef givenFile As iniFile)

        If cmdargs.Count < 2 Then Return

        Dim ind = cmdargs.IndexOf(arg)
        If ind = -1 OrElse ind >= cmdargs.Count - 1 Then Return

        Dim curArg = cmdargs(ind + 1)
        givenFile.Name = curArg

        If curArg.StartsWith("\", StringComparison.InvariantCulture) AndAlso Not curArg.LastIndexOf("\", StringComparison.InvariantCulture) = 0 Then

            Dim split = curArg.Split(CChar("\"))

            For i As Integer = 1 To split.Length - 2

                givenFile.Dir += $"\{split(i)}"

            Next

            givenFile.Name = split.Last

        End If

        cmdargs.RemoveAt(ind)
        cmdargs.RemoveAt(ind)

    End Sub

    ''' <summary>
    ''' Applies a new directory and name to an iniFile object
    ''' </summary>
    ''' 
    ''' <param name="flag">
    ''' The flag preceding the file/path parameter in the arg list
    ''' </param>
    ''' 
    ''' <param name="file">
    ''' An <c> iniFile </c> module property whose path will be modified 
    ''' </param>
    ''' 
    ''' Docs last updated: 2025-08-05 | Code last updated: 2025-08-05
    Private Sub getFileNameAndDir(flag As String,
                            ByRef file As iniFile)

        If cmdargs.Count < 2 Then Return

        Dim ind = cmdargs.IndexOf(flag)
        If ind = -1 OrElse ind >= cmdargs.Count - 1 Then Return

        getFileParams(cmdargs(ind + 1), file)
        cmdargs.RemoveAt(ind)
        cmdargs.RemoveAt(ind)

    End Sub

    ''' <summary>
    ''' Takes in a full form filepath with directory and assigns the directory 
    ''' and filename components to the given <c> <paramref name="file"/> </c>
    ''' </summary>
    ''' 
    ''' <param name="arg">
    ''' A file path argument, eg. the location of a specific .ini file on disk 
    ''' </param>
    ''' 
    ''' <param name="file">
    ''' An <c> iniFile </c> module property whose path will be modified 
    ''' </param>
    ''' 
    ''' Docs last updated: 2025-08-05 | Code last updated: 2025-08-05
    Private Sub getFileParams(ByRef arg As String,
                              ByRef file As iniFile)

        file.Dir = If(arg.StartsWith("\", StringComparison.InvariantCulture), Environment.CurrentDirectory, "")
        Dim splitArg As String() = arg.Split(CChar("\"))

        file.Name = splitArg.Last

        If splitArg.Length < 2 Then Return

        For i = 0 To splitArg.Length - 2

            file.Dir += splitArg(i) & "\"

        Next

        CleanFilePath(file)

    End Sub

    ''' <summary>
    ''' Processes numerically ordered directory (d) and file (f) commandline args on a per-file basis
    ''' </summary>
    ''' 
    ''' <param name="iniFilePropertyNumber">
    ''' The number (1-indexed) of the property associated with <c> <paramref name="someFile"/> </c>
    ''' </param>
    ''' 
    ''' <param name="someFile">
    ''' An <c> iniFile </c> module property whose path will be modified 
    ''' </param>
    ''' 
    ''' Docs last updated: 2025-08-05 | Code last updated: 2025-08-05
    Private Sub getParams(iniFilePropertyNumber As Integer,
                    ByRef someFile As iniFile)

        Dim argStr As String = $"-{iniFilePropertyNumber}"

        If cmdargs.Contains($"{argStr}d") Then getFileNameAndDir($"{argStr}d", someFile)

        If cmdargs.Contains($"{argStr}f") Then getFileName($"{argStr}f", someFile)

        CleanFilePath(someFile)

    End Sub

    ''' <summary>
    ''' Cleans up file paths by removing double, leading, and trailing slashes
    ''' </summary>
    ''' 
    ''' <param name="file">
    ''' An <c> iniFile </c> module property whose path will be modified 
    ''' </param>
    ''' 
    ''' Docs last updated: 2025-08-05 | Code last updated: 2025-08-05
    Private Sub CleanFilePath(ByRef file As iniFile)

        file.Dir = file.Dir.Replace("\\", "\")

        If file.Name.StartsWith("\", StringComparison.InvariantCulture) Then file.Name = file.Name.TrimStart(CChar("\"))

        If file.Dir.EndsWith("\", StringComparison.InvariantCulture) Then file.Dir = file.Dir.TrimEnd(CChar("\"))

    End Sub

    ''' <summary>
    ''' Generates the list of valid file arguments based on the maximum number of files
    ''' </summary>
    ''' 
    ''' <param name="maxFiles">
    ''' The maximum number of <c> iniFiles </c> for which arguments should be generated
    ''' </param>
    ''' 
    ''' <returns>
    ''' Set of valid commandline arguments for file and directory parameters
    ''' <br /> eg. -1d, -1f, -2d, -2f, ..., -maxFilesd, -maxFilesf
    ''' </returns>
    ''' 
    ''' Docs last updated: 2025-08-05 | Code last updated: 2025-08-05
    Private Function generateValidArgs(maxFiles As Integer) As String()

        Dim validArgs As New List(Of String)

        For i = 1 To maxFiles

            validArgs.Add($"-{i}d")
            validArgs.Add($"-{i}f")

        Next

        Return validArgs.ToArray()

    End Function

    ''' <summary>
    ''' Enforces that commandline args are properly formatted and paths exist
    ''' </summary>
    ''' 
    ''' <param name="maxFiles">
    ''' Maximum number of files to for which to validate the arguments 
    ''' </param>
    ''' 
    ''' Docs last updated: 2025-08-05 | Code last updated: 2025-08-05
    Private Sub validateArgs(maxFiles As Integer)

        Dim vArgs As String() = generateValidArgs(maxFiles)

        If cmdargs.Count <= 1 Then Return

        Dim i = 0
        While i < cmdargs.Count - 1

            If Not vArgs.Contains(cmdargs(i)) Then i += 1 : Continue While

            Dim pathArg = cmdargs(i + 1)

            ' Skip the next argument since we've already processed it
            i += 2

        End While

    End Sub

    ''' <summary>
    ''' Prints an error to the user and exits the application after they have pressed a key
    ''' </summary>
    ''' 
    ''' <param name="errTxt">
    ''' The text to be printed to the user
    ''' </param>
    ''' 
    ''' Docs last updated: 2025-08-05 | Code last updated: 2025-08-05
    Private Sub printErrExit(errTxt As String)

        Console.WriteLine($"{errTxt} Press any key to exit.")
        Console.ReadKey()
        Environment.Exit(0)

    End Sub

End Module
