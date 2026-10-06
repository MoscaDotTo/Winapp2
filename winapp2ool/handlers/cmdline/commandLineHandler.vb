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
Imports System.Text

''' <summary>
''' commandLineHandler is a winapp2ool module which handles the command line arguments passed to 
''' winapp2ool by the user. This allows winapp2ool to be called from scripting environments and 
''' enables the use of winapp2ool without having to interact with the UI
''' </summary>
Public Module commandLineHandler

    ''' <summary>
    ''' The command line args still waiting to be handled. Every handler removes the args it
    ''' consumes, so the list shrinks as parsing goes on.
    ''' </summary>
    Public Property cmdargs As List(Of String)

    ''' <summary>
    ''' Indicates whether Winapp2ool was launched from the command line with a module argument.
    ''' When <c> True </c>, <see cref="SaveSettings"/> refuses to write, so a scripted run never
    ''' changes the user's saved configuration.
    ''' </summary>
    Public Property IsCommandLineMode As Boolean = False

    ''' <summary>
    ''' Each module's file slot count and command line handler, keyed by every form of its
    ''' identifier. The lookup is case-sensitive.
    ''' </summary>
    Private ReadOnly ModuleConfigs As Dictionary(Of String, ModuleConfig) = CreateModuleConfigs()

    ''' <summary>
    ''' Builds the module configuration dictionary. Each module goes in under four keys: its
    ''' number and its name, each with and without a leading dash.
    ''' </summary>
    Private Function CreateModuleConfigs() As Dictionary(Of String, ModuleConfig)

        Dim configs As New Dictionary(Of String, ModuleConfig)

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
        addModule("2", "trim", 4, AddressOf Trim.handleCmdLine)
        addModule("3", "transmute", 3, AddressOf Transmute.handleCmdLine)
        addModule("4", "diff", 3, AddressOf Diff.HandleCmdLine)
        addModule("5", "ccdebug", 3, AddressOf CCiniDebug.handleCmdlineArgs)
        addModule("6", "browserbuilder", 2, AddressOf BrowserBuilder.handleCmdLine)
        addModule("7", "combine", 3, AddressOf Combine.handleCmdLine)
        addModule("8", "download", 3, AddressOf Downloader.handleCmdLine)
        addModule("9", "flavorize", 9, AddressOf Flavorizer.handleCmdLine)
        addModule("10", "uwpbuilder", 3, AddressOf UWPBuilder.handleCmdLine)
        addModule("11", "entrybuilder", 3, AddressOf EntryBuilder.handleCmdLine)
        addModule("12", "cc7patcher", 3, AddressOf CC7Patcher.handleCmdLine)

        Return configs

    End Function

    ''' <summary>
    ''' Configuration data for a module's command line handling
    ''' </summary>
    Private Structure ModuleConfig

        Public ReadOnly FileCount As Integer
        Public ReadOnly Handler As Action

        ''' <summary>Creates a new <c> ModuleConfig </c> from its two fields</summary>
        '''
        ''' <param name="fileCount">
        ''' The number of file slots whose <c> -Nd </c> / <c> -Nf </c> args
        ''' <see cref="validateArgs"/> checks. It only decides which args get a warning, and a
        ''' module can bind slots past it.
        ''' </param>
        '''
        ''' <param name="handler">The module's command line handler</param>
        Public Sub New(fileCount As Integer,
                       handler As Action)

            Me.FileCount = fileCount
            Me.Handler = handler

        End Sub

    End Structure

    ''' <summary>
    ''' Flips a boolean setting and removes its argument from <see cref="cmdargs"/>, if the
    ''' argument is there. Only the first copy of the argument is removed.
    ''' </summary>
    '''
    ''' <param name="setting">
    ''' A boolean module setting whose state will be inverted
    ''' </param>
    '''
    ''' <param name="arg">
    ''' A commandline argument targeting <paramref name="setting"/> (case-sensitive)
    ''' </param>
    Public Sub invertSettingAndRemoveArg(ByRef setting As Boolean,
                                               arg As String)

        If Not cmdargs.Contains(arg) Then Return

        gLog($"Found argument: {arg}")
        setting = Not setting
        cmdargs.Remove(arg)

    End Sub

    ''' <summary>
    ''' Reads the command line into <see cref="cmdargs"/>, strips <c> -offline </c> (the launcher
    ''' has already acted on it), checks for and installs an update if <c> -autoupdate </c> is
    ''' present, handles <c> -s </c>, <c> -writelog </c> and the flavor flags, then hands the
    ''' rest to the module the first remaining arg names. <c> -autoupdate </c> stays in
    ''' <see cref="cmdargs"/>.
    ''' Returns without doing anything more when no args remain, so the interactive menu
    ''' opens (under <c> -s </c> the launcher exits instead). An unknown module identifier logs an error and exits with code 1.
    ''' <br /><br />
    ''' Dispatching to a module sets <see cref="IsCommandLineMode"/>, which stops settings
    ''' from being saved for the rest of the run.
    ''' </summary>
    Public Sub processCommandLineArgs()

        cmdargs = ReconstructArgs(Environment.GetCommandLineArgs())
        Dim argStr = String.Join(",", cmdargs)
        gLog($"Found commandline args: {argStr}")


        cmdargs.RemoveAll(Function(a) a.Equals("-offline", StringComparison.OrdinalIgnoreCase))

        checkUpdates(cmdargs.Contains("-autoupdate"))
        If updateIsAvail Then autoUpdate()

        invertSettingAndRemoveArg(SuppressOutput, "-s")
        invertSettingAndRemoveArg(SaveGlobalLogOnExit, "-writelog")
        processFlavorArgs()

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

            gLog($"Invalid module argument provided: {firstArg}")
            printErrExit($"Unknown module identifier: {firstArg}")

        End If

    End Sub

    ''' <summary>
    ''' Generates the list of valid file arguments based on the maximum number of files
    ''' </summary>
    ''' 
    ''' <param name="maxFiles">
    ''' The number of file slots for which arguments should be generated
    ''' </param>
    ''' 
    ''' <returns>
    ''' Set of valid commandline arguments for file and directory parameters
    ''' <br /> eg. -1d, -1f, -2d, -2f, ..., -maxFilesd, -maxFilesf
    ''' </returns>
    Private Function generateValidArgs(maxFiles As Integer) As String()

        Dim validArgs As New List(Of String)

        For i = 1 To maxFiles

            validArgs.Add($"-{i}d")
            validArgs.Add($"-{i}f")

        Next

        Return validArgs.ToArray()

    End Function

    ''' <summary>
    ''' Logs a warning for each <c> -Nd </c> whose directory doesn't exist and each
    ''' <c> -Nf </c> whose value names a parent directory that doesn't exist. It only warns,
    ''' and never stops the run or changes <see cref="cmdargs"/>. We check each path as given,
    ''' so a value with a leading <c> \ </c> is checked against the root of the current drive,
    ''' not the folder the module will resolve it against.
    ''' </summary>
    '''
    ''' <param name="maxFiles">
    ''' The number of file slots whose arguments we check
    ''' </param>
    Private Sub validateArgs(maxFiles As Integer)

        Dim vArgs As String() = generateValidArgs(maxFiles)

        If cmdargs.Count <= 1 Then Return

        Dim i = 0
        While i < cmdargs.Count - 1

            If Not vArgs.Contains(cmdargs(i)) Then i += 1 : Continue While

            Dim pathArg = cmdargs(i + 1)

            ' Validate that the path exists (for directory arguments) or parent directory exists (for file arguments)
            If cmdargs(i).EndsWith("d") Then

                If Not String.IsNullOrEmpty(pathArg) AndAlso Not System.IO.Directory.Exists(pathArg) Then

                    gLog($"Warning: Directory does not exist: {pathArg}. If you are downloading, this path will be created.")

                End If

            ElseIf cmdargs(i).EndsWith("f") Then

                If pathArg.Contains("\") Then

                    Dim parentDir = System.IO.Path.GetDirectoryName(pathArg)

                    gLog($"Warning: Parent directory does not exist: {parentDir}", Not String.IsNullOrEmpty(parentDir) AndAlso Not System.IO.Directory.Exists(parentDir))

                End If

            End If

            ' Skip the next argument since we've already processed it
            i += 2

        End While

    End Sub

    ''' <summary>
    ''' Logs an error, prints it and waits for a key unless output is suppressed, then exits
    ''' with code 1. We save the global log first when output is suppressed or
    ''' <see cref="SaveGlobalLogOnExit"/> is set.
    ''' </summary>
    ''' 
    ''' <param name="errTxt">
    ''' The text to be printed to the user
    ''' </param>
    Private Sub printErrExit(errTxt As String)

        gLog(errTxt)

        If Not SuppressOutput Then

            Console.WriteLine($"{errTxt} Press any key to exit.")
            crk()

        End If

        saveGlobalLog(SuppressOutput OrElse SaveGlobalLogOnExit)
        Environment.Exit(1)

    End Sub

    ''' <summary>
    ''' Sets <see cref="CurrentWinappFlavor"/> to <paramref name="flavorValue"/> and removes
    ''' <paramref name="flavorArg"/> from <see cref="cmdargs"/>, if the argument is there
    ''' </summary>
    ''' 
    ''' <param name="flavorArg">
    ''' The command line argument for flavor selection
    ''' </param>
    ''' 
    ''' <param name="flavorValue">
    ''' The flavor to set
    ''' </param>
    Private Sub handleFlavorArg(flavorArg As String, flavorValue As WinappFlavor)

        If Not cmdargs.Contains(flavorArg) Then Return

        gLog($"Found flavor argument: {flavorArg}")

        CurrentWinappFlavor = flavorValue
        cmdargs.Remove(flavorArg)

        gLog($"Flavor set to: {flavorValue}")

    End Sub

    ''' <summary>
    ''' Sets the winapp2 flavor from the flavor flags in <see cref="cmdargs"/>. When more than
    ''' one flavor flag is given, the one checked last here wins, whatever their order on the
    ''' command line.
    ''' </summary>
    Public Sub processFlavorArgs()

        handleFlavorArg("-ccleaner", WinappFlavor.CCleaner)
        handleFlavorArg("-cc", WinappFlavor.CCleaner)
        handleFlavorArg("-bleachbit", WinappFlavor.BleachBit)
        handleFlavorArg("-bb", WinappFlavor.BleachBit)
        handleFlavorArg("-systemninja", WinappFlavor.SystemNinja)
        handleFlavorArg("-sn", WinappFlavor.SystemNinja)
        handleFlavorArg("-tron", WinappFlavor.Tron)
        handleFlavorArg("-base", WinappFlavor.NonCCleaner)
        handleFlavorArg("-ncc", WinappFlavor.NonCCleaner)
        handleFlavorArg("-ccleaner7", WinappFlavor.CCleaner7)
        handleFlavorArg("-cc7", WinappFlavor.CCleaner7)
        handleFlavorArg("-fluentcleaner", WinappFlavor.FluentCleaner)
        handleFlavorArg("-fc", WinappFlavor.FluentCleaner)

    End Sub

    ''' <summary>
    ''' Rejoins args that still carry quote marks into one arg per quoted run, joined with
    ''' single spaces, and strips the quotes
    ''' </summary>
    '''
    ''' <param name="args">
    ''' The command line as <c> Environment.GetCommandLineArgs </c> returns it. We skip the
    ''' first element, which is the program itself.
    ''' </param>
    '''
    ''' <returns>
    ''' The args after the program name, with quoted runs rejoined
    ''' </returns>
    Private Function ReconstructArgs(args As String()) As List(Of String)

        Dim reconstructed As New List(Of String)
        Dim i As Integer = 1

        While i < args.Length

            Dim currentArg = args(i)

            If currentArg.StartsWith("""", StringComparison.InvariantCulture) Then

                If currentArg.EndsWith("""", StringComparison.InvariantCulture) AndAlso currentArg.Length > 1 Then

                    reconstructed.Add(currentArg.Trim(CChar("""")))

                Else

                    Dim quotedArg As New StringBuilder(currentArg.Substring(1))
                    i += 1

                    While i < args.Length

                        If args(i).EndsWith("""", StringComparison.InvariantCulture) Then

                            quotedArg.Append(" " & args(i).Substring(0, args(i).Length - 1))
                            Exit While

                        Else

                            quotedArg.Append(" " & args(i))

                        End If

                        i += 1

                    End While

                    reconstructed.Add(quotedArg.ToString())

                End If

            Else

                reconstructed.Add(currentArg)

            End If

            i += 1

        End While

        Return reconstructed

    End Function

End Module
