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
''' CLI argument specification for a single module. Declares which file slots,
''' boolean flags, and download support a module accepts, then parses those declarations
''' against <c> cmdargs </c> when <c> Parse </c> is called.
''' </summary>
'''
''' <remarks>
''' Modules create and invoke a <c> CliArgSpec </c> at the top of their
''' <c> handleCmdLine </c> implementation, after settings have been initialized:
''' 
''' <code>
''' New CliArgSpec("modulename") _
'''     .WithFile(1, Me.InputFile, Alias (optional)) _
'''     .WithFlag("-flag", Sub() Setting = Not Setting) _
'''     .Parse()
''' </code>
''' 
''' </remarks>
Public Class CliArgSpec

    ''' <summary>
    ''' Binds a numbered file slot to an <c> iniFileChooser </c> instance, with optional aliases.
    ''' </summary>
    Private Structure FileBinding

        Public ReadOnly Slot As Integer
        Public ReadOnly File As iniFileChooser
        Public ReadOnly Aliases As String()

        Public Sub New(slot As Integer, file As iniFileChooser, aliases As String())

            Me.Slot = slot
            Me.File = file
            Me.Aliases = aliases

        End Sub

    End Structure

    Private Structure FlagBinding

        Public ReadOnly Flag As String
        Public ReadOnly Toggle As Action

        Public Sub New(flag As String, toggle As Action)

            Me.Flag = flag
            Me.Toggle = toggle

        End Sub

    End Structure

    ''' <summary>
    ''' Binds a flag to a file rename action, allowing a single flag
    ''' to set a file slot to a well-known source file.
    ''' </summary>
    Private Structure FileAliasBinding

        Public ReadOnly Flag As String
        Public ReadOnly File As iniFileChooser
        Public ReadOnly NewName As String

        Public Sub New(flag As String, file As iniFileChooser, newName As String)

            Me.Flag = flag
            Me.File = file
            Me.NewName = newName

        End Sub

    End Structure

    Private ReadOnly _moduleName As String
    Private ReadOnly _fileBindings As New List(Of FileBinding)
    Private ReadOnly _flagBindings As New List(Of FlagBinding)
    Private ReadOnly _fileAliasBindings As New List(Of FileAliasBinding)
    Private _downloadToggle As Action = Nothing
    Private _downloadGetState As Func(Of Boolean) = Nothing
    Private _positionalHandler As Action(Of String) = Nothing

    ''' <summary>
    ''' Creates a <c> CliArgSpec </c> for the named module
    ''' </summary>
    '''
    ''' <param name="moduleName">
    ''' The module name used in diagnostic log messages
    ''' </param>
    Public Sub New(moduleName As String)

        _moduleName = moduleName

    End Sub

    ''' <summary>
    ''' Binds a numbered file slot to an <c> iniFileChooser </c> instance.
    ''' The <c> -Nf </c> and <c> -Nd </c> command-line arguments (where N equals
    ''' <paramref name="slot"/>) will set the file's path when <c> Parse </c> is called.
    ''' Optional <paramref name="aliases"/> provide human-readable alternatives:
    ''' an alias of <c> "old" </c> also accepts <c> -oldf </c> and <c> -oldd </c>.
    ''' </summary>
    '''
    ''' <param name="slot">
    ''' The 1-indexed slot number corresponding to <c> -Nf </c> / <c> -Nd </c> args
    ''' </param>
    '''
    ''' <param name="file">
    ''' The <c> iniFileChooser </c> instance whose path will be modified
    ''' </param>
    '''
    ''' <param name="aliases">
    ''' Optional human-readable names for this slot. Each alias <c> x </c> accepts
    ''' <c> -xf </c> (filename) and <c> -xd </c> (directory) in addition to the
    ''' numbered forms.
    ''' </param>
    Public Function WithFile(slot As Integer,
                             file As iniFileChooser,
                  ParamArray aliases As String()) As CliArgSpec

        _fileBindings.Add(New FileBinding(slot, file, aliases))
        Return Me

    End Function

    ''' <summary>
    ''' Binds a named flag to a toggle action.
    ''' When <paramref name="flag"/> is present in <c> cmdargs </c>,
    ''' <paramref name="toggle"/> is invoked and the flag is removed.
    ''' </summary>
    '''
    ''' <param name="flag">
    ''' The command-line flag to match, including its leading dash (e.g. <c> "-savelog" </c>)
    ''' </param>
    '''
    ''' <param name="toggle">
    ''' An action that toggles the associated module setting
    ''' </param>
    Public Function WithFlag(flag As String,
                             toggle As Action) As CliArgSpec

        _flagBindings.Add(New FlagBinding(flag, toggle))
        Return Me

    End Function

    ''' <summary>
    ''' Binds a named flag to a file rename action.
    ''' When <paramref name="flag"/> is present in <c> cmdargs </c>,
    ''' <paramref name="file"/>.Name is set to <paramref name="newName"/>.
    ''' Used for preset filename shortcuts where a short flag selects a well-known source file.
    ''' </summary>
    '''
    ''' <param name="flag">
    ''' The command-line flag to match, including its leading dash (e.g. <c> "-r" </c>)
    ''' </param>
    '''
    ''' <param name="file">
    ''' The <c> iniFileChooser </c> instance whose <c> Name </c> will be replaced
    ''' </param>
    '''
    ''' <param name="newName">
    ''' The filename to assign when <paramref name="flag"/> is found
    ''' </param>
    Public Function WithFileAlias(flag As String,
                                  file As iniFileChooser,
                                  newName As String) As CliArgSpec

        _fileAliasBindings.Add(New FileAliasBinding(flag, file, newName))
        Return Me

    End Function

    ''' <summary>
    ''' Enables the <c> -d </c> download flag for this module.
    ''' When <c> -d </c> is present in <c> cmdargs </c>, <paramref name="toggle"/> is invoked.
    ''' If <paramref name="getState"/> is provided, an offline error is raised whenever
    ''' the resolved download state is <c> True </c> and the application is offline.
    ''' </summary>
    '''
    ''' <param name="toggle">
    ''' An action that toggles the module's download setting
    ''' </param>
    '''
    ''' <param name="getState">
    ''' A function that returns the download setting's current value, used to detect
    ''' offline violations after the toggle has been applied. <br /><br />
    ''' Optional, Default: <c> Nothing </c> <br />
    ''' omit if no offline guard is needed.
    ''' </param>
    Public Function WithDownload(toggle As Action,
                        Optional getState As Func(Of Boolean) = Nothing) As CliArgSpec

        _downloadToggle = toggle
        _downloadGetState = getState
        Return Me

    End Function

    ''' <summary>
    ''' Registers a handler for the first positional (non-flag) argument.
    ''' <c> ParsePositionalArg </c> runs before file binding, so the handler
    ''' can set filenames that file-slot args then override.
    ''' </summary>
    '''
    ''' <param name="handler">
    ''' An action invoked with the first non-flag token found in <c> cmdargs </c>
    ''' </param>
    Public Function WithPositional(handler As Action(Of String)) As CliArgSpec

        _positionalHandler = handler
        Return Me

    End Function

    ''' <summary>
    ''' Parses <c> cmdargs </c> against this spec. Processes positional args, file slots,
    ''' boolean flags, file aliases, and the download flag in order, then emits a
    ''' diagnostic warning for any unrecognized <c> - </c>-prefixed args that remain.
    ''' </summary>
    Public Sub Parse()

        ParsePositionalArg()
        ParseFileArgs()
        ParseFlagArgs()
        ParseFileAliasArgs()
        ParseDownloadArg()
        WarnUnknownArgs()

    End Sub

    ''' <summary>
    ''' Invokes the positional handler with the first non-flag token in <c> cmdargs </c>,
    ''' then removes that token. Does nothing if no positional handler was registered.
    ''' </summary>
    Private Sub ParsePositionalArg()

        If _positionalHandler Is Nothing Then Return

        Dim positional = cmdargs.FirstOrDefault(Function(a) Not a.StartsWith("-", StringComparison.InvariantCulture))
        If positional Is Nothing Then Return

        _positionalHandler(positional)
        cmdargs.Remove(positional)

    End Sub

    ''' <summary>
    ''' For each registered file binding, applies any matching <c> -Nd </c> / <c> -Nf </c>
    ''' args and any alias-based args to the bound <c> iniFileChooser </c>,
    ''' then cleans the resulting path.
    ''' </summary>
    Private Sub ParseFileArgs()

        For Each binding In _fileBindings

            ApplySlotArgs($"-{binding.Slot}", binding.File)

            For Each slotAlias In binding.Aliases : ApplySlotArgs($"-{slotAlias}", binding.File) : Next

            CleanPath(binding.File)

        Next

    End Sub

    ''' <summary>
    ''' Applies the <c> -Xd </c> and <c> -Xf </c> variants of a given prefix to a file,
    ''' where X is the prefix string (e.g. <c> "-1" </c> or <c> "-old" </c>).
    ''' </summary>
    '''
    ''' <param name="prefix">
    ''' The base flag prefix, e.g. <c> "-1" </c> or <c> "-old" </c>
    ''' </param>
    '''
    ''' <param name="file">
    ''' The <c> iniFileChooser </c> whose path will be modified
    ''' </param>
    Private Sub ApplySlotArgs(prefix As String, file As iniFileChooser)

        If cmdargs.Contains($"{prefix}d") Then ApplyDirArg($"{prefix}d", file)
        If cmdargs.Contains($"{prefix}f") Then ApplyFileArg($"{prefix}f", file)

    End Sub

    ''' <summary>
    ''' For each registered flag binding, invokes the toggle action and removes the flag
    ''' if it is present in <c> cmdargs </c>.
    ''' </summary>
    Private Sub ParseFlagArgs()

        For Each binding In _flagBindings

            If Not cmdargs.Contains(binding.Flag) Then Continue For

            gLog($"Found argument {binding.Flag}")
            binding.Toggle()
            cmdargs.Remove(binding.Flag)

        Next

    End Sub

    ''' <summary>
    ''' For each registered file alias binding, sets the file's name and removes the flag
    ''' if it is present in <c> cmdargs </c>.
    ''' </summary>
    Private Sub ParseFileAliasArgs()

        For Each binding In _fileAliasBindings

            If Not cmdargs.Contains(binding.Flag) Then Continue For

            gLog($"Found argument {binding.Flag}")
            binding.File.Name = binding.NewName
            cmdargs.Remove(binding.Flag)

        Next

    End Sub

    ''' <summary>
    ''' Processes the <c> -d </c> download flag. If the flag is present, invokes the
    ''' registered toggle. If a state reader was provided and the resolved download state
    ''' is <c> True </c> while the application is offline, exits with an error message.
    ''' </summary>
    Private Sub ParseDownloadArg()

        If _downloadToggle Is Nothing Then Return

        If cmdargs.Contains("-d") Then

            gLog("Found download argument -d")
            _downloadToggle()
            cmdargs.Remove("-d")

        End If

        If _downloadGetState Is Nothing Then Return

        If _downloadGetState() AndAlso isOffline AndAlso Not SuppressOutput Then

            Console.WriteLine("Winapp2ool Is currently in offline mode, but you have issued commands that require a network connection. Please try again with a network connection. Press any key to exit.")
            Console.ReadKey()
            Environment.Exit(0)

        End If

    End Sub

    ''' <summary>
    ''' Logs a diagnostic warning for each remaining <c> - </c>-prefixed token in
    ''' <c> cmdargs </c> that was not consumed during parsing.
    ''' </summary>
    Private Sub WarnUnknownArgs()

        For Each arg In cmdargs

            If Not arg.StartsWith("-", StringComparison.InvariantCulture) Then Continue For

            gLog($"Warning unrecognized argument '{arg}' for module '{_moduleName}'")

            Next

    End Sub

    ''' <summary>
    ''' Applies a full directory path argument (<c> -Nd </c>) to an <c> iniFileChooser </c>.
    ''' Splits the path into directory and filename components. Removes the flag and its
    ''' value from <c> cmdargs </c>.
    ''' </summary>
    '''
    ''' <param name="flag">
    ''' The directory flag, e.g. <c> "-1d" </c>
    ''' </param>
    '''
    ''' <param name="file">
    ''' The <c> iniFileChooser </c> whose <c> Dir </c> and <c> Name </c> will be modified
    ''' </param>
    Private Sub ApplyDirArg(flag As String,
                            file As iniFileChooser)

        If cmdargs.Count < 2 Then Return

        Dim ind = cmdargs.IndexOf(flag)
        If ind = -1 OrElse ind >= cmdargs.Count - 1 Then Return

        Dim arg = cmdargs(ind + 1)
        file.Dir = If(arg.StartsWith("\", StringComparison.InvariantCulture), Environment.CurrentDirectory, "")

        Dim splitArg As String() = arg.Split(CChar("\"))

        If splitArg.Last.Contains(".") Then file.Name = splitArg.Last

        If splitArg.Length >= 2 Then

            Dim terminus = If(splitArg.Last.Contains("."), splitArg.Length - 2, splitArg.Length - 1)

            For i = 0 To terminus : file.Dir += splitArg(i) & "\" : Next

        End If

        cmdargs.RemoveAt(ind)
        cmdargs.RemoveAt(ind)

    End Sub

    ''' <summary>
    ''' Applies a filename argument (<c> -Nf </c>) to an <c> iniFileChooser </c>.
    ''' Supports leading relative subdirectory paths (e.g. <c> "\subfolder\file.ini" </c>).
    ''' Removes the flag and its value from <c> cmdargs </c>.
    ''' </summary>
    '''
    ''' <param name="flag">
    ''' The file flag, e.g. <c> "-1f" </c>
    ''' </param>
    '''
    ''' <param name="file">
    ''' The <c> iniFileChooser </c> whose <c> Name </c> (and optionally <c> Dir </c>) will be modified
    ''' </param>
    Private Sub ApplyFileArg(flag As String,
                             file As iniFileChooser)

        If cmdargs.Count < 2 Then Return

        Dim ind = cmdargs.IndexOf(flag)
        If ind = -1 OrElse ind >= cmdargs.Count - 1 Then Return

        Dim curArg = cmdargs(ind + 1)
        file.Name = curArg

        If curArg.StartsWith("\", StringComparison.InvariantCulture) AndAlso
           Not curArg.LastIndexOf("\", StringComparison.InvariantCulture) = 0 Then

            Dim split = curArg.Split(CChar("\"))

            For i As Integer = 1 To split.Length - 2 : file.Dir += $"\{split(i)}" : Next

            file.Name = split.Last

        End If

        cmdargs.RemoveAt(ind)
        cmdargs.RemoveAt(ind)

    End Sub

    ''' <summary>
    ''' Removes double slashes and trims leading/trailing slashes from an
    ''' <c> iniFileChooser </c>'s path components.
    ''' </summary>
    '''
    ''' <param name="file">
    ''' The <c> iniFileChooser </c> whose path will be cleaned
    ''' </param>
    Private Sub CleanPath(file As iniFileChooser)

        file.Dir = file.Dir.Replace("\\", "\")

        If file.Name.StartsWith("\", StringComparison.InvariantCulture) Then file.Name = file.Name.TrimStart(CChar("\"))

        If file.Dir.EndsWith("\", StringComparison.InvariantCulture) Then file.Dir = file.Dir.TrimEnd(CChar("\"))

    End Sub

End Class
