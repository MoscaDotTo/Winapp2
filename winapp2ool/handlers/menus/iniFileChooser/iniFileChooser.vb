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

''' <summary>
''' An interactive file chooser for <c> iniFile </c> objects, using a single unified menu
''' </summary>
Public Class iniFileChooser

    ''' <summary>
    ''' The directory of the chosen file. <see cref="Path"/> joins it to <see cref="Name"/> with a backslash.
    ''' </summary>
    Public Property Dir As String

    ''' <summary>
    ''' The current filename
    ''' </summary>
    Public Property Name As String

    ''' <summary>
    ''' The starting filename, offered as a numbered option and restored by <see cref="ResetParams"/>
    ''' </summary>
    Public ReadOnly Property InitName As String

    ''' <summary>
    ''' The starting directory, saved at construction time for use by <see cref="ResetParams"/>
    ''' </summary>
    Public ReadOnly Property InitDir As String

    ''' <summary>
    ''' An optional alternate default filename offered as a numbered option
    ''' </summary>
    Public ReadOnly Property SecondName As String

    ''' <summary>
    ''' Indicates whether picking a filename that doesn't exist keeps the chooser open. When
    ''' <c> False </c>, any name closes it, which suits output files.
    ''' </summary>
    Public ReadOnly Property MustExist As Boolean

    Private _tmpRename As String = ""

    ''' <summary>
    ''' Creates a new <c> iniFileChooser </c> with the given starting directory and filename
    ''' </summary>
    '''
    ''' <param name="dir">
    ''' Initial directory
    ''' </param>
    '''
    ''' <param name="name">
    ''' Initial filename
    ''' </param>
    '''
    ''' <param name="secondName">
    ''' Alternate default filename offered as a quick-select option <br /><br />
    ''' Optional, Default: <c> "" </c> <br />
    ''' no alternate is offered
    ''' </param>
    '''
    ''' <param name="mustExist">
    ''' Indicates whether picking a filename that doesn't exist keeps the chooser open <br /><br />
    ''' Optional, Default: <c> True </c>
    ''' </param>
    Public Sub New(dir As String, name As String,
                   Optional secondName As String = "",
                   Optional mustExist As Boolean = True)

        Me.Dir = dir
        Me.Name = name
        Me.InitDir = dir
        Me.InitName = name
        Me.SecondName = secondName
        Me.MustExist = mustExist

    End Sub

    ''' <summary>
    ''' Returns the full path of the current file
    ''' </summary>
    Public Function Path() As String
        Return $"{Dir}\{Name}"
    End Function

    ''' <summary>
    ''' Returns whether the current file (or just directory, when
    ''' <paramref name="checkPath"/> is <c> False </c>) exists on disk
    ''' </summary>
    '''
    ''' <param name="checkPath">
    ''' Indicates whether to check the full file path rather than only the directory <br /><br />
    ''' Optional, Default: <c> True </c>
    ''' </param>
    Public Function Exists(Optional checkPath As Boolean = True) As Boolean

        Return If(checkPath, File.Exists(Path()), Directory.Exists(Dir))

    End Function

    ''' <summary>
    ''' Returns the chosen file as a parsed <c> iniFile </c>. If the file doesn't exist, we open the
    ''' chooser menu once, and return <c> Nothing </c> if the user leaves it without an existing
    ''' file. In silent mode that menu can't run, so a missing file ends the process instead.
    ''' <br /> When the caller identifies which setting this chooser backs, a file picked through the
    ''' prompt is persisted to the settings file so the choice survives the session
    ''' </summary>
    '''
    ''' <param name="settingsChangedSetting">
    ''' The owning module's settings-changed flag. Set to <c> True </c> if we persist a newly
    ''' chosen file <br /><br />
    ''' Optional, Default: <c> False </c>
    ''' </param>
    '''
    ''' <param name="callingModule">
    ''' The name of the module owning this chooser as it appears in the settings file <br /><br />
    ''' Optional, Default: <c> "" </c>
    ''' </param>
    '''
    ''' <param name="settingName">
    ''' The name of this chooser as it appears in the codebase <br /><br />
    ''' Optional, Default: <c> "" </c>
    ''' </param>
    '''
    ''' <param name="settingChangedName">
    ''' The name of <paramref name="settingsChangedSetting"/> as it appears in the codebase <br /><br />
    ''' Optional, Default: <c> "" </c> <br />
    ''' we persist a new choice only when all three names are given
    ''' </param>
    '''
    ''' <returns>
    ''' The parsed file, or <c> Nothing </c> if the user left the chooser without picking a file
    ''' that exists
    ''' </returns>
    Public Function Load(Optional ByRef settingsChangedSetting As Boolean = False,
                         Optional callingModule As String = "",
                         Optional settingName As String = "",
                         Optional settingChangedName As String = "") As iniFile

        Dim curName = Name
        Dim curDir = Dir

        While Not Exists()

            initModule("File Chooser", AddressOf PrintMenu, AddressOf HandleInput)

            If Not Exists() Then Return Nothing

        End While

        Dim fileChanged = Not Name = curName OrElse Not Dir = curDir
        Dim isSettingsBacked = callingModule.Length > 0 AndAlso settingName.Length > 0 AndAlso settingChangedName.Length > 0

        If fileChanged AndAlso isSettingsBacked Then

            gLog($"Saving {settingName}'s newly chosen parameters to disk")
            saveChooserParams(Me, settingsChangedSetting, callingModule, settingName, settingChangedName)

        End If

        Return iniFile.FromFile(Path())

    End Function

    ''' <summary>
    ''' Returns the file at the current path, parsed by <see cref="iniFile.FromFile"/>, without
    ''' showing the chooser menu. Use this when the file is known to exist. For interactive
    ''' validation, use <see cref="Load"/> instead.
    ''' </summary>
    '''
    ''' <returns>
    ''' The parsed file, or one with no sections if the file doesn't exist. A missing directory
    ''' throws instead.
    ''' </returns>
    Public Function Read() As iniFile

        Return iniFile.FromFile(Path())

    End Function

    ''' <summary>
    ''' Returns a loaded <c> iniFile </c> if <c> Name </c> is non-empty and the file exists on disk,
    ''' otherwise returns <c> Nothing </c>. Use for optional files that may not be configured.
    ''' </summary>
    Public Function ReadIfSet() As iniFile

        If Name.Length = 0 OrElse Not Exists() Then Return Nothing
        Return iniFile.FromFile(Path())

    End Function

    ''' <summary>
    ''' Builds the File Chooser menu with all options and their dispatch handlers registered inline.
    ''' Called by both <see cref="PrintMenu"/> (to render) and <see cref="HandleInput"/>
    ''' (to dispatch), so the displayed option numbers and the dispatch table are always in sync.
    ''' </summary>
    Private Function buildMenu() As MenuSection

        Dim hasInitName = InitName.Length <> 0
        Dim hasSecondName = SecondName.Length <> 0
        Dim hasReset = Dir <> InitDir OrElse Name <> InitName

        Return MenuSection.CreateCompleteMenu("File Chooser",
            {"Choose a file name or directory",
             "Type a filename to change file  |  Type a path with \ to change directory"}) _
            .AddDispatchedOption(InitName, "Use the default name", Sub() Rename(InitName), hasInitName) _
            .AddDispatchedOption(SecondName, "Use the default rename", Sub() Rename(SecondName), hasSecondName) _
            .AddDispatchedOption("Default directory", "Use the same folder as winapp2ool.exe",
                Sub()
                    Dir = Environment.CurrentDirectory
                    exitModule()
                End Sub) _
            .AddDispatchedOption("Parent folder", "Go up one level",
                Sub()
                    Dir = Directory.GetParent(Dir).ToString()
                    exitModule()
                End Sub) _
            .AddDispatchedOption("Restore defaults", "Restore the original name and directory",
                Sub()
                    ResetParams()
                    ExitIfExists()
                End Sub, hasReset) _
            .AddBlank() _
            .AddFileInfo("Current File:      ", Path()) _
            .AddColoredLine($"Current Directory: {replDir(Dir)}", GetRedGreen(Not Directory.Exists(Dir)))

    End Function

    ''' <summary>
    ''' Prints the unified file/directory chooser menu
    ''' </summary>
    Public Sub PrintMenu()

        buildMenu().Print()

    End Sub

    ''' <summary>
    ''' Handles user input for the file/directory chooser menu. A number selects an option and
    ''' <c> 0 </c> leaves the chooser. Blank input leaves it too, subject to <see cref="MustExist"/>.
    ''' Input containing a backslash becomes the new directory if that directory exists, and
    ''' closes the chooser. Anything else becomes the new filename, except that with
    ''' <see cref="MustExist"/> on, a name with no file behind it reverts to the old one.
    ''' </summary>
    '''
    ''' <param name="input">
    ''' The user's input string
    ''' </param>
    Public Sub HandleInput(input As String)

        If input Is Nothing Then argIsNull(NameOf(input)) : Return

        Dim intInput As Integer

        If Integer.TryParse(input, intInput) Then

            If intInput = 0 Then exitModule() : Return
            If Not buildMenu().Dispatch(intInput) Then setNextMenuHeaderText(invInpStr, printColor:=ConsoleColor.Red)
            Return

        End If

        If input.Length = 0 Then ExitIfExists() : Return

        If input.Contains("\"c) Then

            Dim tmpDir = Dir
            Dir = input
            If Not Exists(False) Then

                setNextMenuHeaderText($"{Dir} does not exist", printColor:=ConsoleColor.Red)
                Dir = tmpDir

            Else

                exitModule()

            End If

            Return

        End If

        Rename(input)

    End Sub

    Private Sub Rename(nname As String)

        _tmpRename = Name
        Name = nname
        ExitIfExists(True)

    End Sub

    ''' <summary>
    ''' Closes the chooser if the file exists or <see cref="MustExist"/> is <c> False </c>.
    ''' Otherwise we set a header saying the file doesn't exist and keep the chooser open.
    ''' </summary>
    '''
    ''' <param name="undoPendingRename">
    ''' Indicates whether the rename which called this sub should be rolled back if the renamed file doesn't exist on disk <br /><br />
    ''' Optional, Default: <c> False </c>
    ''' </param>
    Private Sub ExitIfExists(Optional undoPendingRename As Boolean = False)

        If Not MustExist OrElse Exists() Then
            exitModule()
            Return
        End If

        setNextMenuHeaderText($"{Name} does not exist", printColor:=ConsoleColor.Red)
        If undoPendingRename Then Name = _tmpRename

    End Sub

    ''' <summary>
    ''' Restores the <see cref="Dir"/> and <see cref="Name"/> properties to the values used at construction time
    ''' </summary>
    Public Sub ResetParams()

        Dir = InitDir
        Name = InitName

    End Sub

End Class
