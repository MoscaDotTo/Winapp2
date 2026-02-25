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

''' <summary>An interactive file chooser for <c>iniFile2</c> objects, using a single unified menu</summary>
Public Class iniFileChooser

    ''' <summary>The current directory</summary>
    Public Property Dir As String

    ''' <summary>The current filename</summary>
    Public Property Name As String

    ''' <summary>The default filename offered as a numbered option</summary>
    Public ReadOnly Property InitName As String

    ''' <summary>An optional alternate default filename offered as a numbered option</summary>
    Public ReadOnly Property SecondName As String

    ''' <summary>When True, <c>Load</c> will loop until the chosen file exists on disk</summary>
    Public ReadOnly Property MustExist As Boolean

    Private _tmpRename As String = ""

    ''' <summary>
    ''' Creates an <c>iniFileChooser</c> with the given starting directory and filename
    ''' </summary>
    ''' <param name="dir">Starting directory</param>
    ''' <param name="name">Starting filename</param>
    ''' <param name="initName">Default filename offered as a quick-select option</param>
    ''' <param name="secondName">Alternate default filename offered as a quick-select option</param>
    ''' <param name="mustExist">When True, Load loops until the chosen file exists on disk</param>
    Public Sub New(dir As String, name As String,
                   Optional initName As String = "",
                   Optional secondName As String = "",
                   Optional mustExist As Boolean = True)
        Me.Dir = dir
        Me.Name = name
        Me.InitName = initName
        Me.SecondName = secondName
        Me.MustExist = mustExist
    End Sub

    ''' <summary>Returns the full path of the current file</summary>
    Public Function Path() As String
        Return $"{Dir}\{Name}"
    End Function

    ''' <summary>Returns whether the current file (or just directory, when <paramref name="checkPath"/> is False) exists on disk</summary>
    ''' <param name="checkPath">When True, checks the full file path; when False, checks only the directory</param>
    Public Function Exists(Optional checkPath As Boolean = True) As Boolean
        Return If(checkPath, File.Exists(Path()), Directory.Exists(Dir))
    End Function

    ''' <summary>
    ''' Validates the file, looping until it exists (or the user exits without selecting a valid file).
    ''' Returns a loaded <c>iniFile2</c>, or <c>Nothing</c> if the user exited without a valid selection.
    ''' </summary>
    Public Function Load() As iniFile2
        While Not Exists()
            initModule("File Chooser", AddressOf PrintMenu, AddressOf HandleInput)
            If Not Exists() Then Return Nothing
        End While
        Return iniFile2.FromFile(Path())
    End Function

    ''' <summary>Prints the unified file/directory chooser menu</summary>
    Public Sub PrintMenu()
        Dim hasInitName = InitName.Length <> 0
        Dim hasSecondName = SecondName.Length <> 0
        MenuSection.CreateCompleteMenu("File Chooser",
            {"Choose a file name or directory",
             "Type a filename to change file  |  Type a path with \ to change directory"}).
            AddOption(InitName, "Use the default name", hasInitName).
            AddOption(SecondName, "Use the default rename", hasSecondName).
            AddOption("Default directory", "Use the same folder as winapp2ool.exe").
            AddOption("Parent folder", "Go up one level").
            AddBlank().
            AddFileInfo("Current File:      ", Path()).
            AddColoredLine($"Current Directory: {replDir(Dir)}", GetRedGreen(Not Directory.Exists(Dir))).
            Print()
    End Sub

    ''' <summary>Handles user input for the file/directory chooser menu</summary>
    ''' <param name="input">The user's input string</param>
    Public Sub HandleInput(input As String)
        If input Is Nothing Then argIsNull(NameOf(input)) : Return
        Dim hasInitName = InitName.Length <> 0
        Dim hasSecondName = SecondName.Length <> 0
        Select Case True
            Case input = "0"
                exitModule()
            Case input.Length = 0
                ExitIfExists()
            Case hasInitName AndAlso input = "1"
                ReName(InitName)
            Case hasSecondName AndAlso input = computeMenuNumber(2, {Not hasInitName}, {-1})
                ReName(SecondName)
            Case input = computeMenuNumber(3, {Not hasInitName, Not hasSecondName}, {-1, -1})
                Dir = Environment.CurrentDirectory
                exitModule()
            Case input = computeMenuNumber(4, {Not hasInitName, Not hasSecondName}, {-1, -1})
                Dir = Directory.GetParent(Dir).ToString()
                exitModule()
            Case input.Contains("\"c)
                Dim tmpDir = Dir
                Dir = input
                If Not Exists(False) Then
                    setNextMenuHeaderText($"{Dir} does not exist", printColor:=ConsoleColor.Red)
                    Dir = tmpDir
                Else
                    exitModule()
                End If
            Case Else
                ReName(input)
        End Select
    End Sub

    Private Sub ReName(nname As String)
        _tmpRename = Name
        Name = nname
        ExitIfExists(True)
    End Sub

    Private Sub ExitIfExists(Optional undoPendingRename As Boolean = False)
        If Not Exists() AndAlso MustExist Then
            setNextMenuHeaderText($"{Name} does not exist", printColor:=ConsoleColor.Red)
            If undoPendingRename Then Name = _tmpRename
        Else
            exitModule()
        End If
    End Sub

End Class
