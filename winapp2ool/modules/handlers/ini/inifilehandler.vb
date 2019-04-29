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
Imports System.Text.RegularExpressions

''' <summary>
''' A collection of methods and classes for interacting with .ini files 
''' </summary>
Public Module iniFileHandler

    ''' <summary>Enforces that a user selected file exists</summary>
    ''' <param name="someFile">An iniFile object with user defined path and name parameters</param>
    Public Sub chkFileExist(someFile As iniFile)
        If pendingExit() Then Exit Sub
        While Not File.Exists(someFile.path)
            If pendingExit() Then Exit Sub
            setHeaderText("Error", True)
            clrConsole()
            printMenuTop({$"{someFile.name} does not exist."})
            print(1, "File Chooser (default)", "Change the file name")
            print(1, "Directory Chooser", "Change the directory", closeMenu:=True)
            Dim input As String = Console.ReadLine
            Select Case input
                Case "0"
                    exitCode = True
                    Exit Sub
                Case "1", ""
                    fileChooser(someFile)
                Case "2"
                    dirChooser(someFile.dir)
                Case Else
                    setHeaderText(invInpStr, True)
            End Select
            If Not File.Exists(someFile.path) And Not menuHeaderText = invInpStr Then setHeaderText("Error", True)
        End While
    End Sub

    ''' <summary>Enforces that a user defined directory exists, either by selecting a new one or creating one.</summary>
    ''' <param name="dir">A user defined windows directory</param>
    Public Sub chkDirExist(ByRef dir As String)
        If pendingExit() Or Directory.Exists(dir) Then Exit Sub
        Dim iExitCode As Boolean = False
        setHeaderText("Error", True)
        While Not iExitCode
            clrConsole()
            printMenuTop({$"{dir} does not exist."})
            print(1, "Create Directory", "Create this directory")
            print(1, "Directory Chooser (default)", "Specify a new directory", closeMenu:=True)
            Dim input As String = Console.ReadLine()
            Select Case input
                Case "0"
                    dir = Environment.CurrentDirectory
                    exitCode = iExitCode = True
                    Exit Sub
                Case "1"
                    Directory.CreateDirectory(dir)
                Case "2", ""
                    dirChooser(dir)
                Case Else
                    setHeaderText(invInpStr, True)
            End Select
            If Not Directory.Exists(dir) And Not menuHeaderText = invInpStr Then setHeaderText("Error", True)
            If Directory.Exists(dir) Then iExitCode = True
        End While
    End Sub

    ''' <summary>Presents a menu to the user allowing them to perform some file and directory operations</summary>
    ''' <param name="someFile">An iniFile object with user definable parameters</param>
    Public Sub fileChooser(ByRef someFile As iniFile)
        If pendingExit() Then Exit Sub
        clrConsole()
        handleFileChooserChoice(someFile)
        If pendingExit() Then Exit Sub
        handleFileChooserConfirm(someFile)
    End Sub

    ''' <summary>Allows the user change the file name or directory of a given iniFile object</summary>
    ''' <param name="someFile">The iniFile object whose parameters are being modified by the user</param>
    Private Sub handleFileChooserChoice(ByRef someFile As iniFile)
        setHeaderText("File Chooser")
        printMenuTop({"Choose a file name, or open the directory chooser to choose a directory"})
        print(1, someFile.initName, "Use the default name", someFile.initName <> "")
        print(1, someFile.secondName, "Use the default rename", someFile.secondName <> "")
        print(1, "Directory Chooser", "Choose a new directory", trailingBlank:=True)
        print(0, $"Current Directory: {replDir(someFile.dir)}")
        print(0, $"Current File:      {someFile.name}", closeMenu:=True)
        Console.Write($"{Environment.NewLine} Enter a number, {If(Not someFile.name = "", "", "or ")}a new file name{If(someFile.name = "", "", $", or leave blank to continue using '{someFile.name}'")}: ")
        Dim input As String = Console.ReadLine
        Select Case True
            Case input = "0"
                exitCode = True
                clrConsole()
                Exit Sub
            Case input = ""
            Case input = "1" And someFile.initName <> ""
                someFile.name = someFile.initName
            Case (input = "1" And someFile.initName = "") Or (input = "2" And someFile.secondName <> "")
                someFile.name = someFile.secondName
            Case (input = "2" And someFile.secondName = "") Or (input = "3" And someFile.initName <> "" And someFile.secondName <> "")
                dirChooser(someFile.dir)
            Case Else
                someFile.name = input
        End Select
    End Sub

    ''' <summary>Confirms the user's choice of a file's parameters in the File Chooser and allows them to make changes before saving</summary>
    ''' <param name="someFile">The iniFile object whose parameters are being modified by the user</param>
    Public Sub handleFileChooserConfirm(ByRef someFile As iniFile)
        setHeaderText("File Chooser")
        Dim iExitCode As Boolean = False
        Do Until iExitCode
            clrConsole()
            printMenuTop({"Confirm your settings or return to the options to change them."})
            print(1, "File Chooser", "Change the file name")
            print(1, "Directory Chooser", "Change the directory")
            print(1, "Confirm (default)", "Save changes")
            print(0, $"Current Directory: {replDir(someFile.dir)}", leadingBlank:=True)
            print(0, $"Current File     : {someFile.name}", closeMenu:=True)
            Console.Write(Environment.NewLine & "Enter a number, or leave blank to run the default: ")
            Dim input As String = Console.ReadLine()
            Select Case input
                Case "0"
                    exitCode = iExitCode = True
                Case "1"
                    fileChooser(someFile)
                    iExitCode = True
                Case "2"
                    dirChooser(someFile.dir)
                    iExitCode = True
                Case "3", ""
                    iExitCode = True
                Case Else
                    setHeaderText(invInpStr, True)
            End Select
        Loop
    End Sub

    ''' <summary>Presents an interface to the user allowing them to operate on windows directory parameters</summary>
    ''' <param name="dir">A user definable windows directory path</param>
    Public Sub dirChooser(ByRef dir As String)
        If pendingExit() Then Exit Sub
        clrConsole()
        handleDirChooserChoice(dir)
        If pendingExit() Then Exit Sub
        handleDirChooserConfirm(dir)
    End Sub

    ''' <summary>Allows the user to select a directory using a similar interface to the File Chooser</summary>
    ''' <param name="dir">The String containing the directory the user is parameterizing</param>
    Private Sub handleDirChooserChoice(ByRef dir As String)
        clrConsole()
        setHeaderText("Directory Chooser")
        printMenuTop({"Choose a directory"})
        print(1, "Use default (default)", "Use the same folder as winapp2ool.exe")
        print(1, "Parent Folder", "Go up a level")
        print(1, "Current folder", "Continue using the same folder as below")
        print(0, $"Current Directory: {dir}", leadingBlank:=True, closeMenu:=True)
        Console.Write(Environment.NewLine & "Choose a number from above, enter a new directory, or leave blank to run the default: ")
        Dim input As String = Console.ReadLine()
        Select Case input
            Case "0"
                exitCode = True
            Case "1", ""
                dir = Environment.CurrentDirectory
            Case "2"
                dir = Directory.GetParent(dir).ToString
            Case "3"
                clrConsole()
                chkDirExist(dir)
            Case Else
                dir = input
                clrConsole()
                chkDirExist(dir)
        End Select
    End Sub

    ''' <summary>Confirms the user's choice of directory and allows them to change it </summary>
    ''' <param name="dir">The String containing the directory the user is parameterizing</param>
    Private Sub handleDirChooserConfirm(ByRef dir As String)
        setHeaderText("Directory Chooser")
        Dim iExitCode As Boolean = False
        Do Until iExitCode
            If pendingExit() Then Exit Sub
            clrConsole()
            printMenuTop({"Choose a directory"})
            print(1, "Directory Chooser", "Change the directory")
            print(1, "Confirm (default)", "Use this directory")
            print(0, $"Current Directory: {dir}", leadingBlank:=True, closeMenu:=True)
            Console.Write(Environment.NewLine & "Choose a number from above, or leave blank to run the default: ")
            Dim input As String = Console.ReadLine()
            Select Case input
                Case "0"
                    iExitCode = True
                    exitCode = True
                Case "1"
                    dirChooser(dir)
                    iExitCode = True
                Case "2", ""
                    iExitCode = True
                Case Else
                    setHeaderText(invInpStr, True)
            End Select
        Loop
    End Sub
End Module