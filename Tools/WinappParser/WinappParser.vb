Option Strict On
Imports System.IO
Imports Microsoft.Win32

Module Module1
    Dim winfile As String = Environment.CurrentDirectory & "\scripts\winapp2.ini"
    Sub Main()

        'Individual line of the file
        Dim command As String

        'Stores an entire entry
        Dim currentEntry As String = Nothing

        Dim newFile As String = ""

        Try

            'Check if the winapp2.ini file is in the current directory. End program if it isn't.
            If File.Exists(winfile) = False Then
                Console.WriteLine("winapp2.ini file could not be located in the current working directory (" & winfile & ")")
                Console.ReadKey()
                End
            End If

            'Load the winapp2.ini file
            Dim r As IO.StreamReader
            r = New IO.StreamReader(winfile)
            Do While (r.Peek() > -1)

                'Update the line that is being tested
                command = (r.ReadLine.ToString)

                'Assume a space means the start of a new entry
                If command = "" Then

                    'Trim newlines from the start
                    currentEntry = currentEntry.Trim()

                    'Parse the current filec
                    ' MsgBox(currentEntry)
                    parse_entry(currentEntry)

                    'Reset the current entry
                    currentEntry = Nothing

                Else
                    'Combine each line into entries
                    currentEntry = currentEntry & Environment.NewLine & command
                End If

            Loop

            'Close stream
            r.Close()

        Catch ex As Exception
        End Try
        Console.ReadKey()
    End Sub

    'Parse the entries
    Private Sub parse_entry(ByVal entry As String)

        'remove the first entry
        If entry.StartsWith("; Version: v1.0.121112") Then
            Exit Sub
        End If

        'Instance variables
        Dim app_name As String = ""
        Dim detect_path As String = ""

        'Loop through the individual entry
        For Each line As String In entry.Split(CChar(Environment.NewLine))

            'Trim whitespace from each line
            line = (line.Trim())

            'Get the name of the app
            If line.StartsWith("[") Then
                app_name = line.Replace("[", "")
                app_name = app_name.Replace("*]", "")
                'MsgBox(app_name)
            End If

            'Special detects
            If line.StartsWith("SpecialDetect") Then

                'Mozilla detection
                If line.StartsWith("SpecialDetect=DET_MOZILLA") Then
                    If mozilla_path_exists() = False Then
                        Exit Sub
                    End If
                End If

                'Thunderbird detection
                If line.StartsWith("SpecialDetect=DET_THUNDERBIRD") Then
                    If thunderbird_path_exists() = False Then
                        Exit Sub
                    End If
                End If

                'Opera detection
                If line.StartsWith("SpecialDetect=OPERA") Then
                    If opera_path_exists() = False Then
                        Exit Sub
                    End If
                End If

            End If

            'Parse the detection rule
            If line.StartsWith("Detect") Then

                'Convert the path into WinAPI readable format
                detect_path = get_proper_path(line)

                'Determine if it's a directory detect
                If FileOrPathExists(detect_path) Then
                ElseIf RegKeyExists(detect_path) Then

                Else
                    'Requirement not fullfilled
                    Exit Sub
                End If

            End If

            'Don't run entries that contain a warning. They tend to do bad things
            If line.StartsWith("Warning=") Then
                Exit Sub
            End If

            'Now let's actually compute the FileKey's
            If line.StartsWith("FileKey") Then

                'Remove "FileKeyX="
                line = line.Remove(0, line.IndexOf("=") + 1)

                'Get a valid path (with footer)
                line = (get_proper_path(line))

                'Determine if the filepath should be recursed
                Dim recurse As Boolean = False
                If line.EndsWith("|RECURSE") Then
                    recurse = True
                    line = line.Replace("|RECURSE", "")
                End If

                'Get the clean directory
                Dim directory As String = line.Remove(line.IndexOf("|"))

                'Parse wildcards in the directory path
                If directory.Contains("*") Then

                    'Get the path containing the wildcards
                    Dim base_dir As String = directory.Remove(directory.IndexOf("*"))

                    'Confirm the base dir exists
                    If IO.Directory.Exists(base_dir) Then

                        'Loop through and get all subdirs
                        For Each sub_dir As String In IO.Directory.GetDirectories(base_dir)

                            'Connect the base directory will all subdirectories
                            directory = (directory.Replace(base_dir & "*", sub_dir))


                            MsgBox(directory)
                        Next
                        '  MsgBox("base of " & directory & " is " & base_dir)
                    End If

                End If

                'Ensure the directory exists
                If IO.Directory.Exists(directory) = False Then
                    Continue For
                End If


                'Get the file names/extensions that should be scanned
                Dim extensions As New List(Of String)

                'Get the list of file patterns from the line
                Dim filePatterns As String = line.Replace(directory & "|", "")
                For Each single_pattern As String In Split(filePatterns, ";")

                    'Add each individual pattern into the list of extensions
                    extensions.Add(single_pattern)

                Next


                'Run a scan to detect all the specified file
                If recurse = True Then
                    For Each foundFile As String In My.Computer.FileSystem.GetFiles(directory, FileIO.SearchOption.SearchAllSubDirectories, "*.*")

                        'Match each filePattern
                        For Each pattern As String In extensions
                            If foundFile Like pattern Then
                                Console.WriteLine(foundFile)
                            End If
                        Next
                    Next

                Else
                    For Each foundFile As String In My.Computer.FileSystem.GetFiles(directory, FileIO.SearchOption.SearchTopLevelOnly, "*.*")

                        'Match each filePattern
                        For Each pattern As String In extensions
                            If foundFile Like pattern Then
                                Console.WriteLine(foundFile)
                            End If

                        Next
                    Next
                End If


            End If

        Next

    End Sub
    Private Function mozilla_path_exists() As Boolean

        'List all possible Firefox paths
        If IO.Directory.Exists(System.Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) & "\Mozilla\Firefox\") Then
            Return True
        ElseIf IO.Directory.Exists(System.Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) & "\Moonchild Productions\PaleMoon\") Then
            Return True
        ElseIf IO.Directory.Exists(System.Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) & "\Mozilla\Waterfox\") Then
            Return True
        ElseIf IO.Directory.Exists(System.Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) & "\Moonchild Productions\Pale Moon\") Then
            Return True
        Else
            Return False
        End If

    End Function
    Private Function opera_path_exists() As Boolean
        If IO.Directory.Exists(System.Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) & "\Opera") Then
            Return True
        Else
            Return False
        End If
    End Function
    Private Function thunderbird_path_exists() As Boolean
        If IO.Directory.Exists(System.Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) & "\Thunderbird\") Then
            Return True
        ElseIf IO.Directory.Exists(System.Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) & "\Mozilla Messaging\Thunderbird\") Then
            Return True
        Else
            Return False
        End If
    End Function
    Private Function FileOrPathExists(ByVal path As String) As Boolean
        Try
            If IO.File.Exists(path) Then
                Return True
            ElseIf IO.Directory.Exists(path) Then
                Return True
            Else
                Return False

            End If
        Catch ex As Exception
            If IO.Directory.Exists(path) Then
                Return True
            Else
                Return False
            End If
        End Try
    End Function

    Private Function get_proper_path(ByVal Command As String) As String

        'Replace detect entries
        Command = Command.Replace("Detect=", "")
        Command = Command.Replace("DetectFile=", "")

        'Replace environmental variables
        Command = Command.Replace("%ProgramFiles%", System.Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles))
        Command = Command.Replace("%windir%", "C:\Windows")
        Command = Command.Replace("%appdata%", System.Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData))
        Command = Command.Replace("%AppData%", System.Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData))
        Command = Command.Replace("%LocalAppData%", System.Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData))
        Command = Command.Replace("%CommonAppData%", System.Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData))

        'Username directory
        If IO.Directory.Exists("C:\Users\" & Environment.UserName) Then
            Command = Command.Replace("%UserProfile%", "C:\Users\" & Environment.UserName)
            Command = Command.Replace("%UserProfile%", "C:\Users\" & Environment.UserName)
        Else
            Command = Command.Replace("%UserProfile%", "C:\Documents and Settings\" & Environment.UserName)
            Command = Command.Replace("%UserProfile%", "C:\Documents and Settings\" & Environment.UserName)
        End If

        'Application data
        ' Command = Command.Replace(Application.ProductName & "\", "")
        '  Command = Command.Replace(Application.CompanyName & "\", "")
        '  Command = Command.Replace(Application.ProductVersion & "\", "")

        'System directories
        Command = Command.Replace("%rootdir%", "C:")
        Command = Command.Replace("%homedrive%", "C:")
        Command = Command.Replace("%WinDir%", "C:\Windows")
        Command = Command.Replace("%systemdrive%", "C:")
        Command = Command.Replace("%CommonProgramFiles%", System.Environment.GetFolderPath(Environment.SpecialFolder.CommonProgramFiles))

        Return Command

    End Function

    'Check if registry value exists.
    Public Function RegKeyExists(ByVal key As String) As Boolean
        Try
            'Replace any '|' characters with slashes for compatibility with this algorithm
            If key.Contains("|") = True Then
                key = key.Replace("|", "\")
            End If

            'Filter any trailing backslashes, as they make everything go to hell.
            If key.EndsWith("\") Then
                key = key.TrimEnd(CChar("\"))
            End If

            'preserve the original path
            Dim original_key As String = key

            'Check for the existence of the subkey, assigning the key value to a predefined registry hive
            Dim regKey As Microsoft.Win32.RegistryKey
            If key.StartsWith("HKLM\") Then
                key = key.Replace("HKLM\", "")
                regKey = Registry.LocalMachine.OpenSubKey(key, True)
            ElseIf key.StartsWith("HKCU\") Then
                key = key.Replace("HKCU\", "")
                regKey = Registry.CurrentUser.OpenSubKey(key, True)
            ElseIf key.StartsWith("HKCR\") Then
                key = key.Replace("HKCR\", "")
                regKey = Registry.ClassesRoot.OpenSubKey(key, True)
            ElseIf key.StartsWith("HKCC\") Then
                key = key.Replace("HKCC\", "")
                regKey = Registry.CurrentConfig.OpenSubKey(key, True)
            ElseIf key.StartsWith("HKUS\") Then
                key = key.Replace("HKUS\", "")
                regKey = Registry.Users.OpenSubKey(key, True)
            Else
                'prevent a null reference exception in rare circumstances
                regKey = Nothing
            End If

            'Conditional to check if key exists
            If regKey Is Nothing Then 'doesn't exist. Check for values.

                'Values require the hive name to be specified. Do the replacement.
                If original_key.StartsWith("HKLM\") Then
                    original_key = original_key.Replace("HKLM\", "HKEY_LOCAL_MACHINE\")
                ElseIf key.StartsWith("HKCU\") Then
                    original_key = original_key.Replace("HKCU\", "HKEY_CURRENT_USER\")
                ElseIf original_key.StartsWith("HKCR\") Then
                    original_key = original_key.Replace("HKCR\", "HKEY_CLASSES_ROOT\")
                ElseIf original_key.StartsWith("HKUS\") Then
                    original_key = original_key.Replace("HKUS\", "HKEY_USERS\")
                ElseIf original_key.StartsWith("HKCC\") Then
                    original_key = original_key.Replace("HKCC\", "HKEY_CURRENT_CONFIG\")
                End If

                'Work out the valuename and subkey
                Dim last_index As Integer = original_key.LastIndexOf("\")
                Dim valueName As String = original_key.Remove(0, (last_index + 1))
                original_key = original_key.Remove(last_index)

                'Do the conditional on the specified value in the subkey
                Dim regValue As Object = Registry.GetValue(original_key, valueName, Nothing)
                If regValue Is Nothing Then
                    Return False
                Else
                    Return True
                End If

            Else
                'It was the subkey after all, return true.
                Return True
            End If

        Catch ex As Exception
            'An exception has occurred. Return false to be on the safe side.
            Return False
        End Try
    End Function

End Module
