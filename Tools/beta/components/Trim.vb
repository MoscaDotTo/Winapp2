Option Strict On
Imports System.IO

Public Module Trim

    Dim dir As String = Environment.CurrentDirectory
    Dim name As String = "\winapp2.ini"
    Dim waName As String = "\winapp2.ini"
    Dim waDir As String = Environment.CurrentDirectory
    Dim winVer As Double = getWinVer()

    Dim hasMenuTopper As Boolean = False
    Dim exitCode As Boolean = False

    Private Sub printmenu()
        tmenu("Trim")
        menu(menuStr03)
        menu("This tool will trim winapp2.ini such that it contains only entries relevant to your machine,", "c")
        menu("greatly reducing both application load time and the winapp2.ini file size.", "c")
        menu(menuStr04)
        menu("0. Exit                - Return to the winapp2ool menu", "l")
        menu("1. Trim (default)      - Trim winapp2.ini and overwrite the existing file", "l")
        menu("2. Trim (custom)       - Change the save directory and/or file name", "l")
        menu(menuStr02)
    End Sub

    Public Sub main()
        exitCode = False
        hasMenuTopper = False
        Console.Clear()

        Dim input As String
        Do Until exitCode
            printmenu()
            Console.WriteLine()
            Console.Write("Enter a number, or leave blank to run the default: ")
            input = Console.ReadLine()

            Select Case input
                Case "0"
                    Console.WriteLine("Returning to winapp2ool menu...")
                    exitCode = True
                Case "1", ""
                    trim()
                    revertMenu(exitCode)
                    Console.Clear()
                Case "2"
                    changeParams()
                    trim()
                    revertMenu(exitCode)
                    Console.Clear()
                Case Else
                    Console.Write("Invalid input. Please try again: ")
                    input = Console.ReadLine
            End Select
        Loop
    End Sub

    Private Sub changeParams()
        If exitCode Then
            Exit Sub
        End If
        dChooser(dir, exitCode)
        While Not Directory.Exists(dir)
            dChooser(dir, exitCode)
        End While
        If exitCode Then
            Exit Sub
        End If
        fChooser(dir, name, exitCode, "\winapp2.ini", "\winapp2-trimmed.ini")
        If exitCode Then
            Exit Sub
        End If
    End Sub

    Private Sub trim()

        'load winapp2.ini into memory
        validate(waDir, waName, exitCode, "\winapp2.ini", "\winapp2-2.ini")
        If exitCode Then
            Exit Sub
        End If
        Dim winappfile As New iniFile(waDir, waName)
        Console.Clear()
        tmenu("Trim")
        menu("Trimming...", "l")

        'create holders for the trimmed file
        Dim trimmedfile As New iniFile
        Dim cEntries As New iniFile
        Dim fxEntries As New iniFile
        Dim tbEntries As New iniFile

        Dim detChrome As New List(Of String)
        detChrome.AddRange(New String() {"%AppData%\ChromePlus\chrome.exe", "%LocalAppData%\Chromium\Application\chrome.exe", "%LocalAppData%\Chromium\chrome.exe", "%LocalAppData%\Flock\Application\flock.exe", "%LocalAppData%\Google\Chrome SxS\Application\chrome.exe",
                           "%LocalAppData%\Google\Chrome\Application\chrome.exe", "%LocalAppData%\RockMelt\Application\rockmelt.exe", "%LocalAppData%\SRWare Iron\iron.exe", "%ProgramFiles%\Chromium\Application\chrome.exe", "%ProgramFiles%\SRWare Iron\iron.exe",
                           "%ProgramFiles%\Chromium\chrome.exe", "%ProgramFiles%\Flock\Application\flock.exe", "%ProgramFiles%\Google\Chrome SxS\Application\chrome.exe", "%ProgramFiles%\Google\Chrome\Application\chrome.exe", "%ProgramFiles%\RockMelt\Application\rockmelt.exe",
                           "Software\Chromium", "Software\SuperBird", "Software\Torch", "Software\Vivaldi"})

        For i As Integer = 0 To winappfile.sections.Count - 1

            Dim cursection As iniSection = winappfile.sections.Values(i)
            Dim hasDetOS As Boolean = False
            Dim hasMetDetOS As Boolean = True
            Dim hasDets As Boolean = False
            Dim exists As Boolean = False
            For j As Integer = 0 To cursection.keys.Count - 1
                Dim key As iniKey = cursection.keys(j)

                If exists Then
                    Exit For
                End If

                Dim type As String = key.keyType

                Select Case type
                    Case "Detect"
                        hasDets = True
                        If checkExist(key.value) Then
                            trimmedfile.sections.Add(cursection.name, cursection)
                            Exit For
                        End If
                    Case "DetectFile"
                        hasDets = True
                        If checkExist(key.value) Then
                            trimmedfile.sections.Add(cursection.name, cursection)
                            Exit For
                        End If
                    Case "DetectOS"
                        hasDetOS = True
                        hasMetDetOS = detOSCheck(key.value)

                        'move on if we don't meet the DetectOS criteria
                        If Not hasMetDetOS Then
                            Exit For
                        End If
                    Case "SpecialDetect"
                        hasDets = True

                        'handle our SpecialDetects
                        Select Case key.value
                            Case "DET_CHROME"
                                For Each path As String In detChrome
                                    If checkExist(path) Then
                                        cEntries.sections.Add(cursection.name, cursection)
                                        Exit For
                                    End If
                                Next
                                Exit For
                            Case "DET_MOZILLA"
                                If checkExist("%AppData%\Mozilla\Firefox") Then
                                    fxEntries.sections.Add(cursection.name, cursection)
                                    Exit For
                                End If
                            Case "DET_THUNDERBIRD"
                                If checkExist("%AppData%\Thunderbird") Then
                                    tbEntries.sections.Add(cursection.name, cursection)
                                    Exit For
                                End If
                            Case "DET_OPERA"
                                If checkExist("%AppData%\Opera Software") Then
                                    trimmedfile.sections.Add(cursection.name, cursection)
                                    Exit For
                                End If
                        End Select
                End Select
            Next

            If hasDetOS Then
                If Not hasDets And hasMetDetOS Then
                    trimmedfile.sections.Add(cursection.name, cursection)
                End If
            Else
                If Not hasDets Then
                    trimmedfile.sections.Add(cursection.name, cursection)
                End If
            End If

        Next
        menu("...done.", "l")
        Console.Clear()
        tmenu("Trim Complete")
        moMenu("Results")
        menu(menuStr03)
        menu("Number of entries before trimming: " & winappfile.sections.Count, "l")
        Dim entrycount As Integer = trimmedfile.sections.Count + cEntries.sections.Count + fxEntries.sections.Count + tbEntries.sections.Count
        menu("Number of entries after trimming: " & entrycount, "l")
        menu(menuStr01)
        menu("Press any key to return to the winapp2ool menu", "l")
        menu(menuStr02)
        Try
            Dim file As New StreamWriter(dir & name, False)

            Dim comNum As Integer
            'contingency for non-cc ini
            'does not correctly restore for removed entries (todo) 
            If winappfile.comments.Count > 16 Then
                comNum = 9
            Else
                comNum = 8
            End If

            file.WriteLine(winappfile.getStreamOfComments(0, comNum))
            comNum += 1
            file.WriteLine()
            file.WriteLine(cEntries.toString)

            file.WriteLine(winappfile.getStreamOfComments(comNum, comNum + 2))
            comNum += 3
            file.WriteLine()
            file.WriteLine(fxEntries.toString)

            file.WriteLine(winappfile.getStreamOfComments(comNum, comNum + 2))
            file.WriteLine()
            file.WriteLine(tbEntries.toString)

            comNum += 3
            file.WriteLine(winappfile.getStreamOfComments(comNum, comNum))
            file.WriteLine()

            file.Write(trimmedfile)
            file.Close()
        Catch ex As Exception
            Console.WriteLine("Error: " & ex.ToString)
            Console.WriteLine("Please report this error on GitHub")
            Console.WriteLine()
        Finally
            Console.ReadKey()
        End Try
    End Sub

    Private Function checkExist(key As String) As Boolean

        'Observe the registry paths
        If key.StartsWith("HK") Then
            Return checkRegExist(key)
        Else
            Return checkFileExist(key)
        End If
    End Function

    Private Function checkRegExist(key As String) As Boolean

        Dim dir As String = key
        Dim splitDir As String() = dir.Split(Convert.ToChar("\"))

        Try
            Select Case splitDir(0)
                Case "HKCU"
                    dir = dir.Replace("HKCU\", "")
                    If Microsoft.Win32.Registry.CurrentUser.OpenSubKey(dir, True) IsNot Nothing Then
                        Return True
                    End If
                Case "HKLM"
                    dir = dir.Replace("HKLM\", "")
                    If Microsoft.Win32.Registry.LocalMachine.OpenSubKey(dir, True) IsNot Nothing Then
                        Return True
                    Else
                        Dim rDir As String = dir.ToLower.Replace("software\", "Software\WOW6432Node\")
                        If Microsoft.Win32.Registry.LocalMachine.OpenSubKey(rDir, True) IsNot Nothing Then
                            Return True
                        End If
                    End If
                Case "HKU"
                    dir = dir.Replace("HKU\", "")
                    If Microsoft.Win32.Registry.Users.OpenSubKey(dir, True) IsNot Nothing Then
                        Return True
                    End If
                Case "HKCR"
                    dir = dir.Replace("HKCR\", "")
                    If Microsoft.Win32.Registry.ClassesRoot.OpenSubKey(dir, True) IsNot Nothing Then
                        Return True
                    End If
            End Select
        Catch ex As Exception
            'The most common exception here is a permissions one, so assume true if we hit
            'because a permissions exception implies the key exists anyway.
            Return True
        End Try
        Return False
    End Function

    Private Function checkFileExist(key As String) As Boolean

        Dim isProgramFiles As Boolean = False
        Dim dir As String = key

        'make sure we get the proper path for environment variables
        If dir.Contains("%") Then
            Dim splitDir As String() = dir.Split(Convert.ToChar("%"))
            Dim var As String = splitDir(1)
            Dim envDir As String = Environment.GetEnvironmentVariable(var)
            Select Case var
                Case "ProgramFiles"
                    isProgramFiles = True
                Case "Documents"
                    envDir = Environment.GetEnvironmentVariable("UserProfile") & "\Documents"
                Case "CommonAppData"
                    envDir = Environment.GetEnvironmentVariable("ProgramData")
            End Select
            dir = envDir + splitDir(2)
        End If

        Try
            'check out those file/folder paths
            If Directory.Exists(dir) Or File.Exists(dir) Then
                Return True
            End If

            'if we didn't find it and we're looking in Program Files, check the (x86) directory
            If isProgramFiles Then

                Dim envDir As String = Environment.GetEnvironmentVariable("ProgramFiles(x86)")
                dir = envDir & key.Split(CChar("%"))(2)
                Return (Directory.Exists(dir) Or File.Exists(dir))

            End If

        Catch ex As Exception
            Console.WriteLine("Error: " & ex.ToString)
            Console.WriteLine("Please report this error on GitHub")
            Console.WriteLine()
            Return True
        End Try

        Return False
    End Function

    Private Function detOSCheck(value As String) As Boolean

        Dim splitKey As String() = value.Split(Convert.ToChar("|"))

        If value.StartsWith("|") Then
            If winVer > Double.Parse(splitKey(1)) Then
                Return False
            Else
                Return True
            End If
        Else
            If winVer < Double.Parse(splitKey(0)) Then
                Return False
            Else
                Return True
            End If
        End If
    End Function

    Private Function getWinVer() As Double

        Dim winver As String = My.Computer.Info.OSFullName.ToString
        Select Case True
            Case winver.Contains("XP")
                Return 5.1
            Case winver.Contains("Vista")
                Return 6.0
            Case winver.Contains("7")
                Return 6.1
            Case winver.Contains("8") And Not winver.Contains("8.1")
                Return 6.2
            Case winver.Contains("8.1")
                Return 6.3
            Case winver.Contains("10")
                Return 10.0
            Case Else
                Console.WriteLine("Unable to determine which version of Windows you are running.")
                Console.WriteLine()
                Return 0.0
        End Select
    End Function
End Module