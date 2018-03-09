Option Strict On
Imports System.IO

Public Module Trim

    Dim dir As String = Environment.CurrentDirectory
    Dim name As String = "\winapp2.ini"
    Dim waName As String = "\winapp2.ini"
    Dim waDir As String = Environment.CurrentDirectory
    Dim winVer As Double = getWinVer()
    Dim detChrome As List(Of String)

    'This boolean will prevent us from printing so we can call inner functions silently (eg. trim from commandline), WIP
    Public suppressOutput As Boolean

    Dim hasMenuTopper As Boolean = False
    Dim exitCode As Boolean = False

    Private Sub printmenu()
        printMenuLine(tmenu("Trim"))
        printMenuLine(menuStr03)
        printMenuLine("This tool will trim winapp2.ini such that it contains only entries relevant to your machine,", "c")
        printMenuLine("greatly reducing both application load time and the winapp2.ini file size.", "c")
        printMenuLine(menuStr04)
        printMenuLine("0. Exit                - Return to the winapp2ool menu", "l")
        printMenuLine("1. Run (default)      - Trim winapp2.ini and overwrite the existing file", "l")
        printMenuLine("2. Run (custom)       - Change the save directory and/or file name", "l")
        printMenuLine("3. Run (download)     - Download and trim the latest winapp2.ini and save it to the current folder", "l")
        printMenuLine(menuStr02)
    End Sub

    Public Sub main()
        detChrome = New List(Of String)
        detChrome.AddRange(New String() {"%AppData%\ChromePlus\chrome.exe", "%LocalAppData%\Chromium\Application\chrome.exe", "%LocalAppData%\Chromium\chrome.exe", "%LocalAppData%\Flock\Application\flock.exe", "%LocalAppData%\Google\Chrome SxS\Application\chrome.exe",
                           "%LocalAppData%\Google\Chrome\Application\chrome.exe", "%LocalAppData%\RockMelt\Application\rockmelt.exe", "%LocalAppData%\SRWare Iron\iron.exe", "%ProgramFiles%\Chromium\Application\chrome.exe", "%ProgramFiles%\SRWare Iron\iron.exe",
                           "%ProgramFiles%\Chromium\chrome.exe", "%ProgramFiles%\Flock\Application\flock.exe", "%ProgramFiles%\Google\Chrome SxS\Application\chrome.exe", "%ProgramFiles%\Google\Chrome\Application\chrome.exe", "%ProgramFiles%\RockMelt\Application\rockmelt.exe",
                           "HKCU\Software\Chromium", "HKCU\Software\SuperBird", "HKCU\Software\Torch", "HKCU\Software\Vivaldi"})


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
                    suppressOutput = False
                    initTrim()
                    revertMenu(exitCode)
                    Console.Clear()
                Case "2"
                    suppressOutput = False
                    fChooser(dir, name, exitCode, "\winapp2.ini", "\winapp2-trimmed.ini")
                    initTrim()
                    revertMenu(exitCode)
                    Console.Clear()
                Case "3"
                    suppressOutput = False
                    initDownloadedTrim(Downloader.wa2Link, "\winapp2.ini")
                    revertMenu(exitCode)
                    Console.Clear()
                Case Else
                    Console.Write("Invalid input. Please try again: ")
                    input = Console.ReadLine
            End Select
        Loop
    End Sub

    'Grab a remote ini file and toss it into the trimmer
    Public Sub initDownloadedTrim(link As String, name As String)
        Dim winapp2 As New winapp2file(getRemoteIniFile(link, name))
        trim(winapp2)
    End Sub

    Private Function processEntryExistence(ByRef entry As winapp2entry) As Boolean
        'Returns true if an entry would be detected by CCleaner on parse, returns false otherwise

        Dim hasDetOS As Boolean = Not entry.detectOS.Count = 0
        Dim hasMetDetOS As Boolean = False
        Dim exists As Boolean = False

        'Process the DetectOS if we have one, return true if we meet its criteria
        If hasDetOS Then
            If detOSCheck(entry.detectOS(0).value) Then
                hasMetDetOS = True
            Else Return False
            End If
        End If

        'Process our SpecialDetect if we have one
        If Not entry.specialDetect.Count = 0 Then Return checkSpecialDetects(entry.specialDetect(0))

        'Process our Detects
        For Each key In entry.detects
            If checkRegExist(key.value) Then Return True
        Next

        'Process our DetectFiles
        For Each key In entry.detectFiles
            If checkFileExist(key.value) Then Return True
        Next

        'Return true for the special case where we have only a DetectOS and we meet its criteria 
        If hasMetDetOS And entry.specialDetect.Count = 0 And entry.detectFiles.Count = 0 And entry.detects.Count = 0 Then Return True

        Return False
    End Function

    Private Sub processEntryList(ByRef entryList As List(Of winapp2entry))
        'Processes a list of winapp2.ini entries and removes any from the list that wouldn't be detected by CCleaner

        Dim sectionsToBePruned As New List(Of winapp2entry)
        For Each entry In entryList
            If Not processEntryExistence(entry) Then sectionsToBePruned.Add(entry)
        Next

        removeEntries(entryList, sectionsToBePruned)
    End Sub

    Private Sub initTrim()
        'Validate winapp2.ini's existence, load it as an inifile and convert it to a winapp2file before passing it to trim
        validate(waDir, waName, exitCode, "\winapp2.ini", "\winapp2-2.ini")
        If exitCode Then Exit Sub
        Dim winappfile As New iniFile(waDir, waName)
        Dim winapp2 As New winapp2file(winappfile)
        trim(winapp2)
    End Sub

    Private Sub trim(winapp2 As winapp2file)
        Console.Clear()
        If Not suppressOutput Then
            printMenuLine(tmenu("Trim"))
            printMenuLine("Trimming...", "l")
        End If

        'Save our inital # of entries
        Dim entryCountBeforeTrim As Integer = winapp2.count

        'Process our entries
        processEntryList(winapp2.cEntriesW)
        processEntryList(winapp2.fxEntriesW)
        processEntryList(winapp2.tbEntriesW)
        processEntryList(winapp2.mEntriesW)

        'Update the internal inifile objects
        winapp2.rebuildToIniFiles()

        'Sort the file so that entries are written back alphabetically
        winapp2.sortInneriniFiles()

        If Not suppressOutput Then printTrimComplete(entryCountBeforeTrim, winapp2.count)

        Try
            'Rewrite our rebuilt ini back to disk
            Dim file As New StreamWriter(dir & name, False)

            file.Write(winapp2.winapp2string())
            file.Close()
        Catch ex As Exception
            exc(ex)
        Finally
            If Not suppressOutput Then Console.ReadKey()
        End Try
    End Sub

    Private Sub printTrimComplete(before As Integer, after As Integer)
        'Print out the results from the trim
        printMenuLine("...done.", "l")
        Console.Clear()
        printMenuLine(tmenu("Trim Complete"))
        printMenuLine(moMenu("Results"))
        printMenuLine(menuStr03)
        printMenuLine("Number of entries before trimming: " & before, "l")
        printMenuLine("Number of entries after trimming: " & after, "l")
        printMenuLine(menuStr01)
        printMenuLine("Press any key to return to the winapp2ool menu", "l")
        printMenuLine(menuStr02)

    End Sub

    Private Function checkSpecialDetects(ByVal key As iniKey) As Boolean
        'Return true if a SpecialDetect location exists

        Select Case key.value
            Case "DET_CHROME"
                For Each path As String In detChrome
                    If checkExist(path) Then Return True
                Next
            Case "DET_MOZILLA"
                If checkExist("%AppData%\Mozilla\Firefox") Then Return True
            Case "DET_THUNDERBIRD"
                If checkExist("%AppData%\Thunderbird") Then Return True
            Case "DET_OPERA"
                If checkExist("%AppData%\Opera Software") Then Return True
        End Select
        Return False
    End Function

    Private Function checkExist(key As String) As Boolean
        'The only remaining use for this is the DET_CHROME which has a mix of file and registry detections 

        If key.StartsWith("HK") Then
            Return checkRegExist(key)
        Else
            Return checkFileExist(key)
        End If
    End Function

    Private Function checkRegExist(key As String) As Boolean
        'Returns True if a given Detect key exists on the system, otherwise returns False

        Dim dir As String = key
        Dim splitDir As String() = dir.Split(Convert.ToChar("\"))

        Try
            Select Case splitDir(0)
                Case "HKCU"
                    dir = dir.Replace("HKCU\", "")
                    If Microsoft.Win32.Registry.CurrentUser.OpenSubKey(dir, True) IsNot Nothing Then Return True
                Case "HKLM"
                    dir = dir.Replace("HKLM\", "")
                    If Microsoft.Win32.Registry.LocalMachine.OpenSubKey(dir, True) IsNot Nothing Then
                        Return True
                    Else
                        Dim rDir As String = dir.ToLower.Replace("software\", "Software\WOW6432Node\")
                        If Microsoft.Win32.Registry.LocalMachine.OpenSubKey(rDir, True) IsNot Nothing Then Return True
                    End If
                Case "HKU"
                    dir = dir.Replace("HKU\", "")
                    If Microsoft.Win32.Registry.Users.OpenSubKey(dir, True) IsNot Nothing Then Return True
                Case "HKCR"
                    dir = dir.Replace("HKCR\", "")
                    If Microsoft.Win32.Registry.ClassesRoot.OpenSubKey(dir, True) IsNot Nothing Then Return True
            End Select
        Catch ex As Exception
            'The most common (only?) exception here is a permissions one, so assume true if we hit
            'because a permissions exception implies the key exists anyway.
            Return True
        End Try
        Return False
    End Function

    Private Function checkFileExist(key As String) As Boolean
        'Returns True if a DetectFile path exists on the system, otherwise returns False
        Dim isProgramFiles As Boolean = False
        Dim dir As String = key

        'make sure we get the proper path for environment variables
        If dir.Contains("%") Then
            Dim splitDir As String() = dir.Split(CChar("%"))
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
            If Directory.Exists(dir) Or File.Exists(dir) Then Return True

            'if we didn't find it and we're looking in Program Files, check the (x86) directory
            If isProgramFiles Then
                Dim envDir As String = Environment.GetEnvironmentVariable("ProgramFiles(x86)")
                dir = envDir & key.Split(CChar("%"))(2)
                Return (Directory.Exists(dir) Or File.Exists(dir))
            End If

        Catch ex As Exception
            exc(ex)
            Return True
        End Try

        Return False
    End Function

    Private Function detOSCheck(value As String) As Boolean
        'Return True if we satisfy the DetectOS criteria, otherwise return False

        Dim splitKey As String() = value.Split(Convert.ToChar("|"))

        If value.StartsWith("|") Then
            Return Not winVer > Double.Parse(splitKey(1))
        Else
            Return Not winVer < Double.Parse(splitKey(0))
        End If
    End Function

    Private Function getWinVer() As Double
        'Returns the Windows Version mumber, 0.0 if it cannot determine the proper number.

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