Option Strict On
Imports System.IO

Public Module Trim

    Dim winappFile As iniFile = New iniFile(Environment.CurrentDirectory, "winapp2.ini")
    Dim outputFile As iniFile = New iniFile(Environment.CurrentDirectory, "winapp2.ini", "winapp2-trimmed.ini")
    Dim winVer As Double = getWinVer()
    Dim detChrome As New List(Of String) _
        From {"%AppData%\ChromePlus\chrome.exe", "%LocalAppData%\Chromium\Application\chrome.exe", "%LocalAppData%\Chromium\chrome.exe", "%LocalAppData%\Flock\Application\flock.exe", "%LocalAppData%\Google\Chrome SxS\Application\chrome.exe",
                           "%LocalAppData%\Google\Chrome\Application\chrome.exe", "%LocalAppData%\RockMelt\Application\rockmelt.exe", "%LocalAppData%\SRWare Iron\iron.exe", "%ProgramFiles%\Chromium\Application\chrome.exe", "%ProgramFiles%\SRWare Iron\iron.exe",
                           "%ProgramFiles%\Chromium\chrome.exe", "%ProgramFiles%\Flock\Application\flock.exe", "%ProgramFiles%\Google\Chrome SxS\Application\chrome.exe", "%ProgramFiles%\Google\Chrome\Application\chrome.exe", "%ProgramFiles%\RockMelt\Application\rockmelt.exe",
                           "HKCU\Software\Chromium", "HKCU\Software\SuperBird", "HKCU\Software\Torch", "HKCU\Software\Vivaldi"}

    Dim settingsChanged As Boolean
    Dim download As Boolean = False
    Dim downloadNCC As Boolean = False

    'Restore the default state of the module parameters
    Private Sub initDefaultSettings()
        settingsChanged = False
        winappFile.resetParams()
        outputFile.resetParams()
        download = False
        downloadNCC = False
    End Sub

    'Pass the default module parameters to the commandline handler
    Public Sub initTrimParams(ByRef firstFile As iniFile, ByRef secondFile As iniFile, ByRef d As Boolean, ByRef dncc As Boolean)
        initDefaultSettings()
        firstFile = winappFile
        secondFile = outputFile
        d = download
        dncc = downloadNCC
    End Sub

    'Handle command line input and initiate the trim
    Public Sub remoteTrim(firstFile As iniFile, secondFile As iniFile, d As Boolean, ncc As Boolean)
        winappFile = firstFile
        outputFile = secondFile
        download = d
        downloadNCC = ncc
        initTrim()
    End Sub

    Private Sub resetSettings()
        initDefaultSettings()
        menuTopper = "Trim settings have been reset to their default state"
    End Sub

    Private Sub printmenu()
        printMenuTop({"Trim winapp2.ini such that it contains only entries relevant to your machine,", "greatly reducing both application load time and the winapp2.ini file size."}, True)
        printMenuOpt("Run (default)", "Trim winapp2.ini")

        printBlankMenuLine()
        printMenuOpt("Toggle Download", enStr(download) & " using the latest winapp2.ini as the input file")
        printMenuOpt("Toggle Download (Non-CCleaner)", enStr(downloadNCC) & " using the latest Non-CCleaner winapp2.ini as the input file")
        printIf(Not (download Or downloadNCC), "opt", "File Chooser (winapp2.ini)", "Change the winapp2.ini name or location")

        printBlankMenuLine()
        printMenuOpt("File Chooser (save)", "Change the save file name or location")

        printBlankMenuLine()
        printMenuLine("Current winapp2.ini location: " & If(download, If(downloadNCC, "Online (Non-CCleaner)", "Online"), replDir(winappFile.path)), "l")
        printMenuLine("Current save location: " & replDir(outputFile.path), "l")
        printIf(settingsChanged, "reset", "Trim", "")
        printMenuLine(menuStr02)
    End Sub

    Public Sub main()
        initMenu("Trim", 35)
        Dim input As String
        Do Until exitCode
            Console.Clear()
            printmenu()
            Console.WriteLine()
            Console.Write(promptStr)
            input = Console.ReadLine()
            Dim areDownloading As Boolean = download Or downloadNCC
            Select Case True
                Case input = "0"
                    Console.WriteLine("Returning to winapp2ool menu...")
                    exitCode = True
                Case (input = "1" Or input = "")
                    initTrim()
                Case input = "2"
                    toggleSettingParam(download, "Downloading ", settingsChanged)
                    If (Not download) And downloadNCC Then toggleSettingParam(downloadNCC, "Downloading ", settingsChanged)
                Case input = "3"
                    If Not download Then toggleSettingParam(download, "Downloading ", settingsChanged)
                    toggleSettingParam(downloadNCC, "Downloading ", settingsChanged)
                Case input = "4" And Not areDownloading
                    changeFileParams(winappFile, settingsChanged)
                Case input = "4" And areDownloading
                    changeFileParams(outputFile, settingsChanged)
                Case input = "5" And Not areDownloading
                    changeFileParams(outputFile, settingsChanged)
                Case input = "5" And settingsChanged And Not areDownloading
                    resetSettings()
                Case input = "6" And settingsChanged And areDownloading
                    resetSettings()
                Case Else
                    menuTopper = invInpStr
            End Select
        Loop
        revertMenu()
    End Sub

    Private Sub initTrim()
        'Validate winapp2.ini's existence, load it as an inifile and convert it to a winapp2file before passing it to trim
        If Not download Then
            winappFile.validate()
            If exitCode Then Exit Sub
            Dim winapp2 As New winapp2file(winappFile)
            trim(winapp2)
        Else
            Dim link As String = If(downloadNCC, nonccLink, wa2Link)
            initDownloadedTrim(link, "\winapp2.ini")
        End If
        menuTopper = "Trim Complete"
        Console.Clear()
    End Sub

    'Grab a remote ini file and toss it into the trimmer
    Public Sub initDownloadedTrim(link As String, name As String)
        trim(New winapp2file(getRemoteIniFile(link, name)))
    End Sub

    Private Sub trim(winapp2 As winapp2file)
        Console.Clear()
        printMenuLine("Trimming...", "l")

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

        'Print out the results from the trim
        printMenuLine("...done.", "l")
        Console.Clear()
        printMenuLine(tmenu("Trim Complete"))
        printMenuLine(menuStr03)
        printMenuLine("Results", "c")
        printBlankMenuLine()
        printMenuLine("Number of entries before trimming: " & entryCountBeforeTrim, "l")
        printMenuLine("Number of entries after trimming: " & winapp2.count, "l")
        printBlankMenuLine()
        printMenuLine("Press any key to return to the winapp2ool menu", "l")
        printMenuLine(menuStr02)

        Try
            'Rewrite our rebuilt ini back to disk
            Dim file As New StreamWriter(outputFile.path, False)
            file.Write(winapp2.winapp2string())
            file.Close()
        Catch ex As Exception
            exc(ex)
        Finally
            If Not suppressOutput Then Console.ReadKey()
        End Try
    End Sub

    'Returns true if an entry would be detected by CCleaner on parse, returns false otherwise
    Private Function processEntryExistence(ByRef entry As winapp2entry) As Boolean

        Dim hasDetOS As Boolean = Not entry.detectOS.Count = 0
        Dim hasMetDetOS As Boolean = False
        Dim exists As Boolean = False

        'Process the DetectOS if we have one, take note if we meet the criteria, otherwise return false
        If hasDetOS Then
            hasMetDetOS = detOSCheck(entry.detectOS(0).value)
            If Not hasMetDetOS Then Return False 
        End If

        'Process our SpecialDetect if we have one
        If Not entry.specialDetect.Count = 0 Then checkSpecialDetects(entry.specialDetect(0))

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

    'Processes a list of winapp2.ini entries and removes any from the list that wouldn't be detected by CCleaner
    Private Sub processEntryList(ByRef entryList As List(Of winapp2entry))
        Dim sectionsToBePruned As New List(Of winapp2entry)
        For Each entry In entryList
            If Not processEntryExistence(entry) Then sectionsToBePruned.Add(entry)
        Next
        removeEntries(entryList, sectionsToBePruned)
    End Sub

    'Return true if a SpecialDetect location exists
    Private Function checkSpecialDetects(ByVal key As iniKey) As Boolean
        Select Case key.value
            Case "DET_CHROME"
                For Each path As String In detChrome
                    If checkExist(path) Then Return True
                Next
            Case "DET_MOZILLA"
                Return checkFileExist("%AppData%\Mozilla\Firefox")
            Case "DET_THUNDERBIRD"
                Return checkFileExist("%AppData%\Thunderbird")
            Case "DET_OPERA"
                Return checkFileExist("%AppData%\Opera Software")
        End Select
        'If we didn't return above, SpecialDetect definitely doesn't exist
        Return False
    End Function

    'The only remaining use for this is the DET_CHROME which has a mix of file and registry detections 
    Private Function checkExist(key As String) As Boolean
        Return If(key.StartsWith("HK"), checkRegExist(key), checkFileExist(key))
    End Function

    'Returns True if a given Detect key exists on the system, otherwise returns False
    Private Function checkRegExist(key As String) As Boolean

        Dim dir As String = key
        Dim splitDir As String() = dir.Split(CChar("\"))

        Try
            'splitDir(0) sould contain the registry hive location, remove it and query the hive for existence
            Select Case splitDir(0)
                Case "HKCU"
                    dir = dir.Replace("HKCU\", "")
                    Return Microsoft.Win32.Registry.CurrentUser.OpenSubKey(dir, True) IsNot Nothing
                Case "HKLM"
                    dir = dir.Replace("HKLM\", "")
                    If Microsoft.Win32.Registry.LocalMachine.OpenSubKey(dir, True) IsNot Nothing Then Return True

                    'Small workaround for newer x64 versions of Windows
                    Dim rDir As String = dir.ToLower.Replace("software\", "Software\WOW6432Node\")
                    Return Microsoft.Win32.Registry.LocalMachine.OpenSubKey(rDir, True) IsNot Nothing
                Case "HKU"
                    dir = dir.Replace("HKU\", "")
                    Return Microsoft.Win32.Registry.Users.OpenSubKey(dir, True) IsNot Nothing
                Case "HKCR"
                    dir = dir.Replace("HKCR\", "")
                    Return Microsoft.Win32.Registry.ClassesRoot.OpenSubKey(dir, True) IsNot Nothing
            End Select
        Catch ex As Exception
            'The most common (only?) exception here is a permissions one, so assume true if we hit
            'because a permissions exception implies the key exists anyway.
            Return True
        End Try
        'If we didn't return anything above, registry location probably doesn't exist
        Return False
    End Function

    'Returns True if a DetectFile path exists on the system, otherwise returns False
    Private Function checkFileExist(key As String) As Boolean
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

            'Process wildcards appropriately if we have them 
            If dir.Contains("*") Then
                Dim exists As Boolean = False
                exists = expandWildcard(dir)

                'Small contingency for the isProgramFiles case
                If Not exists And isProgramFiles Then
                    Dim envDir As String = Environment.GetEnvironmentVariable("ProgramFiles(x86)")
                    dir = envDir & key.Split(CChar("%"))(2)
                    exists = expandWildcard(dir)
                End If
                Return exists
            End If

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

    'Correctly expand any wildcards for detection on the system, assumes that 
    Private Function expandWildcard(dir As String) As Boolean

        'This should handle wildcards anywhere in a path even though CCleaner only supports them at the end for DetectFiles
        Dim possibleDirs As New List(Of String)
        Dim currentPaths As New List(Of String)

        'Give an empty starter string for the first directory path to be added to 
        currentPaths.Add("")

        'Split the given string into sections by directory 
        Dim splitDir As String() = dir.Split(CChar("\"))

        For Each pathPart In splitDir

            'If this directory parameterization includes a wildcard, expand it appropriately
            'This probably wont work if a string for some reason starts with a *
            If pathPart.Contains("*") Then
                For Each currentPath In currentPaths
                    Try
                        'Query the existence of child paths for each current path we hold 
                        Dim possibilities As String() = Directory.GetDirectories(currentPath, pathPart)

                        'If there are any, add them to our possibility list
                        If Not possibilities.Count = 0 Then possibleDirs.AddRange(possibilities)
                    Catch
                        'Pretty much exception we'll encounter here is going to be the result of directories not existing.
                        'The exception will be thrown from the GetDirectories call and will prevent us from attempt to add new
                        'items to the possibility list. In this instance, we can silently fail (here). 
                    End Try
                Next

                'If no possibilities remain, the wildcard parameterization hasn't left us with any real paths on the system, so we may return false.
                If possibleDirs.Count = 0 Then Return False

                'Otherwise, clear the current paths and repopulate them with the possible paths 
                currentPaths.Clear()
                currentPaths.AddRange(possibleDirs)
                possibleDirs.Clear()
            Else
                If currentPaths.Count = 0 Then
                    currentPaths.Add(pathPart & "\")
                Else
                    Dim newCurPaths As New List(Of String)
                    For Each path As String In currentPaths
                        If Directory.Exists(path & pathPart & "\") Then newCurPaths.Add(path & pathPart & "\")
                    Next
                    currentPaths = newCurPaths
                End If
            End If
        Next

        'If any file/path exists, return true
        For Each currDir In currentPaths
            If Directory.Exists(currDir) Or File.Exists(currDir) Then Return True
        Next

        Return False
    End Function

    'Return True if we satisfy the DetectOS criteria, otherwise return False
    Private Function detOSCheck(value As String) As Boolean
        Dim splitKey As String() = value.Split(CChar("|"))
        Return If(value.StartsWith("|"), Not winVer > Double.Parse(splitKey(1)), Not winVer < Double.Parse(splitKey(0)))
    End Function

    'Returns the Windows Version mumber, 0.0 if it cannot determine the proper number.
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
                'I've never actually tested this function on anything but Windows 10, so if something goes wrong I hope someone reports it! 
                Console.WriteLine("Unable to determine which version of Windows you are running.")
                Console.WriteLine()
                Console.WriteLine("If you see this message, please report your Windows version on GitHub along with the following information:")
                Console.WriteLine("Winver:" & winver)
                Return 0.0
        End Select
    End Function
End Module