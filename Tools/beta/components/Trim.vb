'    Copyright (C) 2018 Robbie Ward
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
'''    <summary>
'''    This module parses a winapp2.ini file and checks each entry therein 
'''    removing any whose detection parameters do not exist on the current system
'''    and outputting a "trimmed" file containing only entries that exist on the system
'''    to the user.
'''   </summary>
Public Module Trim

    Dim winappFile As iniFile = New iniFile(Environment.CurrentDirectory, "winapp2.ini")
    Dim outputFile As iniFile = New iniFile(Environment.CurrentDirectory, "winapp2.ini", "winapp2-trimmed.ini")
    Dim winVer As Double
    Dim detChrome As New List(Of String) _
        From {"%AppData%\ChromePlus\chrome.exe", "%LocalAppData%\Chromium\Application\chrome.exe", "%LocalAppData%\Chromium\chrome.exe", "%LocalAppData%\Flock\Application\flock.exe", "%LocalAppData%\Google\Chrome SxS\Application\chrome.exe",
                           "%LocalAppData%\Google\Chrome\Application\chrome.exe", "%LocalAppData%\RockMelt\Application\rockmelt.exe", "%LocalAppData%\SRWare Iron\iron.exe", "%ProgramFiles%\Chromium\Application\chrome.exe", "%ProgramFiles%\SRWare Iron\iron.exe",
                           "%ProgramFiles%\Chromium\chrome.exe", "%ProgramFiles%\Flock\Application\flock.exe", "%ProgramFiles%\Google\Chrome SxS\Application\chrome.exe", "%ProgramFiles%\Google\Chrome\Application\chrome.exe", "%ProgramFiles%\RockMelt\Application\rockmelt.exe",
                           "HKCU\Software\Chromium", "HKCU\Software\SuperBird", "HKCU\Software\Torch", "HKCU\Software\Vivaldi"}

    Dim settingsChanged As Boolean
    Dim download As Boolean = False
    Dim downloadNCC As Boolean = False
    Dim areDownloading As Boolean = False

    ''' <summary>
    ''' Restores the default state of the module's parameters
    ''' </summary>
    Private Sub initDefaultSettings()
        settingsChanged = False
        winappFile.resetParams()
        outputFile.resetParams()
        download = False
        downloadNCC = False
    End Sub

    ''' <summary>
    ''' Initializes the default module settings and returns references to them to the calling function
    ''' </summary>
    ''' <param name="firstFile">The winapp2.ini file</param>
    ''' <param name="secondFile">The output file</param>
    ''' <param name="d">Download boolean</param>
    ''' <param name="dncc">Non-CCleaner download boolean</param>
    Public Sub initTrimParams(ByRef firstFile As iniFile, ByRef secondFile As iniFile, ByRef d As Boolean, ByRef dncc As Boolean)
        initDefaultSettings()
        firstFile = winappFile
        secondFile = outputFile
        d = download
        dncc = downloadNCC
    End Sub

    ''' <summary>
    ''' Runs the trimmer from outside the module
    ''' </summary>
    ''' <param name="firstFile">The winapp2.ini file</param>
    ''' <param name="secondFile">The output file</param>
    ''' <param name="d">Download boolean</param>
    ''' <param name="ncc">Non-CCleaner download boolean</param>
    Public Sub remoteTrim(firstFile As iniFile, secondFile As iniFile, d As Boolean, ncc As Boolean)
        winappFile = firstFile
        outputFile = secondFile
        download = d
        downloadNCC = ncc
        initTrim()
    End Sub

    ''' <summary>
    ''' Resets the module settings
    ''' </summary>
    Private Sub resetSettings()
        initDefaultSettings()
        menuTopper = "Trim settings have been reset to their default state"
    End Sub

    ''' <summary>
    ''' Prints the main menu to the user
    ''' </summary>
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

    ''' <summary>
    ''' The main event loop for the trimmer
    ''' </summary>
    Public Sub main()
        initMenu("Trim", 35)
        Do Until exitCode
            Console.Clear()
            printmenu()
            Console.WriteLine()
            Console.Write(promptStr)
            handleUserInput(Console.ReadLine())
        Loop
        revertMenu()
    End Sub

    ''' <summary>
    ''' Handles the user input from the menu
    ''' </summary>
    ''' <param name="input">The String containing the user's input</param>
    Private Sub handleUserInput(input As String)
        areDownloading = download Or downloadNCC
        Select Case True
            Case input = "0"
                Console.WriteLine("Returning to winapp2ool menu...")
                exitCode = True
            Case (input = "1" Or input = "")
                initTrim()
            Case input = "2"
                If Not denySettingOffline() Then
                    toggleSettingParam(download, "Downloading ", settingsChanged)
                    If (Not download) And downloadNCC Then toggleSettingParam(downloadNCC, "Downloading ", settingsChanged)
                End If
            Case input = "3"
                If Not denySettingOffline() Then
                    If Not download Then toggleSettingParam(download, "Downloading ", settingsChanged)
                    toggleSettingParam(downloadNCC, "Downloading ", settingsChanged)
                End If
            Case input = "4" And Not areDownloading
                changeFileParams(winappFile, settingsChanged)
            Case (input = "4" And areDownloading) Or (input = "5" And Not areDownloading)
                changeFileParams(outputFile, settingsChanged)
            Case (input = "5" And settingsChanged And areDownloading) Or (input = "6" And settingsChanged And Not areDownloading)
                resetSettings()
            Case Else
                menuTopper = invInpStr
        End Select
    End Sub

    ''' <summary>
    ''' Initiates the trim after validating our ini files
    ''' </summary>
    Private Sub initTrim()
        If Not download Then
            winappFile.validate()
            If exitCode Then Exit Sub
            Dim winapp2 As New winapp2file(winappFile)
            trim(winapp2)
        Else
            Dim link As String = If(downloadNCC, nonccLink, wa2Link)
            initDownloadedTrim(link)
        End If
        menuTopper = "Trim Complete"
        Console.Clear()
    End Sub

    ''' <summary>
    ''' Load a remote .ini file and trim it
    ''' </summary>
    ''' <param name="link">The link to the online resource</param>
    Public Sub initDownloadedTrim(link As String)
        trim(New winapp2file(getRemoteIniFile(link)))
    End Sub

    ''' <summary>
    ''' Performs the trim
    ''' </summary>
    ''' <param name="winapp2">A winapp2.ini file</param>
    Private Sub trim(winapp2 As winapp2file)
        Console.Clear()
        printMenuLine("Trimming...", "l")
        'Save our inital # of entries
        Dim entryCountBeforeTrim As Integer = winapp2.count

        'Process our entries
        For Each entryList In winapp2.winapp2entries
            processEntryList(entryList)
        Next

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

        'Write our rebuilt ini back to disk
        outputFile.overwriteToFile(winapp2.winapp2string)
        If Not suppressOutput Then Console.ReadKey()
    End Sub

    ''' <summary>
    ''' Returns true if an entry's detection criteria is matched by the system, false otherwise.
    ''' </summary>
    ''' <param name="entry">A winapp2.ini entry</param>
    ''' <returns></returns>
    Private Function processEntryExistence(ByRef entry As winapp2entry) As Boolean
        Dim hasDetOS As Boolean = Not entry.detectOS.Count = 0
        Dim hasMetDetOS As Boolean = False

        'Process the DetectOS if we have one, take note if we meet the criteria, otherwise return false
        If hasDetOS Then
            If winVer = Nothing Then winVer = getWinVer()
            hasMetDetOS = detOSCheck(entry.detectOS(0).value)
            If Not hasMetDetOS Then Return False
        End If

        'Process our SpecialDetect if we have one
        For Each key In entry.specialDetect
            If checkSpecialDetects(key) Then Return True
        Next

        'Process our Detects
        For Each key In entry.detects
            If checkRegExist(key.value) Then Return True
        Next

        'Process our DetectFiles
        For Each key In entry.detectFiles
            If checkFileExist(key.value) Then Return True
        Next

        'Return true for the case where we have only a DetectOS and we meet its criteria 
        If hasMetDetOS And entry.specialDetect.Count = 0 And entry.detectFiles.Count = 0 And entry.detects.Count = 0 Then Return True
        'Return true for the case where we have no valid detect criteria
        If entry.detectOS.Count + entry.detectFiles.Count + entry.detects.Count + entry.specialDetect.Count = 0 Then Return True

        Return False
    End Function

    ''' <summary>
    ''' Processess a list of winapp2.ini entries and removes any from the list that wouldn't be detected by CCleaner
    ''' </summary>
    ''' <param name="entryList">The list of winapp2entry objects to check existence for</param>
    Private Sub processEntryList(ByRef entryList As List(Of winapp2entry))
        Dim sectionsToBePruned As New List(Of winapp2entry)
        For Each entry In entryList
            If Not processEntryExistence(entry) Then sectionsToBePruned.Add(entry)
        Next
        removeEntries(entryList, sectionsToBePruned)
    End Sub

    ''' <summary>
    ''' Returns true if a SpecialDetect location exist, false otherwise
    ''' </summary>
    ''' <param name="key">A SpecialDetect format iniKey</param>
    ''' <returns></returns>
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

    ''' <summary>
    ''' Handles passing off checks for the DET_CHROME case
    ''' </summary>
    ''' <param name="key"></param>
    ''' <returns></returns>
    Private Function checkExist(key As String) As Boolean
        Return If(key.StartsWith("HK"), checkRegExist(key), checkFileExist(key))
    End Function

    ''' <summary>
    ''' Returns True if a given Detect path exists in the system registry, false otherwise.
    ''' </summary>
    ''' <param name="key">A registry path from a Detect key</param>
    ''' <returns></returns>
    Private Function checkRegExist(key As String) As Boolean
        Dim dir As String = key
        Dim root As String = getFirstDir(key)
        Try
            Select Case root
                Case "HKCU"
                    dir = dir.Replace("HKCU\", "")
                    Return Microsoft.Win32.Registry.CurrentUser.OpenSubKey(dir, True) IsNot Nothing
                Case "HKLM"
                    dir = dir.Replace("HKLM\", "")
                    If Microsoft.Win32.Registry.LocalMachine.OpenSubKey(dir, True) IsNot Nothing Then Return True
                    'Small workaround for newer x64 versions of Windows
                    dir = dir.ToLower.Replace("software\", "Software\WOW6432Node\")
                    Return Microsoft.Win32.Registry.LocalMachine.OpenSubKey(dir, True) IsNot Nothing
                Case "HKU"
                    dir = dir.Replace("HKU\", "")
                    Return Microsoft.Win32.Registry.Users.OpenSubKey(dir, True) IsNot Nothing
                Case "HKCR"
                    dir = dir.Replace("HKCR\", "")
                    Return Microsoft.Win32.Registry.ClassesRoot.OpenSubKey(dir, True) IsNot Nothing
            End Select
        Catch ex As Exception
            'The most common (only?) exception here is a permissions one, so assume true if we hit because a permissions exception implies the key exists anyway.
            Return True
        End Try
        'If we didn't return anything above, registry location probably doesn't exist
        Return False
    End Function

    ''' <summary>
    ''' Returns Truei f a DetectFile path exists on the system, false otherwise.
    ''' </summary>
    ''' <param name="key">A DetectFile path</param>
    ''' <returns></returns>
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
                    swapDir(dir, key)
                    exists = expandWildcard(dir)
                End If
                Return exists
            End If

            'check out those file/folder paths
            If Directory.Exists(dir) Or File.Exists(dir) Then Return True
            'if we didn't find it and we're looking in Program Files, check the (x86) directory
            If isProgramFiles Then
                swapDir(dir, key)
                Return (Directory.Exists(dir) Or File.Exists(dir))
            End If
        Catch ex As Exception
            exc(ex)
            Return True
        End Try
        Return False
    End Function

    ''' <summary>
    ''' Swaps out a directory with the ProgramFiles parameterization on 64bit computers
    ''' </summary>
    ''' <param name="dir">The text to be modified</param>
    ''' <param name="key">The original state of the text</param>
    Private Sub swapDir(ByRef dir As String, key As String)
        Dim envDir As String = Environment.GetEnvironmentVariable("ProgramFiles(x86)")
        dir = envDir & key.Split(CChar("%"))(2)
    End Sub

    ''' <summary>
    ''' Interpret parameterized wildcards for the current system
    ''' </summary>
    ''' <param name="dir">A path containing a wildcard</param>
    ''' <returns></returns>
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
                        'The exception will be thrown from the GetDirectories call and will prevent us from attempting to add new
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

    ''' <summary>
    ''' Returns true if we satisfy the DetectOS citeria, false otherwise
    ''' </summary>
    ''' <param name="value">The DetectOS criteria to be checked</param>
    ''' <returns></returns>
    Private Function detOSCheck(value As String) As Boolean
        Dim splitKey As String() = value.Split(CChar("|"))
        Return If(value.StartsWith("|"), Not winVer > Double.Parse(splitKey(1)), Not winVer < Double.Parse(splitKey(0)))
    End Function
End Module