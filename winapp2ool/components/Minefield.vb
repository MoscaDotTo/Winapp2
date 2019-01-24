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
''' <summary>
''' This module serves as a simple interface for testing new ideas that don't fit neatly into an existing module
''' The ideas here should be considered alpha and the code here should be considered spaghetti 
''' </summary>
Module Minefield

    Public Sub printMenu()
        printMenuTop({"A testing ground for new ideas/features, watch your step!"})
        print(1, "Java Entry Maker", "A meta script approach to entries in winapp2.ini", closeMenu:=True)
    End Sub

    Public Sub handleUserInput(input As String)
        Select Case input
            Case "0"
                exitCode = True
            Case "1"
                initModule("Java Entry Maker", AddressOf printJMMenu, AddressOf handleJMInput)
        End Select
    End Sub

    Private Sub printJMMenu()
        printMenuTop({"Creates a winapp2.ini entry that cleans up after old Java versions"})
        print(1, "Run (Default)", "Attempt to create an entry based on the current system", closeMenu:=True)
    End Sub

    Private Sub handleJMInput(input As String)
        Select Case input
            Case "0"
                exitCode = True
            Case "1", ""
                ' Load the java.ini file
                Dim javaFile As New iniFile(Environment.CurrentDirectory, "java.ini")
                javaFile.validate()
                ' Get JavaPlugin and JavaScript keys in HKCR\
                Dim JavaPluginKeys As New keyyList(getRegKeys(Microsoft.Win32.Registry.ClassesRoot, {"JavaPlugin", "JavaScript"}.ToList))
                ' Get the CLSIDs present on the current system
                Dim clsids As New keyyList("CLSID")
                Dim kll As New List(Of keyyList) From {clsids, New keyyList("Errors")}
                cwl("Running an intense registry query, this will take a few moments...")
                Dim jSect As iniSection = javaFile.sections.Item("Previous Java Installation Cleanup *")
                jSect.constKeyLists(kll)
                Dim IDS = clsids.toListOfStr(True)
                Dim typeLib As New keyyList(getRegKeys(Microsoft.Win32.Registry.ClassesRoot.OpenSubKey("WOW6432Node\TypeLib\"), {"{5852F5E0-8BF4-11D4-A245-0080C6F74284}"}.ToList))
                Dim classRootIDs As New keyyList(getRegKeys(Microsoft.Win32.Registry.ClassesRoot.OpenSubKey("WOW6432Node\CLSID\"), IDS))
                Dim localMachineClassesIDs As New keyyList(getRegKeys(Microsoft.Win32.Registry.LocalMachine.OpenSubKey("SOFTWARE\Classes\WOW6432Node\CLSID\"), IDS))
                Dim localMachineWOWIds As New keyyList(getRegKeys(Microsoft.Win32.Registry.LocalMachine.OpenSubKey("SOFTWARE\WOW6432Node\Classes\CLSID\"), IDS))
                Dim defClassesIDs As New keyyList(getRegKeys(Microsoft.Win32.Registry.Users.OpenSubKey(".DEFAULT\Software\Classes\CLSID\"), IDS))
                Dim s1518ClassesIDs As New keyyList(getRegKeys(Microsoft.Win32.Registry.Users.OpenSubKey("S-1-5-18\Software\Classes\CLSID\"), IDS))
                Dim localMachineJREs As New keyyList(getRegKeys(Microsoft.Win32.Registry.LocalMachine.OpenSubKey("Software\JavaSoft\Java Runtime Environment"), {"1"}.ToList))
                Dim lmJREminorIDs, cuJREminorIDs As New keyyList
                localMachineJREs.keys.ForEach(Sub(key) lmJREminorIDs.add(key, key.toString.Replace("HKEY_LOCAL_MACHINE", "").Contains("_")))
                localMachineJREs.remove(lmJREminorIDs.keys)
                Dim currentUserJREs As New keyyList(getRegKeys(Microsoft.Win32.Registry.CurrentUser.OpenSubKey("Software\JavaSoft\Java Runtime Environment\"), {"1"}.ToList))
                currentUserJREs.keys.ForEach(Sub(key) cuJREminorIDs.add(key, key.toString.Replace("HKEY_CURRENT_USER", "").Contains("_")))
                currentUserJREs.remove(cuJREminorIDs.keys)
                ' Generate the list of Registry Keys
                Dim regKeyList As keyyList = mkEntry({classRootIDs, localMachineClassesIDs, localMachineWOWIds, defClassesIDs, s1518ClassesIDs,
                        localMachineJREs, lmJREminorIDs, currentUserJREs, cuJREminorIDs, JavaPluginKeys})
                ' Renumber them 
                regKeyList.renumberKeys(replaceAndSort(regKeyList.toListOfStr(True), "-", "-"))
                ' Generate the new entry
                Dim entry As New List(Of String)
                entry.Add("[Java Installation Cleanup *]")
                entry.Add("Section=Experimental")
                entry.Add($"Detect={regKeyList.keys(0).value}")
                entry.Add("Default=False")
                entry.AddRange(regKeyList.toListOfStr)
                Dim out As New iniSection(entry)
                cwl(out.ToString)
                cwl() : cwl("Done")
                Console.ReadLine()
        End Select
    End Sub

    ''' <summary>
    ''' Generates the RegKeylist for the current system
    ''' </summary>
    ''' <param name="kls">A list of keylists containing the RegKeys that will be in the generated entry</param>
    ''' <returns></returns>
    Private Function mkEntry(kls As keyyList()) As keyyList
        Dim out As New keyyList
        For Each lst In kls
            lst.removeLast()
            out.add(lst.keys)
        Next
        Return out
    End Function

    Private Function getRegKeys(reg As Microsoft.Win32.RegistryKey, searches As List(Of String)) As List(Of iniKey)
        Dim out As New List(Of iniKey)
        Try
            Dim keys As List(Of String) = reg.GetSubKeyNames.ToList
            searches.ForEach(Sub(search) keys.ForEach(Sub(key) If key.Contains(search) Then out.Add(New iniKey($"RegKey1={reg.ToString}\{key}"))))
        Catch ex As Exception
            ' The only Exception we can expect here is that reg is not set to an instance of an object
            ' This occurs when the requested registry key does not exist on the current system
            ' We can just silently fail if that's the case 
        End Try
        Return out
    End Function

    Private Sub printGMMenu()
        printMenuTop({"This tool will allow for a more meta approach to creating entries for games, particularly steam."})
        print(1, "Run (Disabled)", "Attempt to generate entries", closeMenu:=True)
    End Sub

    Private Sub handleGMInput(input As String)
        Select Case input
            Case "0"
                exitCode = True
        End Select
    End Sub
End Module
