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
''' </summary>
Module Minefield

    Dim CLSIDS As List(Of iniKey)

    Public Sub printMenu()
        printMenuTop({"A testing ground for new ideas/features, watch your step!"})
        print(1, "Java  Entry Maker", "A meta script approach to entries in winapp2.ini", closeMenu:=True)
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
                Dim entryKeyList As New keyyList
                ' Get JavaPlugin and JavaScript keys in HKCR\
                Dim JavaPluginKeys As List(Of iniKey) = getRegKeys(Microsoft.Win32.Registry.ClassesRoot, {"JavaPlugin", "JavaScript"}.ToList)
                ' Separate out the JavaScript Author keys
                Dim JSAKeys As New List(Of iniKey)
                Dim JSKeys As New List(Of iniKey)
                JavaPluginKeys.ForEach(Sub(key) If key.toString.Contains("JavaScript") Then JSKeys.Add(key))
                JSKeys.ForEach(Sub(key) JavaPluginKeys.Remove(key))
                JSKeys.ForEach(Sub(key) If key.toString.Contains("Author") Then JSAKeys.Add(key))
                JSAKeys.ForEach(Sub(key) JSKeys.Remove(key))
                Dim errs As New List(Of iniKey)
                Dim clsids As New List(Of iniKey)
                Dim kll As New List(Of List(Of iniKey)) : kll.Add(clsids) : kll.Add(errs)
                Dim javaFile As New iniFile(Environment.CurrentDirectory, "java.ini")
                javaFile.validate()
                cwl("Running an intense registry query, this will take a few moments...")
                Dim jSect As iniSection = javaFile.sections.Item("Previous Java Installation Cleanup *")
                jSect.constructKeyLists({"clsid"}.ToList, kll)
                Dim IDS As List(Of String) = getValues(clsids)
                clsids.Clear()
                Dim typeLib As List(Of iniKey) = getRegKeys(Microsoft.Win32.Registry.ClassesRoot.OpenSubKey("WOW6432Node\TypeLib\"), {"{5852F5E0-8BF4-11D4-A245-0080C6F74284}"}.ToList)
                Dim crids As List(Of iniKey) = getRegKeys(Microsoft.Win32.Registry.ClassesRoot.OpenSubKey("WOW6432Node\CLSID\"), IDS)
                Dim lcids As List(Of iniKey) = getRegKeys(Microsoft.Win32.Registry.LocalMachine.OpenSubKey("SOFTWARE\Classes\WOW6432Node\CLSID\"), IDS)
                Dim lmids As List(Of iniKey) = getRegKeys(Microsoft.Win32.Registry.LocalMachine.OpenSubKey("SOFTWARE\WOW6432Node\Classes\CLSID\"), IDS)
                Dim kdids As List(Of iniKey) = getRegKeys(Microsoft.Win32.Registry.Users.OpenSubKey(".DEFAULT\Software\Classes\CLSID\"), IDS)
                Dim ksids As List(Of iniKey) = getRegKeys(Microsoft.Win32.Registry.Users.OpenSubKey("S-1-5-18\Software\Classes\CLSID\"), IDS)
                Dim lmjkeys As List(Of iniKey) = getRegKeys(Microsoft.Win32.Registry.LocalMachine.OpenSubKey("Software\JavaSoft\Java Runtime Environment"), {"1"}.ToList)
                Dim extlmKeys As New List(Of iniKey)
                lmjkeys.ForEach(Sub(key) If key.toString.Contains("_") Then extlmKeys.Add(key))
                extlmKeys.ForEach(Sub(key) lmjkeys.Remove(key))
                Dim cujkeys As List(Of iniKey) = getRegKeys(Microsoft.Win32.Registry.CurrentUser.OpenSubKey("Software\JavaSoft\Java Runtime Environment\"), {"1"}.ToList)
                Dim extcukeys As New List(Of iniKey)
                cujkeys.ForEach(Sub(key) If key.toString.Contains("_") Then extcukeys.Add(key))
                extcukeys.ForEach(Sub(key) cujkeys.Remove(key))
                If Not crids.Count = 0 Then crids.Remove(crids.Last)
                If Not lcids.Count = 0 Then lcids.Remove(lcids.Last)
                If Not lmids.Count = 0 Then lcids.Remove(lmids.Last)
                If Not kdids.Count = 0 Then kdids.Remove(kdids.Last)
                If Not ksids.Count = 0 Then ksids.Remove(ksids.Last)
                If Not lmjkeys.Count = 0 Then lmjkeys.Remove(lmjkeys.Last)
                If Not extlmKeys.Count = 0 Then extlmKeys.Remove(extlmKeys.Last)
                If Not cujkeys.Count = 0 Then cujkeys.Remove(cujkeys.Last)
                If Not extcukeys.Count = 0 Then extcukeys.Remove(extcukeys.Last)
                entryKeyList.add(crids)
                entryKeyList.add(typeLib)
                entryKeyList.add(lcids)
                entryKeyList.add(lmids)
                entryKeyList.add(lmjkeys)
                entryKeyList.add(extlmKeys)
                entryKeyList.add(cujkeys)
                entryKeyList.add(extcukeys)
                renumberKeys(entryKeyList.keys, replaceAndSort(getValues(entryKeyList.keys), "-", "-"))
                Dim entry As New List(Of String)
                entry.Add("[Java Installation Cleanup *]")
                entry.Add("Section=Experimental")
                entry.Add("Default=False")
                entryKeyList.keys.ForEach(Sub(key) entry.Add(key.toString))
                Dim out As New iniSection(entry)
                cwl(out.ToString)
                cwl() : cwl("Done")
                Console.ReadLine()
        End Select
    End Sub

    Private Function getRegKeys(reg As Microsoft.Win32.RegistryKey, searches As List(Of String)) As List(Of iniKey)
        Dim out As New List(Of iniKey)
        Try
            Dim keys As List(Of String) = reg.GetSubKeyNames.ToList
            searches.ForEach(Sub(search As String) keys.ForEach(
                Sub(key As String) If key.Contains(search) Then out.Add(New iniKey($"RegKey1={reg.ToString}\{key}"))))
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
