'    Copyright (C) 2018-2025 Hazel Ward
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
Public Module Executor

    Public Property AllEntriesSettings As Dictionary(Of String, Dictionary(Of String, Boolean))

    ''' <summary>Prints the main Executor menu</summary>
    Public Sub printExecutorMenu()

        printMenuTop({"Deletes files and registry keys from your system"})
        print(1, "Scan (Default)", "Run the tool in scanning mode (nothing will be deleted)")
        print(1, "Toggle scan settings settings", "Modify which entries are executed by the executor")
        print(1, "Execute", "Run the tool in deletion mode", closeMenu:=True)

    End Sub

    ''' <summary> 
    ''' Handles the user input for Executor 
    ''' </summary>
    ''' 
    ''' <param name="input">
    ''' The String containing the user's input
    ''' </param>
    ''' 
    Public Sub handleExecutorInput(input As String)

        Select Case input
            Case "0"
                exitModule()
            Case "1", ""

                ' Load winapp2.ini from disk 
                Dim wa2 As New iniFile(Environment.CurrentDirectory, "winapp2.ini")
                wa2.validate()
                Dim winapp2 As New winapp2file(wa2)



                ' Trim irrelevant entries 
                Trim.trimFile(winapp2)

                ' Create a new settings dictionary containing the entries, default all values to false 
                Dim EntriesSettings As New Dictionary(Of String, Boolean)

                ' Later, we'll support loading these from disk and modifying them individually from the commandline, profiles, etc. For now, lets just
                ' make everything false to start with 
                For Each entrySection In winapp2.Winapp2entries

                    For Each entry In entrySection
                        EntriesSettings.Add(entry.Name, False)
                    Next

                Next

                cwl()


            Case "2"

        End Select
    End Sub
End Module
