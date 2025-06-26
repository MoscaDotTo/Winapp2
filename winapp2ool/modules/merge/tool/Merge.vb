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

''' <summary> 
''' Facilitates the merger of one <c> iniFile </c> into another
''' </summary>
''' 
''' Docs last updated: 2020-09-14 | Code last updated: 2020-09-14
Module Merge

    ''' <summary> 
    ''' Handles the commandline args for Merge 
    ''' </summary>
    ''' 
    ''' <remarks> 
    ''' Merge args: <br /> 
    ''' -mm         : toggle mergemode from replace conflicts to remove conflicts <br />
    ''' Preset merge file choices <br />
    ''' -r          : removed entries.ini  <br />
    ''' -c          : custom.ini  <br />
    ''' -w          : winapp3.ini <br />
    ''' -a          : archived entries.ini <br /> 
    ''' </remarks>
    ''' Docs last updated: 2020-09-14 | Code last updated: 2020-09-14
    Public Sub handleCmdLine()

        initDefaultMergeSettings()

        invertSettingAndRemoveArg(mergeMode, "-mm")
        invertSettingAndRemoveArg(False, "-r", MergeFile2.Name, "Removed Entries.ini")
        invertSettingAndRemoveArg(False, "-c", MergeFile2.Name, "Custom.ini")
        invertSettingAndRemoveArg(False, "-w", MergeFile2.Name, "winapp3.ini")
        invertSettingAndRemoveArg(False, "-a", MergeFile2.Name, "Archived Entries.ini")

        getFileAndDirParams(MergeFile1, MergeFile2, MergeFile3)
        If MergeFile2.Name.Length <> 0 Then initMerge()

    End Sub

    ''' <summary> 
    ''' Validates the <c> iniFiles </c> and kicks off the merging process 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2020-09-14 | Code last updated: 2020-09-14
    Public Sub initMerge()

        clrConsole()

        If Not (enforceFileHasContent(MergeFile1) AndAlso enforceFileHasContent(MergeFile2)) Then Return

        print(4, $"Merging {MergeFile1.Name} with {MergeFile2.Name}")

        mergeFiles(True)

        print(0, "", closeMenu:=True)
        print(3, $"Finished merging files. {anyKeyStr}")
        crk()

    End Sub

    ''' <summary> 
    ''' Conducts the merger of our two iniFiles 
    ''' </summary>
    ''' 
    ''' <param name="isWinapp2"> 
    ''' Indicates that the <c> iniFiles </c> being operated on are of winapp2.ini syntax 
    ''' </param>
    ''' 
    ''' Docs last updated: 2020-09-14 | Code last updated: 2020-09-14
    Private Sub mergeFiles(isWinapp2 As Boolean)

        resolveConflicts(MergeFile1, MergeFile2)

        ' After conflicts are resolved, remaining entries need only be added 
        For Each section In MergeFile2.Sections.Values

            MergeFile1.Sections.Add(section.Name, section)
            print(0, $"Adding {section.Name}", colorLine:=True, enStrCond:=True)

        Next

        If isWinapp2 Then

            Dim tmp As New winapp2file(MergeFile1)
            tmp.sortInneriniFiles()
            MergeFile1.Sections = tmp.toIni.Sections
            MergeFile3.overwriteToFile(tmp.winapp2string)

        Else

            MergeFile1.sortSections(MergeFile1.namesToStrList)
            MergeFile3.overwriteToFile(MergeFile1.toString)

        End If

    End Sub

    ''' <summary> 
    ''' Facilitates merging <c> iniFiles </c> from outside the module 
    ''' </summary>
    ''' 
    ''' <param name="mergeFile">
    ''' An <c> iniFile </c> to whom content will be added 
    ''' </param>
    ''' 
    ''' <param name="sourceFile">
    ''' An <c> iniFile </c> whose content will be added to <c> <paramref name="mergeFile"/> </c> 
    ''' </param>
    ''' 
    ''' <param name="isWinapp"> 
    ''' Indicates that the <c> iniFiles </c> being worked with contain winapp2.ini syntax 
    ''' </param>
    ''' 
    ''' <param name="mm"> 
    ''' Indicates the <c> MergeMode </c> for the merger <br/> 
    ''' <c> True </c> to replace conflicts, <c> False </c> to remove them 
    ''' </param>
    ''' 
    ''' Docs last updated: 2020-09-14 | Code last updated: 2020-09-14
    Public Sub RemoteMerge(ByRef mergeFile As iniFile,
                           ByRef sourceFile As iniFile,
                           isWinapp As Boolean, mm As Boolean)

        MergeFile1 = mergeFile
        MergeFile2 = sourceFile
        mergeMode = mm
        mergeFiles(isWinapp)

    End Sub

    ''' <summary> 
    ''' Performs conflict resolution for the merge process, handling the case where <c> iniSections </c> 
    ''' in both files have the same <c> Name </c> 
    ''' </summary>
    ''' 
    ''' <param name="first"> 
    ''' The master <c> iniFile </c> 
    ''' </param>
    ''' 
    ''' <param name="second"> 
    ''' The <c> iniFile </c> whose contents will be merged into <c> <paramref name="first"/> </c> 
    ''' </param>
    ''' 
    ''' Docs last updated: 2020-09-14 | Code last updated: 2020-09-14
    Private Sub resolveConflicts(ByRef first As iniFile,
                                 ByRef second As iniFile)

        Dim removeList As New List(Of String)

        For Each section In second.Sections.Keys

            If Not first.Sections.Keys.Contains(section) Then Continue For

            print(0, $"{If(mergeMode, "Replacing", "Removing")} {first.Sections.Item(section).Name}", colorLine:=True, enStrCond:=mergeMode)

            ' If mergemode is true, replace the match. otherwise, remove the match
            If mergeMode Then first.Sections.Item(section) = second.Sections.Item(section) Else first.Sections.Remove(section)
            removeList.Add(section)

        Next

        ' Remove any processed sections from the second file so that only entries to add remain
        For Each section In removeList

            second.Sections.Remove(section)

        Next

    End Sub

End Module