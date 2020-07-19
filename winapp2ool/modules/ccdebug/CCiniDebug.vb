'    Copyright (C) 2018-2020 Robbie Ward
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
''' <summary> This module performs housekeeping on ccleaner.ini to clean up leftovers from winapp2.ini </summary>
Module CCiniDebug

    ''' <summary> The winapp2.ini file that ccleaner.ini may optionally be checked against </summary>
    ''' Docs last updated: 2020-07-18 | Code last updated: 2020-07-18
    Public Property CCDebugFile1 As New iniFile(Environment.CurrentDirectory, "winapp2.ini")

    '''<summary> The ccleaner.ini file to be debugged </summary>
    ''' Docs last updated: 2020-07-18 | Code last updated: 2020-07-18
    Public Property CCDebugFile2 As New iniFile(Environment.CurrentDirectory, "ccleaner.ini", mExist:=True)

    '''<summary> Holds the path for the debugged file that will be saved to disk. Overwrites ccleaner.ini by default </summary>
    ''' Docs last updated: 2020-07-18 | Code last updated: 2020-07-18
    Public Property CCDebugFile3 As New iniFile(Environment.CurrentDirectory, "ccleaner.ini", "ccleaner-debugged.ini")

    '''<summary> Indicates that stale winapp2.ini entries should be pruned from ccleaner.ini <br/> Default: <c> True </c> </summary>
    '''<remarks> A "stale" entry is one that appears in an (app) key in <c> CCDebugFile2 </c> but does not have a corresponding section in <c> CCDebugFile1</c> </remarks>
    ''' Docs last updated: 2020-07-18 | Code last updated: 2020-07-18
    Public Property PruneStaleEntries As Boolean = True

    '''<summary> Indicates that the debugged file should be saved back to disk <br/> Default: <c> True </c> </summary>
    ''' Docs last updated: 2020-07-18 | Code last updated: 2020-07-18
    Public Property SaveDebuggedFile As Boolean = True

    '''<summary> Indicates that the contents of ccleaner.ini should be sorted alphabetically <br /> Default: <c> True </c> </summary>
    ''' Docs last updated: 2020-07-18 | Code last updated: 2020-07-18
    Public Property SortFileForOutput As Boolean = True

    '''<summary> Indicates that the module's settings have been modified from their defaults </summary>
    ''' Docs last updated: 2020-07-18 | Code last updated: 2020-07-18
    Public Property CCDBSettingsChanged As Boolean = False

    ''' <summary> Handles the commandline args for CCiniDebug </summary>
    '''  CCiniDebug args:
    ''' -noprune    : disable pruning of stale winapp2.ini entries
    ''' -nosort     : disable sorting ccleaner.ini alphabetically
    ''' -nosave     : disable saving the modified ccleaner.ini back to file
    ''' Docs last updated: 2020-07-18 | Code last updated: 2020-07-18
    Public Sub handleCmdlineArgs()
        initDefaultCCDBSettings()
        invertSettingAndRemoveArg(PruneStaleEntries, "-noprune")
        invertSettingAndRemoveArg(SortFileForOutput, "-nosort")
        invertSettingAndRemoveArg(SaveDebuggedFile, "-nosave")
        getFileAndDirParams(CCDebugFile1, CCDebugFile2, CCDebugFile3)
        initCCDebug()
    End Sub

    '''<summary> Prunes, sorts, and/or saves ccleaner.ini </summary>
    ''' Docs last updated: 2020-07-18 | Code last updated: 2020-07-18
    Private Sub ccDebug()
        If PruneStaleEntries Then prune(CCDebugFile2.Sections("Options"))
        If SortFileForOutput Then sortCC()
        CCDebugFile3.overwriteToFile(CCDebugFile2.toString, SaveDebuggedFile)
    End Sub

    ''' <summary> Sets up the debug and prints its results </summary>
    ''' Docs last updated: 2020-07-18 | Code last updated: 2020-07-18
    Public Sub initCCDebug()
        If Not enforceFileHasContent(CCDebugFile2) Then Return
        ' winapp2.ini is not really required to be populated for this task so we do not need to enforce that it has content
        If PruneStaleEntries Then CCDebugFile1.validate()
        clrConsole()
        print(4, "CCiniDebug Results", conjoin:=True)
        gLog("Debugging CCleaner.ini", ascend:=True)
        ccDebug()
        gLog("Debug complete", descend:=True)
        print(0, $"{If(SaveDebuggedFile, $"{CCDebugFile3.Name} saved", "Analysis complete")}. {anyKeyStr}", isCentered:=True, closeMenu:=True)
        crk()
    End Sub

    ''' <summary> Scans for and removes stale winapp2.ini entry settings from given Options section of a ccleaner.ini file </summary>
    ''' <param name="optionsSec"> The Options section from ccleaner.ini </param>
    ''' Docs last updated: 2020-07-18 | Code last updated: 2020-07-18
    Private Sub prune(ByRef optionsSec As iniSection)
        gLog($"Scanning {CCDebugFile2.Name} for settings left over from removed winapp2.ini entries", ascend:=True)
        print(0, $"Scanning {CCDebugFile2.Name} for settings left over from removed winapp2.ini entries", leadingBlank:=True, trailingBlank:=True)
        Dim tbTrimmed As New List(Of Integer)
        For i = 0 To optionsSec.Keys.KeyCount - 1
            Dim optionStr = optionsSec.Keys.Keys(i).toString
            ' Only operate on (app) keys belonging to winapp2.ini, marked with a * 
            If optionStr.StartsWith("(App)", StringComparison.InvariantCulture) AndAlso optionStr.Contains("*") Then
                Dim toRemove As New List(Of String) From {"(App)", "=True", "=False"}
                toRemove.ForEach(Sub(param) optionStr = optionStr.Replace(param, ""))
                If Not CCDebugFile1.hasSection(optionStr) Then
                    tbTrimmed.Add(i)
                    Dim foundStr = $"Orphaned entry detected: {optionStr}"
                    print(0, foundStr, colorLine:=True)
                    gLog(foundStr, indent:=True)
                End If
            End If
        Next
        print(0, $"{tbTrimmed.Count} orphaned settings detected", leadingBlank:=True, trailingBlank:=True)
        gLog($"{tbTrimmed.Count} orphaned settings detected", indent:=True, descend:=True)
        optionsSec.removeKeys(tbTrimmed)
    End Sub

    ''' <summary> Sorts the keys in the Options section of <c> CCiniDebugFile2 </c></summary>
    ''' Docs last updated: 2020-07-18 | Code last updated: 2020-07-18
    Private Sub sortCC()
        gLog($"Sorting {CCDebugFile2.Name}", indent:=True)
        Dim lineList = CCDebugFile2.Sections("Options").getKeysAsStrList
        lineList.Items.Sort()
        lineList.Items.Insert(0, "[Options]")
        CCDebugFile2.Sections("Options") = New iniSection(lineList.Items)
        gLog("Done", indent:=True, ascend:=True, descend:=True)
    End Sub

End Module