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
''' CCiniDebug is a winapp2ool module which performs housekeeping on the CCleaner configuration file
''' ccleaner.ini to clean up leftovers from winapp2.ini 
''' <br /> 
''' As entries are removed or renamed in winapp2.ini over time, stale configuration keys are 
''' leftover in ccleaner.ini. CCiniDebug processes a ccleaner.ini for its entry configuration keys
''' and checks them each against the names of entries in winapp2.ini. Any configuration keys found 
''' with the winapp2.ini indicator (*) in the name which do not have a corresponding entry by the 
''' same name in the most recent winapp2.ini can then be easily and automatically removed 
''' </summary>
''' 
''' Docs last updated: 2025-08-05 | Code last updated: 2025-08-05
Module CCiniDebug

    ''' <summary>
    ''' Handles the commandline args for CCiniDebug 
    ''' </summary>
    ''' 
    ''' <remarks>
    ''' CCiniDebug setting toggles:
    ''' -noprune    : disable pruning of stale winapp2.ini entries
    ''' -nosort     : disable sorting ccleaner.ini alphabetically
    ''' -nosave     : disable saving the modified ccleaner.ini back to file
    ''' </remarks>
    ''' 
    ''' Docs last updated: 2020-07-18 | Code last updated: 2020-07-18
    Public Sub handleCmdlineArgs()

        initDefaultCCDBSettings()
        invertSettingAndRemoveArg(PruneStaleEntries, "-noprune")
        invertSettingAndRemoveArg(SortFileForOutput, "-nosort")
        invertSettingAndRemoveArg(SaveDebuggedFile, "-nosave")

        getFileAndDirParams(CCDebugFile1, CCDebugFile2, CCDebugFile3)

        initCCDebug()

    End Sub

    '''<summary> 
    '''Prunes, sorts, and/or saves ccleaner.ini 
    '''</summary>
    '''
    ''' Docs last updated: 2020-07-18 | Code last updated: 2020-07-18
    Private Sub ccDebug()

        If PruneStaleEntries Then prune(CCDebugFile2.Sections("Options"))
        If SortFileForOutput Then sortCC()
        CCDebugFile3.overwriteToFile(CCDebugFile2.toString, SaveDebuggedFile)

    End Sub

    ''' <summary>
    ''' Sets up the debug and prints its results 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2020-07-18 | Code last updated: 2025-08-05
    Public Sub initCCDebug()

        If Not enforceFileHasContent(CCDebugFile2) Then Return

        If PruneStaleEntries Then CCDebugFile1.validate()

        clrConsole()
        LogAndPrint(4, $"Analyzing {CCDebugFile2.Name}", conjoin:=True, ascend:=True)

        ccDebug()

        Dim out = $"Analysis complete{If(SaveDebuggedFile, $"{CCDebugFile3.Name} saved", "")}. {anyKeyStr}"

        LogAndPrint(0, out, isCentered:=True, closeMenu:=True, descend:=True)

        crk()

    End Sub

    ''' <summary> 
    ''' Scans for and removes stale winapp2.ini entry settings 
    ''' from a given Options section of a ccleaner.ini file 
    ''' </summary>
    ''' 
    ''' <param name="optionsSec"> 
    ''' The Options section from ccleaner.ini 
    ''' </param>
    ''' 
    ''' Docs last updated: 2020-07-18 | Code last updated: 2025-08-05
    Private Sub prune(ByRef optionsSec As iniSection)

        Dim logStr = $"Scanning {CCDebugFile2.Name} for settings left over from removed winapp2.ini entries"
        LogAndPrint(0, logStr, ascend:=True, leadingBlank:=True, trailingBlank:=True)

        Dim tbTrimmed As New List(Of Integer)
        For i = 0 To optionsSec.Keys.KeyCount - 1

            Dim optionStr = optionsSec.Keys.Keys(i).toString

            If Not optionStr.StartsWith("(App)", StringComparison.InvariantCulture) OrElse Not optionStr.Contains("*") Then Continue For

            Dim toRemove As New List(Of String) From {"(App)", "=True", "=False"}
            toRemove.ForEach(Sub(param) optionStr = optionStr.Replace(param, ""))

            If CCDebugFile1.hasSection(optionStr) Then Continue For

            tbTrimmed.Add(i)
            logStr = $"Orphaned entry detected: {optionStr}"
            LogAndPrint(0, logStr, colorLine:=True, indent:=True)

        Next

        logStr = $"{tbTrimmed.Count} orphaned settings detected"
        LogAndPrint(0, logStr, leadingBlank:=True, trailingBlank:=True, indent:=True, descend:=True)

        optionsSec.removeKeys(tbTrimmed)

    End Sub

    ''' <summary> 
    ''' Sorts the keys in the Options section of <c> CCiniDebugFile2 </c>
    ''' </summary>
    ''' 
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
