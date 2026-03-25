'    Copyright (C) 2018-2026 Hazel Ward
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
''' Holds the settings for the CCiniDebug module, which is responsible for debugging ccleaner.ini files.
''' </summary>
Public Module ccdebugsettings

    ''' <summary>
    ''' The winapp2.ini file that ccleaner.ini may optionally be checked against
    ''' <br /> Default: <c> winapp2.ini </c>
    ''' </summary>
    Public Property CCDebugFile1 As iniFileChooser = New iniFileChooser(Environment.CurrentDirectory, "winapp2.ini", mustExist:=False)

    ''' <summary>
    ''' The ccleaner.ini file to be debugged
    ''' <br /> Default: ccleaner.ini
    ''' </summary>
    Public Property CCDebugFile2 As iniFileChooser = New iniFileChooser(Environment.CurrentDirectory, "ccleaner.ini")

    ''' <summary>
    ''' Holds the path for the debugged file that will be saved to disk.
    ''' Defaults to <c> ccleaner-debugged.ini </c> in the current directory
    ''' <br /> Default: <c> ccleaner.ini </c>
    ''' <br /> Default rename: <c> ccleaner-debugged.ini </c>
    ''' </summary>
    Public Property CCDebugFile3 As iniFileChooser = New iniFileChooser(Environment.CurrentDirectory, "ccleaner.ini", "ccleaner-debugged.ini", mustExist:=False)

    ''' <summary>
    ''' Indicates that stale winapp2.ini entries should be pruned from ccleaner.ini
    ''' <br/> Default: <c> True </c>
    ''' </summary>
    '''
    ''' <remarks>
    ''' A "stale" entry is one that appears in a winapp2.ini (app) key in <c> CCDebugFile2 </c>
    ''' but does not have a corresponding section in <c> CCDebugFile1 </c>
    ''' </remarks>
    Public Property PruneStaleEntries As Boolean = True

    ''' <summary>
    ''' Indicates that the debugged file should be saved back to disk
    ''' <br/> Default: <c> True </c>
    ''' </summary>
    Public Property SaveDebuggedFile As Boolean = True

    ''' <summary>
    ''' Indicates that the contents of ccleaner.ini should be sorted alphabetically
    ''' <br /> Default: <c> True </c>
    ''' </summary>
    Public Property SortFileForOutput As Boolean = True

    ''' <summary>
    ''' Indicates that the module's settings have been modified from their defaults
    ''' </summary>
    Public Property CCDBSettingsChanged As Boolean = False

    ''' <summary>
    ''' Resets the CCiniDebug module's settings to their defaults and persists them via <c> SaveModule2 </c>.
    ''' </summary>
    Public Sub initDefaultCCDBSettings()

        CCDebugFile1.ResetParams()
        CCDebugFile2.ResetParams()
        CCDebugFile3.ResetParams()
        PruneStaleEntries = True
        SaveDebuggedFile = True
        SortFileForOutput = True
        CCDBSettingsChanged = False
        SaveModule2(NameOf(CCiniDebug), GetType(ccdebugsettings))

    End Sub

End Module
