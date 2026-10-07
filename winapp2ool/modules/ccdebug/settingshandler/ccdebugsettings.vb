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
    ''' The winapp2.ini file that ccleaner.ini is checked against when <c> PruneStaleEntries </c>
    ''' is <c> True </c>
    ''' <br /> Default: <c> winapp2.ini </c>
    ''' </summary>
    Public Property CCDebugFile1 As iniFileChooser = New iniFileChooser(Environment.CurrentDirectory, "winapp2.ini", mustExist:=False)

    ''' <summary>
    ''' The ccleaner.ini file to be debugged
    ''' <br /> Default: <c> ccleaner.ini </c>
    ''' </summary>
    Public Property CCDebugFile2 As iniFileChooser = New iniFileChooser(Environment.CurrentDirectory, "ccleaner.ini")

    ''' <summary>
    ''' Holds the path for the debugged file that will be saved to disk. By default this
    ''' overwrites <c> ccleaner.ini </c> in the current directory.
    ''' <br /> Default: <c> ccleaner.ini </c>
    ''' <br /> Default rename: <c> ccleaner-debugged.ini </c>
    ''' </summary>
    Public Property CCDebugFile3 As iniFileChooser = New iniFileChooser(Environment.CurrentDirectory, "ccleaner.ini", "ccleaner-debugged.ini", mustExist:=False)

    ''' <summary>
    ''' Indicates whether stale winapp2.ini entries should be pruned from ccleaner.ini
    ''' <br /> Default: <c> True </c>
    ''' </summary>
    '''
    ''' <remarks>
    ''' A "stale" entry is one named by an <c> (App) </c> key containing <c> * </c> in the
    ''' <c> [Options] </c> section of <c> CCDebugFile2 </c> that has no corresponding section
    ''' in <c> CCDebugFile1 </c>
    ''' </remarks>
    Public Property PruneStaleEntries As Boolean = True

    ''' <summary>
    ''' Indicates whether the debugged file should be saved back to disk
    ''' <br /> Default: <c> True </c>
    ''' </summary>
    Public Property SaveDebuggedFile As Boolean = True

    ''' <summary>
    ''' Indicates whether the keys of the <c> [Options] </c> section of ccleaner.ini should be
    ''' sorted by name. Other sections keep their order.
    ''' <br /> Default: <c> True </c>
    ''' </summary>
    Public Property SortFileForOutput As Boolean = True

    ''' <summary>
    ''' Indicates whether the module's settings have been modified from their defaults
    ''' </summary>
    Public Property CCDBSettingsChanged As Boolean = False

    ''' <summary>
    ''' Resets the CCiniDebug module's settings to their defaults and records them in the settings
    ''' file through <see cref="SaveModule"/>. Nothing is written to disk here.
    ''' <see cref="FlushIfDirty"/> does that later, if the save gate allows it
    ''' </summary>
    Public Sub initDefaultCCDBSettings()

        CCDebugFile1.ResetParams()
        CCDebugFile2.ResetParams()
        CCDebugFile3.ResetParams()
        PruneStaleEntries = True
        SaveDebuggedFile = True
        SortFileForOutput = True
        CCDBSettingsChanged = False
        SaveModule(NameOf(CCiniDebug), GetType(ccdebugsettings))

    End Sub

End Module
