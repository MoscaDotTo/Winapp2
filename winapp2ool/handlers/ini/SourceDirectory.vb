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
Imports System.IO

''' <summary>
''' Reads a folder of per-letter source files into one <c> iniFile </c>, for the generators
''' whose sources are split that way (EntryBuilder and UWPBuilder's <c> AppInfo\ </c>)
''' </summary>
Public Module SourceDirectory

    ''' <summary>
    ''' Combines the top-level <c> *.ini </c> files in <paramref name="dir"/> into a single
    ''' in-memory <c> iniFile </c>, reading them in sorted path order. The first definition of a
    ''' section name (case-insensitive) wins. A repeat, in a later file or the same one, is
    ''' dropped with a warning naming its file and line and the file whose definition is kept,
    ''' and fails the run.
    ''' </summary>
    '''
    ''' <param name="dir">
    ''' Path to the directory containing the source files
    ''' </param>
    '''
    ''' <param name="sectionKind">
    ''' What the warning calls a section, such as <c> AppInfo </c>
    ''' </param>
    '''
    ''' <param name="menuOutput">
    ''' The <c> MenuSection </c> receiving the duplicate-section warnings
    ''' </param>
    '''
    ''' <returns>
    ''' The merged <c> iniFile </c>, or an empty one if the directory does not exist
    ''' or contains no parseable sections
    ''' </returns>
    Public Function Combine(dir As String, sectionKind As String, menuOutput As MenuSection) As iniFile

        Dim combined = iniFile.Empty("", "")

        If menuOutput Is Nothing Then argIsNull(NameOf(menuOutput)) : Return combined
        If Not Directory.Exists(dir) Then Return combined

        Dim files = Directory.GetFiles(dir, "*.ini", SearchOption.TopDirectoryOnly).ToList()
        files.Sort()

        Dim definedIn As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

        For Each filePath In files

            Dim f = iniFile.FromFile(filePath)

            ' The parser keeps a file's first occurrence of each name, so it reaches this loop
            ' ahead of that file's own repeats.
            For Each section In f.Concat(f.DuplicateSections)

                If definedIn.ContainsKey(section.Name) Then

                    Dim dupeMsg = $"Duplicate {sectionKind} section [{section.Name}] at {f.Name} line {section.StartingLineNumber} ignored; keeping the first definition, in {definedIn(section.Name)}"
                    gLog(dupeMsg)
                    menuOutput.AddWarning(dupeMsg)
                    markRunFailed()

                Else

                    combined.AddSection(section)
                    definedIn.Add(section.Name, f.Name)

                End If

            Next

        Next

        Return combined

    End Function

End Module
