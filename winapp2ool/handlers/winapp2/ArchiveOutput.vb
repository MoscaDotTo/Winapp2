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
''' The <c> Archive= </c> key the three generators share. A source entry carrying it is built
''' exactly like a live one, then goes to an archive output, <c> archived.ini </c>, instead of
''' the live one. The key's value is a free-text reason that we never emit, and a bare
''' <c> Archive= </c> archives too. <br /><br />
'''
''' Two forms are refused, each leaving the entry unbuilt and failing the run:
''' <c> Archive=False </c>, since deleting the line is how an entry comes back, and an entry
''' that also says <c> Skip= </c>, since it can't be both.
''' </summary>
Public Module ArchiveOutput

    ''' <summary>
    ''' Returns why an entry's <c> Archive= </c> key is refused, or <c> "" </c> when it isn't
    ''' </summary>
    '''
    ''' <param name="context">
    ''' Where the key came from, such as <c> [Discord *] </c>, put at the start of the message
    ''' </param>
    '''
    ''' <param name="archiveValue">
    ''' The <c> Archive= </c> key's value
    ''' </param>
    '''
    ''' <param name="skipDeclared">
    ''' Indicates whether the same entry declared <c> Skip= </c>
    ''' </param>
    Public Function ArchiveRefusal(context As String,
                                   archiveValue As String,
                                   skipDeclared As Boolean) As String

        If archiveValue.Trim().Equals("False", StringComparison.OrdinalIgnoreCase) Then

            Return $"{context} declares Archive=False; delete the Archive line to keep the entry live. The entry wasn't built"

        End If

        If skipDeclared Then

            Return $"{context} declares both Skip and Archive; keep one of them. The entry wasn't built"

        End If

        Return ""

    End Function

    ''' <summary>
    ''' Warns about a refused <c> Archive= </c> key and fails the run, when
    ''' <see cref="ArchiveRefusal"/> refuses it
    ''' </summary>
    '''
    ''' <param name="context">
    ''' Where the key came from, such as <c> [Discord *] </c>
    ''' </param>
    '''
    ''' <param name="archiveValue">
    ''' The <c> Archive= </c> key's value
    ''' </param>
    '''
    ''' <param name="skipDeclared">
    ''' Indicates whether the same entry declared <c> Skip= </c>
    ''' </param>
    '''
    ''' <param name="menuOutput">
    ''' The <c> MenuSection </c> receiving the warning
    ''' </param>
    '''
    ''' <returns>
    ''' <c> True </c> when the key is refused and the entry must not be built
    ''' </returns>
    Public Function RefuseArchive(context As String,
                                  archiveValue As String,
                                  skipDeclared As Boolean,
                                  menuOutput As MenuSection) As Boolean

        Dim refusal = ArchiveRefusal(context, archiveValue, skipDeclared)
        If refusal.Length = 0 Then

            Return False

        End If

        gLog(refusal)
        menuOutput.AddWarning(refusal)
        markRunFailed()

        Return True

    End Function

    ''' <summary>
    ''' Returns the archive folder line a builder's menu shows: the chooser's <c> Dir </c>, or a
    ''' note that there is none
    ''' </summary>
    '''
    ''' <param name="archiveFolder">
    ''' The builder's archive folder setting
    ''' </param>
    Public Function ArchiveFolderText(archiveFolder As iniFileChooser) As String

        Return If(archiveFolder.Dir.Length = 0, "none (a run with archived entries fails)", archiveFolder.Dir)

    End Function

    ''' <summary>
    ''' Normalizes a generator's archived entries through <see cref="remotedebugGuarded"/> and
    ''' writes them to <c> archived.ini </c> in <paramref name="archiveDir"/> under an
    ''' <see cref="ArtifactHeader"/>, even when there are none, so the artifact set is the same on
    ''' every run. The file is never split. <br /><br />
    '''
    ''' With no archive folder we write nothing. Archived entries then reach neither output, so
    ''' we warn with their count and fail the run.
    ''' </summary>
    '''
    ''' <param name="archiveFile">
    ''' The archived entries, generated exactly as the live ones are. It's consumed, as
    ''' <see cref="remotedebugGuarded"/>'s argument is.
    ''' </param>
    '''
    ''' <param name="archiveDir">
    ''' The folder to write <c> archived.ini </c> to, or <c> "" </c> for none
    ''' </param>
    '''
    ''' <param name="generatorName">
    ''' The module writing the file, such as <c> EntryBuilder </c>
    ''' </param>
    '''
    ''' <param name="sourceFolder">
    ''' The repository-relative folder holding the generator's sources, linked in the header
    ''' </param>
    '''
    ''' <param name="menuOutput">
    ''' The <c> MenuSection </c> receiving progress lines and warnings
    ''' </param>
    Public Sub WriteArchive(archiveFile As iniFile,
                            archiveDir As String,
                            generatorName As String,
                            sourceFolder As String,
                            menuOutput As MenuSection)

        If archiveDir.Length = 0 Then

            If archiveFile.Count = 0 Then

                Return

            End If

            Dim droppedMsg = $"{archiveFile.Count} archived entries left out of the output: no archive folder was given (-4d, or -3d for BrowserBuilder)"
            gLog(droppedMsg)
            menuOutput.AddWarning(droppedMsg)
            markRunFailed()
            Return

        End If

        Dim toWrite = iniFile.Empty(archiveDir, winapp2file.ArchiveFileName)

        For Each section In archiveFile

            toWrite.AddSection(section)

        Next

        toWrite = remotedebugGuarded(toWrite, generatorName, menuOutput)

        Dim header = ArtifactHeader.Text(toWrite.Name, generatorName, toWrite.Count, sourceFolder, sourceIsFolder:=True,
                                         description:="It contains the entries marked Archive= in those sources.")

        If toWrite.OverwriteToFile(header & toWrite.ToString()) Then

            Dim archiveMsg = $"Wrote {toWrite.Count} archived entries to {toWrite.Path()}"
            menuOutput.AddColoredLine(archiveMsg, ConsoleColor.Yellow)
            gLog(archiveMsg)

        End If

    End Sub

End Module
