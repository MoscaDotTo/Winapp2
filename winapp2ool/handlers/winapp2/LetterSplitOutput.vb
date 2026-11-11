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

Imports System.Globalization

''' <summary>
''' Writes a generator's output as 27 per-letter artifact files (<c> #.ini </c>,
''' <c> A.ini </c> ... <c> Z.ini </c>), each under an <see cref="ArtifactHeader"/> that links
''' the one source file of the same letter. EntryBuilder and UWPBuilder use it for their
''' <c> -split </c> output.
''' </summary>
Public Module LetterSplitOutput

    ''' <summary>
    ''' The 27 bucket names in file order: <c> # </c>, then <c> A </c> to <c> Z </c>
    ''' </summary>
    Public ReadOnly Property Letters As IReadOnlyList(Of String) =
        {"#"}.Concat(Enumerable.Range(AscW("A"c), 26).Select(Function(c) ChrW(c).ToString(CultureInfo.InvariantCulture))).ToList()

    ''' <summary>
    ''' Classifies an entry name into its per-letter artifact bucket: the first character,
    ''' uppercased with the invariant culture, when that gives <c> A </c> to <c> Z </c>, and
    ''' otherwise <c> # </c> (digits, punctuation, any other character, and an empty or
    ''' <c> Nothing </c> name). Bucketing is by entry name rather than by source file so an
    ''' entry filed in the wrong source letter still lands in its canonical artifact.
    ''' </summary>
    '''
    ''' <param name="entryName">
    ''' The entry name (section header text) to classify
    ''' </param>
    '''
    ''' <returns>
    ''' The single-character bucket name: <c> A </c>–<c> Z </c> or <c> # </c>
    ''' </returns>
    Public Function LetterBucketFor(entryName As String) As String

        If String.IsNullOrEmpty(entryName) Then

            Return "#"

        End If

        Dim c = Char.ToUpperInvariant(entryName(0))

        Return If(c >= "A"c AndAlso c <= "Z"c, c.ToString(), "#")

    End Function

    ''' <summary>
    ''' Writes <paramref name="outputFile"/>'s sections into the 27 per-letter files in
    ''' <paramref name="targetDir"/>, bucketed by <see cref="LetterBucketFor"/>. We write all 27
    ''' on every run, a letter with no entries getting a header-only file, so the artifact set
    ''' is stable across runs and deleting a letter's last entry can't leave a stale file
    ''' behind. The summary line reports how many writes succeeded and turns red when any
    ''' failed.
    ''' </summary>
    '''
    ''' <param name="outputFile">
    ''' The fully generated, normalized output whose sections will be bucketed
    ''' </param>
    '''
    ''' <param name="targetDir">
    ''' The folder the 27 files are written to
    ''' </param>
    '''
    ''' <param name="generatorName">
    ''' The module named in each file's header, such as <c> EntryBuilder </c>
    ''' </param>
    '''
    ''' <param name="sourceFolder">
    ''' The repository-relative folder holding one source file per letter, such as
    ''' <see cref="ArtifactHeader.EntryBuilderSources"/>. Each file's header links the source
    ''' file of its own letter there.
    ''' </param>
    '''
    ''' <param name="menuOutput">
    ''' The <c> MenuSection </c> receiving the summary line
    ''' </param>
    Public Sub WriteSplitOutput(outputFile As iniFile,
                                targetDir As String,
                                generatorName As String,
                                sourceFolder As String,
                                menuOutput As MenuSection)

        Dim buckets = Letters.ToDictionary(Function(l) l, Function(l) iniFile.Empty(targetDir, l & ".ini"))

        For Each section In outputFile

            buckets(LetterBucketFor(section.Name)).AddSection(section)

        Next

        Dim written = 0

        For Each letter In Letters

            Dim bucket = buckets(letter)
            Dim letterDesc = If(letter = "#", "symbols or numbers", $"the letter {letter}")

            Dim header = ArtifactHeader.Text(bucket.Name, generatorName, bucket.Count,
                                             $"{sourceFolder}/{letter}.ini", sourceIsFolder:=False,
                                             description:=$"It contains entries beginning with {letterDesc}.")

            If bucket.OverwriteToFile(header & bucket.ToString()) Then

                written += 1

            End If

        Next

        Dim splitMsg = $"Wrote {written} of {Letters.Count} per-letter artifact files to {targetDir}"
        menuOutput.AddColoredLine(splitMsg, If(written = Letters.Count, ConsoleColor.Green, ConsoleColor.Red))
        gLog(splitMsg)

    End Sub

End Module
