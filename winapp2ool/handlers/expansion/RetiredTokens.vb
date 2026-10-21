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

Imports System.Text.RegularExpressions

''' <summary>
''' Finds generator tokens written in a <c> %...% </c> spelling the builders no longer accept.
''' <c> %X% </c> is winapp2's environment-variable syntax, so a stale token that reached the
''' output would read as a plausible path. The builders refuse a key holding one instead of
''' passing it through.
''' </summary>
Public Module RetiredTokens

    ''' <summary>
    ''' Returns a message naming the new spelling of the first retired token in
    ''' <paramref name="template"/>. Each stem matches as a case-insensitive prefix, so
    ''' <c> %Package </c> catches <c> %Package% </c>, <c> %package2% </c>, and
    ''' <c> %Package\LocalCache </c> with its closing <c> % </c> missing.
    ''' <br /><br />
    '''
    ''' A token with nothing after the stem is now spelled <c> &lt;Name&gt; </c>, and one with
    ''' only digits after it (<c> %Package2% </c>) is now <c> &lt;Name[N]&gt; </c>, whether or
    ''' not its closing <c> % </c> is there. A token with other letters after the stem gets
    ''' <c> &lt;Name&gt; </c> as a hint. The new spelling takes the stem's capitalization, not
    ''' the template's.
    ''' </summary>
    '''
    ''' <param name="template">
    ''' The key value to scan
    ''' </param>
    '''
    ''' <param name="retiredStems">
    ''' The retired token stems, each written with its leading <c> % </c> and no closing one,
    ''' e.g. <c> %Package </c>
    ''' </param>
    '''
    ''' <param name="context">
    ''' Where the template came from (e.g. <c> [Cortana *].FileKey1 </c>), embedded in the message
    ''' </param>
    '''
    ''' <returns>
    ''' The message for the first stem that matches, or an empty string when none does
    ''' </returns>
    Public Function FindRetiredToken(template As String,
                                     retiredStems As IEnumerable(Of String),
                                     context As String) As String

        For Each stem In retiredStems

            Dim name = stem.TrimStart("%"c)
            Dim m = Regex.Match(template, $"%{Regex.Escape(name)}([^%\\|;]*)(%?)", RegexOptions.IgnoreCase Or RegexOptions.CultureInvariant)

            If Not m.Success Then Continue For

            Dim suffix = m.Groups(1).Value
            Dim closed = m.Groups(2).Value.Length > 0

            Dim numbered = Regex.IsMatch(suffix, "^[0-9]+$")

            If suffix.Length > 0 AndAlso Not numbered Then Return $"`{m.Value}` in {context} looks like the retired `%{name}%` token, which is now spelled `<{name}>`"

            Dim newSpelling = If(numbered, $"<{name}[{suffix}]>", $"<{name}>")

            If closed Then Return $"`{m.Value}` in {context} is now spelled `{newSpelling}`"

            Return $"`{m.Value}` in {context} looks like the retired `%{name}{suffix}%` token, which is now spelled `{newSpelling}`"

        Next

        Return ""

    End Function

End Module
