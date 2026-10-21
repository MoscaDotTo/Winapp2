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
    ''' <paramref name="template"/>, for tokens that kept their name and changed only their
    ''' spelling (<c> %Package% </c> is now <c> &lt;Package&gt; </c>). The matching rules are
    ''' those of the overload taking renames.
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

        Dim renames = retiredStems.Select(Function(stem) New KeyValuePair(Of String, String)(stem, stem.TrimStart("%"c)))
        Return FindRetiredToken(template, renames, context)

    End Function

    ''' <summary>
    ''' Returns a message naming the new spelling of the first retired token in
    ''' <paramref name="template"/>. Each stem matches as a case-insensitive prefix, so
    ''' <c> %Package </c> catches <c> %Package% </c>, <c> %package2% </c>, and
    ''' <c> %Package\LocalCache </c> with its closing <c> % </c> missing.
    ''' <br /><br />
    '''
    ''' A token with nothing after the stem is now spelled <c> &lt;NewName&gt; </c>, and one
    ''' with only digits after it (<c> %Package2% </c>) is now <c> &lt;NewName[N]&gt; </c>,
    ''' whether or not its closing <c> % </c> is there. A token with other letters after the
    ''' stem gets <c> &lt;NewName&gt; </c> as a hint, as does a numbered one when
    ''' <paramref name="indexable"/> is <c> False </c>. The new spelling takes the new name's
    ''' capitalization, not the template's.
    ''' </summary>
    '''
    ''' <param name="template">
    ''' The key value to scan
    ''' </param>
    '''
    ''' <param name="renames">
    ''' Each retired stem, written with its leading <c> % </c> and no closing one, paired with
    ''' the name its token is now spelled with (e.g. <c> %BrowserPath </c> and
    ''' <c> UserDataParent </c>). Stems are tried in order.
    ''' </param>
    '''
    ''' <param name="context">
    ''' Where the template came from (e.g. <c> [Cortana *].FileKey1 </c>), embedded in the message
    ''' </param>
    '''
    '''
    ''' <param name="indexable">
    ''' Indicates whether the new tokens take an index, so <c> %Package2% </c> can be told
    ''' <c> &lt;Package[2]&gt; </c> <br /><br />
    ''' Optional, Default: <c> True </c>
    ''' </param>
    '''
    ''' <returns>
    ''' The message for the first stem that matches, or an empty string when none does
    ''' </returns>
    Public Function FindRetiredToken(template As String,
                                     renames As IEnumerable(Of KeyValuePair(Of String, String)),
                                     context As String,
                            Optional indexable As Boolean = True) As String

        For Each pair In renames

            Dim name = pair.Key.TrimStart("%"c)
            Dim newName = pair.Value
            Dim m = Regex.Match(template, $"%{Regex.Escape(name)}([^%\\|;]*)(%?)", RegexOptions.IgnoreCase Or RegexOptions.CultureInvariant)

            If Not m.Success Then Continue For

            Dim suffix = m.Groups(1).Value
            Dim closed = m.Groups(2).Value.Length > 0

            Dim numbered = indexable AndAlso Regex.IsMatch(suffix, "^[0-9]+$")

            If suffix.Length > 0 AndAlso Not numbered Then Return $"`{m.Value}` in {context} looks like the retired `%{name}%` token, which is now spelled `<{newName}>`"

            Dim newSpelling = If(numbered, $"<{newName}[{suffix}]>", $"<{newName}>")

            If closed Then Return $"`{m.Value}` in {context} is now spelled `{newSpelling}`"

            Return $"`{m.Value}` in {context} looks like the retired `%{name}{suffix}%` token, which is now spelled `{newSpelling}`"

        Next

        Return ""

    End Function

End Module
