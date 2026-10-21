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
''' Generates one browser's entry from an EntryScaffold's templates. The templates use a closed
''' set of four tokens, matched in any case: <c> &lt;UserDataPath&gt; </c>,
''' <c> &lt;UserDataParent&gt; </c> and <c> &lt;LocalUserDataPath&gt; </c> in a
''' <c> FileKeyBase= </c>, and <c> &lt;RegistryRoot&gt; </c> in a <c> RegKeyBase= </c>. After
''' we substitute them, each value goes through <see cref="VariableExpander.Expand"/> with no
''' declared variables, so an inline list such as <c> &lt;Cache,Code Cache&gt; </c> fans out.
''' <see cref="findTemplateError"/> refuses a template that breaks these rules before any
''' browser sees it.
''' </summary>
Friend Module BrowserEntryGenerator

    ''' <summary> The name of the token for a browser's user data path </summary>
    Private Const UserDataPathName As String = "UserDataPath"

    ''' <summary> The name of the token for the folder one level above the user data path </summary>
    Private Const UserDataParentName As String = "UserDataParent"

    ''' <summary> The name of the Gecko-only token for the <c> %LocalAppData% </c> copy of the user data path </summary>
    Private Const LocalUserDataPathName As String = "LocalUserDataPath"

    ''' <summary> The name of the token for a browser's registry root </summary>
    Private Const RegistryRootName As String = "RegistryRoot"

    ''' <summary> The <c> &lt;UserDataPath&gt; </c> token </summary>
    Private Const UserDataPathToken As String = "<" & UserDataPathName & ">"

    ''' <summary> The <c> &lt;UserDataParent&gt; </c> token </summary>
    Private Const UserDataParentToken As String = "<" & UserDataParentName & ">"

    ''' <summary> The <c> &lt;LocalUserDataPath&gt; </c> token </summary>
    Private Const LocalUserDataPathToken As String = "<" & LocalUserDataPathName & ">"

    ''' <summary> The <c> &lt;RegistryRoot&gt; </c> token </summary>
    Private Const RegistryRootToken As String = "<" & RegistryRootName & ">"

    ''' <summary>
    ''' The retired <c> %...% </c> spellings of the four tokens, as stems for
    ''' <see cref="RetiredTokens.FindRetiredToken"/>, each paired with the name it is now
    ''' spelled with. <c> %BrowserPath% </c> and <c> %LocalDataPath% </c> were renamed as well
    ''' as respelled. The tokens take no index, so we pass <c> indexable:=False </c>.
    ''' </summary>
    Private ReadOnly RetiredBrowserTokens As KeyValuePair(Of String, String)() = {
        New KeyValuePair(Of String, String)("%UserDataPath", UserDataPathName),
        New KeyValuePair(Of String, String)("%BrowserPath", UserDataParentName),
        New KeyValuePair(Of String, String)("%LocalDataPath", LocalUserDataPathName),
        New KeyValuePair(Of String, String)("%RegistryRoot", RegistryRootName)
    }

    ''' <summary> The names of the three file tokens, in the order <see cref="processUserDataPaths"/> tries them </summary>
    Private ReadOnly FileTokenNames As String() = {UserDataPathName, UserDataParentName, LocalUserDataPathName}

    ''' <summary> Matches one <c> &lt;name&gt; </c> token, as <see cref="VariableExpander"/> does </summary>
    Private ReadOnly TokenPattern As New Regex("<([^<>]+)>", RegexOptions.Compiled)

    ''' <summary>
    ''' Returns why <paramref name="template"/> can't be used, or an empty string when it can.
    ''' A template is refused when it:
    ''' <list type="bullet">
    '''
    ''' <item> holds a retired <c> %...% </c> token (see <see cref="RetiredBrowserTokens"/>) </item>
    '''
    ''' <item> holds an empty, stray or nested <c> &lt; </c> or <c> &gt; </c> </item>
    '''
    ''' <item> holds a <c> &lt;name&gt; </c> token that isn't one of the four and isn't an
    ''' inline list, or an inline list with no non-empty value (<c> &lt;,&gt; </c>). A list
    ''' with at least one value, such as <c> &lt;a,&gt; </c>, passes, and the expander drops
    ''' its empty value. </item>
    '''
    ''' <item> is a <c> RegKeyBase= </c> with a file token, or without <c> &lt;RegistryRoot&gt; </c> </item>
    '''
    ''' <item> is a <c> FileKeyBase= </c> with <c> &lt;RegistryRoot&gt; </c>, with no file
    ''' token, or with two kinds of file token </item>
    '''
    ''' <item> is a Chromium <c> FileKeyBase= </c> with <c> &lt;LocalUserDataPath&gt; </c> </item>
    '''
    ''' </list>
    ''' Each of these would otherwise produce nothing, or ship a token as literal text.
    ''' </summary>
    '''
    ''' <param name="template">
    ''' The <c> FileKeyBase= </c> or <c> RegKeyBase= </c> value
    ''' </param>
    '''
    ''' <param name="isRegKey">
    ''' Indicates whether <paramref name="template"/> is a <c> RegKeyBase= </c>
    ''' </param>
    '''
    ''' <param name="isGecko">
    ''' Indicates whether the template's scaffold belongs to the Gecko ruleset
    ''' </param>
    '''
    ''' <param name="context">
    ''' Where the template came from (e.g. <c> [EntryScaffold: Caches].FileKeyBase </c>),
    ''' embedded in the message
    ''' </param>
    '''
    ''' <returns>
    ''' The reason the template is refused, or an empty string
    ''' </returns>
    Friend Function findTemplateError(template As String,
                                      isRegKey As Boolean,
                                      isGecko As Boolean,
                                      context As String) As String

        Dim retiredMsg = RetiredTokens.FindRetiredToken(template, RetiredBrowserTokens, context, indexable:=False)
        If retiredMsg.Length > 0 Then Return retiredMsg

        ' A bracket left over once every token is gone is either empty, stray or nested, and
        ' would otherwise ship as text
        Dim residue = TokenPattern.Replace(template, "")
        If residue.IndexOfAny({"<"c, ">"c}) >= 0 Then Return $"Malformed `<...>` token in {context}; a token is one <Name> or one <a,b> list, not empty or nested"

        Dim fileTokens As New List(Of String)
        Dim hasRegistryRoot = False

        For Each m As Match In TokenPattern.Matches(template)

            Dim body = m.Groups(1).Value

            If body.IndexOf(","c) >= 0 Then

                If body.Split(","c).All(Function(value) value.Trim().Length = 0) Then Return $"Inline list `<{body}>` in {context} has no values"
                Continue For

            End If

            If body.Equals(RegistryRootName, StringComparison.OrdinalIgnoreCase) Then

                hasRegistryRoot = True
                Continue For

            End If

            Dim fileToken = FileTokenNames.FirstOrDefault(Function(name) name.Equals(body, StringComparison.OrdinalIgnoreCase))

            If fileToken Is Nothing Then Return $"Unknown token `<{body}>` in {context}; BrowserBuilder recognizes {UserDataPathToken}, {UserDataParentToken}, {LocalUserDataPathToken} and {RegistryRootToken}"

            If Not fileTokens.Contains(fileToken) Then fileTokens.Add(fileToken)

        Next

        If isRegKey Then

            If fileTokens.Count > 0 Then Return $"{context} uses the file token `<{fileTokens(0)}>`; a RegKeyBase takes only {RegistryRootToken}"
            If Not hasRegistryRoot Then Return $"{context} has no {RegistryRootToken} token"
            Return ""

        End If

        If hasRegistryRoot Then Return $"{context} uses {RegistryRootToken}, which only a RegKeyBase takes"
        If fileTokens.Count = 0 Then Return $"{context} has no file token; a FileKeyBase needs one of {UserDataPathToken}, {UserDataParentToken} or {LocalUserDataPathToken}"
        If fileTokens.Count > 1 Then Return $"{context} mixes `<{fileTokens(0)}>` and `<{fileTokens(1)}>`; a FileKeyBase takes one kind of file token"
        If Not isGecko AndAlso fileTokens(0) = LocalUserDataPathName Then Return $"{context} uses {LocalUserDataPathToken}, which only Gecko browsers have"

        Return ""

    End Function

    ''' <summary>
    ''' Returns whether <paramref name="value"/> holds <paramref name="token"/>, in any case
    ''' </summary>
    '''
    ''' <param name="value">
    ''' The template to search
    ''' </param>
    '''
    ''' <param name="token">
    ''' The token to look for
    ''' </param>
    Private Function hasToken(value As String, token As String) As Boolean

        Return value.IndexOf(token, StringComparison.OrdinalIgnoreCase) >= 0

    End Function

    ''' <summary>
    ''' Replaces every occurrence of <paramref name="token"/> in <paramref name="value"/>, in
    ''' any case, with <paramref name="replacement"/>. The replacement is inserted as written,
    ''' so a <c> $ </c> in it is never read as a substitution pattern.
    ''' </summary>
    '''
    ''' <param name="value">
    ''' The template to substitute into
    ''' </param>
    '''
    ''' <param name="token">
    ''' The token to replace
    ''' </param>
    '''
    ''' <param name="replacement">
    ''' The text to put in its place
    ''' </param>
    Private Function replaceToken(value As String, token As String, replacement As String) As String

        Return Regex.Replace(value, Regex.Escape(token), Function(m) replacement, RegexOptions.IgnoreCase Or RegexOptions.CultureInvariant)

    End Function

    ''' <summary>
    ''' Returns the FileKey values for the Opera GX browser, transforming the EntryScaffold
    ''' template in accordance with the non-standard way Opera GX stores itself on the disk by
    ''' default. Each supported template gives two values: one for the user data folder itself
    ''' and one for its <c> _side_profiles\* </c> folders. Applies only when
    ''' <paramref name="browserName"/> is exactly <c> Opera GX </c>. The prefixes match in any
    ''' case. A template that doesn't start with <c> &lt;UserDataPath&gt;\ </c> or
    ''' <c> &lt;UserDataPath&gt;| </c> is logged as unsupported, without a menu warning.
    ''' </summary>
    '''
    ''' <param name="browserName">
    ''' The name of the browser being processed
    ''' </param>
    '''
    ''' <param name="fileKeyBase">
    ''' A FileKey template to be transformed such that it is appropriate for Opera GX
    ''' </param>
    '''
    ''' <param name="userDataPath">
    ''' The resolved <c> &lt;UserDataPath&gt; </c> value for the current browser, provided by
    ''' its <c> BrowserInfo </c> section
    ''' </param>
    '''
    ''' <returns>
    ''' The two values when the current browser is Opera GX and the template is supported;
    ''' <c> Nothing </c> otherwise, in which case the caller substitutes the template normally
    ''' </returns>
    Private Function operaGXValues(browserName As String,
                                   fileKeyBase As String,
                                   userDataPath As String) As List(Of String)

        If Not browserName.Equals("Opera GX") Then Return Nothing

        Dim baseKeyValue As String
        Dim sideProfileKeyValue As String

        Dim wildcardPipe = UserDataPathToken & "\*|"
        Dim wildcardFolder = UserDataPathToken & "\*\"
        Dim pipe = UserDataPathToken & "|"
        Dim folder = UserDataPathToken & "\"

        ' Opera GX stores files both in the main user data directory (without a folder such as Default)
        ' and any profiles created in _side_profiles\* subdirectories instead of the usual Profile1, etc.
        ' Accordingly, we transform our scaffold templates to accommodate this
        ' <UserDataPath>\*|<params> -> <UserDataPath>|<params> + <UserDataPath>\_side_profiles\*|<params>
        ' <UserDataPath>\*\<directory> -> <UserDataPath>\<directory> + <UserDataPath>\_side_profiles\*\<directory>
        ' <UserDataPath>|<params> -> <UserDataPath>|<params> + <UserDataPath>\_side_profiles\*|<params>
        ' <UserDataPath>\<directory> -> <UserDataPath>\<directory> + <UserDataPath>\_side_profiles\*\<directory>

        Select Case True

            ' <UserDataPath>\*|<params> -> remove \* + create side profiles variant
            Case fileKeyBase.StartsWith(wildcardPipe, StringComparison.OrdinalIgnoreCase)

                Dim afterWildcardPipe As String = fileKeyBase.Substring(wildcardPipe.Length)
                baseKeyValue = userDataPath & "|" & afterWildcardPipe
                sideProfileKeyValue = userDataPath & "\_side_profiles\*|" & afterWildcardPipe

            ' <UserDataPath>\*\<directory> -> unnest one level + create side profiles variant
            Case fileKeyBase.StartsWith(wildcardFolder, StringComparison.OrdinalIgnoreCase)

                Dim afterWildcard As String = fileKeyBase.Substring(wildcardFolder.Length)
                baseKeyValue = userDataPath & "\" & afterWildcard
                sideProfileKeyValue = userDataPath & "\_side_profiles\*\" & afterWildcard

            ' <UserDataPath>|<params> -> keep + create side profiles variant
            Case fileKeyBase.StartsWith(pipe, StringComparison.OrdinalIgnoreCase)

                Dim afterPipe As String = fileKeyBase.Substring(pipe.Length)
                baseKeyValue = userDataPath & "|" & afterPipe
                sideProfileKeyValue = userDataPath & "\_side_profiles\*|" & afterPipe

            ' <UserDataPath>\<directory> -> (non-wildcard subdirectory - RETAINS wildcards in directory names)
            Case fileKeyBase.StartsWith(folder, StringComparison.OrdinalIgnoreCase)

                Dim afterBackslash As String = fileKeyBase.Substring(folder.Length)
                baseKeyValue = userDataPath & "\" & afterBackslash
                sideProfileKeyValue = userDataPath & "\_side_profiles\*\" & afterBackslash

            Case Else

                Dim errMsg = $"Unsupported OperaGX path provided in {fileKeyBase}"
                gLog(errMsg)
                Return Nothing

        End Select

        Return New List(Of String) From {baseKeyValue, sideProfileKeyValue}

    End Function

    ''' <summary>
    ''' Generates the entry <c> {browser} {scaffold} * </c> and adds it to
    ''' <paramref name="outputFile"/>: a <c> Section= </c> key, the DetectFile and FileKey keys
    ''' from <see cref="processUserDataPaths"/>, then the RegKeys. If the file already holds an
    ''' entry of that name, the new one is silently dropped. The templates must already have
    ''' passed <see cref="findTemplateError"/>.
    ''' </summary>
    '''
    ''' <param name="browserInfo">
    ''' Pre-computed <c> BrowserInfo </c> structure containing all browser parameters
    ''' </param>
    '''
    ''' <param name="scaffoldName">
    ''' The name of the <c> EntryScaffold </c> used as a template for the current entry being generated
    ''' </param>
    '''
    ''' <param name="fileKeyBases">
    ''' A set of FileKey templates that will be personalized for the current browser
    ''' </param>
    '''
    ''' <param name="regKeyBases">
    ''' A set of RegKey templates that will be personalized for the current browser
    ''' </param>
    '''
    ''' <param name="outputFile">
    ''' The <c> iniFile </c> containing all the entries being generated by Browser Builder
    ''' </param>
    '''
    ''' <param name="isGecko">
    ''' Indicates whether the browser being processed is Gecko-based
    ''' </param>
    '''
    ''' <param name="menuOutput">
    ''' The <c> MenuSection </c> accumulating user-visible output for this run
    ''' </param>
    Friend Sub generateBrowserEntry(browserInfo As BrowserInfo,
                                    scaffoldName As String,
                                    fileKeyBases As List(Of String),
                                    regKeyBases As List(Of String),
                              ByRef outputFile As iniFile,
                                    isGecko As Boolean,
                              ByRef menuOutput As MenuSection)

        Dim generatingMsg = $"Generating entry for browser: {browserInfo.Name}"
        gLog($"  {generatingMsg}")
        menuOutput.AddColoredLine(generatingMsg, ConsoleColor.Magenta)

        Dim entryName As String = $"{browserInfo.Name} {scaffoldName} *"
        Dim newSection As New iniSection(entryName)
        newSection.AddKey(New iniKey($"Section={browserInfo.SectionName}"))

        processUserDataPaths(newSection, browserInfo.UserDataPaths, browserInfo.UserDataParentPaths, browserInfo.TruncateDetect, fileKeyBases, browserInfo.Name, isGecko, menuOutput)

        processRegKeyBases(regKeyBases, newSection, browserInfo.RegistryRoots, menuOutput)

        outputFile.AddSection(newSection)

        Dim generatedText = $"Generated entry: {entryName}"
        menuOutput.AddColoredLine(generatedText, ConsoleColor.Yellow)
        gLog($"        {generatedText}")

    End Sub

    ''' <summary>
    ''' Adds RegKeys to <paramref name="newSection"/> for each template and registry root
    ''' pair, numbered from 1, with <c> &lt;RegistryRoot&gt; </c> replaced by the root and any
    ''' inline list fanned out. A browser with no roots gets no RegKeys.
    ''' </summary>
    '''
    ''' <param name="regKeyBases">
    ''' The set of templates for RegKeys that will be personalized for the current browser
    ''' </param>
    '''
    ''' <param name="newSection">
    ''' The <c> iniSection </c> to which the generated RegKeys will be added
    ''' </param>
    '''
    ''' <param name="RegistryRoots">
    ''' The set of root paths in the registry for the current browser's RegKeys
    ''' </param>
    '''
    ''' <param name="menuOutput">
    ''' The <c> MenuSection </c> receiving any expansion warnings
    ''' </param>
    Private Sub processRegKeyBases(regKeyBases As List(Of String),
                             ByRef newSection As iniSection,
                                   RegistryRoots As List(Of String),
                                   menuOutput As MenuSection)

        Dim regKeyNum As Integer = 1

        For Each regKeyBase In regKeyBases

            For Each root In RegistryRoots

                addExpandedKeys(newSection, "RegKey", regKeyNum, replaceToken(regKeyBase, RegistryRootToken, root), ExpansionDomain.Registry, menuOutput)

            Next

        Next

    End Sub

    ''' <summary>
    ''' Runs <paramref name="keyValue"/> through <see cref="VariableExpander.Expand"/> with no
    ''' declared variables and adds one numbered key per value it returns, logging its
    ''' diagnostics and showing its warnings
    ''' </summary>
    '''
    ''' <param name="section">
    ''' The <c> iniSection </c> to which the keys will be added
    ''' </param>
    '''
    ''' <param name="keyName">
    ''' The name of the keys, without numbers
    ''' </param>
    '''
    ''' <param name="keyNum">
    ''' The number of the next key, advanced past every key we add
    ''' </param>
    '''
    ''' <param name="keyValue">
    ''' The value, with its browser tokens already substituted
    ''' </param>
    '''
    ''' <param name="domain">
    ''' The expansion domain of the key
    ''' </param>
    '''
    ''' <param name="menuOutput">
    ''' The <c> MenuSection </c> receiving any expansion warnings
    ''' </param>
    Private Sub addExpandedKeys(section As iniSection,
                                keyName As String,
                          ByRef keyNum As Integer,
                                keyValue As String,
                                domain As ExpansionDomain,
                                menuOutput As MenuSection)

        Dim result = VariableExpander.Expand(keyValue, New VariableSet, domain, $"[{section.Name}].{keyName}")

        For Each diag In result.Diagnostics

            gLog(diag.Message)
            If diag.Severity = DiagnosticSeverity.Warning Then menuOutput.AddWarning(diag.Message)

        Next

        For Each value In result.Values

            CreateAndAddKey(section, keyName, keyNum, value)

        Next

    End Sub

    ''' <summary>
    ''' Creates a numbered <c> iniKey </c> with the provided <paramref name="keyName"/>,
    ''' <paramref name="keyNum"/>, and <paramref name="keyValue"/>, then adds it to
    ''' <paramref name="section"/>, incrementing <paramref name="keyNum"/>
    ''' </summary>
    '''
    ''' <param name="section">
    ''' An <c> iniSection </c> to which a key will be added
    ''' </param>
    '''
    ''' <param name="keyName">
    ''' The name of the key, does not include numbers
    ''' </param>
    '''
    ''' <param name="keyNum">
    ''' The number appended after <paramref name="keyName"/>
    ''' </param>
    '''
    ''' <param name="keyValue">
    ''' The value of the key
    ''' </param>
    Private Sub CreateAndAddKey(ByRef section As iniSection,
                                      keyName As String,
                                ByRef keyNum As Integer,
                                      keyValue As String)

        section.AddKey(New iniKey($"{keyName}{keyNum}={keyValue}"))
        keyNum += 1

    End Sub

    ''' <summary>
    ''' Iterates through the user data paths and generates DetectFile and FileKey entries
    ''' for each path, adding them to <paramref name="newSection"/>. The DetectFile is
    ''' unnumbered when there is only one path. FileKeys are numbered on across all paths.
    ''' <br /><br />
    '''
    ''' Each template holds exactly one kind of file token, matched in any case, which
    ''' <see cref="findTemplateError"/> has already checked:
    ''' <list type="bullet">
    '''
    ''' <item>
    ''' <c> &lt;UserDataPath&gt; </c>: the user data path, or the Opera GX pair from
    ''' <see cref="operaGXValues"/>
    ''' </item>
    '''
    ''' <item>
    ''' <c> &lt;UserDataParent&gt; </c>: the parent path. For Chromium we first add a copy with
    ''' <c> %LocalAppData% </c> and <c> %AppData% </c> changed to <c> %ProgramFiles% </c>.
    ''' </item>
    '''
    ''' <item>
    ''' <c> &lt;LocalUserDataPath&gt; </c>, Gecko only: the user data path with
    ''' <c> %AppData% </c> changed to <c> %LocalAppData% </c>
    ''' </item>
    '''
    ''' </list>
    '''
    ''' Each substituted value then goes through <see cref="addExpandedKeys"/>, so an inline
    ''' list fans it out.
    ''' </summary>
    '''
    ''' <param name="newSection">
    ''' The <c> iniSection </c> currently being built by Browser Builder
    ''' </param>
    '''
    ''' <param name="userDataPaths">
    ''' The set of user data paths provided from the ruleset file for the current browser
    ''' </param>
    '''
    ''' <param name="userDataParentPaths">
    ''' The set of parent paths for the user data paths, used for truncation
    ''' </param>
    '''
    ''' <param name="truncate">
    ''' Indicates whether the DetectFile keys use the parent path instead of the user data path
    ''' </param>
    '''
    ''' <param name="fileKeyBases">
    ''' The set of FileKey templates provided by the current <c> EntryScaffold </c>
    ''' </param>
    '''
    ''' <param name="BrowserName">
    ''' The name of the browser for which the current entry is being built
    ''' </param>
    '''
    ''' <param name="isGecko">
    ''' Indicates whether the current browser is Gecko-based
    ''' </param>
    '''
    ''' <param name="menuOutput">
    ''' The <c> MenuSection </c> receiving any expansion warnings
    ''' </param>
    Private Sub processUserDataPaths(ByRef newSection As iniSection,
                                     ByRef userDataPaths As List(Of String),
                                     ByRef userDataParentPaths As List(Of String),
                                           truncate As Boolean,
                                     ByRef fileKeyBases As List(Of String),
                                           BrowserName As String,
                                           isGecko As Boolean,
                                           menuOutput As MenuSection)

        Dim fileKeyNum As Integer = 1
        Dim multiPath = userDataPaths.Count > 1

        For i = 0 To userDataPaths.Count - 1

            Dim DetectFileNum = If(multiPath, (i + 1).ToString, "")
            newSection.AddKey(New iniKey($"DetectFile{DetectFileNum}={If(truncate, userDataParentPaths(i), userDataPaths(i))}"))

            For Each fileKeyBase In fileKeyBases

                Dim fileKeyValues As New List(Of String)

                Select Case True

                    Case hasToken(fileKeyBase, UserDataPathToken)

                        Dim operaValues = operaGXValues(BrowserName, fileKeyBase, userDataPaths(i))

                        If operaValues IsNot Nothing Then

                            fileKeyValues.AddRange(operaValues)

                        Else

                            fileKeyValues.Add(replaceToken(fileKeyBase, UserDataPathToken, userDataPaths(i)))

                        End If

                    Case hasToken(fileKeyBase, UserDataParentToken)

                        Dim parentValue = replaceToken(fileKeyBase, UserDataParentToken, userDataParentPaths(i))

                        If Not isGecko Then

                            Dim programFilesValue = parentValue.Replace("%LocalAppData%", "%ProgramFiles%")
                            fileKeyValues.Add(programFilesValue.Replace("%AppData%", "%ProgramFiles%"))

                        End If

                        fileKeyValues.Add(parentValue)

                    Case hasToken(fileKeyBase, LocalUserDataPathToken)

                        Dim localUserDataPath = userDataPaths(i).Replace("%AppData%", "%LocalAppData%")
                        fileKeyValues.Add(replaceToken(fileKeyBase, LocalUserDataPathToken, localUserDataPath))

                End Select

                For Each fileKeyValue In fileKeyValues

                    addExpandedKeys(newSection, "FileKey", fileKeyNum, fileKeyValue, ExpansionDomain.Filesystem, menuOutput)

                Next

            Next

        Next

    End Sub

End Module
