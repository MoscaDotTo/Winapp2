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

Imports System.Text

''' <summary>
''' Tests for BrowserBuilder's EntryScaffold tokens: what each of <c> &lt;UserDataPath&gt; </c>,
''' <c> &lt;UserDataParent&gt; </c>, <c> &lt;LocalUserDataPath&gt; </c> and
''' <c> &lt;RegistryRoot&gt; </c> becomes for a Chromium and a Gecko browser, the Opera GX side
''' profile shapes, case-insensitive and <c> $ </c>-safe matching, inline lists, and the
''' templates the builder refuses (each dropped once, with a failing exit code), including the
''' retired <c> %...% </c> spellings
''' </summary>
<TestClass()> Public Class BrowserBuilderTests

    ''' <summary> A Chromium browser with two registry roots </summary>
    Private Const Chrome As String =
        "[BrowserInfo: Chrome]" & vbCrLf &
        "Section=Google Chrome" & vbCrLf &
        "UserDataPath=%LocalAppData%\Google\Chrome\User Data" & vbCrLf &
        "RegistryRoot=HKCU\Software\Google\Chrome" & vbCrLf &
        "RegistryRoot=HKCU\Software\Chromium" & vbCrLf

    ''' <summary> A second Chromium browser, so a refused template could repeat per browser </summary>
    Private Const Brave As String =
        "[BrowserInfo: Brave]" & vbCrLf &
        "Section=Brave" & vbCrLf &
        "UserDataPath=%LocalAppData%\BraveSoftware\Brave-Browser\User Data" & vbCrLf

    ''' <summary> A Gecko browser with one registry root </summary>
    Private Const Firefox As String =
        "[BrowserInfo: Firefox]" & vbCrLf &
        "Section=Firefox" & vbCrLf &
        "UserDataPath=%AppData%\Mozilla\Firefox\Profiles" & vbCrLf &
        "RegistryRoot=HKCU\Software\Mozilla" & vbCrLf

    ''' <summary> Opera GX, whose FileKeys get a side profile copy </summary>
    Private Const OperaGX As String =
        "[BrowserInfo: Opera GX]" & vbCrLf &
        "Section=Opera" & vbCrLf &
        "UserDataPath=%AppData%\Opera Software\Opera GX Stable" & vbCrLf

    ''' <summary> Opera GX's user data path </summary>
    Private Const GXPath As String = "%AppData%\Opera Software\Opera GX Stable"

    ''' <summary>Returns the sections parsed from <paramref name="text"/></summary>
    ''' <param name="text">Literal ini text</param>
    Private Shared Function Sections(text As String) As List(Of winapp2ool.iniSection)

        Dim bytes = Encoding.UTF8.GetBytes(text)
        Using ms As New IO.MemoryStream(bytes)
            Using reader As New IO.StreamReader(ms)
                Return winapp2ool.iniFile.FromStream(reader, "", "test.ini").ToList()
            End Using
        End Using

    End Function

    ''' <summary>
    ''' Parses the BrowserInfo sections of <paramref name="text"/> and runs each EntryScaffold
    ''' section in it through <c> processEntryScaffold </c>, as <c> buildScaffolds </c> does
    ''' </summary>
    '''
    ''' <param name="text">Literal ini text holding BrowserInfo and EntryScaffold sections</param>
    '''
    ''' <param name="isGecko">Indicates whether the text is a Gecko ruleset</param>
    '''
    ''' <param name="log">Set to every line the build wrote to the log</param>
    '''
    ''' <returns>The file holding the generated entries</returns>
    Private Shared Function Generate(text As String, isGecko As Boolean, ByRef log As List(Of String)) As winapp2ool.iniFile

        Dim menu As New winapp2ool.MenuSection
        Dim output = winapp2ool.iniFile.Empty("", "browsers.ini")

        Using cap = winapp2ool.gLogCapture()

            Dim parsed = Sections(text)
            Dim browsers = parsed.Where(Function(s) s.Name.StartsWith("BrowserInfo:", StringComparison.Ordinal)) _
                                 .Select(Function(s) winapp2ool.BrowserBuilder.parseBrowserInfo(s, menu)).ToList()

            For Each scaffold In parsed.Where(Function(s) s.Name.StartsWith("EntryScaffold:", StringComparison.Ordinal))
                winapp2ool.BrowserBuilder.processEntryScaffold(scaffold, browsers, isGecko, output, menu)
            Next

            log = cap.Lines.ToList()

        End Using

        Return output

    End Function

    ''' <summary>
    ''' Runs <paramref name="action"/> from a zero exit code and returns the exit code it leaves,
    ''' restoring the previous one afterwards
    ''' </summary>
    '''
    ''' <param name="action">The work to run</param>
    Private Shared Function ExitCodeAfter(action As Action) As Integer

        Dim savedExitCode = Environment.ExitCode

        Try

            Environment.ExitCode = 0
            action()
            Return Environment.ExitCode

        Finally

            Environment.ExitCode = savedExitCode

        End Try

    End Function

    ''' <summary>Returns every key of the entry <paramref name="entryName"/> as <c> Name=Value </c></summary>
    '''
    ''' <param name="output">The generated file</param>
    '''
    ''' <param name="entryName">The entry's name</param>
    Private Shared Function KeysOf(output As winapp2ool.iniFile, entryName As String) As List(Of String)

        Dim section = output.GetSection(entryName)
        Assert.IsNotNull(section, $"Expected an entry named {entryName}")
        Return section.Keys.Select(Function(k) k.ToString()).ToList()

    End Function

    ''' <summary>Returns the FileKey and RegKey values of the entry <paramref name="entryName"/></summary>
    '''
    ''' <param name="output">The generated file</param>
    '''
    ''' <param name="entryName">The entry's name</param>
    Private Shared Function PathKeysOf(output As winapp2ool.iniFile, entryName As String) As List(Of String)

        Return KeysOf(output, entryName).Where(Function(k) k.StartsWith("FileKey", StringComparison.Ordinal) OrElse k.StartsWith("RegKey", StringComparison.Ordinal)).ToList()

    End Function

    ''' <summary>
    ''' Builds <paramref name="text"/> with a clean exit code and returns the output
    ''' </summary>
    '''
    ''' <param name="text">Literal ini text</param>
    '''
    ''' <param name="isGecko">Indicates whether the text is a Gecko ruleset</param>
    Private Shared Function GenerateClean(text As String, isGecko As Boolean) As winapp2ool.iniFile

        Dim output As winapp2ool.iniFile = Nothing
        Dim log As List(Of String) = Nothing

        Assert.AreEqual(0, ExitCodeAfter(Sub() output = Generate(text, isGecko, log)), String.Join("; ", log))

        Return output

    End Function

    ' ----- Substitution -----

    ''' <summary>
    ''' <c> &lt;UserDataPath&gt; </c> becomes a Chromium browser's user data path
    ''' </summary>
    <TestMethod()> Public Sub UserDataPath_Chromium()

        Dim output = GenerateClean(Chrome & "[EntryScaffold: Caches]" & vbCrLf & "FileKeyBase=<UserDataPath>\*\Cache|*|REMOVESELF" & vbCrLf, False)

        CollectionAssert.AreEqual(New List(Of String) From {
            "Section=Google Chrome",
            "DetectFile=%LocalAppData%\Google\Chrome\User Data",
            "FileKey1=%LocalAppData%\Google\Chrome\User Data\*\Cache|*|REMOVESELF"
        }, KeysOf(output, "Chrome Caches *"))

    End Sub

    ''' <summary>
    ''' <c> &lt;UserDataPath&gt; </c> becomes a Gecko browser's profiles path
    ''' </summary>
    <TestMethod()> Public Sub UserDataPath_Gecko()

        Dim output = GenerateClean(Firefox & "[EntryScaffold: Cookies]" & vbCrLf & "FileKeyBase=<UserDataPath>\*|cookies*" & vbCrLf, True)

        CollectionAssert.AreEqual(New List(Of String) From {"FileKey1=%AppData%\Mozilla\Firefox\Profiles\*|cookies*"},
                                  PathKeysOf(output, "Firefox Cookies *"))

    End Sub

    ''' <summary>
    ''' For Chromium, <c> &lt;UserDataParent&gt; </c> gives a <c> %ProgramFiles% </c> copy first,
    ''' then the parent of the user data path
    ''' </summary>
    <TestMethod()> Public Sub UserDataParent_Chromium_AddsProgramFilesCopy()

        Dim output = GenerateClean(Chrome & "[EntryScaffold: Logs]" & vbCrLf & "FileKeyBase=<UserDataParent>\Application|debug.log" & vbCrLf, False)

        CollectionAssert.AreEqual(New List(Of String) From {
            "FileKey1=%ProgramFiles%\Google\Chrome\Application|debug.log",
            "FileKey2=%LocalAppData%\Google\Chrome\Application|debug.log"
        }, PathKeysOf(output, "Chrome Logs *"))

    End Sub

    ''' <summary>
    ''' For Gecko, <c> &lt;UserDataParent&gt; </c> is the parent of the profiles path alone
    ''' </summary>
    <TestMethod()> Public Sub UserDataParent_Gecko_HasNoCopy()

        Dim output = GenerateClean(Firefox & "[EntryScaffold: Logs]" & vbCrLf & "FileKeyBase=<UserDataParent>|*.log|RECURSE" & vbCrLf, True)

        CollectionAssert.AreEqual(New List(Of String) From {"FileKey1=%AppData%\Mozilla\Firefox|*.log|RECURSE"},
                                  PathKeysOf(output, "Firefox Logs *"))

    End Sub

    ''' <summary>
    ''' <c> &lt;LocalUserDataPath&gt; </c> is the Gecko profiles path moved to <c> %LocalAppData% </c>
    ''' </summary>
    <TestMethod()> Public Sub LocalUserDataPath_Gecko()

        Dim output = GenerateClean(Firefox & "[EntryScaffold: Caches]" & vbCrLf & "FileKeyBase=<LocalUserDataPath>\*\thumbnails|*|REMOVESELF" & vbCrLf, True)

        CollectionAssert.AreEqual(New List(Of String) From {"FileKey1=%LocalAppData%\Mozilla\Firefox\Profiles\*\thumbnails|*|REMOVESELF"},
                                  PathKeysOf(output, "Firefox Caches *"))

    End Sub

    ''' <summary>
    ''' Each of the four Opera GX template shapes gives the main folder and its side profiles
    ''' </summary>
    <TestMethod()> Public Sub OperaGX_AllFourShapes()

        Dim scaffold = "[EntryScaffold: Mixed]" & vbCrLf &
                       "FileKeyBase=<UserDataPath>\*|Bookmarks.bak" & vbCrLf &
                       "FileKeyBase=<UserDataPath>\*\Sessions|*|REMOVESELF" & vbCrLf &
                       "FileKeyBase=<UserDataPath>|*.log" & vbCrLf &
                       "FileKeyBase=<UserDataPath>\Crashpad|*|REMOVESELF" & vbCrLf

        Dim output = GenerateClean(OperaGX & scaffold, False)

        CollectionAssert.AreEqual(New List(Of String) From {
            $"FileKey1={GXPath}|Bookmarks.bak",
            $"FileKey2={GXPath}\_side_profiles\*|Bookmarks.bak",
            $"FileKey3={GXPath}\Sessions|*|REMOVESELF",
            $"FileKey4={GXPath}\_side_profiles\*\Sessions|*|REMOVESELF",
            $"FileKey5={GXPath}|*.log",
            $"FileKey6={GXPath}\_side_profiles\*|*.log",
            $"FileKey7={GXPath}\Crashpad|*|REMOVESELF",
            $"FileKey8={GXPath}\_side_profiles\*\Crashpad|*|REMOVESELF"
        }, PathKeysOf(output, "Opera GX Mixed *"))

    End Sub

    ''' <summary>
    ''' The Opera GX shapes are recognized whatever the token's case
    ''' </summary>
    <TestMethod()> Public Sub OperaGX_PrefixIgnoresCase()

        Dim output = GenerateClean(OperaGX & "[EntryScaffold: Caches]" & vbCrLf & "FileKeyBase=<userdatapath>\*\Cache|*|REMOVESELF" & vbCrLf, False)

        CollectionAssert.AreEqual(New List(Of String) From {
            $"FileKey1={GXPath}\Cache|*|REMOVESELF",
            $"FileKey2={GXPath}\_side_profiles\*\Cache|*|REMOVESELF"
        }, PathKeysOf(output, "Opera GX Caches *"))

    End Sub

    ''' <summary>
    ''' <c> &lt;RegistryRoot&gt; </c> gives one RegKey per root, numbered on across templates
    ''' </summary>
    <TestMethod()> Public Sub RegistryRoot_OncePerRoot()

        Dim scaffold = "[EntryScaffold: Telemetry]" & vbCrLf &
                       "FileKeyBase=<UserDataPath>|*.log" & vbCrLf &
                       "RegKeyBase=<RegistryRoot>\BlBeacon|state" & vbCrLf &
                       "RegKeyBase=<RegistryRoot>\StabilityMetrics" & vbCrLf

        Dim output = GenerateClean(Chrome & scaffold, False)

        CollectionAssert.AreEqual(New List(Of String) From {
            "FileKey1=%LocalAppData%\Google\Chrome\User Data|*.log",
            "RegKey1=HKCU\Software\Google\Chrome\BlBeacon|state",
            "RegKey2=HKCU\Software\Chromium\BlBeacon|state",
            "RegKey3=HKCU\Software\Google\Chrome\StabilityMetrics",
            "RegKey4=HKCU\Software\Chromium\StabilityMetrics"
        }, PathKeysOf(output, "Chrome Telemetry *"))

    End Sub

    ''' <summary>
    ''' Tokens match in any case
    ''' </summary>
    <TestMethod()> Public Sub Tokens_MatchInAnyCase()

        Dim scaffold = "[EntryScaffold: Mixed]" & vbCrLf &
                       "FileKeyBase=<USERDATAPATH>\*|History*" & vbCrLf &
                       "FileKeyBase=<userDataParent>\Application|debug.log" & vbCrLf &
                       "RegKeyBase=<registryroot>\Installer|LastChecked" & vbCrLf

        Dim output = GenerateClean(Chrome & scaffold, False)

        CollectionAssert.AreEqual(New List(Of String) From {
            "FileKey1=%LocalAppData%\Google\Chrome\User Data\*|History*",
            "FileKey2=%ProgramFiles%\Google\Chrome\Application|debug.log",
            "FileKey3=%LocalAppData%\Google\Chrome\Application|debug.log",
            "RegKey1=HKCU\Software\Google\Chrome\Installer|LastChecked",
            "RegKey2=HKCU\Software\Chromium\Installer|LastChecked"
        }, PathKeysOf(output, "Chrome Mixed *"))

    End Sub

    ''' <summary>
    ''' A <c> $ </c> in a browser's path is inserted as written, never read as a substitution
    ''' </summary>
    <TestMethod()> Public Sub DollarInPath_IsLiteral()

        Dim browser = "[BrowserInfo: Odd]" & vbCrLf &
                      "Section=Odd" & vbCrLf &
                      "UserDataPath=%LocalAppData%\$0 $1 Browser\User Data" & vbCrLf &
                      "RegistryRoot=HKCU\Software\$1" & vbCrLf

        Dim scaffold = "[EntryScaffold: Mixed]" & vbCrLf &
                       "FileKeyBase=<UserDataPath>\*|History*" & vbCrLf &
                       "RegKeyBase=<RegistryRoot>\Installer" & vbCrLf

        Dim output = GenerateClean(browser & scaffold, False)

        CollectionAssert.AreEqual(New List(Of String) From {
            "FileKey1=%LocalAppData%\$0 $1 Browser\User Data\*|History*",
            "RegKey1=HKCU\Software\$1\Installer"
        }, PathKeysOf(output, "Odd Mixed *"))

    End Sub

    ''' <summary>
    ''' An inline list fans a FileKeyBase out, and a RegKeyBase too
    ''' </summary>
    <TestMethod()> Public Sub InlineList_FansOut()

        Dim scaffold = "[EntryScaffold: Caches]" & vbCrLf &
                       "FileKeyBase=<UserDataPath>\*\<Cache,Code Cache>|*|REMOVESELF" & vbCrLf &
                       "RegKeyBase=<RegistryRoot>\<BlBeacon,Installer>" & vbCrLf

        Dim output = GenerateClean(Chrome & scaffold, False)

        CollectionAssert.AreEqual(New List(Of String) From {
            "FileKey1=%LocalAppData%\Google\Chrome\User Data\*\Cache|*|REMOVESELF",
            "FileKey2=%LocalAppData%\Google\Chrome\User Data\*\Code Cache|*|REMOVESELF",
            "RegKey1=HKCU\Software\Google\Chrome\BlBeacon",
            "RegKey2=HKCU\Software\Google\Chrome\Installer",
            "RegKey3=HKCU\Software\Chromium\BlBeacon",
            "RegKey4=HKCU\Software\Chromium\Installer"
        }, PathKeysOf(output, "Chrome Caches *"))

    End Sub

    ''' <summary>
    ''' An inline list works in the Opera GX shapes, after the side profile rewrite
    ''' </summary>
    <TestMethod()> Public Sub InlineList_InOperaGX()

        Dim output = GenerateClean(OperaGX & "[EntryScaffold: Caches]" & vbCrLf & "FileKeyBase=<UserDataPath>\*\<Cache,GPUCache>|*|REMOVESELF" & vbCrLf, False)

        CollectionAssert.AreEqual(New List(Of String) From {
            $"FileKey1={GXPath}\Cache|*|REMOVESELF",
            $"FileKey2={GXPath}\GPUCache|*|REMOVESELF",
            $"FileKey3={GXPath}\_side_profiles\*\Cache|*|REMOVESELF",
            $"FileKey4={GXPath}\_side_profiles\*\GPUCache|*|REMOVESELF"
        }, PathKeysOf(output, "Opera GX Caches *"))

    End Sub

    ''' <summary>
    ''' A scaffold whose templates all pass leaves the exit code at 0 and logs no drop
    ''' </summary>
    <TestMethod()> Public Sub CleanScaffold_LeavesExitCodeZero()

        Dim log As List(Of String) = Nothing
        Dim scaffold = "[EntryScaffold: Mixed]" & vbCrLf &
                       "FileKeyBase=<UserDataPath>\*|History*" & vbCrLf &
                       "FileKeyBase=<UserDataParent>\Application|debug.log" & vbCrLf &
                       "RegKeyBase=<RegistryRoot>\Installer" & vbCrLf

        Assert.AreEqual(0, ExitCodeAfter(Sub() Generate(Chrome & Brave & scaffold, False, log)))
        Assert.IsFalse(log.Any(Function(l) l.Contains("dropping")), String.Join("; ", log))

    End Sub

    ' ----- Refused templates -----

    ''' <summary>
    ''' Builds <paramref name="badTemplate"/> beside a good FileKeyBase for two browsers of the
    ''' ruleset, asserting that the run failed, that exactly one log line drops the template and
    ''' contains <paramref name="expected"/>, and that each browser keeps only the good key
    ''' </summary>
    '''
    ''' <param name="badTemplate">The refused key line, e.g. <c> FileKeyBase=&lt;Typo&gt;|* </c></param>
    '''
    ''' <param name="expected">Text the drop message must contain</param>
    '''
    ''' <param name="isGecko">
    ''' Indicates whether to build with the Gecko ruleset <br /><br />
    ''' Optional, Default: <c> False </c>
    ''' </param>
    '''
    ''' <returns>The drop message</returns>
    Private Shared Function AssertRefused(badTemplate As String, expected As String, Optional isGecko As Boolean = False) As String

        Dim browsers = If(isGecko, Firefox & Firefox.Replace("[BrowserInfo: Firefox]", "[BrowserInfo: Waterfox]"), Chrome & Brave)
        Dim scaffold = "[EntryScaffold: Mixed]" & vbCrLf &
                       badTemplate & vbCrLf &
                       "FileKeyBase=<UserDataPath>\*|History*" & vbCrLf

        Dim output As winapp2ool.iniFile = Nothing
        Dim log As List(Of String) = Nothing

        Assert.AreEqual(1, ExitCodeAfter(Sub() output = Generate(browsers & scaffold, isGecko, log)))

        Dim drops = log.Where(Function(l) l.Contains("; dropping `")).ToList()
        Assert.AreEqual(1, drops.Count, String.Join("; ", log))
        StringAssert.Contains(drops(0), expected)
        StringAssert.Contains(drops(0), $"; dropping `{badTemplate}`")

        For Each section In output
            CollectionAssert.AreEqual(New List(Of String) From {"FileKey1"},
                                      section.Keys.Where(Function(k) k.Name.StartsWith("FileKey", StringComparison.Ordinal) OrElse k.Name.StartsWith("RegKey", StringComparison.Ordinal)) _
                                                  .Select(Function(k) k.Name).ToList(), section.Name)
            StringAssert.Contains(section.Keys.First(Function(k) k.Name = "FileKey1").Value, "|History*")
        Next

        Assert.AreEqual(2, output.Count)

        Return drops(0)

    End Function

    ''' <summary> A FileKeyBase with no file token is refused </summary>
    <TestMethod()> Public Sub NoFileToken_IsRefused()

        AssertRefused("FileKeyBase=%LocalAppData%\Google\Chrome\User Data|*.log", "has no file token")

    End Sub

    ''' <summary> <c> &lt;LocalUserDataPath&gt; </c> in a Chromium template is refused </summary>
    <TestMethod()> Public Sub LocalUserDataPath_InChromium_IsRefused()

        AssertRefused("FileKeyBase=<LocalUserDataPath>\*\Cache|*", "uses <LocalUserDataPath>, which only Gecko browsers have")

    End Sub

    ''' <summary> <c> &lt;LocalUserDataPath&gt; </c> is refused in Chromium whatever its case </summary>
    <TestMethod()> Public Sub LocalUserDataPath_InChromium_AnyCase_IsRefused()

        AssertRefused("FileKeyBase=<localuserdatapath>\*\Cache|*", "which only Gecko browsers have")

    End Sub

    ''' <summary> A FileKeyBase mixing two kinds of file token is refused </summary>
    <TestMethod()> Public Sub MixedFileTokens_AreRefused()

        AssertRefused("FileKeyBase=<UserDataParent>\Application|<UserDataPath>", "mixes `<UserDataParent>` and `<UserDataPath>`")

    End Sub

    ''' <summary> A Gecko FileKeyBase mixing <c> &lt;UserDataPath&gt; </c> and <c> &lt;LocalUserDataPath&gt; </c> is refused </summary>
    <TestMethod()> Public Sub MixedFileTokens_Gecko_AreRefused()

        AssertRefused("FileKeyBase=<LocalUserDataPath>\<UserDataPath>|*", "mixes `<LocalUserDataPath>` and `<UserDataPath>`", isGecko:=True)

    End Sub

    ''' <summary> <c> &lt;RegistryRoot&gt; </c> in a FileKeyBase is refused </summary>
    <TestMethod()> Public Sub RegistryRoot_InFileKeyBase_IsRefused()

        AssertRefused("FileKeyBase=<RegistryRoot>\Cache|*", "uses <RegistryRoot>, which only a RegKeyBase takes")

    End Sub

    ''' <summary> A file token in a RegKeyBase is refused </summary>
    <TestMethod()> Public Sub FileToken_InRegKeyBase_IsRefused()

        AssertRefused("RegKeyBase=<RegistryRoot>\<UserDataPath>", "uses the file token `<UserDataPath>`; a RegKeyBase takes only <RegistryRoot>")

    End Sub

    ''' <summary> A RegKeyBase without <c> &lt;RegistryRoot&gt; </c> is refused rather than emitted once per root </summary>
    <TestMethod()> Public Sub RegKeyBase_WithoutRegistryRoot_IsRefused()

        AssertRefused("RegKeyBase=HKCU\Software\Google\Chrome\BlBeacon", "has no <RegistryRoot> token")

    End Sub

    ''' <summary> An unknown token in a FileKeyBase is refused </summary>
    <TestMethod()> Public Sub UnknownToken_InFileKeyBase_IsRefused()

        AssertRefused("FileKeyBase=<UserDataPath>\<Typo>|*", "Unknown token `<Typo>`")

    End Sub

    ''' <summary>
    ''' An unknown token in a RegKeyBase is refused too. The registry domain would keep it as
    ''' literal text.
    ''' </summary>
    <TestMethod()> Public Sub UnknownToken_InRegKeyBase_IsRefused()

        AssertRefused("RegKeyBase=<RegistryRoot>\<Typo>", "Unknown token `<Typo>`")

    End Sub

    ''' <summary> A misspelled token is unknown, not silently skipped </summary>
    <TestMethod()> Public Sub MisspelledToken_IsRefused()

        AssertRefused("FileKeyBase=<UserDatePath>\*|*.log", "Unknown token `<UserDatePath>`")

    End Sub

    ''' <summary> An inline list with no non-empty value is refused, in either key </summary>
    <TestMethod()> Public Sub EmptyInlineList_IsRefused()

        AssertRefused("FileKeyBase=<UserDataPath>\*\Cache|<,>", "Inline list `<,>` in [EntryScaffold: Mixed].FileKeyBase has no values")
        AssertRefused("RegKeyBase=<RegistryRoot>\< , >", "Inline list `< , >` in [EntryScaffold: Mixed].RegKeyBase has no values")

    End Sub

    ''' <summary>
    ''' An inline list with one non-empty value passes, and the expander drops the empty one
    ''' </summary>
    <TestMethod()> Public Sub InlineList_WithOneValue_Passes()

        Dim output = GenerateClean(Chrome & "[EntryScaffold: Caches]" & vbCrLf & "FileKeyBase=<UserDataPath>\*\<Cache,>|*" & vbCrLf, False)

        CollectionAssert.AreEqual(New List(Of String) From {"FileKey1=%LocalAppData%\Google\Chrome\User Data\*\Cache|*"},
                                  PathKeysOf(output, "Chrome Caches *"))

    End Sub

    ''' <summary> An empty <c> &lt;&gt; </c> is refused rather than shipped </summary>
    <TestMethod()> Public Sub EmptyBrackets_AreRefused()

        AssertRefused("FileKeyBase=<UserDataPath>\<>|x.log", "Malformed `<...>` token in [EntryScaffold: Mixed].FileKeyBase")
        AssertRefused("RegKeyBase=<RegistryRoot>\<>", "Malformed `<...>` token in [EntryScaffold: Mixed].RegKeyBase")

    End Sub

    ''' <summary> A stray <c> &gt; </c> or <c> &lt; </c> is refused rather than shipped </summary>
    <TestMethod()> Public Sub StrayBracket_IsRefused()

        AssertRefused("FileKeyBase=<UserDataPath>>\y|x.log", "Malformed `<...>` token in [EntryScaffold: Mixed].FileKeyBase")
        AssertRefused("FileKeyBase=<UserDataPath>\<y|x.log", "Malformed `<...>` token in [EntryScaffold: Mixed].FileKeyBase")
        AssertRefused("RegKeyBase=<RegistryRoot>>\y", "Malformed `<...>` token in [EntryScaffold: Mixed].RegKeyBase")

    End Sub

    ''' <summary> A token nested in an inline list is refused rather than shipped as the list's other value </summary>
    <TestMethod()> Public Sub NestedToken_IsRefused()

        AssertRefused("FileKeyBase=<<UserDataPath>,z>\w|*", "Malformed `<...>` token in [EntryScaffold: Mixed].FileKeyBase")
        AssertRefused("RegKeyBase=<<RegistryRoot>,z>\w", "Malformed `<...>` token in [EntryScaffold: Mixed].RegKeyBase")

    End Sub

    ' ----- BrowserInfo values -----

    ''' <summary>
    ''' Builds a scaffold for a browser carrying one extra <paramref name="badKey"/>, asserting
    ''' that the run failed, that exactly one log line skips that value and says BrowserInfo
    ''' values are literal paths, and returns the generated keys
    ''' </summary>
    '''
    ''' <param name="badKey">The refused BrowserInfo key line</param>
    Private Shared Function AssertBrowserValueRefused(badKey As String) As List(Of String)

        Dim scaffold = "[EntryScaffold: Mixed]" & vbCrLf &
                       "FileKeyBase=<UserDataPath>\*|History*" & vbCrLf &
                       "RegKeyBase=<RegistryRoot>\Installer" & vbCrLf

        Dim output As winapp2ool.iniFile = Nothing
        Dim log As List(Of String) = Nothing

        Assert.AreEqual(1, ExitCodeAfter(Sub() output = Generate(Chrome & badKey & vbCrLf & scaffold, False, log)))

        Dim skips = log.Where(Function(l) l.Contains("; skipping `")).ToList()
        Assert.AreEqual(1, skips.Count, String.Join("; ", log))
        StringAssert.Contains(skips(0), "BrowserInfo values are literal paths")
        StringAssert.Contains(skips(0), $"; skipping `{badKey}`")

        Return PathKeysOf(output, "Chrome Mixed *")

    End Function

    ''' <summary>
    ''' A <c> UserDataPath= </c> holding a token is skipped, and the browser's other path still builds
    ''' </summary>
    <TestMethod()> Public Sub UserDataPath_WithToken_IsSkipped()

        CollectionAssert.AreEqual(New List(Of String) From {
            "FileKey1=%LocalAppData%\Google\Chrome\User Data\*|History*",
            "RegKey1=HKCU\Software\Google\Chrome\Installer",
            "RegKey2=HKCU\Software\Chromium\Installer"
        }, AssertBrowserValueRefused("UserDataPath=%LocalAppData%\<Foo>\User Data"))

    End Sub

    ''' <summary>
    ''' A <c> RegistryRoot= </c> holding an inline list is skipped rather than fanned out
    ''' </summary>
    <TestMethod()> Public Sub RegistryRoot_WithInlineList_IsSkipped()

        CollectionAssert.AreEqual(New List(Of String) From {
            "FileKey1=%LocalAppData%\Google\Chrome\User Data\*|History*",
            "RegKey1=HKCU\Software\Google\Chrome\Installer",
            "RegKey2=HKCU\Software\Chromium\Installer"
        }, AssertBrowserValueRefused("RegistryRoot=HKCU\Software\<a,b>"))

    End Sub

    ' ----- Retired spellings -----

    ''' <summary>
    ''' A numbered <c> %RegistryRoot2% </c> gets the plain hint, since BrowserBuilder's tokens take no index
    ''' </summary>
    <TestMethod()> Public Sub RetiredNumberedToken_GetsNoIndexHint()

        Dim msg = AssertRefused("RegKeyBase=%RegistryRoot2%\BlBeacon",
                                "`%RegistryRoot2%` in [EntryScaffold: Mixed].RegKeyBase looks like the retired `%RegistryRoot%` token, which is now spelled `<RegistryRoot>`")

        Assert.IsFalse(msg.Contains("[2]"), msg)

    End Sub

    ''' <summary> <c> %UserDataPath% </c> is refused, naming <c> &lt;UserDataPath&gt; </c> </summary>
    <TestMethod()> Public Sub RetiredUserDataPath_IsRefused()

        AssertRefused("FileKeyBase=%UserDataPath%\*|*.log", "`%UserDataPath%` in [EntryScaffold: Mixed].FileKeyBase is now spelled `<UserDataPath>`")

    End Sub

    ''' <summary> <c> %BrowserPath% </c> is refused, naming its new name <c> &lt;UserDataParent&gt; </c> </summary>
    <TestMethod()> Public Sub RetiredBrowserPath_NamesUserDataParent()

        AssertRefused("FileKeyBase=%BrowserPath%\Application|debug.log", "`%BrowserPath%` in [EntryScaffold: Mixed].FileKeyBase is now spelled `<UserDataParent>`")

    End Sub

    ''' <summary> A <c> %BrowserPath </c> missing its closing <c> % </c> is refused too </summary>
    <TestMethod()> Public Sub RetiredBrowserPath_Unclosed_IsRefused()

        AssertRefused("FileKeyBase=%BrowserPath\Application|x",
                      "`%BrowserPath` in [EntryScaffold: Mixed].FileKeyBase looks like the retired `%BrowserPath%` token, which is now spelled `<UserDataParent>`")

    End Sub

    ''' <summary> A lowercase <c> %browserpath% </c> is refused with the new name's capitalization </summary>
    <TestMethod()> Public Sub RetiredBrowserPath_Lowercase_IsRefused()

        AssertRefused("FileKeyBase=%browserpath%\Application|x", "`%browserpath%` in [EntryScaffold: Mixed].FileKeyBase is now spelled `<UserDataParent>`")

    End Sub

    ''' <summary> <c> %LocalDataPath% </c> is refused, naming <c> &lt;LocalUserDataPath&gt; </c> </summary>
    <TestMethod()> Public Sub RetiredLocalDataPath_NamesLocalUserDataPath()

        AssertRefused("FileKeyBase=%LocalDataPath%\*\cache2|*", "`%LocalDataPath%` in [EntryScaffold: Mixed].FileKeyBase is now spelled `<LocalUserDataPath>`", isGecko:=True)

    End Sub

    ''' <summary> <c> %RegistryRoot% </c> is refused, naming <c> &lt;RegistryRoot&gt; </c> </summary>
    <TestMethod()> Public Sub RetiredRegistryRoot_IsRefused()

        AssertRefused("RegKeyBase=%RegistryRoot%\BlBeacon|state", "`%RegistryRoot%` in [EntryScaffold: Mixed].RegKeyBase is now spelled `<RegistryRoot>`")

    End Sub

End Class
