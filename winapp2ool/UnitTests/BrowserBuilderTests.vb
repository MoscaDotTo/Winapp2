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

    ''' <summary> The Opera GX layout applies to the browser name in any case </summary>
    <TestMethod()> Public Sub OperaGX_NameIgnoresCase()

        Dim output = GenerateClean(OperaGX.Replace("[BrowserInfo: Opera GX]", "[BrowserInfo: opera gx]") &
                                   "[EntryScaffold: Caches]" & vbCrLf & "FileKeyBase=<UserDataPath>\*\Cache|*|REMOVESELF" & vbCrLf, False)

        CollectionAssert.AreEqual(New List(Of String) From {
            $"FileKey1={GXPath}\Cache|*|REMOVESELF",
            $"FileKey2={GXPath}\_side_profiles\*\Cache|*|REMOVESELF"
        }, PathKeysOf(output, "opera gx Caches *"))

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
    ''' <param name="declarations">
    ''' Variable declaration lines, each ending in a line break, put ahead of the template <br /><br />
    ''' Optional, Default: none
    ''' </param>
    '''
    ''' <returns>The drop message</returns>
    Private Shared Function AssertRefused(badTemplate As String, expected As String, Optional isGecko As Boolean = False,
                                          Optional declarations As String = "") As String

        Dim browsers = If(isGecko, Firefox & Firefox.Replace("[BrowserInfo: Firefox]", "[BrowserInfo: Waterfox]"), Chrome & Brave)
        Dim scaffold = "[EntryScaffold: Mixed]" & vbCrLf &
                       declarations &
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
    ''' Builds a scaffold for a browser carrying the extra key lines <paramref name="extra"/>,
    ''' asserting that the run failed, that exactly one log line skips <paramref name="badKey"/>
    ''' and contains <paramref name="expected"/>, and returns the generated keys
    ''' </summary>
    '''
    ''' <param name="badKey">The refused BrowserInfo key line</param>
    '''
    ''' <param name="expected">Text the skip message must contain</param>
    '''
    ''' <param name="extra">
    ''' Other key lines for the browser, each ending in a line break <br /><br />
    ''' Optional, Default: none
    ''' </param>
    Private Shared Function AssertBrowserValueRefused(badKey As String, expected As String,
                                                      Optional extra As String = "") As List(Of String)

        Dim scaffold = "[EntryScaffold: Mixed]" & vbCrLf &
                       "FileKeyBase=<UserDataPath>\*|History*" & vbCrLf &
                       "RegKeyBase=<RegistryRoot>\Installer" & vbCrLf

        Dim output As winapp2ool.iniFile = Nothing
        Dim log As List(Of String) = Nothing

        Assert.AreEqual(1, ExitCodeAfter(Sub() output = Generate(Chrome & extra & badKey & vbCrLf & scaffold, False, log)))

        Dim skips = log.Where(Function(l) l.Contains("; skipping `")).ToList()
        Assert.AreEqual(1, skips.Count, String.Join("; ", log))
        StringAssert.Contains(skips(0), expected)
        StringAssert.Contains(skips(0), $"; skipping `{badKey}`")

        Return PathKeysOf(output, "Chrome Mixed *")

    End Function

    ''' <summary> Chrome's keys under <c> [EntryScaffold: Mixed] </c> when every extra value was skipped </summary>
    Private Shared ReadOnly ChromeMixedKeys As New List(Of String) From {
        "FileKey1=%LocalAppData%\Google\Chrome\User Data\*|History*",
        "RegKey1=HKCU\Software\Google\Chrome\Installer",
        "RegKey2=HKCU\Software\Chromium\Installer"
    }

    ''' <summary>
    ''' A <c> UserDataPath= </c> holding a token the browser doesn't declare is skipped, and the
    ''' browser's other path still builds
    ''' </summary>
    <TestMethod()> Public Sub UserDataPath_WithUndeclaredVariable_IsSkipped()

        CollectionAssert.AreEqual(ChromeMixedKeys, AssertBrowserValueRefused("UserDataPath=%LocalAppData%\<Foo>\User Data", "Unknown variable `<Foo>`"))

    End Sub

    ''' <summary> A BrowserInfo value can't use one of the four scaffold tokens </summary>
    <TestMethod()> Public Sub BrowserValue_WithScaffoldToken_IsSkipped()

        CollectionAssert.AreEqual(ChromeMixedKeys, AssertBrowserValueRefused("RegistryRoot=<RegistryRoot>\Beta", "which only a scaffold takes"))

    End Sub

    ''' <summary> A BrowserInfo value with a malformed bracket is skipped </summary>
    <TestMethod()> Public Sub BrowserValue_Malformed_IsSkipped()

        CollectionAssert.AreEqual(ChromeMixedKeys, AssertBrowserValueRefused("UserDataPath=%LocalAppData%\Foo>\User Data", "Malformed `<...>` token"))

    End Sub

    ''' <summary> A BrowserInfo value using a variable with no values is skipped </summary>
    <TestMethod()> Public Sub BrowserValue_EmptyVariable_IsSkipped()

        CollectionAssert.AreEqual(ChromeMixedKeys, AssertBrowserValueRefused("RegistryRoot=HKCU\Software\<Channel>", "has no values", extra:="Channel=" & vbCrLf))

    End Sub

    ''' <summary>
    ''' A BrowserInfo variable that holds a token nothing declares leaves a bracket in the value,
    ''' which is skipped
    ''' </summary>
    <TestMethod()> Public Sub BrowserValue_VariableHoldingUnknownToken_IsSkipped()

        CollectionAssert.AreEqual(ChromeMixedKeys, AssertBrowserValueRefused("RegistryRoot=HKCU\Software\<Channel>", "still holds a `<...>` token", extra:="Channel=<Nope>" & vbCrLf))

    End Sub

    ''' <summary> An inline list in a <c> RegistryRoot= </c> gives one root per value </summary>
    <TestMethod()> Public Sub RegistryRoot_WithInlineList_FansOut()

        Dim browser = "[BrowserInfo: Edge]" & vbCrLf &
                      "Section=Edge" & vbCrLf &
                      "UserDataPath=%LocalAppData%\Microsoft\Edge\User Data" & vbCrLf &
                      "RegistryRoot=HKCU\Software\Microsoft\<Edge,Edge Beta>" & vbCrLf
        Dim output = GenerateClean(browser & "[EntryScaffold: Installer]" & vbCrLf & "RegKeyBase=<RegistryRoot>\Installer" & vbCrLf, False)

        CollectionAssert.AreEqual(New List(Of String) From {
            "RegKey1=HKCU\Software\Microsoft\Edge\Installer",
            "RegKey2=HKCU\Software\Microsoft\Edge Beta\Installer"
        }, PathKeysOf(output, "Edge Installer *"))

    End Sub

    ''' <summary>
    ''' A browser variable in a <c> UserDataPath= </c> gives one path per value, each with its
    ''' own DetectFile and FileKeys, exactly as separate keys would
    ''' </summary>
    <TestMethod()> Public Sub BrowserVariable_FansOutUserDataPaths()

        Dim browser = "[BrowserInfo: Thunderbird]" & vbCrLf &
                      "Section=Thunderbird" & vbCrLf &
                      "Packages=MZLA,ThunderbirdBeta" & vbCrLf &
                      "UserDataPath=%LocalAppData%\Packages\<Packages>\Profiles" & vbCrLf
        Dim output = GenerateClean(browser & "[EntryScaffold: History]" & vbCrLf & "FileKeyBase=<UserDataPath>\*|places*" & vbCrLf, True)

        CollectionAssert.AreEqual(New List(Of String) From {
            "Section=Thunderbird",
            "DetectFile1=%LocalAppData%\Packages\MZLA\Profiles",
            "FileKey1=%LocalAppData%\Packages\MZLA\Profiles\*|places*",
            "DetectFile2=%LocalAppData%\Packages\ThunderbirdBeta\Profiles",
            "FileKey2=%LocalAppData%\Packages\ThunderbirdBeta\Profiles\*|places*"
        }, KeysOf(output, "Thunderbird History *"))

    End Sub

    ''' <summary>
    ''' Browser variables resolve against each other, and the same variable can feed both a
    ''' path and a registry root
    ''' </summary>
    <TestMethod()> Public Sub BrowserVariable_NestedAndShared()

        Dim browser = "[BrowserInfo: Edge]" & vbCrLf &
                      "Section=Edge" & vbCrLf &
                      "Channel=<Product>,<Product> Beta" & vbCrLf &
                      "Product=Edge" & vbCrLf &
                      "UserDataPath=%LocalAppData%\Microsoft\<Channel>\User Data" & vbCrLf &
                      "RegistryRoot=HKCU\Software\Microsoft\<Channel>" & vbCrLf
        Dim output = GenerateClean(browser & "[EntryScaffold: Mixed]" & vbCrLf &
                                   "FileKeyBase=<UserDataPath>|*.log" & vbCrLf &
                                   "RegKeyBase=<RegistryRoot>\Installer" & vbCrLf, False)

        CollectionAssert.AreEqual(New List(Of String) From {
            "FileKey1=%LocalAppData%\Microsoft\Edge\User Data|*.log",
            "FileKey2=%LocalAppData%\Microsoft\Edge Beta\User Data|*.log",
            "RegKey1=HKCU\Software\Microsoft\Edge\Installer",
            "RegKey2=HKCU\Software\Microsoft\Edge Beta\Installer"
        }, PathKeysOf(output, "Edge Mixed *"))

    End Sub

    ''' <summary>
    ''' A misspelled reserved key becomes a variable nothing uses, which warns without failing
    ''' the run
    ''' </summary>
    <TestMethod()> Public Sub BrowserVariable_Unreferenced_Warns()

        Dim output As winapp2ool.iniFile = Nothing
        Dim log As List(Of String) = Nothing
        Dim browser = Chrome & "RegistyRoot=HKCU\Software\Typo" & vbCrLf

        Assert.AreEqual(0, ExitCodeAfter(Sub() output = Generate(browser & "[EntryScaffold: Logs]" & vbCrLf & "FileKeyBase=<UserDataPath>|*.log" & vbCrLf, False, log)))
        Assert.AreEqual(1, log.Where(Function(l) l.Contains("Variable 'RegistyRoot=' declared in [BrowserInfo: Chrome] but never referenced")).Count(), String.Join("; ", log))

    End Sub

    ''' <summary> A bare <c> Skip </c> line, with no value, skips the browser </summary>
    <TestMethod()> Public Sub BareSkip_SkipsTheBrowser()

        Dim output As String = Nothing

        Assert.AreEqual(0, RunBuild(Chrome & "Skip" & vbCrLf & Brave & "[EntryScaffold: Logs]" & vbCrLf & "FileKeyBase=<UserDataPath>|*.log" & vbCrLf, Nothing, output))
        Assert.IsFalse(output.Contains("[Chrome Logs *]"), output)
        StringAssert.Contains(output, "[Brave Logs *]")

    End Sub

    ''' <summary>
    ''' An empty <c> UserDataPath= </c> warns and is skipped without failing the run or
    ''' stopping the other browsers
    ''' </summary>
    <TestMethod()> Public Sub EmptyUserDataPath_IsSkipped()

        Dim output As String = Nothing
        Dim empty = "[BrowserInfo: Blank]" & vbCrLf & "Section=Blank" & vbCrLf & "UserDataPath=" & vbCrLf

        Assert.AreEqual(0, RunBuild(empty & Brave & Backups, Nothing, output))
        StringAssert.Contains(output, "[Brave Bookmark Backups *]")

    End Sub

    ''' <summary>
    ''' An empty <c> RegistryRoot= </c> gives the browser no root, so a scaffold that requires
    ''' one skips it rather than writing a RegKey with no root
    ''' </summary>
    <TestMethod()> Public Sub EmptyRegistryRoot_GivesNoRoot()

        Dim browser = Brave & "RegistryRoot=" & vbCrLf
        Dim output = GenerateClean(browser & "[EntryScaffold: Prefs]" & vbCrLf & "RequiresRegistryRoot" & vbCrLf & "RegKeyBase=<RegistryRoot>\Prefs" & vbCrLf, False)

        Assert.AreEqual(0, output.Count)

    End Sub

    ''' <summary> A user data path with no parent folder is skipped and fails the run </summary>
    <TestMethod()> Public Sub UserDataPath_WithNoParent_IsSkipped()

        Dim output As winapp2ool.iniFile = Nothing
        Dim log As List(Of String) = Nothing
        Dim browser = "[BrowserInfo: Flat]" & vbCrLf & "Section=Flat" & vbCrLf & "UserDataPath=NoFolders" & vbCrLf

        Assert.AreEqual(1, ExitCodeAfter(Sub() output = Generate(browser & Backups, False, log)))
        Assert.AreEqual(1, log.Where(Function(l) l.Contains("gives `NoFolders`, which has no parent folder")).Count(), String.Join("; ", log))

    End Sub

    ''' <summary> An empty <c> RegKeyBase= </c> is refused as a RegKeyBase, not a FileKeyBase </summary>
    <TestMethod()> Public Sub EmptyRegKeyBase_IsRefusedAsRegKeyBase()

        AssertRefused("RegKeyBase=", "RegKeyBase has no <RegistryRoot> token")

    End Sub

    ''' <summary>
    ''' Numbered keys declare one variable, since the name drops its digits, and the second
    ''' declaration replaces the first with a warning that doesn't fail the run
    ''' </summary>
    <TestMethod()> Public Sub RedeclaredVariable_WarnsAndLastWins()

        Dim output As winapp2ool.iniFile = Nothing
        Dim log As List(Of String) = Nothing

        Assert.AreEqual(0, ExitCodeAfter(Sub() output = Generate(Chrome & "[EntryScaffold: Caches]" & vbCrLf &
                                                                 "Dirs1=Cache" & vbCrLf &
                                                                 "Dirs2=Code Cache" & vbCrLf &
                                                                 "FileKeyBase=<UserDataPath>\*\<Dirs>|*" & vbCrLf, False, log)))
        Assert.AreEqual(1, log.Where(Function(l) l.Contains("declares `Dirs` more than once")).Count(), String.Join("; ", log))
        CollectionAssert.AreEqual(New List(Of String) From {
            "FileKey1=%LocalAppData%\Google\Chrome\User Data\*\Code Cache|*"
        }, PathKeysOf(output, "Chrome Caches *"))

    End Sub

    ' ----- Scaffold variables -----

    ''' <summary>
    ''' A scaffold variable fans a template out, and may be declared after the template that
    ''' uses it
    ''' </summary>
    <TestMethod()> Public Sub ScaffoldVariable_FansOutTemplate()

        Dim output = GenerateClean(Chrome & "[EntryScaffold: Caches]" & vbCrLf &
                                   "FileKeyBase=<UserDataPath>\*\<Dirs>|*|REMOVESELF" & vbCrLf &
                                   "Dirs=Cache,Code Cache" & vbCrLf, False)

        CollectionAssert.AreEqual(New List(Of String) From {
            "FileKey1=%LocalAppData%\Google\Chrome\User Data\*\Cache|*|REMOVESELF",
            "FileKey2=%LocalAppData%\Google\Chrome\User Data\*\Code Cache|*|REMOVESELF"
        }, PathKeysOf(output, "Chrome Caches *"))

    End Sub

    ''' <summary> One scaffold variable can feed several templates, including a RegKeyBase </summary>
    <TestMethod()> Public Sub ScaffoldVariable_SharedAcrossTemplates()

        Dim output = GenerateClean(Chrome & "[EntryScaffold: Beacon]" & vbCrLf &
                                   "Names=BlBeacon" & vbCrLf &
                                   "FileKeyBase=<UserDataPath>\<Names>|*" & vbCrLf &
                                   "RegKeyBase=<RegistryRoot>\<Names>|state" & vbCrLf, False)

        CollectionAssert.AreEqual(New List(Of String) From {
            "FileKey1=%LocalAppData%\Google\Chrome\User Data\BlBeacon|*",
            "RegKey1=HKCU\Software\Google\Chrome\BlBeacon|state",
            "RegKey2=HKCU\Software\Chromium\BlBeacon|state"
        }, PathKeysOf(output, "Chrome Beacon *"))

    End Sub

    ''' <summary> Scaffold variables nest, and <c> &lt;Name[N]&gt; </c> picks one value </summary>
    <TestMethod()> Public Sub ScaffoldVariable_NestedAndIndexed()

        Dim output = GenerateClean(Chrome & "[EntryScaffold: Caches]" & vbCrLf &
                                   "Base=Cache" & vbCrLf &
                                   "Dirs=<Base>,Code <Base>,GPU<Base>" & vbCrLf &
                                   "FileKeyBase=<UserDataPath>\*\<Dirs[2]>|*|REMOVESELF" & vbCrLf, False)

        CollectionAssert.AreEqual(New List(Of String) From {
            "FileKey1=%LocalAppData%\Google\Chrome\User Data\*\Code Cache|*|REMOVESELF"
        }, PathKeysOf(output, "Chrome Caches *"))

    End Sub

    ''' <summary> A scaffold variable works inside the Opera GX rewrite </summary>
    <TestMethod()> Public Sub ScaffoldVariable_InOperaGX()

        Dim output = GenerateClean(OperaGX & "[EntryScaffold: Caches]" & vbCrLf &
                                   "Dirs=Cache,Code Cache" & vbCrLf &
                                   "FileKeyBase=<UserDataPath>\*\<Dirs>|*|REMOVESELF" & vbCrLf, False)

        CollectionAssert.AreEqual(New List(Of String) From {
            $"FileKey1={GXPath}\Cache|*|REMOVESELF",
            $"FileKey2={GXPath}\Code Cache|*|REMOVESELF",
            $"FileKey3={GXPath}\_side_profiles\*\Cache|*|REMOVESELF",
            $"FileKey4={GXPath}\_side_profiles\*\Code Cache|*|REMOVESELF"
        }, PathKeysOf(output, "Opera GX Caches *"))

    End Sub

    ''' <summary> An index past the end of a scaffold variable's list refuses the template once </summary>
    <TestMethod()> Public Sub ScaffoldVariable_IndexPastEnd_IsRefused()

        Dim drop = AssertRefused("FileKeyBase=<UserDataPath>\*\<Dirs[3]>|*", "<Dirs[3]> is out of range", declarations:="Dirs=Cache,Code Cache" & vbCrLf)
        Assert.IsFalse(drop.Contains("dropping key;"), drop)

    End Sub

    ''' <summary> A scaffold variable with no values refuses the template that uses it </summary>
    <TestMethod()> Public Sub ScaffoldVariable_Empty_IsRefused()

        AssertRefused("FileKeyBase=<UserDataPath>\*\<Dirs>|*", "has no values", declarations:="Dirs=" & vbCrLf)

    End Sub

    ''' <summary> A scaffold variable holding a token nothing declares refuses the template </summary>
    <TestMethod()> Public Sub ScaffoldVariable_HoldingUnknownToken_IsRefused()

        AssertRefused("FileKeyBase=<UserDataPath>\*\<Dirs>|*", "still holds a `<...>` token", declarations:="Dirs=<Nope>" & vbCrLf)

    End Sub

    ''' <summary> A scaffold can't use a variable a browser declares </summary>
    <TestMethod()> Public Sub BrowserVariable_IsNotVisibleToScaffolds()

        Dim output As winapp2ool.iniFile = Nothing
        Dim log As List(Of String) = Nothing
        Dim browser = Chrome & "Channel=Beta" & vbCrLf

        Assert.AreEqual(1, ExitCodeAfter(Sub() output = Generate(browser & "[EntryScaffold: Logs]" & vbCrLf & "FileKeyBase=<UserDataPath>\<Channel>|*.log" & vbCrLf, False, log)))
        Assert.AreEqual(1, log.Where(Function(l) l.Contains("Unknown token `<Channel>`")).Count(), String.Join("; ", log))

    End Sub

    ''' <summary> Declaring a variable named after one of the four tokens fails the run </summary>
    <TestMethod()> Public Sub TokenNamedVariable_IsRefused()

        Dim output As winapp2ool.iniFile = Nothing
        Dim log As List(Of String) = Nothing

        Assert.AreEqual(1, ExitCodeAfter(Sub() output = Generate(Chrome & "[EntryScaffold: Logs]" & vbCrLf &
                                                                 "userdataparent=x" & vbCrLf &
                                                                 "FileKeyBase=<UserDataPath>|*.log" & vbCrLf, False, log)))
        Assert.AreEqual(1, log.Where(Function(l) l.Contains("declares `userdataparent`, which is a token, not a variable")).Count(), String.Join("; ", log))
        Assert.AreEqual(1, PathKeysOf(output, "Chrome Logs *").Count)

    End Sub

    ''' <summary> A scaffold variable nothing uses warns without failing the run </summary>
    <TestMethod()> Public Sub ScaffoldVariable_Unreferenced_Warns()

        Dim output As winapp2ool.iniFile = Nothing
        Dim log As List(Of String) = Nothing

        Assert.AreEqual(0, ExitCodeAfter(Sub() output = Generate(Chrome & "[EntryScaffold: Logs]" & vbCrLf &
                                                                 "RequiresRegistyRoot=True" & vbCrLf &
                                                                 "FileKeyBase=<UserDataPath>|*.log" & vbCrLf, False, log)))
        Assert.AreEqual(1, log.Where(Function(l) l.Contains("Variable 'RequiresRegistyRoot=' declared in [EntryScaffold: Logs]")).Count(), String.Join("; ", log))

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

    ''' <summary>
    ''' A section prefix matches in any case, with or without a space after its colon, and the
    ''' name after it is trimmed. A prefix with no name after it names nothing.
    ''' </summary>
    <TestMethod()> Public Sub NameAfterPrefix_IgnoresCaseAndSpacing()

        Assert.AreEqual("Chrome", winapp2ool.BrowserBuilder.nameAfterPrefix("BrowserInfo: Chrome", "BrowserInfo:"))
        Assert.AreEqual("Chrome", winapp2ool.BrowserBuilder.nameAfterPrefix("BrowserInfo:Chrome", "BrowserInfo:"))
        Assert.AreEqual("Chrome", winapp2ool.BrowserBuilder.nameAfterPrefix("browserinfo: Chrome", "BrowserInfo:"))
        Assert.AreEqual("Chrome", winapp2ool.BrowserBuilder.nameAfterPrefix("BROWSERINFO:   Chrome ", "BrowserInfo:"))
        Assert.AreEqual("", winapp2ool.BrowserBuilder.nameAfterPrefix("BrowserInfo: ", "BrowserInfo:"))
        Assert.AreEqual("", winapp2ool.BrowserBuilder.nameAfterPrefix("BrowserInfo:", "BrowserInfo:"))
        Assert.AreEqual("", winapp2ool.BrowserBuilder.nameAfterPrefix("BrowserInf: Chrome", "BrowserInfo:"))

    End Sub

    ''' <summary> A scaffold every test ruleset below shares </summary>
    Private Const Backups As String =
        "[EntryScaffold: Bookmark Backups]" & vbCrLf &
        "FileKeyBase=<UserDataPath>\*|Bookmarks.bak" & vbCrLf

    ''' <summary>
    ''' Runs a whole BrowserBuilder build over a scratch source directory holding the given
    ''' rulesets, and returns the exit code it leaves and the output file's text, or
    ''' <c> Nothing </c> when no output was written
    ''' </summary>
    '''
    ''' <param name="chromium">The text of <c> chromium.ini </c>, or <c> Nothing </c> to leave it out</param>
    '''
    ''' <param name="gecko">The text of <c> gecko.ini </c>, or <c> Nothing </c> to leave it out</param>
    '''
    ''' <param name="output">Set to the text of <c> browsers.ini </c>, or <c> Nothing </c></param>
    '''
    ''' <param name="makeDir">Indicates whether the source directory exists at all</param>
    Private Shared Function RunBuild(chromium As String, gecko As String, ByRef output As String,
                                     Optional makeDir As Boolean = True) As Integer

        Dim dir = IO.Path.Combine(IO.Path.GetTempPath(), $"winapp2ool-browserbuilder-{Guid.NewGuid():N}")
        Dim savedSource = winapp2ool.BuilderFile1
        Dim savedTarget = winapp2ool.BuilderFile2
        Dim savedSuppress = winapp2ool.SuppressOutput
        Dim savedExitCode = Environment.ExitCode

        Try

            If makeDir Then IO.Directory.CreateDirectory(dir)
            If chromium IsNot Nothing Then IO.File.WriteAllText(IO.Path.Combine(dir, "chromium.ini"), chromium)
            If gecko IsNot Nothing Then IO.File.WriteAllText(IO.Path.Combine(dir, "gecko.ini"), gecko)

            Dim outDir = IO.Path.Combine(IO.Path.GetTempPath(), $"winapp2ool-browserbuilder-out-{Guid.NewGuid():N}")
            IO.Directory.CreateDirectory(outDir)

            Try

                winapp2ool.BuilderFile1 = New winapp2ool.iniFileChooser(dir, "", "")
                winapp2ool.BuilderFile2 = New winapp2ool.iniFileChooser(outDir, "browsers.ini", mustExist:=False)
                winapp2ool.SuppressOutput = True
                Environment.ExitCode = 0

                winapp2ool.BrowserBuilder.initBrowserBuilder()

                Dim outPath = IO.Path.Combine(outDir, "browsers.ini")
                output = If(IO.File.Exists(outPath), IO.File.ReadAllText(outPath), Nothing)
                Return Environment.ExitCode

            Finally

                IO.Directory.Delete(outDir, recursive:=True)

            End Try

        Finally

            winapp2ool.BuilderFile1 = savedSource
            winapp2ool.BuilderFile2 = savedTarget
            winapp2ool.SuppressOutput = savedSuppress
            Environment.ExitCode = savedExitCode
            If IO.Directory.Exists(dir) Then IO.Directory.Delete(dir, recursive:=True)

        End Try

    End Function

    ''' <summary> A source directory with only <c> chromium.ini </c> builds and leaves the exit code at 0 </summary>
    <TestMethod()> Public Sub OnlyChromium_BuildsWithoutFailing()

        Dim output As String = Nothing
        Assert.AreEqual(0, RunBuild(Chrome & Backups, Nothing, output))
        Assert.IsNotNull(output)
        StringAssert.Contains(output, "[Chrome Bookmark Backups *]")

    End Sub

    ''' <summary> A source directory with only <c> gecko.ini </c> builds and leaves the exit code at 0 </summary>
    <TestMethod()> Public Sub OnlyGecko_BuildsWithoutFailing()

        Dim output As String = Nothing
        Assert.AreEqual(0, RunBuild(Nothing, Firefox & Backups, output))
        Assert.IsNotNull(output)
        StringAssert.Contains(output, "[Firefox Bookmark Backups *]")

    End Sub

    ''' <summary> A source directory with neither ruleset writes nothing and fails the run </summary>
    <TestMethod()> Public Sub NoRulesets_WritesNothingAndFails()

        Dim output As String = Nothing
        Assert.AreEqual(1, RunBuild(Nothing, Nothing, output))
        Assert.IsNull(output)

    End Sub

    ''' <summary> A source directory that doesn't exist writes nothing and fails the run </summary>
    <TestMethod()> Public Sub MissingSourceDirectory_WritesNothingAndFails()

        Dim output As String = Nothing
        Assert.AreEqual(1, RunBuild(Nothing, Nothing, output, makeDir:=False))
        Assert.IsNull(output)

    End Sub

    ''' <summary>
    ''' Section headers with no space after the colon, or in another case, build with the name
    ''' intact and leave the exit code at 0
    ''' </summary>
    <TestMethod()> Public Sub LooseSectionHeaders_Build()

        Dim loose = Chrome.Replace("[BrowserInfo: Chrome]", "[BrowserInfo:Chrome]") &
                    Brave.Replace("[BrowserInfo: Brave]", "[browserinfo: Brave]") &
                    Backups.Replace("[EntryScaffold: Bookmark Backups]", "[ENTRYSCAFFOLD:Bookmark Backups]")
        Dim output As String = Nothing
        Assert.AreEqual(0, RunBuild(loose, Nothing, output))
        StringAssert.Contains(output, "[Chrome Bookmark Backups *]")
        StringAssert.Contains(output, "[Brave Bookmark Backups *]")

    End Sub

    ''' <summary>
    ''' A misspelled prefix, and a prefix with no name after it, are ignored as invalid and fail
    ''' the run, and the valid browser beside them still builds
    ''' </summary>
    <TestMethod()> Public Sub InvalidSections_AreIgnoredAndFail()

        Dim misspelled = Chrome.Replace("[BrowserInfo: Chrome]", "[BrowserInf: Chrome]")
        Dim output As String = Nothing
        Assert.AreEqual(1, RunBuild(misspelled & "[BrowserInfo: ]" & vbCrLf & Brave & Backups, Nothing, output))
        StringAssert.Contains(output, "[Brave Bookmark Backups *]")
        Assert.IsFalse(output.Contains("[Chrome "), output)

    End Sub

End Class
