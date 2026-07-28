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
''' Tests for Transmute's global operations: the <c> [*] </c> sentinel section (global
''' add / remove / replace of keys across every base section) and <c> [*Map: label] </c>
''' key mapping rules (KeyType+Value match → whole-key replacement, including the Name,
''' one-to-many via numbered Replace keys). <br />
''' Covers the mode refusal table, numbered-key add refusal, first-match-wins/no-re-match
''' semantics, in-place ordering, one-to-many expansion, malformed and zero-hit rules,
''' global-then-named processing order, the <c> %EntryName% </c> token, wildcard
''' <c> Match=...=* </c> fallback rules, and the <c> RecognizeGlobalSections </c> opt-out
''' </summary>
<TestClass()> Public Class TransmuteTests

    ''' <summary>
    ''' Helper: parse an <c> iniFile2 </c> from literal ini text
    ''' </summary>
    Private Shared Function MakeIni(text As String) As winapp2ool.iniFile2

        Dim bytes = Encoding.UTF8.GetBytes(text)
        Using ms As New IO.MemoryStream(bytes)
            Using reader As New IO.StreamReader(ms)
                Return winapp2ool.iniFile2.FromStream(reader, "", "test.ini")
            End Using
        End Using

    End Function

    ''' <summary>
    ''' Helper: run a transmutation of <paramref name="sourceText"/> against
    ''' <paramref name="baseText"/> in memory and return the mutated base file
    ''' </summary>
    Private Shared Function RunTransmute(baseText As String,
                                         sourceText As String,
                                         mode As winapp2ool.TransmuteMode,
                                Optional replaceMode As winapp2ool.ReplaceMode = winapp2ool.ReplaceMode.ByKey,
                                Optional removeMode As winapp2ool.RemoveMode = winapp2ool.RemoveMode.ByKey,
                                Optional removeKeyMode As winapp2ool.RemoveKeyMode = winapp2ool.RemoveKeyMode.ByName) As winapp2ool.iniFile2

        Dim baseFile = MakeIni(baseText)
        Dim sourceFile = MakeIni(sourceText)
        Dim outputFile = winapp2ool.iniFile2.Empty("", "out.ini")
        Dim menuOutput As New winapp2ool.MenuSection

        winapp2ool.RemoteTransmute(baseFile, sourceFile, outputFile, False, menuOutput,
                                   mode, replaceMode, removeMode, removeKeyMode, skipFormat:=True)

        Return baseFile

    End Function

    ''' <summary>
    ''' Sentinel recognition is on by default; reset it before every test since
    ''' the opt-out test disables it
    ''' </summary>
    <TestInitialize()> Public Sub ResetGlobalRecognition()

        winapp2ool.transmuteSettings.RecognizeGlobalSections = True

    End Sub

    ''' <summary>
    ''' A [*] section in Add mode adds its unnumbered keys to every base section,
    ''' and the [*] section itself is never added to the base file
    ''' </summary>
    <TestMethod()> Public Sub GlobalAdd_AddsUnnumberedKeyToEverySection()

        Dim result = RunTransmute(
            "[Alpha *]" & vbCrLf & "Detect=HKCU\Software\Alpha" & vbCrLf &
            "[Beta *]" & vbCrLf & "Detect=HKCU\Software\Beta" & vbCrLf,
            "[*]" & vbCrLf & "Author=Winapp2.ini Project" & vbCrLf,
            winapp2ool.TransmuteMode.Add)

        Assert.AreEqual(2, result.Count)
        Assert.IsFalse(result.Contains("*"))

        For Each section In result
            Assert.AreEqual("Winapp2.ini Project", section.GetKey("Author").Value)
            Assert.AreEqual(2, section.Keys.Count)
        Next

    End Sub

    ''' <summary>
    ''' Numbered keys are refused by global Add: adding FileKey1= to every section
    ''' would create instant duplicates and standalone Transmute has no renumber pass
    ''' </summary>
    <TestMethod()> Public Sub GlobalAdd_RefusesNumberedKey()

        Dim result = RunTransmute(
            "[Alpha *]" & vbCrLf & "Detect=HKCU\Software\Alpha" & vbCrLf,
            "[*]" & vbCrLf & "FileKey1=%AppData%\Alpha|*.log" & vbCrLf & "Author=Winapp2.ini Project" & vbCrLf,
            winapp2ool.TransmuteMode.Add)

        Dim section = result.GetSection("Alpha *")
        Assert.IsFalse(section.HasKey("FileKey1"))
        ' The unnumbered key in the same [*] section is still applied
        Assert.IsTrue(section.HasKey("Author"))

    End Sub

    ''' <summary>
    ''' Global Remove ByName removes every key with a matching Name from every section,
    ''' regardless of value, leaving other keys untouched
    ''' </summary>
    <TestMethod()> Public Sub GlobalRemove_ByName_RemovesFromEverySection()

        Dim result = RunTransmute(
            "[Alpha *]" & vbCrLf & "Section=Games" & vbCrLf & "Detect=HKCU\Software\Alpha" & vbCrLf &
            "[Beta *]" & vbCrLf & "Section=Utilities" & vbCrLf & "Detect=HKCU\Software\Beta" & vbCrLf &
            "[Gamma *]" & vbCrLf & "LangSecRef=3021" & vbCrLf,
            "[*]" & vbCrLf & "Section=anything" & vbCrLf,
            winapp2ool.TransmuteMode.Remove,
            removeKeyMode:=winapp2ool.RemoveKeyMode.ByName)

        Assert.IsFalse(result.GetSection("Alpha *").HasKey("Section"))
        Assert.IsFalse(result.GetSection("Beta *").HasKey("Section"))
        Assert.IsTrue(result.GetSection("Alpha *").HasKey("Detect"))
        Assert.IsTrue(result.GetSection("Gamma *").HasKey("LangSecRef"))

    End Sub

    ''' <summary>
    ''' Global Remove ByValue removes keys matching KeyType+Value from every section:
    ''' key numbering is ignored on both sides and non-matching values survive
    ''' </summary>
    <TestMethod()> Public Sub GlobalRemove_ByValue_IgnoresNumberingAndKeepsOtherValues()

        Dim result = RunTransmute(
            "[Alpha *]" & vbCrLf & "DetectFile1=%LocalAppData%\Chromium" & vbCrLf & "DetectFile2=%AppData%\Alpha" & vbCrLf &
            "[Beta *]" & vbCrLf & "DetectFile=%LocalAppData%\Chromium" & vbCrLf,
            "[*]" & vbCrLf & "DetectFile=%LocalAppData%\Chromium" & vbCrLf,
            winapp2ool.TransmuteMode.Remove,
            removeKeyMode:=winapp2ool.RemoveKeyMode.ByValue)

        Dim alpha = result.GetSection("Alpha *")
        Assert.AreEqual(1, alpha.Keys.Count)
        Assert.AreEqual("%AppData%\Alpha", alpha.GetKey("DetectFile2").Value)
        Assert.AreEqual(0, result.GetSection("Beta *").Keys.Count)

    End Sub

    ''' <summary>
    ''' Global Replace ByKey sets the value of matching key Names in every section that
    ''' has them, and does not create the key in sections that lack it
    ''' </summary>
    <TestMethod()> Public Sub GlobalReplace_ByKey_ReplacesOnlyWherePresent()

        Dim result = RunTransmute(
            "[Alpha *]" & vbCrLf & "LangSecRef=3021" & vbCrLf &
            "[Beta *]" & vbCrLf & "Section=Games" & vbCrLf,
            "[*]" & vbCrLf & "LangSecRef=3005" & vbCrLf,
            winapp2ool.TransmuteMode.Replace,
            replaceMode:=winapp2ool.ReplaceMode.ByKey)

        Assert.AreEqual("3005", result.GetSection("Alpha *").GetKey("LangSecRef").Value)
        Assert.IsFalse(result.GetSection("Beta *").HasKey("LangSecRef"))
        Assert.AreEqual("Games", result.GetSection("Beta *").GetKey("Section").Value)

    End Sub

    ''' <summary>
    ''' A [*] section under Remove BySection is refused: it would remove every
    ''' section from the base file
    ''' </summary>
    <TestMethod()> Public Sub GlobalRemove_BySection_IsRefused()

        Dim result = RunTransmute(
            "[Alpha *]" & vbCrLf & "Detect=HKCU\Software\Alpha" & vbCrLf,
            "[*]" & vbCrLf & "Section=anything" & vbCrLf,
            winapp2ool.TransmuteMode.Remove,
            removeMode:=winapp2ool.RemoveMode.BySection)

        Assert.AreEqual(1, result.Count)
        Assert.IsTrue(result.GetSection("Alpha *").HasKey("Detect"))

    End Sub

    ''' <summary>
    ''' A [*] section under Replace BySection is refused as incoherent
    ''' </summary>
    <TestMethod()> Public Sub GlobalReplace_BySection_IsRefused()

        Dim result = RunTransmute(
            "[Alpha *]" & vbCrLf & "Detect=HKCU\Software\Alpha" & vbCrLf,
            "[*]" & vbCrLf & "Detect=HKCU\Software\Replaced" & vbCrLf,
            winapp2ool.TransmuteMode.Replace,
            replaceMode:=winapp2ool.ReplaceMode.BySection)

        Assert.AreEqual("HKCU\Software\Alpha", result.GetSection("Alpha *").GetKey("Detect").Value)

    End Sub

    ''' <summary>
    ''' Globals are processed before named sections, so a named section can refine
    ''' the result of a global operation in the same run
    ''' </summary>
    <TestMethod()> Public Sub GlobalThenNamed_NamedSectionRefinesGlobalResult()

        Dim result = RunTransmute(
            "[Alpha *]" & vbCrLf & "LangSecRef=3021" & vbCrLf &
            "[Beta *]" & vbCrLf & "LangSecRef=3021" & vbCrLf,
            "[*]" & vbCrLf & "LangSecRef=3005" & vbCrLf &
            "[Beta *]" & vbCrLf & "LangSecRef=3029" & vbCrLf,
            winapp2ool.TransmuteMode.Replace,
            replaceMode:=winapp2ool.ReplaceMode.ByKey)

        Assert.AreEqual("3005", result.GetSection("Alpha *").GetKey("LangSecRef").Value)
        Assert.AreEqual("3029", result.GetSection("Beta *").GetKey("LangSecRef").Value)

    End Sub

    ''' <summary>
    ''' A basic [*Map:] rule replaces the whole key — including its Name — in place,
    ''' preserving the section's key order and updating the type index
    ''' </summary>
    <TestMethod()> Public Sub Map_BasicRename_ReplacesKeyInPlace()

        Dim result = RunTransmute(
            "[Alpha *]" & vbCrLf & "Detect=HKCU\Software\Alpha" & vbCrLf & "Section=Google Chrome Web Browser" & vbCrLf & "FileKey1=%AppData%\Alpha|*.log" & vbCrLf &
            "[Beta *]" & vbCrLf & "Section=Games" & vbCrLf,
            "[*Map: Chrome]" & vbCrLf & "Match=Section=Google Chrome Web Browser" & vbCrLf & "Replace=LangSecRef=3029" & vbCrLf,
            winapp2ool.TransmuteMode.Replace,
            replaceMode:=winapp2ool.ReplaceMode.ByKey)

        Dim alpha = result.GetSection("Alpha *")
        Dim alphaKeys = alpha.Keys.ToList()

        ' Replaced in place: same ordinal position, new Name and Value
        Assert.AreEqual(3, alphaKeys.Count)
        Assert.AreEqual("LangSecRef", alphaKeys(1).Name)
        Assert.AreEqual("3029", alphaKeys(1).Value)

        ' The type index reflects the replacement
        Assert.AreEqual(1, alpha.Keys.GetByType("LangSecRef").Count)
        Assert.AreEqual(0, alpha.Keys.GetByType("Section").Count)

        ' Non-matching values are untouched
        Assert.AreEqual("Games", result.GetSection("Beta *").GetKey("Section").Value)

    End Sub

    ''' <summary>
    ''' Numbered Match criteria (Match1=, Match2=, ...) express many-to-one mappings
    ''' </summary>
    <TestMethod()> Public Sub Map_NumberedMatches_ManyToOne()

        Dim result = RunTransmute(
            "[Alpha *]" & vbCrLf & "LangSecRef=3005" & vbCrLf &
            "[Beta *]" & vbCrLf & "LangSecRef=3021" & vbCrLf &
            "[Gamma *]" & vbCrLf & "LangSecRef=3029" & vbCrLf,
            "[*Map: CC7 apps tag]" & vbCrLf & "Match1=LangSecRef=3005" & vbCrLf & "Match2=LangSecRef=3021" & vbCrLf & "Replace=Tags=ccapps" & vbCrLf,
            winapp2ool.TransmuteMode.Replace,
            replaceMode:=winapp2ool.ReplaceMode.ByKey)

        Assert.AreEqual("ccapps", result.GetSection("Alpha *").GetKey("Tags").Value)
        Assert.AreEqual("ccapps", result.GetSection("Beta *").GetKey("Tags").Value)
        Assert.IsFalse(result.GetSection("Gamma *").HasKey("Tags"))
        Assert.AreEqual("3029", result.GetSection("Gamma *").GetKey("LangSecRef").Value)

    End Sub

    ''' <summary>
    ''' Match criteria compare KeyTypes with numbers stripped from both sides, so a
    ''' Match=FileKey=... criterion matches FileKey3= with the same value
    ''' </summary>
    <TestMethod()> Public Sub Map_MatchesNumberedBaseKeysByType()

        Dim result = RunTransmute(
            "[Alpha *]" & vbCrLf & "FileKey3=%AppData%\Alpha|*.log" & vbCrLf,
            "[*Map: relocate]" & vbCrLf & "Match=FileKey=%AppData%\Alpha|*.log" & vbCrLf & "Replace=FileKey=%LocalAppData%\Alpha|*.log" & vbCrLf,
            winapp2ool.TransmuteMode.Replace,
            replaceMode:=winapp2ool.ReplaceMode.ByKey)

        Dim keys = result.GetSection("Alpha *").Keys.ToList()
        Assert.AreEqual("FileKey", keys(0).Name)
        Assert.AreEqual("%LocalAppData%\Alpha|*.log", keys(0).Value)

    End Sub

    ''' <summary>
    ''' Rules are single-pass with first-match-wins semantics: a key replaced by an
    ''' earlier rule is never re-matched by a later rule in the same run, while keys
    ''' that originally match the later rule are still converted
    ''' </summary>
    <TestMethod()> Public Sub Map_FirstMatchWins_ReplacedKeysNotReMatched()

        Dim result = RunTransmute(
            "[Alpha *]" & vbCrLf & "Section=Google Chrome Web Browser" & vbCrLf &
            "[Beta *]" & vbCrLf & "LangSecRef=3029" & vbCrLf,
            "[*Map: first]" & vbCrLf & "Match=Section=Google Chrome Web Browser" & vbCrLf & "Replace=LangSecRef=3029" & vbCrLf &
            "[*Map: second]" & vbCrLf & "Match=LangSecRef=3029" & vbCrLf & "Replace=Tags=Google,Chrome,Browser" & vbCrLf,
            winapp2ool.TransmuteMode.Replace,
            replaceMode:=winapp2ool.ReplaceMode.ByKey)

        ' Alpha's key was replaced by the first rule and must NOT be re-matched by the second
        Assert.AreEqual("3029", result.GetSection("Alpha *").GetKey("LangSecRef").Value)
        Assert.IsFalse(result.GetSection("Alpha *").HasKey("Tags"))

        ' Beta's key originally matched the second rule and is converted
        Assert.AreEqual("Google,Chrome,Browser", result.GetSection("Beta *").GetKey("Tags").Value)
        Assert.IsFalse(result.GetSection("Beta *").HasKey("LangSecRef"))

    End Sub

    ''' <summary>
    ''' [*Map:] rules are only recognized in Replace ByKey mode; under any other mode
    ''' they are skipped without touching the base file (and without being treated
    ''' as removal criteria)
    ''' </summary>
    <TestMethod()> Public Sub Map_OutsideReplaceByKey_IsSkipped()

        Dim result = RunTransmute(
            "[Alpha *]" & vbCrLf & "Section=Google Chrome Web Browser" & vbCrLf,
            "[*Map: Chrome]" & vbCrLf & "Match=Section=Google Chrome Web Browser" & vbCrLf & "Replace=LangSecRef=3029" & vbCrLf,
            winapp2ool.TransmuteMode.Remove,
            removeKeyMode:=winapp2ool.RemoveKeyMode.ByName)

        Assert.AreEqual("Google Chrome Web Browser", result.GetSection("Alpha *").GetKey("Section").Value)

    End Sub

    ''' <summary>
    ''' Malformed rules are skipped individually: a rule with no Replace=, and a rule
    ''' whose Match value is not a Name=Value pair, leave the base file untouched while
    ''' a well-formed sibling rule in the same source file still applies
    ''' </summary>
    <TestMethod()> Public Sub Map_MalformedRules_SkippedIndividually()

        Dim result = RunTransmute(
            "[Alpha *]" & vbCrLf & "Section=Games" & vbCrLf & "LangSecRef=3021" & vbCrLf,
            "[*Map: no replace]" & vbCrLf & "Match=Section=Games" & vbCrLf &
            "[*Map: bad match]" & vbCrLf & "Match=NotAKeyValuePair" & vbCrLf & "Replace=Tags=broken" & vbCrLf &
            "[*Map: good]" & vbCrLf & "Match=LangSecRef=3021" & vbCrLf & "Replace=Tags=ccapps" & vbCrLf,
            winapp2ool.TransmuteMode.Replace,
            replaceMode:=winapp2ool.ReplaceMode.ByKey)

        Dim alpha = result.GetSection("Alpha *")

        ' Malformed rules changed nothing
        Assert.AreEqual("Games", alpha.GetKey("Section").Value)
        ' The well-formed rule still applied
        Assert.AreEqual("ccapps", alpha.GetKey("Tags").Value)
        Assert.IsFalse(alpha.HasKey("LangSecRef"))

    End Sub

    ''' <summary>
    ''' A well-formed rule that matches nothing leaves the file untouched and does not throw
    ''' (it emits a stale-rule warning to the menu)
    ''' </summary>
    <TestMethod()> Public Sub Map_ZeroHits_LeavesFileUntouched()

        Dim result = RunTransmute(
            "[Alpha *]" & vbCrLf & "Section=Games" & vbCrLf,
            "[*Map: stale]" & vbCrLf & "Match=Section=No Such Value" & vbCrLf & "Replace=LangSecRef=1234" & vbCrLf,
            winapp2ool.TransmuteMode.Replace,
            replaceMode:=winapp2ool.ReplaceMode.ByKey)

        Assert.AreEqual("Games", result.GetSection("Alpha *").GetKey("Section").Value)
        Assert.IsFalse(result.GetSection("Alpha *").HasKey("LangSecRef"))

    End Sub

    ''' <summary>
    ''' A [*Map:] rule with numbered Replace keys (Replace1=, Replace2=, ...) expands one
    ''' matched key into several: the first replacement takes the matched key's ordinal
    ''' position, the rest follow it immediately, and later keys shift down intact
    ''' </summary>
    <TestMethod()> Public Sub Map_OneToMany_ExpandsKeyInPlace()

        Dim result = RunTransmute(
            "[Alpha *]" & vbCrLf & "Section=Games" & vbCrLf & "DetectFile=%UserProfile%\Documents\My Games\Borderlands*" & vbCrLf & "FileKey1=%UserProfile%\Documents\My Games\Borderlands|*.log" & vbCrLf,
            "[*Map: Borderlands DetectFile]" & vbCrLf &
            "Match=DetectFile=%UserProfile%\Documents\My Games\Borderlands*" & vbCrLf &
            "Replace1=DetectFile1=%UserProfile%\Documents\My Games\Borderlands" & vbCrLf &
            "Replace2=DetectFile2=%UserProfile%\Documents\My Games\Borderlands 2" & vbCrLf &
            "Replace3=DetectFile3=%UserProfile%\Documents\My Games\Borderlands 3" & vbCrLf,
            winapp2ool.TransmuteMode.Replace,
            replaceMode:=winapp2ool.ReplaceMode.ByKey)

        Dim alpha = result.GetSection("Alpha *")
        Dim keys = alpha.Keys.ToList()

        ' 3 original keys - 1 matched + 3 replacements = 5
        Assert.AreEqual(5, keys.Count)

        ' First replacement at the matched key's ordinal, the rest immediately after
        Assert.AreEqual("DetectFile1", keys(1).Name)
        Assert.AreEqual("%UserProfile%\Documents\My Games\Borderlands", keys(1).Value)
        Assert.AreEqual("DetectFile2", keys(2).Name)
        Assert.AreEqual("%UserProfile%\Documents\My Games\Borderlands 2", keys(2).Value)
        Assert.AreEqual("DetectFile3", keys(3).Name)
        Assert.AreEqual("%UserProfile%\Documents\My Games\Borderlands 3", keys(3).Value)

        ' Keys following the matched key shift down intact
        Assert.AreEqual("FileKey1", keys(4).Name)

        ' The type index reflects the expansion
        Assert.AreEqual(3, alpha.Keys.GetByType("DetectFile").Count)

    End Sub

    ''' <summary>
    ''' Replacement keys inserted by a one-to-many rule are never evaluated in the same
    ''' run: a later rule matching an inserted key's KeyType+Value must not fire
    ''' </summary>
    <TestMethod()> Public Sub Map_OneToMany_InsertedKeysNotReMatched()

        Dim result = RunTransmute(
            "[Alpha *]" & vbCrLf & "DetectFile=%LocalAppData%\Google\Chrome*" & vbCrLf,
            "[*Map: expand]" & vbCrLf &
            "Match=DetectFile=%LocalAppData%\Google\Chrome*" & vbCrLf &
            "Replace1=DetectFile1=%LocalAppData%\Google\Chrome" & vbCrLf &
            "Replace2=DetectFile2=%LocalAppData%\Google\Chrome Beta" & vbCrLf &
            "[*Map: would re-match]" & vbCrLf &
            "Match=DetectFile=%LocalAppData%\Google\Chrome Beta" & vbCrLf &
            "Replace=Tags=must-not-appear" & vbCrLf,
            winapp2ool.TransmuteMode.Replace,
            replaceMode:=winapp2ool.ReplaceMode.ByKey)

        Dim alpha = result.GetSection("Alpha *")

        ' The inserted DetectFile2 survives untouched; the second rule never fired
        Assert.AreEqual("%LocalAppData%\Google\Chrome Beta", alpha.GetKey("DetectFile2").Value)
        Assert.IsFalse(alpha.HasKey("Tags"))

    End Sub

    ''' <summary>
    ''' A single one-to-many rule expands the same matched wildcard independently in
    ''' every section that carries it
    ''' </summary>
    <TestMethod()> Public Sub Map_OneToMany_ExpandsInMultipleSections()

        Dim result = RunTransmute(
            "[Alpha *]" & vbCrLf & "DetectFile=%LocalAppData%\BraveSoftware\Brave-Browser*" & vbCrLf &
            "[Beta *]" & vbCrLf & "DetectFile=%LocalAppData%\BraveSoftware\Brave-Browser*" & vbCrLf &
            "[Gamma *]" & vbCrLf & "DetectFile=%AppData%\Unrelated" & vbCrLf,
            "[*Map: Brave DetectFile]" & vbCrLf &
            "Match=DetectFile=%LocalAppData%\BraveSoftware\Brave-Browser*" & vbCrLf &
            "Replace1=DetectFile1=%LocalAppData%\BraveSoftware\Brave-Browser" & vbCrLf &
            "Replace2=DetectFile2=%LocalAppData%\BraveSoftware\Brave-Browser-Beta" & vbCrLf &
            "Replace3=DetectFile3=%LocalAppData%\BraveSoftware\Brave-Browser-Nightly" & vbCrLf,
            winapp2ool.TransmuteMode.Replace,
            replaceMode:=winapp2ool.ReplaceMode.ByKey)

        For Each name In {"Alpha *", "Beta *"}
            Dim section = result.GetSection(name)
            Assert.AreEqual(3, section.Keys.Count)
            Assert.AreEqual(3, section.Keys.GetByType("DetectFile").Count)
            Assert.AreEqual("%LocalAppData%\BraveSoftware\Brave-Browser-Nightly", section.GetKey("DetectFile3").Value)
        Next

        ' Non-matching sections are untouched
        Assert.AreEqual(1, result.GetSection("Gamma *").Keys.Count)
        Assert.AreEqual("%AppData%\Unrelated", result.GetSection("Gamma *").GetKey("DetectFile").Value)

    End Sub

    ''' <summary>
    ''' Two one-to-many rules firing on different keys of the same section expand
    ''' independently without disturbing each other's insertions (the Dell Stage Suite
    ''' shape: two wildcards under different roots in one entry)
    ''' </summary>
    <TestMethod()> Public Sub Map_TwoOneToManyRules_SameSection()

        Dim result = RunTransmute(
            "[Alpha *]" & vbCrLf & "Section=Games" & vbCrLf & "DetectFile1=%AppData%\Dell\*Stage" & vbCrLf & "DetectFile2=%LocalAppData%\Dell\*Stage" & vbCrLf,
            "[*Map: Dell roaming]" & vbCrLf &
            "Match=DetectFile=%AppData%\Dell\*Stage" & vbCrLf &
            "Replace1=DetectFile1=%AppData%\Dell\Dell Stage" & vbCrLf &
            "Replace2=DetectFile2=%AppData%\Dell\MusicStage" & vbCrLf &
            "[*Map: Dell local]" & vbCrLf &
            "Match=DetectFile=%LocalAppData%\Dell\*Stage" & vbCrLf &
            "Replace=DetectFile3=%LocalAppData%\Dell\VideoStage" & vbCrLf,
            winapp2ool.TransmuteMode.Replace,
            replaceMode:=winapp2ool.ReplaceMode.ByKey)

        Dim keys = result.GetSection("Alpha *").Keys.ToList()

        Assert.AreEqual(4, keys.Count)
        Assert.AreEqual("Section", keys(0).Name)
        Assert.AreEqual("%AppData%\Dell\Dell Stage", keys(1).Value)
        Assert.AreEqual("%AppData%\Dell\MusicStage", keys(2).Value)
        Assert.AreEqual("%LocalAppData%\Dell\VideoStage", keys(3).Value)

    End Sub

    ''' <summary>
    ''' A numbered Replace key whose value is not a Name=Value pair makes the rule
    ''' malformed: it is skipped whole while a well-formed sibling rule still applies
    ''' </summary>
    <TestMethod()> Public Sub Map_MalformedNumberedReplace_SkippedIndividually()

        Dim result = RunTransmute(
            "[Alpha *]" & vbCrLf & "Section=Games" & vbCrLf & "LangSecRef=3021" & vbCrLf,
            "[*Map: bad numbered replace]" & vbCrLf & "Match=Section=Games" & vbCrLf & "Replace1=Tags=ok" & vbCrLf & "Replace2=NotAKeyValuePair" & vbCrLf &
            "[*Map: good]" & vbCrLf & "Match=LangSecRef=3021" & vbCrLf & "Replace=Tags=ccapps" & vbCrLf,
            winapp2ool.TransmuteMode.Replace,
            replaceMode:=winapp2ool.ReplaceMode.ByKey)

        Dim alpha = result.GetSection("Alpha *")

        ' The malformed rule changed nothing - not even its well-formed Replace1
        Assert.AreEqual("Games", alpha.GetKey("Section").Value)
        ' The well-formed sibling rule still applied
        Assert.AreEqual("ccapps", alpha.GetKey("Tags").Value)
        Assert.IsFalse(alpha.HasKey("LangSecRef"))

    End Sub

    ''' <summary>
    ''' With RecognizeGlobalSections disabled, [*] and [*Map:] fall through to normal
    ''' section name matching: in Add mode they are added to the base file as literal sections
    ''' </summary>
    <TestMethod()> Public Sub NoGlobal_SentinelsTreatedAsOrdinarySections()

        winapp2ool.transmuteSettings.RecognizeGlobalSections = False

        Dim result = RunTransmute(
            "[Alpha *]" & vbCrLf & "Detect=HKCU\Software\Alpha" & vbCrLf,
            "[*]" & vbCrLf & "Author=Winapp2.ini Project" & vbCrLf &
            "[*Map: Chrome]" & vbCrLf & "Match=Section=Games" & vbCrLf & "Replace=LangSecRef=3029" & vbCrLf,
            winapp2ool.TransmuteMode.Add)

        ' The sentinels were added as literal sections instead of being interpreted
        Assert.IsTrue(result.Contains("*"))
        Assert.IsTrue(result.Contains("*Map: Chrome"))
        ' And no global application took place
        Assert.IsFalse(result.GetSection("Alpha *").HasKey("Author"))

    End Sub

    ''' <summary>
    ''' The %EntryName% token in a [*] key value is replaced with each receiving section's
    ''' name — including the trailing " *" of winapp2 entry names — so a single global key
    ''' yields a distinct value per section
    ''' </summary>
    <TestMethod()> Public Sub GlobalAdd_EntryNameToken_SubstitutesSectionName()

        Dim result = RunTransmute(
            "[Alpha *]" & vbCrLf & "Detect=HKCU\Software\Alpha" & vbCrLf &
            "[Beta *]" & vbCrLf & "Detect=HKCU\Software\Beta" & vbCrLf,
            "[*]" & vbCrLf & "ID=%EntryName%" & vbCrLf,
            winapp2ool.TransmuteMode.Add)

        Assert.AreEqual("Alpha *", result.GetSection("Alpha *").GetKey("ID").Value)
        Assert.AreEqual("Beta *", result.GetSection("Beta *").GetKey("ID").Value)

    End Sub

    ''' <summary>
    ''' The token is recognized case-insensitively, consistent with all other Transmute matching
    ''' </summary>
    <TestMethod()> Public Sub GlobalAdd_EntryNameToken_CaseInsensitive()

        Dim result = RunTransmute(
            "[Alpha *]" & vbCrLf & "Detect=HKCU\Software\Alpha" & vbCrLf,
            "[*]" & vbCrLf & "ID=%entryname%" & vbCrLf,
            winapp2ool.TransmuteMode.Add)

        Assert.AreEqual("Alpha *", result.GetSection("Alpha *").GetKey("ID").Value)

    End Sub

    ''' <summary>
    ''' The token is only interpreted by [*] global sections: a named-section Add leaves
    ''' it as a literal value
    ''' </summary>
    <TestMethod()> Public Sub NamedAdd_EntryNameToken_StaysLiteral()

        Dim result = RunTransmute(
            "[Alpha *]" & vbCrLf & "Detect=HKCU\Software\Alpha" & vbCrLf,
            "[Alpha *]" & vbCrLf & "ID=%EntryName%" & vbCrLf,
            winapp2ool.TransmuteMode.Add)

        Assert.AreEqual("%EntryName%", result.GetSection("Alpha *").GetKey("ID").Value)

    End Sub

    ''' <summary>
    ''' Token-bearing and token-free keys coexist in the same [*] section: the token-free
    ''' key follows the original build-once path and applies identically everywhere
    ''' </summary>
    <TestMethod()> Public Sub GlobalAdd_MixedTokenAndPlainKeys_BothApplied()

        Dim result = RunTransmute(
            "[Alpha *]" & vbCrLf & "Detect=HKCU\Software\Alpha" & vbCrLf &
            "[Beta *]" & vbCrLf & "Detect=HKCU\Software\Beta" & vbCrLf,
            "[*]" & vbCrLf & "ID=%EntryName%" & vbCrLf & "Author=Winapp2.ini Project" & vbCrLf,
            winapp2ool.TransmuteMode.Add)

        For Each name In {"Alpha *", "Beta *"}
            Assert.AreEqual(name, result.GetSection(name).GetKey("ID").Value)
            Assert.AreEqual("Winapp2.ini Project", result.GetSection(name).GetKey("Author").Value)
        Next

    End Sub

    ''' <summary>
    ''' A [*Map:] match Value of exactly * matches any value of the named KeyType,
    ''' leaving keys of other KeyTypes untouched
    ''' </summary>
    <TestMethod()> Public Sub Map_WildcardMatch_MatchesAnyValueOfKeyType()

        Dim result = RunTransmute(
            "[Alpha *]" & vbCrLf & "LangSecRef=3021" & vbCrLf &
            "[Beta *]" & vbCrLf & "LangSecRef=9999" & vbCrLf &
            "[Gamma *]" & vbCrLf & "Section=Games" & vbCrLf & "Detect=HKCU\Software\Gamma" & vbCrLf,
            "[*Map: any LangSecRef]" & vbCrLf & "Match=LangSecRef=*" & vbCrLf & "Replace=Tags=ccapps" & vbCrLf,
            winapp2ool.TransmuteMode.Replace,
            replaceMode:=winapp2ool.ReplaceMode.ByKey)

        Assert.AreEqual("ccapps", result.GetSection("Alpha *").GetKey("Tags").Value)
        Assert.AreEqual("ccapps", result.GetSection("Beta *").GetKey("Tags").Value)

        ' Other KeyTypes are not matched by the LangSecRef wildcard
        Dim gamma = result.GetSection("Gamma *")
        Assert.IsFalse(gamma.HasKey("Tags"))
        Assert.AreEqual("Games", gamma.GetKey("Section").Value)
        Assert.AreEqual("HKCU\Software\Gamma", gamma.GetKey("Detect").Value)

    End Sub

    ''' <summary>
    ''' A specific rule listed before a wildcard fallback wins for its value while the
    ''' fallback catches everything else — the CC7 category-to-tag shape in miniature
    ''' </summary>
    <TestMethod()> Public Sub Map_SpecificRuleBeforeWildcardFallback_Wins()

        Dim result = RunTransmute(
            "[Alpha *]" & vbCrLf & "LangSecRef=3029" & vbCrLf &
            "[Beta *]" & vbCrLf & "LangSecRef=3001" & vbCrLf &
            "[Gamma *]" & vbCrLf & "Section=Custom Category" & vbCrLf,
            "[*Map: Chrome tag]" & vbCrLf & "Match=LangSecRef=3029" & vbCrLf & "Replace=Tags=Google,Chrome,Browser" & vbCrLf &
            "[*Map: tag fallback]" & vbCrLf & "Match1=LangSecRef=*" & vbCrLf & "Match2=Section=*" & vbCrLf & "Replace=Tags=ccapps" & vbCrLf,
            winapp2ool.TransmuteMode.Replace,
            replaceMode:=winapp2ool.ReplaceMode.ByKey)

        Assert.AreEqual("Google,Chrome,Browser", result.GetSection("Alpha *").GetKey("Tags").Value)
        Assert.AreEqual("ccapps", result.GetSection("Beta *").GetKey("Tags").Value)
        Assert.AreEqual("ccapps", result.GetSection("Gamma *").GetKey("Tags").Value)

    End Sub

    ''' <summary>
    ''' Wildcard rules obey the same mode gate as every other [*Map:] rule: outside
    ''' Replace ByKey they are skipped without touching the base file
    ''' </summary>
    <TestMethod()> Public Sub Map_Wildcard_OutsideReplaceByKey_IsSkipped()

        Dim result = RunTransmute(
            "[Alpha *]" & vbCrLf & "LangSecRef=3021" & vbCrLf,
            "[*Map: any LangSecRef]" & vbCrLf & "Match=LangSecRef=*" & vbCrLf & "Replace=Tags=ccapps" & vbCrLf,
            winapp2ool.TransmuteMode.Remove,
            removeKeyMode:=winapp2ool.RemoveKeyMode.ByName)

        Assert.AreEqual("3021", result.GetSection("Alpha *").GetKey("LangSecRef").Value)
        Assert.IsFalse(result.GetSection("Alpha *").HasKey("Tags"))

    End Sub

    ''' <summary>
    ''' A [*Name: scaffold] Add applies its payload to every section whose name ends with
    ''' " scaffold *", and the anchored suffix does not catch a longer scaffold sharing the
    ''' same prefix words (Web Browsing Session must not select Web Browsing Session Backups)
    ''' </summary>
    <TestMethod()> Public Sub NameFilter_AnchoredSuffix_SelectsScaffoldNotItsSuperstring()

        Dim result = RunTransmute(
            "[Google Chrome Web Browsing Session *]" & vbCrLf & "Section=Google Chrome Web Browser" & vbCrLf &
            "[Mozilla Firefox Web Browsing Session *]" & vbCrLf & "Section=Mozilla Firefox Web Browser" & vbCrLf &
            "[Mozilla Firefox Web Browsing Session Backups *]" & vbCrLf & "Section=Mozilla Firefox Web Browser" & vbCrLf,
            "[*Name: Web Browsing Session]" & vbCrLf & "Default=False" & vbCrLf,
            winapp2ool.TransmuteMode.Add)

        Assert.AreEqual("False", result.GetSection("Google Chrome Web Browsing Session *").GetKey("Default").Value)
        Assert.AreEqual("False", result.GetSection("Mozilla Firefox Web Browsing Session *").GetKey("Default").Value)
        ' The Backups superstring scaffold is NOT selected
        Assert.IsFalse(result.GetSection("Mozilla Firefox Web Browsing Session Backups *").HasKey("Default"))

    End Sub

    ''' <summary>
    ''' The OR'd Match= predicate set confirms entry kind: a section selected by name is only
    ''' modified when it also satisfies a predicate, so a same-named non-browser entry is skipped
    ''' while both a Web Browser and an Email Client (Thunderbird) are covered
    ''' </summary>
    <TestMethod()> Public Sub NameFilter_MatchPredicateSet_ConfirmsEntryKind()

        Dim result = RunTransmute(
            "[Google Chrome Bookmark Backups *]" & vbCrLf & "Section=Google Chrome Web Browser" & vbCrLf &
            "[Thunderbird Bookmark Backups *]" & vbCrLf & "Section=Thunderbird Email Client" & vbCrLf &
            "[Some App Bookmark Backups *]" & vbCrLf & "Section=Productivity" & vbCrLf,
            "[*Name: Bookmark Backups]" & vbCrLf & "Match1=Section=* Web Browser" & vbCrLf & "Match2=Section=* Email Client" & vbCrLf & "Default=False" & vbCrLf,
            winapp2ool.TransmuteMode.Add)

        Assert.AreEqual("False", result.GetSection("Google Chrome Bookmark Backups *").GetKey("Default").Value)
        Assert.AreEqual("False", result.GetSection("Thunderbird Bookmark Backups *").GetKey("Default").Value)
        ' The non-browser entry matched the name suffix but failed both predicates
        Assert.IsFalse(result.GetSection("Some App Bookmark Backups *").HasKey("Default"))

    End Sub

    ''' <summary>
    ''' A Match= value glob with a leading wildcard matches the messy real-world Section values
    ''' (leading dots, "Secure", "Origin") that a fixed enumeration would miss
    ''' </summary>
    <TestMethod()> Public Sub NameFilter_ValueGlob_MatchesMessySectionValues()

        Dim result = RunTransmute(
            "[.360 Secure Browser Web Browsing Cookies *]" & vbCrLf & "Section=.360 Secure Browser Web Browser" & vbCrLf &
            "[AVG Secure Browser Web Browsing Cookies *]" & vbCrLf & "Section=AVG Secure Web Browser" & vbCrLf &
            "[Brave Origin Web Browsing Cookies *]" & vbCrLf & "Section=Brave Origin Web Browser" & vbCrLf,
            "[*Name: Web Browsing Cookies]" & vbCrLf & "Match=Section=* Web Browser" & vbCrLf & "Default=False" & vbCrLf,
            winapp2ool.TransmuteMode.Add)

        For Each name In {".360 Secure Browser Web Browsing Cookies *", "AVG Secure Browser Web Browsing Cookies *", "Brave Origin Web Browsing Cookies *"}
            Assert.AreEqual("False", result.GetSection(name).GetKey("Default").Value)
        Next

    End Sub

    ''' <summary>
    ''' With no Match= predicates the name suffix alone decides selection
    ''' </summary>
    <TestMethod()> Public Sub NameFilter_NoPredicates_NameSuffixAloneSelects()

        Dim result = RunTransmute(
            "[Anything Pinned Tabs *]" & vbCrLf & "Section=Whatever" & vbCrLf &
            "[Unrelated Entry *]" & vbCrLf & "Section=Whatever" & vbCrLf,
            "[*Name: Pinned Tabs]" & vbCrLf & "Default=False" & vbCrLf,
            winapp2ool.TransmuteMode.Add)

        Assert.AreEqual("False", result.GetSection("Anything Pinned Tabs *").GetKey("Default").Value)
        Assert.IsFalse(result.GetSection("Unrelated Entry *").HasKey("Default"))

    End Sub

    ''' <summary>
    ''' In Add mode a payload key already present in a selected section is skipped, so the
    ''' operation never duplicates or overwrites an existing key
    ''' </summary>
    <TestMethod()> Public Sub NameFilter_Add_SkipsKeyAlreadyPresent()

        Dim result = RunTransmute(
            "[Chrome Download History *]" & vbCrLf & "Section=Chrome Web Browser" & vbCrLf & "Default=True" & vbCrLf &
            "[Edge Download History *]" & vbCrLf & "Section=Edge Web Browser" & vbCrLf,
            "[*Name: Download History]" & vbCrLf & "Match=Section=* Web Browser" & vbCrLf & "Default=False" & vbCrLf,
            winapp2ool.TransmuteMode.Add)

        Dim chrome = result.GetSection("Chrome Download History *")
        ' Pre-existing Default is left untouched and not duplicated
        Assert.AreEqual("True", chrome.GetKey("Default").Value)
        Assert.AreEqual(1, chrome.Keys.GetByType("Default").Count)
        ' The section that lacked Default receives it
        Assert.AreEqual("False", result.GetSection("Edge Download History *").GetKey("Default").Value)

    End Sub

    ''' <summary>
    ''' A [*Name:] payload can also drive Replace ByKey: it sets the value of the payload key
    ''' in selected sections that already have it, and does not create it where absent
    ''' </summary>
    <TestMethod()> Public Sub NameFilter_ReplaceByKey_ReplacesOnlyWherePresent()

        Dim result = RunTransmute(
            "[Chrome Extension Cookies *]" & vbCrLf & "Section=Chrome Web Browser" & vbCrLf & "Default=True" & vbCrLf &
            "[Edge Extension Cookies *]" & vbCrLf & "Section=Edge Web Browser" & vbCrLf,
            "[*Name: Extension Cookies]" & vbCrLf & "Match=Section=* Web Browser" & vbCrLf & "Default=False" & vbCrLf,
            winapp2ool.TransmuteMode.Replace,
            replaceMode:=winapp2ool.ReplaceMode.ByKey)

        Assert.AreEqual("False", result.GetSection("Chrome Extension Cookies *").GetKey("Default").Value)
        ' Replace never creates a key that was absent
        Assert.IsFalse(result.GetSection("Edge Extension Cookies *").HasKey("Default"))

    End Sub

    ''' <summary>
    ''' The %EntryName% token is expanded per selected section in a [*Name:] payload
    ''' </summary>
    <TestMethod()> Public Sub NameFilter_EntryNameToken_SubstitutesSectionName()

        Dim result = RunTransmute(
            "[Chrome Autofill Data *]" & vbCrLf & "Section=Chrome Web Browser" & vbCrLf &
            "[Edge Autofill Data *]" & vbCrLf & "Section=Edge Web Browser" & vbCrLf,
            "[*Name: Autofill Data]" & vbCrLf & "Match=Section=* Web Browser" & vbCrLf & "ID=%EntryName%" & vbCrLf,
            winapp2ool.TransmuteMode.Add)

        Assert.AreEqual("Chrome Autofill Data *", result.GetSection("Chrome Autofill Data *").GetKey("ID").Value)
        Assert.AreEqual("Edge Autofill Data *", result.GetSection("Edge Autofill Data *").GetKey("ID").Value)

    End Sub

    ''' <summary>
    ''' Numbered payload keys are refused by a [*Name:] Add for the same reason [*] refuses
    ''' them; an unnumbered payload key in the same rule still applies
    ''' </summary>
    <TestMethod()> Public Sub NameFilter_Add_RefusesNumberedPayloadKey()

        Dim result = RunTransmute(
            "[Chrome Web Browsing Cookies *]" & vbCrLf & "Section=Chrome Web Browser" & vbCrLf,
            "[*Name: Web Browsing Cookies]" & vbCrLf & "Match=Section=* Web Browser" & vbCrLf & "FileKey1=%AppData%\X|*.log" & vbCrLf & "Default=False" & vbCrLf,
            winapp2ool.TransmuteMode.Add)

        Dim section = result.GetSection("Chrome Web Browsing Cookies *")
        Assert.IsFalse(section.HasKey("FileKey1"))
        Assert.AreEqual("False", section.GetKey("Default").Value)

    End Sub

    ''' <summary>
    ''' Section-level modes are refused: a [*Name:] under Remove BySection or Replace BySection
    ''' leaves the base file untouched (the payload keys would be ignored)
    ''' </summary>
    <TestMethod()> Public Sub NameFilter_BySectionModes_AreRefused()

        Dim baseText = "[Chrome Pinned Tabs *]" & vbCrLf & "Section=Chrome Web Browser" & vbCrLf & "Default=True" & vbCrLf
        Dim source = "[*Name: Pinned Tabs]" & vbCrLf & "Default=False" & vbCrLf

        Dim removed = RunTransmute(baseText, source, winapp2ool.TransmuteMode.Remove, removeMode:=winapp2ool.RemoveMode.BySection)
        Assert.AreEqual(1, removed.Count)
        Assert.AreEqual("True", removed.GetSection("Chrome Pinned Tabs *").GetKey("Default").Value)

        Dim replaced = RunTransmute(baseText, source, winapp2ool.TransmuteMode.Replace, replaceMode:=winapp2ool.ReplaceMode.BySection)
        Assert.AreEqual("True", replaced.GetSection("Chrome Pinned Tabs *").GetKey("Default").Value)

    End Sub

    ''' <summary>
    ''' A [*Name:] rule that selects no section leaves the file untouched and does not throw
    ''' (it emits a stale-rule warning)
    ''' </summary>
    <TestMethod()> Public Sub NameFilter_ZeroSelected_LeavesFileUntouched()

        Dim result = RunTransmute(
            "[Chrome Web Browsing Cookies *]" & vbCrLf & "Section=Chrome Web Browser" & vbCrLf,
            "[*Name: No Such Scaffold]" & vbCrLf & "Match=Section=* Web Browser" & vbCrLf & "Default=False" & vbCrLf,
            winapp2ool.TransmuteMode.Add)

        Assert.IsFalse(result.GetSection("Chrome Web Browsing Cookies *").HasKey("Default"))

    End Sub

    ''' <summary>
    ''' With RecognizeGlobalSections disabled, [*Name:] falls through to normal section name
    ''' matching: in Add mode it is added to the base file as a literal section
    ''' </summary>
    <TestMethod()> Public Sub NameFilter_NoGlobal_TreatedAsOrdinarySection()

        winapp2ool.transmuteSettings.RecognizeGlobalSections = False

        Dim result = RunTransmute(
            "[Chrome Pinned Tabs *]" & vbCrLf & "Section=Chrome Web Browser" & vbCrLf,
            "[*Name: Pinned Tabs]" & vbCrLf & "Default=False" & vbCrLf,
            winapp2ool.TransmuteMode.Add)

        ' Added as a literal section instead of being interpreted
        Assert.IsTrue(result.Contains("*Name: Pinned Tabs"))
        Assert.IsFalse(result.GetSection("Chrome Pinned Tabs *").HasKey("Default"))

    End Sub

End Class
