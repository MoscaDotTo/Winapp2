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
''' Pure-function tests for <c> VariableExpander </c> and <c> VariableSet </c>.
''' Covers cartesian ordering, co-varying repeats, single-element degenerate,
''' undeclared-token domain handling, case-insensitivity, empty-list semantics,
''' commutativity with %% tokens, and the unreferenced-variable typo backstop.
''' </summary>
<TestClass()> Public Class VariableExpanderTests

    ''' <summary>
    ''' Helper: build a <c> VariableSet </c> from name/csv pairs in declaration order
    ''' </summary>
    Private Shared Function BuildVars(ParamArray decls As String()) As winapp2ool.VariableSet

        Dim vars As New winapp2ool.VariableSet
        For i = 0 To decls.Length - 1 Step 2
            vars.Add(decls(i), decls(i + 1))
        Next
        Return vars

    End Function

    ''' <summary>
    ''' A template with no <c> &lt;&gt; </c> tokens returns a single-element list
    ''' containing the template verbatim, with no diagnostics
    ''' </summary>
    <TestMethod()> Public Sub Expand_NoTokens_ReturnsTemplateVerbatim()

        Dim vars = BuildVars()
        Dim result = winapp2ool.VariableExpander.Expand(
            "HKCU\Software\Foo|bar",
            vars,
            winapp2ool.ExpansionDomain.Registry,
            "[Entry]")

        Assert.AreEqual(1, result.Values.Count)
        Assert.AreEqual("HKCU\Software\Foo|bar", result.Values(0))
        Assert.AreEqual(0, result.Diagnostics.Count)

    End Sub

    ''' <summary>
    ''' A single declared variable with multiple values produces one substituted
    ''' string per value, in declaration order
    ''' </summary>
    <TestMethod()> Public Sub Expand_SingleDeclaredToken_FansOutOverValues()

        Dim vars = BuildVars("version", "9.0,10.0,11.0")
        Dim result = winapp2ool.VariableExpander.Expand(
            "HKCU\Software\Microsoft\Office\<version>\Outlook",
            vars,
            winapp2ool.ExpansionDomain.Registry,
            "[Entry]")

        Assert.AreEqual(3, result.Values.Count)
        Assert.AreEqual("HKCU\Software\Microsoft\Office\9.0\Outlook", result.Values(0))
        Assert.AreEqual("HKCU\Software\Microsoft\Office\10.0\Outlook", result.Values(1))
        Assert.AreEqual("HKCU\Software\Microsoft\Office\11.0\Outlook", result.Values(2))

    End Sub

    ''' <summary>
    ''' Two distinct declared variables fan out as a cartesian product, with the
    ''' leftmost-occurring variable varying slowest (outermost loop)
    ''' </summary>
    <TestMethod()> Public Sub Expand_TwoDeclaredTokens_CartesianLeftmostOutermost()

        Dim vars = BuildVars("version", "9.0,10.0", "mru", "1,2,3")
        Dim result = winapp2ool.VariableExpander.Expand(
            "HKCU\Software\Microsoft\Office\<version>\Outlook\Office Finder|MRU <mru>",
            vars,
            winapp2ool.ExpansionDomain.Registry,
            "[Entry]")

        Assert.AreEqual(6, result.Values.Count)
        ' Outer (version) slowest — every version's mru block contiguous
        Assert.AreEqual("HKCU\Software\Microsoft\Office\9.0\Outlook\Office Finder|MRU 1", result.Values(0))
        Assert.AreEqual("HKCU\Software\Microsoft\Office\9.0\Outlook\Office Finder|MRU 2", result.Values(1))
        Assert.AreEqual("HKCU\Software\Microsoft\Office\9.0\Outlook\Office Finder|MRU 3", result.Values(2))
        Assert.AreEqual("HKCU\Software\Microsoft\Office\10.0\Outlook\Office Finder|MRU 1", result.Values(3))
        Assert.AreEqual("HKCU\Software\Microsoft\Office\10.0\Outlook\Office Finder|MRU 2", result.Values(4))
        Assert.AreEqual("HKCU\Software\Microsoft\Office\10.0\Outlook\Office Finder|MRU 3", result.Values(5))

    End Sub

    ''' <summary>
    ''' A variable referenced multiple times in the same template co-varies — the
    ''' same chosen value is substituted at every occurrence on each iteration,
    ''' producing one result per value (not a self-cartesian)
    ''' </summary>
    <TestMethod()> Public Sub Expand_RepeatedTokenInSameTemplate_CoVaries()

        Dim vars = BuildVars("x", "a,b,c")
        Dim result = winapp2ool.VariableExpander.Expand(
            "<x>\middle\<x>",
            vars,
            winapp2ool.ExpansionDomain.Registry,
            "[Entry]")

        Assert.AreEqual(3, result.Values.Count)
        Assert.AreEqual("a\middle\a", result.Values(0))
        Assert.AreEqual("b\middle\b", result.Values(1))
        Assert.AreEqual("c\middle\c", result.Values(2))

    End Sub

    ''' <summary>
    ''' A declared variable with one value degenerates to plain substitution — no
    ''' fan-out, just a single result with the value substituted
    ''' </summary>
    <TestMethod()> Public Sub Expand_SingleElementList_NoFanOut()

        Dim vars = BuildVars("only", "solo")
        Dim result = winapp2ool.VariableExpander.Expand(
            "prefix-<only>-suffix",
            vars,
            winapp2ool.ExpansionDomain.Filesystem,
            "[Entry]")

        Assert.AreEqual(1, result.Values.Count)
        Assert.AreEqual("prefix-solo-suffix", result.Values(0))

    End Sub

    ''' <summary>
    ''' An undeclared <c> &lt;name&gt; </c> in a filesystem-domain key drops the key
    ''' entirely (empty <c> Values </c>) and emits a Warning diagnostic
    ''' </summary>
    <TestMethod()> Public Sub Expand_UndeclaredTokenInFilesystemKey_DropsKey()

        Dim vars = BuildVars()
        Dim result = winapp2ool.VariableExpander.Expand(
            "%LocalAppData%\App\<typo>\file",
            vars,
            winapp2ool.ExpansionDomain.Filesystem,
            "[Entry]")

        Assert.AreEqual(0, result.Values.Count)
        Assert.AreEqual(1, result.Diagnostics.Count)
        Assert.AreEqual(winapp2ool.DiagnosticSeverity.Warning, result.Diagnostics(0).Severity)
        StringAssert.Contains(result.Diagnostics(0).Message, "<typo>")
        StringAssert.Contains(result.Diagnostics(0).Message, "dropping")

    End Sub

    ''' <summary>
    ''' An undeclared <c> &lt;name&gt; </c> in a registry-domain key passes through
    ''' verbatim with an Advisory diagnostic — registry paths can legitimately
    ''' contain literal angle brackets
    ''' </summary>
    <TestMethod()> Public Sub Expand_UndeclaredTokenInRegistryKey_KeepsLiteral()

        Dim vars = BuildVars()
        Dim result = winapp2ool.VariableExpander.Expand(
            "HKCU\Software\<weirdname>\Value",
            vars,
            winapp2ool.ExpansionDomain.Registry,
            "[Entry]")

        Assert.AreEqual(1, result.Values.Count)
        Assert.AreEqual("HKCU\Software\<weirdname>\Value", result.Values(0))
        Assert.AreEqual(1, result.Diagnostics.Count)
        Assert.AreEqual(winapp2ool.DiagnosticSeverity.Advisory, result.Diagnostics(0).Severity)
        StringAssert.Contains(result.Diagnostics(0).Message, "<weirdname>")

    End Sub

    ''' <summary>
    ''' Variable name lookup is case-insensitive: <c> &lt;VERSION&gt; </c> binds to
    ''' a <c> version= </c> declaration and vice versa
    ''' </summary>
    <TestMethod()> Public Sub Expand_CaseInsensitiveLookup_BindsRegardlessOfCase()

        Dim vars = BuildVars("Version", "9.0,10.0")
        Dim result = winapp2ool.VariableExpander.Expand(
            "x-<VERSION>-<version>-<vErSiOn>",
            vars,
            winapp2ool.ExpansionDomain.Registry,
            "[Entry]")

        Assert.AreEqual(2, result.Values.Count)
        Assert.AreEqual("x-9.0-9.0-9.0", result.Values(0))
        Assert.AreEqual("x-10.0-10.0-10.0", result.Values(1))

    End Sub

    ''' <summary>
    ''' <c> %name% </c> environment-variable tokens are not expansion targets; they
    ''' pass through untouched, demonstrating that the %% and &lt;&gt; token domains
    ''' are closed (the two-phase composition relies on this)
    ''' </summary>
    <TestMethod()> Public Sub Expand_PercentTokensAreUntouched()

        Dim vars = BuildVars("v", "9,10")
        Dim result = winapp2ool.VariableExpander.Expand(
            "%LocalAppData%\Microsoft\Office\<v>\%UserName%",
            vars,
            winapp2ool.ExpansionDomain.Filesystem,
            "[Entry]")

        Assert.AreEqual(2, result.Values.Count)
        Assert.AreEqual("%LocalAppData%\Microsoft\Office\9\%UserName%", result.Values(0))
        Assert.AreEqual("%LocalAppData%\Microsoft\Office\10\%UserName%", result.Values(1))
        Assert.AreEqual(0, result.Diagnostics.Count)

    End Sub

    ''' <summary>
    ''' A declared variable with an empty value list, referenced by a template,
    ''' drops the key and emits a Warning. This is the empty-list-semantics decision
    ''' from the plan (lean: warn + drop, parallels Scaffolds=present-but-empty)
    ''' </summary>
    <TestMethod()> Public Sub Expand_EmptyValueList_DropsKey()

        Dim vars = BuildVars("v", "")
        Dim result = winapp2ool.VariableExpander.Expand(
            "key-<v>",
            vars,
            winapp2ool.ExpansionDomain.Registry,
            "[Entry]")

        Assert.AreEqual(0, result.Values.Count)
        Assert.AreEqual(1, result.Diagnostics.Count)
        Assert.AreEqual(winapp2ool.DiagnosticSeverity.Warning, result.Diagnostics(0).Severity)
        StringAssert.Contains(result.Diagnostics(0).Message, "<v>")
        StringAssert.Contains(result.Diagnostics(0).Message, "empty value list")

    End Sub

    ''' <summary>
    ''' Mixed declared + undeclared in a registry-domain template: declared tokens
    ''' expand, undeclared tokens stay literal, advisory diagnostic emitted
    ''' </summary>
    <TestMethod()> Public Sub Expand_MixedDeclaredAndUndeclared_Registry_ExpandsDeclaredKeepsLiteral()

        Dim vars = BuildVars("v", "9,10")
        Dim result = winapp2ool.VariableExpander.Expand(
            "HKCU\<v>\<typo>\stuff",
            vars,
            winapp2ool.ExpansionDomain.Registry,
            "[Entry]")

        Assert.AreEqual(2, result.Values.Count)
        Assert.AreEqual("HKCU\9\<typo>\stuff", result.Values(0))
        Assert.AreEqual("HKCU\10\<typo>\stuff", result.Values(1))
        Assert.AreEqual(1, result.Diagnostics.Count)
        Assert.AreEqual(winapp2ool.DiagnosticSeverity.Advisory, result.Diagnostics(0).Severity)

    End Sub

    ''' <summary>
    ''' Mixed declared + undeclared in a filesystem-domain template: the whole key
    ''' is dropped because the undeclared token cannot be safely passed through —
    ''' literal <c> &lt; </c>/<c> &gt; </c> is impossible in a Win32 path
    ''' </summary>
    <TestMethod()> Public Sub Expand_MixedDeclaredAndUndeclared_Filesystem_DropsKey()

        Dim vars = BuildVars("v", "9,10")
        Dim result = winapp2ool.VariableExpander.Expand(
            "%LocalAppData%\<v>\<typo>\file",
            vars,
            winapp2ool.ExpansionDomain.Filesystem,
            "[Entry]")

        Assert.AreEqual(0, result.Values.Count)
        Assert.AreEqual(1, result.Diagnostics.Count)
        Assert.AreEqual(winapp2ool.DiagnosticSeverity.Warning, result.Diagnostics(0).Severity)

    End Sub

    ''' <summary>
    ''' First-occurrence ordering uses position across the whole template, including
    ''' across the <c> | </c> boundary in winapp2 key values — proving the engine
    ''' doesn't special-case the pipe
    ''' </summary>
    <TestMethod()> Public Sub Expand_FirstOccurrenceSpansThePipe()

        Dim vars = BuildVars("a", "1,2", "b", "x,y")
        ' <b> appears first (left of pipe), <a> appears second (right of pipe)
        Dim result = winapp2ool.VariableExpander.Expand(
            "path\<b>|key <a>",
            vars,
            winapp2ool.ExpansionDomain.Registry,
            "[Entry]")

        Assert.AreEqual(4, result.Values.Count)
        ' Outer = b (first occurrence), inner = a
        Assert.AreEqual("path\x|key 1", result.Values(0))
        Assert.AreEqual("path\x|key 2", result.Values(1))
        Assert.AreEqual("path\y|key 1", result.Values(2))
        Assert.AreEqual("path\y|key 2", result.Values(3))

    End Sub

    ''' <summary>
    ''' Driving migration case from the plan: the Outlook RegKey template fans out
    ''' over both Office <c> &lt;version&gt; </c> (7 values, no Office 13.0) and MRU
    ''' <c> &lt;mrunums&gt; </c> (3 values), producing 21 keys with version varying
    ''' slowest
    ''' </summary>
    <TestMethod()> Public Sub Expand_OutlookRegKeyTemplate_FansOutVersionAndMRU()

        Dim vars = BuildVars(
            "version", "9.0,10.0,11.0,12.0,14.0,15.0,16.0",
            "MRUNums", "1,2,3")

        Dim result = winapp2ool.VariableExpander.Expand(
            "HKCU\Software\Microsoft\Office\<version>\Outlook\Office Finder|MRU <MRUNums>",
            vars,
            winapp2ool.ExpansionDomain.Registry,
            "[Microsoft Outlook *].RegKey")

        Assert.AreEqual(21, result.Values.Count)
        Assert.AreEqual("HKCU\Software\Microsoft\Office\9.0\Outlook\Office Finder|MRU 1", result.Values(0))
        Assert.AreEqual("HKCU\Software\Microsoft\Office\9.0\Outlook\Office Finder|MRU 3", result.Values(2))
        Assert.AreEqual("HKCU\Software\Microsoft\Office\10.0\Outlook\Office Finder|MRU 1", result.Values(3))
        Assert.AreEqual("HKCU\Software\Microsoft\Office\16.0\Outlook\Office Finder|MRU 3", result.Values(20))
        Assert.AreEqual(0, result.Diagnostics.Count)

    End Sub

    ''' <summary>
    ''' Calling <c> Expand </c> records every observed declared token as referenced,
    ''' so <c> UnreferencedNames </c> only returns variables never seen
    ''' </summary>
    <TestMethod()> Public Sub Expand_MarksReferencedVariablesOnVariableSet()

        Dim vars = BuildVars("used", "1,2", "unused", "x,y", "alsoused", "a,b")
        winapp2ool.VariableExpander.Expand("k-<used>-<alsoused>", vars,
                                           winapp2ool.ExpansionDomain.Registry, "[Entry]")

        Dim unref = vars.UnreferencedNames()
        Assert.AreEqual(1, unref.Count)
        Assert.AreEqual("unused", unref(0))

    End Sub

    ''' <summary>
    ''' <c> MarkReferenced </c> on an undeclared name is silently ignored; it does
    ''' not add the name to the declared set or affect unreferenced enumeration
    ''' </summary>
    <TestMethod()> Public Sub VariableSet_MarkReferencedOnUndeclared_NoEffect()

        Dim vars = BuildVars("declared", "1,2")
        vars.MarkReferenced("notdeclared")

        Assert.IsFalse(vars.IsDeclared("notdeclared"))
        Dim unref = vars.UnreferencedNames()
        Assert.AreEqual(1, unref.Count)
        Assert.AreEqual("declared", unref(0))

    End Sub

    ''' <summary>
    ''' Empty input templates pass through verbatim
    ''' </summary>
    <TestMethod()> Public Sub Expand_EmptyTemplate_ReturnsEmptyString()

        Dim vars = BuildVars("v", "9,10")
        Dim result = winapp2ool.VariableExpander.Expand("", vars,
                                                         winapp2ool.ExpansionDomain.Registry, "[Entry]")

        Assert.AreEqual(1, result.Values.Count)
        Assert.AreEqual("", result.Values(0))

    End Sub

    ''' <summary>
    ''' Driving Auslogics-style case: a variable whose values themselves contain
    ''' <c> &lt;&gt; </c> tokens (RegRoots = two templates each referencing a
    ''' different version list) resolves before any user-template expansion so
    ''' that <c> &lt;RegRoots&gt; </c> later behaves like a flat 7-element list
    ''' </summary>
    <TestMethod()> Public Sub ResolveAll_NestedVariable_FansOutEachValueOverItsDeps()

        Dim vars As New winapp2ool.VariableSet
        vars.Add("HKCUVersions", "2.x,3.x,4.x,5.x")
        vars.Add("HKLMVersions", "3.x,4.x,5.x")
        vars.Add("RegRoots", "HKCU\App\<HKCUVersions>,HKLM\Wow6432\App\<HKLMVersions>")

        Dim diags = winapp2ool.VariableExpander.ResolveAll(vars)

        Assert.AreEqual(0, diags.Count)
        Dim regRoots = vars.Values("RegRoots")
        Assert.AreEqual(7, regRoots.Count)
        Assert.AreEqual("HKCU\App\2.x", regRoots(0))
        Assert.AreEqual("HKCU\App\3.x", regRoots(1))
        Assert.AreEqual("HKCU\App\4.x", regRoots(2))
        Assert.AreEqual("HKCU\App\5.x", regRoots(3))
        Assert.AreEqual("HKLM\Wow6432\App\3.x", regRoots(4))
        Assert.AreEqual("HKLM\Wow6432\App\4.x", regRoots(5))
        Assert.AreEqual("HKLM\Wow6432\App\5.x", regRoots(6))

    End Sub

    ''' <summary>
    ''' End-to-end Auslogics-style template: after <c> ResolveAll </c>, the
    ''' RegKeyBase template <c> &lt;RegRoots&gt;\Settings|&lt;SettingsKeys&gt; </c>
    ''' fans 7×3 to 21 keys, with each RegRoots value carrying its own resolved
    ''' version segment
    ''' </summary>
    <TestMethod()> Public Sub ResolveAll_ThenExpand_AuslogicsRegKeyBase()

        Dim vars As New winapp2ool.VariableSet
        vars.Add("HKCUVersions", "2.x,3.x")
        vars.Add("HKLMVersions", "3.x,4.x,5.x")
        vars.Add("RegRoots", "HKCU\Auslogics\<HKCUVersions>,HKLM\Wow6432Node\Auslogics\<HKLMVersions>")
        vars.Add("SettingsKeys", "ErrorsFound,ErrorsRemoved,LastScan")

        winapp2ool.VariableExpander.ResolveAll(vars)

        Dim result = winapp2ool.VariableExpander.Expand(
            "<RegRoots>\Settings|<SettingsKeys>",
            vars,
            winapp2ool.ExpansionDomain.Registry,
            "[Auslogics].RegKeyBase")

        Assert.AreEqual(15, result.Values.Count)
        ' Outer = RegRoots (5 values: 2 HKCU + 3 HKLM), inner = SettingsKeys (3)
        Assert.AreEqual("HKCU\Auslogics\2.x\Settings|ErrorsFound", result.Values(0))
        Assert.AreEqual("HKCU\Auslogics\2.x\Settings|LastScan", result.Values(2))
        Assert.AreEqual("HKCU\Auslogics\3.x\Settings|ErrorsFound", result.Values(3))
        Assert.AreEqual("HKLM\Wow6432Node\Auslogics\5.x\Settings|LastScan", result.Values(14))

    End Sub

    ''' <summary>
    ''' Multi-level nesting: A references B, B references C. After ResolveAll,
    ''' A's values reflect both layers
    ''' </summary>
    <TestMethod()> Public Sub ResolveAll_TransitiveDependency_ResolvesAllLayers()

        Dim vars As New winapp2ool.VariableSet
        vars.Add("C", "x,y")
        vars.Add("B", "b1<C>,b2<C>")
        vars.Add("A", "a-<B>")

        winapp2ool.VariableExpander.ResolveAll(vars)

        Dim aValues = vars.Values("A")
        Assert.AreEqual(4, aValues.Count)
        ' B resolves to [b1x, b1y, b2x, b2y]; A then becomes a-<each>
        Assert.AreEqual("a-b1x", aValues(0))
        Assert.AreEqual("a-b1y", aValues(1))
        Assert.AreEqual("a-b2x", aValues(2))
        Assert.AreEqual("a-b2y", aValues(3))

    End Sub

    ''' <summary>
    ''' Resolution order is independent of declaration order: declaring the
    ''' dependent (B) before its dependency (C) still resolves correctly via
    ''' topological order
    ''' </summary>
    <TestMethod()> Public Sub ResolveAll_DeclarationOrderIndependent()

        Dim vars As New winapp2ool.VariableSet
        vars.Add("B", "<C>-foo")           ' declared before C
        vars.Add("C", "1,2")

        winapp2ool.VariableExpander.ResolveAll(vars)

        Dim bValues = vars.Values("B")
        Assert.AreEqual(2, bValues.Count)
        Assert.AreEqual("1-foo", bValues(0))
        Assert.AreEqual("2-foo", bValues(1))

    End Sub

    ''' <summary>
    ''' A two-variable cycle (A ↔ B) is detected: both variables remain at their
    ''' authored templated values and a Warning diagnostic is emitted for each
    ''' </summary>
    <TestMethod()> Public Sub ResolveAll_TwoVariableCycle_LeavesUnresolvedAndWarns()

        Dim vars As New winapp2ool.VariableSet
        vars.Add("A", "<B>")
        vars.Add("B", "<A>")

        Dim diags = winapp2ool.VariableExpander.ResolveAll(vars)

        ' Both vars unresolved; their values still contain the tokens
        Assert.AreEqual("<B>", vars.Values("A")(0))
        Assert.AreEqual("<A>", vars.Values("B")(0))
        ' One Warning per cyclic var
        Assert.AreEqual(2, diags.Count)
        Assert.IsTrue(diags.All(Function(d) d.Severity = winapp2ool.DiagnosticSeverity.Warning))

    End Sub

    ''' <summary>
    ''' A self-cycle (A references &lt;A&gt;) is detected just like a multi-var cycle
    ''' </summary>
    <TestMethod()> Public Sub ResolveAll_SelfCycle_Detected()

        Dim vars As New winapp2ool.VariableSet
        vars.Add("A", "loop-<A>-end")

        Dim diags = winapp2ool.VariableExpander.ResolveAll(vars)

        Assert.AreEqual("loop-<A>-end", vars.Values("A")(0))
        Assert.AreEqual(1, diags.Count)
        Assert.AreEqual(winapp2ool.DiagnosticSeverity.Warning, diags(0).Severity)
        StringAssert.Contains(diags(0).Message, "<A>")

    End Sub

    ''' <summary>
    ''' An undeclared token inside a variable's value resolves under registry
    ''' rules (advisory + literal preserved), letting the user-template's later
    ''' domain check catch it if the literal would be invalid downstream
    ''' </summary>
    <TestMethod()> Public Sub ResolveAll_UndeclaredInValue_EmitsAdvisoryKeepsLiteral()

        Dim vars As New winapp2ool.VariableSet
        vars.Add("V", "foo-<NoSuch>-bar")

        Dim diags = winapp2ool.VariableExpander.ResolveAll(vars)

        Assert.AreEqual(1, vars.Values("V").Count)
        Assert.AreEqual("foo-<NoSuch>-bar", vars.Values("V")(0))
        Assert.AreEqual(1, diags.Count)
        Assert.AreEqual(winapp2ool.DiagnosticSeverity.Advisory, diags(0).Severity)

    End Sub

    ''' <summary>
    ''' Resolution marks a variable as referenced when its <c> &lt;name&gt; </c> token
    ''' appears inside another variable's value list, so a variable consumed only via
    ''' a nested chain does not false-positive the unreferenced-variable typo backstop.
    ''' A variable reached by neither a user template nor a nested chain (here,
    ''' <c> Outer </c>) remains unreferenced and is reported.
    ''' </summary>
    <TestMethod()> Public Sub ResolveAll_MarksReferencedViaNestedChains()

        Dim vars As New winapp2ool.VariableSet
        vars.Add("Inner", "x,y")
        vars.Add("Outer", "<Inner>-z")

        winapp2ool.VariableExpander.ResolveAll(vars)

        ' Inner was consumed by Outer's value list during resolution → marked.
        ' Outer was not consumed by any template → still unreferenced.
        Dim unref = vars.UnreferencedNames()
        Assert.AreEqual(1, unref.Count)
        Assert.IsTrue(unref.Contains("Outer"))

    End Sub

    ''' <summary>
    ''' An inline ad-hoc list <c> &lt;a,b,c&gt; </c> needs no declaration: the
    ''' comma-bearing token body is split into an anonymous axis and fanned out
    ''' exactly like a declared variable, in body order
    ''' </summary>
    <TestMethod()> Public Sub Expand_InlineList_FansOutWithoutDeclaration()

        Dim vars = BuildVars()
        Dim result = winapp2ool.VariableExpander.Expand(
            "%WinDir%\<System32,SysWOW64>\Macromed\Flash|*.log",
            vars,
            winapp2ool.ExpansionDomain.Filesystem,
            "[Entry]")

        Assert.AreEqual(2, result.Values.Count)
        Assert.AreEqual("%WinDir%\System32\Macromed\Flash|*.log", result.Values(0))
        Assert.AreEqual("%WinDir%\SysWOW64\Macromed\Flash|*.log", result.Values(1))
        Assert.AreEqual(0, result.Diagnostics.Count)

    End Sub

    ''' <summary>
    ''' Inline list values are trimmed; surrounding whitespace in the body does not
    ''' leak into the substituted output
    ''' </summary>
    <TestMethod()> Public Sub Expand_InlineList_TrimsValues()

        Dim vars = BuildVars()
        Dim result = winapp2ool.VariableExpander.Expand(
            "App\<Cache, Logs , Temp>",
            vars,
            winapp2ool.ExpansionDomain.Filesystem,
            "[Entry]")

        Assert.AreEqual(3, result.Values.Count)
        Assert.AreEqual("App\Cache", result.Values(0))
        Assert.AreEqual("App\Logs", result.Values(1))
        Assert.AreEqual("App\Temp", result.Values(2))

    End Sub

    ''' <summary>
    ''' Two <em>different</em> inline lists in one template cross-product like any two
    ''' axes, with the leftmost varying slowest — no advisory, this is legitimate use
    ''' </summary>
    <TestMethod()> Public Sub Expand_TwoDistinctInlineLists_CrossProduct()

        Dim vars = BuildVars()
        Dim result = winapp2ool.VariableExpander.Expand(
            "%LocalAppData%\App\<v1,v2>\<Cache,Logs>|*",
            vars,
            winapp2ool.ExpansionDomain.Filesystem,
            "[Entry]")

        Assert.AreEqual(4, result.Values.Count)
        Assert.AreEqual("%LocalAppData%\App\v1\Cache|*", result.Values(0))
        Assert.AreEqual("%LocalAppData%\App\v1\Logs|*", result.Values(1))
        Assert.AreEqual("%LocalAppData%\App\v2\Cache|*", result.Values(2))
        Assert.AreEqual("%LocalAppData%\App\v2\Logs|*", result.Values(3))
        Assert.AreEqual(0, result.Diagnostics.Count)

    End Sub

    ''' <summary>
    ''' An <em>identical</em> inline list repeated in one template collapses to a
    ''' single co-varying axis (not a self-cartesian) and emits exactly one Advisory,
    ''' because an anonymous list cannot express same-axis intent
    ''' </summary>
    <TestMethod()> Public Sub Expand_RepeatedIdenticalInlineList_CoVariesAndAdvises()

        Dim vars = BuildVars()
        Dim result = winapp2ool.VariableExpander.Expand(
            "%WinDir%\<System32,SysWOW64>\Macromed\backup-<System32,SysWOW64>|*.log",
            vars,
            winapp2ool.ExpansionDomain.Filesystem,
            "[Entry]")

        ' One axis, two values — NOT 2x2. Positions move together.
        Assert.AreEqual(2, result.Values.Count)
        Assert.AreEqual("%WinDir%\System32\Macromed\backup-System32|*.log", result.Values(0))
        Assert.AreEqual("%WinDir%\SysWOW64\Macromed\backup-SysWOW64|*.log", result.Values(1))
        Assert.AreEqual(1, result.Diagnostics.Count)
        Assert.AreEqual(winapp2ool.DiagnosticSeverity.Advisory, result.Diagnostics(0).Severity)
        StringAssert.Contains(result.Diagnostics(0).Message, "co-vary")

    End Sub

    ''' <summary>
    ''' A declared variable and an inline list compose in one template: both become
    ''' axes, declared-first varying slowest, with no diagnostics
    ''' </summary>
    <TestMethod()> Public Sub Expand_DeclaredAndInlineList_Compose()

        Dim vars = BuildVars("ver", "1.0,2.0")
        Dim result = winapp2ool.VariableExpander.Expand(
            "App\<ver>\<Cache,Logs>",
            vars,
            winapp2ool.ExpansionDomain.Filesystem,
            "[Entry]")

        Assert.AreEqual(4, result.Values.Count)
        Assert.AreEqual("App\1.0\Cache", result.Values(0))
        Assert.AreEqual("App\1.0\Logs", result.Values(1))
        Assert.AreEqual("App\2.0\Cache", result.Values(2))
        Assert.AreEqual("App\2.0\Logs", result.Values(3))
        Assert.AreEqual(0, result.Diagnostics.Count)

    End Sub

    ''' <summary>
    ''' A degenerate inline list whose body trims to nothing (<c> &lt; , &gt; </c>)
    ''' has no values, so the cartesian product is empty and the key is dropped with
    ''' a Warning — same path as a declared empty list
    ''' </summary>
    <TestMethod()> Public Sub Expand_DegenerateInlineList_DropsKey()

        Dim vars = BuildVars()
        Dim result = winapp2ool.VariableExpander.Expand(
            "App\< , >\file",
            vars,
            winapp2ool.ExpansionDomain.Filesystem,
            "[Entry]")

        Assert.AreEqual(0, result.Values.Count)
        Assert.AreEqual(1, result.Diagnostics.Count)
        Assert.AreEqual(winapp2ool.DiagnosticSeverity.Warning, result.Diagnostics(0).Severity)

    End Sub

    ''' <summary>
    ''' An inline list inside a DECLARATION value survives comma-splitting: the
    ''' declaration is one value carrying an inline <c> &lt;a,b,c&gt; </c> list, not
    ''' three fragments split at the list's commas. This is the [Ability Office Write *]
    ''' Root scenario — <c> Root=...\&lt;6.0,7.0,11.0&gt;\... </c> must resolve to three
    ''' clean roots so root-detection inference does not false-skip on a stray bracket.
    ''' </summary>
    <TestMethod()> Public Sub Add_InlineListInDeclaration_NotShreddedByCommaSplit()

        Dim vars As New winapp2ool.VariableSet
        vars.Add("Root", "HKCU\Software\Ability <6.0,7.0,11.0>\Ability Write")

        ' One declared value, inline list intact (NOT 3 fragments at the inner commas)
        Dim raw = vars.Values("Root")
        Assert.AreEqual(1, raw.Count)
        Assert.AreEqual("HKCU\Software\Ability <6.0,7.0,11.0>\Ability Write", raw(0))

        ' After resolution the inline list fans the single value out to three clean roots
        winapp2ool.VariableExpander.ResolveAll(vars)
        Dim resolved = vars.Values("Root")
        Assert.AreEqual(3, resolved.Count)
        Assert.AreEqual("HKCU\Software\Ability 6.0\Ability Write", resolved(0))
        Assert.AreEqual("HKCU\Software\Ability 7.0\Ability Write", resolved(1))
        Assert.AreEqual("HKCU\Software\Ability 11.0\Ability Write", resolved(2))
        ' None retains a < bracket, so inferRootDetection classifies rather than skips
        Assert.IsFalse(resolved.Any(Function(v) v.Contains("<")))

    End Sub

    ''' <summary>
    ''' Top-level commas still separate declaration values when an inline list is also
    ''' present: <c> x,&lt;a,b&gt;,y </c> is three values, the middle carrying the list
    ''' </summary>
    <TestMethod()> Public Sub Add_MixedTopLevelAndInlineCommas_SplitsOnlyTopLevel()

        Dim vars As New winapp2ool.VariableSet
        vars.Add("V", "x,<a,b>,y")

        Dim raw = vars.Values("V")
        Assert.AreEqual(3, raw.Count)
        Assert.AreEqual("x", raw(0))
        Assert.AreEqual("<a,b>", raw(1))
        Assert.AreEqual("y", raw(2))

    End Sub

    ''' <summary>
    ''' A plain comma-separated declaration with no brackets splits on every comma as
    ''' before — the bracket-aware split is backward-compatible for the common case
    ''' </summary>
    <TestMethod()> Public Sub Add_PlainCsvDeclaration_UnchangedSplit()

        Dim vars As New winapp2ool.VariableSet
        vars.Add("Versions", "6.0,7.0,11.0")

        Dim raw = vars.Values("Versions")
        Assert.AreEqual(3, raw.Count)
        Assert.AreEqual("6.0", raw(0))
        Assert.AreEqual("7.0", raw(1))
        Assert.AreEqual("11.0", raw(2))

    End Sub

    ''' <summary>
    ''' Stray <c> &lt; </c> or <c> &gt; </c> characters without a complete
    ''' <c> &lt;name&gt; </c> form are left literal — the tokenizer requires the
    ''' full pattern to match
    ''' </summary>
    <TestMethod()> Public Sub Expand_StrayAngleBrackets_AreLiteral()

        Dim vars = BuildVars("v", "9,10")
        Dim result = winapp2ool.VariableExpander.Expand(
            "x < y > z",
            vars,
            winapp2ool.ExpansionDomain.Registry,
            "[Entry]")

        Assert.AreEqual(1, result.Values.Count)
        ' Note: "< y >" actually matches <[^<>]+> because " y " has no < or > — so this
        ' DOES match a token "< y >" with name " y ". Verify it's treated as undeclared.
        Assert.AreEqual("x < y > z", result.Values(0))
        Assert.AreEqual(1, result.Diagnostics.Count)
        Assert.AreEqual(winapp2ool.DiagnosticSeverity.Advisory, result.Diagnostics(0).Severity)

    End Sub

End Class
