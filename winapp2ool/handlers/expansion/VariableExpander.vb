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
''' The domain a key value lives in, used to decide how undeclared <c> &lt;token&gt; </c>
''' references are resolved. Filesystem keys (FileKey, DetectFile, etc.) cannot contain
''' literal <c> &lt; </c> or <c> &gt; </c> because those are reserved Win32 path
''' characters; registry keys can. The engine therefore drops filesystem keys that
''' carry an undeclared token (hard error) and lets registry keys keep the literal
''' bracket (advisory only).
''' </summary>
Public Enum ExpansionDomain

    ''' <summary>
    ''' The key value will be written into a file system location (FileKey, FileKeyBase,
    ''' DetectFile, ExcludeKey with File/Path flag, DetectOS, anything unclassified).
    ''' Undeclared <c> &lt;token&gt; </c> references drop the key.
    ''' </summary>
    Filesystem

    ''' <summary>
    ''' The key value will be written into a registry location (RegKey, RegKeyBase,
    ''' Detect, ExcludeKey with Reg flag). Undeclared <c> &lt;token&gt; </c> references
    ''' pass through verbatim with an advisory diagnostic.
    ''' </summary>
    Registry

End Enum

''' <summary>
''' Severity of an <see cref="ExpansionDiagnostic"/>. The engine raises no output
''' itself; the caller maps these onto <c> gLog </c> / <c> MenuSection.AddWarning </c>.
''' </summary>
Public Enum DiagnosticSeverity

    ''' <summary>
    ''' A note worth surfacing but not indicative of a problem. Currently issued for
    ''' undeclared tokens in registry-domain keys (could be a typo, could be a
    ''' genuine literal-bracket registry path).
    ''' </summary>
    Advisory

    ''' <summary>
    ''' A real problem that altered output: a key was dropped, or zero expansions
    ''' were produced from a referenced variable that had no values.
    ''' </summary>
    [Warning]

End Enum

''' <summary>
''' A single non-fatal note produced during expansion. The engine returns these
''' alongside the expanded values so the caller can route them to its own logging
''' and menu output without the engine knowing anything about those subsystems.
''' </summary>
Public Structure ExpansionDiagnostic

    ''' <summary>
    ''' The severity of this diagnostic — <see cref="DiagnosticSeverity.Advisory"/>
    ''' for informational notes, <see cref="DiagnosticSeverity.Warning"/> for actual
    ''' problems that altered the produced output.
    ''' </summary>
    Public Severity As DiagnosticSeverity

    ''' <summary>
    ''' Human-readable message describing the condition, with the caller-supplied
    ''' context already embedded so the caller can log the string directly.
    ''' </summary>
    Public Message As String

    ''' <summary>
    ''' Creates a diagnostic with the given severity and message
    ''' </summary>
    '''
    ''' <param name="severity">
    ''' The severity level — see <see cref="DiagnosticSeverity"/>
    ''' </param>
    '''
    ''' <param name="message">
    ''' The full message text, including any caller-supplied context (entry name,
    ''' key name, etc.) so the caller can log it without further formatting
    ''' </param>
    Public Sub New(severity As DiagnosticSeverity, message As String)

        Me.Severity = severity
        Me.Message = message

    End Sub

End Structure

''' <summary>
''' The fan-out of one template plus any diagnostics raised while producing it.
''' Returned by <see cref="VariableExpander.Expand"/>.
''' </summary>
Public Structure ExpansionResult

    ''' <summary>
    ''' The expanded strings, one per cartesian combination. Empty when the template
    ''' was dropped (filesystem-domain undeclared token, or a referenced variable
    ''' had no values).
    ''' </summary>
    Public Values As List(Of String)

    ''' <summary>
    ''' Non-fatal notes raised while expanding the template. The engine produces
    ''' these but does not log them; the caller routes each onto its own logging
    ''' surface.
    ''' </summary>
    Public Diagnostics As List(Of ExpansionDiagnostic)

    ''' <summary>
    ''' Creates a result wrapping the given values and diagnostics lists
    ''' </summary>
    '''
    ''' <param name="values">
    ''' The expanded string set — empty when the template was dropped
    ''' </param>
    '''
    ''' <param name="diagnostics">
    ''' The accompanying diagnostic notes — empty when expansion was clean
    ''' </param>
    Public Sub New(values As List(Of String), diagnostics As List(Of ExpansionDiagnostic))

        Me.Values = values
        Me.Diagnostics = diagnostics

    End Sub

End Structure

''' <summary>
''' An entry's declared list variables, looked up case-insensitively. The same set
''' is shared across every <see cref="VariableExpander.Expand"/> call for a single
''' entry, so reference tracking accumulates correctly and the caller can ask which
''' declared variables were never used.
''' </summary>
Public Class VariableSet

    ''' <summary>
    ''' Backing store for declared variables. Key matching is case-insensitive so
    ''' <c> Version= </c> binds <c> &lt;version&gt; </c>, <c> &lt;VERSION&gt; </c>,
    ''' and <c> &lt;Version&gt; </c> identically.
    ''' </summary>
    Private ReadOnly _vars As New Dictionary(Of String, List(Of String))(StringComparer.InvariantCultureIgnoreCase)

    ''' <summary>
    ''' Set of declared variable names that have been observed as <c> &lt;name&gt; </c>
    ''' tokens in at least one template, used to compute <see cref="UnreferencedNames"/>
    ''' for the open-vocabulary typo backstop.
    ''' </summary>
    Private ReadOnly _referenced As New HashSet(Of String)(StringComparer.InvariantCultureIgnoreCase)

    ''' <summary>
    ''' Adds a variable declaration with values parsed from a comma-separated list.
    ''' Trimmed, non-empty tokens only. Re-adding an existing name replaces the
    ''' prior values.
    ''' <br /><br />
    '''
    ''' Splitting is bracket-aware: only commas at <c> &lt;&gt; </c> depth zero separate
    ''' values, so an inline list inside a declaration — e.g.
    ''' <c> Root=HKCU\...\&lt;6.0,7.0,11.0&gt;\... </c> — survives as a single value for
    ''' the expander to fan out later, instead of being shredded into broken fragments
    ''' at the inline list's commas.
    ''' </summary>
    '''
    ''' <param name="name">
    ''' The variable name (case-insensitive)
    ''' </param>
    '''
    ''' <param name="csv">
    ''' The raw comma-separated value string from the source key
    ''' </param>
    Public Sub Add(name As String, csv As String)

        Dim parsed = SplitTopLevelCsv(csv) _
                        .Select(Function(s) s.Trim()) _
                        .Where(Function(s) s.Length > 0) _
                        .ToList()

        _vars(name) = parsed

    End Sub

    ''' <summary>
    ''' Splits a declaration value on its top-level commas — those at <c> &lt;&gt; </c>
    ''' bracket depth zero — leaving commas inside an inline <c> &lt;a,b,c&gt; </c> list
    ''' intact. <c> &lt;&gt; </c> nesting is forbidden by the token grammar, so a simple
    ''' depth counter suffices; an unbalanced <c> &lt; </c> keeps the remainder on one
    ''' value, which is harmless for the malformed input that would produce it.
    ''' </summary>
    '''
    ''' <param name="csv">
    ''' The raw declaration value to split
    ''' </param>
    '''
    ''' <returns>
    ''' The value split on top-level commas only; inline-list commas are preserved
    ''' </returns>
    Private Shared Function SplitTopLevelCsv(csv As String) As List(Of String)

        Dim parts As New List(Of String)
        Dim current As New System.Text.StringBuilder()
        Dim depth = 0

        For Each ch In csv

            If ch = "<"c Then

                depth += 1
                current.Append(ch)

            ElseIf ch = ">"c Then

                If depth > 0 Then depth -= 1
                current.Append(ch)

            ElseIf ch = ","c AndAlso depth = 0 Then

                parts.Add(current.ToString())
                current.Clear()

            Else

                current.Append(ch)

            End If

        Next

        parts.Add(current.ToString())
        Return parts

    End Function

    ''' <summary>
    ''' Reports whether <paramref name="name"/> was declared via <see cref="Declare"/>
    ''' </summary>
    '''
    ''' <param name="name">
    ''' The variable name to query (case-insensitive)
    ''' </param>
    '''
    ''' <returns>
    ''' <c> True </c> when <paramref name="name"/> has been declared, otherwise <c> False </c>
    ''' </returns>
    Public Function IsDeclared(name As String) As Boolean

        Return _vars.ContainsKey(name)

    End Function

    ''' <summary>
    ''' Returns the value list for a declared variable. Callers should check
    ''' <see cref="IsDeclared"/> first; this returns an empty list for unknown names
    ''' rather than throwing.
    ''' </summary>
    '''
    ''' <param name="name">
    ''' The variable name to look up (case-insensitive)
    ''' </param>
    '''
    ''' <returns>
    ''' The declared value list, or an empty list when <paramref name="name"/> is
    ''' undeclared
    ''' </returns>
    Public Function Values(name As String) As List(Of String)

        Dim result As List(Of String) = Nothing
        If _vars.TryGetValue(name, result) Then Return result
        Return New List(Of String)

    End Function

    ''' <summary>
    ''' Records that <paramref name="name"/> appeared as a <c> &lt;name&gt; </c> token
    ''' in at least one template. Idempotent. Names not present in <see cref="_vars"/>
    ''' are silently ignored, so callers can mark every token they see without first
    ''' checking <see cref="IsDeclared"/>.
    ''' </summary>
    '''
    ''' <param name="name">
    ''' The variable name observed in a template (case-insensitive)
    ''' </param>
    Public Sub MarkReferenced(name As String)

        If _vars.ContainsKey(name) Then _referenced.Add(name)

    End Sub

    ''' <summary>
    ''' Lists declared variable names that were never observed as a <c> &lt;name&gt; </c>
    ''' token in any template. Used by callers to surface the open-vocabulary typo
    ''' backstop ("declared but never used; possible typo").
    ''' </summary>
    '''
    ''' <returns>
    ''' Declared names that have never been passed to <see cref="MarkReferenced"/>,
    ''' in declaration order
    ''' </returns>
    Public Function UnreferencedNames() As List(Of String)

        Return _vars.Keys.Where(Function(n) Not _referenced.Contains(n)).ToList()

    End Function

    ''' <summary>
    ''' Lists every declared variable name regardless of reference state. Used by
    ''' <see cref="VariableExpander.ResolveAll"/> to walk the full set when building
    ''' the dependency graph.
    ''' </summary>
    Public Function DeclaredNames() As List(Of String)

        Return _vars.Keys.ToList()

    End Function

    ''' <summary>
    ''' Counts declared variables whose value references at least one <em>other</em>
    ''' declared variable (an out-edge in the dependency graph). This is a proxy for the
    ''' "topological-sort tax" a reader pays to follow a chain of nested declarations —
    ''' the cost that makes a large symbol table noise rather than abstraction. Must be
    ''' called <em>before</em> <see cref="VariableExpander.ResolveAll"/> flattens the
    ''' values, since resolution replaces the <c> &lt;token&gt; </c> references with
    ''' literals.
    ''' </summary>
    '''
    ''' <returns>
    ''' The number of declared variables that reference another declared variable in any
    ''' of their values; <c> 0 </c> when the symbol table is flat (no nesting)
    ''' </returns>
    Public Function NestedReferenceCount() As Integer

        Dim names = _vars.Keys.ToList()
        Dim count = 0

        For Each name In names

            Dim referencesAnother = False

            For Each value In _vars(name)

                For Each other In names

                    If String.Equals(other, name, StringComparison.InvariantCultureIgnoreCase) Then Continue For
                    If value.IndexOf("<" & other & ">", StringComparison.InvariantCultureIgnoreCase) >= 0 Then referencesAnother = True : Exit For

                Next

                If referencesAnother Then Exit For

            Next

            If referencesAnother Then count += 1

        Next

        Return count

    End Function

    ''' <summary>
    ''' Replaces the value list for an existing declaration. Used by
    ''' <see cref="VariableExpander.ResolveAll"/> after nested-token expansion;
    ''' callers other than the resolver should prefer <see cref="Add"/>.
    ''' </summary>
    '''
    ''' <param name="name">
    ''' The variable name (case-insensitive). Must already be declared via
    ''' <see cref="Add"/>; otherwise the call is a no-op.
    ''' </param>
    '''
    ''' <param name="newValues">
    ''' The fully-resolved replacement value list
    ''' </param>
    Public Sub Replace(name As String, newValues As List(Of String))

        If _vars.ContainsKey(name) Then _vars(name) = newValues

    End Sub

End Class

''' <summary>
''' Shared cross-DSL preprocessing engine that fans one template out over a set of
''' declared list variables. Variables are declared as entry-level keys
''' (e.g. <c> Version=9.0,10.0,11.0 </c>) and referenced via <c> &lt;name&gt; </c>
''' placeholders inside other key values; the engine returns one expanded string
''' per cartesian combination of the variables present in the template, with
''' leftmost-first occurrence varying slowest (outermost loop). A placeholder body
''' containing a comma (<c> &lt;a,b,c&gt; </c>) is an inline ad-hoc list — an
''' anonymous axis declared at its use site, for one-off leaf lists not worth a
''' named declaration.
''' <br /><br />
'''
''' The engine is string-generic — it knows nothing about winapp2 key types, ini
''' parsing, or output formatting. The caller passes one template at a time, tags
''' the call with an <see cref="ExpansionDomain"/> (so the engine can apply the
''' filesystem-vs-registry rule for undeclared tokens), and routes any returned
''' <see cref="ExpansionDiagnostic"/> values onto its own logging surface.
''' <br /><br />
'''
''' This module is the phase-2 pass in the two-phase builder DSL pipeline. Phase 1
''' is each builder's own <c> %% </c> environment-style substitution (e.g.
''' <c> %WebViewRoot% </c>, <c> %Package% </c>), which runs first and produces the
''' templates fed to this engine. The two phases compose because their token
''' domains are closed: <c> %% </c> expansion never emits <c> &lt;&gt; </c> tokens
''' and this engine never emits <c> %% </c> tokens.
''' </summary>
Public Module VariableExpander

    ''' <summary>
    ''' Regex matching one <c> &lt;name&gt; </c> placeholder. The body
    ''' (<c> [^&lt;&gt;]+ </c>) requires at least one character and forbids nesting,
    ''' so a stray <c> &lt; </c> or <c> &gt; </c> elsewhere in the template is left
    ''' alone. Compiled once at module load.
    ''' </summary>
    Private ReadOnly TokenPattern As New Regex("<([^<>]+)>", RegexOptions.Compiled)

    ''' <summary>
    ''' Expands a template into its fan-out over the declared list variables it
    ''' references. Returns one string per cartesian combination, with leftmost-first
    ''' occurrence varying slowest. A template with no declared tokens returns a
    ''' single-element list containing the template verbatim.
    ''' <br /><br />
    '''
    ''' A token whose body contains a comma is an inline ad-hoc list — e.g.
    ''' <c> &lt;System32,SysWOW64&gt; </c> is an anonymous two-value axis, fanned out
    ''' exactly like a declared variable but without a separate declaration. It is the
    ''' natural form for a single-use leaf list. Two <em>different</em> inline lists in
    ''' one template cross-product like any two axes; an <em>identical</em> inline list
    ''' repeated in one template co-varies (one axis) and emits an Advisory, because an
    ''' anonymous list cannot state same-axis intent — declare a named variable when
    ''' co-variance across positions is wanted. Inline detection takes priority over the
    ''' declared-name lookup, but declared names never contain commas in practice so the
    ''' two never collide.
    ''' <br /><br />
    '''
    ''' Undeclared <c> &lt;name&gt; </c> tokens are resolved by <paramref name="domain"/>:
    ''' filesystem keys drop the template (empty <see cref="ExpansionResult.Values"/>
    ''' plus a Warning diagnostic), registry keys keep the literal text and emit an
    ''' Advisory. A referenced variable with an empty value list also drops the
    ''' template with a Warning.
    ''' <br /><br />
    '''
    ''' The engine itself produces no logging — every diagnostic is returned in
    ''' <see cref="ExpansionResult.Diagnostics"/> for the caller to route onto
    ''' <c> gLog </c> / <c> MenuSection.AddWarning </c>. <paramref name="context"/>
    ''' is embedded verbatim in diagnostic messages so the caller does not need to
    ''' wrap them in another format string.
    ''' </summary>
    '''
    ''' <param name="template">
    ''' The single key value to expand. May contain zero or more <c> &lt;name&gt; </c>
    ''' tokens at any position, including inline <c> &lt;a,b,c&gt; </c> ad-hoc lists
    ''' </param>
    '''
    ''' <param name="vars">
    ''' The entry's declared variable set. Tokens whose name is declared in
    ''' <paramref name="vars"/> are expanded; others are handled per
    ''' <paramref name="domain"/>. Each declared token observed is marked referenced
    ''' on <paramref name="vars"/> for the caller's later typo backstop check.
    ''' </param>
    '''
    ''' <param name="domain">
    ''' Whether <paramref name="template"/> is a filesystem-domain or registry-domain
    ''' key value. Determines the behavior on undeclared tokens
    ''' </param>
    '''
    ''' <param name="context">
    ''' Caller-supplied label (e.g. entry name and key type) embedded into any
    ''' returned diagnostic messages
    ''' </param>
    '''
    ''' <returns>
    ''' The expanded value set and any diagnostics raised during expansion
    ''' </returns>
    Public Function Expand(template As String,
                            vars As VariableSet,
                            domain As ExpansionDomain,
                            context As String) As ExpansionResult

        Dim diagnostics As New List(Of ExpansionDiagnostic)

        ' Phase 1: discover tokens, building the ordered list of cartesian axes and
        ' preserving first-occurrence order across the entire template (path and
        ' value, spanning the '|'). A token whose body contains a comma is an INLINE
        ' ad-hoc list (e.g. <System32,SysWOW64>) whose axis key is the body itself and
        ' whose values are the comma-split body; otherwise the token is a declared-
        ' variable reference or an undeclared token. Inline lists never touch 'vars' —
        ' keying their axis on the body text lets the existing name-keyed cartesian and
        ' SubstituteTokens machinery handle them unchanged. Axis keys are distinct: a
        ' repeated declared name co-varies silently (intended idiom), but a repeated
        ' identical inline list co-varies WITH an advisory, since an anonymous list
        ' cannot express "the same axis as that other one" — the author who meant
        ' co-variance should declare a named variable.
        Dim axisKeys As New List(Of String)
        Dim axisValues As New List(Of List(Of String))
        Dim axisKeySet As New HashSet(Of String)(StringComparer.InvariantCultureIgnoreCase)
        Dim undeclaredNames As New List(Of String)
        Dim undeclaredSet As New HashSet(Of String)(StringComparer.InvariantCultureIgnoreCase)

        For Each match As Match In TokenPattern.Matches(template)

            Dim body = match.Groups(1).Value

            If body.IndexOf(","c) >= 0 Then

                If axisKeySet.Add(body) Then

                    axisKeys.Add(body)
                    axisValues.Add(body.Split(","c) _
                                       .Select(Function(s) s.Trim()) _
                                       .Where(Function(s) s.Length > 0) _
                                       .ToList())

                Else

                    diagnostics.Add(New ExpansionDiagnostic(DiagnosticSeverity.Advisory,
                        $"Inline list <{body}> appears more than once in {context}; occurrences co-vary — declare a named variable if that is not intended"))

                End If

            ElseIf vars.IsDeclared(body) Then

                vars.MarkReferenced(body)

                If axisKeySet.Add(body) Then

                    axisKeys.Add(body)
                    axisValues.Add(vars.Values(body))

                End If

            Else

                If undeclaredSet.Add(body) Then undeclaredNames.Add(body)

            End If

        Next

        ' Phase 2: handle undeclared tokens by domain.
        If undeclaredNames.Count > 0 Then

            If domain = ExpansionDomain.Filesystem Then

                For Each name In undeclaredNames

                    diagnostics.Add(New ExpansionDiagnostic(DiagnosticSeverity.Warning,
                        $"Undeclared variable <{name}> in filesystem-domain key {context}; dropping key"))

                Next

                Return New ExpansionResult(New List(Of String), diagnostics)

            Else

                For Each name In undeclaredNames

                    diagnostics.Add(New ExpansionDiagnostic(DiagnosticSeverity.Advisory,
                        $"Undeclared variable <{name}> in registry-domain key {context}; emitted as literal"))

                Next

            End If

        End If

        ' Phase 3: any axis with no values yields an empty cartesian product, so the
        ' template would expand to nothing — drop it with a Warning. Covers both a
        ' declared variable with an empty value list and a degenerate inline list
        ' (e.g. <,,> or < , >) whose body trims to nothing.
        For axisIdx = 0 To axisKeys.Count - 1

            If axisValues(axisIdx).Count = 0 Then

                diagnostics.Add(New ExpansionDiagnostic(DiagnosticSeverity.Warning,
                    $"Axis <{axisKeys(axisIdx)}> has no values, referenced by {context}; dropping key"))

                Return New ExpansionResult(New List(Of String), diagnostics)

            End If

        Next

        ' Phase 4: cartesian enumerate over the axes. Leftmost-first means the axis at
        ' index 0 varies slowest (outermost loop) — decoded by walking the index
        ' dimensions back-to-front when computing the modulus.
        Dim totalCombinations = 1
        For Each lst In axisValues : totalCombinations *= lst.Count : Next

        Dim values As New List(Of String)(totalCombinations)

        For combination = 0 To totalCombinations - 1

            Dim chosen As New Dictionary(Of String, String)(StringComparer.InvariantCultureIgnoreCase)
            Dim remaining = combination

            For dimIdx = axisKeys.Count - 1 To 0 Step -1

                Dim valuesForDim = axisValues(dimIdx)
                chosen(axisKeys(dimIdx)) = valuesForDim(remaining Mod valuesForDim.Count)
                remaining = remaining \ valuesForDim.Count

            Next

            values.Add(SubstituteTokens(template, chosen))

        Next

        Return New ExpansionResult(values, diagnostics)

    End Function

    ''' <summary>
    ''' Resolves nested <c> &lt;name&gt; </c> references inside the value lists of
    ''' <paramref name="vars"/> itself, before any user template is expanded. Each
    ''' variable's values are fanned out against the variables they reference, in
    ''' topological order so each dependency resolves before its dependents. Cyclic
    ''' dependencies are detected and left unresolved with a Warning.
    ''' <br /><br />
    '''
    ''' After this call returns, every variable's value list contains literal strings
    ''' (or, for variables in a cycle, the original templated strings). User templates
    ''' that reference these variables therefore receive fully-resolved substitution
    ''' without the engine needing a second expansion pass.
    ''' <br /><br />
    '''
    ''' Reference tracking on <paramref name="vars"/> IS updated during resolution: a
    ''' variable consumed only via a nested chain (e.g. <c> RegRoots=...&lt;HKLMVersions&gt;... </c>
    ''' with no template referencing <c> &lt;HKLMVersions&gt; </c> directly) still counts
    ''' as referenced, so the unreferenced-variable typo backstop does not false-positive
    ''' on it. A variable declared but reached by neither a user template nor a nested
    ''' chain remains unreferenced and is reported by <see cref="VariableSet.UnreferencedNames"/>.
    ''' </summary>
    '''
    ''' <param name="vars">
    ''' The variable set to resolve in place
    ''' </param>
    '''
    ''' <returns>
    ''' Diagnostics raised during resolution — cycle warnings, undeclared-token
    ''' advisories inside variable definitions, etc. The caller routes them onto
    ''' <c> gLog </c> / <c> MenuSection.AddWarning </c>.
    ''' </returns>
    Public Function ResolveAll(vars As VariableSet) As List(Of ExpansionDiagnostic)

        Dim diagnostics As New List(Of ExpansionDiagnostic)
        Dim allNames = vars.DeclaredNames()

        ' Build dependency map: which other declared variables each variable's
        ' value list references. Undeclared tokens don't contribute to deps.
        Dim deps As New Dictionary(Of String, HashSet(Of String))(StringComparer.InvariantCultureIgnoreCase)

        For Each name In allNames

            Dim varDeps As New HashSet(Of String)(StringComparer.InvariantCultureIgnoreCase)

            For Each value In vars.Values(name)

                For Each m As Match In TokenPattern.Matches(value)

                    Dim refName = m.Groups(1).Value
                    If vars.IsDeclared(refName) Then varDeps.Add(refName)

                Next

            Next

            deps(name) = varDeps

        Next

        ' Topological resolution: in each pass, resolve every variable whose
        ' dependencies are already resolved. Stop when no progress is made.
        Dim resolved As New HashSet(Of String)(StringComparer.InvariantCultureIgnoreCase)

        Do

            Dim progress = False

            For Each name In allNames

                If resolved.Contains(name) Then Continue For

                Dim allDepsResolved = True
                For Each d In deps(name)
                    If Not resolved.Contains(d) Then allDepsResolved = False : Exit For
                Next
                If Not allDepsResolved Then Continue For

                Dim newValues As New List(Of String)

                For Each value In vars.Values(name)

                    ' Registry domain so undeclared tokens stay literal with an
                    ' advisory rather than being silently dropped. The user
                    ' template's later domain still applies its own rules.
                    Dim result = Expand(value, vars, ExpansionDomain.Registry, $"<{name}> declaration")
                    diagnostics.AddRange(result.Diagnostics)
                    newValues.AddRange(result.Values)

                Next

                vars.Replace(name, newValues)
                resolved.Add(name)
                progress = True

            Next

            If Not progress Then Exit Do

        Loop

        ' Anything still unresolved is in a cycle. Leave values as authored and warn.
        For Each name In allNames

            If resolved.Contains(name) Then Continue For

            Dim unresolvedDeps = deps(name).Where(Function(d) Not resolved.Contains(d)).Select(Function(d) $"<{d}>").ToList()
            Dim cycleMsg = $"Variable <{name}> participates in a cyclic dependency with {String.Join(", ", unresolvedDeps)}; values left unresolved"
            diagnostics.Add(New ExpansionDiagnostic(DiagnosticSeverity.Warning, cycleMsg))

        Next

        Return diagnostics

    End Function

    ''' <summary>
    ''' Replaces every <c> &lt;name&gt; </c> token in <paramref name="template"/>
    ''' that matches an entry in <paramref name="chosen"/> with the corresponding
    ''' value, case-insensitively. Tokens not present in <paramref name="chosen"/>
    ''' are left literal so registry-domain undeclared tokens survive untouched.
    ''' </summary>
    '''
    ''' <param name="template">
    ''' The string in which to substitute tokens
    ''' </param>
    '''
    ''' <param name="chosen">
    ''' The current iteration's variable-name → value bindings (case-insensitive)
    ''' </param>
    '''
    ''' <returns>
    ''' The template with all matching tokens replaced
    ''' </returns>
    Private Function SubstituteTokens(template As String,
                                       chosen As Dictionary(Of String, String)) As String

        Return TokenPattern.Replace(template, Function(m)

                                                  Dim name = m.Groups(1).Value
                                                  Dim chosenValue As String = Nothing
                                                  If chosen.TryGetValue(name, chosenValue) Then Return chosenValue
                                                  Return m.Value

                                              End Function)

    End Function

End Module
