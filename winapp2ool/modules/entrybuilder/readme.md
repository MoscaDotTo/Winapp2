# EntryBuilder

**EntryBuilder** is a winapp2ool module that generates winapp2.ini entries from a shorthand DSL. You write one ini section per application: anything that is already valid winapp2 syntax passes through unchanged, and a small set of winapp2ool-private keys expand into standard winapp2 form at generation time. Think of it as a preprocessor for winapp2.ini: declare the repetitive parts once and let the tool write them out.

EntryBuilder is a devops module used to maintain the base entries of the winapp2.ini distribution. Source files at `Assembler\EntryBuilder\{letter}.ini` are combined at runtime and processed alongside the shared scaffold catalogs in `Assembler\Scaffolds` to produce a single intermediate `entrybuilder.ini`, which the build pipeline then merges into the final database.

### What does EntryBuilder do?

EntryBuilder reads every `*.ini` file in its source directory, combines them in-memory, parses each section into an entry definition, expands the shorthand, normalizes the result via WinappDebug, and writes the output to `entrybuilder.ini`. The two main shorthand categories are:

- **Scaffold families**: entries that declare a `WebViewRoot=` (embedded WebView2/EBWebView data folder) or `QtWebEngineRoot=` (embedded QtWebEngine data folder) receive baseline cleaning FileKeys drawn from a shared catalog, one set per declared root.
- **List variables**: entries can declare comma-separated lists (`Versions=10.0,11.0,12.0`) and reference them from any key with `<Versions>` tokens; each referencing key fans out into one output key per combination.

### Why EntryBuilder?

- **Shorthand DSL**: Replaces 20-line repetitive entries with a few list-variable declarations and templated keys. 
- **Shared scaffold catalogs**: One catalog of Chromium cleaning patterns covers every embedded WebView2 host. Discord, Slack, VS Code, Teams, and dozens of small Electron apps get consistent baseline coverage without per-app duplication. A second catalog does the same for QtWebEngine hosts.
- **Pass-through compatible**: Any standard winapp2 key passes through unchanged. You can paste a raw entry in and incrementally enrich it with shorthand, or leave it untouched.
- **Consistent DSL surface**: The same `<Name>` variable syntax and scaffold key names work across winapp2ool's builder modules.

---

# Table of Contents

1. [Requirements](#requirements)
2. [Quick Start](#quick-start)
3. [Menu Options](#menu-options)
4. [How Generation Works](#how-generation-works)
5. [Source File Format](#source-file-format)
6. [Pass-Through Keys](#pass-through-keys)
7. [Shorthand Keys](#shorthand-keys)
   - [WebViewRoot](#webviewroot)
   - [WebViewScaffolds / ExcludeWebViewScaffolds](#webviewscaffolds--excludewebviewscaffolds)
   - [QtWebEngineRoot / QtWebEngineScaffolds](#qtwebengineroot--qtwebenginescaffolds)
   - [FileKeyBase / RegKeyBase / ExcludeKeyBase](#filekeybase--regkeybase--excludekeybase)
   - [Skip](#skip)
8. [List Variables and Token Expansion](#list-variables-and-token-expansion)
   - [Declaration](#declaration)
   - [Reference](#reference)
   - [Inline Lists](#inline-lists)
   - [Cartesian Fan-out and Co-variance](#cartesian-fan-out-and-co-variance)
   - [Nested References](#nested-references)
   - [The Reserved Root Variable](#the-reserved-root-variable)
   - [Undeclared Tokens](#undeclared-tokens)
   - [Typo Backstops](#typo-backstops)
9. [Validation and Warnings](#validation-and-warnings)
10. [Run Output](#run-output)
11. [Command-Line Arguments](#command-line-arguments)
    - [File Selection](#file-selection)
    - [Examples](#examples)
12. [Tips & Best Practices](#tips--best-practices)
13. [Troubleshooting](#troubleshooting)
14. [Usage Examples](#usage-examples)
    - [Pass-Through and Variables](#pass-through-and-variables)
      - [Example 1: A pass-through entry with a single-value variable](#example-1-a-pass-through-entry-with-a-single-value-variable)
      - [Example 2: Automatic detection from the reserved Root variable](#example-2-automatic-detection-from-the-reserved-root-variable)
      - [Example 3: List-variable fan-out](#example-3-list-variable-fan-out)
      - [Example 4: Co-variance of repeated variables](#example-4-co-variance-of-repeated-variables)
      - [Example 5: Inline lists](#example-5-inline-lists)
      - [Example 6: Nested references](#example-6-nested-references)
      - [Example 7: Literal angle brackets in registry keys](#example-7-literal-angle-brackets-in-registry-keys)
    - [Scaffold Families](#scaffold-families)
      - [Example 8: Default WebView scaffolds](#example-8-default-webview-scaffolds)
      - [Example 9: Everything except All with exclusions](#example-9-everything-except-all-with-exclusions)
      - [Example 10: QtWebEngine and a non-Default profile](#example-10-qtwebengine-and-a-non-default-profile)
    - [Housekeeping](#housekeeping)
      - [Example 11: Skipping a defunct entry](#example-11-skipping-a-defunct-entry)

---

# Requirements

- A source directory containing one or more `*.ini` files with at least one parseable section
- The shared WebView scaffold catalog (typically `Assembler\Scaffolds\webview.ini`) if any entry declares `WebViewRoot=`
- The shared QtWebEngine scaffold catalog (typically `Assembler\Scaffolds\qtwebengine.ini`) if any entry declares `QtWebEngineRoot=`

If the source directory is empty, missing, or contains no parseable sections, EntryBuilder reports `No EntryBuilder source definitions found in: <directory>` and writes no output. If a scaffold catalog is missing or empty, generation continues with zero scaffold FileKeys emitted from that catalog and a warning logged. Entries with no corresponding root key are unaffected.

---

# Quick Start

### Common Workflow

1. Place per-letter source files (`A.ini`, `B.ini`, ...) in the source directory
2. Open EntryBuilder from the **Entry Lab** main menu, or invoke `winapp2ool -entrybuilder` from the command line
3. Run. EntryBuilder combines the files, expands the shorthand, and writes the result to `entrybuilder.ini`

The default source directory and save target are both the current directory. The default scaffold catalogs are `webview.ini` and `qtwebengine.ini` in the current directory; in normal use, `entrybuilder.ini` and both catalogs live next to the winapp2ool executable, and the source directory points at `..\..\Assembler\EntryBuilder\`.

---

# Menu Options

| Option | Effect | Notes |
|:-|:-|:-|
| Run (default) | Generate entries from the configured source directory | Does nothing if no parseable sections found |
| Choose source directory | Select the directory containing per-letter source files | Only the directory is used; the file name is ignored |
| Choose save target | Select where to save the generated entries | Default: `entrybuilder.ini` in current directory |
| Choose webview scaffolds | Select the shared WebView scaffold catalog | Default: `webview.ini` in current directory |
| Choose QtWebEngine scaffolds | Select the shared QtWebEngine scaffold catalog | Default: `qtwebengine.ini` in current directory |
| Reset Settings | Restore all settings to their defaults | Only shown when settings have been changed |

---

# How Generation Works

Each run proceeds through the same fixed steps:

1. **Combine**: every `*.ini` in the source directory is loaded in alphabetical filename order and concatenated in-memory. 
2. **Parse**: each section becomes one entry definition. Recognised winapp2 and shorthand keys are claimed by name; *any other key* is treated as a list-variable declaration.
3. **Resolve variables**: variable values that reference other variables (`<Other>` tokens inside a declaration) are resolved so every value list is fully literal. Cyclic references are warned about and not expanded.
4. **Infer detection**: if the entry declares a variable named `Root`, a `Detect=<Root>` or `DetectFile=<Root>` is generated automatically (see [The Reserved Root Variable](#the-reserved-root-variable)).
5. **Validate**: structurally invalid entries are skipped with warnings; questionable ones are emitted with warnings (see [Validation and Warnings](#validation-and-warnings)).
6. **Generate**: keys are emitted per entry in canonical winapp2 order. Each content key goes through, in order: scaffold expansion (catalog templates copied in for each selected scaffold and declared root), **root substitution** (`%WebViewRoot%` / `%QtWebEngineRoot%` replaced with each declared root), then **token expansion** (`<Name>` fan-out). Numbered key families are renumbered from 1.
7. **Normalize**: the whole in-memory file is passed through WinappDebug with all repairs and optimizations enabled.
8. **Save**: the output file is written with a generated comment header (see [Run Output](#run-output)).

Because root substitution (step 6) happens before token expansion, `<Name>` tokens inside a `WebViewRoot=` / `QtWebEngineRoot=` value survive substitution and are expanded like any other token in the resulting key.

Not every key family receives every processing step:

| Key family | Root placeholders | Expansion domain |
|:-|:-|:-|
| `Detect` | No  | Registry |
| `DetectFile` | No  | Filesystem |
| `FileKey` / `FileKeyBase` | Yes | Filesystem |
| `RegKey` / `RegKeyBase` |  Yes | Registry |
| `ExcludeKey` / `ExcludeKeyBase` |  Yes | By flag: `REG` → Registry, otherwise Filesystem |
| Scaffold catalog templates | Yes (their purpose) | Filesystem |

The expansion domain matters when a `<Name>` token doesn't match any declared variable, see [Undeclared Tokens](#undeclared-tokens). Note that `%WebViewRoot%` / `%QtWebEngineRoot%` are **not** substituted in `Detect` / `DetectFile`. Write the path (or a `<Variable>`) directly in detection keys.

---

# Source File Format

EntryBuilder source files are standard ini format. Each section describes one application and corresponds to exactly one output winapp2.ini entry; the section header is the entry name and is round-tripped verbatim into the output.

###### Example input

```ini
[A-PDF PDF Content Splitter *]
Root=HKCU\Software\A-PDF\PDFContentSplitter

LangSecRef=3021
RegKeyBase=<Root>\Setting|<hloadpdfpath,hSaveallpdfpath,LastRulename>
```

###### Example output 

```ini
[A-PDF PDF Content Splitter *]
LangSecRef=3021
Detect=HKCU\Software\A-PDF\PDFContentSplitter
RegKey1=HKCU\Software\A-PDF\PDFContentSplitter\Setting|hloadpdfpath
RegKey2=HKCU\Software\A-PDF\PDFContentSplitter\Setting|hSaveallpdfpath
RegKey3=HKCU\Software\A-PDF\PDFContentSplitter\Setting|LastRulename
```

An application requiring multiple sibling entries (e.g. a primary entry for caches and a separate entry for login data) is expressed as multiple sections, repeating shared metadata.

In the production tree, entries are split per-letter (`#.ini` through `Z.ini`), but EntryBuilder itself only requires at least 1 ini.

---

# Pass-Through Keys

Any of the following standard winapp2 keys passes through into the output, subject to the processing steps in the [table above](#how-generation-works):

| Key | Notes |
|:-|:-|
| `Section` | Mutually exclusive with `LangSecRef`; declaring both warns and `LangSecRef` wins |
| `LangSecRef` | See above |
| `Detect` | Multiple values (after fan-out) produce numbered `Detect1`, `Detect2`, ...; a single value is emitted as bare `Detect` |
| `DetectFile` | Same numbering rule as `Detect` |
| `FileKey` | Same treatment as `FileKeyBase`, including root substitution; renumbered after expansion |
| `RegKey` | Same treatment as `RegKeyBase`; renumbered after expansion |
| `ExcludeKey` | Same treatment as `ExcludeKeyBase`; renumbered after expansion |


Key numbers in the source files are ignored. Every numbered key family is renumbered from 1 after expansion. Key numbers have been stripped from the EntryBuilder source files to improve readability. 

The `Base=` and plain forms of each content key behave identically in the output; the distinction only affects the run statistics, which count `FileKeyBase=` / `RegKeyBase=` / `ExcludeKeyBase=` as *generated* content and the plain forms as *pass-through* content (see [Run Output](#run-output)). Convention: use the `Base=` form when the key is a template you expect to expand, the plain form when pasting literal keys.

---

# Shorthand Keys

These are winapp2ool-private keys that are consumed during generation and never appear in the output.

## WebViewRoot

The full path to a Chromium-data folder for this application, typically `%AppData%\<AppName>` or `%LocalAppData%\<AppName>\<Subfolder>\EBWebView`.

- Multiple `WebViewRoot=` keys are permitted; each scaffold FileKey is generated once per root.
- The literal string `%WebViewRoot%` can be used in any `FileKey` / `RegKey` / `ExcludeKey` value (`Base=` or plain form); it is substituted against each declared root at generation time, fanning the key out per root.
- An entry with no roots and no FileKey/RegKey content is skipped with a warning.

## WebViewScaffolds / ExcludeWebViewScaffolds

`WebViewScaffolds=` is a comma-separated list of scaffold names from the shared catalog to apply to this entry.

| Form | Behavior |
|:-|:-|
| Key absent | Default scaffold set is used: `Caches`, `Telemetry` |
| Key present, value empty | No scaffold FileKeys are generated |
| `WebViewScaffolds=Caches,Telemetry,DRMData` | Exactly the named scaffolds are applied |
| `WebViewScaffolds=All` | Expands to every scaffold in the catalog, **including host-risk categories** (cookies, history, sessions, web storage, login data) |

`ExcludeWebViewScaffolds=` is a comma-separated list of scaffold names to subtract from the active selection. The natural idiom for "everything except X":

###### Example

```ini
WebViewScaffolds=All
ExcludeWebViewScaffolds=WebStorage,LoginData
```

Combining a non-`All` explicit `WebViewScaffolds=` list with `ExcludeWebViewScaffolds=` warns (`Both WebViewScaffolds and ExcludeWebViewScaffolds set in [Entry]; applying exclusions to explicit list`), since the same effect is usually more clearly expressed by listing only the desired scaffolds. Unknown scaffold names are dropped with a warning.

Scaffold names track the catalog, not this document. At the time of writing the WebView catalog defines: `Autofill`, `Autoplay`, `BookmarkBackups`, `BookmarkFavicons`, `Caches`, `DefaultApps`, `DownloadHistory`, `DRMData`, `ExtensionCookies`, `ProgressiveWebApps`, `PrivacySandbox`, `LoginData`, `Security`, `Shopping`, `StorageQuota`, `Telemetry`, `WebCookies`, `WebHistory`, `WebSession`, `WebStorage`. Only `Caches` and `Telemetry` are default-on; scaffolds that can remove user-visible state (`WebCookies`, `WebHistory`, `WebSession`, `WebStorage`, `LoginData`) always require explicit opt-in.

## QtWebEngineRoot / QtWebEngineScaffolds

QtWebEngine is a second, independent scaffold family for apps that embed Qt's WebEngine rather than WebView2. It behaves **exactly like the WebView family above**, with these substitutions:

| WebView family | QtWebEngine family |
|:-|:-|
| `WebViewRoot=` | `QtWebEngineRoot=` |
| `WebViewScaffolds=` | `QtWebEngineScaffolds=` |
| `ExcludeWebViewScaffolds=` | `ExcludeQtWebEngineScaffolds=` |
| `%WebViewRoot%` placeholder | `%QtWebEngineRoot%` placeholder |
| `webview.ini` catalog | `qtwebengine.ini` catalog |

The default set is the same (`Caches`, `Telemetry`), the `All` sentinel and exclusion idiom work identically, and a single entry may declare **both** families (each expands against its own catalog and placeholder). The QtWebEngine catalog is smaller: `Caches`, `Telemetry`, `WebCookies`, `WebHistory`, `WebSession`, `WebStorage`.

Two QtWebEngine-specific notes:

- **Root points at the `QtWebEngine` folder**, not a profile inside it. The catalog templates bake the `\Default\` profile segment in, because the overwhelming majority of QtWebEngine hosts use the Default profile. So `QtWebEngineRoot=%LocalAppData%\AppName\QtWebEngine` is correct; do not append `\Default`.
- **Non-`Default` profiles** (e.g. calibre's `OffTheRecord` / `viewer-lookup`) are not covered by the scaffolds. Use `FileKeyBase=%QtWebEngineRoot%\OffTheRecord\...` for those. The placeholder substitutes in any content key.

## FileKeyBase / RegKeyBase / ExcludeKeyBase

Non-scaffold templates, applied **in addition to** the catalog scaffolds. Each behaves exactly like its plain-form counterpart (`FileKey=` etc. see [Pass-Through Keys](#pass-through-keys)): root substitution, then token expansion, then renumbering from 1. Use these for app-specific cleaning targets outside the catalog scaffolds: application logs, crash dumps, version-numbered paths, non-`Default` QtWebEngine profiles, etc.

## Skip

Tells EntryBuilder to omit this entry from generation. Used to discontinue support for an entry without losing its configuration.

```ini
Skip=
```

**The value of `Skip=` is ignored**: `Skip=False`, `Skip=True`, and `Skip=0` all skip the entry. Provide the key if and only if you want the entry skipped.

---

# List Variables and Token Expansion

Repeated content (version numbers, registry roots, path fragments, key suffixes) can be declared once as a list variable on the entry and referenced from other keys using `<Name>` tokens. At generation time, EntryBuilder fans each referencing key out into one output key per cartesian combination of the variables it references. Keys with no `<>` tokens pass through unchanged.

## Declaration

Any entry-level key whose name is not a recognised winapp2 key or winapp2ool shorthand key is interpreted as a list-variable declaration. The value is a comma-separated list with whitespace trimmed from each token. Variable names are case-insensitive.

###### Example

```ini
Versions=9.0,10.0,11.0,12.0,14.0,15.0,16.0
MRUNum=1,2,3
RegRoot=HKCU\Software\Microsoft\Office
```

A single-value declaration (`RegRoot` above) is permitted and useful for abbreviating long literal strings even when no fan-out is needed.

## Reference

`<Name>` tokens can appear anywhere inside any key value: path, pattern, flag region, on either side of the `|` separator, in any key that participates in expansion (see the [key-family table](#how-generation-works)). Tokens are matched against the entry's declared variables case-insensitively.

```ini
Versions=9.0,10.0,11.0,12.0,14.0,15.0,16.0
MRUNum=1,2,3
RegRoot=HKCU\Software\Microsoft\Office

Detect=<RegRoot>\<Versions>\Outlook
RegKeyBase=<RegRoot>\<Versions>\Outlook\Office Finder|<MRUNum>
```

## Inline Lists

A token whose body contains a comma is an **inline list**: an anonymous, single-use variable that needs no declaration:

```ini
RegKeyBase=HKCU\Software\7-Zip\<Compression,Extraction>|ArcHistory
```

fans out into two RegKeys, exactly as if `<X>` referenced a declared two-value variable. Inline lists are the natural form for a short list used in exactly one key; promote to a named declaration once the list is reused or grows large.

Two *different* inline lists in one key cross-multiply like any two variables. An *identical* inline list repeated within one key co-varies (both occurrences take the same value in each output key). This emits a message in the log, because an anonymous list can't state that intent; declare a named variable when you want co-variance across positions.

Inline lists also work inside declaration values. A declaration's comma-split respects `<>` brackets (only commas outside a token separate list items) and the inline list is flattened into the variable's value list before anything references it. So

```ini
Root=HKCU\Software\Ability <6.0,7.0,11.0>\Ability Write
```

declares a three-value `Root`, which composes with [Root detection inference](#the-reserved-root-variable) to produce `Detect1` through `Detect3`.

## Cartesian Fan-out and Co-variance

A key that references N variables of sizes A, B, C, ... produces A × B × C × ... output keys, one per combination. The leftmost reference varies slowest (outermost loop), so the `RegKeyBase` example under [Reference](#reference) produces 7 × 3 = 21 RegKey lines ordered as `(Versions[0], MRUNum[0])`, `(Versions[0], MRUNum[1])`, `(Versions[0], MRUNum[2])`, `(Versions[1], MRUNum[0])`, ...

Repeated references to the same *declared* variable within one key **co-vary**: they all bind to the same value within a single combination rather than producing an inner loop. This is the right tool for "version N's install path and version N's log folder": they always pair, never cross-multiply (see [Example 4](#example-4-co-variance-of-repeated-variables)).

## Nested References

A variable's value list can itself contain `<Other>` tokens that reference other declared variables. EntryBuilder resolves these before any key is expanded, so each referencing key sees a fully-literal value list.

```ini
Versions=10.0,10.1
ArcMapKey=ArcMap,Desktop<Versions>\ArcMap
Detect=HKCU\Software\ESRI\<ArcMapKey>
```

resolves `ArcMapKey` to `(ArcMap, Desktop10.0\ArcMap, Desktop10.1\ArcMap)`, then `Detect` fans across that 3-value list. Cyclic dependencies (A references B references A) are reported (`Variable <A> participates in a cyclic dependency with B; values left unresolved`) and the involved variables are passed through unresolved. 

## The Reserved Root Variable

The variable name `Root` is reserved and carries an extra behavior: **declaring `Root=` guarantees the entry a detection key derived from it.** EntryBuilder classifies each `Root` value as registry (leading segment is a registry hive) or filesystem (anything else) and injects the token `<Root>` as the first `Detect` (registry) or `DetectFile` (filesystem) of the entry:

```ini
[A Hat in Time *]
Root=%ProgramFiles%\Steam\steamapps\common\HatinTime\HatinTimeGame

Section=Games
FileKeyBase=<Root>\Logs|*
```

generates `DetectFile=%ProgramFiles%\Steam\steamapps\common\HatinTime\HatinTimeGame` without the entry declaring any detection key. The injected token expands like any other, so a multi-valued `Root` produces `DetectFile1`, `DetectFile2`, ..., etc.

Rules and edge cases:

- Injection is **idempotent**: a hand-written `Detect=<Root>` / `DetectFile=<Root>` is never duplicated, and any other detection keys are left untouched. `Root` inference only ever *adds* a missing anchor.
- Because inference marks `Root` referenced, a `Root`-only entry does not trigger the unreferenced-variable warning.
- A `Root` list mixing registry and filesystem values warns (`[Entry] Root mixes registry and filesystem values; inferring detection from the first value's domain`). Split heterogeneous roots into separate variables instead.
- A `Root` whose value still contains unresolved `<tokens>` (typically from a cycle) skips inference with a warning.
- **Want a base-path variable *without* automatic detection? Name it anything else**: `DiskRoot`, `AppRoot`, etc. (see [Example 1](#example-1-a-pass-through-entry-with-a-single-value-variable)).

## Undeclared Tokens

Resolution of `<Name>` tokens whose name matches neither a declared variable nor an inline list depends on the key's expansion domain (see the [key-family table](#how-generation-works)):

| Domain | Behavior | Message | Why |
|:-|:-|:-|:-|
| Filesystem | Drop the key | `Undeclared variable <X> in filesystem-domain key [Entry].FileKey; dropping key` | Windows path syntax forbids `<` and `>`, so any such token is unambiguously a typo |
| Registry | Emit the literal text | `Undeclared variable <X> in registry-domain key [Entry].RegKey; emitted as literal` | Registry key names can legitimately contain `<` and `>`; the engine cannot assume a typo |

The filesystem-domain drop is a warning shown in the menu output; the registry-domain literal emission is an advisory written to the log only (see [Validation and Warnings](#validation-and-warnings)).

## Typo Backstops

EntryBuilder warns at generation time if a declared variable is never referenced by any `<Name>` token in any expanded key: `Variable 'Versions=' declared in [Entry] but never referenced by any <Versions> token; possible typo`. This catches mis-spelled references (`Versions=` declared but only `<Versoin>` appearing in keys) that would otherwise fail silently in the registry domain.

A variable declaration whose name collides with a reserved key name also warns (`Variable declaration 'X=' in [Entry] shadows reserved key name; possible typo`).

---

# Validation and Warnings

Warnings are shown in the menu output after a run and recorded in the winapp2ool log; advisories (marked below) are recorded in the log only. Messages are shown here with `[Entry]` standing in for the entry name and `X` / `N` for the specific value.

| Condition | Outcome | Message |
|:-|:-|:-|
| Both `Section` and `LangSecRef` declared | `LangSecRef` wins, `Section` dropped | `Both Section and LangSecRef present in [Entry], using LangSecRef` |
| Neither `Section` nor `LangSecRef` declared | Entry **skipped** | `No Section or LangSecRef in [Entry], skipping` |
| No root keys and no FileKey/RegKey content in any form | Entry **skipped** | `[Entry] declares no WebViewRoot, QtWebEngineRoot, FileKeyBase, or RegKeyBase - nothing to emit, skipping` |
| No detection (after `Root` inference) | Entry emitted, always-on | `[Entry] declares no detection (Detect / DetectFile / DetectOS); generated entry will be always-on` |
| `Default=` declared | Key dropped | `Default= declared in [Entry]; EntryBuilder never emits Default, ignoring` |
| `All` mixed with other scaffold names | Extras ignored | `WebViewScaffolds=All in [Entry] with redundant additional names (X, Y); ignoring` |
| Explicit scaffold list mixed with exclusions | Exclusions applied | `Both WebViewScaffolds and ExcludeWebViewScaffolds set in [Entry]; applying exclusions to explicit list` |
| Unknown scaffold name requested | Scaffold dropped | `Unknown WebView scaffold 'X' requested by [Entry], skipping` |
| Scaffold catalog missing or empty | Zero scaffold keys from that catalog | `WebViewScaffold: catalog at <path> is empty or missing` |
| Undeclared `<X>` token, filesystem domain | Key dropped | `Undeclared variable <X> in filesystem-domain key [Entry].FileKey; dropping key` |
| Undeclared `<X>` token, registry domain | Literal emitted *(advisory)* | `Undeclared variable <X> in registry-domain key [Entry].RegKey; emitted as literal` |
| Referenced variable or inline list has no values | Key dropped | `Axis <X> has no values, referenced by [Entry].FileKey; dropping key` |
| Cyclic variable references | Values left as authored | `Variable <X> participates in a cyclic dependency with Y; values left unresolved` |
| Identical inline list repeated in one key | Occurrences co-vary *(advisory)* | `Inline list <a,b> appears more than once in [Entry].FileKey; occurrences co-vary - declare a named variable if that is not intended` |
| `Root` value has unresolved tokens | Detection inference skipped | `[Entry] Root has unresolved <token> references; skipping detection inference` |
| `Root` mixes registry and filesystem values | First value's domain wins | `[Entry] Root mixes registry and filesystem values; inferring detection from the first value's domain` |
| Declared variable never referenced | Warning only | `Variable 'X=' declared in [Entry] but never referenced by any <X> token; possible typo` |
| Variable name collides with a reserved key | Warning only | `Variable declaration 'X=' in [Entry] shadows reserved key name; possible typo` |

The scaffold-family messages appear with `QtWebEngine` in place of `WebView` for the QtWebEngine family.

---

# Run Output

### The output file

After expansion, the entire result is normalized by WinappDebug with all repairs and optimizations enabled, so the final key order and numbering follow WinappDebug's alphabetization, not the order keys appear in your source or the order EntryBuilder generated them.

The output file begins with a generated comment header:

```ini
; Version <YYMMDD>
; # of entries: <count>
; entrybuilder.ini is generated by the Winapp2ool Entry Builder
; Entries in this file may be incomplete and are not intended to be used directly with any cleaning software
; They are utilized by winapp2ool to create the final winapp2.ini file for distribution
; If you are not maintaining winapp2.ini for distribution, you probably don't need this file!
; Refer to the Winapp2ool documentation for more information: <url>
```

This file is consumed by the build pipeline and merged with the base entries, BrowserBuilder, and UWPBuilder outputs to produce the final winapp2.ini.

### The statistics summary

Each run also renders a statistics box covering: entry counts (read / generated / skipped / scaffold-bearing / `Root`-inferred), shorthand declaration counts, the generated-key breakdown by category, and a generated vs pass-through provenance split: how many output keys the builder produced from templates (scaffolds + `Base=` forms) versus how many were passed through from literal declarations. An abstraction payoff section watch-lists "tangled" entries whose variables heavily reference other variables. 

---

# Command-Line Arguments

EntryBuilder is invoked as `winapp2ool -entrybuilder`. Combine with the global `-s` flag for silent (non-interactive) execution in the build pipeline.

### File Selection

Each file slot has a corresponding index for use with the `-Nd` (directory) and `-Nf` (file name) argument pattern:

| Index | File | Default |
|:-|:-|:-|
| 1 | Source directory | Current directory (directory only; file name is ignored) |
| 2 | Save target | `entrybuilder.ini` in current directory |
| 3 | WebView scaffold catalog | `webview.ini` in current directory |
| 4 | QtWebEngine scaffold catalog | `qtwebengine.ini` in current directory |

| Arg | Effect |
|:-|:-|
| `-Nd path` | Set directory for file slot N |
| `-Nf name` | Set file name for file slot N |
| `-Nf subdir\name` | Set file name within a subdirectory of its path |

### Examples

| Command | Effect |
|:-|:-|
| `winapp2ool -entrybuilder` | Run with current settings |
| `winapp2ool -entrybuilder -1d ..\..\Assembler\EntryBuilder` | Read sources from the assembler folder, save `entrybuilder.ini` in the current directory |
| `winapp2ool -entrybuilder -1d ..\..\Assembler\EntryBuilder -3d ..\..\Assembler\Scaffolds -4d ..\..\Assembler\Scaffolds -s` | Full build-pipeline invocation: source and both catalogs from sibling folders, silent mode |
| `winapp2ool -entrybuilder -2f test-output.ini` | Override the save target file name |

---

# Tips & Best Practices

### Start with pass-through, enrich incrementally

The DSL is opt-in. Pasting a raw winapp2.ini entry into a source file produces an identical output entry. Add shorthand only when it removes repetition.

### Use single-value variables to name long literals

Even with no fan-out, `RegRoot=HKCU\Software\Vendor\Product` lets you write `<RegRoot>\Settings|Foo` instead of repeating the full path on every key. 

### `Root` is a contract, not just a name

Naming a variable `Root` means "this is the app's anchor path, derive detection from it." If you want the abbreviation without the inferred detection key, pick another name (`DiskRoot`, `AppRoot`). 

### Inline lists for one-offs, declarations for reuse

`<System32,SysWOW64>` beats declaring a variable for a list used exactly once. The moment the same list appears in a second key or two occurrences must co-vary, declare it with a name.

### Lean on co-variance for paired tokens

When the same declared variable appears multiple times in a single key, references co-vary rather than fanning out independently. This is the right tool for "version N's install path and version N's log folder." They always pair, never cross-multiply.

### Default scaffold set is intentionally minimal

`Caches` and `Telemetry` are the only opt-in-free scaffolds. Anything that could remove user-visible state (cookies, history, sessions, saved logins, web storage) requires explicit opt-in: either `Scaffolds=All` with selective exclusion, or an explicit list naming the desired host-risk scaffolds.

### Prefer `<Variable>` to numbered duplication

`RegKeyBase=<RegRoot>\<Versions>\Outlook` is easier to extend than seven hand-numbered `RegKey1..7` lines. When a new version ships, append it to the `Versions=` list and every dependent key grows automatically.

---

# Troubleshooting

| Symptom | Likely Cause | Fix |
|:-|:-|:-|
| `No EntryBuilder source definitions found in: <dir>` | Source directory is empty, missing, or contains no `*.ini` files with parseable sections | Verify the source directory in **Choose source directory** or via `-1d` |
| Entry is silently absent from the output | `Skip=` is set, or the entry was skipped by validation (`No Section or LangSecRef...` / `...nothing to emit, skipping`), or a duplicate section name in an alphabetically-earlier file won | Check the run log for the corresponding message |
| Output is missing expected scaffold FileKeys | Catalog failed to load (`...catalog at <path> is empty or missing`), scaffold name misspelled (`Unknown WebView scaffold 'X'...`), or scaffold excluded by `Exclude*Scaffolds=` | Check the `Loaded N WebView scaffold(s)` / `Loaded N QtWebEngine scaffold(s)` lines and the warnings |
| Output key contains a literal `%WebViewRoot%` / `%QtWebEngineRoot%` | The matching root key wasn't declared on the entry - or the placeholder was used in a `Detect`/`DetectFile`, where it is never substituted | Declare the root key, or write the path / a `<Variable>` directly in detection keys |
| Entry gained a `Detect`/`DetectFile` you didn't write | The entry declares a variable named `Root` - detection inference is automatic | Intended? Delete your redundant hand-written detection. Not intended? Rename the variable (`DiskRoot`, `AppRoot`, ...) |
| QtWebEngine scaffold FileKeys point at a wrong/empty `Default` folder | `QtWebEngineRoot=` was pointed at a profile (e.g. `...\QtWebEngine\Default`) instead of the `QtWebEngine` folder, double-baking `\Default\` | Point `QtWebEngineRoot=` at the `QtWebEngine` folder; use `FileKeyBase=` for non-`Default` profiles |
| Output RegKey contains a literal `<Name>` | The variable wasn't declared - the registry domain emits literal text and logs `...emitted as literal` as an advisory | Declare the variable, or accept the literal if intentional |
| Output is missing a FileKey/DetectFile entirely | An undeclared `<Name>` token dropped the key (`...dropping key`), or a referenced variable had no values | Fix the typo, or declare the variable |
| `Variable 'X=' declared ... but never referenced` warning | The declaration is unused - usually a typo in a `<X>` reference somewhere | Search the entry for the intended reference and correct it |
| Output reorders keys differently from the source | WinappDebug normalization rewrites to canonical order and alphabetised numbering on every run | This is intentional; reformat the source to match if the diff bothers you |
| `Default=` was stripped from an entry | EntryBuilder never emits `Default=` by design | Move the entry to a non-EntryBuilder source if `Default=` is required |

---

# Usage Examples

The production source files these examples are drawn from live in [`Assembler/EntryBuilder`](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/EntryBuilder) in the main repo. Every example follows the same recipe: put the source snippet in a `.ini` file (the production letter file is named per example), run

```
winapp2ool -entrybuilder -1d <folder containing the source file>
```

and read the generated `entrybuilder.ini` in the current directory. The scaffold examples (8–10) additionally need the shared catalogs reachable; their full commands are shown inline.

###### Note: The outputs below are shown as EntryBuilder generates them, before WinappDebug's normalization pass alphabetizes key numbering within each family, and without the generated file header. The *key set* is exactly as shown. Long fan-outs are truncated with `; ...` comments.

## Pass-Through and Variables

### Example 1: A pass-through entry with a single-value variable

**Context**

A simple game with one log file. The entry is already plain winapp2 syntax; the only annoyance is writing the same long path twice.

**Intent**

We want the standard entry, with the path declared once. We deliberately name the variable `DiskRoot` (*not* `Root`) because we're declaring detection by hand and don't want the [reserved-`Root` inference](#the-reserved-root-variable) involved (here it would be harmless, but the naming choice should be deliberate).

**Files**

###### **Source (`A.ini`)**

```ini
[Annie's Millions *]
DiskRoot=%AppData%\PoBros\Annies Millions
Section=Games
DetectFile=<DiskRoot>
FileKey=<DiskRoot>|logfile.txt
```

**Command**

```
winapp2ool -entrybuilder -1d ..\..\Assembler\EntryBuilder
```

###### Note: `-1d` points at the directory containing the source files; `entrybuilder.ini` is written to the current directory

**Output**

```ini
[Annie's Millions *]
Section=Games
DetectFile=%AppData%\PoBros\Annies Millions
FileKey1=%AppData%\PoBros\Annies Millions|logfile.txt
```

**Explanation**

- `DiskRoot=` is not a recognised key, so it becomes a single-value variable
- `<DiskRoot>` expands in both keys; a single-value variable produces no fan-out
- The un-numbered `FileKey=` is renumbered to `FileKey1`; the single `DetectFile` stays bare
- `DiskRoot=` itself is consumed and does not appear in the output

---

### Example 2: Automatic detection from the reserved Root variable

**Context**

Most entries anchor both their detection and their cleaning targets on one path. Declaring that path as `Root=` lets EntryBuilder write the detection key for you.

**Intent**

We want a complete entry for a utility whose data lives under one folder, without writing any detection key.

**Files**

###### **Source (`A.ini`)**

```ini
[A Hat in Time *]
Root=%ProgramFiles%\Steam\steamapps\common\HatinTime\HatinTimeGame

Section=Games
FileKeyBase=<Root>\Logs|*
```

**Output**

```ini
[A Hat in Time *]
Section=Games
DetectFile=%ProgramFiles%\Steam\steamapps\common\HatinTime\HatinTimeGame
FileKey1=%ProgramFiles%\Steam\steamapps\common\HatinTime\HatinTimeGame\Logs|*
```

**Explanation**

- `Root=` is the reserved variable: EntryBuilder classifies `%ProgramFiles%\Steam\steamapps\common\HatinTime\HatinTimeGame` as a filesystem path and injects `DetectFile=<Root>`, which then expands normally
- A registry-hive value (`Root=HKCU\Software\Valve\Steam\Apps\253230`) would have produced `Detect=` instead
- A multi-valued `Root=%AppData%\Foo,%LocalAppData%\Foo` would fan the injected token into `DetectFile1` / `DetectFile2`

**Notes**

The log records the inference: `[A Hat in Time *] inferred DetectFile=<Root>`. If the entry already contained a hand-written `DetectFile=<Root>`, nothing would be added (inference never duplicates).

---

### Example 3: List-variable fan-out

**Context**

Microsoft Office entries span many historical versions, each with its own registry node. Hand-writing them means dozens of near-identical lines that must all be touched when a version is added.

**Intent**

We want one `Detect` per Office version and one `RegKey` per (version, MRU-slot) combination, declared once.

**Files**

###### **Source (`M.ini`)**

```ini
[Microsoft Outlook *]
Versions=9.0,10.0,11.0,12.0,14.0,15.0,16.0
MRUNum=1,2,3
RegRoot=HKCU\Software\Microsoft\Office
LangSecRef=3024
Detect=<RegRoot>\<Versions>\Outlook
RegKeyBase=<RegRoot>\<Versions>\Outlook\Office Finder|<MRUNum>
```

**Output**

```ini
[Microsoft Outlook *]
LangSecRef=3024
Detect1=HKCU\Software\Microsoft\Office\9.0\Outlook
Detect2=HKCU\Software\Microsoft\Office\10.0\Outlook
; ... Detect3 – Detect6 ...
Detect7=HKCU\Software\Microsoft\Office\16.0\Outlook
RegKey1=HKCU\Software\Microsoft\Office\9.0\Outlook\Office Finder|1
RegKey2=HKCU\Software\Microsoft\Office\9.0\Outlook\Office Finder|2
RegKey3=HKCU\Software\Microsoft\Office\9.0\Outlook\Office Finder|3
RegKey4=HKCU\Software\Microsoft\Office\10.0\Outlook\Office Finder|1
; ... RegKey5 – RegKey20, version-major / MRU-minor ...
RegKey21=HKCU\Software\Microsoft\Office\16.0\Outlook\Office Finder|3
```

**Explanation**

- `Detect` references one 7-value variable → 7 keys, numbered `Detect1`–`Detect7`
- `RegKeyBase` references two variables → 7 × 3 = 21 keys; the leftmost reference (`<Versions>`) varies slowest
- `<RegRoot>` is single-valued, so it abbreviates without multiplying
- Adding Office vNext is a one-token edit to `Versions=`: both key families grow automatically

**Notes**

WinappDebug's alphabetization is number-aware: digit runs compare by numeric value, so `..\10.0` sorts after `..\9.0` and the ascending version order shown here survives normalization unchanged. The pre-normalization caveat above bites only when declaration order differs from sort order - [Example 4](#example-4-co-variance-of-repeated-variables)'s `Stable,Beta` keys, for instance, come back renumbered with `Beta` first.

---

### Example 4: Co-variance of repeated variables

**Context**

Some applications embed the same version number twice in one path: once as a folder name, once inside a file name. Those two positions must always agree - version 1.0's folder never contains version 2.0's log.

**Intent**

We want one cache key per (version, channel) combination, and one log key per version, where the version in the file name always matches the version in the folder.

**Files**

###### **Source (`H.ini`)**

```ini
[Hazel Vault *]
Versions=1.0,2.0
Channels=Stable,Beta
Root=%AppData%\Hazel Vault

LangSecRef=3021
FileKeyBase=<Root>\<Versions>\<Channels>\Cache|*|RECURSE
FileKeyBase=<Root>\<Versions>|app-<Versions>.log
```

**Output**

```ini
[Hazel Vault *]
LangSecRef=3021
DetectFile=%AppData%\Hazel Vault
FileKey1=%AppData%\Hazel Vault\1.0\Stable\Cache|*|RECURSE
FileKey2=%AppData%\Hazel Vault\1.0\Beta\Cache|*|RECURSE
FileKey3=%AppData%\Hazel Vault\2.0\Stable\Cache|*|RECURSE
FileKey4=%AppData%\Hazel Vault\2.0\Beta\Cache|*|RECURSE
FileKey5=%AppData%\Hazel Vault\1.0|app-1.0.log
FileKey6=%AppData%\Hazel Vault\2.0|app-2.0.log
```

**Explanation**

- The first `FileKeyBase` references two *different* multi-value variables, so they cross-multiply: 2 × 2 = 4 keys (`<Root>` is single-valued and just abbreviates)
- The second `FileKeyBase` references `<Versions>` *twice*, so the occurrences co-vary: 2 keys, not 4. Both occurrences bind to the same value in each output key - there is no `\1.0|app-2.0.log`
- Co-variance is automatic for repeated *declared* variables. A repeated identical *inline* list also co-varies, but logs an advisory because an anonymous list can't state that intent (see [Inline Lists](#inline-lists))

---

### Example 5: Inline lists

**Context**

7-Zip keeps an archive-history value under two sibling registry keys. Declaring a named variable for a two-item list used once is ceremony without benefit.

**Intent**

We want one `RegKey` per sibling key, from a single line, and a free detection anchor while we're at it.

**Files**

###### **Source (`#.ini`)**

```ini
[7-Zip *]
LangSecRef=3024
Root=HKCU\Software\7-Zip
RegKeyBase=<Root>\<Compression,Extraction>|ArcHistory
```

**Output**

```ini
[7-Zip *]
LangSecRef=3024
Detect=HKCU\Software\7-Zip
RegKey1=HKCU\Software\7-Zip\Compression|ArcHistory
RegKey2=HKCU\Software\7-Zip\Extraction|ArcHistory
```

**Explanation**

- `<Compression,Extraction>` contains a comma, so it is an inline list: an anonymous two-value axis needing no declaration
- `Root=` is a registry-hive value here, so the inferred detection is a `Detect`, not a `DetectFile`
- Repeating the *same* inline list twice in one key makes the occurrences co-vary and logs an advisory. Declare a named variable when co-variance (or independence) needs to be explicit

---

### Example 6: Nested references

**Context**

ArcGIS ArcMap's registry layout changed across versions: the oldest releases used a bare `ArcMap` key, later ones `Desktop<version>\ArcMap`. The set of key names is itself version-derived.

**Intent**

We want the version list declared once, the derived key-name list built from it, and every dependent key fanned across the result.

**Files**

###### **Source (`A.ini`)**

```ini
[ArcGIS ArcMap *]
Versions=10.0,10.1
ArcMapKey=ArcMap,Desktop<Versions>\ArcMap

LangSecRef=3021
Detect=HKCU\Software\ESRI\<ArcMapKey>
RegKeyBase=HKCU\Software\ESRI\<ArcMapKey>\Recent File List
```

**Output**

```ini
[ArcGIS ArcMap *]
LangSecRef=3021
Detect1=HKCU\Software\ESRI\ArcMap
Detect2=HKCU\Software\ESRI\Desktop10.0\ArcMap
Detect3=HKCU\Software\ESRI\Desktop10.1\ArcMap
RegKey1=HKCU\Software\ESRI\ArcMap\Recent File List
RegKey2=HKCU\Software\ESRI\Desktop10.0\ArcMap\Recent File List
RegKey3=HKCU\Software\ESRI\Desktop10.1\ArcMap\Recent File List
```

**Explanation**

- `ArcMapKey`'s value list contains `<Versions>`, so it resolves first: the 2-value `Versions` expands the second list item, yielding the 3-value list `(ArcMap, Desktop10.0\ArcMap, Desktop10.1\ArcMap)`
- `Detect` and `RegKeyBase` then fan across the fully-literal 3-value list
- A cycle (`A` referencing `B` referencing `A`) would leave both variables as authored, with a warning naming the participants

---

### Example 7: Literal angle brackets in registry keys

**Context**

Unlike file paths, registry key and value names may legitimately contain `<` and `>`. Some applications store values with names like `<default>`.

**Intent**

We want to target a registry value literally named `<default>` and understand why EntryBuilder lets this through instead of treating it as a broken variable reference.

**Files**

###### **Source (`H.ini`)**

```ini
[Hazel's Silly Little System Utility *]
Root=HKCU\Software\Hazel\HSLSU

LangSecRef=3024
RegKeyBase=<Root>\Recent|<default>
```

**Output**

```ini
[Hazel's Silly Little System Utility *]
LangSecRef=3024
Detect=HKCU\Software\Hazel\HSLSU
RegKey1=HKCU\Software\Hazel\HSLSU\Recent|<default>
```

**Explanation**

- `<Root>` matches a declared variable and expands; `<default>` matches nothing
- Because `RegKey` is a registry-domain key, the undeclared token is emitted **as literal text**, with a log-only advisory: `Undeclared variable <default> in registry-domain key [Hazel's Silly Little System Utility *].RegKey; emitted as literal`
- The same token in a `FileKeyBase` would instead **drop the key** with a menu-visible warning, since `<` / `>` can never appear in a Windows path

**Notes**

This is why the unreferenced-variable backstop matters: if you declare `Versions=` but typo every reference as `<Versoin>`, the registry domain would silently emit the literal. The `Variable 'Versions=' declared ... but never referenced` warning is what surfaces the mistake.

---

## Scaffold Families

### Example 8: Default WebView scaffolds

**Context**

Discord is one of dozens of Electron/WebView apps whose embedded Chromium leaves the same cache and telemetry residue. The shared catalog exists so none of those patterns are written per-app.

**Intent**

We want baseline cache and telemetry cleaning for Discord's Chromium data folder, from a four-line definition.

**Files**

###### **Source (`D.ini`)**

```ini
[Discord *]
WebViewRoot=%AppData%\discord
LangSecRef=3023
DetectFile=%AppData%\discord
```

**Command**

```
winapp2ool -entrybuilder -1d ..\..\Assembler\EntryBuilder -3d ..\..\Assembler\Scaffolds
```

###### Note: `-3d` points file slot 3 (the WebView catalog, `webview.ini`) at the shared scaffolds folder

**Output**

```ini
[Discord *]
LangSecRef=3023
DetectFile=%AppData%\discord
FileKey1=%AppData%\discord|*-journal|RECURSE
FileKey2=%AppData%\discord|Module Info Cache
FileKey3=%AppData%\discord\Default|*.ldb;CURRENT;LOCK;MANIFEST-*;ServerCertificate
FileKey4=%AppData%\discord\Default\*Cache*|*|REMOVESELF
; ... 16 further Caches FileKeys ...
FileKey21=%AppData%\discord|*.pma;LOG;LOG.old|RECURSE
FileKey22=%AppData%\discord|*_shutdown_ms.txt;*.log;Breadcrumbs;BrowsingTopics*;Last Browser;Last Version
; ... 15 further Telemetry FileKeys, 37 scaffold FileKeys in total from the default set
```

**Explanation**

- No `WebViewScaffolds=` key is present, so the default set applies: `Caches` + `Telemetry`
- Every `FileKeyBase=` template in each selected catalog scaffold is copied in with `%WebViewRoot%` replaced by `%AppData%\discord`
- The exact keys (and their count, 37 here) track the catalog: when the catalog gains a pattern, every entry that selects that scaffold gains the key on the next build, with no source edits

**Notes**

A second `WebViewRoot=` line would double the scaffold output: one full set per root.

---

### Example 9: Everything except All with exclusions

**Context**

Adobe Photoshop's embedded WebView is used for plugin UI; nearly all of its browsing-type residue is safe to remove, but wiping web storage and saved logins would break plugin state users care about.

**Intent**

We want every catalog scaffold **except** `WebStorage` and `LoginData`, plus one hand-written FileKey for Photoshop's own logs.

**Files**

###### **Source (`A.ini`)**

```ini
[Adobe Photoshop *]
WebViewRoot=%AppData%\Adobe\UXP\PluginsStorage\PHSP\26\Shared\EBWebView
WebViewScaffolds=All
ExcludeWebViewScaffolds=WebStorage,LoginData
LangSecRef=3023

DetectFile=%AppData%\Adobe\Adobe Photoshop *
FileKey=%AppData%\Adobe\Adobe Photoshop *\Logs|*
```

**Command**

```
winapp2ool -entrybuilder -1d ..\..\Assembler\EntryBuilder -3d ..\..\Assembler\Scaffolds
```

**Output**

```ini
[Adobe Photoshop *]
LangSecRef=3023
DetectFile=%AppData%\Adobe\Adobe Photoshop *
FileKey1=%AppData%\Adobe\UXP\PluginsStorage\PHSP\26\Shared\EBWebView\Default|*Web Data
FileKey2=%AppData%\Adobe\UXP\PluginsStorage\PHSP\26\Shared\EBWebView\Default\AutoFill*|*|REMOVESELF
FileKey3=%AppData%\Adobe\UXP\PluginsStorage\PHSP\26\Shared\EBWebView\AutoFill*|*|REMOVESELF
FileKey4=%AppData%\Adobe\UXP\PluginsStorage\PHSP\26\Shared\EBWebView\MEIPreload|*|REMOVESELF
; ... FileKey5 – FileKey77: the remaining safe scaffold sets (Caches, Security, Telemetry, ...) ...
FileKey78=%AppData%\Adobe\UXP\PluginsStorage\PHSP\26\Shared\EBWebView\Default\Network|Cookies*;Device Bound Sessions*
FileKey79=%AppData%\Adobe\UXP\PluginsStorage\PHSP\26\Shared\EBWebView\Default|History*;Network Action Predictor*;Top Sites*;shortcuts*;Visited Links*
FileKey80=%AppData%\Adobe\UXP\PluginsStorage\PHSP\26\Shared\EBWebView\Default\Extension State|*|REMOVESELF
FileKey81=%AppData%\Adobe\UXP\PluginsStorage\PHSP\26\Shared\EBWebView\Default\Sessions|*|REMOVESELF
FileKey82=%AppData%\Adobe\Adobe Photoshop *\Logs|*
```

**Explanation**

- `All` expands to every scaffold in the catalog (20 at the time of writing) and the log records it: `WebViewScaffolds=All in [Adobe Photoshop *]; expanded to 20 scaffold(s) from catalog`
- The two exclusions subtract from that expansion, leaving 18 scaffolds and 81 scaffold FileKeys. That includes the host-risk sets `All` deliberately opts into: `WebCookies` (FileKey78), `WebHistory` (FileKey79), and `WebSession` (FileKey80–81)
- `All` + `Exclude...` is the sanctioned idiom and does not warn; an explicit non-`All` list combined with exclusions does
- The hand-written `FileKey=` is appended after all scaffold keys (FileKey82) and participates in the same renumbering

---

### Example 10: QtWebEngine and a non-Default profile

**Context**

calibre embeds QtWebEngine. Its main profile residue is covered by the QtWebEngine catalog, but calibre also creates `OffTheRecord` and `viewer-lookup` profiles the catalog's `Default`-profile templates can't reach.

**Intent**

We want the default QtWebEngine scaffolds plus two hand-written keys reaching the non-`Default` profiles, reusing the root declaration for all of it.

**Files**

###### **Source (`C.ini`)**

```ini
[Calibre *]
QtWebEngineRoot=%LocalAppData%\calibre-ebook.com\calibre\QtWebEngine
LangSecRef=3023
DetectFile=%LocalAppData%\calibre-ebook.com
FileKeyBase=%QtWebEngineRoot%\OffTheRecord\GPUCache|*|RECURSE
FileKeyBase=%QtWebEngineRoot%\viewer-lookup\blob_storage|*|REMOVESELF
```

**Command**

```
winapp2ool -entrybuilder -1d ..\..\Assembler\EntryBuilder -4d ..\..\Assembler\Scaffolds
```

###### Note: `-4d` points file slot 4 (the QtWebEngine catalog, `qtwebengine.ini`) at the shared scaffolds folder

**Output**

```ini
[Calibre *]
LangSecRef=3023
DetectFile=%LocalAppData%\calibre-ebook.com
FileKey1=%LocalAppData%\calibre-ebook.com\calibre\QtWebEngine\Default\*Cache|*|RECURSE
FileKey2=%LocalAppData%\calibre-ebook.com\calibre\QtWebEngine\Default\GPUCache|*|REMOVESELF
FileKey3=%LocalAppData%\calibre-ebook.com\calibre\QtWebEngine\Default\blob_storage|*|REMOVESELF
FileKey4=%LocalAppData%\calibre-ebook.com\calibre\QtWebEngine\Default\File System|*|REMOVESELF
FileKey5=%LocalAppData%\calibre-ebook.com\calibre\QtWebEngine\Default\Platform Notifications|*|REMOVESELF
FileKey6=%LocalAppData%\calibre-ebook.com\calibre\QtWebEngine\Default\Service Worker|*|REMOVESELF
FileKey7=%LocalAppData%\calibre-ebook.com\calibre\QtWebEngine\Default|*-journal;*.log;*.old;LOG;LOG.old;Network Persistent State
FileKey8=%LocalAppData%\calibre-ebook.com\calibre\QtWebEngine\Default\VideoDecodeStats|*|REMOVESELF
FileKey9=%LocalAppData%\calibre-ebook.com\calibre\QtWebEngine\OffTheRecord\GPUCache|*|RECURSE
FileKey10=%LocalAppData%\calibre-ebook.com\calibre\QtWebEngine\viewer-lookup\blob_storage|*|REMOVESELF
```

**Explanation**

- `QtWebEngineRoot=` points at the `QtWebEngine` **folder**: the catalog templates bake the `\Default\` profile segment in themselves (FileKey1–8: the default `Caches` + `Telemetry` sets)
- The two `FileKeyBase=` lines reach the non-`Default` profiles by reusing `%QtWebEngineRoot%`, the escape hatch for anything outside the catalog's `Default` assumption
- Pointing the root at `..\QtWebEngine\Default` instead would double-bake the profile segment and clean nothing
- Host-risk scaffolds (`WebCookies`, `WebHistory`, `WebSession`, `WebStorage`) would require `QtWebEngineScaffolds=All` or an explicit list, exactly as in the WebView family

---

## Housekeeping

### Example 11: Skipping a defunct entry

**Context**

An application is discontinued, but deleting its definition would lose the research embedded in it should the app return or a fork appear.

**Intent**

We want the entry omitted from the output while keeping its definition in source.

**Files**

###### **Source (`D.ini`)**

```ini
[Discontinued App *]
Skip=
WebViewRoot=%AppData%\DiscontinuedApp
LangSecRef=3023
DetectFile=%AppData%\DiscontinuedApp
```

**Output**

The entry does not appear in `entrybuilder.ini`. Re-enabling it later is a one-line deletion.

**Explanation**

- `Skip=` omits the entry regardless of its value: even `Skip=False` skips (see [Skip](#skip))
- Skipped entries are counted in the run statistics, so a `Skip=` never disappears unaccounted
