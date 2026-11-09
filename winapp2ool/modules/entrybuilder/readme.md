# EntryBuilder

**EntryBuilder** is a winapp2ool build-chain module for winapp2.ini maintainers that generates winapp2.ini entries from a shorthand DSL. You write one ini section per application: anything that is already valid winapp2 syntax passes through unchanged, and a small set of winapp2ool-private keys expand into standard winapp2 form at generation time. Think of it as a preprocessor for winapp2.ini: declare the repetitive parts once and let the tool write them out. If you only use winapp2.ini with a cleaner, you don't need EntryBuilder.

EntryBuilder is a devops module used to maintain the base entries of the winapp2.ini distribution. Every non-browser, non-UWP entry lives in its source files. Source files at `Assembler\EntryBuilder\{letter}.ini` are combined and processed alongside the shared scaffold catalogs in `Assembler\Scaffolds` to produce either a single `entrybuilder.ini` (the default) or, with `-split`, the 27 per-letter artifact files the build pipeline commits at `Assembler\Entries\` and merges into the final database.

### What does EntryBuilder do?

EntryBuilder reads every `*.ini` file in its source directory, combines them in-memory, parses each section into an entry definition, expands the shorthand, normalizes the result via WinappDebug, and writes the output to `entrybuilder.ini`. The two main shorthand categories are:

- **Scaffold families**: entries that declare a `WebViewRoot=` (embedded WebView2/EBWebView data folder), `QtWebEngineRoot=` (embedded QtWebEngine data folder), or `ElectronRoot=` (an Electron app's `userData` folder) receive baseline cleaning FileKeys drawn from a shared catalog, one set per declared root.
- **List variables**: entries can declare comma-separated lists (`Versions=10.0,11.0,12.0`) and reference them from any key with `<Versions>` tokens; each referencing key fans out into one output key per combination.

### Why EntryBuilder?

- **Shorthand DSL**: Replaces 20-line repetitive entries with a few list-variable declarations and templated keys. 
- **Shared scaffold catalogs**: One catalog of Chromium cleaning patterns covers every embedded WebView2 host, without per-app duplication. A second does the same for QtWebEngine hosts, and a third for Electron apps.
- **Pass-through compatible**: Any standard winapp2 key passes through unchanged. You can paste a raw entry in and incrementally enrich it with shorthand, or leave it untouched.
- **Consistent DSL surface**: The same `<Name>` variable syntax and scaffold key names work across winapp2ool's builder modules.

---

# Table of Contents

**Using EntryBuilder**

1. [Requirements](#requirements)
2. [Quick Start](#quick-start)
3. [Common Uses](#common-uses)
4. [Menu Options](#menu-options)
5. [Command-Line Arguments](#command-line-arguments)
   - [File Selection](#file-selection)
   - [Toggles](#toggles)
6. [How Generation Works](#how-generation-works)
7. [Source File Format](#source-file-format)
8. [Pass-Through Keys](#pass-through-keys)
9. [Shorthand Keys](#shorthand-keys)
   - [WebViewRoot](#webviewroot)
   - [WebViewScaffolds / ExcludeWebViewScaffolds](#webviewscaffolds--excludewebviewscaffolds)
   - [QtWebEngineRoot / QtWebEngineCacheRoot / QtWebEngineScaffolds](#qtwebengineroot--qtwebenginecacheroot--qtwebenginescaffolds)
   - [ElectronRoot / ElectronUpdaterRoot / ElectronScaffolds](#electronroot--electronupdaterroot--electronscaffolds)
   - [FileKeyBase / RegKeyBase / ExcludeKeyBase](#filekeybase--regkeybase--excludekeybase)
   - [Skip](#skip)
10. [List Variables and Token Expansion](#list-variables-and-token-expansion)
    - [Declaration](#declaration)
    - [Reference](#reference)
    - [Inline Lists](#inline-lists)
    - [Indexed References](#indexed-references)
    - [Cartesian Fan-out and Co-variance](#cartesian-fan-out-and-co-variance)
    - [Nested References](#nested-references)
    - [The Reserved Root Variable](#the-reserved-root-variable)
    - [Undeclared Tokens](#undeclared-tokens)
    - [Typo Backstops](#typo-backstops)
11. [Validation and Warnings](#validation-and-warnings)
12. [Run Output](#run-output)
13. [Tips & Best Practices](#tips--best-practices)
14. [Troubleshooting](#troubleshooting)

**Examples**

15. **[Examples](EXAMPLES.md)**: thirteen worked examples with captured input and output

---

# Requirements

- A source directory containing one or more `*.ini` files with at least one parseable section
- The shared scaffold directory (typically `Assembler\Scaffolds`) if any entry declares a scaffold root. Every `*.ini` in it is read, and each catalog's engine family comes from its section headers rather than its filename: `webview.ini` (`[WebViewScaffold: ...]`) serves `WebViewRoot=`, `qtwebengine.ini` (`[QtWebEngineScaffold: ...]`) serves `QtWebEngineRoot=` / `QtWebEngineCacheRoot=`, and `electron.ini` (`[ElectronScaffold: ...]`) serves `ElectronRoot=` / `ElectronUpdaterRoot=`

If the source directory is empty, missing, or contains no parseable sections, EntryBuilder reports `No EntryBuilder source definitions found in: <directory>` and writes no output. If the scaffold directory or a catalog within it is missing or empty, generation continues with zero scaffold FileKeys emitted for the affected families and a warning logged. Entries with no corresponding root key are unaffected.

---

# Quick Start

### Common Workflow

1. Place per-letter source files (`A.ini`, `B.ini`, ...) in the source directory
2. Open EntryBuilder from the **Entry Lab** main menu, or invoke `winapp2ool -entrybuilder` from the command line
3. Run. EntryBuilder combines the files, expands the shorthand, and writes the result to `entrybuilder.ini`

The default source directory, save target, and scaffold directory are all the current directory; in normal use `entrybuilder.ini` lives next to the winapp2ool executable, the source directory points at `..\..\Assembler\EntryBuilder\`, and the scaffold directory at `..\..\Assembler\Scaffolds\`.

---

# Common Uses

| To | Run | Example |
|:-|:-|:-|
| Generate entries from the source files in the current directory | `winapp2ool -entrybuilder` | |
| Generate from the project's source files, saving `entrybuilder.ini` in the current directory | `winapp2ool -entrybuilder -1d ..\..\Assembler\EntryBuilder` | [A pass-through entry with a single-value variable](EXAMPLES.md#example-1-a-pass-through-entry-with-a-single-value-variable) |
| Generate with the shared scaffold catalogs in silent mode | `winapp2ool -entrybuilder -1d ..\..\Assembler\EntryBuilder -3d ..\..\Assembler\Scaffolds -s` | [Default WebView scaffolds](EXAMPLES.md#example-8-default-webview-scaffolds) |
| Save the output under another file name | `winapp2ool -entrybuilder -2f test-output.ini` | |
| Write the per-letter artifact files into `Assembler\Entries` | `winapp2ool -entrybuilder -split -2d ..\..\Assembler\Entries` | |

In the menu, the same choices are the source, save target and scaffolds directory choosers and the output splitting toggle, set before **Run**; see [Menu Options](#menu-options). `-s` is one of winapp2ool's [global args](../../Readme.md#global-args).

---

# Menu Options

| Option | Effect | Notes |
|:-|:-|:-|
| Run (default) | Generate entries from the configured source directory | Does nothing if no parseable sections found |
| Choose source directory | Select the directory containing per-letter source files | Only the directory is used; the file name is ignored |
| Choose save target | Select where to save the generated entries | Default: `entrybuilder.ini` in current directory |
| Choose scaffolds directory | Select the folder holding the shared scaffold catalogs | Default: current directory |
| Reset Settings | Restore all settings to their defaults | Only shown when settings have been changed |

---

# Command-Line Arguments

EntryBuilder is invoked as `winapp2ool -entrybuilder`. Combine with the global `-s` flag for silent (non-interactive) execution in the build pipeline.

### File Selection

Each file slot has a corresponding index for use with the `-Nd` (directory) and `-Nf` (file name) argument pattern:

| Index | File | Default |
|:-|:-|:-|
| 1 | Source directory | Current directory (directory only; file name is ignored) |
| 2 | Save target | `entrybuilder.ini` in current directory |
| 3 | Shared scaffold directory (only `.Dir` is read, so `-3f` has no effect) | Current directory |

| Arg | Effect |
|:-|:-|
| `-Nd path` | Set directory for file slot N |
| `-Nf name` | Set file name for file slot N |
| `-Nf subdir\name` | Set file name within a subdirectory of its path |

### Toggles

| Arg | Effect |
|:-|:-|
| `-split` | Write per-letter artifact files (`#.ini`, `A.ini` ... `Z.ini`) into the save target's directory instead of a single output file. Entries are bucketed by the first character of their name (non-letters go to `#.ini`); all 27 files are always written, each with a do-not-edit header |

---

# How Generation Works

Each run proceeds through the same fixed steps:

1. **Combine**: every `*.ini` in the source directory is loaded in alphabetical filename order and concatenated in-memory. 
2. **Parse**: each section becomes one entry definition. Recognised winapp2 and shorthand keys are claimed by name; *any other key* is treated as a list-variable declaration.
3. **Resolve variables**: variable values that reference other variables (`<Other>` tokens inside a declaration) are resolved so every value list is fully literal. Cyclic references are warned about and not expanded.
4. **Infer detection**: if the entry declares a variable named `Root`, a `Detect=<Root>` or `DetectFile=<Root>` is generated automatically (see [The Reserved Root Variable](#the-reserved-root-variable)).
5. **Validate**: structurally invalid entries are skipped with warnings; questionable ones are emitted with warnings (see [Validation and Warnings](#validation-and-warnings)).
6. **Generate**: keys are emitted per entry in canonical winapp2 order. Each content key goes through, in order: scaffold expansion (catalog templates copied in for each selected scaffold and declared root), **root substitution** (`<WebViewRoot>`, `<QtWebEngineRoot>`, `<QtWebEngineCacheRoot>`, `<ElectronRoot>` and `<ElectronUpdaterRoot>`, in any case, replaced with each declared root), then **token expansion** (`<Name>` fan-out). Numbered key families are renumbered from 1.
7. **Normalize**: the whole in-memory file is passed through WinappDebug with all repairs and optimizations enabled.
8. **Save**: the output file is written with a generated comment header (see [Run Output](#run-output)).

Because root substitution (step 6) happens before token expansion, `<Name>` tokens inside a `WebViewRoot=` / `QtWebEngineRoot=` value survive substitution and are expanded like any other token in the resulting key. Root placeholders share the `<Name>` spelling but aren't variables: they are substituted as text first, so a `<Root>` inside a root value and a `<Root>` in the template are one token and co-vary.

Not every key family receives every processing step:

| Key family | Root placeholders | Expansion domain |
|:-|:-|:-|
| `Detect` | No  | Registry |
| `DetectFile` | No  | Filesystem |
| `FileKey` / `FileKeyBase` | Yes | Filesystem |
| `RegKey` / `RegKeyBase` |  Yes | Registry |
| `ExcludeKey` / `ExcludeKeyBase` |  Yes | By flag: `REG` → Registry, otherwise Filesystem |
| Scaffold catalog templates | Yes (their purpose) | Filesystem |

The expansion domain matters when a `<Name>` token doesn't match any declared variable, see [Undeclared Tokens](#undeclared-tokens). Note that root placeholders are **not** substituted in `Detect` / `DetectFile`, where they reach token expansion as undeclared tokens: a `DetectFile` holding one is dropped, and a `Detect` keeps it literally. Write the path (or a `<Variable>`) directly in detection keys.

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
| `DetectOS` | Kernel version filter, e.g. `6.1\|10.0`; emitted verbatim, never subject to variable expansion. Counts as detection for the always-on warning |
| `Warning` | Emitted verbatim; never subject to variable expansion |
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

- The root is the folder that holds the profiles, not a profile. Profile-scoped templates use `<WebViewRoot>\*\`, so `Default` and named profiles such as `WV2Profile_*` are all covered.
- Multiple `WebViewRoot=` keys are permitted; each scaffold FileKey is generated once per root.
- The literal string `<WebViewRoot>` can be used in any `FileKey` / `RegKey` / `ExcludeKey` value (`Base=` or plain form); it is substituted against each declared root at generation time, fanning the key out per root.
- An entry with no roots and no FileKey/RegKey content is skipped with a warning.

### The retired `%WebViewRoot%` spelling

Earlier versions spelled the root placeholders `%WebViewRoot%`, `%QtWebEngineRoot%`, `%QtWebEngineCacheRoot%`, `%ElectronRoot%` and `%ElectronUpdaterRoot%`. A key still written that way, including one with the closing `%` missing (`%ElectronRoot\logs|*`), is dropped with a warning naming the new spelling, and the run fails with exit code 1:

```
`%QtWebEngineCacheRoot%` in [Test App *].FileKeyBase is now spelled `<QtWebEngineCacheRoot>`; key dropped
`%ElectronRoot` in [Test App *].FileKey looks like the retired `%ElectronRoot%` token, which is now spelled `<ElectronRoot>`; key dropped
```

`%...%` is winapp2.ini's environment-variable syntax, so a stale token would otherwise reach the output looking like a real path. Every key is checked except `Warning=` prose. UWPBuilder's retired `%Package%` means nothing here and is left alone. The shared catalogs in `Assembler\Scaffolds` are checked as they load, and a catalog template holding a retired token is dropped and fails the run the same way.

## WebViewScaffolds / ExcludeWebViewScaffolds

`WebViewScaffolds=` is a comma-separated list of scaffold names from the shared catalog to apply to this entry.

| Form | Behavior |
|:-|:-|
| Key absent | Default scaffold set is used: `Caches`, `Telemetry` |
| Key present, value empty | No scaffold FileKeys are generated |
| `WebViewScaffolds=Caches,Telemetry,DRMData` | Exactly the named scaffolds are applied |
| `WebViewScaffolds=All` | Expands to every scaffold in the catalog except the legacy tier, **including host-risk categories** (cookies, history, sessions, web storage, login data) |
| `WebViewScaffolds=All,LegacyTelemetry` | `All`, plus the named legacy scaffold |

`ExcludeWebViewScaffolds=` is a comma-separated list of scaffold names to subtract from the active selection. The natural idiom for "everything except X":

###### Example

```ini
WebViewScaffolds=All
ExcludeWebViewScaffolds=WebStorage,LoginData
```

Combining a non-`All` explicit `WebViewScaffolds=` list with `ExcludeWebViewScaffolds=` warns (`Both WebViewScaffolds and ExcludeWebViewScaffolds set in [Entry]; applying exclusions to explicit list`), since the same effect is usually more clearly expressed by listing only the desired scaffolds. Unknown scaffold names are dropped with a warning, whether or not they sit beside `All`.

Scaffold names track the catalog, not this document. At the time of writing the WebView catalog defines: `Autofill`, `Autoplay`, `BookmarkFavicons`, `Caches`, `DefaultApps`, `DownloadHistory`, `DRMData`, `PrivacySandbox`, `LoginData`, `Security`, `Shopping`, `Sync Data`, `Telemetry`, `WebCookies`, `WebHistory`, `WebSession`, `WebStorage`, and in the legacy tier `LegacyCaches`, `LegacyDownloadHistory`, `LegacyDRMData`, `LegacyExtensionCookies`, `LegacyProgressiveWebApps`, `LegacyStorageQuota`, `LegacyTelemetry`, `LegacyWebHistory`. Only `Caches` and `Telemetry` are default-on; scaffolds that can remove user-visible state (`WebCookies`, `WebHistory`, `WebSession`, `WebStorage`, `LoginData`) always require explicit opt-in.

### The legacy tier

A catalog section carrying `Tier=Legacy` is withheld from `All`. It is emitted only when an entry names it, alone or beside `All`. Legacy scaffolds hold patterns that no current host writes but an older or unusual one may, and their names start with `Legacy` so the tier shows in a selection line. Naming a non-legacy scaffold beside `All` is redundant and warns; naming a legacy one is the only reason to list anything beside it.

## QtWebEngineRoot / QtWebEngineCacheRoot / QtWebEngineScaffolds

QtWebEngine is a second independent scaffold family for apps that embed Qt's WebEngine rather than WebView2. It behaves **exactly like the WebView family above**, with these substitutions:

| WebView family | QtWebEngine family |
|:-|:-|
| `WebViewRoot=` | `QtWebEngineRoot=` |
| `WebViewScaffolds=` | `QtWebEngineScaffolds=` |
| `ExcludeWebViewScaffolds=` | `ExcludeQtWebEngineScaffolds=` |
| `<WebViewRoot>` placeholder | `<QtWebEngineRoot>` placeholder |
| `webview.ini` catalog | `qtwebengine.ini` catalog |

The `All` value for scaffolds, the legacy tier, and exclusion format work identically, and a single entry may declare both families. The QtWebEngine catalog is smaller: `Caches`, `Favicons`, `PrivacySandbox`, `Security`, `StorageQuota`, `Telemetry`, `VisitedLinks`, `WebCookies`, `WebHistory`, `WebSession`, `WebStorage`, and in the legacy tier `LegacyTelemetry`, `LegacyWebStorage`.

The default set is: `Caches`, `StorageQuota`, `Telemetry` and `VisitedLinks`.

- `QtWebEngineRoot=` names one profile directory, profile segment included. Use`QtWebEngineRoot=%LocalAppData%\AppName\QtWebEngine\Default`, not the `QtWebEngine` folder above it. 


```ini
QtRoot=%LocalAppData%\calibre-ebook.com\calibre\QtWebEngine
QtWebEngineRoot=<QtRoot>\<Default,OffTheRecord,viewer-lookup>
```

### QtWebEngineCacheRoot

`QtWebEngineCacheRoot=` names the profile's HTTP cache directory, bound to `<QtWebEngineCacheRoot>`. Qt keeps the cache under its CacheLocation rather than inside the profile directory, so `QtWebEngineRoot=` can't reach it:

```ini
QtWebEngineRoot=%LocalAppData%\VideoKeeper\QtWebEngine\Default
QtWebEngineCacheRoot=%LocalAppData%\VideoKeeper\cache\QtWebEngine\Default
```

It works like `ElectronUpdaterRoot=` below: either root alone opts the entry into the family, and a template whose placeholder has no declared root is dropped.

## ElectronRoot / ElectronUpdaterRoot / ElectronScaffolds

Electron is a third independent scaffold family, for apps that bundle their own Chromium via Electron rather than embedding a browser view. The grammar is identical WebView and QTWebEngine

| WebView family | Electron family |
|:-|:-|
| `WebViewRoot=` | `ElectronRoot=` |
| `WebViewScaffolds=` | `ElectronScaffolds=` |
| `ExcludeWebViewScaffolds=` | `ExcludeElectronScaffolds=` |
| `<WebViewRoot>` placeholder | `<ElectronRoot>` placeholder |
| `webview.ini` catalog | `electron.ini` catalog |

The default set is: `AppLogs`, `Caches`, `StorageQuota`, `Telemetry` and `UpdaterCache`. The full catalog adds `MediaDRM`, `PrivacySandbox`, `Security`, `TempFiles` (opt-in, safe) and `WebCookies`, `WebStorage` (opt-in, host-risk).

`ElectronRoot=` declares the app's `userData` folder, which for Electron is the Chromium profile directory. Typically `%AppData%\<ProductName>`.

```ini
[Signal Messenger *]
LangSecRef=3022
DetectFile=%AppData%\Signal
ElectronRoot=%AppData%\Signal
```

### ElectronUpdaterRoot

Like QtWebEngine, Electron has a second root key: `ElectronUpdaterRoot=` names the electron-updater download cache, a sibling of `userData` under `%LocalAppData%` holding `installer.exe`, `package.7z` and a `pending\` folder. It must be declared separately.

### Multiple roots

Root keys are repeatable:

```ini
ElectronRoot1=%AppData%\Notion
ElectronRoot2=%AppData%\Notion\partitions\*
```

## FileKeyBase / RegKeyBase / ExcludeKeyBase

Non-scaffold templates, applied **in addition to** the catalog scaffolds. Each behaves exactly like its plain-form counterpart (`FileKey=` etc. see [Pass-Through Keys](#pass-through-keys)): root substitution, then token expansion, then renumbering from 1. Use these for app-specific cleaning targets outside the catalog scaffolds: application logs, crash dumps, version-numbered paths, etc.

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

## Indexed References

`<Name[N]>` picks the Nth value of a declared variable, indexed from 1, and doesn't fan out:

```ini
Versions=9.0,10.0,11.0
DetectFile=%ProgramFiles%\App <Versions[3]>
```

gives `DetectFile=%ProgramFiles%\App 11.0`. The index counts the variable's values after nested references and inline lists in its declaration have been flattened. An index of 0 or past the end of the list drops the key with a warning, in registry keys too, since the name is declared and the token can't be a literal. An indexed name that isn't declared is an [undeclared token](#undeclared-tokens) like any other.

`<Versions>` and `<Versions[2]>` in one key are two separate variables as far as fan-out goes, so they cross-multiply rather than co-vary. A comma still makes an inline list, so `<a[1],b>` is the two literal values `a[1]` and `b`.

## Cartesian Fan-out and Co-variance

A key that references N variables of sizes A, B, C, ... produces A × B × C × ... output keys, one per combination. The leftmost reference varies slowest (outermost loop), so the `RegKeyBase` example under [Reference](#reference) produces 7 × 3 = 21 RegKey lines ordered as `(Versions[0], MRUNum[0])`, `(Versions[0], MRUNum[1])`, `(Versions[0], MRUNum[2])`, `(Versions[1], MRUNum[0])`, ...

Repeated references to the same *declared* variable within one key **co-vary**: they all bind to the same value within a single combination rather than producing an inner loop. This is the right tool for "version N's install path and version N's log folder": they always pair, never cross-multiply (see [Example 4](EXAMPLES.md#example-4-co-variance-of-repeated-variables)).

## Nested References

A variable's value list can itself contain `<Other>` tokens that reference other declared variables. EntryBuilder resolves these before any key is expanded, so each referencing key sees a fully-literal value list.

```ini
Versions=10.0,10.1
ArcMapKey=ArcMap,Desktop<Versions>\ArcMap
Detect=HKCU\Software\ESRI\<ArcMapKey>
```

resolves `ArcMapKey` to `(ArcMap, Desktop10.0\ArcMap, Desktop10.1\ArcMap)`, then `Detect` fans across that 3-value list. Cyclic dependencies (A references B references A) are reported (`Variable <A> participates in a cyclic dependency with <B>; values left unresolved`) and the involved variables are passed through unresolved. 

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
- **Want a base-path variable *without* automatic detection? Name it anything else**: `DiskRoot`, `AppRoot`, etc. (see [Example 1](EXAMPLES.md#example-1-a-pass-through-entry-with-a-single-value-variable)).

## Undeclared Tokens

Resolution of `<Name>` tokens whose name matches neither a declared variable nor an inline list depends on the key's expansion domain (see the [key-family table](#how-generation-works)):

| Domain | Behavior | Message | Why |
|:-|:-|:-|:-|
| Filesystem | Drop the key | `Undeclared variable <X> in filesystem-domain key [Entry].FileKey; dropping key` | Windows path syntax forbids `<` and `>`, so any such token is unambiguously a typo |
| Registry | Emit the literal text | `Undeclared variable <X> in registry-domain key [Entry].RegKey; passed through as literal` | Registry key names can legitimately contain `<` and `>`; the engine cannot assume a typo |

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
| No root keys and no FileKey/RegKey content in any form | Entry **skipped** | `[Entry] declares no WebViewRoot, QtWebEngineRoot, ElectronRoot, FileKeyBase, or RegKeyBase — nothing to emit, skipping` |
| No detection (after `Root` inference) | Entry emitted, always-on | `[Entry] declares no detection (Detect / DetectFile / DetectOS); generated entry will be always-on` |
| `Default=` declared | Key dropped | `Default= declared in [Entry]; EntryBuilder never emits Default, ignoring` |
| `SpecialDetect=` declared | Key dropped | `SpecialDetect is deprecated; key dropped from [Entry]. Replace with Detect or DetectFile` |
| `All` mixed with non-legacy scaffold names | Extras ignored | `WebViewScaffolds=All in [Entry] with redundant additional names (X, Y); ignoring` |
| Explicit scaffold list mixed with exclusions | Exclusions applied | `Both WebViewScaffolds and ExcludeWebViewScaffolds set in [Entry]; applying exclusions to explicit list` |
| Unknown scaffold name requested | Scaffold dropped | `Unknown WebView scaffold 'X' requested by [Entry], skipping` |
| Scaffold directory missing | Zero scaffold keys from every catalog | `Scaffold directory not found: <dir>; no scaffold keys will be generated` |
| Catalog file holds a section with no `Scaffold:` marker | Section ignored | `Unexpected section in scaffold catalog <file>: [Section]` |
| Catalog section has a `Tier=` value other than `Legacy` | Tier ignored, scaffold stays in `All` | `Unknown tier 'X' in WebViewScaffold [WebViewScaffold: Name]; the only tier is Legacy` |
| Undeclared `<X>` token, filesystem domain | Key dropped | `Undeclared variable <X> in filesystem-domain key [Entry].FileKey; dropping key` |
| Undeclared `<X>` token, registry domain | Literal emitted *(advisory)* | `Undeclared variable <X> in registry-domain key [Entry].RegKey; passed through as literal` |
| `<X[N]>` index is 0 or past the end of `X` | Key dropped | `[Entry].FileKey: <X[N]> is out of range (M values); dropping key` |
| The same, inside a variable declaration | That one value dropped from the variable | `[Entry]: <Cache> declaration: <X[N]> is out of range (M values); dropping key` |
| Referenced variable or inline list has no values | Key dropped | `Axis <X> has no values, referenced by [Entry].FileKey; dropping key` |
| Cyclic variable references | Values left as authored | `[Entry]: Variable <X> participates in a cyclic dependency with <Y>; values left unresolved` |
| Identical inline list repeated in one key | Occurrences co-vary *(advisory)* | `Inline list <a,b> appears more than once in [Entry].FileKey; occurrences co-vary: declare a named variable if that is not intended` |
| `Root` value has unresolved tokens | Detection inference skipped | `[Entry] Root has unresolved <token> references; skipping detection inference` |
| `Root` mixes registry and filesystem values | First value's domain wins | `[Entry] Root mixes registry and filesystem values; inferring detection from the first value's domain` |
| Declared variable never referenced | Warning only | `Variable 'X=' declared in [Entry] but never referenced by any <X> token; possible typo` |
| Variable name collides with a reserved key | Warning only | `Variable declaration 'X=' in [Entry] shadows reserved key name; possible typo` |

The scaffold-family messages appear with `QtWebEngine` or `Electron` in place of `WebView` for those families. The log also records each `All` expansion: `WebViewScaffolds=All in [Entry]; expanded to N scaffold(s) from catalog, M legacy withheld`.

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

`Caches` and `Telemetry` are the only opt-in-free WebView scaffolds. Anything that could remove user-visible state (cookies, history, sessions, saved logins, web storage) requires explicit opt-in: either `Scaffolds=All` with selective exclusion, or an explicit list naming the desired host-risk scaffolds. Legacy scaffolds are never selected implicitly, not even by `All`.

### Prefer `<Variable>` to numbered duplication

`RegKeyBase=<RegRoot>\<Versions>\Outlook` is easier to extend than seven hand-numbered `RegKey1..7` lines. When a new version ships, append it to the `Versions=` list and every dependent key grows automatically.

---

# Troubleshooting

| Symptom | Likely Cause | Fix |
|:-|:-|:-|
| `No EntryBuilder source definitions found in: <dir>` | Source directory is empty, missing, or contains no `*.ini` files with parseable sections | Verify the source directory in **Choose source directory** or via `-1d` |
| Entry is silently absent from the output | `Skip=` is set, or the entry was skipped by validation (`No Section or LangSecRef...` / `...nothing to emit, skipping`) | Check the run log for the corresponding message |
| `Duplicate EntryBuilder section [<name>] at <file> line <n> ignored; keeping the first definition, in <file>` | Two sections share a name (in any case), in one file or in two. Files are read in name order and only the first definition is used. The output is still written, but the run fails with exit code 1 | Merge the two sections, or rename one |
| Output is missing expected scaffold FileKeys | Scaffold directory not found (`Scaffold directory not found: <dir>...`), scaffold name misspelled (`Unknown WebView scaffold 'X'...`), or scaffold excluded by `Exclude*Scaffolds=` | Check the `Loaded N WebView scaffold(s)` / `Loaded N QtWebEngine scaffold(s)` lines and the warnings |
| A key holding `<WebViewRoot>` / `<QtWebEngineRoot>` / `<QtWebEngineCacheRoot>` / `<ElectronRoot>` / `<ElectronUpdaterRoot>` is missing from the output | The matching root key wasn't declared on the entry, which drops the key silently, or the placeholder was used in a `DetectFile`, where it is never substituted and drops as an undeclared token (a `Detect` keeps it literally) | Declare the root key, or write the path / a `<Variable>` directly in detection keys |
| `... is now spelled <WebViewRoot>; key dropped` and exit code 1 | The key still uses a retired `%WebViewRoot%`-style spelling | Respell it as the message says (see [The retired `%WebViewRoot%` spelling](#the-retired-webviewroot-spelling)) |
| Entry gained a `Detect`/`DetectFile` you didn't write | The entry declares a variable named `Root` - detection inference is automatic | Intended? Delete your redundant hand-written detection. Not intended? Rename the variable (`DiskRoot`, `AppRoot`, ...) |
| Output RegKey contains a literal `<Name>` | The variable wasn't declared - the registry domain emits literal text and logs `...passed through as literal` as an advisory | Declare the variable, or accept the literal if intentional |
| Output is missing a FileKey/DetectFile entirely | An undeclared `<Name>` token dropped the key (`...dropping key`), or a referenced variable had no values | Fix the typo, or declare the variable |
| `Variable 'X=' declared ... but never referenced` warning | The declaration is unused - usually a typo in a `<X>` reference somewhere | Search the entry for the intended reference and correct it |
| Output reorders keys differently from the source | WinappDebug normalization rewrites to canonical order and alphabetised numbering on every run | This is intentional; reformat the source to match if the diff bothers you |
| `Default=` was stripped from an entry | EntryBuilder never emits `Default=` by design | Move the entry to a non-EntryBuilder source if `Default=` is required |
