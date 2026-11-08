# UWPBuilder

**UWPBuilder** is a winapp2ool build-chain module for winapp2.ini maintainers that generates winapp2.ini entries for Universal Windows Platform (UWP) applications from a small templating DSL. Every UWP app stores its data under `%LocalAppData%\Packages\PACKAGE_FOLDER`, and part of that layout is identical from app to app, so UWPBuilder applies a shared scaffold of baseline cleaning targets to every application entry.

UWPBuilder is a build-pipeline module: its output is an intermediate file consumed by winapp2ool to produce the final winapp2.ini, and is not intended to be used directly with any cleaning software. If you are not maintaining winapp2.ini, you probably don't need this module.

### What does UWPBuilder do?
UWPBuilder reads a source directory containing a scaffold template (`UWP.ini`) and per-letter application definition files (`AppInfo\*.ini`), and generates one winapp2.ini entry per defined application. Each entry receives package detection and a consistent baseline set of FileKeys automatically; app-specific keys are layered on top, with the `<Package>` variable saving you from ever writing a package path out in full.

### Why UWPBuilder?
- Consistency: every UWP entry receives the same baseline cleaning coverage, generated from a single template
- Brevity: `Package=` + a category key is a complete entry; `<Package>` abstracts away verbose package folder paths everywhere else
- Hybrid app support: apps with both a UWP package and a win32 installation can carry both sets of keys in one definition
- Embedded browser coverage: apps using WebView2 (EBWebView), QtWebEngine, or Electron can generate FileKeys from shared scaffold catalogs

---

# Table of Contents

**Using UWPBuilder**

1. [Requirements](#requirements)
2. [Quick Start](#quick-start)
3. [Common Uses](#common-uses)
4. [Menu Options](#menu-options)
5. [Command-Line Arguments](#command-line-arguments)
   - [File Selection](#file-selection)
6. [How Generation Works](#how-generation-works)
7. [Source Files](#source-files)
   - [The source directory](#the-source-directory)
   - [UWP.ini: the entry scaffold](#uwpini-the-entry-scaffold)
   - [AppInfo files](#appinfo-files)
8. [AppInfo Key Reference](#appinfo-key-reference)
   - [Standard winapp2.ini keys](#standard-winapp2ini-keys)
   - [Winapp2ool-exclusive keys](#winapp2ool-exclusive-keys)
   - [Package variables](#package-variables)
   - [Root placeholders in AppInfo keys](#root-placeholders-in-appinfo-keys)
   - [Variables](#variables)
   - [Skip and SkipUWPFileKeys](#skip-and-skipuwpfilekeys)
9. [WebView Scaffolds](#webview-scaffolds)
   - [Opting in](#opting-in)
   - [Selecting scaffolds](#selecting-scaffolds)
   - [The scaffold catalog](#the-scaffold-catalog)
10. [QtWebEngine Scaffolds](#qtwebengine-scaffolds)
11. [Electron Scaffolds](#electron-scaffolds)
12. [Output Formatting](#output-formatting)
    - [When the lint pass changes meaning](#when-the-lint-pass-changes-meaning)
13. [Tips & Best Practices](#tips--best-practices)
14. [Troubleshooting](#troubleshooting)

**Examples**

15. **[Examples](EXAMPLES.md)**: eight worked examples with captured input and output

---

# Requirements

- A source directory containing:
  - `UWP.ini`: the UWP scaffold template
  - An `AppInfo\` subdirectory of `*.ini` files defining the applications
- Optionally, a scaffold directory holding the shared catalogs consumed when entries opt into embedded-browser scaffolding. Every `*.ini` in it is read, and each catalog's engine family comes from its section headers rather than its filename:
  - `webview.ini`: `[WebViewScaffold: ...]`, the WebView2/EBWebView catalog
  - `qtwebengine.ini`: `[QtWebEngineScaffold: ...]`, the QtWebEngine catalog
  - `electron.ini`: `[ElectronScaffold: ...]`, the Electron catalog

The winapp2.ini project's source files live in [`Assembler/UWP/`](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/UWP) (source directory) and [`Assembler/Scaffolds/`](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/Scaffolds) (scaffold directory)

###### Note: The scaffold directory is only required if an AppInfo entry declares a scaffold root. A missing directory, or a missing catalog within it, logs a warning and generation continues with zero scaffold FileKeys for the affected families.

---

# Quick Start

### Common Workflow
1. Point the source directory at your `UWP.ini` + `AppInfo\` folder
2. Point the scaffolds directory chooser at the shared `Scaffolds\` folder
3. Run UWPBuilder

To add a new application, add a section to the appropriate `AppInfo\{letter}.ini` file

---

# Common Uses

| To | Run | Example |
|:-|:-|:-|
| Generate `uwp.ini` from the sources in `UWP\`, using the shared scaffold catalogs | `winapp2ool -uwpbuilder -1d \UWP -2f uwp.ini -3d \Scaffolds` | [A Minimal Entry](EXAMPLES.md#example-1-a-minimal-entry) |
| Write the committed artifact silently and offline, as the build script does | `winapp2ool -s -offline -uwpbuilder -1d \UWP -2d \Entries\UWP -2f uwp.ini -3d \Scaffolds` | |

Every worked example runs the first command; what changes between them is the AppInfo section it reads. In the menu, the same choices are **Choose source directory**, **Choose save target** and **Choose scaffolds directory**, set before **Run**; see [Menu Options](#menu-options). `-s` and `-offline` are winapp2ool [global args](../../Readme.md#global-args).

---

# Menu Options

| Option | Effect | Notes |
|:-|:-|:-|
| Run (default) | Generate UWP app winapp2.ini entries | |
| Choose source directory | Select the directory containing `UWP.ini` and `AppInfo\` | Default: current directory |
| Choose save target | Select where to save the generated entries | Default: `uwp.ini` |
| Choose scaffolds directory | Select the folder holding the shared scaffold catalogs | Default: current directory |
| Reset Settings | Restore all settings to their defaults | Only shown when settings have been changed |

The menu also displays the current source directory, save target, and scaffolds directory.

---

# Command-Line Arguments

UWPBuilder supports command-line automation for scripting environments and is invoked as `winapp2ool -uwpbuilder`. It has no mode toggles.

### File Selection

| Arg | Effect | Default Value |
|:-|:-|:-|
| `-1d path` | Set the source directory (containing `UWP.ini` and `AppInfo\`) | Current directory |
| `-2d path` | Set the output directory | Current directory |
| `-2f name` | Set the output file name | `uwp.ini` |
| `-3d path` | Set the shared scaffold directory | Current directory |

Paths may be absolute, or relative to the working directory using a leading backslash (e.g. `-1d \UWP`).

###### Note: Slot 3 reads only its directory, so `-3f` has no effect on it. If it is left at its default and the working directory holds no catalogs, generation still succeeds with a warning and zero scaffold FileKeys. When scripting, always pass `-3d` explicitly. 

---

# How Generation Works

| Stage | What happens |
|:-|:-|
| Combine | All `AppInfo\*.ini` files are combined in-memory, in alphabetical path order. No intermediate file is created |
| Load scaffolds | `[EntryScaffold: ...]` sections are read from `UWP.ini`; every catalog in the scaffolds directory is loaded |
| Validate | Each AppInfo section is checked: at least one `Package=` and at least one category key are required. Sections failing either check are skipped with a warning; a section declaring *both* category keys warns but is still generated, using `LangSecRef` |
| Generate | One entry is generated per valid section: package detection, scaffold FileKeys, app-specific keys, and any selected WebView/QtWebEngine/Electron scaffold keys, with `<Package>` and every other variable expanded throughout |
| Lint | The complete output is normalized by WinappDebug with optimizations force-enabled (see [Output Formatting](#output-formatting)) |
| Write | The output file is written to disk |

###### Note: Sections with duplicate names across AppInfo files are silently ignored

---

# Source Files

## The source directory

UWPBuilder reads everything from a single source directory:

```
UWP\
├── UWP.ini            ← scaffold template
└── AppInfo\
    ├── A.ini          ← app definitions, split alphabetically
    ├── B.ini
    └── ...
```

The alphabetical split of `AppInfo\` is a convention for maintainability, not a requirement — all `*.ini` files in the subdirectory are combined regardless of name.

## UWP.ini: the entry scaffold

`UWP.ini` contains `[EntryScaffold: ...]` sections defining the baseline applied to every generated entry. Two key types are recognized:

| Key | Effect |
|:-|:-|
| `DetectFileBase=` | A DetectFile path template, expanded once per package to produce the entry's detection paths |
| `FileKeyBase=` | A FileKey template applied to every application as a baseline cleaning target |

The scaffold currently shipped with the winapp2.ini project:

```ini
[EntryScaffold: UWP App]
DetectFileBase=<Package>
FileKeyBase=<Package>\AC|*|RECURSE
FileKeyBase=<Package>\Settings|*.log*
FileKeyBase=<Package>\SystemAppData\Helium|*.log*
FileKeyBase=<Package>\TempState|*|REMOVESELF
```

## AppInfo files

Each section in an AppInfo file describes one winapp2.ini entry. The section header is used verbatim as the generated entry's name, so it should follow winapp2.ini naming rules (ending in ` *]`):

```ini
[ASUS Quick Launch *]
Package=B9ECED6F.ASUSQuickLaunch_qmba6cd70vzyy
LangSecRef=3024
```

Comments are welcome in all source files for your own reference; they are not carried into the output.

---

# AppInfo Key Reference

## Standard winapp2.ini keys

These keys carry their normal winapp2.ini meaning and are passed into the generated entry:

| Key | Behavior | Notes |
|:-|:-|:-|
| `LangSecRef=` / `Section=` | Category of the generated entry | Exactly one is required. If both are present, a warning is issued and `LangSecRef` is used |
| `Detect=` | Registry detection | Repeatable. Unnumbered in the output when only one is present. `<>`-expanded |
| `DetectFile=` | File system detection, appended after the scaffold-generated detection paths | Repeatable. `<>`-expanded, `<Package>` included. Use for hybrid apps that need win32 detection alongside package detection |
| `DetectOS=` | Windows kernel version filter, e.g. `6.1\|10.0` | Emitted verbatim, never expanded |
| `FileKey=` | Cleaning target; equivalent to `FileKeyBase=` | [Root placeholders](#root-placeholders-in-appinfo-keys) substituted, then `<>`-expanded, `<Package>` included; keys without a token pass through verbatim |
| `RegKey=` | Registry cleaning target | Renumbered from 1. [Root placeholders](#root-placeholders-in-appinfo-keys) substituted, then `<>`-expanded |
| `ExcludeKey=` | Exclusion; equivalent to `ExcludeKeyBase=` | [Root placeholders](#root-placeholders-in-appinfo-keys) substituted, then `<>`-expanded, `<Package>` included, in the domain the substituted value's flag names |
| `Warning=` | Warning shown to the user | Emitted verbatim, never expanded |

###### Note: Key numbers in the source are ignored: everything is renumbered during generation and normalized again by the lint pass.

`SpecialDetect=` is deprecated and ignored with a warning. `Default=` is warned and ignored.

Every other unrecognised key name is a [variable declaration](#variables). There is no longer an "unexpected key type" error

## Winapp2ool-exclusive keys

| Key | Effect | Notes |
|:-|:-|:-|
| `Package=` | The application's package folder name under `%LocalAppData%\Packages` | At least one is required. Repeatable as `Package1=`, `Package2=`, ... for multi-package entries |
| `FileKeyBase=` | A FileKey template; equivalent to `FileKey=` | Convention: use `FileKeyBase=` for `<Package>`-relative keys and `FileKey=` for win32 paths |
| `ExcludeKeyBase=` | An ExcludeKey template; equivalent to `ExcludeKey=` | |
| `RegKeyBase=` | A RegKey template; equivalent to `RegKey=` | |
| `Skip` | Omits this application from generation entirely | Presence-based: see below |
| `SkipUWPFileKeys` | Suppresses the scaffold `FileKeyBase=` templates for this entry only | Presence-based: see below. Detection keys are unaffected |
| `WebViewRoot=` (alias `WebViewPath=`) | Declares an embedded WebView2/EBWebView data root; opts the entry into [WebView scaffolds](#webview-scaffolds) | Repeatable (numbered or not). `<Package>` variables are expanded |
| `WebViewScaffolds=` | Comma-separated scaffold selection, replacing the default set | `All` sentinel expands to the whole catalog except the legacy tier |
| `ExcludeWebViewScaffolds=` | Comma-separated scaffold names subtracted from the active selection | |
| `QtWebEngineRoot=` (alias `QtWebEnginePath=`) | Declares one embedded QtWebEngine profile directory (profile segment included, e.g. `...\QtWebEngine\Default`); opts the entry into [QtWebEngine scaffolds](#qtwebengine-scaffolds) | Repeatable, one per profile. `<Package>` variables are expanded |
| `QtWebEngineCacheRoot=` (alias `QtWebEngineCachePath=`) | Declares a QtWebEngine HTTP cache directory for `<QtWebEngineCacheRoot>`; also opts the entry into QtWebEngine scaffolds | Repeatable. `<Package>` variables are expanded |
| `QtWebEngineScaffolds=` | Scaffold selection for the QtWebEngine family | Same grammar as `WebViewScaffolds=` |
| `ExcludeQtWebEngineScaffolds=` | Exclusions for the QtWebEngine family | |
| `ElectronRoot=` (alias `ElectronPath=`) | Declares an Electron `userData` folder; opts the entry into [Electron scaffolds](#electron-scaffolds) | Repeatable. `<Package>` variables are expanded |
| `ElectronUpdaterRoot=` (alias `ElectronUpdaterPath=`) | Declares an electron-updater download cache for `<ElectronUpdaterRoot>`; also opts the entry into Electron scaffolds | Repeatable. `<Package>` variables are expanded |
| `ElectronScaffolds=` | Scaffold selection for the Electron family | Same grammar as `WebViewScaffolds=` |
| `ExcludeElectronScaffolds=` | Exclusions for the Electron family | |

## Package variables

`<Package>` is a [variable](#variables) UWPBuilder declares on every entry, with one value per `Package=` key: `%LocalAppData%\Packages\PACKAGE_FOLDER`. It works in the value of any expanded key, the root keys (`WebViewRoot=` and the rest) and other variable declarations included:

- **`<Package>`** expands once per package: a key containing it produces one output key per declared package, in declaration order
- **`<Package[N]>`** selects exactly one package by position: a key containing `<Package[2]>` produces a single output key using the second package. An `N` of 0 or past the number of packages drops the key with a warning
- A key containing no package variable passes through verbatim as a single key

`<Package>` and `<Package[2]>` in one key are two separate variables, so they multiply like any two variables. The same goes for `<Package>` beside a variable declared from it (`Cache=<Package>\LocalCache`): a key using both pairs every package with every other package's cache, so use one or the other in a key. See [Example 4](EXAMPLES.md#example-4-multi-package-applications) for both forms in action.

Package detection comes from the scaffold's `DetectFileBase=`, so an app-specific `DetectFile=<Package>` is never needed.

### The retired `%Package%` and `%WebViewRoot%` spellings

Winapp2ool v1.7.x.x versions spelled these `%Package%` and `%Package1%`, `%Package2%`, and the [root placeholders](#root-placeholders-in-appinfo-keys) `%WebViewRoot%`, `%QtWebEngineRoot%`, `%QtWebEngineCacheRoot%`, `%ElectronRoot%` and `%ElectronUpdaterRoot%`. A key still written that way, including one with the closing `%` missing (`%Package\LocalCache`, `%ElectronRoot\logs`), is dropped with a warning naming the new spelling, and the run fails with exit code 1:

```
`%WebViewRoot%` in [Test App *].FileKeyBase is now spelled `<WebViewRoot>`; key dropped
```

`Package=` and `Warning=` values are not checked. The shared scaffold catalogs refuse the same spellings as they load: a catalog template holding one is dropped (`...; template dropped`) and the run fails.

## Root placeholders in AppInfo keys

An app's own `FileKey=` / `FileKeyBase=`, `RegKey=` / `RegKeyBase=` and `ExcludeKey=` / `ExcludeKeyBase=` values take the same root placeholders as the scaffold catalogs: `<WebViewRoot>`, `<QtWebEngineRoot>`, `<QtWebEngineCacheRoot>`, `<ElectronRoot>` and `<ElectronUpdaterRoot>`, matched in any case. Each is replaced with every path the entry declared for it (`WebViewRoot=` / `WebViewPath=` and so on) before variables are expanded, so a key fans out once per path, and a `<Package>` inside a path expands per package along with the key:

```ini
[Hazel Viewer *]
Package=Hazel.Viewer_abc123
LangSecRef=3021
WebViewPath1=<Package>\LocalState
WebViewPath2=%LocalAppData%\Hazel Viewer
FileKeyBase=<WebViewRoot>\EBWebView\*\Cache|*
```

generates `%LocalAppData%\Packages\Hazel.Viewer_abc123\LocalState\EBWebView\*\Cache|*` and `%LocalAppData%\Hazel Viewer\EBWebView\*\Cache|*`, alongside the baseline FileKeys and the default WebView scaffolds for both paths. A key naming a root the entry declared no path for is dropped without a warning, as a catalog template is. An ExcludeKey's flag is read after substitution. `Detect=` and `DetectFile=` values are never root-substituted: a placeholder there is an undeclared token.

## Variables

Any key whose name UWPBuilder does not recognise is a variable declaration: a comma-delimited list of values, referenced elsewhere in the entry as `<Name>`. Each reference fans the containing key out into one key per value.

```ini
[Hazel Office *]
Package=Hazel.Office_abc123
LangSecRef=3021
Version=11.0,16.0
RegKeyBase=HKCU\Software\Hazel\<Version>\File MRU
```

generates

```ini
RegKey1=HKCU\Software\Hazel\11.0\File MRU
RegKey2=HKCU\Software\Hazel\16.0\File MRU
```

This is the same engine EntryBuilder uses:

- Two distinct variables in one template produce the **cartesian product** of their values
- A repeated token in one template **co-varies** rather than multiplying
- An **inline list** needs no declaration: `<Default,OffTheRecord>`
- Variables may reference other variables; declarations are flattened before generation
- Matching is case-insensitive

### Undeclared tokens

What happens to a `<Token>` with no declaration depends on where the key lands:

| Domain | Keys | Behavior |
|:-|:-|:-|
| Registry | `Detect=`, `RegKey=`/`RegKeyBase=`, `REG` `ExcludeKey=` | Kept literally, advisory logged |
| File system | `DetectFile=`, `FileKey=`/`FileKeyBase=`, all other `ExcludeKey=` | Dropped, warning logged |

### What is never expanded

- **`Package=`**: a fan-out would change the package list's length and silently shift the meaning of every `<Package[N]>` selector in the entry. A `<` or `>` in a `Package=` value is warned about and used literally, and the `Package` variable holds it unexpanded
- **`Warning=`** and **`DetectOS=`**: prose and a version range respectively, not path templates

### Typo backstops

Because the vocabulary is open, a misspelled key name no longer errors; it becomes a variable. There are two checks for this:

1. Undeclared-token diagnostics catch a misspelled *reference* (`<Verison>` against a `Version=` declaration)
2. Unreferenced-declaration warnings catch a misspelled *declaration*: any variable that no `<Token>` references is reported as a probable typo after generation.

###### Note: like every other diagnostic, these are invisible during a scripted build. See [Troubleshooting](#troubleshooting).

## Skip and SkipUWPFileKeys

Both keys are presence-based, not boolean. The value does not matter

```ini
Skip=False
```

still skips the entry. Provide the key if and only if you want the effect; to disable it, delete the key entirely.

- `Skip`: the application is omitted from generation. Used to discontinue support for an app without deleting its configuration
- `SkipUWPFileKeys`: the scaffold `FileKeyBase=` templates from `UWP.ini` are not generated for this entry, but everything else is generated normally. See [Example 7](EXAMPLES.md#example-7-secondary-entries-with-skipuwpfilekeys)

---

# WebView Scaffolds

Many applications embed a Chromium-based WebView2 browser whose data folder (usually named `EBWebView`) follows a standard layout. Rather than hand-writing the same cleaning keys for every such app, UWPBuilder draws them from a shared catalog of `[WebViewScaffold: ...]` sections: the same catalog consumed by [EntryBuilder](../entrybuilder/readme.md).

## Opting in

Declaring one or more `WebViewRoot=` keys. No `WebViewRoot=` means zero WebView keys are generated, regardless of any other WebView key:

```ini
[GoldKey Acellus *]
Package=GoldKeyCorporation.Acellus_v2rn50m687qry
WebViewRoot=<Package>\LocalState\EBWebView
LangSecRef=3021
```

Each selected scaffold's templates are expanded once per declared root, with `<WebViewRoot>` substituted for the (package-expanded) root path. The root is the folder that holds the profiles; profile-scoped templates use `<WebViewRoot>\*\`, so `Default` and named profiles such as `WV2Profile_*` are all covered. An app with two WebView folders declares two roots (`WebViewRoot1=`, `WebViewRoot2=`) and receives the full selection for each. The root does not have to be package-relative: a win32 path like `%AppData%\Microsoft\Teams` is equally valid.

## Selecting scaffolds

| Declaration | Resulting selection |
|:-|:-|
| No `WebViewScaffolds=` key | The default set: Caches, Telemetry |
| `WebViewScaffolds=Name1,Name2` | Exactly the named scaffolds |
| `WebViewScaffolds=` (present but empty) | Nothing  |
| `WebViewScaffolds=All` | Every scaffold in the catalog except the legacy tier, including host-risk categories |
| `WebViewScaffolds=All,LegacyTelemetry` | `All`, plus the named legacy scaffold |
| `ExcludeWebViewScaffolds=Name1,Name2` | Subtracted from whichever selection is active |

- `All` is a reserved value and cannot be used as a catalog scaffold name. Naming a non-legacy scaffold alongside `All` is redundant and warns
- A scaffold marked `Tier=Legacy` in the catalog is withheld from `All` and generated only when named, alone or beside `All`
- `All` + `ExcludeWebViewScaffolds=` is "everything except these" and does not warn
- Unknown names in `WebViewScaffolds=` are dropped with a warning, including beside `All`
- Unknown names in `ExcludeWebViewScaffolds=` are silent.
- All scaffold name matching is case-insensitive

## The scaffold catalog

The current set of scaffolds for WebView2 is: `Autofill`, `Autoplay`, `BookmarkFavicons`, `Caches`, `DefaultApps`, `DownloadHistory`, `DRMData`, `PrivacySandbox`, `LoginData`, `Security`, `Shopping`, `Sync Data`, `Telemetry`, `WebCookies`, `WebHistory`, `WebSession`, `WebStorage`.

The legacy tier holds patterns that modern host don't write but an older ones may: `LegacyCaches`, `LegacyDownloadHistory`, `LegacyDRMData`, `LegacyExtensionCookies`, `LegacyProgressiveWebApps`, `LegacyStorageQuota`, `LegacyTelemetry`, `LegacyWebHistory`.

Only Caches and Telemetry are generated by default. The remaining scaffolds require explicit opt-in, and legacy scaffolds must be named even when `All` is selected.

---

# QtWebEngine Scaffolds

A second, structurally identical scaffold family covers applications embedding QtWebEngine, from `qtwebengine.ini` in the scaffolds directory.

The grammar mirrors the WebView family: `QtWebEngineRoot=` (alias `QtWebEnginePath=`) opts in, and `QtWebEngineScaffolds=` / `ExcludeQtWebEngineScaffolds=` select (with the same `All` value and warning behavior). Two things differ.

The root names a profile directory, profile segment included: `...\QtWebEngine\Default`, not the `QtWebEngine` folder above it. Qt keeps the HTTP cache outside that directory, so a second root, `QtWebEngineCacheRoot=` (alias `QtWebEngineCachePath=`), declares it for `<QtWebEngineCacheRoot>`. Either root alone opts the entry in, and a template whose placeholder has no declared root is dropped.

The default selection is wider: Caches + StorageQuota + Telemetry + VisitedLinks. QtWebEngine's catalog breaks out two low-risk targets (`QuotaManager`, `Visited Links`) that the WebView catalog keeps out of its default set, in the legacy `LegacyStorageQuota` and the host-risk `WebHistory` respectively; see the [EntryBuilder readme](../entrybuilder/readme.md#qtwebengineroot--qtwebenginecacheroot--qtwebenginescaffolds) or the catalog header for why.

The QtWebEngine catalog is smaller. The current set is: `Caches`, `Favicons`, `PrivacySandbox`, `Security`, `StorageQuota`, `Telemetry`, `VisitedLinks`, `WebCookies`, `WebHistory`, `WebSession`, `WebStorage`, plus the legacy `LegacyTelemetry` and `LegacyWebStorage`. Caches, StorageQuota, Telemetry and VisitedLinks are generated by default; the rest require explicit opt-in.

---

# Electron Scaffolds

A third family covers applications embedding Electron, from `electron.ini` in the scaffolds directory.

`ElectronRoot=` (alias `ElectronPath=`) opts in, and `ElectronScaffolds=` / `ExcludeElectronScaffolds=` select, with the same `All` value and warning behavior as the other families. 

Like QtWebEngine, this family has a second root. `ElectronUpdaterRoot=` declares the electron-updater download cache. Either root alone opts the entry into the family. A template whose placeholder has no declared root is dropped.

```ini
[Hazel Music *]
Package=Hazel.Music_abc123
LangSecRef=3021
ElectronRoot=%LocalAppData%\Hazel Music
ElectronUpdaterRoot=%LocalAppData%\Hazel-music-updater
```

The default selection is `AppLogs`, `Caches`, `StorageQuota`, `Telemetry`, `UpdaterCache`. The full set adds `MediaDRM`, `PrivacySandbox`, `Security` and `TempFiles`, plus the host-risk `WebCookies` and `WebStorage`, all of which require explicit opt-in. The Electron catalog has no legacy tier.

---

# Output Formatting

Before being written to disk, the generated output is normalized by WinappDebug with Optimizations enabled. As a result:

- Entries are sorted alphabetically
- Keys are grouped and ordered by type (category -> detection -> cleaning), renumbered sequentially, and FileKeys are sorted alphabetically by value
- FileKeys targeting the same location are merged into a single key with a combined parameter list

This means the output will not visibly reflect the order in which keys were defined or generated.

The output file opens with a generated header comment block:

```ini
; # of entries: 187
; uwp.ini is generated by the Winapp2ool UWP Builder
; Entries in this file may be incomplete and are not intended to be used directly with any cleaning software
; They are utilized by winapp2ool to create the final winapp2.ini file for distribution
; If you are not maintaining winapp2.ini for distribution, you probably don't need this file!
; Refer to the Winapp2ool documentation for more information: https://github.com/MoscaDotTo/Winapp2/blob/master/winapp2ool/Readme.md
; You can find the complete winapp2.ini file here: https://raw.githubusercontent.com/MoscaDotTo/Winapp2/master/Non-CCleaner/Winapp2.ini
```

## When the lint pass changes meaning

The normalization above is supposed to be formatting-only. Every entry is decomposed into semantic units before and after the lint pass and compared. The repairs and optimizations (reordering, renumbering, pattern alphabetization, same-path merges, exact-duplicate removal) are invisible to that comparison, so a difference means the generator output something the linter silently rewrote or destroyed.

On a difference the run reports every finding, writes the pre-lint content to `uwp.prelint.ini` beside the output, and sets exit code 1, which fails the scripted build rather than committing bad data.

This indicates a bug in the data output by the generator, not a normal validation warning: reaching it means either an AppInfo value or a scaffold template produces a key the linter considers wrong. Every generative module (BrowserBuilder, UWPBuilder, EntryBuilder) runs this check.

---

# Tips & Best Practices

- **Use `FileKeyBase=` for package-relative keys and `FileKey=` for win32 paths.** The two are functionally identical, but the convention keeps hybrid entries readable
- **Never write `Skip=False`.** `Skip` and `SkipUWPFileKeys` are presence-based; delete the key to disable the effect
- **Don't add key numbers.** Source numbering is ignored and the lint pass renumbers, reorders, and may merge FileKeys
- **Comment the AppInfo sections freely.** Comments never reach the output

---

# Troubleshooting

###### Note: None of the messages below are visible during a scripted build. Silent mode (`-s`) suppresses all menu output, and the global log is only flushed to `winapp2ool.log` when the run exits nonzero

| Message | Cause |
|:-|:-|
| "UWP.ini not found or empty in: {dir}" | The source directory doesn't contain a parseable `UWP.ini`. |
| "No app definitions found in: {dir}" | The `AppInfo\` subdirectory is missing, empty, or contains no parseable sections |
| "Unexpected section in template file: [{name}]" | `UWP.ini` contains a section that isn't an `[EntryScaffold: ...]` |
| "Unexpected key in scaffold [{name}]: {key}" | An `[EntryScaffold: ...]` section contains a key other than `DetectFileBase=`/`FileKeyBase=` |
| "`%Package%` in [{name}].{key} is now spelled `<Package>`; key dropped" | The key uses a [retired spelling](#the-retired-package-and-webviewroot-spellings); the run fails with exit code 1. A shared catalog template reports `; template dropped` instead |
| "Unexpected key type in [{name}]: {key}" | An AppInfo section contains an unsupported key; it is dropped from the generated entry |
| "No Package key in [{name}], skipping" | The AppInfo section has no `Package=` key; the entry is not generated |
| "No LangSecRef or Section in [{name}], skipping" | The AppInfo section has no category key; the entry is not generated |
| "Both LangSecRef and Section present in [{name}], using LangSecRef" | The AppInfo section declares both category keys; the entry is generated with `LangSecRef` and the `Section` value is discarded |
| "Skipping scaffold FileKeys for: {name}" | Informational; the entry declares `SkipUWPFileKeys` |
| "WebViewScaffold catalog at {path} is empty or missing" | The catalog file wasn't found; generation continues with zero WebView scaffold FileKeys. |
| "Unknown WebView scaffold '{name}' requested by [{entry}], skipping" | A `WebViewScaffolds=` name doesn't exist in the catalog. Note that unknown `ExcludeWebViewScaffolds=` names produce **no** message |
| "WebViewScaffolds=All in [{entry}] with redundant additional names ({names}); ignoring" | A non-legacy scaffold was named beside `All`, which already selects it |
| "Both WebViewScaffolds and ExcludeWebViewScaffolds set in [{entry}]; applying exclusions to explicit list" | Exclusions usually pair with the implicit defaults or `All` |
| "Duplicate WebViewScaffold name '{name}'; last definition wins" | The catalog defines the same scaffold twice |
| "Unexpected section in WebViewScaffold catalog: [{name}]" | The catalog contains a section without the family prefix; it is ignored |
| "Unknown tier '{value}' in WebViewScaffold [{section}]; the only tier is Legacy" | A catalog section has a `Tier=` value other than `Legacy`; the tier is ignored and the scaffold stays in `All` |

QtWebEngine- and Electron-family messages are identical with `QtWebEngine` or `Electron` in place of `WebView`. Each `All` expansion is also logged: `WebViewScaffolds=All in [{entry}]; expanded to N scaffold(s) from catalog, M legacy withheld`.
