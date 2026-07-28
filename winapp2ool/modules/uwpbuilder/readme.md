# UWPBuilder

**UWPBuilder** is a winapp2ool module that generates winapp2.ini entries for Universal Windows Platform (UWP) applications from a small templating DSL. Every UWP app stores its data under `%LocalAppData%\Packages\<PackageFolderName>`, and part of that layout is identical from app to app, so UWPBuilder applies a shared scaffold of baseline cleaning targets to every application entry.

UWPBuilder is a build-pipeline module: its output is an intermediate file consumed by winapp2ool to produce the final winapp2.ini, and is not intended to be used directly with any cleaning software. If you are not maintaining winapp2.ini, you probably don't need this module.

### What does UWPBuilder do?
UWPBuilder reads a source directory containing a scaffold template (`UWP.ini`) and per-letter application definition files (`AppInfo\*.ini`), and generates one winapp2.ini entry per defined application. Each entry receives package detection and a consistent baseline set of FileKeys automatically; app-specific keys are layered on top, with the `%Package%` variable saving you from ever writing a package path out in full.

### Why UWPBuilder?
- Consistency: every UWP entry receives the same baseline cleaning coverage, generated from a single template
- Brevity: `Package=` + a category key is a complete entry; `%Package%` abstracts away verbose package folder paths everywhere else
- Hybrid app support: apps with both a UWP package and a win32 installation can carry both sets of keys in one definition
- Embedded browser coverage: apps using WebView2 (EBWebView) or QTWebEngine can generate FileKeys from a shared scaffold catalog

---

# Table of Contents
1. [Requirements](#requirements)
2. [Quick Start](#quick-start)
3. [Menu Options](#menu-options)
4. [How Generation Works](#how-generation-works)
5. [Source Files](#source-files)
   - [The source directory](#the-source-directory)
   - [UWP.ini — the entry scaffold](#uwpini--the-entry-scaffold)
   - [AppInfo files](#appinfo-files)
6. [AppInfo Key Reference](#appinfo-key-reference)
   - [Standard winapp2.ini keys](#standard-winapp2ini-keys)
   - [Winapp2ool-exclusive keys](#winapp2ool-exclusive-keys)
   - [Package variables](#package-variables)
   - [Skip and SkipUWPFileKeys](#skip-and-skipuwpfilekeys)
7. [WebView Scaffolds](#webview-scaffolds)
   - [Opting in](#opting-in)
   - [Selecting scaffolds](#selecting-scaffolds)
   - [The scaffold catalog](#the-scaffold-catalog)
8. [QtWebEngine Scaffolds](#qtwebengine-scaffolds)
9. [Output Formatting](#output-formatting)
   - [When the lint pass changes meaning](#when-the-lint-pass-changes-meaning)
10. [Command-Line Arguments](#command-line-arguments)
    - [File Selection](#file-selection)
    - [Examples](#examples)
11. [Tips & Best Practices](#tips--best-practices)
12. [Troubleshooting](#troubleshooting)
13. [Usage Examples](#usage-examples)

---

# Requirements

- A source directory containing:
  - `UWP.ini`: the UWP scaffold template
  - An `AppInfo\` subdirectory of `*.ini` files defining the applications
- Optionally, the shared scaffold catalogs consumed when entries opt into embedded-browser scaffolding:
  - `webview.ini`: the WebView2/EBWebView catalog
  - `qtwebengine.ini`: the QtWebEngine catalog

The winapp2.ini project's source files live in [`Assembler/UWP/`](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/UWP) (source directory) and [`Assembler/Scaffolds/`](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/Scaffolds) (catalogs)

###### Note: The scaffold catalogs are only required if an AppInfo entry declares a WebView or QtWebEngine root. A missing or empty catalog logs a warning and generation continues with zero scaffold FileKeys for that family.

---

# Quick Start

### Common Workflow
1. Point the source directory at your `UWP.ini` + `AppInfo\` folder
2. Point the WebView and QtWebEngine catalog choosers at the shared `Scaffolds\` files
3. Run UWPBuilder

To add a new application, add a section to the appropriate `AppInfo\{letter}.ini` file

---

# Menu Options

| Option | Effect | Notes |
|:-|:-|:-|
| Run (default) | Generate UWP app winapp2.ini entries | |
| Choose source directory | Select the directory containing `UWP.ini` and `AppInfo\` | Default: current directory |
| Choose save target | Select where to save the generated entries | Default: `uwp.ini` |
| Choose webview scaffold | Select the shared WebView scaffold catalog | Default: `webview.ini` |
| Choose QtWebEngine scaffold | Select the shared QtWebEngine scaffold catalog | Default: `qtwebengine.ini` |
| Reset Settings | Restore all settings to their defaults | Only shown when settings have been changed |

The menu also displays the current source directory, save target, and both catalog paths.

---

# How Generation Works

| Stage | What happens |
|:-|:-|
| Combine | All `AppInfo\*.ini` files are combined in-memory, in alphabetical path order. No intermediate file is created |
| Load scaffolds | `[EntryScaffold: ...]` sections are read from `UWP.ini`; the WebView and QtWebEngine catalogs are loaded from their configured paths |
| Validate | Each AppInfo section is checked: at least one `Package=` and at least one category key are required. Sections failing either check are skipped with a warning; a section declaring *both* category keys warns but is still generated, using `LangSecRef` |
| Generate | One entry is generated per valid section: package detection, scaffold FileKeys, app-specific keys, and any selected WebView/QtWebEngine scaffold keys, with `%Package%` variables expanded throughout |
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
DetectFileBase=%Package%
FileKeyBase=%Package%\AC|*|RECURSE
FileKeyBase=%Package%\Settings|*.log*
FileKeyBase=%Package%\SystemAppData\Helium|*.log*
FileKeyBase=%Package%\TempState|*|REMOVESELF
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
| `Detect=` | Registry detection, passed through verbatim | Repeatable. Unnumbered in the output when only one is present |
| `DetectFile=` | File system detection, passed through verbatim and appended after the scaffold-generated detection paths | Repeatable. Not `%Package%`-expanded. Use for hybrid apps that need win32 detection alongside package detection |
| `FileKey=` | Cleaning target; equivalent to `FileKeyBase=` | `%Package%` variables are expanded; keys without variables pass through verbatim |
| `RegKey=` | Registry cleaning target, passed through verbatim | Renumbered from 1 in the output |
| `ExcludeKey=` | Exclusion; equivalent to `ExcludeKeyBase=` | `%Package%` variables are expanded |

###### Note: Key numbers in the source are ignored: everything is renumbered during generation and normalized again by the lint pass.

Any key type not listed on this page is reported with a warning and dropped: it does not appear in the output. UWPBuilder does not currently pass `Warning=` keys through.

## Winapp2ool-exclusive keys

| Key | Effect | Notes |
|:-|:-|:-|
| `Package=` | The application's package folder name under `%LocalAppData%\Packages` | At least one is required. Repeatable as `Package1=`, `Package2=`, ... for multi-package entries |
| `FileKeyBase=` | A FileKey template; equivalent to `FileKey=` | Convention: use `FileKeyBase=` for `%Package%`-relative keys and `FileKey=` for win32 paths |
| `ExcludeKeyBase=` | An ExcludeKey template; equivalent to `ExcludeKey=` | |
| `Skip` | Omits this application from generation entirely | Presence-based: see below |
| `SkipUWPFileKeys` | Suppresses the scaffold `FileKeyBase=` templates for this entry only | Presence-based: see below. Detection keys are unaffected |
| `WebViewRoot=` (alias `WebViewPath=`) | Declares an embedded WebView2/EBWebView data root; opts the entry into [WebView scaffolds](#webview-scaffolds) | Repeatable (numbered or not). `%Package%` variables are expanded |
| `WebViewScaffolds=` | Comma-separated scaffold selection, replacing the default set | `All` sentinel expands to the whole catalog |
| `ExcludeWebViewScaffolds=` | Comma-separated scaffold names subtracted from the active selection | |
| `QtWebEngineRoot=` (alias `QtWebEnginePath=`) | Declares one embedded QtWebEngine profile directory (profile segment included, e.g. `...\QtWebEngine\Default`); opts the entry into [QtWebEngine scaffolds](#qtwebengine-scaffolds) | Repeatable, one per profile. `%Package%` variables are expanded |
| `QtWebEngineScaffolds=` | Scaffold selection for the QtWebEngine family | Same grammar as `WebViewScaffolds=` |
| `ExcludeQtWebEngineScaffolds=` | Exclusions for the QtWebEngine family | |

## Package variables

`%Package%` expands to `%LocalAppData%\Packages\<PackageFolderName>` and may be used in the value of any expanded key (`FileKeyBase=`/`FileKey=`, `ExcludeKeyBase=`/`ExcludeKey=`, `WebViewRoot=`, `QtWebEngineRoot=`):

- **Unnumbered `%Package%`** expands once per package: a key containing it produces one output key per declared package, in order
- **Numbered `%PackageN%`** selects exactly one package by position: a key containing `%Package2%` produces a single output key using the second package
- A key containing no package variable passes through verbatim as a single key

See [Example 4](#example-4-multi-package-applications) for both forms in action.

Note that `Detect=`, `DetectFile=`, and `RegKey=` in an AppInfo section are **not** expanded: they pass through verbatim. `%Package%` in a `DetectFile=` reaches the output as a literal. Package detection comes from the scaffold's `DetectFileBase=`, so you should not need one.

## Skip and SkipUWPFileKeys

Both keys are presence-based, not boolean. The value does not matter

```ini
Skip=False
```

still skips the entry. Provide the key if and only if you want the effect; to disable it, delete the key entirely.

- `Skip`: the application is omitted from generation. Used to discontinue support for an app without deleting its configuration
- `SkipUWPFileKeys`: the scaffold `FileKeyBase=` templates from `UWP.ini` are not generated for this entry, but everything else is generated normally. See [Example 7](#example-7-secondary-entries-with-skipuwpfilekeys)

---

# WebView Scaffolds

Many applications embed a Chromium-based WebView2 browser whose data folder (usually named `EBWebView`) follows a standard layout. Rather than hand-writing the same cleaning keys for every such app, UWPBuilder draws them from a shared catalog of `[WebViewScaffold: ...]` sections: the same catalog consumed by [EntryBuilder](../entrybuilder/readme.md).

## Opting in

Declaring one or more `WebViewRoot=` keys. No `WebViewRoot=` means zero WebView keys are generated, regardless of any other WebView key:

```ini
[GoldKey Acellus *]
Package=GoldKeyCorporation.Acellus_v2rn50m687qry
WebViewRoot=%Package%\LocalState\EBWebView
LangSecRef=3021
```

Each selected scaffold's templates are expanded once per declared root, with `%WebViewRoot%` substituted for the (package-expanded) root path. An app with two WebView folders declares two roots (`WebViewRoot1=`, `WebViewRoot2=`) and receives the full selection for each. The root does not have to be package-relative: a win32 path like `%AppData%\Microsoft\Teams` is equally valid.

## Selecting scaffolds

| Declaration | Resulting selection |
|:-|:-|
| No `WebViewScaffolds=` key | The default set: Caches, Telemetry |
| `WebViewScaffolds=Name1,Name2` | Exactly the named scaffolds |
| `WebViewScaffolds=` (present but empty) | Nothing  |
| `WebViewScaffolds=All` | Every scaffold in the catalog, including host-risk categories |
| `ExcludeWebViewScaffolds=Name1,Name2` | Subtracted from whichever selection is active |

- `All` is a reserved value and cannot be used as a catalog scaffold name. Naming additional scaffolds alongside `All` is redundant and warns
- `All` + `ExcludeWebViewScaffolds=` is "everything except these" and does not warn
- Unknown names in `WebViewScaffolds=` are dropped with a warning
- Unknown names in `ExcludeWebViewScaffolds=` are silent.
- All scaffold name matching is case-insensitive

## The scaffold catalog

The current set of scaffolds for WebView2 is: `Autofill`, `Autoplay`, `BookmarkBackups`, `BookmarkFavicons`, `Caches`, `DefaultApps`, `DownloadHistory`, `DRMData`, `ExtensionCookies`, `ProgressiveWebApps`, `PrivacySandbox`, `LoginData`, `Security`, `Shopping`, `StorageQuota`, `Sync Data`, `Telemetry`, `WebCookies`, `WebHistory`, `WebSession`, `WebStorage`.

Only Caches and Telemetry are generated by default. The remaining scaffolds require explicit opt-in.

---

# QtWebEngine Scaffolds

A second, structurally identical scaffold family covers applications embedding QtWebEngine (with its own catalog at `Assembler/Scaffolds/qtwebengine.ini`, configurable via `-4f`/`-4d`).

The grammar mirrors the WebView family: `QtWebEngineRoot=` (alias `QtWebEnginePath=`) opts in, and `QtWebEngineScaffolds=` / `ExcludeQtWebEngineScaffolds=` select (with the same `All` value and warning behavior). Two things differ.

The root names a profile directory, profile segment included: `...\QtWebEngine\Default`, not the `QtWebEngine` folder above it.

The default selection is wider: Caches + StorageQuota + Telemetry + VisitedLinks. QtWebEngine's catalog breaks out two low-risk targets (`QuotaManager`, `Visited Links`) that the WebView catalog leaves inside host-risk scaffolds; see the [EntryBuilder readme](../entrybuilder/readme.md#qtwebengineroot--qtwebenginescaffolds) or the catalog header for why.

The QtWebEngine catalog is smaller. The current set is: `Caches`, `StorageQuota`, `Telemetry`, `VisitedLinks`, `WebCookies`, `WebHistory`, `WebSession`, `WebStorage`. The first four are generated by default; the remaining four require explicit opt-in.

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

# Command-Line Arguments

UWPBuilder supports command-line automation for scripting environments and is invoked as `winapp2ool -uwpbuilder`. It has no mode toggles.

### File Selection

| Arg | Effect | Default Value |
|:-|:-|:-|
| `-1d path` | Set the source directory (containing `UWP.ini` and `AppInfo\`) | Current directory |
| `-2d path` | Set the output directory | Current directory |
| `-2f name` | Set the output file name | `uwp.ini` |
| `-3d path` / `-3f name` | Set the WebView scaffold catalog location | Current directory / `webview.ini` |
| `-4d path` / `-4f name` | Set the QtWebEngine scaffold catalog location | Current directory / `qtwebengine.ini` |

Paths may be absolute, or relative to the working directory using a leading backslash (e.g. `-1d \UWP`).

###### Note: If the catalogs are left at their defaults and no `webview.ini`/`qtwebengine.ini` exists in the working directory, generation still succeeds with a warning and zero scaffold FileKeys for the affected family. When scripting, always pass `-3d`/`-4d` explicitly.

### Examples

| Command | Effect |
|:-|:-|
| `winapp2ool -uwpbuilder -1d \UWP -2f uwp.ini -3d \Scaffolds -4d \Scaffolds` | Generate entries from `.\UWP\` into `.\uwp.ini`, reading both catalogs from `.\Scaffolds\` |
| `winapp2ool -s -offline -uwpbuilder -1d \UWP -2d \Entries\UWP -2f uwp.ini -3d \Scaffolds -4d \Scaffolds` | Silent, offline, output staged to the committed artifact location `Entries\UWP\uwp.ini`; this is the invocation used by the winapp2.ini build script |

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
| "Unexpected key type in [{name}]: {key}" | An AppInfo section contains an unsupported key; it is dropped from the generated entry |
| "No Package key in [{name}], skipping" | The AppInfo section has no `Package=` key; the entry is not generated |
| "No LangSecRef or Section in [{name}], skipping" | The AppInfo section has no category key; the entry is not generated |
| "Both LangSecRef and Section present in [{name}], using LangSecRef" | The AppInfo section declares both category keys; the entry is generated with `LangSecRef` and the `Section` value is discarded |
| "Skipping scaffold FileKeys for: {name}" | Informational; the entry declares `SkipUWPFileKeys` |
| "WebViewScaffold catalog at {path} is empty or missing" | The catalog file wasn't found; generation continues with zero WebView scaffold FileKeys. |
| "Unknown WebView scaffold '{name}' requested by [{entry}], skipping" | A `WebViewScaffolds=` name doesn't exist in the catalog. Note that unknown `ExcludeWebViewScaffolds=` names produce **no** message |
| "WebViewScaffolds=All in [{entry}] with redundant additional names ({names}); ignoring" | `All` already selects everything |
| "Both WebViewScaffolds and ExcludeWebViewScaffolds set in [{entry}]; applying exclusions to explicit list" | Exclusions usually pair with the implicit defaults or `All` |
| "Duplicate WebViewScaffold name '{name}'; last definition wins" | The catalog defines the same scaffold twice |
| "Unexpected section in WebViewScaffold catalog: [{name}]" | The catalog contains a section without the family prefix; it is ignored |

QtWebEngine-family messages are identical with `QtWebEngine` in place of `WebView`.

---

# Usage Examples

All examples below were generated with real runs of UWPBuilder against the scaffold template and catalogs shipped in `Assembler/`. The source files shown are AppInfo sections; the command for every example is the standard invocation:

```
winapp2ool -uwpbuilder -1d \UWP -2f uwp.ini -3d \Scaffolds -4d \Scaffolds
```

###### Note: The generated header comment block (see [Output Formatting](#output-formatting)) is omitted from the outputs below for brevity.

### Example 1: A Minimal Entry

**Context**

The vast majority of UWP applications store nothing cleanable outside the standard package layout. For these apps, the scaffold baseline is the entire entry.

**Intent**

We want a complete winapp2.ini entry for ASUS Quick Launch by declaring only its package and category.

**Files**

###### **AppInfo section (`AppInfo\A.ini`)**
```ini
[ASUS Quick Launch *]
Package=B9ECED6F.ASUSQuickLaunch_qmba6cd70vzyy
LangSecRef=3024
```

**Output**

```ini
[ASUS Quick Launch *]
LangSecRef=3024
DetectFile=%LocalAppData%\Packages\B9ECED6F.ASUSQuickLaunch_qmba6cd70vzyy
FileKey1=%LocalAppData%\Packages\B9ECED6F.ASUSQuickLaunch_qmba6cd70vzyy\AC|*|RECURSE
FileKey2=%LocalAppData%\Packages\B9ECED6F.ASUSQuickLaunch_qmba6cd70vzyy\Settings|*.log*
FileKey3=%LocalAppData%\Packages\B9ECED6F.ASUSQuickLaunch_qmba6cd70vzyy\SystemAppData\Helium|*.log*
FileKey4=%LocalAppData%\Packages\B9ECED6F.ASUSQuickLaunch_qmba6cd70vzyy\TempState|*|REMOVESELF
```

**Explanation**
- `DetectFile` is generated from the scaffold's `DetectFileBase=%Package%` template
- The four FileKeys are the scaffold's `FileKeyBase=` templates with `%Package%` expanded
- The category key passes through as written
- Nothing in this entry was written by hand except the package name and category

---

### Example 2: App-Specific FileKeys

**Context**

Apple Devices keeps logs in several locations under its package that the baseline scaffold doesn't cover.

**Intent**

We want to extend the baseline with three additional package-relative FileKeys.

**Files**

###### **AppInfo section (`AppInfo\A.ini`)**
```ini
[Apple Devices *]
Package=AppleInc.AppleDevices_nzyj5cx40ttqa
LangSecRef=3024
FileKeyBase=%Package%\LocalCache\Local\Logs|*
FileKeyBase=%Package%\LocalCache\Roaming\Apple Computer\Logs|*
FileKeyBase=%Package%\LocalState\Logs|*
```

**Output**

```ini
[Apple Devices *]
LangSecRef=3024
DetectFile=%LocalAppData%\Packages\AppleInc.AppleDevices_nzyj5cx40ttqa
FileKey1=%LocalAppData%\Packages\AppleInc.AppleDevices_nzyj5cx40ttqa\AC|*|RECURSE
FileKey2=%LocalAppData%\Packages\AppleInc.AppleDevices_nzyj5cx40ttqa\LocalCache\Local\Logs|*
FileKey3=%LocalAppData%\Packages\AppleInc.AppleDevices_nzyj5cx40ttqa\LocalCache\Roaming\Apple Computer\Logs|*
FileKey4=%LocalAppData%\Packages\AppleInc.AppleDevices_nzyj5cx40ttqa\LocalState\Logs|*
FileKey5=%LocalAppData%\Packages\AppleInc.AppleDevices_nzyj5cx40ttqa\Settings|*.log*
FileKey6=%LocalAppData%\Packages\AppleInc.AppleDevices_nzyj5cx40ttqa\SystemAppData\Helium|*.log*
FileKey7=%LocalAppData%\Packages\AppleInc.AppleDevices_nzyj5cx40ttqa\TempState|*|REMOVESELF
```

**Explanation**
- The three app-specific keys are interleaved with the scaffold baseline: the lint pass sorts FileKeys alphabetically by value, so source ordering is lost
- All seven keys were renumbered sequentially after sorting

---

### Example 3: Hybrid UWP + win32 Applications

**Context**

Amazon Music ships as both a UWP package and a win32 installation. A single entry can clean both.

**Intent**

We want package-relative keys (`FileKeyBase=` with `%Package%`) for the UWP install, verbatim win32 keys (`FileKey=`), and a registry detection for the win32 version alongside the generated package detection.

**Files**

###### **AppInfo section (`AppInfo\A.ini`)**
```ini
[Amazon Music *]
Detect=HKCU\Software\Amazon\Amazon Music
LangSecRef=3023
Package=AmazonMobileLLC.AmazonMusic_kc6t79cpj4tp0
FileKeyBase=%Package%\LocalCache\Local\Amazon Music\Data\App Cache\*Cache*|RECURSE
FileKeyBase=%Package%\LocalCache\Local\Amazon Music\Data\Artwork Cache|*.jpg;*.png|RECURSE
FileKeyBase=%Package%\LocalCache\Local\Amazon Music\Data\Hammer Cache|*
FileKeyBase=%Package%\LocalCache\Local\Amazon Music\Data\S*Cache|*
FileKeyBase=%Package%\LocalCache\Local\Amazon Music\Logs|*
FileKeyBase=%Package%\LocalCache\Local\Amazon Music\User Data|*.pma
FileKeyBase=%Package%\LocalCache\Local\Amazon Music\User Data\Crashpad\Reports|*
FileKey=%LocalAppData%\Amazon Music\Crash Dumps|*
FileKey=%LocalAppData%\Amazon Music\data\App Cache|*-journal;ChromeDWriteFontCache;data_*;f_*;index
FileKey=%LocalAppData%\Amazon Music\data\App Cache\*Cache|*|RECURSE
FileKey=%LocalAppData%\Amazon Music\data\Artwork Cache|*.jpg;*.png|RECURSE
FileKey=%LocalAppData%\Amazon Music\data\Hammer Cache|*
FileKey=%LocalAppData%\Amazon Music\data\S*Cache|*
FileKey=%LocalAppData%\Amazon Music\Logs|*
FileKey=%LocalAppData%\Amazon Music\User Data|*.pma
FileKey=%LocalAppData%\Amazon Music\User Data\Crashpad\Reports|*
FileKey=%UserProfile%\.amu\updates|*.exe
FileKey=%UserProfile%\Documents\Amazon Music Importer\Logs|*|RECURSE
```

**Output**

```ini
[Amazon Music *]
LangSecRef=3023
Detect=HKCU\Software\Amazon\Amazon Music
DetectFile=%LocalAppData%\Packages\AmazonMobileLLC.AmazonMusic_kc6t79cpj4tp0
FileKey1=%LocalAppData%\Amazon Music\Crash Dumps|*
FileKey2=%LocalAppData%\Amazon Music\data\App Cache|*-journal;ChromeDWriteFontCache;data_*;f_*;index
FileKey3=%LocalAppData%\Amazon Music\data\App Cache\*Cache|*|RECURSE
FileKey4=%LocalAppData%\Amazon Music\data\Artwork Cache|*.jpg;*.png|RECURSE
FileKey5=%LocalAppData%\Amazon Music\data\Hammer Cache|*
FileKey6=%LocalAppData%\Amazon Music\data\S*Cache|*
FileKey7=%LocalAppData%\Amazon Music\Logs|*
FileKey8=%LocalAppData%\Amazon Music\User Data|*.pma
FileKey9=%LocalAppData%\Amazon Music\User Data\Crashpad\Reports|*
FileKey10=%LocalAppData%\Packages\AmazonMobileLLC.AmazonMusic_kc6t79cpj4tp0\AC|*|RECURSE
FileKey11=%LocalAppData%\Packages\AmazonMobileLLC.AmazonMusic_kc6t79cpj4tp0\LocalCache\Local\Amazon Music\Data\App Cache\*Cache*|RECURSE
FileKey12=%LocalAppData%\Packages\AmazonMobileLLC.AmazonMusic_kc6t79cpj4tp0\LocalCache\Local\Amazon Music\Data\Artwork Cache|*.jpg;*.png|RECURSE
FileKey13=%LocalAppData%\Packages\AmazonMobileLLC.AmazonMusic_kc6t79cpj4tp0\LocalCache\Local\Amazon Music\Data\Hammer Cache|*
FileKey14=%LocalAppData%\Packages\AmazonMobileLLC.AmazonMusic_kc6t79cpj4tp0\LocalCache\Local\Amazon Music\Data\S*Cache|*
FileKey15=%LocalAppData%\Packages\AmazonMobileLLC.AmazonMusic_kc6t79cpj4tp0\LocalCache\Local\Amazon Music\Logs|*
FileKey16=%LocalAppData%\Packages\AmazonMobileLLC.AmazonMusic_kc6t79cpj4tp0\LocalCache\Local\Amazon Music\User Data|*.pma
FileKey17=%LocalAppData%\Packages\AmazonMobileLLC.AmazonMusic_kc6t79cpj4tp0\LocalCache\Local\Amazon Music\User Data\Crashpad\Reports|*
FileKey18=%LocalAppData%\Packages\AmazonMobileLLC.AmazonMusic_kc6t79cpj4tp0\Settings|*.log*
FileKey19=%LocalAppData%\Packages\AmazonMobileLLC.AmazonMusic_kc6t79cpj4tp0\SystemAppData\Helium|*.log*
FileKey20=%LocalAppData%\Packages\AmazonMobileLLC.AmazonMusic_kc6t79cpj4tp0\TempState|*|REMOVESELF
FileKey21=%UserProfile%\.amu\updates|*.exe
FileKey22=%UserProfile%\Documents\Amazon Music Importer\Logs|*|RECURSE
```

**Explanation**
- `Detect=` passes through verbatim, providing win32 detection; the package `DetectFile` is generated alongside it
- The win32 `FileKey=` values contain no `%Package%` variable and pass through verbatim
- One entry covers both installations

---

### Example 4: Multi-Package Applications

**Context**

Cortana has shipped under two different package identities. Some data locations exist under both packages, others only under the newer one.

**Intent**

We want one entry covering both packages: shared keys should expand for every package (`%Package%`), package-specific keys should target only one (`%Package2%`).

**Files**

###### **AppInfo section (`AppInfo\C.ini`)**
```ini
[Cortana *]
Package1=Microsoft.549981C3F5F10_8wekyb3d8bbwe
Package2=Microsoft.Cortana_8wekyb3d8bbwe
LangSecRef=3025
FileKeyBase=%Package%\LocalCache|*|RECURSE
FileKeyBase=%Package2%\LocalState|*|RECURSE
ExcludeKey1=FILE|%Package2%\LocalState\DeviceSearchCache\|SettingsCache.txt
```

**Output**

```ini
[Cortana *]
LangSecRef=3025
DetectFile1=%LocalAppData%\Packages\Microsoft.549981C3F5F10_8wekyb3d8bbwe
DetectFile2=%LocalAppData%\Packages\Microsoft.Cortana_8wekyb3d8bbwe
FileKey1=%LocalAppData%\Packages\Microsoft.549981C3F5F10_8wekyb3d8bbwe\AC|*|RECURSE
FileKey2=%LocalAppData%\Packages\Microsoft.549981C3F5F10_8wekyb3d8bbwe\LocalCache|*|RECURSE
FileKey3=%LocalAppData%\Packages\Microsoft.549981C3F5F10_8wekyb3d8bbwe\Settings|*.log*
FileKey4=%LocalAppData%\Packages\Microsoft.549981C3F5F10_8wekyb3d8bbwe\SystemAppData\Helium|*.log*
FileKey5=%LocalAppData%\Packages\Microsoft.549981C3F5F10_8wekyb3d8bbwe\TempState|*|REMOVESELF
FileKey6=%LocalAppData%\Packages\Microsoft.Cortana_8wekyb3d8bbwe\AC|*|RECURSE
FileKey7=%LocalAppData%\Packages\Microsoft.Cortana_8wekyb3d8bbwe\LocalCache|*|RECURSE
FileKey8=%LocalAppData%\Packages\Microsoft.Cortana_8wekyb3d8bbwe\LocalState|*|RECURSE
FileKey9=%LocalAppData%\Packages\Microsoft.Cortana_8wekyb3d8bbwe\Settings|*.log*
FileKey10=%LocalAppData%\Packages\Microsoft.Cortana_8wekyb3d8bbwe\SystemAppData\Helium|*.log*
FileKey11=%LocalAppData%\Packages\Microsoft.Cortana_8wekyb3d8bbwe\TempState|*|REMOVESELF
ExcludeKey1=FILE|%LocalAppData%\Packages\Microsoft.Cortana_8wekyb3d8bbwe\LocalState\DeviceSearchCache\|SettingsCache.txt
```

**Explanation**
- One `DetectFile` was generated per package from the scaffold's `DetectFileBase=%Package%`
- The scaffold baseline and the unnumbered `%Package%` FileKey expanded once per package: keys 1–5 for `Package1`, keys 6–7 and 9–11 for `Package2`
- The `%Package2%` keys (FileKey8's `LocalState`, and the ExcludeKey) expanded exactly once, targeting only the second package

---

### Example 5: WebView Scaffolds (Default Selection)

**Context**

GoldKey Acellus is an education app that embeds WebView2 and keeps its browser data under `LocalState\EBWebView`.

**Intent**

We want the entry to draw the default scaffold selection (Caches + Telemetry) from the shared catalog by declaring only the WebView root.

**Files**

###### **AppInfo section**
```ini
[GoldKey Acellus *]
Package=GoldKeyCorporation.Acellus_v2rn50m687qry
WebViewRoot=%Package%\LocalState\EBWebView
LangSecRef=3021
```

###### Note: The entry for this app shipped in the winapp2.ini project selects `WebViewScaffolds=All` with exclusions — that variant is [Example 6](#example-6-webview-scaffolds-all-with-exclusions). This example omits the selector keys to demonstrate the default behavior.

**Output**

```ini
[GoldKey Acellus *]
LangSecRef=3021
DetectFile=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry
FileKey1=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\AC|*|RECURSE
FileKey2=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView|*.log;*_shutdown_ms.txt;Breadcrumbs;BrowsingTopics*;Last Browser;Last Version;Module Info Cache
FileKey3=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView|*.pma;LOG;LOG.old;*-journal|RECURSE
FileKey4=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView\*BrowserMetrics|*|REMOVESELF
FileKey5=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView\*Cache*|*|REMOVESELF
FileKey6=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView\Avatars|*|REMOVESELF
FileKey7=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView\Crash Reports|*|REMOVESELF
FileKey8=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView\Crashpad|*|REMOVESELF
FileKey9=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView\Default|*.ldb;CURRENT;LOCK;MANIFEST-*;ServerCertificate;*.log
FileKey10=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView\Default\*Cache*|*|REMOVESELF
FileKey11=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView\Default\blob_storage|*|REMOVESELF
FileKey12=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView\Default\BudgetDatabase|*|REMOVESELF
FileKey13=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView\Default\DataSharing|*|REMOVESELF
FileKey14=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView\Default\Download Service|*|REMOVESELF
FileKey15=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView\Default\Feature Engagement Tracker|*|REMOVESELF
FileKey16=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView\Default\File System|*|REMOVESELF
FileKey17=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView\Default\GCM Store|*|REMOVESELF
FileKey18=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView\Default\JumpListIcons*|*|REMOVESELF
FileKey19=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView\Default\Network|Network Persistent State*;Reporting and NEL*;SCT Auditing Pending Reports*
FileKey20=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView\Default\optimization*|*|REMOVESELF
FileKey21=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView\Default\PersistentOriginTrials|*|REMOVESELF
FileKey22=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView\Default\Platform Notifications|*|REMOVESELF
FileKey23=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView\Default\Service Worker|*|REMOVESELF
FileKey24=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView\Default\Shared Dictionary\cache|*|REMOVESELF
FileKey25=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView\Default\Site Characteristics Database|*|REMOVESELF
FileKey26=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView\Default\Storage\ext\*\def\*cache*|*|REMOVESELF
FileKey27=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView\Default\Storage\ext\*\def\Platform Notifications|*|REMOVESELF
FileKey28=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView\Default\VideoDecodeStats|*|REMOVESELF
FileKey29=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView\Default\WebRTC Logs|*|REMOVESELF
FileKey30=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView\Default\WebrtcVideoStats|*|REMOVESELF
FileKey31=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView\Local Traces|*|REMOVESELF
FileKey32=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView\Optimization*|*|REMOVESELF
FileKey33=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView\OriginTrials|*|REMOVESELF
FileKey34=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView\Stability|*|REMOVESELF
FileKey35=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\Settings|*.log*
FileKey36=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\SystemAppData\Helium|*.log*
FileKey37=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\TempState|*|REMOVESELF
```

**Explanation**
- Declaring `WebViewRoot=` opted the entry into scaffold emission
- With no `WebViewScaffolds=` key, the default Caches + Telemetry selection applied
- The catalog templates had `%WebViewRoot%` substituted with the package-expanded root path
- 33 WebView FileKeys were generated from a single declaration line
- The lint pass merged scaffold templates targeting the same folder (e.g. FileKey9 combines a Caches template and a Telemetry template both targeting `\Default`)

---

### Example 6: WebView Scaffolds (All With Exclusions)

**Context**

For a dedicated single-purpose app like Acellus, the browsing data inside the WebView (history, cookies, sessions) can be safely deleted, but wiping `WebStorage` or `LoginData` would log the user out or destroy locally-stored user data.

**Intent**

We want every scaffold in the catalog *except* the two that remove primary user state. This is the variant of this entry actually shipped by the winapp2.ini project.

**Files**

###### **AppInfo section (`AppInfo\G.ini`)**
```ini
[GoldKey Acellus *]
Package=GoldKeyCorporation.Acellus_v2rn50m687qry
WebViewRoot=%Package%\LocalState\EBWebView
WebViewScaffolds=All
ExcludeWebViewScaffolds=WebStorage,LoginData
LangSecRef=3021
```

**Output**

The full output carries **68 FileKeys**; the excerpt below shows the shape, with the middle trimmed for brevity:

```ini
[GoldKey Acellus *]
LangSecRef=3021
DetectFile=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry
FileKey1=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\AC|*|RECURSE
FileKey2=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView|*.log;*_shutdown_ms.txt;Breadcrumbs;BrowsingTopics*;Last Browser;Last Version;*first_party_sets*;Module Info Cache
FileKey3=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView|*.pma;LOG;LOG.old;*-journal|RECURSE
FileKey4=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView\*BrowserMetrics|*|REMOVESELF
FileKey5=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView\*Cache*|*|REMOVESELF
FileKey6=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView\AutoFill*|*|REMOVESELF
FileKey7=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView\Avatars|*|REMOVESELF
FileKey8=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView\CertificateRevocation|*|REMOVESELF
FileKey9=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView\CookieReadinessList|*|REMOVESELF
FileKey10=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView\Crash Reports|*|REMOVESELF
FileKey11=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView\Crashpad|*|REMOVESELF
FileKey12=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView\Crowd Deny|*|REMOVESELF
FileKey13=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView\Default|*.ldb;CURRENT;LOCK;MANIFEST-*;ServerCertificate;*.log;*Web Data;Bookmarks.bak;BrowsingTopics*;Conversions*;InterestGroups;MediaDeviceSalts;PrivateAggregation*;SharedStorage*;DIPS*;DownloadMetadata;Extension Cookies;favicons*;History*;Network Action Predictor*;shortcuts*;Top Sites*;Visited Links*;PreferredApps;QuotaManager

; ... FileKey14 through FileKey65: the remaining Autofill, Autoplay, BookmarkBackups, DownloadHistory,
; DRMData, ProgressiveWebApps, PrivacySandbox, Security, Shopping, WebCookies, WebHistory and
; WebSession scaffold keys, plus the rest of Caches and Telemetry ...

FileKey66=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\Settings|*.log*
FileKey67=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\SystemAppData\Helium|*.log*
FileKey68=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\TempState|*|REMOVESELF
```

**Explanation**
- `WebViewScaffolds=All` expanded to every scaffold in the catalog; `ExcludeWebViewScaffolds=WebStorage,LoginData` then removed those two from the selection
- Compare FileKey13 with Example 5's FileKey9: the same `\Default` folder now aggregates parameters from many more scaffolds, merged into one key by the lint pass

---

### Example 7: Secondary Entries With SkipUWPFileKeys

**Context**

Spotify's downloaded songs occupy significant disk space, but users won't necessarily expect a cleaner to remove their offline music. This warrants a *separate* entry so the two cleaning surfaces can be toggled independently, but a second entry for the same package would duplicate the scaffold baseline.

**Intent**

We want a primary `[Spotify *]` entry with full coverage, and a secondary `[Spotify - Downloaded Songs *]` entry that shares the package but carries only its own keys.

**Files**

###### **AppInfo sections (`AppInfo\S.ini`)**
```ini
[Spotify *]
Package=SpotifyAB.SpotifyMusic_zpdnekdrzrea0
LangSecRef=3023
DetectFile=%AppData%\Spotify
FileKeyBase=%Package%\LocalCache\Spotify|*|REMOVESELF
FileKeyBase=%Package%\LocalState\Users\*|ad-state-storage.bnk;context_player_state_restore;cached
FileKey=%AppData%\Spotify|*.log
FileKey=%AppData%\Spotify\users\*|ad-state-storage.bnk;frequency.pb;local-files.bnk;local_sync_state;watch-sources.bnk
FileKey=%LocalAppData%\Spotify|*|REMOVESELF
ExcludeKey1=PATH|%LocalAppData%\Spotify\Storage\*\|*

[Spotify - Downloaded Songs *]
Package=SpotifyAB.SpotifyMusic_zpdnekdrzrea0
SkipUWPFileKeys=True
LangSecRef=3023
DetectFile=%AppData%\Spotify
FileKeyBase=%Package%\LocalState\Spotify\Storage|*|REMOVESELF
FileKey=%LocalAppData%\Spotify\Storage|*|REMOVESELF
```

**Output**

```ini
[Spotify - Downloaded Songs *]
LangSecRef=3023
DetectFile1=%AppData%\Spotify
DetectFile2=%LocalAppData%\Packages\SpotifyAB.SpotifyMusic_zpdnekdrzrea0
FileKey1=%LocalAppData%\Packages\SpotifyAB.SpotifyMusic_zpdnekdrzrea0\LocalState\Spotify\Storage|*|REMOVESELF
FileKey2=%LocalAppData%\Spotify\Storage|*|REMOVESELF

[Spotify *]
LangSecRef=3023
DetectFile1=%AppData%\Spotify
DetectFile2=%LocalAppData%\Packages\SpotifyAB.SpotifyMusic_zpdnekdrzrea0
FileKey1=%AppData%\Spotify|*.log
FileKey2=%AppData%\Spotify\users\*|ad-state-storage.bnk;frequency.pb;local-files.bnk;local_sync_state;watch-sources.bnk
FileKey3=%LocalAppData%\Packages\SpotifyAB.SpotifyMusic_zpdnekdrzrea0\AC|*|RECURSE
FileKey4=%LocalAppData%\Packages\SpotifyAB.SpotifyMusic_zpdnekdrzrea0\LocalCache\Spotify|*|REMOVESELF
FileKey5=%LocalAppData%\Packages\SpotifyAB.SpotifyMusic_zpdnekdrzrea0\LocalState\Users\*|ad-state-storage.bnk;cached;context_player_state_restore
FileKey6=%LocalAppData%\Packages\SpotifyAB.SpotifyMusic_zpdnekdrzrea0\Settings|*.log*
FileKey7=%LocalAppData%\Packages\SpotifyAB.SpotifyMusic_zpdnekdrzrea0\SystemAppData\Helium|*.log*
FileKey8=%LocalAppData%\Packages\SpotifyAB.SpotifyMusic_zpdnekdrzrea0\TempState|*|REMOVESELF
FileKey9=%LocalAppData%\Spotify|*|REMOVESELF
ExcludeKey1=PATH|%LocalAppData%\Spotify\Storage\*\|*
```

**Explanation**
- `SkipUWPFileKeys` suppressed only the four scaffold FileKeys on the secondary entry; its detection was still generated (both entries carry the same two DetectFiles)
- The primary entry keeps full baseline coverage but excludes the `Storage` folder, which belongs to the secondary entry.
- Note the entries are alphabetized in the output, and FileKey5's parameter list was alphabetized by the lint pass

---

### Example 8: Validation Warnings and Skips

**Context**

AppInfo sections are validated before generation. It's worth seeing what rejection actually looks like, and what happens to entries that only *partially* misbehave.

**Intent**

We run three deliberately flawed sections through the builder: one with no `Package=`, one declaring both category keys plus an unsupported key, and one carrying `Skip=False`.

**Files**

###### **AppInfo section (`AppInfo\src.ini`)**
```ini
[Hazel Notes *]
LangSecRef=3021
FileKeyBase=%Package%\LocalState\Logs|*

[Hazel Weather *]
Package=Hazel.Weather_123abc
LangSecRef=3021
Section=Hazel Apps
Author=Hazel

[Hazel Legacy *]
Package=Hazel.Legacy_123abc
LangSecRef=3021
Skip=False
```

**Warnings**

The run reports:

```
No Package key in [Hazel Notes *], skipping
Unexpected key type in [Hazel Weather *]: Author
Both LangSecRef and Section present in [Hazel Weather *], using LangSecRef
```

**Output**

```ini
[Hazel Weather *]
LangSecRef=3021
DetectFile=%LocalAppData%\Packages\Hazel.Weather_123abc
FileKey1=%LocalAppData%\Packages\Hazel.Weather_123abc\AC|*|RECURSE
FileKey2=%LocalAppData%\Packages\Hazel.Weather_123abc\Settings|*.log*
FileKey3=%LocalAppData%\Packages\Hazel.Weather_123abc\SystemAppData\Helium|*.log*
FileKey4=%LocalAppData%\Packages\Hazel.Weather_123abc\TempState|*|REMOVESELF
```

**Explanation**
- `[Hazel Notes *]` was skipped entirely because it has no `Package=` key
- `[Hazel Weather *]` was generated with warnings: the conflicting `Section=` was discarded in favor of `LangSecRef`, and the unsupported `Author=` key was dropped entirely
- `[Hazel Legacy *]` was skipped **silently**: `Skip` is presence-based, so `Skip=False` skips just like `Skip=True`
- Only one of the three sections produced an entry

---