# UWPBuilder

**UWPBuilder** is a winapp2ool build-chain module for winapp2.ini maintainers that generates winapp2.ini entries for Universal Windows Platform (UWP) apps from a few lines of description each. Every UWP app keeps its data under `%LocalAppData%\Packages\PACKAGE_FOLDER`, and part of that folder looks the same from app to app, so UWPBuilder gives every entry the same set of baseline cleaning targets.

Its output, `uwp.ini`, is an intermediate file that winapp2ool combines into the published winapp2.ini. It isn't meant for use with any cleaner on its own, and unless you maintain winapp2.ini you probably don't need this module.

### What does UWPBuilder do?
UWPBuilder reads a source directory holding a scaffold template, `UWP.ini`, and app definitions split into per-letter files under `AppInfo\`. It writes one winapp2.ini entry for each app. Every entry gets package detection and the baseline FileKeys automatically, and the app's own keys go on top. The `<Package>` variable stands in for the package folder, so you never have to write a package path out in full.

### Why UWPBuilder?
A package name and a category make a complete entry. UWPBuilder fills in the detection and the baseline FileKeys from one template, the same way for every app.

Apps that ship both a UWP package and a win32 install keep both sets of keys in one definition. Apps that embed a browser engine (WebView2, QtWebEngine or Electron) can take their cleaning keys from the shared scaffold catalogs.

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

15. **[Examples](EXAMPLES.md)**: eight worked examples with real input and output

---

# Requirements

- A source directory holding:
  - `UWP.ini`, the scaffold template
  - An `AppInfo\` subdirectory of `*.ini` files that define the apps
- Optionally, a scaffold directory holding the shared catalogs, read when an entry opts into embedded-browser scaffolds. UWPBuilder reads every `*.ini` in it and takes each catalog's engine family from its section headers, not its file name:
  - `webview.ini`: `[WebViewScaffold: ...]`, the WebView2/EBWebView catalog
  - `qtwebengine.ini`: `[QtWebEngineScaffold: ...]`, the QtWebEngine catalog
  - `electron.ini`: `[ElectronScaffold: ...]`, the Electron catalog

The winapp2.ini project keeps its source files in [`Assembler/UWP/`](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/UWP) (source directory) and [`Assembler/Scaffolds/`](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/Scaffolds) (scaffold directory).

###### Note: You only need the scaffold directory if an AppInfo entry declares a scaffold root. If the directory doesn't exist, UWPBuilder warns `Scaffold directory not found` and carries on without scaffold keys. If it exists but has no catalog for a family, every entry using that family warns that its scaffolds are unknown. The run succeeds either way.

---

# Quick Start

### Common Workflow
1. Point the source directory at the folder holding `UWP.ini` and `AppInfo\`
2. Point the scaffolds directory at the shared `Scaffolds\` folder
3. Choose **Run**

To add an app, add a section to the matching `AppInfo\{letter}.ini` file.

---

# Common Uses

| To | Run | Example |
|:-|:-|:-|
| Generate `uwp.ini` from the sources in `UWP\`, using the shared scaffold catalogs | `winapp2ool -uwpbuilder -1d \UWP -2f uwp.ini -3d \Scaffolds` | [A Minimal Entry](EXAMPLES.md#example-1-a-minimal-entry) |
| Write the committed artifact silently and offline, as the build script does | `winapp2ool -s -offline -uwpbuilder -1d \UWP -2d \Entries\UWP -2f uwp.ini -3d \Scaffolds` | |

Every worked example runs the first command. Only the AppInfo section it reads changes. In the menu, set **Choose source directory**, **Choose save target** and **Choose scaffolds directory**, then choose **Run**; see [Menu Options](#menu-options). `-s` and `-offline` are winapp2ool [global args](../../Readme.md#global-args).

---

# Menu Options

| Option | Effect | Notes |
|:-|:-|:-|
| Run (default) | Generate entries for the apps in the source directory | |
| Choose source directory | Select the directory holding `UWP.ini` and `AppInfo\` | Default: current directory |
| Choose save target | Select where to save the generated entries | Default: `uwp.ini` |
| Choose scaffolds directory | Select the folder holding the shared scaffold catalogs | Default: current directory |
| Reset Settings | Restore all settings to their defaults | Only shown when settings have been changed |

Below the options, the menu shows the current source directory, save target and scaffolds directory.

---

# Command-Line Arguments

UWPBuilder is invoked as `winapp2ool -uwpbuilder`. It takes three directory arguments and an output file name, and no toggles.

### File Selection

| Arg | Effect | Default Value |
|:-|:-|:-|
| `-1d path` | Set the source directory (holding `UWP.ini` and `AppInfo\`) | Current directory |
| `-2d path` | Set the directory the output is saved to | Current directory |
| `-2f name` | Set the output file name | `uwp.ini` |
| `-3d path` | Set the shared scaffold directory | Current directory |

Paths can be absolute, or relative to the working directory with a leading backslash (`-1d \UWP`).

###### Note: Slot 3 only reads a directory, so `-3f` does nothing there. It defaults to the current directory, and if that holds no catalogs the run still succeeds, minus every scaffold key, with only an `Unknown ... scaffold` warning for each entry that asked for them. Always pass `-3d` when scripting.

---

# How Generation Works

| Stage | What happens |
|:-|:-|
| Combine | Every `AppInfo\*.ini` file is combined in memory, in alphabetical path order. No intermediate file is written |
| Load scaffolds | The `[EntryScaffold: ...]` sections are read from `UWP.ini`, and every catalog in the scaffolds directory is loaded |
| Validate | Each AppInfo section needs at least one `Package=` and a category key. A section missing either is skipped with a warning. A section with *both* category keys warns but is still generated, using `LangSecRef` |
| Generate | Each valid section becomes one entry: package detection, the baseline FileKeys, the app's own keys, and any WebView, QtWebEngine or Electron scaffold keys it selected, with `<Package>` and every other variable expanded |
| Lint | WinappDebug normalizes the whole output with Optimizations forced on (see [Output Formatting](#output-formatting)) |
| Write | The output file is saved |

###### Note: When two AppInfo files define a section with the same name, the file that sorts first wins and the other section is dropped with a warning

---

# Source Files

## The source directory

UWPBuilder reads everything from one source directory:

```
UWP\
├── UWP.ini            ← scaffold template
└── AppInfo\
    ├── A.ini          ← app definitions, split alphabetically
    ├── B.ini
    └── ...
```

Splitting `AppInfo\` by letter is only a convention to keep the files manageable. UWPBuilder combines every `*.ini` file in the folder, whatever it's called.

## UWP.ini: the entry scaffold

`UWP.ini` holds `[EntryScaffold: ...]` sections, the baseline every generated entry gets. They take two keys:

| Key | Effect |
|:-|:-|
| `DetectFileBase=` | A DetectFile template, expanded once per package into the entry's detection paths |
| `FileKeyBase=` | A FileKey template every app gets as a baseline cleaning target |

The scaffold the winapp2.ini project ships:

```ini
[EntryScaffold: UWP App]
DetectFileBase=<Package>
FileKeyBase=<Package>\AC|*|RECURSE
FileKeyBase=<Package>\Settings|*.log*
FileKeyBase=<Package>\SystemAppData\Helium|*.log*
FileKeyBase=<Package>\TempState|*|REMOVESELF
```

## AppInfo files

Each section in an AppInfo file describes one winapp2.ini entry. The section header becomes the entry's name exactly as written, so it has to follow winapp2.ini's naming rules and end in ` *]`:

```ini
[ASUS Quick Launch *]
Package=B9ECED6F.ASUSQuickLaunch_qmba6cd70vzyy
LangSecRef=3024
```

Comment the source files as much as you like. Comments never reach the output.

---

# AppInfo Key Reference

## Standard winapp2.ini keys

These keys mean what they always mean in winapp2.ini, and they're passed on to the generated entry:

| Key | Behavior | Notes |
|:-|:-|:-|
| `LangSecRef=` / `Section=` | Category of the generated entry | Exactly one is required. With both, UWPBuilder warns and uses `LangSecRef` |
| `Detect=` | Registry detection | Can repeat. Unnumbered in the output when there's only one. `<>`-expanded |
| `DetectFile=` | File system detection, added after the detection paths the scaffold generates | Can repeat. `<>`-expanded, `<Package>` included. For hybrid apps that need win32 detection beside the package detection |
| `DetectOS=` | Windows kernel version filter, e.g. `6.1\|10.0` | Written out as is, never expanded |
| `FileKey=` | Cleaning target; same as `FileKeyBase=` | [Root placeholders](#root-placeholders-in-appinfo-keys) filled in, then `<>`-expanded, `<Package>` included. A key with no token passes through as is |
| `RegKey=` | Registry cleaning target | Renumbered from 1. [Root placeholders](#root-placeholders-in-appinfo-keys) filled in, then `<>`-expanded |
| `ExcludeKey=` | Exclusion; same as `ExcludeKeyBase=` | [Root placeholders](#root-placeholders-in-appinfo-keys) filled in, then `<>`-expanded, `<Package>` included, in the domain named by the flag of the filled-in value |
| `Warning=` | Warning shown to the user | Written out as is, never expanded |

###### Note: Key numbers in the source are ignored. UWPBuilder numbers every key itself, and the lint pass numbers them again.

`SpecialDetect=` is deprecated and dropped with a warning. `Default=` is ignored with a warning, since UWPBuilder never writes one.

Any other key name declares a [variable](#variables), so an unfamiliar key is never an error on its own.

## Winapp2ool-exclusive keys

| Key | Effect | Notes |
|:-|:-|:-|
| `Package=` | The app's package folder name under `%LocalAppData%\Packages` | At least one is required. Can repeat as `Package1=`, `Package2=`, ... for an app with more than one package |
| `FileKeyBase=` | A FileKey template; same as `FileKey=` | By convention, `FileKeyBase=` for `<Package>` paths and `FileKey=` for win32 paths |
| `ExcludeKeyBase=` | An ExcludeKey template; same as `ExcludeKey=` | |
| `RegKeyBase=` | A RegKey template; same as `RegKey=` | |
| `Skip` | Leaves this app out of the output | Takes effect whenever the key is there; see below |
| `SkipUWPFileKeys` | Leaves out the scaffold `FileKeyBase=` templates for this entry only | Takes effect whenever the key is there; see below. Detection is unaffected |
| `WebViewRoot=` (alias `WebViewPath=`) | Declares an embedded WebView2/EBWebView data root, opting the entry into [WebView scaffolds](#webview-scaffolds) | Can repeat, numbered or not. `<Package>` is expanded |
| `WebViewScaffolds=` | Comma-separated list of scaffolds to use instead of the default set | `All` selects the whole catalog except the legacy tier |
| `ExcludeWebViewScaffolds=` | Comma-separated scaffold names taken out of the active selection | |
| `QtWebEngineRoot=` (alias `QtWebEnginePath=`) | Declares one embedded QtWebEngine profile directory, profile folder included (e.g. `...\QtWebEngine\Default`), opting the entry into [QtWebEngine scaffolds](#qtwebengine-scaffolds) | Can repeat, one per profile. `<Package>` is expanded |
| `QtWebEngineCacheRoot=` (alias `QtWebEngineCachePath=`) | Declares a QtWebEngine HTTP cache directory for `<QtWebEngineCacheRoot>`. Also opts the entry into QtWebEngine scaffolds | Can repeat. `<Package>` is expanded |
| `QtWebEngineScaffolds=` | Scaffold selection for the QtWebEngine family | Same rules as `WebViewScaffolds=` |
| `ExcludeQtWebEngineScaffolds=` | Exclusions for the QtWebEngine family | |
| `ElectronRoot=` (alias `ElectronPath=`) | Declares an Electron `userData` folder, opting the entry into [Electron scaffolds](#electron-scaffolds) | Can repeat. `<Package>` is expanded |
| `ElectronUpdaterRoot=` (alias `ElectronUpdaterPath=`) | Declares an electron-updater download cache for `<ElectronUpdaterRoot>`. Also opts the entry into Electron scaffolds | Can repeat. `<Package>` is expanded |
| `ElectronScaffolds=` | Scaffold selection for the Electron family | Same rules as `WebViewScaffolds=` |
| `ExcludeElectronScaffolds=` | Exclusions for the Electron family | |

## Package variables

`<Package>` is a [variable](#variables) UWPBuilder declares on every entry, with one value for each `Package=` key: `%LocalAppData%\Packages\PACKAGE_FOLDER`. It works in any key that gets expanded, including the root keys (`WebViewRoot=` and the rest) and your own variable declarations:

- **`<Package>`** expands once per package: a key that uses it becomes one key for each package, in the order they're declared
- **`<Package[N]>`** picks one package by position: a key using `<Package[2]>` becomes a single key for the second package. An `N` of 0, or past the last package, drops the key with a warning
- A key with no package variable passes through as is

`<Package>` and `<Package[2]>` in the same key are two separate variables, so they multiply like any other two. So do `<Package>` and a variable declared from it (`Cache=<Package>\LocalCache`): a key using both pairs each package with each package's cache, mismatched pairs included. Use one or the other in a key. [Example 4](EXAMPLES.md#example-4-multi-package-applications) shows both forms.

The scaffold's `DetectFileBase=` already detects every package, so an entry never needs its own `DetectFile=<Package>`.

### The retired `%Package%` and `%WebViewRoot%` spellings

Older versions spelled these `%Package%`, `%Package1%`, `%Package2%` and so on, and the [root placeholders](#root-placeholders-in-appinfo-keys) `%WebViewRoot%`, `%QtWebEngineRoot%`, `%QtWebEngineCacheRoot%`, `%ElectronRoot%` and `%ElectronUpdaterRoot%`. UWPBuilder drops a key still written that way, including one missing its closing `%` (`%Package\LocalCache`, `%ElectronRoot\logs`), warns with the new spelling, and fails the run with exit code 1:

```
`%WebViewRoot%` in [Test App *].FileKeyBase is now spelled `<WebViewRoot>`; key dropped
```

`Package=` and `Warning=` values aren't checked. The shared scaffold catalogs refuse the same spellings when they load: a catalog template holding one is dropped (`...; template dropped`) and the run fails.

## Root placeholders in AppInfo keys

An app's own `FileKey=` / `FileKeyBase=`, `RegKey=` / `RegKeyBase=` and `ExcludeKey=` / `ExcludeKeyBase=` values take the same root placeholders the scaffold catalogs do: `<WebViewRoot>`, `<QtWebEngineRoot>`, `<QtWebEngineCacheRoot>`, `<ElectronRoot>` and `<ElectronUpdaterRoot>`, in any case. Each one is replaced with every path the entry declared for it (`WebViewRoot=` / `WebViewPath=` and so on) before variables are expanded. A key fans out once per path, and a `<Package>` inside a path expands per package along with the rest of the key:

```ini
[Hazel Viewer *]
Package=Hazel.Viewer_abc123
LangSecRef=3021
WebViewPath1=<Package>\LocalState
WebViewPath2=%LocalAppData%\Hazel Viewer
FileKeyBase=<WebViewRoot>\EBWebView\*\Cache|*
```

generates `%LocalAppData%\Packages\Hazel.Viewer_abc123\LocalState\EBWebView\*\Cache|*` and `%LocalAppData%\Hazel Viewer\EBWebView\*\Cache|*`, beside the baseline FileKeys and the default WebView scaffolds for both paths. A key naming a root the entry has no path for is dropped without a warning, the same as a catalog template. An ExcludeKey's flag is read after the paths are filled in. `Detect=` and `DetectFile=` values never get root placeholders filled in, so a placeholder there is just an undeclared token.

## Variables

Any key whose name UWPBuilder doesn't recognize declares a variable: a comma-separated list of values, used elsewhere in the entry as `<Name>`. Each use fans the key out into one key per value.

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

It's the same engine EntryBuilder uses:

- Two different variables in one template produce the **cartesian product** of their values
- The same variable used twice in one template **co-varies** instead of multiplying
- An **inline list** needs no declaration: `<Default,OffTheRecord>`
- Variables can use other variables; UWPBuilder flattens the declarations before generating anything
- Names match in any case

### Undeclared tokens

What happens to a `<Token>` nothing declares depends on the key it's in:

| Domain | Keys | Behavior |
|:-|:-|:-|
| Registry | `Detect=`, `RegKey=`/`RegKeyBase=`, `REG` `ExcludeKey=` | Kept as written, with an advisory in the log |
| File system | `DetectFile=`, `FileKey=`/`FileKeyBase=`, all other `ExcludeKey=` | Dropped, with a warning |

### What is never expanded

- **`Package=`**: expanding it could change how many packages there are and quietly shift what every `<Package[N]>` in the entry points at. A `<` or `>` in a `Package=` value gets a warning and is used as written, and so is the `Package` variable built from it
- **`Warning=`** and **`DetectOS=`**: one is prose and the other a version range, and neither is a path

### Typo backstops

Since any unknown key name declares a variable, misspelling a key name doesn't raise an error. Two checks catch it instead:

1. An undeclared-token warning catches a misspelled *use* (`<Verison>` when the declaration is `Version=`)
2. An unused-variable warning catches a misspelled *declaration*: after generating, UWPBuilder reports every variable that no `<Token>` uses as a possible typo. [Example 8](EXAMPLES.md#example-8-validation-warnings-and-skips) shows one

###### Note: Like every other diagnostic, these don't show during a scripted build. See [Troubleshooting](#troubleshooting).

## Skip and SkipUWPFileKeys

Both keys take effect whenever they're present. The value doesn't matter, so

```ini
Skip=False
```

still skips the entry. Add the key only when you want its effect, and delete it to turn the effect off.

- `Skip`: leaves the app out of the output. Use it to drop support for an app without deleting its definition
- `SkipUWPFileKeys`: leaves out the baseline `FileKeyBase=` templates from `UWP.ini` for this entry, and generates everything else as usual. See [Example 7](EXAMPLES.md#example-7-secondary-entries-with-skipuwpfilekeys)

---

# WebView Scaffolds

Many apps embed WebView2, a Chromium-based browser whose data folder (usually named `EBWebView`) has a standard layout. Instead of hand-writing the same cleaning keys for each of them, UWPBuilder takes them from a shared catalog of `[WebViewScaffold: ...]` sections, the same catalog [EntryBuilder](../entrybuilder/readme.md) uses.

## Opting in

An entry opts in by declaring at least one `WebViewRoot=` key. An entry without one gets no WebView keys, whatever other WebView keys it has:

```ini
[GoldKey Acellus *]
Package=GoldKeyCorporation.Acellus_v2rn50m687qry
WebViewRoot=<Package>\LocalState\EBWebView
LangSecRef=3021
```

Each selected scaffold's templates are expanded once per declared root, with `<WebViewRoot>` replaced by the root's path after `<Package>` is expanded. The root is the folder that holds the profiles. Templates for profile data use `<WebViewRoot>\*\`, which covers `Default` and named profiles like `WV2Profile_*`. An app with two WebView folders declares two roots (`WebViewRoot1=`, `WebViewRoot2=`) and gets the full selection for each. A root doesn't have to be inside the package: a win32 path like `%AppData%\Microsoft\Teams` works just as well.

## Selecting scaffolds

| Declaration | Resulting selection |
|:-|:-|
| No `WebViewScaffolds=` key | The default set: Caches, Telemetry |
| `WebViewScaffolds=Name1,Name2` | Exactly the named scaffolds |
| `WebViewScaffolds=` (present but empty) | No WebView scaffolds |
| `WebViewScaffolds=All` | Every scaffold in the catalog except the legacy tier, host-risk scaffolds included |
| `WebViewScaffolds=All,LegacyTelemetry` | `All`, plus the named legacy scaffold |
| `ExcludeWebViewScaffolds=Name1,Name2` | Taken out of whichever selection is active |

- `All` is reserved, so no scaffold can be named that. Naming a non-legacy scaffold beside `All` is redundant and warns
- A scaffold marked `Tier=Legacy` in the catalog is left out of `All`, and is only generated when it's named, alone or beside `All`
- `All` with `ExcludeWebViewScaffolds=` means "everything except these" and doesn't warn
- An unknown name in `WebViewScaffolds=` is dropped with a warning, beside `All` too
- An unknown name in `ExcludeWebViewScaffolds=` is ignored without a warning
- Scaffold names match in any case

## The scaffold catalog

The WebView2 catalog's scaffolds are `Autofill`, `Autoplay`, `BookmarkFavicons`, `Caches`, `DefaultApps`, `DownloadHistory`, `DRMData`, `LoginData`, `PrivacySandbox`, `Security`, `Shopping`, `Sync Data`, `Telemetry`, `WebCookies`, `WebHistory`, `WebSession` and `WebStorage`.

The legacy tier holds patterns that modern hosts don't write but older ones may: `LegacyCaches`, `LegacyDownloadHistory`, `LegacyDRMData`, `LegacyExtensionCookies`, `LegacyProgressiveWebApps`, `LegacyStorageQuota`, `LegacyTelemetry` and `LegacyWebHistory`.

Only Caches and Telemetry are generated by default. Everything else has to be selected, and a legacy scaffold has to be named even when `All` is selected.

---

# QtWebEngine Scaffolds

A second family, built the same way, covers apps that embed QtWebEngine. Its catalog is `qtwebengine.ini` in the scaffolds directory.

It works like the WebView family: `QtWebEngineRoot=` (alias `QtWebEnginePath=`) opts in, and `QtWebEngineScaffolds=` / `ExcludeQtWebEngineScaffolds=` select, with the same `All` value and the same warnings. Two things differ.

First, the root is a profile directory, profile folder included: `...\QtWebEngine\Default`, not the `QtWebEngine` folder above it. Qt keeps the HTTP cache outside that directory, so a second root, `QtWebEngineCacheRoot=` (alias `QtWebEngineCachePath=`), says where it is, for `<QtWebEngineCacheRoot>`. Either root opts the entry in on its own, and a template whose placeholder has no declared root is dropped.

Second, the default selection is wider: Caches, StorageQuota, Telemetry and VisitedLinks. The QtWebEngine catalog gives two low-risk targets, `QuotaManager` and `Visited Links`, scaffolds of their own, while the WebView catalog keeps them outside its defaults, in the legacy `LegacyStorageQuota` and the host-risk `WebHistory`. The [EntryBuilder readme](../entrybuilder/readme.md#qtwebengineroot--qtwebenginecacheroot--qtwebenginescaffolds) and the catalog header explain why.

The QtWebEngine catalog is smaller. Its scaffolds are `Caches`, `Favicons`, `PrivacySandbox`, `Security`, `StorageQuota`, `Telemetry`, `VisitedLinks`, `WebCookies`, `WebHistory`, `WebSession` and `WebStorage`, plus the legacy `LegacyTelemetry` and `LegacyWebStorage`. Anything outside the default four has to be selected.

---

# Electron Scaffolds

A third family covers apps that embed Electron, from `electron.ini` in the scaffolds directory.

`ElectronRoot=` (alias `ElectronPath=`) opts in, and `ElectronScaffolds=` / `ExcludeElectronScaffolds=` select, with the same `All` value and warnings as the other families.

Like QtWebEngine, Electron has a second root: `ElectronUpdaterRoot=` declares the electron-updater download cache. Either root opts the entry in on its own, and a template whose placeholder has no declared root is dropped.

```ini
[Hazel Music *]
Package=Hazel.Music_abc123
LangSecRef=3021
ElectronRoot=%LocalAppData%\Hazel Music
ElectronUpdaterRoot=%LocalAppData%\Hazel-music-updater
```

The default selection is `AppLogs`, `Caches`, `StorageQuota`, `Telemetry` and `UpdaterCache`. The full set adds `MediaDRM`, `PrivacySandbox`, `Security` and `TempFiles`, plus the host-risk `WebCookies` and `WebStorage`, and all of those have to be selected. The Electron catalog has no legacy tier.

---

# Output Formatting

Before saving, UWPBuilder runs the output through WinappDebug with Optimizations on. As a result:

- Entries are sorted alphabetically
- Keys are grouped by type (category, then detection, then cleaning) and renumbered, and FileKeys are sorted by value
- FileKeys with the same path are merged into one key with a combined pattern list

So the output won't show the order you wrote or generated the keys in.

The output file starts with a generated header:

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

The lint pass is supposed to change formatting only, so UWPBuilder checks. It breaks every entry into its meaningful parts before and after the pass and compares the two. The changes WinappDebug is allowed to make (reordering, renumbering, sorting patterns, merging keys with the same path, removing exact duplicates) don't register in that comparison, so any difference means the generator wrote something the linter quietly rewrote or threw away.

When that happens, the run reports every difference, saves the pre-lint output to `uwp.prelint.ini` beside the output, and sets exit code 1, so a scripted build fails instead of committing bad data.

This isn't an ordinary validation warning. It means an AppInfo value or a scaffold template produced a key the linter considers wrong, and that's a bug in the sources. All three generators (BrowserBuilder, UWPBuilder and EntryBuilder) run this check.

---

# Tips & Best Practices

- **Use `FileKeyBase=` for package paths and `FileKey=` for win32 paths.** They do the same thing, but the convention keeps hybrid entries readable
- **Never write `Skip=False`.** `Skip` and `SkipUWPFileKeys` act whenever they're present, so delete the key to turn the effect off
- **Don't number your keys.** UWPBuilder ignores source numbering, and the lint pass renumbers, reorders, and may merge FileKeys anyway
- **Comment AppInfo sections freely.** Comments never reach the output

---

# Troubleshooting

###### Note: You won't see any of these messages during a scripted build. Silent mode (`-s`) hides all menu output, and the global log is only saved to `winapp2ool.log` when the run exits nonzero or you pass the global `-writelog`

| Message | Cause |
|:-|:-|
| "UWP.ini not found or empty in: {dir}" | The source directory has no readable `UWP.ini`. Nothing is generated, and the run fails with exit code 1 |
| "No app definitions found in: {dir}" | The `AppInfo\` subdirectory is missing or empty, or has no readable sections. Nothing is generated, and the run fails with exit code 1 |
| "[{name}].{key} has no value; key dropped" | A key that needs a value was left empty. Only `Skip=`, `SkipUWPFileKeys=` and the scaffold lists mean something with no value |
| "Duplicate AppInfo section [{name}] at {file} line {n} ignored; keeping the first definition, in {file}" | Two AppInfo sections share a name (in any case), in one file or in two. Files are read in name order, and only the first definition is used. The output is still written, but the run fails with exit code 1 |
| "Unexpected section in template file: [{name}]" | `UWP.ini` has a section that isn't an `[EntryScaffold: ...]` |
| "Unexpected key in scaffold [{name}]: {key}" | An `[EntryScaffold: ...]` section has a key other than `DetectFileBase=` or `FileKeyBase=` |
| "`%Package%` in [{name}].{key} is now spelled `<Package>`; key dropped" | The key uses a [retired spelling](#the-retired-package-and-webviewroot-spellings), and the run fails with exit code 1. A shared catalog template reports `; template dropped` instead |
| "Variable '{name}=' declared in [{entry}] but never referenced by any <{name}> token; possible typo" | No key in the entry uses this variable. Usually it's a misspelled key name, which declared a variable instead, so nothing from that key reaches the entry |
| "Package= value in [{name}] contains '<' or '>'; Package is never variable-expanded and the value is used literally" | A `Package=` value holds a bracket. It's used exactly as written |
| "SpecialDetect is deprecated; key dropped from [{name}]. Replace with Detect or DetectFile" | The section uses `SpecialDetect=`, which is dropped |
| "Default= declared in [{name}]; UWPBuilder never emits Default, ignoring" | The section has a `Default=` key, which is dropped |
| "No Package key in [{name}], skipping" | The AppInfo section has no `Package=` key, so no entry is generated |
| "No LangSecRef or Section in [{name}], skipping" | The AppInfo section has no category key, so no entry is generated |
| "Both LangSecRef and Section present in [{name}], using LangSecRef" | The AppInfo section has both category keys. The entry is generated with `LangSecRef` and the `Section` value is thrown away |
| "Skipping scaffold FileKeys for: {name}" | Just information: the entry has `SkipUWPFileKeys` |
| "Scaffold directory not found: {dir}; no scaffold keys will be generated" | The scaffolds directory doesn't exist. Generation carries on without any scaffold keys |
| "Unknown WebView scaffold '{name}' requested by [{entry}], skipping" | A `WebViewScaffolds=` name isn't in the catalog. This is also what a missing `webview.ini` looks like: every scaffold the entry wants, defaults included, is unknown. Unknown `ExcludeWebViewScaffolds=` names produce **no** message |
| "WebViewScaffolds=All in [{entry}] with redundant additional names ({names}); ignoring" | A non-legacy scaffold was named beside `All`, which already selects it |
| "Both WebViewScaffolds and ExcludeWebViewScaffolds set in [{entry}]; applying exclusions to explicit list" | Exclusions usually go with the default set or `All` |
| "Scaffold '{name}' selected by [{entry}] produced no keys; its templates reference a root the entry did not declare" | A scaffold the entry selected by name only has templates for a root the entry didn't declare, such as `UpdaterCache` on an entry with only `ElectronRoot=` |
| "Duplicate WebViewScaffold: name '{name}'; last definition wins" | The catalog defines the same scaffold twice |
| "Unexpected section in scaffold catalog {file}: [{name}]" | A section in the scaffolds directory has no `Scaffold:` in its header, often a misspelling. It's ignored |
| "Scaffold family '{family}' in {dir} is not consumed by any builder; check the spelling of its section headers" | A catalog section names a family no builder reads, such as `[ElctronScaffold: ...]` |
| "Unknown tier '{value}' in WebViewScaffold: [{section}]; the only tier is Legacy" | A catalog section has a `Tier=` value other than `Legacy`. The tier is ignored and the scaffold stays in `All` |

QtWebEngine and Electron messages read the same, with `QtWebEngine` or `Electron` in place of `WebView`. Each `All` expansion is also logged: `WebViewScaffolds=All in [{entry}]; expanded to N scaffold(s) from catalog, M legacy withheld`.
