# EntryBuilder

**EntryBuilder** is a winapp2ool build-chain module for winapp2.ini maintainers that generates winapp2.ini entries from a shorthand DSL. You write one ini section per application. Anything already in winapp2 syntax passes through unchanged, and a few keys that only winapp2ool understands expand into ordinary winapp2 keys at generation time. It's a preprocessor for winapp2.ini: you write the repetitive parts once and EntryBuilder writes them out. If you only use winapp2.ini with a cleaner, you don't need EntryBuilder.

The winapp2.ini build uses it for every entry that isn't a browser, Microsoft Store apps included. Their sources are the per-letter files in `Assembler\Sources\EntryBuilder\`, which EntryBuilder reads alongside the shared scaffold catalogs in `Assembler\Sources\Scaffolds\`. By default it writes a single `entrybuilder.ini`. With `-split` it writes the 27 per-letter files that the build commits to `Assembler\Generated\EntryBuilder\` and combines into the final winapp2.ini.

### What does EntryBuilder do?

EntryBuilder reads every `*.ini` file in its source directory, combines them in memory, and turns each section into an entry. It expands the shorthand, runs the result through WinappDebug, and saves it to `entrybuilder.ini`. There are two kinds of shorthand:

- Scaffold families. An entry that declares `WebViewRoot=` (an embedded WebView2 data folder), `QtWebEngineRoot=` (an embedded QtWebEngine profile), `ElectronRoot=` (an Electron app's `userData` folder) or `Package=` (a Store app's package folder) gets a baseline set of cleaning FileKeys from a shared catalog, one set per root or package.
- List variables. An entry can declare a comma-separated list (`Versions=10.0,11.0,12.0`) and use it in any key as `<Versions>`. Each key that uses it becomes one key per value, or one per combination when it uses several lists.

### Why EntryBuilder?

EntryBuilder makes it possible to maintain winapp2.ini correctly at scale by providing a small shorthand DSL that lowers to formal winapp2.ini syntax.

Apps that embed a browser engine leave the same Chromium residue behind whatever the app is, and every Store app keeps the same few folders inside its package. EntryBuilder keeps those cleaning patterns in shared catalogs, one each for WebView2, QtWebEngine, Electron and packaged apps, so they aren't copied into every entry.

None of it is mandatory. A plain winapp2 entry pasted into a source file comes out unchanged, so you can add shorthand to an entry a piece at a time, or never. The `<Name>` variables work the same way in BrowserBuilder.

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
   - [Archive](#archive)
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
11. [Store and hybrid apps](#store-and-hybrid-apps)
    - [Package](#package)
    - [The Package variable](#the-package-variable)
    - [PackageScaffolds / ExcludePackageScaffolds](#packagescaffolds--excludepackagescaffolds)
    - [Hybrid apps](#hybrid-apps)
    - [UWPBuilder's leftover keys](#uwpbuilders-leftover-keys)
    - [Real entries](#real-entries)
12. [Validation and Warnings](#validation-and-warnings)
13. [Run Output](#run-output)
14. [Tips & Best Practices](#tips--best-practices)
15. [Troubleshooting](#troubleshooting)

**Examples**

16. **[Examples](EXAMPLES.md)**: thirteen worked examples with captured input and output

---

# Requirements

- A source directory containing one or more `*.ini` files with at least one parseable section
- The shared scaffold directory (normally `Assembler\Sources\Scaffolds`), if any entry declares a scaffold root or a package. EntryBuilder reads every `*.ini` in it and tells the catalogs apart by their section headers, not their file names: `[WebViewScaffold: ...]` sections (in `webview.ini`) serve `WebViewRoot=`, `[QtWebEngineScaffold: ...]` sections (`qtwebengine.ini`) serve `QtWebEngineRoot=` and `QtWebEngineCacheRoot=`, `[ElectronScaffold: ...]` sections (`electron.ini`) serve `ElectronRoot=` and `ElectronUpdaterRoot=`, and `[PackageScaffold: ...]` sections (`package.ini`) serve `Package=`

If the source directory is empty, missing, or holds no parseable sections, EntryBuilder reports `No EntryBuilder source definitions found in: <directory>`, writes nothing, and fails the run (exit code 1). If the scaffold directory is missing, it warns `Scaffold directory not found: <dir>; no scaffold keys will be generated` and builds the entries without scaffold keys. A directory with no catalog for a family gives no such warning, but every entry that uses the family warns `Unknown … scaffold` for each scaffold it selects, defaults included. Neither of these fails the run, and entries that declare no root or package aren't affected. A Store app built without `package.ini` still gets its package detection, just not the baseline FileKeys.

---

# Quick Start

### Common Workflow

1. Place per-letter source files (`A.ini`, `B.ini`, ...) in the source directory
2. Open EntryBuilder from the **Entry Lab** main menu, or run `winapp2ool -entrybuilder` from the command line
3. Choose **Run**. EntryBuilder combines the files, expands the shorthand, and writes the result to `entrybuilder.ini`

The source directory, save target and scaffold directory all default to the current directory. The build script runs EntryBuilder from the `Assembler` folder with `-1d \Sources\EntryBuilder -3d \Sources\Scaffolds`, where a leading `\` means the current directory, and adds `-split` and an output folder.

---

# Common Uses

| To | Run | Example |
|:-|:-|:-|
| Generate entries from the source files in the current directory | `winapp2ool -entrybuilder` | |
| Generate from the project's source files, saving `entrybuilder.ini` in the current directory | `winapp2ool -entrybuilder -1d ..\..\Assembler\Sources\EntryBuilder` | [A pass-through entry with a single-value variable](EXAMPLES.md#example-1-a-pass-through-entry-with-a-single-value-variable) |
| Generate with the shared scaffold catalogs in silent mode | `winapp2ool -entrybuilder -1d ..\..\Assembler\Sources\EntryBuilder -3d ..\..\Assembler\Sources\Scaffolds -s` | [Default WebView scaffolds](EXAMPLES.md#example-8-default-webview-scaffolds) |
| Save the output under another file name | `winapp2ool -entrybuilder -2f test-output.ini` | |
| Write the per-letter artifact files into `Assembler\Generated\EntryBuilder` | `winapp2ool -entrybuilder -split -2d ..\..\Assembler\Generated\EntryBuilder` | |

In the menu, you make the same choices with the source, save target and scaffolds directory choosers and the output splitting toggle before choosing **Run**; see [Menu Options](#menu-options). `-s` is one of winapp2ool's [global args](../../Readme.md#global-args).

---

# Menu Options

| Option | Effect | Notes |
|:-|:-|:-|
| Run (default) | Generate entries from the configured source directory | Reports `No EntryBuilder source definitions found`, writes nothing, and fails the run if there's nothing to read |
| Choose source directory | Select the directory containing per-letter source files | Only the directory is used; the file name is ignored |
| Choose save target | Select where to save the generated entries | Default: `entrybuilder.ini` in the current directory. With output splitting on, only its folder is used |
| Choose scaffolds directory | Select the folder holding the shared scaffold catalogs | Default: current directory |
| Choose archive folder | Select the folder where `archived.ini` is written | None by default. See [Archive](#archive) |
| Toggle output splitting | Write per-letter files into the save target's folder instead of one file | Off by default. The menu equivalent of `-split` |
| Reset Settings | Restore all settings to their defaults | Only shown when settings have been changed |

---

# Command-Line Arguments

Run EntryBuilder with `winapp2ool -entrybuilder`. The build adds the global `-s` flag to run it without the menu.

### File Selection

Each file slot has an index for the `-Nd` (directory) and `-Nf` (file name) arguments:

| Index | File | Default |
|:-|:-|:-|
| 1 | Source directory | Current directory (directory only; file name is ignored) |
| 2 | Save target | `entrybuilder.ini` in current directory |
| 3 | Shared scaffold directory (only `.Dir` is read, so `-3f` has no effect) | Current directory |
| 4 | Archive folder, also `-archived` (only `.Dir` is read). See [Archive](#archive) | None |

| Arg | Effect |
|:-|:-|
| `-Nd path` | Set directory for file slot N |
| `-Nf name` | Set file name for file slot N |
| `-Nf subdir\name` | Set file name within a subdirectory of its path |

### Toggles

| Arg | Effect |
|:-|:-|
| `-split` | Write per-letter files (`#.ini`, `A.ini` ... `Z.ini`) into the save target's folder instead of one file. Each entry goes in the file for the first character of its name, and anything that doesn't start with a letter goes in `#.ini`. All 27 files are written every time, empty or not, each with a do-not-edit header |

---

# How Generation Works

Every run goes through the same steps:

1. **Combine**: EntryBuilder loads every `*.ini` in the source directory, in alphabetical order by file name, and joins them in memory. When two sections share a name, the first one wins and the run fails (see [Troubleshooting](#troubleshooting)).
2. **Parse**: each section becomes one entry definition. Keys EntryBuilder knows, winapp2 or shorthand, are claimed by name. *Every other key* declares a list variable.
3. **Resolve variables**: a variable whose values use other variables (`<Other>` inside a declaration) gets them filled in, so every value list is plain text. A cycle is reported and left unexpanded.
4. **Infer detection**: an entry that declares a variable named `Root` gets a `Detect=<Root>` or `DetectFile=<Root>` (see [The Reserved Root Variable](#the-reserved-root-variable)), and an entry that declares `Package=` gets a `DetectFile` for each package, ahead of any other (see [Store and hybrid apps](#store-and-hybrid-apps)).
5. **Validate**: an entry that can't be built is skipped with a warning, and a questionable one is built with a warning (see [Validation and Warnings](#validation-and-warnings)).
6. **Generate**: EntryBuilder writes each entry's keys in winapp2 order. Each content key goes through scaffold expansion (catalog templates copied in for each selected scaffold and each declared root or package), then root substitution (`<WebViewRoot>`, `<QtWebEngineRoot>`, `<QtWebEngineCacheRoot>`, `<ElectronRoot>` and `<ElectronUpdaterRoot>`, in any case, replaced with each declared root), then token expansion (`<Name>` fan-out). Numbered key families are renumbered from 1.
7. **Normalize**: WinappDebug runs over the whole file with every repair and optimization on.
8. **Save**: EntryBuilder writes the output with a comment header (see [Run Output](#run-output)).

Root substitution comes before token expansion, so a `<Name>` token inside a root value (`QtWebEngineRoot=<QtRoot>\Default`) lands in the key and expands like any other token there. Root placeholders look like variables but aren't: they're replaced as plain text first, so a `<Root>` in a root value and a `<Root>` in the template end up as one token and co-vary.

Not every key family goes through every step:

| Key family | Root placeholders | Expansion domain |
|:-|:-|:-|
| `Detect` | No  | Registry |
| `DetectFile` | No  | Filesystem |
| `FileKey` / `FileKeyBase` | Yes | Filesystem |
| `RegKey` / `RegKeyBase` |  Yes | Registry |
| `ExcludeKey` / `ExcludeKeyBase` |  Yes | By flag: `REG` → Registry, otherwise Filesystem |
| Scaffold catalog templates | Yes (their purpose) | Filesystem |

The expansion domain decides what happens to a `<Name>` token that matches no variable (see [Undeclared Tokens](#undeclared-tokens)). Root placeholders are **not** substituted in `Detect` or `DetectFile`, so they reach token expansion as undeclared tokens: a `DetectFile` holding one is dropped, and a `Detect` keeps it as literal text. Write the path, or a `<Variable>`, in detection keys instead.

---

# Source File Format

Source files are ordinary ini files. Each section is one application and becomes one winapp2.ini entry, and the section name is copied to the output unchanged as the entry name.

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

An application that needs more than one entry (say, one for caches and another for saved logins) gets one section per entry, each repeating the keys they share.

The project's sources are split into `#.ini` through `Z.ini`, but EntryBuilder only needs one file.

---

# Pass-Through Keys

These standard winapp2 keys pass through to the output, after the processing in the [table above](#how-generation-works):

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


EntryBuilder ignores key numbers in the sources and renumbers every family from 1 after expansion. The project's source files leave the numbers off.

The `Base=` and plain form of each content key give the same output. The only difference is in the run statistics, which count `FileKeyBase=`, `RegKeyBase=` and `ExcludeKeyBase=` as *generated* content and the plain forms as *pass-through* content (see [Run Output](#run-output)). By convention, use the `Base=` form for a template you expect to expand and the plain form for a literal key.

---

# Shorthand Keys

Only winapp2ool understands these keys. EntryBuilder uses them up during generation, and they never reach the output. `Package=` and its selectors have a section of their own, [Store and hybrid apps](#store-and-hybrid-apps).

## WebViewRoot

The full path to the app's WebView2 data folder, usually `%AppData%\<AppName>` or `%LocalAppData%\<AppName>\<Subfolder>\EBWebView`.

- The root is the folder that holds the profiles, not a profile. Profile-scoped templates use `<WebViewRoot>\*\`, so `Default` and named profiles such as `WV2Profile_*` are all covered.
- An entry can declare more than one `WebViewRoot=`, and gets each scaffold FileKey once per root.
- `<WebViewRoot>` also works in any `FileKey`, `RegKey` or `ExcludeKey` value, `Base=` or plain. EntryBuilder puts each declared root in its place, one key per root.
- An entry with no roots and no FileKey or RegKey content is skipped with a warning.

### The retired `%WebViewRoot%` spelling

Older versions spelled the root placeholders `%WebViewRoot%`, `%QtWebEngineRoot%`, `%QtWebEngineCacheRoot%`, `%ElectronRoot%` and `%ElectronUpdaterRoot%`. EntryBuilder drops a key still written that way, including one missing its closing `%` (`%ElectronRoot\logs|*`), warns with the new spelling, and fails the run with exit code 1:

```
`%QtWebEngineCacheRoot%` in [Test App *].FileKeyBase is now spelled `<QtWebEngineCacheRoot>`; key dropped
`%ElectronRoot` in [Test App *].FileKey looks like the retired `%ElectronRoot%` token, which is now spelled `<ElectronRoot>`; key dropped
```

`%...%` is winapp2.ini's environment variable syntax, so a stale token would otherwise reach the output looking like a real path. EntryBuilder checks every key except `Warning=` and `Package=`, and refuses the retired `%Package%` / `%Package1%` spelling of [`<Package>`](#the-package-variable) the same way. The shared catalogs in `Assembler\Sources\Scaffolds` are checked as they load, and a catalog template with a retired token is dropped and fails the run the same way.

## WebViewScaffolds / ExcludeWebViewScaffolds

`WebViewScaffolds=` is a comma-separated list of the catalog scaffolds to apply to the entry.

| Form | Behavior |
|:-|:-|
| Key absent | The default set: `Caches`, `Telemetry` |
| Key present, value empty | No scaffolds from this family |
| `WebViewScaffolds=Caches,Telemetry,DRMData` | Exactly the named scaffolds |
| `WebViewScaffolds=All` | Every scaffold in the catalog except the legacy tier, **including the host-risk ones** (cookies, history, sessions, web storage, login data) |
| `WebViewScaffolds=All,LegacyTelemetry` | `All`, plus the named legacy scaffold |

`ExcludeWebViewScaffolds=` is a comma-separated list of scaffolds to take out of the selection. Pair it with `All` for everything except a few:

###### Example

```ini
WebViewScaffolds=All
ExcludeWebViewScaffolds=WebStorage,LoginData
```

Exclusions also work without a `WebViewScaffolds=` key, where they come out of the default set. On an explicit list other than `All` they still apply, but warn (`Both WebViewScaffolds and ExcludeWebViewScaffolds set in [Entry]; applying exclusions to explicit list`); list only the scaffolds you want instead. An unknown scaffold name is dropped with a warning, beside `All` or not.

The catalog, not this page, decides the scaffold names, and [`Assembler/Sources/Scaffolds/readme.md`](../../../Assembler/Sources/Scaffolds/readme.md) keeps the current lists. When this was written, the WebView catalog defined `Autofill`, `Autoplay`, `BookmarkFavicons`, `Caches`, `DefaultApps`, `DownloadHistory`, `DRMData`, `PrivacySandbox`, `LoginData`, `Security`, `Shopping`, `Sync Data`, `Telemetry`, `WebCookies`, `WebHistory`, `WebSession`, `WebStorage`, and in the legacy tier `LegacyCaches`, `LegacyDownloadHistory`, `LegacyDRMData`, `LegacyExtensionCookies`, `LegacyProgressiveWebApps`, `LegacyStorageQuota`, `LegacyTelemetry`, `LegacyWebHistory`. Only `Caches` and `Telemetry` are on by default. The scaffolds that remove state a user would notice (`WebCookies`, `WebHistory`, `WebSession`, `WebStorage`, `LoginData`) always need an opt-in, through `All` or by name.

### The legacy tier

`All` leaves out any catalog section marked `Tier=Legacy`. An entry gets a legacy scaffold only by naming it, alone or beside `All`. These scaffolds hold patterns that no current host writes but an older or unusual one might, and their names start with `Legacy` so the tier shows in a selection line. A non-legacy name beside `All` is redundant and warns, so a legacy name is the only thing worth listing there.

## QtWebEngineRoot / QtWebEngineCacheRoot / QtWebEngineScaffolds

QtWebEngine is a second scaffold family, for apps that embed Qt WebEngine instead of WebView2. It works **exactly like the WebView family above**, under these names:

| WebView family | QtWebEngine family |
|:-|:-|
| `WebViewRoot=` | `QtWebEngineRoot=` |
| `WebViewScaffolds=` | `QtWebEngineScaffolds=` |
| `ExcludeWebViewScaffolds=` | `ExcludeQtWebEngineScaffolds=` |
| `<WebViewRoot>` placeholder | `<QtWebEngineRoot>` placeholder |
| `webview.ini` catalog | `qtwebengine.ini` catalog |

`All`, the legacy tier and exclusions work the same way, and one entry can use any mix of the three families. The QtWebEngine catalog is smaller: `Caches`, `Favicons`, `PrivacySandbox`, `Security`, `StorageQuota`, `Telemetry`, `VisitedLinks`, `WebCookies`, `WebHistory`, `WebSession`, `WebStorage`, and in the legacy tier `LegacyTelemetry`, `LegacyWebStorage`.

The default set is `Caches`, `StorageQuota`, `Telemetry` and `VisitedLinks`.

`QtWebEngineRoot=` names one profile directory, profile folder included: `QtWebEngineRoot=%LocalAppData%\AppName\QtWebEngine\Default`, not the `QtWebEngine` folder above it. For an app with several profiles, declare each one, or put the profile names in an inline list, as calibre's entry does ([Example 10](EXAMPLES.md#example-10-qtwebengine-across-several-profiles)):

```ini
QtRoot=%LocalAppData%\calibre-ebook.com\calibre\QtWebEngine
QtWebEngineRoot=<QtRoot>\<Default,OffTheRecord,viewer-lookup>
```

### QtWebEngineCacheRoot

`QtWebEngineCacheRoot=` names the profile's HTTP cache folder, for templates that use `<QtWebEngineCacheRoot>`. Qt keeps the cache outside the profile directory, so `QtWebEngineRoot=` can't reach it:

```ini
QtWebEngineRoot=%LocalAppData%\VideoKeeper\QtWebEngine\Default
QtWebEngineCacheRoot=%LocalAppData%\VideoKeeper\cache\QtWebEngine\Default
```

It works like `ElectronUpdaterRoot=` below: either root on its own opts the entry into the family, and a template whose placeholder has no declared root is dropped.

## ElectronRoot / ElectronUpdaterRoot / ElectronScaffolds

Electron is the third family, for apps built on Electron, which bundle their own Chromium instead of embedding a browser view. Its keys follow the same pattern as the other two:

| WebView family | Electron family |
|:-|:-|
| `WebViewRoot=` | `ElectronRoot=` |
| `WebViewScaffolds=` | `ElectronScaffolds=` |
| `ExcludeWebViewScaffolds=` | `ExcludeElectronScaffolds=` |
| `<WebViewRoot>` placeholder | `<ElectronRoot>` placeholder |
| `webview.ini` catalog | `electron.ini` catalog |

The default set is `AppLogs`, `Caches`, `StorageQuota`, `Telemetry` and `UpdaterCache`. The rest of the catalog is opt-in: `MediaDRM`, `PrivacySandbox`, `Security` and `TempFiles` are safe, and `WebCookies` and `WebStorage` are host-risk.

`ElectronRoot=` is the app's `userData` folder, usually `%AppData%\<ProductName>`. In Electron that folder is the Chromium profile itself, with no `Default` folder inside it.

```ini
[Signal Messenger *]
LangSecRef=3022
DetectFile=%AppData%\Signal
ElectronRoot=%AppData%\Signal
```

### ElectronUpdaterRoot

Electron has a second root too. `ElectronUpdaterRoot=` names electron-updater's download cache, a folder under `%LocalAppData%` holding `installer.exe`, `package.7z` and a `pending\` folder. You have to declare it yourself, because its name comes from the app's build ID rather than its `userData` path (`%AppData%\Humble App` against `%LocalAppData%\humble-client-updater`). Without it, the default `UpdaterCache` scaffold produces nothing.

### Multiple roots

Every root key can repeat. Number the copies; a comma in a root value is taken as part of the path, and EntryBuilder warns about one outside a `<>` list:

```ini
ElectronRoot1=%AppData%\Notion
ElectronRoot2=%AppData%\Notion\partitions\*
```

## FileKeyBase / RegKeyBase / ExcludeKeyBase

Templates of your own, added to whatever the scaffolds produce. Each works like its plain form (`FileKey=` and so on, see [Pass-Through Keys](#pass-through-keys)): root substitution, then token expansion, then renumbering from 1. Use them for anything the catalogs don't cover, such as the app's own logs, crash dumps or version-numbered folders.

## Skip

Leaves the entry out of the output. Use it to set an entry aside without deleting its definition. To retire an entry, use [Archive](#archive) instead.

```ini
Skip=
```

EntryBuilder ignores the value: a bare `Skip=`, `Skip=True` and even `Skip=False` all skip the entry.

## Archive

Sends the entry to the archive instead of the live output. The entry is parsed, expanded and checked like any other, then written to `archived.ini` in the archive folder (`-4d`, or **Choose archive folder** in the menu). The value is a reason for whomever reads the source, EntryBuilder never writes it anywhere. A bare `Archive=` archives too.

```ini
[LEGO Universe *]
Archive=Discontinued (2012)

LangSecRef=3021
DetectFile=%AppData%\LEGO Universe
FileKey=%AppData%\LEGO Universe|*.log
```

To bring the entry back, delete the `Archive=` line.

- `archived.ini` is one file, never split, with the same kind of do-not-edit header as the other output. Given an archive folder, EntryBuilder always writes it, with only the header when nothing is archived.
- With no archive folder, archived entries go to neither output, and the run warns with their count and fails (exit code 1).
- An archived entry and a live one can't share a name: the two meet in the source merge, which keeps the first and fails the run.
- An entry that can't be built is skipped with its usual warning, `Archive=` or not.
- `Archive=False` is refused, and so is `Archive=` beside `Skip=`. Either way the entry isn't built and the run fails.

---

# List Variables and Token Expansion

Anything an entry repeats, such as version numbers, registry roots, path fragments or value names, can be declared once as a list variable and used in other keys as `<Name>`. EntryBuilder turns each key that uses variables into one key per combination of their values. Keys without `<>` tokens pass through unchanged.

## Declaration

Any key that isn't a winapp2 key or one of EntryBuilder's shorthand keys declares a list variable. The value is a comma-separated list, with spaces trimmed from each item, and variable names are case-insensitive.

###### Example

```ini
Versions=9.0,10.0,11.0,12.0,14.0,15.0,16.0
MRUNum=1,2,3
RegRoot=HKCU\Software\Microsoft\Office
```

A variable with one value, like `RegRoot` above, is a handy abbreviation for a long path even when nothing fans out.

## Reference

A `<Name>` token can go anywhere in the value of a key that takes expansion (see the [key-family table](#how-generation-works)): the path, the pattern, the flag, either side of the `|`. Tokens match the entry's variables in any case.

```ini
Versions=9.0,10.0,11.0,12.0,14.0,15.0,16.0
MRUNum=1,2,3
RegRoot=HKCU\Software\Microsoft\Office

Detect=<RegRoot>\<Versions>\Outlook
RegKeyBase=<RegRoot>\<Versions>\Outlook\Office Finder|<MRUNum>
```

## Inline Lists

A token with a comma in it is an **inline list**, a one-off variable you don't have to declare:

```ini
RegKeyBase=HKCU\Software\7-Zip\<Compression,Extraction>|ArcHistory
```

gives two RegKeys, the same as a declared variable with two values would. Use an inline list for a short list in one key, and declare a variable once the list gets reused or long.

Two different inline lists in one key cross-multiply like any two variables. The same inline list written twice in one key co-varies instead: both copies take the same value in each output key, and the log notes it. Declare a named variable when that's what you mean.

Inline lists work inside declarations too. Only commas outside `<>` separate a declaration's values, and the inline list is flattened into the variable's values before any key uses it. So

```ini
Root=HKCU\Software\Ability <6.0,7.0,11.0>\Ability Write
```

declares a `Root` with three values, and [detection inference](#the-reserved-root-variable) turns it into `Detect1` through `Detect3`.

## Indexed References

`<Name[N]>` picks the Nth value of a declared variable, counting from 1, and doesn't fan out:

```ini
Versions=9.0,10.0,11.0
DetectFile=%ProgramFiles%\App <Versions[3]>
```

gives `DetectFile=%ProgramFiles%\App 11.0`. The index counts the variable's values after any nested references and inline lists in its declaration are flattened. An index of 0 or past the end of the list drops the key with a warning, in registry keys too: the name is declared, so the token can't be meant literally. An indexed name that isn't declared is an [undeclared token](#undeclared-tokens) like any other.

`<Versions>` and `<Versions[2]>` in one key count as two separate variables, so they cross-multiply rather than co-vary. A comma still makes an inline list, so `<a[1],b>` is the two literal values `a[1]` and `b`.

## Cartesian Fan-out and Co-variance

A key that uses variables with A, B and C values gives A × B × C keys, one per combination. The leftmost variable changes slowest, so the `RegKeyBase` example under [Reference](#reference) gives 7 × 3 = 21 RegKeys, in the order `(9.0, 1)`, `(9.0, 2)`, `(9.0, 3)`, `(10.0, 1)`, ...

A declared variable used twice in one key **co-varies**: both uses take the same value in each output key instead of multiplying. That's what you want for a version's folder and the same version's log file name (see [Example 4](EXAMPLES.md#example-4-co-variance-of-repeated-variables)).

## Nested References

A variable's values can use other variables. EntryBuilder fills those in before expanding any key, so every key sees plain values.

```ini
Versions=10.0,10.1
ArcMapKey=ArcMap,Desktop<Versions>\ArcMap
Detect=HKCU\Software\ESRI\<ArcMapKey>
```

resolves `ArcMapKey` to `(ArcMap, Desktop10.0\ArcMap, Desktop10.1\ArcMap)`, then `Detect` fans out across those three values. A cycle (A uses B, B uses A) is reported (`Variable <A> participates in a cyclic dependency with <B>; values left unresolved`) and the variables in it are left as written.

## The Reserved Root Variable

`Root` is a reserved name with one extra behavior: **an entry that declares `Root=` always gets a detection key from it.** If the value starts with a registry hive, EntryBuilder adds `<Root>` as the entry's first `Detect`; otherwise it adds it as the first `DetectFile`:

```ini
[A Hat in Time *]
Root=%ProgramFiles%\Steam\steamapps\common\HatinTime\HatinTimeGame

Section=Games
FileKeyBase=<Root>\Logs|*
```

gets `DetectFile=%ProgramFiles%\Steam\steamapps\common\HatinTime\HatinTimeGame` without writing a detection key. The added token expands like any other, so a `Root` with several values gives `DetectFile1`, `DetectFile2`, and so on.

The details:

- If the entry already has `Detect=<Root>` or `DetectFile=<Root>`, nothing is added. Other detection keys are left alone either way.
- Inference counts as a use of `Root`, so an entry that never writes `<Root>` doesn't get the unreferenced-variable warning.
- A `Root` that mixes registry and filesystem values warns (`[Entry] Root mixes registry and filesystem values; inferring detection from the first value's domain`). Split them into separate variables.
- A `Root` whose value still holds unresolved `<tokens>`, usually from a cycle, skips inference with a warning.
- For a base path without automatic detection, use any other name, such as `DiskRoot` or `AppRoot` (see [Example 1](EXAMPLES.md#example-1-a-pass-through-entry-with-a-single-value-variable)).

## Undeclared Tokens

What happens to a `<Name>` token that's neither a declared variable nor an inline list depends on the key's expansion domain (see the [key-family table](#how-generation-works)):

| Domain | Behavior | Message | Why |
|:-|:-|:-|:-|
| Filesystem | Drop the key | `Undeclared variable <X> in filesystem-domain key [Entry].FileKey; dropping key` | Windows paths can't contain `<` or `>`, so the token has to be a typo |
| Registry | Emit the literal text | `Undeclared variable <X> in registry-domain key [Entry].RegKey; passed through as literal` | Registry names can contain `<` and `>`, so the token may be meant literally |

The filesystem drop shows as a warning in the menu output. The registry pass-through is an advisory that only goes to the log (see [Validation and Warnings](#validation-and-warnings)).

## Typo Backstops

EntryBuilder warns when no key uses a declared variable: `Variable 'Versions=' declared in [Entry] but never referenced by any <Versions> token; possible typo`. That catches a misspelled reference, such as `Versions=` declared but `<Versoin>` used, which the registry domain would otherwise pass through without complaint. It catches a misspelled key name too, since `DetectFiel=` declares a variable that nothing uses.

A variable named after a reserved key also warns (`Variable declaration 'X=' in [Entry] shadows reserved key name; possible typo`).

---

# Store and hybrid apps

A Microsoft Store app (UWP or MSIX) keeps its data under `%LocalAppData%\Packages\PACKAGE_FOLDER`, and part of that folder looks the same from app to app. A Store app's entry is an ordinary EntryBuilder entry that names its package folder with `Package=`. That one key does three things:

- It adds a `DetectFile=%LocalAppData%\Packages\PACKAGE_FOLDER` for each package, ahead of any other detection
- It defines the `<Package>` variable for the entry's own keys
- It opts the entry into the Package scaffold family, whose default `Baseline` scaffold holds the folders every Store app gets cleaned

So a package and a category make a complete entry:

```ini
[BreeZip *]
Package=3138AweZip.AweZip_ffd303wmbhcjt
LangSecRef=3021
```

[Real entries](#real-entries) shows what EntryBuilder makes of it.

## Package

The value is the package's folder name under `%LocalAppData%\Packages`, used exactly as written. An app with more than one package repeats the key as `Package1=`, `Package2=` and so on. EntryBuilder ignores the digits and keeps the packages in the order they're declared.

A package name is never variable-expanded, because a fan-out would change how many packages there are and shift what every `<Package[N]>` in the entry points at. A value holding `<` or `>` warns (`Package= value in [Entry] contains '<' or '>'; Package is never variable-expanded and the value is used literally`) and is used as written.

`Package=` counts as content, so an entry with only packages and a category is built instead of skipped for having nothing to emit. The package detection comes from `Package=` itself, not from the scaffold catalog, so `PackageScaffolds=` never removes it, and an entry with a package never gets the always-on warning.

## The Package variable

`<Package>` has one value per package, `%LocalAppData%\Packages\PACKAGE_FOLDER`. It works in any key that takes [token expansion](#list-variables-and-token-expansion), including root keys (`WebViewRoot=<Package>\LocalState\EBWebView`) and your own variable declarations.

- `<Package>` expands once per package, in declaration order
- `<Package[N]>` picks one package by position, counting from 1, and doesn't fan out. An `N` of 0 or past the last package drops the key with a warning, like any other [indexed reference](#indexed-references)

`<Package>` and `<Package[2]>` in one key are two separate variables, so they multiply. So do `<Package>` and a variable declared from it (`Cache=<Package>\LocalCache`): a key using both pairs each package with every package's cache, mismatched pairs included. Use one or the other in a key.

The package detection is already there, so an entry never needs its own `DetectFile=<Package>`.

## PackageScaffolds / ExcludePackageScaffolds

The Package family's catalog is `package.ini` in the shared scaffold directory. It has one scaffold, `Baseline`, which is on by default:

```ini
[PackageScaffold: Baseline]
FileKeyBase=<Package>\AC|*|RECURSE
FileKeyBase=<Package>\Settings|*.log*
FileKeyBase=<Package>\SystemAppData\Helium|*.log*
FileKeyBase=<Package>\TempState|*|REMOVESELF
```

Each template is written out once per package, with `<Package>` replaced by that package's folder. `PackageScaffolds=` and `ExcludePackageScaffolds=` select from the catalog the same way the [WebView selectors](#webviewscaffolds--excludewebviewscaffolds) do: with no key you get `Baseline`, `All` and a named list work as they do there, and an empty `PackageScaffolds=` selects nothing.

The empty form is for an entry that shares its packages with another entry. `[Spotify - Downloaded Songs *]` declares the same package as `[Spotify *]`, which already gets the baseline, so it declares an empty `PackageScaffolds=` and the two entries don't clean the same folders twice. It keeps the package detection and its own `<Package>` keys.

## Hybrid apps

An app that ships both a Store package and a win32 install keeps both in one entry, and that needs nothing new. Write the win32 keys beside `Package=` as you would in any entry. Other scaffold families mix in too, and their roots can sit inside the package:

```ini
[Be My Eyes *]
Package=BeMyEyes.BeMyEyes_7yeb8xxw19svt
WebViewRoot=<Package>\LocalCache\Local\Be My Eyes\EBWebView
WebViewScaffolds=All
ExcludeWebViewScaffolds=WebStorage,LoginData
LangSecRef=3021
FileKeyBase=<Package>\LocalCache\Local\BeMyEyes\Be My Eyes\logs|*
```

The project's sources use `FileKeyBase=` for `<Package>` paths and `FileKey=` for win32 paths. The two do the same thing, but the split keeps a hybrid readable.

## UWPBuilder's leftover keys

Store apps used to have a builder of their own, UWPBuilder, which winapp2ool no longer has. A few of its keys mean nothing here. Since any key EntryBuilder doesn't know declares a variable, a leftover would otherwise become an unused variable and its scaffold keys would quietly go missing, so EntryBuilder refuses them instead: it drops the key, warns, and fails the run with exit code 1.

| Old key | Use instead |
|:-|:-|
| `SkipUWPFileKeys=` | An empty `PackageScaffolds=` |
| `WebViewPath=` | `WebViewRoot=` |
| `QtWebEnginePath=` | `QtWebEngineRoot=` |
| `QtWebEngineCachePath=` | `QtWebEngineCacheRoot=` |
| `ElectronPath=` | `ElectronRoot=` |
| `ElectronUpdaterPath=` | `ElectronUpdaterRoot=` |

The warning names the replacement:

```
SkipUWPFileKeys= in [Entry] is UWPBuilder's spelling; key dropped. Use an empty PackageScaffolds= instead
```

The retired `%Package%` spelling is refused too (see [The retired `%WebViewRoot%` spelling](#the-retired-webviewroot-spelling)).

## Real entries

These are entries from the project's sources, with the output EntryBuilder wrote for them in `Assembler\Generated\EntryBuilder\` (the file header is left out).

###### `[BreeZip *]`, the minimal entry (`Sources\EntryBuilder\B.ini`)

```ini
[BreeZip *]
Package=3138AweZip.AweZip_ffd303wmbhcjt
LangSecRef=3021
```

###### Output (`Generated\EntryBuilder\B.ini`)

```ini
[BreeZip *]
LangSecRef=3021
DetectFile=%LocalAppData%\Packages\3138AweZip.AweZip_ffd303wmbhcjt
FileKey1=%LocalAppData%\Packages\3138AweZip.AweZip_ffd303wmbhcjt\AC|*|RECURSE
FileKey2=%LocalAppData%\Packages\3138AweZip.AweZip_ffd303wmbhcjt\Settings|*.log*
FileKey3=%LocalAppData%\Packages\3138AweZip.AweZip_ffd303wmbhcjt\SystemAppData\Helium|*.log*
FileKey4=%LocalAppData%\Packages\3138AweZip.AweZip_ffd303wmbhcjt\TempState|*|REMOVESELF
```

The detection comes from `Package=` and the four FileKeys from the `Baseline` scaffold.

###### `[Samsung Bixby *]`, two packages, `<Package[N]>` and a win32 key (`Sources\EntryBuilder\S.ini`)

```ini
[Samsung Bixby *]
Package1=SAMSUNGELECTRONICSCO.LTD.Bixby_3c1yjt4zspk6g
Package2=SAMSUNGELECTRONICSCO.LTD.SamsungIntelligenceVoiceS_3c1yjt4zspk6g
LangSecRef=3021
FileKeyBase=<Package[1]>\LocalState\Samsung\Bixby\Logs|*.log
FileKeyBase=<Package[1]>\LocalState\Microsoft\AppCenter\Errors|*
FileKey=%ProgramData%\Samsung\Intelligence Voice Service\Logs|*
```

###### Output (`Generated\EntryBuilder\S.ini`)

```ini
[Samsung Bixby *]
LangSecRef=3021
DetectFile1=%LocalAppData%\Packages\SAMSUNGELECTRONICSCO.LTD.Bixby_3c1yjt4zspk6g
DetectFile2=%LocalAppData%\Packages\SAMSUNGELECTRONICSCO.LTD.SamsungIntelligenceVoiceS_3c1yjt4zspk6g
FileKey1=%LocalAppData%\Packages\SAMSUNGELECTRONICSCO.LTD.Bixby_3c1yjt4zspk6g\AC|*|RECURSE
FileKey2=%LocalAppData%\Packages\SAMSUNGELECTRONICSCO.LTD.Bixby_3c1yjt4zspk6g\LocalState\Microsoft\AppCenter\Errors|*
FileKey3=%LocalAppData%\Packages\SAMSUNGELECTRONICSCO.LTD.Bixby_3c1yjt4zspk6g\LocalState\Samsung\Bixby\Logs|*.log
FileKey4=%LocalAppData%\Packages\SAMSUNGELECTRONICSCO.LTD.Bixby_3c1yjt4zspk6g\Settings|*.log*
FileKey5=%LocalAppData%\Packages\SAMSUNGELECTRONICSCO.LTD.Bixby_3c1yjt4zspk6g\SystemAppData\Helium|*.log*
FileKey6=%LocalAppData%\Packages\SAMSUNGELECTRONICSCO.LTD.Bixby_3c1yjt4zspk6g\TempState|*|REMOVESELF
FileKey7=%LocalAppData%\Packages\SAMSUNGELECTRONICSCO.LTD.SamsungIntelligenceVoiceS_3c1yjt4zspk6g\AC|*|RECURSE
FileKey8=%LocalAppData%\Packages\SAMSUNGELECTRONICSCO.LTD.SamsungIntelligenceVoiceS_3c1yjt4zspk6g\Settings|*.log*
FileKey9=%LocalAppData%\Packages\SAMSUNGELECTRONICSCO.LTD.SamsungIntelligenceVoiceS_3c1yjt4zspk6g\SystemAppData\Helium|*.log*
FileKey10=%LocalAppData%\Packages\SAMSUNGELECTRONICSCO.LTD.SamsungIntelligenceVoiceS_3c1yjt4zspk6g\TempState|*|REMOVESELF
FileKey11=%ProgramData%\Samsung\Intelligence Voice Service\Logs|*
```

Each package gets its own detection and its own copy of the baseline. The two `<Package[1]>` keys land in the first package only, and the win32 `FileKey` passes through. WinappDebug sorts the FileKeys by value, so the app's own keys end up among the baseline ones.

###### `[Spotify - Downloaded Songs *]`, a package shared with another entry (`Sources\EntryBuilder\S.ini`)

```ini
[Spotify - Downloaded Songs *]
Package=SpotifyAB.SpotifyMusic_zpdnekdrzrea0
PackageScaffolds=
LangSecRef=3023
DetectFile=%AppData%\Spotify
FileKeyBase=<Package>\LocalState\Spotify\Storage|*|REMOVESELF
FileKey=%LocalAppData%\Spotify\Storage|*|REMOVESELF
```

###### Output (`Generated\EntryBuilder\S.ini`)

```ini
[Spotify - Downloaded Songs *]
LangSecRef=3023
DetectFile1=%AppData%\Spotify
DetectFile2=%LocalAppData%\Packages\SpotifyAB.SpotifyMusic_zpdnekdrzrea0
FileKey1=%LocalAppData%\Packages\SpotifyAB.SpotifyMusic_zpdnekdrzrea0\LocalState\Spotify\Storage|*|REMOVESELF
FileKey2=%LocalAppData%\Spotify\Storage|*|REMOVESELF
```

The empty `PackageScaffolds=` leaves the baseline to `[Spotify *]`. The package detection is still there, and WinappDebug's sort puts it after the hand-written `DetectFile`.

---

# Validation and Warnings

Warnings show in the menu output after a run and go to the winapp2ool log. Advisories, marked below, go to the log only. In the messages, `[Entry]` stands for the entry name and `X` / `N` for the value.

| Condition | Outcome | Message |
|:-|:-|:-|
| Two sections share a name | First definition kept, run fails (exit code 1) | `Duplicate EntryBuilder section [<name>] at <file> line <n> ignored; keeping the first definition, in <file>` |
| Both `Section` and `LangSecRef` declared | `LangSecRef` wins, `Section` dropped | `Both Section and LangSecRef present in [Entry], using LangSecRef` |
| Neither `Section` nor `LangSecRef` declared | Entry **skipped** | `No Section or LangSecRef in [Entry], skipping` |
| No package, no root keys and no FileKey/RegKey content in any form | Entry **skipped** | `[Entry] declares no Package, WebViewRoot, QtWebEngineRoot, ElectronRoot, FileKeyBase, or RegKeyBase — nothing to emit, skipping` |
| No detection (after `Root` inference) | Entry emitted, always-on | `[Entry] declares no detection (Detect / DetectFile / DetectOS); generated entry will be always-on` |
| A key other than `Skip=`, `Archive=` or a scaffold list has no value | Key dropped | `[Entry].FileKey has no value; key dropped` |
| `Archive=False` | Entry **not built**, run fails (exit code 1) | `[Entry] declares Archive=False; delete the Archive line to keep the entry live. The entry isn't built` |
| Both `Skip=` and `Archive=` | Entry **not built**, run fails (exit code 1) | `[Entry] declares both Skip and Archive; keep one of them. The entry isn't built` |
| Archived entries and no archive folder | Archived entries left out of both outputs, run fails (exit code 1) | `N archived entries left out of the output: no archive folder was given (-4d, or -3d for BrowserBuilder)` |
| `Default=` declared | Key dropped | `Default= declared in [Entry]; EntryBuilder never emits Default, ignoring` |
| `SpecialDetect=` declared | Key dropped | `SpecialDetect is deprecated; key dropped from [Entry]. Replace with Detect or DetectFile` |
| A key uses a retired `%...Root%` or `%Package%` spelling | Key dropped, run fails (exit code 1) | `` `%X%` in [Entry].FileKeyBase is now spelled `<X>`; key dropped `` |
| A key UWPBuilder used (`SkipUWPFileKeys=`, `WebViewPath=` and the other `...Path=` roots) | Key dropped, run fails (exit code 1) | `X= in [Entry] is UWPBuilder's spelling; key dropped. Use Y instead` |
| A `Package=` value holds `<` or `>` | Value used as written | `Package= value in [Entry] contains '<' or '>'; Package is never variable-expanded and the value is used literally` |
| `All` mixed with non-legacy scaffold names | Extras ignored | `WebViewScaffolds=All in [Entry] with redundant additional names (X, Y); ignoring` |
| Explicit scaffold list mixed with exclusions | Exclusions applied | `Both WebViewScaffolds and ExcludeWebViewScaffolds set in [Entry]; applying exclusions to explicit list` |
| Unknown scaffold name requested | Scaffold dropped | `Unknown WebView scaffold 'X' requested by [Entry], skipping` |
| A scaffold named in `…Scaffolds=` has no template whose root the entry declared | Scaffold produces nothing | `Scaffold 'X' selected by [Entry] produced no keys; its templates reference a root the entry did not declare` |
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

The scaffold-family messages read `QtWebEngine`, `Electron` or `Package` in place of `WebView` for those families. The log also records each `All` expansion: `WebViewScaffolds=All in [Entry]; expanded to N scaffold(s) from catalog, M legacy withheld`.

---

# Run Output

### The output file

WinappDebug normalizes the whole result with every repair and optimization on, so key order and numbering in the output follow WinappDebug's sorting, not your source or the order EntryBuilder generated the keys in. If that pass would change what an entry means rather than just how it's written, the run fails and leaves a `.prelint.ini` copy of the unlinted output beside the save target (see the [reconciliation guard](../winappdebug/README.md#the-reconciliation-guard)).

The output file begins with this comment header, where the first line counts its entries and the second names the generator:

```ini
; # of entries: 2,629
; This file is generated by winapp2ool's EntryBuilder
; Do not edit it directly
; The next build overwrites any changes made to it
;
; Source folder:
;   https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/Sources/EntryBuilder
;
; This file is not intended to be used standalone or with any cleaning software
; Latest winapp2.ini, in every flavor:
;   https://github.com/MoscaDotTo/Winapp2/releases/latest
;
; To add or fix an entry, see:
;   https://github.com/MoscaDotTo/Winapp2/blob/master/CONTRIBUTING.md
```

With `-split`, each per-letter file links its own source file instead of the folder and adds a line saying which letter it holds. This is the build's `A.ini`:

```ini
; # of entries: 370
; This file is generated by winapp2ool's EntryBuilder
; Do not edit it directly
; The next build overwrites any changes made to it
;
; This file contains entries for applications whose names begin with the letter A
; Source file:
;   https://github.com/MoscaDotTo/Winapp2/blob/master/Assembler/Sources/EntryBuilder/A.ini
;
; This file is not intended to be used standalone or with any cleaning software
; Latest winapp2.ini, in every flavor:
;   https://github.com/MoscaDotTo/Winapp2/releases/latest
;
; To add or fix an entry, see:
;   https://github.com/MoscaDotTo/Winapp2/blob/master/CONTRIBUTING.md
```

`#.ini` holds the entries beginning with numbers or symbols, and its link spells the file `%23.ini`, since a bare `#` would end the link. The links always point into the winapp2 repository, whatever source directory the run read.

The build runs EntryBuilder with `-split`, and a strict Combine joins the per-letter files with BrowserBuilder's `browsers.ini` to make the base winapp2.ini.

### The statistics summary

Each run also prints a statistics box: entry counts (read, generated, skipped, with packages, with scaffolds, with an inferred `Root`), how many shorthand keys were declared, the generated keys by category, and how many output keys came from templates (scaffolds and the `Base=` forms) against how many were passed through as written. The counts are taken before WinappDebug runs, so they describe what EntryBuilder generated rather than the saved file. A last section lists "tangled" entries, whose variables lean heavily on each other.

---

# Tips & Best Practices

### Start with pass-through, enrich incrementally

The DSL is opt-in. A raw winapp2.ini entry pasted into a source file comes out the same, apart from WinappDebug's normalization. Add shorthand when it removes repetition.

### Use single-value variables to name long literals

Even with nothing to fan out, `RegRoot=HKCU\Software\Vendor\Product` lets you write `<RegRoot>\Settings|Foo` instead of the full path on every key.

### `Root` is a contract, not just a name

Calling a variable `Root` says it's the app's anchor path and asks for detection from it. For the abbreviation without the detection key, pick another name (`DiskRoot`, `AppRoot`).

### Inline lists for one-offs, declarations for reuse

`<System32,SysWOW64>` is less work than declaring a variable for a list you use once. When the list shows up in a second key, or two copies of it need to co-vary, give it a name.

### Lean on co-variance for paired tokens

Use the same declared variable twice in one key when two positions have to match, like a version's folder and the version in its log file name. The two uses co-vary instead of multiplying.

### Default scaffold set is intentionally minimal

`Caches` and `Telemetry` are the only WebView scaffolds on by default. Anything that could remove state a user would notice (cookies, history, sessions, saved logins, web storage) needs an opt-in: `WebViewScaffolds=All` with exclusions, or a list naming the scaffolds you want. Legacy scaffolds are never picked implicitly, not even by `All`.

### Prefer `<Variable>` to numbered duplication

`RegKeyBase=<RegRoot>\<Versions>\Outlook` is easier to extend than seven hand-numbered `RegKey1..7` lines. When a new version ships, add it to `Versions=` and every key that uses the list grows with it.

---

# Troubleshooting

| Symptom | Likely Cause | Fix |
|:-|:-|:-|
| `No EntryBuilder source definitions found in: <dir>` and exit code 1 | Source directory is empty, missing, or contains no `*.ini` files with parseable sections | Verify the source directory in **Choose source directory** or via `-1d` |
| Entry is silently absent from the output | `Skip=` is set, or the entry was skipped by validation (`No Section or LangSecRef...` / `...nothing to emit, skipping`) | Check the run log for the corresponding message |
| `... is UWPBuilder's spelling; key dropped` and exit code 1 | The entry was written for UWPBuilder | Use the key the message names (see [UWPBuilder's leftover keys](#uwpbuilders-leftover-keys)) |
| `Duplicate EntryBuilder section [<name>] at <file> line <n> ignored; keeping the first definition, in <file>` | Two sections share a name (in any case), in one file or in two. Files are read in name order and only the first definition is used. The output is still written, but the run fails with exit code 1 | Merge the two sections, or rename one |
| `[Entry].X has no value; key dropped` | A key that needs a value, such as a root, a FileKey or a variable, was left empty | Fill in the value or delete the key |
| Output is missing expected scaffold FileKeys | Scaffold directory not found (`Scaffold directory not found: <dir>...`), scaffold name misspelled (`Unknown WebView scaffold 'X'...`), or scaffold excluded by `Exclude*Scaffolds=` | Check the `Loaded N … scaffold(s)` line for each family, and the warnings |
| A key holding `<WebViewRoot>` / `<QtWebEngineRoot>` / `<QtWebEngineCacheRoot>` / `<ElectronRoot>` / `<ElectronUpdaterRoot>` is missing from the output | The matching root key wasn't declared on the entry, which drops the key silently, or the placeholder was used in a `DetectFile`, where it is never substituted and drops as an undeclared token (a `Detect` keeps it literally) | Declare the root key, or write the path / a `<Variable>` directly in detection keys |
| `... is now spelled <WebViewRoot>; key dropped` and exit code 1 | The key still uses a retired `%WebViewRoot%`-style spelling | Respell it as the message says (see [The retired `%WebViewRoot%` spelling](#the-retired-webviewroot-spelling)) |
| Entry gained a `Detect`/`DetectFile` you didn't write | The entry declares a variable named `Root`, and detection inference is automatic | Intended? Delete your redundant hand-written detection. Not intended? Rename the variable (`DiskRoot`, `AppRoot`, ...) |
| Output RegKey contains a literal `<Name>` | The variable wasn't declared. The registry domain emits the literal text and logs `...passed through as literal` as an advisory | Declare the variable, or accept the literal if intentional |
| Output is missing a FileKey/DetectFile entirely | An undeclared `<Name>` token dropped the key (`...dropping key`), or a referenced variable had no values | Fix the typo, or declare the variable |
| `Variable 'X=' declared ... but never referenced` warning | Nothing uses the declaration, usually because of a typo in a `<X>` reference or in a key name | Search the entry for the intended reference or key and correct it |
| The run failed and left a `.prelint.ini` beside the output | The WinappDebug pass would have changed what an entry means. That's a bug in the sources or in EntryBuilder; the log lists each change | See the [reconciliation guard](../winappdebug/README.md#the-reconciliation-guard) |
| Output reorders keys differently from the source | WinappDebug normalization rewrites to canonical order and alphabetized numbering on every run | This is intentional; reformat the source to match if the diff bothers you |
| `Default=` was stripped from an entry | EntryBuilder never emits `Default=` | Add it in a flavor instead, the way the FluentCleaner flavor adds `Default=False` |
