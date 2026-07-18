# BrowserBuilder

**BrowserBuilder** is a winapp2ool module that generates winapp2.ini entries for web browsers using a small scripting interface. Rather than writing individual entries for every browser and every entry type by hand, you define browser metadata and entry templates once. BrowserBuilder then cross-applies them to produce a consistent set of entries for every supported browser automatically.

BrowserBuilder is primarily an internal devops tool used to maintain the browser sections of the winapp2.ini distribution. End users with browsers installed in non-standard locations or using portable installations can also use it to generate entries for those scenarios (see [Example 8](#example-8-generating-entries-for-a-portable-browser)).

### What does BrowserBuilder do?

BrowserBuilder reads two DSL files (`chromium.ini` and `gecko.ini`) from a single configurable source directory, each containing **BrowserInfo** sections describing individual browsers and **EntryScaffold** sections describing entry templates. It cross-applies every scaffold against every browser to produce a complete set of generated entries, applies Flavorizer corrections from the same directory to correct over- and under-coverage, normalizes the result with WinappDebug, and saves it to `browsers.ini`.

### Why BrowserBuilder?

- Consistency: every browser receives the same baseline coverage, generated from the same templates. No browser lags behind when a template location is discovered
- Scale: at the time of writing, 62 `BrowserInfo` sections × 35 `EntryScaffold` templates generate over 1,100 entries from two small ruleset files
- Low-cost maintenance: adding support for a new browser means writing one 3-5 line `BrowserInfo` section, no code changes; fixing a scaffold once fixes it for every browser
- Custom coverage: end users can generate entries for portable or non-standard installations that the distributed winapp2.ini cannot cover

---

# Table of Contents

1. [Requirements](#requirements)
2. [Quick Start](#quick-start)
3. [Menu Options](#menu-options)
4. [Source File Format](#source-file-format)
   - [BrowserInfo Sections](#browserinfo-sections)
   - [EntryScaffold Sections](#entryscaffold-sections)
   - [Template Variables](#template-variables)
   - [Special Case: Opera GX](#special-case-opera-gx)
5. [Post-Generation Flavorizing](#post-generation-flavorizing)
6. [Output Format](#output-format)
7. [Command-Line Arguments](#command-line-arguments)
   - [File Selection](#file-selection)
   - [Examples](#examples)
8. [Tips & Best Practices](#tips--best-practices)
9. [Troubleshooting](#troubleshooting)
10. [Usage Examples](#usage-examples)
    - [Example 1: Generating an Entry](#example-1-generating-an-entry)
    - [Example 2: The Cross-Product](#example-2-the-cross-product)
    - [Example 3: Multiple UserDataPaths](#example-3-multiple-userdatapaths)
    - [Example 4: TruncateDetect](#example-4-truncatedetect)
    - [Example 5: Registry Scaffolds and RequiresRegistryRoot](#example-5-registry-scaffolds-and-requiresregistryroot)
    - [Example 6: BrowserPath and the ProgramFiles Mirror](#example-6-browserpath-and-the-programfiles-mirror)
    - [Example 7: LocalDataPath (Gecko)](#example-7-localdatapath-gecko)
    - [Example 8: Generating Entries for a Portable Browser](#example-8-generating-entries-for-a-portable-browser)

---

# Requirements

- A source directory containing at least one of `chromium.ini` or `gecko.ini` with valid `BrowserInfo` and `EntryScaffold` sections
- Optionally, [flavor correction files](#post-generation-flavorizing) in the same directory, resolved by their fixed names

All input files are read from the single source directory. The individual files are not separately configurable; pointing BrowserBuilder at a directory is the entire configuration.

---

# Quick Start

### Common Workflow

1. Place `chromium.ini` and/or `gecko.ini` (and any flavor correction files) in one directory
2. From the winapp2ool main menu, select Entry Lab, and then select BrowserBuilder 
3. Set the source directory to that location (or launch winapp2ool from it; the default is the current directory)
4. Run: BrowserBuilder generates entries for all browsers defined in the rulesets, applies any flavor corrections found next to them, and saves the result to `browsers.ini`

---

# Menu Options

| Option | Effect | Notes |
|:-|:-|:-|
| Run (default) | Generate browser entries from the configured source directory | Requires at least one of `chromium.ini` or `gecko.ini` in the source directory |
| Choose source directory | Select the directory containing `chromium.ini`, `gecko.ini`, and flavor correction files | Default: current directory |
| Choose save target | Select the output file path | Default: `browsers.ini` in the current directory |
| Reset Settings | Restore all settings to their defaults | Only shown when settings have been changed |

The menu also displays the current source directory and save target below the options.

---

# Source File Format

BrowserBuilder reads standard ini files containing two types of sections: `BrowserInfo` sections and `EntryScaffold` sections. Sections of any other type are ignored with a warning.

```ini
[BrowserInfo: Browser Name]
Section=Category Name
UserDataPath=%LocalAppData%\BrowserName\User Data
RegistryRoot=HKCU\Software\BrowserName

[EntryScaffold: Entry Description]
FileKeyBase=%UserDataPath%\*\Cache|*|REMOVESELF
```

The section type prefixes are **case-sensitive** and must include the space after the colon: exactly `[BrowserInfo: ...]` and `[EntryScaffold: ...]`. A section named `[browserinfo: ...]` or `[BrowserInfo:...]` is ignored as invalid. Key names *within* sections are case-insensitive.

## BrowserInfo Sections

Each `BrowserInfo` section describes a single browser. The browser name from the section header is prepended to the entry name in every generated entry.

| Key | Effect | Notes |
|:-|:-|:-|
| `Section=<value>` | The `Section=` value for all entries generated for this browser | Required |
| `UserDataPath=<path>` | The browser's user data directory | Required; multiple keys allowed for browsers with more than one possible path |
| `RegistryRoot=<path>` | A registry path root for the browser | Optional; multiple keys allowed; used as `%RegistryRoot%` in `RegKeyBase` templates |
| `TruncateDetect` | Use the parent of `UserDataPath` for the `DetectFile` instead | No value needed; useful when the path's parent directory contains a version wildcard |
| `Skip` | Exclude this browser from generation | No value needed; the section is read but no entries are produced |

### Example

```ini
[BrowserInfo: Google Chrome]
Section=Google Chrome Web Browser
UserDataPath=%LocalAppData%\Google\Chrome\User Data
RegistryRoot=HKCU\Software\Google\Chrome
```

`Skip` and `TruncateDetect` are presence-based: providing the key at all (regardless of value) activates the behavior. `Skip=False` **will** cause the browser to be skipped.

Any other key in a `BrowserInfo` section produces a warning and is otherwise ignored.

## EntryScaffold Sections

Each `EntryScaffold` section is a template for one winapp2.ini entry type. The description from the section header is appended to the browser name in every generated entry.

| Key | Effect | Notes |
|:-|:-|:-|
| `FileKeyBase=<template>` | A FileKey template to generate for each browser | Multiple keys allowed; must contain a [template variable](#template-variables), or the line is silently discarded |
| `RegKeyBase=<template>` | A RegKey template to generate for each browser | Multiple keys allowed; use `%RegistryRoot%` for browser-specific registry paths |
| `RequiresRegistryRoot` | Skip this scaffold for any browser without a `RegistryRoot` | No value needed; prevents generating entries with no working keys |

### Example

```ini
[EntryScaffold: Web Browsing Cookies]
FileKeyBase=%UserDataPath%\*\Network|Cookies*;Device Bound Sessions*
```

For every `BrowserInfo` section, this scaffold produces an entry named `<browser name> Web Browsing Cookies *` with the corresponding `FileKey` values substituted in.

Any other key in an `EntryScaffold` section produces a warning and is otherwise ignored.

## Template Variables

| Variable | Replaced with | Available in |
|:-|:-|:-|
| `%UserDataPath%` | The browser's user data path from `BrowserInfo` | `FileKeyBase` |
| `%BrowserPath%` | The parent directory of `UserDataPath` | `FileKeyBase` |
| `%LocalDataPath%` | `UserDataPath` with `%AppData%` swapped for `%LocalAppData%` | `FileKeyBase` (Gecko only) |
| `%RegistryRoot%` | A registry root path from `BrowserInfo` | `RegKeyBase` |

**A `FileKeyBase` which contains none of these variables generates nothing.** There is no warning; the template is silently skipped for every browser. Likewise, `%LocalDataPath%` used in `chromium.ini` is silently skipped. If an expected FileKey is missing from your output, check the template's variable first.

### Variable behavior details

**`%UserDataPath%`** is substituted directly with the `UserDataPath` value. For browsers with multiple `UserDataPath` keys, one `DetectFile` and one set of FileKeys is generated per path. See [Example 3](#example-3-multiple-userdatapaths).

**`%BrowserPath%`** resolves to the parent directory of `UserDataPath`. For Chromium browsers, each `%BrowserPath%` template generates a *second* FileKey with `%LocalAppData%` or `%AppData%` replaced by `%ProgramFiles%`, covering system-wide (All Users) installations. See [Example 6](#example-6-browserpath-and-the-programfiles-mirror). Gecko browsers do not receive the `%ProgramFiles%` mirror.

**`%LocalDataPath%`** is a Gecko-only variable that takes `UserDataPath` and replaces `%AppData%` with `%LocalAppData%`, targeting the local (non-roaming) profile data directory. See [Example 7](#example-7-localdatapath-gecko).

**`%RegistryRoot%`** is substituted in `RegKeyBase` templates. One `RegKey` is generated per `RegistryRoot` key per `RegKeyBase` template. See [Example 5](#example-5-registry-scaffolds-and-requiresregistryroot).

## Special Case: Opera GX

Opera GX stores profile data directly in its user data directory and keeps additional profiles under `_side_profiles\*` rather than the usual `Default`/`Profile 1` folders. BrowserBuilder recognizes the browser name `Opera GX` and rewrites every `%UserDataPath%` template into two FileKeys: one against the user data directory itself and one against `_side_profiles\*`. If you diff the generated output, Opera GX entries will structurally differ from every other Chromium browser's; this is intentional.

---

# Post-Generation Flavorizing

After generating entries, BrowserBuilder applies Flavorizer to the output before saving. This corrects coverage issues that arise from the template-based generation approach. For example: removing entries for features a particular browser does not support, or adding vendor-specific keys that cannot be templated.

The flavor files are read from the source directory. The names are not configurable:

| Stage | Operation | File name (fixed) |
|:-|:-|:-|
| 1 | Section removals | `browser_section_removals.ini` |
| 2 | Key name removals | `browser_name_removals.ini` |
| 3 | Key value removals | `browser_value_removals.ini` |
| 4 | Section replacements | `browser_section_replacements.ini` |
| 5 | Key replacements | `browser_key_replacements.ini` |
| 6 | Additions | `browser_additions.ini` |

All flavor files are optional; missing files are silently skipped.

For additional context, see the [Flavorizer README](../transmute/Flavorizer/readme.md) and the [Transmute README's Usage Examples](../transmute/readme.md#usage-examples).

---

# Output Format

The output file (`browsers.ini` by default) contains standard winapp2.ini format entries. Each generated entry has:

- A section header in the form `[<browser name> <scaffold name> *]`
- A `Section=` key from the browser's `BrowserInfo`
- `DetectFile=` (or numbered `DetectFile1=`, `DetectFile2=` for multiple paths) pointing to the `UserDataPath` (or its parent, if `TruncateDetect` is set)
- Numbered `FileKey` and `RegKey` entries from the scaffold templates with variables substituted

Before saving, the entire output is run through **WinappDebug** with full optimizations forced: entries are alphabetized, keys are placed in standard winapp2.ini order, key numbering is normalized, and parameter lists are alphabetized. 

The saved file begins with a generated comment header:

```ini
; Version <yyMMdd>
; # of entries: #,###
; browsers.ini is generated by the Winapp2ool Browser Builder
; Entries in this file may be incomplete and are not intended to be used directly with any cleaning software
; They are utilized by winapp2ool to create the final winapp2.ini file for distribution
; If you are not maintaining winapp2.ini for distribution, you probably don't need this file!
; Refer to the Winapp2ool documentation for more information: https://github.com/MoscaDotTo/Winapp2/blob/master/winapp2ool/Readme.md
; You can find the complete winapp2.ini file here: https://raw.githubusercontent.com/MoscaDotTo/Winapp2/master/Non-CCleaner/Winapp2.ini
```

---

# Command-Line Arguments

BrowserBuilder accepts two file parameters: the source directory and the output file. There are no toggle arguments. The ruleset and flavor file names within the source directory are fixed and cannot be overridden from the command line.

## File Selection

| Arg | Effect | Default |
|:-|:-|:-|
| `-1d path` | Set the source directory containing `chromium.ini`, `gecko.ini`, and flavor correction files | Current directory |
| `-2d path` | Set the output file path | Current directory |
| `-2f name` | Set the output file name | `browsers.ini` |

## Examples

| Command | Effect |
|:-|:-|
| `winapp2ool -browserbuilder` | Generate from the rulesets in the current directory, saving to `browsers.ini` in the current directory |
| `winapp2ool -browserbuilder -1d "C:\Winapp2\Assembler\BrowserBuilder"` | Generate from the project's rulesets and flavor files, saving to `browsers.ini` in the current directory |
| `winapp2ool -browserbuilder -1d "C:\rulesets" -2d "C:\output" -2f generated.ini` | Generate from `C:\rulesets`, saving to `C:\output\generated.ini` |

---

# Tips & Best Practices

### Always Use Flavor Files

Raw generated output will have coverage issues: entries for features not supported by certain browsers, or missing browser-specific keys and entries. Flavor files are how these are corrected. They must sit in the source directory under their exact fixed names to be picked up.

### Skipping Discontinued Browsers

Use `Skip` in a `BrowserInfo` section to exclude a browser without deleting its data. This preserves the section for potential future re-activation.

### Multiple UserDataPaths

Provide multiple `UserDataPath` keys in a `BrowserInfo` section to support browsers with more than one possible profile location (e.g., a standalone install and a Microsoft Store package). BrowserBuilder generates a `DetectFile` and a full set of FileKeys for each path.

### TruncateDetect

Use `TruncateDetect` when `UserDataPath` contains a wildcard in its direct parent directory name (e.g., `%LocalAppData%\BraveSoftware\Brave-Browser*\User Data`). Some cleaning tools (notably CCleaner) do not support wildcards in `DetectFile` parent paths; truncating moves the detection up to the wildcarded directory itself. See [Example 4](#example-4-truncatedetect).

---

# Troubleshooting

| Symptom | Cause |
|:-|:-|
| "No valid generative rulesets found" | Neither `chromium.ini` nor `gecko.ini` in the source directory contains any parseable sections; check the source directory setting |
| Warning: "Invalid section found and ignored: [section]" | A section's name does not start with exactly `BrowserInfo:` or `EntryScaffold:`; the prefixes are case-sensitive |
| Warning: "Unexpected KeyType in [section]: [key]" | A `BrowserInfo` or `EntryScaffold` section contains a key name the module doesn't recognize; the key is ignored |
| Warning: "No valid UserDataPath key found in [browser]" | The `BrowserInfo` section is missing its `UserDataPath` key; the browser's entries will have no `DetectFile` and no FileKeys |
| Warning: "No valid Section key found in [browser]" | The `BrowserInfo` section is missing its `Section` key; generated entries will carry an empty `Section=` |
| Warning: "Skipping [browser] for [scaffold]: scaffold requires RegistryRoot" | The scaffold declares `RequiresRegistryRoot` and the browser has no `RegistryRoot` key; no entry is generated for that pairing |
| A browser is missing from the output entirely | It may have `Skip` set; check its `BrowserInfo` section |
| One entry is missing for one browser | Likely a `RequiresRegistryRoot` scaffold and a browser without a `RegistryRoot` |
| A scaffold key is missing from every entry | The `FileKeyBase` contains no template variable (or uses `%LocalDataPath%` in `chromium.ini`) and was silently discarded |
| Flavor corrections aren't being applied | The flavor file's name doesn't exactly match its fixed name, or it isn't in the source directory. Missing files are skipped silently |
| Generated entries are incomplete or wrong | Flavor files correct coverage issues; ensure the appropriate flavor files are present in the source directory |

---

# Usage Examples

The examples below use the real rulesets from the winapp2.ini project, found [here](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/BrowserBuilder). Relevant excerpts are provided on this page, and every output shown is the module's real output. Unless stated otherwise, the source directory contains only the rulesets (no flavor files), and the [generated comment header](#output-format) is omitted from outputs for brevity.

## Example 1: Generating an Entry

**Context**

The simplest possible case: one browser crossed with one scaffold. Chromium keeps backup copies of the bookmarks file in each profile, and we want an entry that cleans them.

**Intent**

We want to generate a `Bookmark Backups` entry for Chromium.

**Files**

###### **Ruleset (`chromium.ini`), relevant sections**

```ini
[BrowserInfo: Chromium]
Section=Chromium Web Browser
UserDataPath=%LocalAppData%\Chromium\User Data
RegistryRoot=HKCU\Software\Chromium

[EntryScaffold: Bookmark Backups]
FileKeyBase=%UserDataPath%\*|Bookmarks.bak
```

**Command**

```
winapp2ool -browserbuilder -1d C:\rulesets
```

**Output**

###### **Output file (`browsers.ini`), relevant entry**

```ini
[Chromium Bookmark Backups *]
Section=Chromium Web Browser
DetectFile=%LocalAppData%\Chromium\User Data
FileKey1=%LocalAppData%\Chromium\User Data\*|Bookmarks.bak
```

**Explanation**

- The entry name is the browser name + the scaffold name + ` *`: `Chromium` + `Bookmark Backups` → `[Chromium Bookmark Backups *]`
- `Section=` is copied from the `BrowserInfo`
- `DetectFile=` is the browser's `UserDataPath`
- The `FileKeyBase` template becomes `FileKey1` with `%UserDataPath%` substituted
- The `RegistryRoot` key goes unused here. This scaffold has no `RegKeyBase` templates

---

## Example 2: The Cross-Product

**Context**

The point of BrowserBuilder is that browsers and templates multiply: every scaffold is applied to every browser in the same ruleset. Adding one browser adds a full family of entries; adding one scaffold extends every browser at once.

**Intent**

We want to see what two browsers × two scaffolds produce.

**Files**

###### **Ruleset (`chromium.ini`)**

```ini
[BrowserInfo: CCleaner Browser]
Section=CCleaner Web Browser
UserDataPath=%LocalAppData%\CCleaner Browser\User Data
RegistryRoot=HKCU\Software\Piriform\Browser

[BrowserInfo: Chromium]
Section=Chromium Web Browser
UserDataPath=%LocalAppData%\Chromium\User Data
RegistryRoot=HKCU\Software\Chromium

[EntryScaffold: Autoplay Preferences]
FileKeyBase=%UserDataPath%\MEIPreload|*|REMOVESELF

[EntryScaffold: Bookmark Backups]
FileKeyBase=%UserDataPath%\*|Bookmarks.bak
```

**Command**

```
winapp2ool -browserbuilder -1d C:\rulesets
```

**Output**

###### **Output file (`browsers.ini`)**

```ini
[CCleaner Browser Autoplay Preferences *]
Section=CCleaner Web Browser
DetectFile=%LocalAppData%\CCleaner Browser\User Data
FileKey1=%LocalAppData%\CCleaner Browser\User Data\MEIPreload|*|REMOVESELF

[CCleaner Browser Bookmark Backups *]
Section=CCleaner Web Browser
DetectFile=%LocalAppData%\CCleaner Browser\User Data
FileKey1=%LocalAppData%\CCleaner Browser\User Data\*|Bookmarks.bak

[Chromium Autoplay Preferences *]
Section=Chromium Web Browser
DetectFile=%LocalAppData%\Chromium\User Data
FileKey1=%LocalAppData%\Chromium\User Data\MEIPreload|*|REMOVESELF

[Chromium Bookmark Backups *]
Section=Chromium Web Browser
DetectFile=%LocalAppData%\Chromium\User Data
FileKey1=%LocalAppData%\Chromium\User Data\*|Bookmarks.bak
```

**Explanation**

- 2 browsers × 2 scaffolds = 4 entries, each personalized with its browser's paths and category

---

## Example 3: Multiple UserDataPaths

**Context**

Mozilla Firefox can be installed standalone (profiles under `%AppData%`) or from the Microsoft Store (profiles under the package's `LocalCache`). Both locations need coverage, and a machine may have either.

**Intent**

We want one entry per scaffold that detects and cleans both install types.

**Files**

###### **Ruleset (`gecko.ini`)**

```ini
[BrowserInfo: Mozilla Firefox]
Section=Mozilla Firefox Web Browser
UserDataPath=%AppData%\Mozilla\Firefox\Profiles
UserDataPath=%LocalAppData%\Packages\Mozilla.Firefox_n80bbvh6b1yt2\LocalCache\Roaming\Mozilla\Firefox\Profiles

[EntryScaffold: Autocomplete History]
FileKeyBase=%UserDataPath%\*|formhistory*
```

**Command**

```
winapp2ool -browserbuilder -1d C:\rulesets
```

**Output**

###### **Output file (`browsers.ini`)**

```ini
[Mozilla Firefox Autocomplete History *]
Section=Mozilla Firefox Web Browser
DetectFile1=%AppData%\Mozilla\Firefox\Profiles
DetectFile2=%LocalAppData%\Packages\Mozilla.Firefox_n80bbvh6b1yt2\LocalCache\Roaming\Mozilla\Firefox\Profiles
FileKey1=%AppData%\Mozilla\Firefox\Profiles\*|formhistory*
FileKey2=%LocalAppData%\Packages\Mozilla.Firefox_n80bbvh6b1yt2\LocalCache\Roaming\Mozilla\Firefox\Profiles\*|formhistory*
```

**Explanation**

- With multiple `UserDataPath` keys, the `DetectFile`s become numbered: `DetectFile1`, `DetectFile2`, one per path
- Every `FileKeyBase` template is expanded once per path, so one template line produced two FileKeys

---

## Example 4: TruncateDetect

**Context**

Brave's stable, beta, and nightly channels store data under `Brave-Browser`, `Brave-Browser-Beta`, and `Brave-Browser-Nightly` respectively, which the ruleset covers with a single wildcard: `%LocalAppData%\BraveSoftware\Brave-Browser*\User Data`. But some cleaning tools reject a wildcard in a `DetectFile`'s parent path, so detecting on the full `UserDataPath` would break.

**Intent**

We want the generated `DetectFile` to point at the wildcarded parent directory instead of the `User Data` folder inside it.

**Files**

###### **Ruleset (`chromium.ini`)**

```ini
[BrowserInfo: Brave]
Section=Brave Web Browser
TruncateDetect=True
UserDataPath=%LocalAppData%\BraveSoftware\Brave-Browser*\User Data
RegistryRoot=HKCU\Software\BraveSoftware\Brave-Browser
RegistryRoot=HKCU\Software\BraveSoftware\Brave-Browser-Beta
RegistryRoot=HKCU\Software\BraveSoftware\Brave-Browser-Nightly

[EntryScaffold: Bookmark Backups]
FileKeyBase=%UserDataPath%\*|Bookmarks.bak
```

**Command**

```
winapp2ool -browserbuilder -1d C:\rulesets
```

**Output**

###### **Output file (`browsers.ini`)**

```ini
[Brave Bookmark Backups *]
Section=Brave Web Browser
DetectFile=%LocalAppData%\BraveSoftware\Brave-Browser*
FileKey1=%LocalAppData%\BraveSoftware\Brave-Browser*\User Data\*|Bookmarks.bak
```

**Explanation**

- The `DetectFile` now points at the parent path, `Brave-Browser*`, rather than `Brave-Browser*\User Data`

---

## Example 5: Registry Scaffolds and RequiresRegistryRoot

**Context**

Chromium browsers record pinned-tab state in the registry under the browser vendor's root key. Not every browser in the ruleset has a known registry root and generating a pinned-tabs entry with no keys in it would produce a broken entry.

**Intent**

We want a `Pinned Tabs` entry for every browser with a known registry root, and no entry at all for the rest.

**Files**

###### **Ruleset (`chromium.ini`)**

```ini
[BrowserInfo: Arc]
Section=Arc Web Browser
UserDataPath=%LocalAppData%\Packages\TheBrowserCompany.Arc_ttt1ap7aakyb4\LocalCache\Local\Arc\User Data

[BrowserInfo: Brave]
Section=Brave Web Browser
TruncateDetect=True
UserDataPath=%LocalAppData%\BraveSoftware\Brave-Browser*\User Data
RegistryRoot=HKCU\Software\BraveSoftware\Brave-Browser
RegistryRoot=HKCU\Software\BraveSoftware\Brave-Browser-Beta
RegistryRoot=HKCU\Software\BraveSoftware\Brave-Browser-Nightly

[EntryScaffold: Pinned Tabs]
RequiresRegistryRoot=True
RegKeyBase=%RegistryRoot%\PreferenceMACs\Default|pinned_tabs
```

**Command**

```
winapp2ool -browserbuilder -1d C:\rulesets
```

**Output**

###### **Output file (`browsers.ini`)**

```ini
[Brave Pinned Tabs *]
Section=Brave Web Browser
DetectFile=%LocalAppData%\BraveSoftware\Brave-Browser*
RegKey1=HKCU\Software\BraveSoftware\Brave-Browser\PreferenceMACs\Default|pinned_tabs
RegKey2=HKCU\Software\BraveSoftware\Brave-Browser-Beta\PreferenceMACs\Default|pinned_tabs
RegKey3=HKCU\Software\BraveSoftware\Brave-Browser-Nightly\PreferenceMACs\Default|pinned_tabs
```

There is no `[Arc Pinned Tabs *]` in the output. During the run, the module reports:

```
Skipping Arc for Pinned Tabs: scaffold requires RegistryRoot
```

**Explanation**

- One `RegKey` is generated per `RegistryRoot` per `RegKeyBase` template: Brave's three roots × one template = `RegKey1`-`RegKey3`
- `RequiresRegistryRoot` excludes browsers without a `RegistryRoot` from this scaffold entirely, with a warning naming the browser and scaffold

---

## Example 6: BrowserPath and the ProgramFiles Mirror

**Context**

Some browser files live *outside* the user data directory (e.g. installer logs, setup metrics). When a Chromium browser is installed "for all users", those application files live under `%ProgramFiles%` instead of the per-user location.

**Intent**

We want the Telemetry entry to clean the application directory for both per-user and all-users installations, without writing both variants into the scaffold.

**Files**

###### **Ruleset (`chromium.ini`)**

```ini
[BrowserInfo: Chromium]
Section=Chromium Web Browser
UserDataPath=%LocalAppData%\Chromium\User Data
RegistryRoot=HKCU\Software\Chromium

[EntryScaffold: Telemetry]
FileKeyBase=%BrowserPath%\Application|debug.log
FileKeyBase=%BrowserPath%\Application\SetupMetrics|*|REMOVESELF
```

**Command**

```
winapp2ool -browserbuilder -1d C:\rulesets
```

**Output**

###### **Output file (`browsers.ini`), trimmed to match above**

```ini
[Chromium Telemetry *]
Section=Chromium Web Browser
DetectFile=%LocalAppData%\Chromium\User Data
FileKey1=%LocalAppData%\Chromium\Application|debug.log
FileKey2=%LocalAppData%\Chromium\Application\SetupMetrics|*|REMOVESELF
; ... 
FileKey20=%ProgramFiles%\Chromium\Application|debug.log
FileKey21=%ProgramFiles%\Chromium\Application\SetupMetrics|*|REMOVESELF
```

**Explanation**

- `%BrowserPath%` is the parent of `UserDataPath`. Here: `%LocalAppData%\Chromium`
- Each `%BrowserPath%` template produced two `FileKeys`: the substituted path, and a mirror with the user-profile variable swapped for `%ProgramFiles%`
- The mirror is Chromium-only; Gecko browsers get just the direct substitution

---

## Example 7: LocalDataPath (Gecko)

**Context**

Gecko browsers split their profile data: settings and history roam under `%AppData%`, while caches and other machine-local data sit in a mirrored path under `%LocalAppData%`. Cache-cleaning templates need to target the local mirror, but `BrowserInfo` sections only declare the roaming path.

**Intent**

We want the Caches scaffold to clean the `%LocalAppData%` mirror without every browser declaring a second path.

**Files**

###### **Ruleset (`gecko.ini`)**

```ini
[BrowserInfo: Waterfox]
Section=Waterfox Web Browser
UserDataPath=%AppData%\Waterfox\Profiles

[EntryScaffold: Caches]
FileKeyBase=%LocalDataPath%\*\*cache*|*|REMOVESELF
FileKeyBase=%LocalDataPath%\*\chrome_debugger_profile|*|REMOVESELF
FileKeyBase=%LocalDataPath%\*\thumbnails|*|REMOVESELF
FileKeyBase=%UserDataPath%\*|*.corrupt|RECURSE
FileKeyBase=%UserDataPath%\*|AlternateServices.txt;notificationstore.json;parent.lock;serviceworker.txt;webappsstore.sqlite;cert9.db;ClientAuthRememberList.txt;SiteSecurityServiceState*
FileKeyBase=%UserDataPath%\*\notificationstore|*
FileKeyBase=%UserDataPath%\*\security_state|*
FileKeyBase=%UserDataPath%\*\shader-cache|*
FileKeyBase=%UserDataPath%\*\storage\temporary|*|RECURSE
```

**Command**

```
winapp2ool -browserbuilder -1d C:\rulesets
```

**Output**

###### **Output file (`browsers.ini`)**

```ini
[Waterfox Caches *]
Section=Waterfox Web Browser
DetectFile=%AppData%\Waterfox\Profiles
FileKey1=%AppData%\Waterfox\Profiles\*|*.corrupt|RECURSE
FileKey2=%AppData%\Waterfox\Profiles\*|AlternateServices.txt;cert9.db;ClientAuthRememberList.txt;notificationstore.json;parent.lock;serviceworker.txt;SiteSecurityServiceState*;webappsstore.sqlite
FileKey3=%AppData%\Waterfox\Profiles\*\notificationstore|*
FileKey4=%AppData%\Waterfox\Profiles\*\security_state|*
FileKey5=%AppData%\Waterfox\Profiles\*\shader-cache|*
FileKey6=%AppData%\Waterfox\Profiles\*\storage\temporary|*|RECURSE
FileKey7=%LocalAppData%\Waterfox\Profiles\*\*cache*|*|REMOVESELF
FileKey8=%LocalAppData%\Waterfox\Profiles\*\chrome_debugger_profile|*|REMOVESELF
FileKey9=%LocalAppData%\Waterfox\Profiles\*\thumbnails|*|REMOVESELF
```

**Explanation**

- `%LocalDataPath%` took the declared `UserDataPath` and swapped `%AppData%` → `%LocalAppData%`, producing FileKey7-9
- The `DetectFile` still only points at the roaming path
- The keys and their parameters have been reordered relative to their templates as a result of BrowserBuilder running its output through WinappDebug

---

## Example 8: Generating Entries for a Portable Browser

**Context**

You run Google Chrome Portable from `D:\PortableApps`. The distributed winapp2.ini doesn't cover this: its Chrome entries detect and clean the standard install locations. BrowserBuilder can generate entries for *your* install path, and because scaffolds are just text, you can copy whichever ones you care about from the [project's chromium.ini](https://github.com/MoscaDotTo/Winapp2/blob/master/Assembler/BrowserBuilder/chromium.ini).

**Intent**

We want history and cookie cleaning entries for a portable Chrome, then we want them merged into our local winapp2.ini.

**Files**

###### **Ruleset (`C:\myrules\chromium.ini`)**

```ini
[BrowserInfo: Chrome Portable]
Section=Chrome Portable Web Browser
UserDataPath=D:\PortableApps\GoogleChromePortable\Data\profile

[EntryScaffold: Web Browsing History]
FileKeyBase=%UserDataPath%\*|History*;Visited Links*;Top Sites*;Network Action Predictor*;shortcuts*

[EntryScaffold: Web Browsing Cookies]
FileKeyBase=%UserDataPath%\*\Network|Cookies*;Device Bound Sessions*
```

**Commands**

```
winapp2ool -browserbuilder -1d C:\myrules -2d C:\myrules
winapp2ool -transmute -add -1f winapp2.ini -2d C:\myrules -2f browsers.ini -3f winapp2.ini
```

**Output**

###### **Generated file (`C:\myrules\browsers.ini`) after the first command**

```ini
[Chrome Portable Web Browsing Cookies *]
Section=Chrome Portable Web Browser
DetectFile=D:\PortableApps\GoogleChromePortable\Data\profile
FileKey1=D:\PortableApps\GoogleChromePortable\Data\profile\*\Network|Cookies*;Device Bound Sessions*

[Chrome Portable Web Browsing History *]
Section=Chrome Portable Web Browser
DetectFile=D:\PortableApps\GoogleChromePortable\Data\profile
FileKey1=D:\PortableApps\GoogleChromePortable\Data\profile\*|History*;Network Action Predictor*;shortcuts*;Top Sites*;Visited Links*
```

**Explanation**

Two steps are chained:

| Step | Module | Effect |
|:-|:-|:-|
| 1 | BrowserBuilder | Generates the two entries above from the hand-written ruleset |
| 2 | Transmute (Add) | Merges the generated entries into your local `winapp2.ini` in place |

###### Notes

- Template variables accept absolute paths and also environment variables
- Because your additions live in the ruleset, you can regenerate and re-merge after each winapp2.ini update

