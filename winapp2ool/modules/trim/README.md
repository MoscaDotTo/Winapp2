# Trim

**Trim** is a winapp2ool module that removes irrelevant entries from winapp2.ini based on the current machine. Each entry in winapp2.ini declares detection criteria: registry keys, file system paths, operating system version requirements, or named application identifiers, that indicate whether the software it targets is present on a given system. Trim evaluates those criteria and removes any entry whose software is not detected, producing a smaller, faster winapp2.ini containing only entries relevant to the current machine.

### What does Trim do?

Trim reads a winapp2.ini file (from disk or downloaded directly from GitHub), checks every entry's detection criteria against the machine it is running on, and saves a copy containing only the entries whose targeted software was found. 

### Why Trim?

- Performance: CCleaner slowly evaluates every entry in winapp2.ini when it starts; a trimmed file dramatically improves its startup time
- Automation: `winapp2ool -trim -d` downloads the latest database and produces a machine-specific copy in one command

---

# Table of Contents

1. [Requirements](#requirements)
2. [Quick Start](#quick-start)
3. [Menu Options](#menu-options)
4. [How Trimming Works](#how-trimming-works)
   - [Evaluation Order](#evaluation-order)
   - [DetectOS](#detectos)
   - [Detect](#detect)
   - [DetectFile](#detectfile)
   - [SpecialDetect](#specialdetect)
   - [Entries Without Detection Keys](#entries-without-detection-keys)
   - [Includes and Excludes](#includes-and-excludes)
   - [Environment Variables](#environment-variables)
   - [VirtualStore Handling](#virtualstore-handling)
5. [Output Formatting](#output-formatting)
6. [Command-Line Arguments](#command-line-arguments)
   - [Toggles](#toggles)
   - [File Selection](#file-selection)
   - [Examples](#examples)
7. [Tips & Best Practices](#tips--best-practices)
8. [Troubleshooting](#troubleshooting)
9. [Usage Examples](#usage-examples)
   - [Example 1: Anatomy of a Trim](#example-1-anatomy-of-a-trim)
   - [Example 2: Trimming the Full Database](#example-2-trimming-the-full-database)
   - [Example 3: Downloading and Trimming in One Step](#example-3-downloading-and-trimming-in-one-step)
   - [Example 4: Never-Trim and Always-Trim Overrides](#example-4-never-trim-and-always-trim-overrides)
   - [Example 5: VirtualStore Key Generation](#example-5-virtualstore-key-generation)

---

# Requirements

- A `winapp2.ini` file to trim, **or** an internet connection with downloading enabled (`-d`)

---

# Quick Start

### Common Workflow (local file)

1. Place your `winapp2.ini` in the same directory as winapp2ool
2. Open Trim from the winapp2ool main menu
3. Run. Trim evaluates every entry and saves the result

By default, the trimmed file overwrites the input `winapp2.ini`. Use **Choose save target** to save to a different path first.

### Common Workflow (download and trim)

1. Open Trim
2. Enable **Toggle downloading** to use the latest winapp2.ini from GitHub
3. Run. Trim downloads and trims in one step

---

# Menu Options

| Option | Effect | Notes |
|:-|:-|:-|
| Run (default) | Trim winapp2.ini using current settings | |
| Toggle downloading | Enable or disable downloading the latest winapp2.ini from GitHub as the input | Default: `False`; unavailable in offline mode |
| Toggle include list | Enable or disable the includes file override | Default: `False` |
| Toggle exclude list | Enable or disable the excludes file override | Default: `False` |
| Choose winapp2.ini | Select a different local winapp2.ini to trim | Only shown when downloading is disabled |
| Choose save target | Select a different output file path | Default: overwrites the input file |
| Choose includes file | Select the includes file | Only shown when include list is enabled |
| Choose excludes file | Select the excludes file | Only shown when exclude list is enabled |
| Reset Settings | Restore all settings to their defaults | Only shown when settings have been changed |

---

# How Trimming Works

## Evaluation Order

Each entry is evaluated independently, in the following order, until one yields a result.

1. **Include list** (if enabled): entry name listed → **kept**
2. **Exclude list** (if enabled): entry name listed → **removed**
3. **DetectOS**: present, not satisfied → **removed**
4. **Detect**: any Detect registry path exists → **kept**
5. **DetectFile**: any DetectFile path exists → **kept**
6. **SpecialDetect**: any matches → **kept**
7. Entry has *only* a DetectOS key and it was satisfied in step 3 → **kept**
8. Entry has no detection keys of any kind → **kept**
9. Otherwise → **removed**

## DetectOS

`DetectOS` declares a Windows version requirement as a `major.minor` version number and takes one of three forms:

| Form | Meaning | Entry is kept when |
|:-|:-|:-|
| `DetectOS=6.1\|` | Minimum version | Windows version ≥ 6.1 |
| `DetectOS=\|6.1` | Maximum version | Windows version ≤ 6.1 |
| `DetectOS=6.1\|10.0` | Version range | 6.1 ≤ Windows version ≤ 10.0 |

If the requirement is not satisfied, the entry is removed immediately without evaluating any other detection keys. If it is satisfied and the entry has no other detection keys, the entry is kept.

## Detect

- Paths under `HKLM\Software` are additionally checked under `HKLM\SOFTWARE\WOW6432Node`, so entries for 32-bit applications on 64-bit Windows are retained correctly
- If reading a key raises a permission error, the key is assumed to exist and the entry is retained

## DetectFile

`DetectFile` specifies a file system path checked for existence. The path may resolve to either a file or a directory.

- Wildcards (`*`) are supported anywhere in the path. Each wildcard segment is expanded against the real directories present on the system; the entry is kept if any expansion resolves to an existing file or directory. 
- Paths under `%ProgramFiles%` are checked in both the native Program Files directory and the 32-bit `Program Files (x86)` directory. See [Environment Variables](#environment-variables)
- If a directory cannot be read due to permissions, the target is assumed to exist and the entry is retained

## SpecialDetect

`SpecialDetect` is a deprecated CCleaner variable retained only to support very old winapp2.ini files. New entries do not use it. Trim recognizes four values:

| Value | What is checked |
|:-|:-|
| `DET_CHROME` | A hardcoded list of ~26 Chromium-family browser paths and registry keys |
| `DET_MOZILLA` | `%AppData%\Mozilla\Firefox` |
| `DET_THUNDERBIRD` | `%AppData%\Thunderbird` |
| `DET_OPERA` | `%AppData%\Opera Software` |

If any of the associated paths or keys exists on the system, the entry is kept.

## Entries Without Detection Keys

Entries that have no detection keys of any type are always kept. 

## Includes and Excludes

Trim supports two override files that bypass detection evaluation entirely. Both checks run **before** any detection key is evaluated.

**Include list** (`includes.ini` by default): Any entry whose name appears in this file is always retained, regardless of whether its detection criteria are satisfied. Enable with **Toggle include list** in the menu or `-includes` on the command line.

**Exclude list** (`excludes.ini` by default): Any entry whose name appears in this file is always removed, regardless of whether its detection criteria are satisfied. Enable with **Toggle exclude list** in the menu or `-excludes` on the command line.

Each file is a plain ini file where section names are the entry names to match:

```ini
[Some Application *]
[Another Application *]
```

The key-value content of each section is ignored and need not be included. 

See [Example 4](#example-4-never-trim-and-always-trim-overrides) 

## Environment Variables

Trim expands standard Windows environment variables in `DetectFile` paths, plus several CCleaner-specific variables that do not exist in the Windows environment:

| Variable | Windows XP | Windows Vista and later |
|:-|:-|:-|
| `%Documents%` | `%UserProfile%\My Documents` | `%UserProfile%\Documents` |
| `%CommonAppData%` | `%AllUsersProfile%\Application Data` | `%AllUsersProfile%` |
| `%LocalLowAppData%` | `%UserProfile%\AppData\LocalLow` | `%UserProfile%\AppData\LocalLow` |
| `%Pictures%` | `%UserProfile%\My Documents\My Pictures` | `%UserProfile%\Pictures` |
| `%Music%` | `%UserProfile%\My Documents\My Music` | `%UserProfile%\Music` |
| `%Video%` | `%UserProfile%\My Documents\My Videos` | `%UserProfile%\Videos` |

`%ProgramFiles%` covers both the native Program Files directory and `Program Files (x86)` on 64-bit systems, and the equivalent applies to `HKLM\Software` registry paths via `WOW6432Node`.

**Malformed variables:** a path whose environment variable is malformed (e.g. a value ending at the closing `%` with no path after it) prints an error and pauses:

```
Error: <path> contains a malformatted environment variable and has been ignored
The associated entry will be retained in the final output file
Press any key to continue
```

## VirtualStore Handling

On some systems, particularly those upgraded from older versions of Windows, pre-UAC applications may have had their writes to protected locations redirected into the user's VirtualStore. For entries that pass detection, Trim scans the entry's `FileKey`, `RegKey`, and `ExcludeKey` values and generates additional keys covering the corresponding VirtualStore locations if they exist:

| Key type | Original location | VirtualStore counterpart |
|:-|:-|:-|
| FileKey / ExcludeKey | `%ProgramFiles%` | `%LocalAppData%\VirtualStore\Program Files*` |
| FileKey / ExcludeKey | `%CommonAppData%` | `%LocalAppData%\VirtualStore\ProgramData` |
| FileKey / ExcludeKey | `%CommonProgramFiles%` | `%LocalAppData%\VirtualStore\Program Files*\Common Files` |
| FileKey / ExcludeKey / RegKey | `HKLM\Software` | `HKCU\Software\Classes\VirtualStore\MACHINE\SOFTWARE` |

On most systems no VirtualStore keys are generated at all. When keys are added, the entry's keys are renumbered. See [Example 5](#example-5-virtualstore-key-generation).

---

# Output Formatting

Trim always writes its output as a fully formatted winapp2.ini file:

- Entries are sorted alphabetically and grouped into the winapp2.ini ordering for the selected flavor 
- The winapp2.ini preamble comments are regenerated at the top of the file: the `; Version:` line is carried over from the input file (or written as `; Version: 000000` if the input had none), and the `; # of entries:` line is updated to the post-trim count
- **Comments are not preserved.** Comments in the input file do not appear in the output. Since the default save target overwrites the input file, trimming an annotated custom file in place will lose your comments
- Key order within entries is unchanged, except that entries receiving [VirtualStore keys](#virtualstore-handling) have their keys renumbered

---

# Command-Line Arguments

Invoke Trim from the command line with `winapp2ool -trim`.

Command-line runs always start from Trim's default settings. Settings saved from the menu (alternate file paths, enabled include/exclude lists) are not applied to CLI runs. Pass the appropriate flags each time you invoke Trim from the command line.

### Toggles

| Arg | Effect |
|:-|:-|
| `-d` | Download the latest winapp2.ini from GitHub as the input  |
| `-includes` | Enable the includes file: entries named in it are never trimmed |
| `-excludes` | Enable the excludes file: entries named in it are always trimmed |

### File Selection

| Arg | Effect | Default |
|:-|:-|:-|
| `-1d path` | Set the input winapp2.ini directory | Current directory |
| `-1f name` | Set the input winapp2.ini file name | `winapp2.ini` |
| `-2d path` | Set the includes file directory | Current directory |
| `-2f name` | Set the includes file name | `includes.ini` |
| `-3d path` | Set the output directory | Current directory |
| `-3f name` | Set the output file name | `winapp2.ini` (overwrites input) |
| `-4d path` | Set the excludes file directory | Current directory |
| `-4f name` | Set the excludes file name | `excludes.ini` |

### Examples

| Command | Effect |
|:-|:-|
| `winapp2ool -trim` | Trim `winapp2.ini` in the current directory, overwriting it in place |
| `winapp2ool -trim -3f winapp2-trimmed.ini` | Trim and save the result to a new file |
| `winapp2ool -trim -d` | Download the latest winapp2.ini and trim it |
| `winapp2ool -trim -d -3f winapp2-trimmed.ini` | Download, trim, and save to a named file |
| `winapp2ool -trim -includes -2f keepers.ini` | Trim, but never remove entries named in `keepers.ini` |
| `winapp2ool -trim -3f trimmed.ini -s` | Trim silently and exit |

---

# Tips & Best Practices

### Save to a Different File First

The default output overwrites the input `winapp2.ini`, and [comments are not preserved](#output-formatting). Save to a different file (e.g. `winapp2-trimmed.ini`) the first time so you can verify the results before replacing your working copy.

### Re-trim After Changes

A trimmed file reflects both the contents of winapp2.ini and the installed software at the time it was run. After updating winapp2.ini or installing/uninstalling software, re-trim to keep the file accurate. Newly installed software will not be covered by a stale trimmed file.

### Download and Trim 

Use the **Toggle downloading** option or the `-d` flag to download the latest winapp2.ini and trim it in a single operation. 

### Trimmed Files and CCleaner Performance

CCleaner evaluates every entry in winapp2.ini when it starts. Trimming the file dramatically improves CCleaner's startup time. CC7Patcher also supports trimming before patching.

---

# Troubleshooting

| Message | Cause |
|:-|:-|
| "Internet connection lost! Please check your network connection and try again" | Downloading is enabled but no connection is available |
| File Chooser appears with header "winapp2.ini does not exist" | The input file was not found |
| "winapp2.ini was empty or not found" (red menu header) | The input file exists but contains no entries |
| "Error: \<path\> contains a malformatted environment variable and has been ignored" | A `DetectFile` value has a broken `%Variable%`; the entry is retained, and Trim pauses for a keypress |
| "Your key seems to be malformatted (bad root? ...)" (log only) | A `Detect` key uses a registry root other than `HKCU`/`HKLM`/`HKU`/`HKCR` |

| Symptom | Cause |
|:-|:-|
| An entry for installed software was removed | The `DetectFile`/`Detect` path doesn't match the actual installation location (e.g. a portable or non-standard install location) |
| Download option is unavailable in the menu | winapp2ool is in offline mode |

---

# Usage Examples

The outputs below were captured from real runs on a Windows 11 machine. Trim's results are machine-specific by design, your counts and retained entries will differ.

## Example 1: Anatomy of a Trim

**Context**

The clearest way to see every detection rule at work is a small file exercising each one. This input mixes four entries taken verbatim from the current winapp2.ini with two constructed for demonstration (`[My Custom Cleanup *]`, which has no detection keys, and `[Windows XP Era Application *]`, since the current database no longer ships `DetectOS` entries).

The machine in question has 7-Zip, Firefox, Steam, and Windows Calculator installed.

**Intent**

We want to trim this file in place and observe which entries survive, and why.

**Files**

###### **Input file (`winapp2.ini`)**

```ini
[7-Zip ZS *]
LangSecRef=3024
Detect=HKCU\Software\7-Zip-Zstandard
RegKey1=HKCU\Software\7-Zip-Zstandard\Compression|ArcHistory
RegKey2=HKCU\Software\7-Zip-Zstandard\Extraction|PathHistory
RegKey3=HKCU\Software\7-Zip-Zstandard\FM|CopyHistory
RegKey4=HKCU\Software\7-Zip-Zstandard\FM|FolderHistory

[Mozilla Firefox Autofill Data *]
LangSecRef=3026
DetectFile1=%AppData%\Mozilla\Firefox\Profiles
DetectFile2=%LocalAppData%\Packages\Mozilla.Firefox_*
FileKey1=%AppData%\Mozilla\Firefox\Profiles\*|autofill-profiles.json
FileKey2=%LocalAppData%\Packages\Mozilla.Firefox_*\LocalCache\Roaming\Mozilla\Firefox\Profiles\*|autofill-profiles.json

[My Custom Cleanup *]
Section=Custom Entries
FileKey1=%UserProfile%\Downloads|*.partial

[Steam Packages *]
Section=Games
Detect=HKCU\Software\Valve\Steam
FileKey1=%ProgramFiles%\Steam\package|*.zip.*

[Windows Calculator *]
LangSecRef=3025
DetectFile=%LocalAppData%\Packages\Microsoft.WindowsCalculator_*
FileKey1=%LocalAppData%\Packages\Microsoft.WindowsCalculator_*\AC|*|RECURSE
FileKey2=%LocalAppData%\Packages\Microsoft.WindowsCalculator_*\LocalCache|*|RECURSE
FileKey3=%LocalAppData%\Packages\Microsoft.WindowsCalculator_*\LocalState\Cache|*|RECURSE
FileKey4=%LocalAppData%\Packages\Microsoft.WindowsCalculator_*\Settings|*.log*

[Windows XP Era Application *]
Section=Legacy Applications
DetectOS=|5.2
FileKey1=%WinDir%\Prefetch|*.pf
```

**Command**

```
winapp2ool -trim
```

**Output**

###### **Console**

```
 ╔════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╗
 ║                                                      Trim Complete                                                 ║
 ╠════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╣
 ║ Initial entry count: 6                                                                                             ║
 ║ Trimmed entry count: 4                                                                                             ║
 ║ 2 entries trimmed from winapp2.ini (33%)                                                                           ║
 ╠════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╣
 ║                                          Press any key to return to the menu.                                      ║
 ╚════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╝
```

###### **Output file (`winapp2.ini`) after trimming **

```ini
; Firefox/Mozilla based browsers (1)

[Mozilla Firefox Autofill Data *]
LangSecRef=3026
DetectFile1=%AppData%\Mozilla\Firefox\Profiles
DetectFile2=%LocalAppData%\Packages\Mozilla.Firefox_*
FileKey1=%AppData%\Mozilla\Firefox\Profiles\*|autofill-profiles.json
FileKey2=%LocalAppData%\Packages\Mozilla.Firefox_*\LocalCache\Roaming\Mozilla\Firefox\Profiles\*|autofill-profiles.json

; End of Firefox/Mozilla based browsers

[My Custom Cleanup *]
Section=Custom Entries
FileKey1=%UserProfile%\Downloads|*.partial

[Steam Packages *]
Section=Games
Detect=HKCU\Software\Valve\Steam
FileKey1=%ProgramFiles%\Steam\package|*.zip.*

[Windows Calculator *]
LangSecRef=3025
DetectFile=%LocalAppData%\Packages\Microsoft.WindowsCalculator_*
FileKey1=%LocalAppData%\Packages\Microsoft.WindowsCalculator_*\AC|*|RECURSE
FileKey2=%LocalAppData%\Packages\Microsoft.WindowsCalculator_*\LocalCache|*|RECURSE
FileKey3=%LocalAppData%\Packages\Microsoft.WindowsCalculator_*\LocalState\Cache|*|RECURSE
FileKey4=%LocalAppData%\Packages\Microsoft.WindowsCalculator_*\Settings|*.log*
```

**Explanation**

| Entry | Result | Why |
|:-|:-|:-|
| `[7-Zip ZS *]` | Removed | `HKCU\Software\7-Zip-Zstandard` does not exist. Note that 7-Zip itself is installed, but this entry targets the Zstandard fork's exact registry key. |
| `[Mozilla Firefox Autofill Data *]` | Kept | `DetectFile1` (`%AppData%\Mozilla\Firefox\Profiles`) exists |
| `[My Custom Cleanup *]` | Kept | No detection keys,always retained |
| `[Steam Packages *]` | Kept | `HKCU\Software\Valve\Steam` exists |
| `[Windows Calculator *]` | Kept | The wildcard `DetectFile` expanded to the real package folder `Microsoft.WindowsCalculator_8wekyb3d8bbwe` |
| `[Windows XP Era Application *]` | Removed | `DetectOS=\|5.2` requires Windows ≤ 5.2 (XP/Server 2003); this machine reports 10.0 |

**Notes**

The [output formatting](#output-formatting) is visible in the result: Firefox entry was grouped into its browser-family category block consistent with the style of the CCleaner flavor.

---

## Example 2: Trimming the Full Database

**Context**

The everyday use case: producing a machine-specific copy of the complete winapp2.ini database.

**Intent**

We want to trim the full 3,715-entry winapp2.ini, saving the result to a separate file so the original is preserved.

**Command**

```
winapp2ool -trim -3f winapp2-trimmed.ini
```

**Output**

###### **Console**

```
 ╔════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╗
 ║                                                      Trim Complete                                                 ║
 ╠════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╣
 ║ Initial entry count: 3715                                                                                          ║
 ║ Trimmed entry count: 349                                                                                           ║
 ║ 3366 entries trimmed from winapp2.ini (91%)                                                                        ║
 ╠════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╣
 ║                                          Press any key to return to the menu.                                      ║
 ╚════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╝
```

###### **Output file (`winapp2-trimmed.ini`), first lines**

```ini
; Version: 251109
; # of entries: 349
;
; Winapp2.ini is fully licensed under the CC-BY-SA-4.0 license agreement. ...
```

**Explanation**

- The input file is `winapp2.ini` in the current directory (default)
- The output file is `winapp2-trimmed.ini`; the input file is left untouched
- 91% of the database was removed, only the 349 entries whose software was detected on this machine remain
- The `; Version: 251109` line was carried over from the input file, and the entry count comment was updated to the post-trim count

---

## Example 3: Downloading and Trimming in One Step

**Context**

For keeping a trimmed copy current, there is no need to maintain a local copy of the full database at all.

**Intent**

We want to download the latest winapp2.ini from GitHub and trim it in a single command, starting from an empty directory.

**Command**

```
winapp2ool -trim -d -3f winapp2-trimmed.ini
```

**Output**

###### **Console**

```
 ╔════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╗
 ║                                                      Trim Complete                                                 ║
 ╠════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╣
 ║ Initial entry count: 3715                                                                                          ║
 ║ Trimmed entry count: 349                                                                                           ║
 ║ 3366 entries trimmed from winapp2.ini (91%)                                                                        ║
 ╚════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╝
```

**Explanation**

- No local winapp2.ini is required: `-d` downloads the latest database and trims it directly
- With `-d`, the `-1f`/`-1d` input file settings are ignored
- The output is `winapp2-trimmed.ini` 

---

## Example 4: Never-Trim and Always-Trim Overrides

**Context**

Continuing from [Example 1](#example-1-anatomy-of-a-trim): suppose we plan to install 7-Zip Zstandard soon, so we want its entry to survive trimming even though it isn't detected yet. Meanwhile, we prefer Steam's own cleanup for its package cache and never want `[Steam Packages *]` in our output, even though Steam is installed.

**Intent**

We want to retain `[7-Zip ZS *]` despite its failing detection and remove `[Steam Packages *]` despite its passing detection.

**Files**

Using the same input file as Example 1, plus:

###### **Includes file (`includes.ini`)**

```ini
[7-Zip ZS *]
```

###### **Excludes file (`excludes.ini`)**

```ini
[Steam Packages *]
```

**Command**

```
winapp2ool -trim -includes -excludes -3f winapp2-trimmed.ini
```

**Output**

```ini
[Mozilla Firefox Autofill Data *]
[7-Zip ZS *]
[My Custom Cleanup *]
[Windows Calculator *]
```

**Explanation**

- `[7-Zip ZS *]` is now kept: the include list overrides its failing `Detect`
- `[Steam Packages *]` is now removed: the exclude list overrides its passing `Detect`
- `[Windows XP Era Application *]` is still removed 

**Notes**
- Both override checks run before any detection evaluation
- If an entry appears in both files, the include list wins
 - The `-includes` and `-excludes` flags must be passed on each CLI run where they're required 

---

## Example 5: VirtualStore Key Generation

**Context**

On systems upgraded through older versions of Windows, pre-UAC applications may have had writes to `Program Files` silently redirected into `%LocalAppData%\VirtualStore`. Cleaning the original location alone would miss this redirected data. This example machine has such a leftover: `%LocalAppData%\VirtualStore\Program Files (x86)\Steam\package` exists.

**Intent**

We want entries that pass detection to automatically gain coverage of their VirtualStore counterparts if and only if those locations actually exist.

**Files**

The same input file as [Example 1](#example-1-anatomy-of-a-trim); the relevant entry:

```ini
[Steam Packages *]
Section=Games
Detect=HKCU\Software\Valve\Steam
FileKey1=%ProgramFiles%\Steam\package|*.zip.*
```

**Command**

```
winapp2ool -trim -3f winapp2-trimmed.ini
```

**Output**

###### **The entry in `winapp2-trimmed.ini` after trimming**

```ini
[Steam Packages *]
Section=Games
Detect=HKCU\Software\Valve\Steam
FileKey1=%LocalAppData%\VirtualStore\Program Files*\Steam\package|*.zip.*
FileKey2=%ProgramFiles%\Steam\package|*.zip.*
```

**Explanation**

- `[Steam Packages *]` passed detection, so Trim scanned its FileKeys for VirtualStore-eligible paths
- `FileKey1=%ProgramFiles%\Steam\package|*.zip.*` produced the candidate `%LocalAppData%\VirtualStore\Program Files*\Steam\package`, which exists on this machine, so the key was added; the entry's FileKeys were then renumbered
- On a machine without that VirtualStore path, the entry would pass through unchanged. Compare to Example 1, where the same input produced no VirtualStore keys 
