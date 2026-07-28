# Trim

**Trim** is a winapp2ool module that strips winapp2.ini down to the entries that apply to the machine it runs on. Every entry carries detection criteria (a registry key, a file path, a Windows version requirement) saying whether the software it targets is actually installed. Trim checks those criteria against the current system and drops the entries whose software isn't there.

### What does Trim do?

Trim reads a winapp2.ini (off disk, or downloaded from GitHub), tests every entry against the machine it is running on, and saves a copy with only the entries that passed. 

### Why Trim?

- Performance: CCleaner evaluates every entry in winapp2.ini when it starts, and it is slow about it. Trimming typically removes around 90% of the file, so CCleaner starts faster
- Automation: `winapp2ool -trim -d` downloads the latest database and gives you a machine-specific copy in one command

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

Each entry is checked on its own, in this order, stopping at the first step that gives an answer.

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

A failed version check removes the entry on the spot; Trim never looks at its other detection keys. If the check passes and there are no other detection keys, the entry is kept.

## Detect

`Detect` names a registry path that Trim checks for existence.

- Paths under `HKLM\Software` are also checked under `HKLM\SOFTWARE\WOW6432Node`, so 32-bit applications on 64-bit Windows are found 
- If a key can't be read because of permissions, Trim assumes it exists and keeps the entry

## DetectFile

`DetectFile` names a path that Trim checks for existence. It can point at either a file or a folder.

- Wildcards (`*`) work anywhere in the path. Trim expands each wildcard segment against the folders actually on disk, and keeps the entry if any expansion resolves to something that exists. 
- Paths under `%ProgramFiles%` are checked in both the native Program Files directory and the 32-bit `Program Files (x86)` directory. See [Environment Variables](#environment-variables)
- If a folder can't be read because of permissions, Trim assumes the target exists and keeps the entry

## SpecialDetect

`SpecialDetect` is a deprecated CCleaner variable. Trim still handles it so that very old winapp2.ini files work. Four values are recognized:

| Value | What is checked |
|:-|:-|
| `DET_CHROME` | A hardcoded list of ~26 Chromium-family browser paths and registry keys |
| `DET_MOZILLA` | `%AppData%\Mozilla\Firefox` |
| `DET_THUNDERBIRD` | `%AppData%\Thunderbird` |
| `DET_OPERA` | `%AppData%\Opera Software` |

If any of the associated paths or keys exists on the system, the entry is kept.

## Entries Without Detection Keys

Entries with no detection keys of any type are always kept

## Includes and Excludes

Two optional override files let you overrule detection. Trim consults both before it looks at any detection key.

**Include list** (`includes.ini` by default): Name an entry here and it is always kept, whether or not its detection criteria are satisfied. Turn it on with **Toggle include list** in the menu, or `-includes` on the command line.

**Exclude list** (`excludes.ini` by default): Entries named here are always removed, even when their detection criteria pass. **Toggle exclude list** in the menu, `-excludes` on the command line.

Each file is a plain ini file where section names are the entry names to match:

```ini
[Some Application *]
[Another Application *]
```

Whatever keys those sections contain is ignored, so you can leave them empty. 

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

On 64-bit systems, `%ProgramFiles%` covers both the native Program Files directory and `Program Files (x86)`. `HKLM\Software` registry paths get the same treatment through `WOW6432Node`.

**Malformed variables:** a path whose environment variable is malformed (e.g. a value ending at the closing `%` with no path after it) prints an error and pauses:

```
Error: <path> contains a malformatted environment variable and has been ignored
The associated entry will be retained in the final output file
Press any key to continue
```

## VirtualStore Handling

Pre-UAC applications that wrote to protected locations had those writes quietly redirected into the user's VirtualStore, and machines upgraded from older Windows versions still carry the leftovers. So for every entry that survives detection, Trim reads the `FileKey`, `RegKey`, and `ExcludeKey` values and adds keys for the matching VirtualStore locations, but only where those locations exist:

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

Command-line runs always start from Trim's defaults. Settings you saved from the menu, like alternate file paths or enabled include/exclude lists, do not carry over, so pass the flags you need on every invocation.

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

A trimmed file only describes the winapp2.ini it came from and the software that was installed when it ran. Update winapp2.ini or install a new application and it is out of date, and an out of date trimmed file has no entries covering whatever you just installed. Re-trim after either kind of change.

### Download and Trim 

You can skip keeping a local copy of the database entirely. **Toggle downloading** in the menu, or `-d` on the command line, fetches the current winapp2.ini and trims it in one operation. 

### Trimmed Files and CCleaner Performance

CCleaner evaluates every entry in winapp2.ini when it starts, so a big file means a slow launch. If you are on CCleaner 7, CC7Patcher can trim the file for you before patching it in.

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

The outputs below were captured from real runs on a Windows 11 machine. Trim's results are machine-specific by design, so expect different counts and a different set of survivors on yours.

## Example 1: Anatomy of a Trim

**Context**

Six entries are enough to watch every detection rule fire at once. Four of them are lifted verbatim from the current winapp2.ini; the other two exist for the demonstration. `[My Custom Cleanup *]` has no detection keys at all, and `[Windows XP Era Application *]` is here because the real database no longer ships a single `DetectOS` entry.

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
| `[My Custom Cleanup *]` | Kept | No detection keys, always retained |
| `[Steam Packages *]` | Kept | `HKCU\Software\Valve\Steam` exists |
| `[Windows Calculator *]` | Kept | The wildcard `DetectFile` expanded to the real package folder `Microsoft.WindowsCalculator_8wekyb3d8bbwe` |
| `[Windows XP Era Application *]` | Removed | `DetectOS=\|5.2` requires Windows ≤ 5.2 (XP/Server 2003); this machine reports 10.0 |

**Notes**

You can see the [output formatting](#output-formatting) at work in the result: the Firefox entry came out wrapped in its browser-family comment block, the way the CCleaner flavor is formatted.

---

## Example 2: Trimming the Full Database

**Context**

This is the ordinary case, running Trim over the whole database to get a copy that fits one machine.

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
- 91% of the database was removed; the 349 survivors are the entries whose software was actually detected on this machine
- The `; Version: 251109` line was carried over from the input file, and the entry count comment was updated to the post-trim count

---

## Example 3: Downloading and Trimming in One Step

**Context**

If the trimmed file is all you keep, you never need the full database sitting on disk.

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

On systems upgraded through older versions of Windows, pre-UAC applications may have had writes to `Program Files` silently redirected into `%LocalAppData%\VirtualStore`. Clean only the original location and the redirected copy is missed. This machine has exactly that leftover: `%LocalAppData%\VirtualStore\Program Files (x86)\Steam\package` exists.

**Intent**

We want surviving entries to pick up coverage of their VirtualStore counterparts, but only where those locations really exist.

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
