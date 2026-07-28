# WinappDebug

**WinappDebug** is a winapp2.ini [linter](https://en.wikipedia.org/wiki/Lint_(software)). It performs static analysis on winapp2.ini to detect style and syntax errors across a wide range of configurable categories, and can optionally repair many of those errors automatically. 

### What does WinappDebug do?

WinappDebug reads a winapp2.ini file, validates every entry against the winapp2.ini syntax and style rules, and reports each violation to the console with the offending key. Repairs are applied as errors are found; enabling Saving writes the repaired file back to disk.

### Why WinappDebug?

- Contribution hygiene: validates entries before they're accepted to the winapp2.ini project
- Hand-edit safety: catch typos, ordering, and numbering mistakes after manual changes
- Pipeline normalization: standardize the style of files produced by Transmute or hand-merging
- Release automation: see [WinappDebug in the winapp2.ini Build](#winappdebug-in-the-winapp2ini-build)

---

# Table of Contents

1. [Requirements](#requirements)
2. [Quick Start](#quick-start)
3. [Menu Options](#menu-options)
4. [How It Works](#how-it-works)
   - [Scan, Repair, and Save](#scan-repair-and-save)
   - [Scan Settings](#scan-settings)
   - [Default Key Policy](#default-key-policy)
   - [Flavor-Aware Checks](#flavor-aware-checks)
5. [Detected Errors](#detected-errors)
6. [Command-Line Arguments](#command-line-arguments)
   - [Toggles](#toggles)
   - [File Selection](#file-selection)
   - [Flavor Selection](#flavor-selection)
   - [Examples](#examples)
7. [WinappDebug in the winapp2.ini Build](#winappdebug-in-the-winapp2ini-build)
   - [Where it runs](#where-it-runs)
   - [The reconciliation guard](#the-reconciliation-guard)
8. [Troubleshooting](#troubleshooting)
9. [Usage Examples](#usage-examples)
   - [Scanning](#scanning)
     - [Example 1: Scanning a file for errors](#example-1-scanning-a-file-for-errors)
     - [Example 2: Auditing a single category](#example-2-auditing-a-single-category)
     - [Example 3: Reporting without repairing](#example-3-reporting-without-repairing)
   - [Repairing](#repairing)
     - [Example 4: Repairing a file and saving the corrections](#example-4-repairing-a-file-and-saving-the-corrections)
     - [Example 5: Renumbering and re-alphabetization](#example-5-renumbering-and-re-alphabetization)
     - [Example 6: Structural key repairs](#example-6-structural-key-repairs)
     - [Example 7: Casing, duplicates, and unneeded numbering](#example-7-casing-duplicates-and-unneeded-numbering)
     - [Example 8: Flag and ExcludeKey repairs](#example-8-flag-and-excludekey-repairs)
     - [Example 9: Path and registry validity](#example-9-path-and-registry-validity)
   - [Errors WinappDebug will not repair](#errors-winappdebug-will-not-repair)
     - [Example 10: Errors that cannot be repaired automatically](#example-10-errors-that-cannot-be-repaired-automatically)
   - [Flavors and Default keys](#flavors-and-default-keys)
     - [Example 11: Flavor-aware linting (System Ninja)](#example-11-flavor-aware-linting-system-ninja)
     - [Example 12: Flavor-aware linting (BleachBit)](#example-12-flavor-aware-linting-bleachbit)
     - [Example 13: Auditing Default key values](#example-13-auditing-default-key-values)
     - [Example 14: Preserving Default keys with -keepdefaults](#example-14-preserving-default-keys-with--keepdefaults)
   - [Advanced](#advanced)
     - [Example 15: Merging redundant FileKeys (Optimizations)](#example-15-merging-redundant-filekeys-optimizations)
     - [Example 16: Scripted in-place repair](#example-16-scripted-in-place-repair)

---

# Requirements

- A `winapp2.ini` file to lint

---

# Quick Start

### Scan only (no file changes)

1. Place your `winapp2.ini` in the same directory as winapp2ool
2. Open WinappDebug from the main menu
3. Run. WinappDebug reports all detected errors to the console without making or saving any changes 

### Scan and repair

1. Open WinappDebug
2. Enable **Toggle Saving** to write corrections to disk
3. Run. WinappDebug corrects all auto-repairable errors and saves the result to `winapp2-debugged.ini` 

###### Note: By default the corrected file is saved to `winapp2-debugged.ini` and your input file is left untouched. Use **File Chooser (save)** (or `-3f` on the command line) to overwrite the input in place.

---

# Menu Options

| Option | Effect | Notes |
|:-|:-|:-|
| Run (default) | Lint winapp2.ini and report all detected errors | Pressing Enter with no input also runs |
| File Chooser (winapp2.ini) | Select a different local winapp2.ini to lint | Default: `winapp2.ini` in the current directory |
| Toggle Saving | Enable or disable writing corrected errors back to disk | Default: `False` |
| File Chooser (save) | Select a different output path for the corrected file | Only shown when Saving is enabled; default: `winapp2-debugged.ini` in the current directory |
| Scan Settings | Open the Scan Settings sub-menu to enable or disable individual checks | |
| Toggle Default Value Audit | Switch from removing Default keys to enforcing a specific Default value | Default: `False`; see [Default Key Policy](#default-key-policy) |
| Toggle Expected Default | Switch the enforced `Default` value between `True` and `False` | Only shown when the audit is enabled; default: `False` |
| Reset Settings | Restore all settings to their defaults | Only shown when settings have been changed |
| Log Viewer | Show the most recent lint results | Only shown when errors were found during the last run |

---

# How It Works

## Scan, Repair, and Save

WinappDebug processes each entry in winapp2.ini in turn, validating every key against the active set of lint rules and recording errors, fixing them where possible (if enabled). After processing all entries, it alphabetizes the entry list, reporting any entry that was out of position.

Each lint category operates in two independent modes:

- **Scan**: detect the error and report it to the console
- **Repair**: automatically correct the error in memory

Both are enabled by default for every category except Optimizations. Scanning without repairing reports errors but leaves the content unchanged; repairs without saving are applied in memory and reported, but no file is written. Only when **Saving** is enabled (`-c` on the command line) is the corrected file written to disk.

Each error is reported as a block naming the entry, the error, and (usually) the offending key:

```
Error in [7-Zip *]:
Forward slash (/) detected in lieu of backslash (\).
Key: FileKey1=%AppData%/7-Zip|*.tmp
```

After the run, a summary reports the entry count and the total number of possible errors. The error total is colored green (0), yellow (fewer than 10), or red (10 or more). The full report can be revisited from the menu via **Log Viewer**.

###### Note: Repairs are applied as errors are found, so a later message may name a key by the value or number it received from an earlier repair. [Example 5](#example-5-renumbering-and-re-alphabetization) and [Example 8](#example-8-flag-and-excludekey-repairs) both show this.

## Scan Settings

Each scan category can be independently enabled or disabled from the **Scan Settings** sub-menu. The menu lists all fifteen **Scan Options** first (items 1–15), then the corresponding fifteen **Repair Options** (items 16–30).

The two toggles are linked in one direction each:

- Disabling a category's **scan** also disables its repair
- Enabling a category's **repair** also re-enables its scan
- Disabling only a category's **repair** leaves the scan on, so its errors are still reported but not corrected. See [Example 3](#example-3-reporting-without-repairing)

| Category | What it detects | What the repair does |
|:-|:-|:-|
| Casing | Improper CamelCasing on commands, environment variables, registry roots, and the `RECURSE` / `REMOVESELF` flags | Rewrites to the correct casing |
| Alphabetization | Entries and keys not in alphabetical order | Reorders the entries and keys |
| Improper Numbering | Numbered keys carrying the wrong number | Renumbers the keys sequentially |
| Parameters | FileKey parameter errors (duplicate, empty, or non-alphabetical parameters) | Deletes duplicate and empty parameters, then sorts the rest |
| Flags | Incorrect flag usage in FileKeys and ExcludeKeys (`RECURSE`, `REMOVESELF`, `FILE`, `PATH`, `REG`) | Inserts the missing pipe before a flag. Flag *casing* is repaired by Casing, not here |
| Slashes | Forward slash (`/`) where a backslash (`\`) is required; consecutive backslashes; trailing backslash in `DetectFile` | Replaces `/` with `\`, collapses `\\`, and trims the trailing backslash |
| Defaults | `Default` keys present when they should not be (see [Default Key Policy](#default-key-policy)) | Deletes the `Default` key |
| Duplicates | Keys duplicating another key's value within an entry | Deletes the duplicate key |
| Unneeded Numbering | Keys carrying a number they should not have | Strips the number from the key's name |
| Multiples | Singleton keys (e.g. `LangSecRef`) appearing more than once | Deletes the extra keys |
| Invalid Values | Invalid values for `LangSecRef`, `SpecialDetect`, and similar keys | Nothing, report only  |
| Syntax Errors | Entry configurations that will not function (missing cleaning keys, missing detection, etc.) | Repairs a missing `=` and broken `%EnvironmentVariable%` delimiters |
| Path Validity | Invalid file system or registry path formats | Inserts the missing backslash before an ExcludeKey's pattern pipe; the rest is report only |
| Semicolons | Trailing semicolons; semicolons before pipe symbols in FileKeys | Strips the semicolon |
| Optimizations | FileKeys sharing both a path and a flag, which could be merged into fewer keys *(experimental; disabled by default)* | Merges them into a single key |

###### Note: Typing `alloff` at the Scan Settings menu disables every scan and repair at once.

A number of checks are not governed by any category and run no matter what the Scan Settings say. They are marked *(always on)* in [Detected Errors](#detected-errors)

Scan Settings changes made in the menu are saved and persist between sessions. Command-line runs always lint with the default scan settings: saved customizations apply only to runs started from the menu. There are two flags for controlling scans via the command line: `-opti` enables the Optimizations rule for a single run, and `-keepdefaults` switches the Defaults rule off for a single run. 

## Default Key Policy

Current winapp2.ini style ships entries without `Default=` keys. WinappDebug enforces this by default: any `Default` key is reported (`Entry has a Default key where there should be none`) and removed by the repair.

Enabling **Toggle Default Value Audit** inverts the policy: every entry is then *required* to carry a `Default` key with a specific value. The expected value is `False` unless changed via **Toggle Expected Default**. Under the audit:

- An entry with no `Default` key is reported (`No Default Key found`) and one is inserted with the expected value. This happens regardless of the Defaults scan and repair toggles
- A `Default` key with the wrong value is reported (`Incorrect value for Default Key found`) and corrected

See [Example 13](#example-13-auditing-default-key-values). 

A flag is available on the command line only: `-keepdefaults` leaves `Default=` keys exactly as it finds them, neither removing them nor requiring them. 

## Flavor-Aware Checks

A few checks depend on winapp2ool's global **Flavor** setting (found in the Winapp2ool Settings menu, or set for a single run via the [flavor command-line flags](#flavor-selection)):

| Flavor | Additional check |
|:-|:-|
| System Ninja | `Wildcard (*) found in DetectFile`: System Ninja does not support wildcards in detection keys |
| BleachBit | `ExcludeKey contains REG flag in BleachBit flavor`: BleachBit does not support registry ExcludeKeys |

---

# Detected Errors

The messages below are the exact strings printed in the error report. **Bold** messages are corrected automatically when their category's repair is enabled; plain messages are reported only and must be fixed by hand. The category in brackets is the [Scan Settings](#scan-settings) entry that governs the check — except for those marked *(always on)*, which run whether or not you have disabled anything.

### Entry structure

- `Duplicate entry name detected` *(always on)*
- `All entries must end in ' *'` *(always on)*
- `Section key found alongside LangSecRef key, but only one should be present` *(Syntax Errors)*
- `Entry has no valid classifier key (LangSecRef, Section)` *(Syntax Errors)*
- `Entry has no valid detection keys (Detect, DetectFile, DetectOS, SpecialDetect)` *(always on)*
- `Entry has no valid deletion keys (FileKey, RegKey)` *(Syntax Errors)*
- `Entry has ExcludeKeys but no valid FileKeys or RegKeys` *(Syntax Errors)*
- `Entry has ExcludeKeys pointing to file system locations but no FileKeys` *(always on)*
- `Entry has ExcludeKeys pointing to registry locations but no RegKeys` *(always on)*
- **`Entry alphabetization`** *(Alphabetization)*: the entry is out of alphabetical order; the report names the expected position

### Key formatting (all key types)

- **`Missing '=' detected and repaired in key.`** *(always on)* e.g. `FileKey1%WinDir%\tmp|*`; the report shows the key after repair
- **`{KeyName} is missing a '=' or was not provided with a value. It will be deleted.`** *(always on)*: the repair above could not produce a usable key, so the key is dropped from the entry.
- **`Detected unwanted whitespace in iniKey`** *(always on)*: leading/trailing whitespace on key names or values
- **`Forward slash (/) detected in lieu of backslash (\).`** *(Slashes)*
- **`Extraneous backslashes (\\) detected`** *(Slashes)*
- **`Trailing semicolon (;).`** *(Semicolons)*
- **`Colon (:) found where there should be a semicolon (;)`** *(always on; repaired by Semicolons)*
- **`Double '%' found in environment variable`** *(always on; repaired by Syntax Errors)*
- `Missing backslash (\) after %EnvironmentVariable%.` *(Slashes)*
- **`{Variable} has a casing error.`** *(Casing)* : e.g. `AppData has a casing error.` for `%appdata%`; also covers winapp2.ini command casing and `RECURSE`/`REMOVESELF` flag casing
- `Invalid data provided: {value} in {key}` *(Invalid Values)*: an unrecognized environment variable, command, or `SpecialDetect` value. A second line, `Valid data: {list}`, follows with the accepted values
- `Environment Variable is missing leading %` / `trailing %` / `leading and trailing %` *(Syntax Errors)*

### Numbering, ordering, and duplicates

- **`{KeyType} entry is incorrectly numbered.`** *(Improper Numbering)*: the report shows the expected and found numbers
- **`Detected unnecessary numbering.`** *(Unneeded Numbering)*: e.g. `Detect1` in an entry with a single Detect
- **`{KeyType} alphabetization`** *(Alphabetization)*: a numbered key is out of alphabetical order 
- **`Multiple {KeyType} detected.`** *(Multiples)*: a singleton key appears more than once; the duplicates are deleted
- **`Duplicate key value found`** *(Duplicates)*: the report shows both the duplicate and the key it duplicates; the duplicate is deleted

### FileKey

- `Missing pipe (|) in FileKey.` *(always on)*
- **`Duplicate FileKey parameter found`** / **`Empty FileKey parameter found`** *(Parameters)*
- **`FileKey parameters are not in alphabetical order`** *(Parameters)*: the report shows the current and sorted parameter lists
- **`Trailing semicolon (;) in parameters`** *(Semicolons)*: a semicolon immediately before the pipe
- `RECURSE or REMOVESELF is incorrectly spelled, or there are too many pipe (|) symbols.` *(always on)*
- **`Missing pipe (|) before RECURSE.`** / **`Missing pipe (|) before REMOVESELF.`** *(Flags)*
- **`Backslash (\) found before pipe (|).`** *(Slashes)*

### Detect, DetectFile, and RegKey paths

- **`Incorrect registry root casing.`** *(Casing)*: e.g. `hkcu\` for `HKCU\`
- `LangSecRef holds an invalid value.` *(Invalid Values)*: the value is not one of the recognized CCleaner section numbers
- `Invalid registry path detected.` *(Path Validity)*
- `Invalid file system path detected.` *(Path Validity)*
- `Illegal characters (< > ") detected in filesystem path.` *(Path Validity)*
- **`Trailing backslash (\) found in DetectFile`** *(Slashes)*
- `Nested wildcard found in DetectFile` *(always on)*: supported by winapp2ool's Trim, not supported by CCleaner
- `Wildcard (*) found in DetectFile`: [System Ninja flavor](#flavor-aware-checks) only

### ExcludeKey

- **`Missing pipe (|) after ExcludeKey flag`** *(Flags)*
- **`Missing backslash (\) before pipe (|) in ExcludeKey.`** *(Path Validity)*
- `No valid exclude flag (FILE, PATH, or REG) found in ExcludeKey.` *(Flags)*
- `ExcludeKey has too many flags` *(Flags)*
- `ExcludeKey contains REG flag in BleachBit flavor`: [BleachBit flavor](#flavor-aware-checks) only

### Default

- **`Entry has a Default key where there should be none`** *(Defaults)*: default policy; the key is removed
- **`No Default Key found`** *(always on under the audit)*: [Default Value Audit](#default-key-policy) only. Both the report and the insertion ignore the Defaults toggles. Enabling the audit means every entry gets a `Default` key
- **`Incorrect value for Default Key found`** *(Defaults)*: Default Value Audit only; the value is corrected

### Optimizations (experimental, disabled by default)

- FileKeys matching another key in the same entry on both path and flag are merged into a single key. See [Example 15](#example-15-merging-redundant-filekeys-optimizations)

---

# Command-Line Arguments

WinappDebug is module 1 on the command line: `winapp2ool -debug` (`-debug`, `debug`, `-1`, and `1` are equivalent).

Command-line runs always use the default scan settings, Scan Settings customizations saved from the menu do not apply. `-opti` and `-keepdefaults` are the two exceptions: each adjusts a rule for a single run without touching your saved configuration.

### Toggles

| Arg | Effect |
|:-|:-|
| `-c` | Enable saving corrected errors back to disk |
| `-usedate` | Use the current date (yymmdd) as the version string in the output file's preamble, e.g. `; Version: 260724` |
| `-opti` | Enable the experimental Optimizations rule (FileKey merger) for this run |
| `-keepdefaults` | Leave existing `Default=` keys alone instead of reporting and removing them; see [Default Key Policy](#default-key-policy) |

Two global winapp2ool flags matter for scripted lint runs. They are not WinappDebug arguments, they apply to any module and can go anywhere on the command line:

| Arg | Effect |
|:-|:-|
| `-s` | Silent mode: suppresses **all** console output. A lint run under `-s` prints nothing at all |
| `-writelog` | Force `winapp2ool.log` to be written even on a successful run. Without it, a silent run only writes the log when the exit code is nonzero |

### File Selection

| Arg | Effect | Default |
|:-|:-|:-|
| `-1d path` | Set the input winapp2.ini directory | Current directory |
| `-1f name` | Set the input winapp2.ini file name | `winapp2.ini` |
| `-1f subdir\name` | Set the input file name within a subfolder of its path | |
| `-3d path` | Set the output directory | Current directory |
| `-3f name` | Set the output file name | `winapp2-debugged.ini` |
| `-3f subdir\name` | Set the output file name within a subfolder of its path | |

###### Note: By default the output is saved to `winapp2-debugged.ini` and the input file is left untouched. To repair the input in place, pass its name to `-3f` explicitly.

### Flavor Selection

The global flavor flags select the [flavor-aware checks](#flavor-aware-checks) for a single run without changing the saved Flavor setting's menu default:

| Arg | Alias | Flavor |
|:-|:-|:-|
| `-ccleaner` | `-cc` | CCleaner (default) |
| `-bleachbit` | `-bb` | BleachBit |
| `-systemninja` | `-sn` | System Ninja |
| `-tron` | | Tron |
| `-base` | `-ncc` | Base (non-CCleaner) |
| `-ccleaner7` | `-cc7` | CCleaner 7 |
| `-fluentcleaner` | `-fc` | FluentCleaner |

### Examples

| Command | Effect |
|:-|:-|
| `winapp2ool -debug` | Scan winapp2.ini and report errors without saving |
| `winapp2ool -debug -c` | Scan and save all auto-repairable corrections to winapp2-debugged.ini |
| `winapp2ool -debug -c -3f winapp2.ini` | Scan and repair winapp2.ini in place |
| `winapp2ool -debug -c -opti -3f winapp2.ini` | Repair in place with FileKey merging enabled |
| `winapp2ool -debug -c -usedate -3f winapp2.ini` | Repair in place, stamping today's date as the version string |
| `winapp2ool -debug -c -keepdefaults -3f winapp2.ini` | Repair in place, leaving `Default=` keys untouched |
| `winapp2ool -debug -1d C:\ini\work -1f custom.ini` | Scan a specific file in another directory |
| `winapp2ool -sn -debug` | Scan winapp2.ini with the System Ninja flavor checks active |

---

# WinappDebug in the winapp2.ini Build

WinappDebug is not an optional cleanup step in this project. It runs nine times per build of winapp2.ini: three times inside the entry generators and six times as the final stage of a published flavor

## Where it runs

| Stage | Invocation | Gated? |
|:-|:-|:-|
| BrowserBuilder output | `remotedebugGuarded` (in-process, forced Optimizations) | Yes [reconciliation guard](#the-reconciliation-guard) |
| UWPBuilder output | `remotedebugGuarded` (in-process, forced Optimizations) | Yes |
| EntryBuilder output | `remotedebugGuarded` (in-process, forced Optimizations) | Yes |
| Base winapp2.ini | `-s -offline -debug -usedate -c -opti -1f Winapp2.ini -3f Winapp2.ini` | No |
| CCleaner flavor | same, on `winapp2-ccleaner-flavor.ini` | No |
| FluentCleaner flavor | same, **plus `-keepdefaults`** | No |
| BleachBit flavor | same, on `winapp2-bleachbit-flavor.ini` | No |
| Tron flavor | same, on `winapp2-tron-flavor.ini` | No |
| System Ninja flavor | same, output renamed to `Winapp2.rules` | No |

The CCleaner 7 flavor is derived from the already-linted CCleaner flavor by a Transmute pass and gets no lint pass of its own after that conversion, as the linter has not yet been extended for the CCleaner 7 format.

## The reconciliation guard

The Lint Reconciler compares the semantic output of the generators before and after their linting pass. Dropped entries, discarded malformed keys, and re-written values mean the generator created invalid data that the lint silently destroyed. This indicates a bug in the source data for the generator, and will cause the winapp2.ini build to fail.

---

# Troubleshooting

| Symptom | Cause |
|:-|:-|
| A **File Chooser** appears with the header `winapp2.ini does not exist` | The input file was not found |
| The menu header reports `winapp2.ini was empty or not found` | The input file exists but contains no entries |
| `0 possible errors detected.` on a file you expected to fail | The file is already compliant, the relevant scan categories are disabled, or the check you expected is gated behind a [flavor](#flavor-aware-checks) |
| Errors are reported but no file appears | Saving is disabled (the default)  |
| Log Viewer is not in the menu | No errors were found during the last run |

---

# Usage Examples

Some examples below reuse scenarios from the [Transmute readme](../transmute/readme.md), which corrects Browser Builder output as part of the winapp2.ini build pipeline. 

## Scanning

### Example 1: Scanning a file for errors

**Context**

We have a small winapp2.ini with a few common style mistakes: a forward slash in a path, a gap in FileKey numbering, and `Default` keys, which current winapp2.ini style omits.

**Intent**

We want to see every error WinappDebug detects without changing any file.

**Files**

###### **Input file (`winapp2.ini`)**

```ini
[7-Zip *]
LangSecRef=3021
Detect=HKCU\Software\7-Zip
FileKey1=%AppData%/7-Zip|*.tmp
Default=False

[Notepad++ *]
LangSecRef=3021
Detect=HKCU\Software\Notepad++
FileKey1=%AppData%\Notepad++|*.log
FileKey3=%AppData%\Notepad++\backup|*.*
Default=True
```

**Command**
```
winapp2ool -debug
```

**Output**

The report is rendered to the console:

```
 ╔════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╗
 ║                                                  Linting winapp2.ini                                               ║
 ╠                                                                                                                    ╣
 ║ Error in [7-Zip *]:                                                                                                ║
 ║ Forward slash (/) detected in lieu of backslash (\).                                                               ║
 ║ Key: FileKey1=%AppData%/7-Zip|*.tmp                                                                                ║
 ║                                                                                                                    ║
 ║ Error in [7-Zip *]:                                                                                                ║
 ║ Entry has a Default key where there should be none                                                                 ║
 ║                                                                                                                    ║
 ║ Error in [Notepad++ *]:                                                                                            ║
 ║ FileKey entry is incorrectly numbered.                                                                             ║
 ║ Expected: FileKey2                                                                                                 ║
 ║ Found:    FileKey3                                                                                                 ║
 ║                                                                                                                    ║
 ║ Error in [Notepad++ *]:                                                                                            ║
 ║ Entry has a Default key where there should be none                                                                 ║
 ║                                                                                                                    ║
 ╠                                                                                                                    ╣
 ║                                                     Lint Complete!                                                 ║
 ╠                                                                                                                    ╣
 ║                                                     Entry count: 2                                                 ║
 ║                                              4 possible errors detected.                                           ║
 ║                                                                                                                    ║
 ║                                          Press any key to return to the menu.                                      ║
 ╚════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╝
```

###### Note: The remaining examples show only the report lines, without the console frame.

**Explanation**
- The input file is winapp2.ini (the default)
- Both `Default` keys are flagged: `Default=False` and `Default=True` should not exist at all
- The numbering error names both the expected and the found key
- Saving is disabled, so nothing is written to disk

---

### Example 2: Auditing a single category

**Context**

We want to know if a copy of winapp2.ini's file system and registry paths are well-formed.

**Intent**

We want a report about path problems and nothing else, and we want the file left exactly as it is.

**Files**

###### **Input file (`winapp2.ini`)**

```ini
[HandBrake *]
LangSecRef=3023
Detect=HKCU\Software\HandBrake
DetectFile1=%LocalAppData%\HandBrake\
DetectFile2=%LocalAppData%\HandBrake\*\logs\*
FileKey1=%LocalAppData%\HandBrake\logs|*.txt
FileKey3=%LocalAppData%\HandBrake\cache|*.tmp
RegKey1=Software\HandBrake\Recent
```

**Steps**

`alloff` is a Scan Settings command, not a command-line flag:

1. Open **WinappDebug** from the main menu
2. Select **Toggle Saving** to enable saving
3. Select **Scan Settings**
4. Type `alloff`: every scan and every repair is now disabled
5. Select item **13** (*Path Validity*) under **Scan Options** to re-enable that scan alone
6. Return to the WinappDebug menu with `0`, then select **Run**

**Output**

```
Error in [HandBrake *]:
Nested wildcard found in DetectFile
Key: DetectFile2=%LocalAppData%\HandBrake\*\logs\*

Error in [HandBrake *]:
Invalid registry path detected.
Key: RegKey1=Software\HandBrake\Recent

2 possible errors detected.
```

###### **Output file (`winapp2-debugged.ini`)**

```ini
[HandBrake *]
LangSecRef=3023
Detect=HKCU\Software\HandBrake
DetectFile1=%LocalAppData%\HandBrake\
DetectFile2=%LocalAppData%\HandBrake\*\logs\*
FileKey1=%LocalAppData%\HandBrake\logs|*.txt
FileKey3=%LocalAppData%\HandBrake\cache|*.tmp
RegKey1=Software\HandBrake\Recent
```

**Explanation**
- `RegKey1` is the Path Validity finding
- The nested wildcard is one of the *(always on)* checks and reports regardless.
- The trailing backslash on `DetectFile1` (Slashes) and `FileKey3`'s numbering (Improper Numbering) are ignored, their categories being off
- Both reported errors are report-only checks, so nothing would have been repaired here even with repairs enabled

---

### Example 3: Reporting without repairing

**Context**

We are reviewing a contribution and want WinappDebug's full report, but we do not trust its slash handling on this particular file. We want to inspect the paths ourselves before anything rewrites them.

**Intent**

We want every category to report as usual, with the Slashes repair alone switched off.

**Files**

The input file is unchanged from [Example 1](#example-1-scanning-a-file-for-errors).

**Steps**

1. Open **WinappDebug** from the main menu
2. Select **Toggle Saving** to enable saving
3. Select **Scan Settings**
4. Select item **21** (*Slashes*) under **Repair Options** to disable that repair
5. Return with `0`, then select **Run**

**Output**

```
Error in [7-Zip *]:
Forward slash (/) detected in lieu of backslash (\).
Key: FileKey1=%AppData%/7-Zip|*.tmp

Error in [7-Zip *]:
Missing backslash (\) after %EnvironmentVariable%.
Key: FileKey1=%AppData%/7-Zip|*.tmp

Error in [7-Zip *]:
Entry has a Default key where there should be none

Error in [Notepad++ *]:
FileKey entry is incorrectly numbered.
Expected: FileKey2
Found:    FileKey3

Error in [Notepad++ *]:
Entry has a Default key where there should be none

5 possible errors detected.
```

###### **Output file (`winapp2-debugged.ini`)**

```ini
[7-Zip *]
LangSecRef=3021
Detect=HKCU\Software\7-Zip
FileKey1=%AppData%/7-Zip|*.tmp

[Notepad++ *]
LangSecRef=3021
Detect=HKCU\Software\Notepad++
FileKey1=%AppData%\Notepad++|*.log
FileKey2=%AppData%\Notepad++\backup|*.*
```

**Explanation**
- The forward slash is still reported, but not repaired.
- Every other repair ran normally: both `Default` keys were removed and `FileKey3` was renumbered to `FileKey2`
- The error count went up, from 4 to 5, compared to the same file in [Example 1](#example-1-scanning-a-file-for-errors)

**Notes**

`Missing backslash (\) after %EnvironmentVariable%.` never fires in Example 1 because the Slashes repair has already turned `%AppData%/` into `%AppData%\` by the time the check runs. 

---

## Repairing

### Example 4: Repairing a file and saving the corrections

**Context**

Continuing from [Example 1](#example-1-scanning-a-file-for-errors), we have reviewed the report and want the corrections written to disk.

**Intent**

We want WinappDebug to repair all four errors and save the corrected file.

**Files**

The input file is unchanged from Example 1.

**Command**
```
winapp2ool -debug -c
```

**Output**

The console report is identical to Example 1 with one additional summary line:

```
winapp2-debugged.ini saved with any corrections made
```

###### **Output file (`winapp2-debugged.ini`)**

```ini
[7-Zip *]
LangSecRef=3021
Detect=HKCU\Software\7-Zip
FileKey1=%AppData%\7-Zip|*.tmp

[Notepad++ *]
LangSecRef=3021
Detect=HKCU\Software\Notepad++
FileKey1=%AppData%\Notepad++|*.log
FileKey2=%AppData%\Notepad++\backup|*.*
```

**Explanation**
- The forward slash is corrected to a backslash
- `FileKey3` is renumbered to `FileKey2`
- Both `Default` keys are removed

---

### Example 5: Renumbering and re-alphabetization

**Context**

After hand-editing, an entry's FileKeys are listed out of order and the entries themselves are no longer alphabetical.

**Intent**

We want WinappDebug to renumber the keys, sort them, and sort the entry list.

**Files**

###### **Input file (`winapp2.ini`)**

```ini
[Speccy *]
LangSecRef=3025
Detect=HKCU\Software\Piriform\Speccy
FileKey2=%LocalAppData%\Programs\Speccy|*.log
FileKey1=%AppData%\Speccy|*.xml

[CCleaner *]
LangSecRef=3025
Detect=HKCU\Software\Piriform\CCleaner
FileKey1=%ProgramFiles%\CCleaner|*.log
```

**Command**
```
winapp2ool -debug -c
```

**Output**

```
Error in [Speccy *]:
FileKey entry is incorrectly numbered.
Expected: FileKey1
Found:    FileKey2

Error in [Speccy *]:
FileKey entry is incorrectly numbered.
Expected: FileKey2
Found:    FileKey1

Error in [Speccy *]:
FileKey alphabetization
FileKey1=%LocalAppData%\Programs\Speccy|*.log appears to be out of place
Expected position: 2

Error in Speccy *:
Entry alphabetization
Speccy * appears to be out of place
Expected position: 2

4 possible errors detected.
```

###### **Output file (`winapp2-debugged.ini`)**

```ini
[CCleaner *]
LangSecRef=3025
Detect=HKCU\Software\Piriform\CCleaner
FileKey1=%ProgramFiles%\CCleaner|*.log

[Speccy *]
LangSecRef=3025
Detect=HKCU\Software\Piriform\Speccy
FileKey1=%AppData%\Speccy|*.xml
FileKey2=%LocalAppData%\Programs\Speccy|*.log
```

**Explanation**
- The two FileKeys are sorted alphabetically by value and renumbered to match their new positions
- `[Speccy *]` is moved after `[CCleaner *]` in the output

**Notes**

Repairs are applied as the errors are found, so later messages may reference already-repaired keys. The alphabetization error names `FileKey1`, the key's name after renumbering.

---

### Example 6: Structural key repairs

**Context**

A hand-written entry has three broken keys: one with unwanted whitespace, one missing its `=` separator entirely, and one with a trailing semicolon.

**Intent**

We want all three keys repaired mechanically.

**Files**

###### **Input file (`winapp2.ini`)**

```ini
[Greenshot *]
LangSecRef=3021
Detect=HKCU\Software\Greenshot
FileKey1=%AppData%\Greenshot|*.log;
FileKey2%AppData%\Greenshot\thumbnails|*.png
FileKey3= %LocalAppData%\Greenshot|*.tmp
```

**Command**
```
winapp2ool -debug -c
```

**Output**

```
Error in [Greenshot *]:
Detected unwanted whitespace in iniKey
Key: FileKey3= %LocalAppData%\Greenshot|*.tmp

Error in [Greenshot *]:
Missing '=' detected and repaired in key.
Key: FileKey2=%AppData%\Greenshot\thumbnails|*.png

Error in [Greenshot *]:
Trailing semicolon (;).
Key: FileKey1=%AppData%\Greenshot|*.log;

Error in [Greenshot *]:
FileKey entry is incorrectly numbered.
Expected: FileKey2
Found:    FileKey3

Error in [Greenshot *]:
FileKey entry is incorrectly numbered.
Expected: FileKey3
Found:    FileKey2

Error in [Greenshot *]:
FileKey alphabetization
FileKey2=%LocalAppData%\Greenshot|*.tmp appears to be out of place
Expected position: 3

6 possible errors detected.
```

###### **Output file (`winapp2-debugged.ini`)**

```ini
[Greenshot *]
LangSecRef=3021
Detect=HKCU\Software\Greenshot
FileKey1=%AppData%\Greenshot|*.log
FileKey2=%AppData%\Greenshot\thumbnails|*.png
FileKey3=%LocalAppData%\Greenshot|*.tmp
```

**Explanation**
- `FileKey2` (missing `=`) is inserted 
- The whitespace and trailing semicolon are stripped
- The keys are re-sorted and re-numbered

---

### Example 7: Casing, duplicates, and unneeded numbering

**Context**

An entry has accumulated four separate problems from hand-editing: a stray second `LangSecRef`, a `Detect1` in an entry with only one Detect, a lowercase environment variable, and a FileKey that duplicates another key's value.

**Intent**

We want to see what the four data-deleting and normalizing repairs do to a single entry in one pass.

**Files**

###### **Input file (`winapp2.ini`)**

```ini
[Foobar2000 *]
LangSecRef=3023
LangSecRef=3024
Detect1=HKCU\Software\foobar2000
FileKey1=%appdata%\foobar2000|*.log
FileKey2=%AppData%\foobar2000\cache|*.tmp
FileKey3=%AppData%\foobar2000|*.log
```

**Command**
```
winapp2ool -debug -c
```

**Output**

```
Error in [Foobar2000 *]:
Multiple LangSecRef detected.
Key: LangSecRef=3024

Error in [Foobar2000 *]:
Detected unnecessary numbering.
Expected: Detect
Found:    Detect1

Error in [Foobar2000 *]:
AppData has a casing error.
Key: FileKey1=%appdata%\foobar2000|*.log

Error in [Foobar2000 *]:
Duplicate key value found
Key:            FileKey3=%AppData%\foobar2000|*.log
Duplicates:     FileKey1=%AppData%\foobar2000|*.log

4 possible errors detected.
```

###### **Output file (`winapp2-debugged.ini`)**

```ini
[Foobar2000 *]
LangSecRef=3023
Detect=HKCU\Software\foobar2000
FileKey1=%AppData%\foobar2000|*.log
FileKey2=%AppData%\foobar2000\cache|*.tmp
```

**Explanation**
- `LangSecRef=3024` is deleted: `LangSecRef` is a singleton key, and the repair keeps the first occurrence
- `Detect1` becomes `Detect`, since the entry has only one detection key of that type
- `%appdata%` becomes `%AppData%`
- `FileKey3` is deleted as a duplicate of `FileKey1`. 

---

### Example 8: Flag and ExcludeKey repairs

**Context**

A hand-written entry misuses flags in almost every way available: a `RECURSE` with no pipe before it, an ExcludeKey with no flag at all, one with the flag combined with the path, one missing the backslash before its pattern pipe, and one carrying an incorrect flag.

**Intent**

We want to see which of these the Flags and Path Validity repairs can fix and which need a human.

**Files**

###### **Input file (`winapp2.ini`)**

```ini
[Free Download Manager *]
LangSecRef=3021
Detect=HKCU\Software\FreeDownloadManager
FileKey1=%AppData%\FreeDownloadManager\logs|*.logRECURSE
ExcludeKey1=%AppData%\FreeDownloadManager\sessions|*.dat
ExcludeKey2=FILE%AppData%\FreeDownloadManager\queue|*.dat
ExcludeKey3=FILE|%AppData%\FreeDownloadManager\cache|*.tmp
ExcludeKey4=PATH|%AppData%\FreeDownloadManager\plugins\|*.*|RECURSE
```

**Command**
```
winapp2ool -debug -c
```

**Output**

```
Error in [Free Download Manager *]:
Missing pipe (|) before RECURSE.
Key: FileKey1=%AppData%\FreeDownloadManager\logs|*.logRECURSE

Error in [Free Download Manager *]:
No valid exclude flag (FILE, PATH, or REG) found in ExcludeKey.
Key: ExcludeKey1=%AppData%\FreeDownloadManager\sessions|*.dat

Error in [Free Download Manager *]:
Missing pipe (|) after ExcludeKey flag
Key: ExcludeKey2=FILE%AppData%\FreeDownloadManager\queue|*.dat

Error in [Free Download Manager *]:
Missing backslash (\) before pipe (|) in ExcludeKey.
Key: ExcludeKey2=FILE|%AppData%\FreeDownloadManager\queue|*.dat

Error in [Free Download Manager *]:
Missing backslash (\) before pipe (|) in ExcludeKey.
Key: ExcludeKey3=FILE|%AppData%\FreeDownloadManager\cache|*.tmp

Error in [Free Download Manager *]:
ExcludeKey has too many flags
Key: ExcludeKey4=PATH|%AppData%\FreeDownloadManager\plugins\|*.*|RECURSE

Error in [Free Download Manager *]:
ExcludeKey alphabetization
ExcludeKey2=FILE|%AppData%\FreeDownloadManager\queue\|*.dat appears to be out of place
Expected position: 3

7 possible errors detected.
```

###### **Output file (`winapp2-debugged.ini`)**

```ini
[Free Download Manager *]
LangSecRef=3021
Detect=HKCU\Software\FreeDownloadManager
FileKey1=%AppData%\FreeDownloadManager\logs|*.log|RECURSE
ExcludeKey1=%AppData%\FreeDownloadManager\sessions|*.dat
ExcludeKey2=FILE|%AppData%\FreeDownloadManager\cache\|*.tmp
ExcludeKey3=FILE|%AppData%\FreeDownloadManager\queue\|*.dat
ExcludeKey4=PATH|%AppData%\FreeDownloadManager\plugins\|*.*|RECURSE
```

**Explanation**
- `FileKey1` gains its missing pipe: `*.logRECURSE` becomes `*.log|RECURSE`
- `ExcludeKey1` has no flag and is not repaired
- `ExcludeKey2` is repaired twice: first the flag gets its pipe (`FILE%AppData%` -> `FILE|%AppData%`), and then is found to be missing the backslash before its pattern pipe
- `ExcludeKey4` carries both a `PATH` flag and a `RECURSE` flag; the extra flag is reported but not repaired
- The repaired ExcludeKeys re-sort, so `queue` and `cache` swap places and are renumbered


---

### Example 9: Path and registry validity

**Context**

An entry mixes valid and invalid path forms: a lowercase registry root, a `DetectFile` with a trailing backslash, a `DetectFile` with a wildcard in the middle of the path, and a `RegKey` missing its hive.

**Intent**

We want to see which path problems are mechanical and which indicate a genuinely wrong path.

**Files**

###### **Input file (`winapp2.ini`)**

```ini
[HandBrake *]
LangSecRef=3023
Detect=hkcu\Software\HandBrake
DetectFile1=%LocalAppData%\HandBrake\
DetectFile2=%LocalAppData%\HandBrake\*\logs\*
FileKey1=%LocalAppData%\HandBrake\logs|*.txt
RegKey1=Software\HandBrake\Recent
```

**Command**
```
winapp2ool -debug -c
```

**Output**

```
Error in [HandBrake *]:
Incorrect registry root casing.
Key: Detect=hkcu\Software\HandBrake

Error in [HandBrake *]:
Trailing backslash (\) found in DetectFile
Key: DetectFile1=%LocalAppData%\HandBrake\

Error in [HandBrake *]:
Nested wildcard found in DetectFile
Key: DetectFile2=%LocalAppData%\HandBrake\*\logs\*

Error in [HandBrake *]:
Invalid registry path detected.
Key: RegKey1=Software\HandBrake\Recent

4 possible errors detected.
```

###### **Output file (`winapp2-debugged.ini`)**

```ini
[HandBrake *]
LangSecRef=3023
Detect=HKCU\Software\HandBrake
DetectFile1=%LocalAppData%\HandBrake
DetectFile2=%LocalAppData%\HandBrake\*\logs\*
FileKey1=%LocalAppData%\HandBrake\logs|*.txt
RegKey1=Software\HandBrake\Recent
```

**Explanation**
- `hkcu\` becomes `HKCU\`
- `DetectFile1`'s trailing backslash is trimmed
- `DetectFile2`'s nested wildcard is reported but not repaired
- `RegKey1` has no registry hive; it is reported and not repaired

---

## Errors WinappDebug will not repair

### Example 10: Errors that cannot be repaired automatically

**Context**

An entry carries an invalid `LangSecRef` value and has no detection keys at all. WinappDebug cannot know the correct category or the correct detection path, so there is nothing it can safely repair.

**Intent**

We want to see how report-only errors behave when saving is enabled.

**Files**

###### **Input file (`winapp2.ini`)**

```ini
[Winamp *]
LangSecRef=9999
FileKey1=%AppData%\Winamp|*.log
```

**Command**
```
winapp2ool -debug -c
```

**Output**

```
Error in [Winamp *]:
LangSecRef holds an invalid value.
Key: LangSecRef=9999

Error in [Winamp *]:
Entry has no valid detection keys (Detect, DetectFile, DetectOS, SpecialDetect)

2 possible errors detected.
```

###### **Output file (`winapp2-debugged.ini`)**

```ini
[Winamp *]
LangSecRef=9999
FileKey1=%AppData%\Winamp|*.log
```

**Explanation**
- Both errors are reported, but the entry unchanged
- Report-only errors always require manual correction

---

## Flavors and Default keys

### Example 11: Flavor-aware linting (System Ninja)

**Context**

The System Ninja flavor of winapp2.ini cannot use wildcards in `DetectFile` keys, System Ninja does not support them. The same key is perfectly valid in the base and CCleaner flavors. 

**Intent**

We want to validate a file against System Ninja's rules for a single run.

**Files**

###### **Input file (`winapp2.ini`)**

```ini
[Brave Session *]
Section=Brave Web Browser
DetectFile=%LocalAppData%\BraveSoftware\Brave-Browser*
FileKey1=%LocalAppData%\BraveSoftware\Brave-Browser*\User Data\*\Sessions|*|REMOVESELF
```

**Command**
```
winapp2ool -sn -debug
```

**Output**

```
Error in [Brave Session *]:
Wildcard (*) found in DetectFile
Key: DetectFile=%LocalAppData%\BraveSoftware\Brave-Browser*

1 possible errors detected.
```

**Explanation**
- The `-sn` flag sets the flavor to System Ninja for this run, activating its DetectFile wildcard check
- The same command without `-sn` reports `0 possible errors detected.`
- The check is report-only and cannot be repaired automatically 

---

### Example 12: Flavor-aware linting (BleachBit)

**Context**

BleachBit does not implement registry exclusions. An `ExcludeKey` carrying the `REG` flag should report an error 

**Intent**

We want to find registry ExcludeKeys before shipping a file to BleachBit users.

**Files**

###### **Input file (`winapp2.ini`)**

```ini
[WinRAR *]
LangSecRef=3021
Detect=HKCU\Software\WinRAR
FileKey1=%AppData%\WinRAR|*.log
RegKey1=HKCU\Software\WinRAR\DialogEditHistory
ExcludeKey1=REG|HKCU\Software\WinRAR\DialogEditHistory\ArcName
```

**Command**
```
winapp2ool -bb -debug
```

**Output**

```
Error in [WinRAR *]:
ExcludeKey contains REG flag in BleachBit flavor
Key: ExcludeKey1=REG|HKCU\Software\WinRAR\DialogEditHistory\ArcName

1 possible errors detected.
```

**Explanation**
- The `-bb` flag activates the BleachBit check for this run only; your saved Flavor setting is untouched
- The identical run without `-bb` reports `0 possible errors detected.`
- Like the System Ninja check, this is report-only. 

---

### Example 13: Auditing Default key values

**Context**

We are preparing a file for a consumer that expects every entry to carry `Default=False` 

**Intent**

We want every entry to end up with `Default=False`: inserted where missing, corrected where wrong.

**Files**

###### **Input file (`winapp2.ini`)**

```ini
[Audacity *]
LangSecRef=3023
Detect=HKCU\Software\Audacity
FileKey1=%AppData%\audacity|*.log

[VLC Media Player *]
LangSecRef=3023
Detect=HKCU\Software\VideoLAN\VLC
FileKey1=%AppData%\vlc|*.log
Default=True
```

**Steps**

The Default Value Audit has no command-line flag; it is enabled from the menu:

1. Open **WinappDebug** from the main menu
2. Select **Toggle Saving** to enable saving
3. Select **Toggle Default Value Audit** to enable the audit (the expected value defaults to `False`)
4. Select **Run**

**Output**

```
Error in [Audacity *]:
No Default Key found

Error in [VLC Media Player *]:
Incorrect value for Default Key found
Key: Default=True

2 possible errors detected.
```

###### **Output file (`winapp2-debugged.ini`)**

```ini
[Audacity *]
LangSecRef=3023
Detect=HKCU\Software\Audacity
Default=False
FileKey1=%AppData%\audacity|*.log

[VLC Media Player *]
LangSecRef=3023
Detect=HKCU\Software\VideoLAN\VLC
Default=False
FileKey1=%AppData%\vlc|*.log
```

**Explanation**
- `[Audacity *]` had no Default key, so `Default=False` is inserted
- `[VLC Media Player *]`'s `Default=True` is corrected to `Default=False`
- **Toggle Expected Default** flips the enforced value to `True` for consumers that want everything enabled
- Without the audit, the same run would instead *remove* `Default=True` and flag nothing on `[Audacity *]`. See [Example 1](#example-1-scanning-a-file-for-errors)

---

### Example 14: Preserving Default keys with -keepdefaults

**Context**

The FluentCleaner flavor ships a default-*on* cleaner, so its build marks browser-privacy entries `Default=False` while leaving cache and telemetry entries enabled. Those `Default=` keys are a large point of the flavor. Linting that output with the default Defaults policy would delete every one of them, silently, and the file would still look fine.

**Intent**

We want to lint the FluentCleaner output for real style errors without touching its `Default=` keys.

**Files**

###### **Input file (`winapp2.ini`)**

```ini
[Vivaldi Web Browsing Cookies *]
LangSecRef=3033
DetectFile=%LocalAppData%\Vivaldi\User Data
Default=False
FileKey1=%LocalAppData%\Vivaldi\User Data\*\Network|Cookies;Cookies-journal

[Vivaldi Web Browsing Internet Cache *]
LangSecRef=3033
DetectFile=%LocalAppData%\Vivaldi\User Data
FileKey1=%LocalAppData%\Vivaldi\User Data\*\Cache|*|RECURSE
```

**Commands**

The wrong way, then the right way:

```
winapp2ool -debug -c -3f no-flag.ini
winapp2ool -debug -c -keepdefaults -3f with-flag.ini
```

**Output**

The first command reports the `Default` key as an error:

```
Error in [Vivaldi Web Browsing Cookies *]:
Entry has a Default key where there should be none

1 possible errors detected.
```

The second reports nothing at all:

```
0 possible errors detected.
```

###### **Output file (`no-flag.ini`) without the flag**

```ini
;
; Vivaldi (2)

[Vivaldi Web Browsing Cookies *]
LangSecRef=3033
DetectFile=%LocalAppData%\Vivaldi\User Data
FileKey1=%LocalAppData%\Vivaldi\User Data\*\Network|Cookies;Cookies-journal

[Vivaldi Web Browsing Internet Cache *]
LangSecRef=3033
DetectFile=%LocalAppData%\Vivaldi\User Data
FileKey1=%LocalAppData%\Vivaldi\User Data\*\Cache|*|RECURSE

; End of Vivaldi
```

###### **Output file (`with-flag.ini`) with `-keepdefaults`**

```ini
;
; Vivaldi (2)

[Vivaldi Web Browsing Cookies *]
LangSecRef=3033
DetectFile=%LocalAppData%\Vivaldi\User Data
Default=False
FileKey1=%LocalAppData%\Vivaldi\User Data\*\Network|Cookies;Cookies-journal

[Vivaldi Web Browsing Internet Cache *]
LangSecRef=3033
DetectFile=%LocalAppData%\Vivaldi\User Data
FileKey1=%LocalAppData%\Vivaldi\User Data\*\Cache|*|RECURSE

; End of Vivaldi
```

**Explanation**
- Without the flag, `Default=False` is reported as an error and deleted. 
- With `-keepdefaults`, the key is neither reported nor touched, and the run reports zero errors

---

## Advanced

### Example 15: Merging redundant FileKeys (Optimizations)

**Context**

[Transmute's Example 9](../transmute/readme.md) adds a `FileKey` to a generated entry, leaving it unnumbered and pointing at a path another FileKey already covers. The Optimizations category can fold such keys together.

**Intent**

We want the redundant FileKey merged into the existing key covering the same path, and the entry's syntax normalized.

**Files**

###### **Input file (`winapp2.ini`)**

```ini
[360 Secure Browser Bookmarked Websites *]
Section=.360 Secure Browser Web Browser
DetectFile=%AppData%\360se6\User Data
FileKey1=%AppData%\360se6\User Data\*|bookmarks;BookmarkMergedSurfaceOrdering
FileKey2=%AppData%\360se6\User Data\*\power_bookmarks|*|REMOVESELF
FileKey=%AppData%\360se6\User Data\*|360Bookmarks*
```

**Command**
```
winapp2ool -debug -c -opti
```

###### Note: Optimizations can also be enabled for menu runs from **Scan Settings** 

**Output**

```
Error in [360 Secure Browser Bookmarked Websites *]:
FileKey parameters are not in alphabetical order
Key:      FileKey1=%AppData%\360se6\User Data\*|bookmarks;BookmarkMergedSurfaceOrdering
Current:  bookmarks;BookmarkMergedSurfaceOrdering
Sorted:   BookmarkMergedSurfaceOrdering;bookmarks

Error in [360 Secure Browser Bookmarked Websites *]:
FileKey entry is incorrectly numbered.
Expected: FileKey3
Found:    FileKey

Error in [360 Secure Browser Bookmarked Websites *]:
FileKey alphabetization
FileKey3=%AppData%\360se6\User Data\*|360Bookmarks* appears to be out of place
Expected position: 1

[360 Secure Browser Bookmarked Websites *] has keys which can be merged
The following keys can be merged into other keys:
FileKey2=%AppData%\360se6\User Data\*|BookmarkMergedSurfaceOrdering;bookmarks
The resulting key list will be reduced to:
FileKey1=%AppData%\360se6\User Data\*|360Bookmarks*;BookmarkMergedSurfaceOrdering;bookmarks
FileKey2=%AppData%\360se6\User Data\*\power_bookmarks|*|REMOVESELF

3 possible errors detected.
```

###### **Output file (`winapp2-debugged.ini`)**

```ini
[360 Secure Browser Bookmarked Websites *]
Section=.360 Secure Browser Web Browser
DetectFile=%AppData%\360se6\User Data
FileKey1=%AppData%\360se6\User Data\*|360Bookmarks*;BookmarkMergedSurfaceOrdering;bookmarks
FileKey2=%AppData%\360se6\User Data\*\power_bookmarks|*|REMOVESELF
```

**Explanation**
- The unnumbered `FileKey` targets `%AppData%\360se6\User Data\*` with no flag, matching `FileKey1`, so its `360Bookmarks*` parameter is merged into `FileKey1` and the redundant key is deleted
- The merge is announced in its own report block naming the keys removed and the resulting key list, but is not counted in the error total; the numbering and ordering errors it triggers appear as usual
- The merged parameter list is a straight concatenation in key order (`360Bookmarks*;BookmarkMergedSurfaceOrdering;bookmarks`). 
- This is the same normalization [Transmute's Example 9](../transmute/readme.md) applies after adding a key to a generated entry

---

### Example 16: Scripted in-place repair

**Context**

The winapp2.ini build pipeline lints each flavor as its final stage. Nothing is interactive and nothing lands in a second file: the artifact is repaired where it sits and stamped with the build date.

**Intent**

We want a single non-interactive command that repairs a file in place and records when it was built.

**Files**

###### **Input file (`winapp2.ini`)**

```ini
[Audacity *]
LangSecRef=3023
Detect=HKCU\Software\Audacity
FileKey1=%AppData%\audacity|*.log
FileKey2=%AppData%\audacity\cache|*.tmp
```

**Command**

This is the build's own base-file lint stage, verbatim from `build winapp2.ps1`, with `-writelog` added so the run leaves evidence behind:

```
winapp2ool -s -offline -writelog -debug -usedate -c -opti -1f Winapp2.ini -3f Winapp2.ini
```

**Output**

Nothing on the console: `-s` suppresses all of it. The report goes to `winapp2ool.log`, whose tail reads:

```
Lint complete
Entry count: 1
0 errors detected
2 ms
```

###### **Output file (`winapp2.ini`) first three lines**

```ini
; Version: 260724
; # of entries: 1
;
```

**Explanation**
- `-3f Winapp2.ini` points the save target at the input, so the file is repaired in place. Without it the output would be saved to `winapp2-debugged.ini` and the input would be untouched
- `-usedate` stamps the preamble's version line with today's date in `yymmdd` form. 
- `-offline` skips the startup connection check
- `-s` makes the run non-interactive. It also silences the report entirely, which is why `-writelog` is needed
- `-opti` enables the FileKey merger, as every stage of the real build does
- The run reports zero errors and still rewrites the file, the preamble comment is regenerated on every save regardless of whether any entry changed
