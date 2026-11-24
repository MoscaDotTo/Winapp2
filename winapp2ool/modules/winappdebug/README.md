# WinappDebug

**WinappDebug** is a [linter](https://en.wikipedia.org/wiki/Lint_(software)) for winapp2.ini. It checks winapp2.ini entries for style and syntax errors across a wide range of configurable categories, and repairs many of those errors automatically.

---

# Table of Contents

#### **Using WinappDebug**

1. [Requirements](#requirements)
2. [Quick Start](#quick-start)
3. [Common Uses](#common-uses)
4. [Main Menu Options](#menu-options)
5. [Command-Line Arguments](#command-line-arguments) 
    - [Toggles](#toggles)
    - [File Selection](#file-selection)
    - [Flavor Selection](#flavor-selection)
6. [Configuration](#configuration)
    - [Scan Settings](#scan-settings)
    - [Default Key Policy](#default-key-policy)
    - [Flavor-Aware Checks](#flavor-aware-checks)
7. [Detected Errors](#detected-errors)
8. [Troubleshooting](#troubleshooting)
9. **[Usage Examples](EXAMPLES.md)**

#### **For Contributors**

9. [WinappDebug in the winapp2.ini Build](#winappdebug-in-the-winapp2ini-build)
    - [The reconciliation guard](#the-reconciliation-guard)

---

# Requirements

- A `winapp2.ini` file to lint

---

# Quick Start

### Scan only (no file changes)

1. Put your `winapp2.ini` in the same folder as winapp2ool
2. Open WinappDebug from the main menu
3. Select **Run**. WinappDebug lists every error it finds and doesn't change any file

### Scan and repair

1. Open WinappDebug
2. Select **Toggle Saving** to turn saving on
3. Select **Run**. WinappDebug repairs every error it can and saves the result to `winapp2-debugged.ini`

###### Note: WinappDebug saves to `winapp2-debugged.ini` and leaves your input file alone unless you tell it otherwise. To overwrite the input instead, point **File Chooser (save)** (or `-3f` on the command-line) at it.

---

# Common Uses

| To | Run | Example |
|:-|:-|:-|
| Scan `winapp2.ini` and report errors without saving | `winapp2ool -debug` | [Scanning a file for errors](EXAMPLES.md#example-1-scanning-a-file-for-errors) |
| Repair errors and save the result to `winapp2-debugged.ini` | `winapp2ool -debug -c` | [Repairing a file and saving the corrections](EXAMPLES.md#example-4-repairing-a-file-and-saving-the-corrections) |
| Repair `winapp2.ini` in place | `winapp2ool -debug -c -3f winapp2.ini` | |
| Repair with FileKey merging turned on | `winapp2ool -debug -c -opti` | [Merging redundant FileKeys (Optimizations)](EXAMPLES.md#example-15-merging-redundant-filekeys-optimizations) |
| Repair in place, leaving `Default=` keys alone | `winapp2ool -debug -c -keepdefaults -3f winapp2.ini` | [Preserving Default keys with -keepdefaults](EXAMPLES.md#example-14-preserving-default-keys-with--keepdefaults) |
| Scan with the System Ninja flavor checks on | `winapp2ool -sn -debug` | [Flavor-aware linting (System Ninja)](EXAMPLES.md#example-11-flavor-aware-linting-system-ninja) |
| Repair in place from a script with no console output, writing `winapp2ool.log` | `winapp2ool -s -offline -writelog -debug -usedate -c -opti -1f Winapp2.ini -3f Winapp2.ini` | [Scripted in-place repair](EXAMPLES.md#example-16-scripted-in-place-repair) |

In the menu, you make the same choices with **Toggle Saving** and **File Chooser (save)** before you select **Run**, and **Scan Settings** turns individual checks on and off. See [Menu Options](#menu-options). `-s`, `-offline` and `-writelog` are [winapp2ool global args](../../Readme.md#global-args).

---


# Main Menu Options

| Option | Effect | Notes |
|:-|:-|:-|
| Run (default) | Lint winapp2.ini and report every error found | Pressing Enter on an empty line also runs it |
| File Chooser (winapp2.ini) | Choose a different winapp2.ini to lint | Default: `winapp2.ini` in the current directory |
| Toggle Saving | Turn saving the corrected file on or off | Default: `Off` |
| File Chooser (save) | Choose where the corrected file is saved | Only shown while saving is on; default: `winapp2-debugged.ini` in the current directory |
| Scan Settings | Open the Scan Settings menu to turn individual checks on or off | |
| Toggle Default Value Audit | Require a `Default` key with a set value in every entry, instead of removing `Default` keys | Default: `Off`; see [Default Key Policy](#default-key-policy) |
| Toggle Expected Default | Switch the required `Default` value between `True` and `False` | Only shown while the audit is on; default: `False` |
| Reset Settings | Restore all settings to their defaults | Only shown when settings have been changed |
| Log Viewer | Show the most recent lint results | Only shown after a run that found errors |

---

# Command-Line Arguments

From the command-line, WinappDebug is invoked with `winapp2ool -debug` (`debug`, `-1` and `1` also work).

Command-line runs always use the default scan settings, most of which are on by default. `-opti` and `-keepdefaults` each enable an off-by-default rule for a single run.

### Toggles

| Arg | Effect |
|:-|:-|
| `-c` | Save the corrected file |
| `-usedate` | Use today's date (yymmdd) as the version line in the saved file's header, e.g. `; Version: 260724` |
| `-opti` | Turn on FileKey merging for this run |
| `-keepdefaults` | Leave existing `Default=` keys alone instead of reporting and removing them; see [Default Key Policy](#default-key-policy) |
| `-archive` | Save corrections with the `archive.ini` header |
| `-winapp3` | Save corrections with the `winapp3.ini` header |

### Winapp2 Global Args 
| Arg | Effect |
|:-|:-|
| `-s` | Silent mode. Suppresses **all** console output, so a lint run under `-s` prints nothing at all |
| `-writelog` | Write `winapp2ool.log` on exit to see lint results from scripted runs |

### File Selection

| Arg | Effect | Default |
|:-|:-|:-|
| `-1d path` | Set the input winapp2.ini directory | Current directory |
| `-1f name` | Set the input winapp2.ini file name | `winapp2.ini` |
| `-1f subdir\name` | Use a file in a subfolder of the input directory | |
| `-3d path` | Set the output directory | Current directory |
| `-3f name` | Set the output file name | `winapp2-debugged.ini` |
| `-3f subdir\name` | Save to a subfolder of the output directory | |

### Flavor Selection

The global flavor flags turn on a flavor's [flavor-aware checks](#flavor-aware-checks) for one run.

| Arg | Alias | Flavor |
|:-|:-|:-|
| `-ccleaner` | `-cc` | CCleaner |
| `-bleachbit` | `-bb` | BleachBit |
| `-systemninja` | `-sn` | System Ninja |
| `-tron` | | Tron |
| `-base` | `-ncc` | Base (default) |
| `-ccleaner7` | `-cc7` | CCleaner 7 |
| `-fluentcleaner` | `-fc` | FluentCleaner |

---

# Configuration

## Scan Settings

The **Scan Settings** menu turns each category on or off. It lists all fifteen **Scan Options** first (items 1–15), then their corresponding **Repair Options** (items 16–30).

Turning a category's scan off turns its repair off too and turning a category's repair on turns its scan on.

| Category | What it detects | What the repair does |
|:-|:-|:-|
| Casing | Wrong CamelCasing on commands, environment variables, registry roots, and the `RECURSE` / `REMOVESELF` flags | Rewrites them with the right casing |
| Alphabetization | Entries and keys out of alphabetical order | Reorders the entries and keys |
| Improper Numbering | Numbered keys with the wrong number | Renumbers the keys in order |
| Parameters | FileKey parameter errors (duplicate, empty, or out-of-order parameters) | Deletes duplicate and empty parameters, then sorts the rest |
| Flags | Wrong flag usage in FileKeys and ExcludeKeys (`RECURSE`, `REMOVESELF`, `FILE`, `PATH`, `REG`) | Inserts the missing pipe before a flag. Flag *casing* belongs to Casing, not this category |
| Slashes | A forward slash (`/`) where a backslash (`\`) belongs; doubled backslashes; a trailing backslash in `DetectFile` | Replaces `/` with `\`, collapses `\\`, and trims the trailing backslash |
| Defaults | `Default` keys where there shouldn't be any (see [Default Key Policy](#default-key-policy)) | Deletes the `Default` key |
| Duplicates | Keys with the same value as another key in the entry | Deletes the duplicate key |
| Unneeded Numbering | Keys with a number they shouldn't have | Strips the number from the key's name |
| Multiples | Keys that should appear once (e.g. `LangSecRef`) appearing more than once | Deletes the extra keys |
| Invalid Values | Invalid values for `LangSecRef`, `SpecialDetect`, and similar keys | Nothing, report only |
| Syntax Errors | Entries that can't work as written (no cleaning keys, no detection, etc.) | Repairs a missing `=` and broken `%EnvironmentVariable%` delimiters |
| Path Validity | Badly formed file system or registry paths | Inserts the missing backslash before an ExcludeKey's pattern pipe; everything else is report only |
| Semicolons | Trailing semicolons; semicolons just before the pipe in FileKeys | Strips the semicolon |
| Optimizations | FileKeys with the same path and flag, which could be merged into fewer keys *(experimental; off by default)* | Merges them into one key |

###### Note: Type `alloff` at the Scan Settings menu to turn every scan and repair off at once.

## Default Key Policy

By default, WinappDebug reports `Default=` keys as errors and removes them. Enabling the **Default Value Audit** instead enforces that every entry *should* have a `Default=` key. If it is absent, one is inserted according to the **Expected Default** value (`False` by default). If one is present but has an incorrect value, its value is corrected. See [Example 13](EXAMPLES.md#example-13-auditing-default-key-values).

`-keepdefaults` only exists on the command-line. It leaves `Default=` keys exactly as it finds them, neither removing them nor requiring them.

## Flavor-Aware Checks

There are a few checks for specific flavors of winapp2.ini. These checks only run when winapp2ool's flavor setting is set accordingly. 

| Flavor | Additional check | 
|:-|:-|
| System Ninja | `Wildcard (*) found in DetectFile`: System Ninja doesn't support wildcards in detection keys |
| BleachBit | `ExcludeKey contains REG flag in BleachBit flavor`: BleachBit 5 and older don't support registry ExcludeKeys |

---

# Detected Errors

These are the exact messages the report prints. Messages in **bold** can be automatically repaired. The category in parentheses is the [Scan Settings](#scan-settings) entry that controls the check, unless it's marked *(always on)*.

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
- **`Entry alphabetization`** *(Alphabetization)*: the entry is out of alphabetical order. The report gives the position it belongs in

### Key formatting (all key types)

- **`Missing '=' detected and repaired in key.`** *(always on)*: e.g. `FileKey1%WinDir%\tmp|*`. The report shows the key after the repair
- **`{KeyName} is missing a '=' or was not provided with a value. It will be deleted.`** *(always on)*: the repair above couldn't produce a usable key, so the key is dropped from the entry
- **`Detected unwanted whitespace in iniKey`** *(always on)*: leading or trailing whitespace on a key's name or value
- **`Forward slash (/) detected in lieu of backslash (\).`** *(Slashes)*
- **`Extraneous backslashes (\\) detected`** *(Slashes)*
- **`Trailing semicolon (;).`** *(Semicolons)*
- **`Colon (:) found where there should be a semicolon (;)`** *(always on; repaired by Semicolons)*
- **`Double '%' found in environment variable`** *(always on; repaired by Syntax Errors)*
- `Missing backslash (\) after %EnvironmentVariable%.` *(Slashes)*
- **`{Variable} has a casing error.`** *(Casing)*: e.g. `AppData has a casing error.` for `%appdata%`. This also covers the casing of winapp2.ini commands and the `RECURSE`/`REMOVESELF` flags
- `Invalid data provided: {value} in {key}` *(Invalid Values)*: an environment variable, command, or `SpecialDetect` value WinappDebug doesn't recognize. A second line, `Valid data: {list}`, lists the values it accepts
- `Environment Variable is missing leading %` / `trailing %` / `leading and trailing %` *(Syntax Errors)*

### Numbering, ordering, and duplicates

- **`{KeyType} entry is incorrectly numbered.`** *(Improper Numbering)*: the report shows the expected and found numbers
- **`Detected unnecessary numbering.`** *(Unneeded Numbering)*: e.g. `Detect1` in an entry with only one Detect
- **`{KeyType} alphabetization`** *(Alphabetization)*: a numbered key is out of alphabetical order
- **`Multiple {KeyType} detected.`** *(Multiples)*: a key that should appear once appears more than once. The extras are deleted
- **`Duplicate key value found`** *(Duplicates)*: the report shows the duplicate and the key it duplicates. The duplicate is deleted

### FileKey

- `Missing pipe (|) in FileKey.` *(always on)*
- **`Duplicate FileKey parameter found`** / **`Empty FileKey parameter found`** *(Parameters)*
- **`FileKey parameters are not in alphabetical order`** *(Parameters)*: the report shows the current and sorted parameter lists
- **`Trailing semicolon (;) in parameters`** *(Semicolons)*: a semicolon right before the pipe
- `RECURSE or REMOVESELF is incorrectly spelled, or there are too many pipe (|) symbols.` *(always on)*
- **`Missing pipe (|) before RECURSE.`** / **`Missing pipe (|) before REMOVESELF.`** *(Flags)*
- **`Backslash (\) found before pipe (|).`** *(Slashes)*

### Detect, DetectFile, and RegKey paths

- **`Incorrect registry root casing.`** *(Casing)*: e.g. `hkcu\` for `HKCU\`
- `LangSecRef holds an invalid value.` *(Invalid Values)*: the value isn't one of CCleaner's section numbers that WinappDebug recognizes
- `Invalid registry path detected.` *(Path Validity)*
- `Invalid file system path detected.` *(Path Validity)*
- `Illegal characters (< > ") detected in filesystem path.` *(Path Validity)*
- **`Trailing backslash (\) found in DetectFile`** *(Slashes)*
- `Nested wildcard found in DetectFile` *(always on)*: winapp2ool's Trim handles these, but CCleaner and System Ninja don't.
- `Wildcard (*) found in DetectFile`: [System Ninja flavor](#flavor-aware-checks) only

### ExcludeKey

- **`Missing pipe (|) after ExcludeKey flag`** *(Flags)*
- **`Missing backslash (\) before pipe (|) in ExcludeKey.`** *(Path Validity)*
- `No valid exclude flag (FILE, PATH, or REG) found in ExcludeKey.` *(Flags)*
- `ExcludeKey has too many flags` *(Flags)*
- `ExcludeKey contains REG flag in BleachBit flavor`: [BleachBit flavor](#flavor-aware-checks) only

### Default

- **`Entry has a Default key where there should be none`** *(Defaults)*: the default policy. The key is removed
- **`No Default Key found`** *(always on under the audit)*: [Default Value Audit](#default-key-policy) only. The report and the insertion both ignore the Defaults switches, so with the audit on, every entry ends up with a `Default` key
- **`Incorrect value for Default Key found`** *(Defaults)*: Default Value Audit only. The value is corrected

### Optimizations (Off by default)

- FileKeys that match another key in the same entry on both path and flag are merged into one key. See [Example 15](EXAMPLES.md#example-15-merging-redundant-filekeys-optimizations)

---

# Troubleshooting

| Symptom | Cause |
|:-|:-|
| A **File Chooser** opens with the header `winapp2.ini does not exist` | WinappDebug couldn't find the input file |
| The menu header says `winapp2.ini was empty or not found` | The input file exists but has no entries |
| Errors are reported but no file appears | Saving is off, which is the default |
| The summary says `winapp2-debugged.ini was not saved` | WinappDebug couldn't write to the save target |
| Log Viewer isn't in the menu | The last run found no errors, or you haven't run WinappDebug yet this session |

---

# WinappDebug in the winapp2.ini Build

Every build of winapp2.ini runs WinappDebug inside each entry generator, on its output and on its `archived.ini`, and once more as the last stage of each flavor and of the archive.

## The reconciliation guard

Each generator's lint pass goes through the Lint Reconciler, which compares generator's entries before and after the lint. The lint may make syntactic changes such as reordering or renumbering keys, sorting FileKey patterns, merging FileKeys, dropping duplicates, and changing case. Any semantic differences, such as a dropped entry, a malformed or discarded key, or a rewritten value, means the generator wrote invalid data that was destroyed by the lint, so the guard fails the build. It reports every difference it found and saves the entries as they were before the lint to `<name>.prelint.ini` beside the output.
