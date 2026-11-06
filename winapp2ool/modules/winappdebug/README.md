# WinappDebug

**WinappDebug** is a winapp2.ini [linter](https://en.wikipedia.org/wiki/Lint_(software)) for anyone who writes or edits winapp2.ini entries. It performs static analysis on winapp2.ini to detect style and syntax errors across a wide range of configurable categories, and can optionally repair many of those errors automatically. 

### What does WinappDebug do?

WinappDebug reads a winapp2.ini file, validates every entry against the winapp2.ini syntax and style rules, and reports each violation to the console with the offending key. Repairs are applied as errors are found; enabling Saving writes the repaired file back to disk.

### Why WinappDebug?

- Contribution hygiene: validates entries before they're accepted to the winapp2.ini project
- Hand-edit safety: catch typos, ordering, and numbering mistakes after manual changes
- Pipeline normalization: standardize the style of files produced by Transmute or hand-merging
- Release automation: see [WinappDebug in the winapp2.ini Build](#winappdebug-in-the-winapp2ini-build)

---

# Table of Contents

**Using WinappDebug**

1. [Requirements](#requirements)
2. [Quick Start](#quick-start)
3. [Common Uses](#common-uses)
4. [Menu Options](#menu-options)
5. [Command-Line Arguments](#command-line-arguments)
   - [Toggles](#toggles)
   - [File Selection](#file-selection)
   - [Flavor Selection](#flavor-selection)
6. [How It Works](#how-it-works)
   - [Scan, Repair, and Save](#scan-repair-and-save)
   - [Scan Settings](#scan-settings)
   - [Default Key Policy](#default-key-policy)
   - [Flavor-Aware Checks](#flavor-aware-checks)
7. [Detected Errors](#detected-errors)
8. [Troubleshooting](#troubleshooting)

**For Contributors**

9. [WinappDebug in the winapp2.ini Build](#winappdebug-in-the-winapp2ini-build)
   - [Where it runs](#where-it-runs)
   - [The reconciliation guard](#the-reconciliation-guard)

**Examples**

10. **[Examples](EXAMPLES.md)**: sixteen worked examples with captured input and output

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

# Common Uses

| To | Run | Example |
|:-|:-|:-|
| Scan `winapp2.ini` and report errors without saving | `winapp2ool -debug` | [Scanning a file for errors](EXAMPLES.md#example-1-scanning-a-file-for-errors) |
| Repair errors and save the result to `winapp2-debugged.ini` | `winapp2ool -debug -c` | [Repairing a file and saving the corrections](EXAMPLES.md#example-4-repairing-a-file-and-saving-the-corrections) |
| Repair `winapp2.ini` in place | `winapp2ool -debug -c -3f winapp2.ini` | |
| Repair with FileKey merging enabled | `winapp2ool -debug -c -opti` | [Merging redundant FileKeys (Optimizations)](EXAMPLES.md#example-15-merging-redundant-filekeys-optimizations) |
| Repair in place, leaving `Default=` keys untouched | `winapp2ool -debug -c -keepdefaults -3f winapp2.ini` | [Preserving Default keys with -keepdefaults](EXAMPLES.md#example-14-preserving-default-keys-with--keepdefaults) |
| Scan with the System Ninja flavor checks active | `winapp2ool -sn -debug` | [Flavor-aware linting (System Ninja)](EXAMPLES.md#example-11-flavor-aware-linting-system-ninja) |
| Repair in place from a script with no console output, writing `winapp2ool.log` | `winapp2ool -s -offline -writelog -debug -usedate -c -opti -1f Winapp2.ini -3f Winapp2.ini` | [Scripted in-place repair](EXAMPLES.md#example-16-scripted-in-place-repair) |

In the menu, **Toggle Saving** and **File Chooser (save)** set the same choices before **Run**, and **Scan Settings** switches individual checks; see [Menu Options](#menu-options). `-s`, `-offline` and `-writelog` are winapp2ool [global args](../../Readme.md#global-args).

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

###### Note: Repairs are applied as errors are found, so a later message may name a key by the value or number it received from an earlier repair. [Example 5](EXAMPLES.md#example-5-renumbering-and-re-alphabetization) and [Example 8](EXAMPLES.md#example-8-flag-and-excludekey-repairs) both show this.

## Scan Settings

Each scan category can be independently enabled or disabled from the **Scan Settings** sub-menu. The menu lists all fifteen **Scan Options** first (items 1–15), then the corresponding fifteen **Repair Options** (items 16–30).

The two toggles are linked in one direction each:

- Disabling a category's **scan** also disables its repair
- Enabling a category's **repair** also re-enables its scan
- Disabling only a category's **repair** leaves the scan on, so its errors are still reported but not corrected. See [Example 3](EXAMPLES.md#example-3-reporting-without-repairing)

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

See [Example 13](EXAMPLES.md#example-13-auditing-default-key-values). 

A flag is available on the command line only: `-keepdefaults` leaves `Default=` keys exactly as it finds them, neither removing them nor requiring them. 

## Flavor-Aware Checks

A few checks depend on winapp2ool's global **Flavor** setting (found in the Winapp2ool Settings menu, or set for a single run via the [flavor command-line flags](#flavor-selection)):

| Flavor | Additional check |
|:-|:-|
| System Ninja | `Wildcard (*) found in DetectFile`: System Ninja does not support wildcards in detection keys |
| BleachBit | `ExcludeKey contains REG flag in BleachBit flavor`: BleachBit does not support registry ExcludeKeys |

---

# Detected Errors

The messages below are the exact strings printed in the error report. **Bold** messages are corrected automatically when their category's repair is enabled; plain messages are reported only and must be fixed by hand. The category in brackets is the [Scan Settings](#scan-settings) entry that governs the check, except for those marked *(always on)*.

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

- FileKeys matching another key in the same entry on both path and flag are merged into a single key. See [Example 15](EXAMPLES.md#example-15-merging-redundant-filekeys-optimizations)

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

# WinappDebug in the winapp2.ini Build

This section is for people who maintain the winapp2.ini build, or who want to know where WinappDebug runs in it. You don't need it to use WinappDebug.


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

