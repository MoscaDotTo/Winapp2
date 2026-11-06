# WinappDebug

**WinappDebug** is a winapp2.ini [linter](https://en.wikipedia.org/wiki/Lint_(software)) for anyone who writes or edits winapp2.ini entries. It checks winapp2.ini for style and syntax errors in a wide range of categories you can turn on and off, and it can repair many of those errors for you.

### What does WinappDebug do?

WinappDebug reads a winapp2.ini file, checks every entry against winapp2.ini's syntax and style rules, and prints each problem it finds along with the key responsible. It repairs errors in memory as it finds them. Turn saving on and it writes the repaired file to the save target, which is `winapp2-debugged.ini` unless you choose another.

### Why WinappDebug?

Run it on new entries before you contribute them, and after any hand edit to winapp2.ini, to catch typos and ordering or numbering mistakes. It also brings the style of files from Transmute or a hand merge back in line. WinappDebug is an important build tool for winapp2.ini: see [WinappDebug in the winapp2.ini Build](#winappdebug-in-the-winapp2ini-build).

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

1. Put your `winapp2.ini` in the same folder as winapp2ool
2. Open WinappDebug from the main menu
3. Select **Run**. WinappDebug lists every error it finds and doesn't change any file

### Scan and repair

1. Open WinappDebug
2. Select **Toggle Saving** to turn saving on
3. Select **Run**. WinappDebug repairs every error it can and saves the result to `winapp2-debugged.ini`

###### Note: WinappDebug saves to `winapp2-debugged.ini` and leaves your input file alone unless you tell it otherwise. To overwrite the input instead, point **File Chooser (save)** (or `-3f` on the command line) at it.

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

In the menu, you make the same choices with **Toggle Saving** and **File Chooser (save)** before you select **Run**, and **Scan Settings** turns individual checks on and off. See [Menu Options](#menu-options). `-s`, `-offline` and `-writelog` are winapp2ool [global args](../../Readme.md#global-args), not WinappDebug's own.

---


# Menu Options

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

On the command line, WinappDebug is `winapp2ool -debug` (`debug`, `-1` and `1` also work).

Command-line runs always use the default scan settings. Changes you've saved in Scan Settings from the menu don't apply to them. The exceptions are `-opti` and `-keepdefaults`, which each change one rule for a single run and leave your saved settings alone.

### Toggles

| Arg | Effect |
|:-|:-|
| `-c` | Save the corrected file |
| `-usedate` | Use today's date (yymmdd) as the version line in the saved file's header, e.g. `; Version: 260724` |
| `-opti` | Turn on the experimental Optimizations rule (the FileKey merger) for this run |
| `-keepdefaults` | Leave existing `Default=` keys alone instead of reporting and removing them; see [Default Key Policy](#default-key-policy) |

Two of winapp2ool's global flags matter for scripted lint runs. They aren't WinappDebug arguments: they work with any module and can go anywhere on the command line.

| Arg | Effect |
|:-|:-|
| `-s` | Silent mode. Suppresses **all** console output, so a lint run under `-s` prints nothing at all |
| `-writelog` | Write `winapp2ool.log` even when the run succeeds. Without it, a silent run only writes the log when it exits with a nonzero code |

### File Selection

| Arg | Effect | Default |
|:-|:-|:-|
| `-1d path` | Set the input winapp2.ini directory | Current directory |
| `-1f name` | Set the input winapp2.ini file name | `winapp2.ini` |
| `-1f subdir\name` | Use a file in a subfolder of the input directory | |
| `-3d path` | Set the output directory | Current directory |
| `-3f name` | Set the output file name | `winapp2-debugged.ini` |
| `-3f subdir\name` | Save to a subfolder of the output directory | |

###### Note: The output goes to `winapp2-debugged.ini` unless you name another file, so your input stays as it was. To repair the input in place, pass its name to `-3f`.

### Flavor Selection

The global flavor flags turn on a flavor's [flavor-aware checks](#flavor-aware-checks) for one run. They don't change the Flavor setting saved in the menu.

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

WinappDebug goes through winapp2.ini one entry at a time. It checks every key against the rules that are turned on and records each error, repairing it on the spot if that category's repair is on. When every entry is done, it sorts the entry list alphabetically and reports any entry that was out of place.

Each category has two switches:

- The **scan** finds the error and reports it
- The **repair** fixes it in memory

Both are on for every category except Optimizations. With a scan on and its repair off, you get the report and the content stays as it was. Repairs happen in memory and get reported whether or not you save, but WinappDebug only writes a file when saving is on (`-c` on the command line).

Each error is reported as a block naming the entry, the error, and usually the key responsible:

```
Error in [7-Zip *]:
Forward slash (/) detected in lieu of backslash (\).
Key: FileKey1=%AppData%/7-Zip|*.tmp
```

At the end, a summary gives the entry count and the number of possible errors. That number is green at 0, yellow below 10, and red at 10 or more. **Log Viewer** in the menu shows the full report again.

###### Note: Because WinappDebug repairs errors as it finds them, a later message can name a key by the value or number an earlier repair gave it. [Example 5](EXAMPLES.md#example-5-renumbering-and-re-alphabetization) and [Example 8](EXAMPLES.md#example-8-flag-and-excludekey-repairs) both show this.

## Scan Settings

The **Scan Settings** menu turns each category on or off. It lists all fifteen **Scan Options** first (items 1–15), then the matching fifteen **Repair Options** (items 16–30).

Each switch pulls the other along in one direction:

- Turning a category's scan off turns its repair off too
- Turning a category's repair on turns its scan back on if it is off
- Turning only the repair off leaves the scan on, so you still see the category's errors but they don't get fixed. See [Example 3](EXAMPLES.md#example-3-reporting-without-repairing)

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

Some checks don't belong to any category and always run, whatever Scan Settings says. [Detected Errors](#detected-errors) marks them *(always on)*.

Scan Settings changes you make in the menu are saved and carry over to your next session, but only menu runs use them. Command-line runs always lint with the defaults. Two flags change that for a single run: `-opti` turns the Optimizations rule on, and `-keepdefaults` turns the Defaults rule off.

## Default Key Policy

Current winapp2.ini style leaves `Default=` keys out of entries, and by default WinappDebug enforces that: it reports any `Default` key (`Entry has a Default key where there should be none`), and the repair removes it.

**Toggle Default Value Audit** turns that around. Every entry must then have a `Default` key with a set value, `False` unless you change it with **Toggle Expected Default**. With the audit on:

- An entry with no `Default` key is reported (`No Default Key found`) and gets one with the expected value, whatever the Defaults scan and repair switches say
- A `Default` key with the wrong value is reported (`Incorrect value for Default Key found`) and corrected

See [Example 13](EXAMPLES.md#example-13-auditing-default-key-values).

`-keepdefaults` only exists on the command line. It leaves `Default=` keys exactly as it finds them, neither removing them nor requiring them.

## Flavor-Aware Checks

A few checks depend on winapp2ool's global **Flavor** setting. You'll find it in the Winapp2ool Settings menu, or you can set it for one run with the [flavor command-line flags](#flavor-selection):

| Flavor | Additional check |
|:-|:-|
| System Ninja | `Wildcard (*) found in DetectFile`: System Ninja doesn't support wildcards in detection keys |
| BleachBit | `ExcludeKey contains REG flag in BleachBit flavor`: BleachBit doesn't support registry ExcludeKeys |

---

# Detected Errors

These are the exact messages the report prints. WinappDebug fixes a **bold** message automatically while its category's repair is on. A plain message is only reported, and you'll have to fix it yourself. The category in parentheses is the [Scan Settings](#scan-settings) entry that controls the check, unless it's marked *(always on)*.

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

### Optimizations (experimental, off by default)

- FileKeys that match another key in the same entry on both path and flag are merged into one key. See [Example 15](EXAMPLES.md#example-15-merging-redundant-filekeys-optimizations)

---

# Troubleshooting

| Symptom | Cause |
|:-|:-|
| A **File Chooser** opens with the header `winapp2.ini does not exist` | WinappDebug couldn't find the input file |
| The menu header says `winapp2.ini was empty or not found` | The input file exists but has no entries |
| `0 possible errors detected.` on a file you expected to fail | The file is already clean, the categories you expected are turned off, or the check you expected only runs for a certain [flavor](#flavor-aware-checks) |
| Errors are reported but no file appears | Saving is off, which is the default |
| The summary says `winapp2-debugged.ini was not saved` | WinappDebug couldn't write to the save target |
| Log Viewer isn't in the menu | The last run found no errors, or you haven't run WinappDebug yet this session |

---

# WinappDebug in the winapp2.ini Build

This section is for people who maintain the winapp2.ini build, or who want to know where WinappDebug fits in it. You don't need it to use WinappDebug.

Every build of winapp2.ini runs WinappDebug nine times: three times inside the entry generators, and once as the last stage of each of six flavors.

## Where it runs

| Stage | Invocation | Gated? |
|:-|:-|:-|
| BrowserBuilder output | `remotedebugGuarded` (in-process, Optimizations forced on) | Yes, by the [reconciliation guard](#the-reconciliation-guard) |
| UWPBuilder output | `remotedebugGuarded` (in-process, Optimizations forced on) | Yes |
| EntryBuilder output | `remotedebugGuarded` (in-process, Optimizations forced on) | Yes |
| Base winapp2.ini | `-s -offline -debug -usedate -c -opti -1f Winapp2.ini -3f Winapp2.ini` | No |
| CCleaner flavor | same, on `winapp2-ccleaner-flavor.ini` | No |
| FluentCleaner flavor | same, **plus `-keepdefaults`** | No |
| BleachBit flavor | same, on `winapp2-bleachbit-flavor.ini` | No |
| Tron flavor | same, on `winapp2-tron-flavor.ini` | No |
| System Ninja flavor | same, output renamed to `Winapp2.rules` | No |

###### Note: The CCleaner 7 flavor gets no lint pass of its own because a Transmute pass makes it from the already-linted CCleaner flavor.

## The reconciliation guard

Each generator's lint pass goes through the Lint Reconciler, which compares what the generator's entries mean before and after the lint. The lint may reorder and renumber keys, sort FileKey patterns, merge FileKeys that share a path and flag, drop exact duplicates, and change case. Any other difference, such as a dropped entry, a malformed key thrown away, or a rewritten value, means the generator wrote invalid data and the lint quietly destroyed it. That's a bug in the generator's source data, so the guard fails the build. It reports every difference it found and saves the entries as they were before the lint to `<name>.prelint.ini` beside the output.

