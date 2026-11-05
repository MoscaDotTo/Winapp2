# Trim

**Trim** is a winapp2ool module that reduces the full database down to the set of entries that apply to the machine on which it runs.

### What does Trim do?

Every entry in winapp2.ini has detection criteria (a registry key or a file path) used to determine whether the software it cleans is installed. Trim reads a winapp2.ini, from disk or straight from GitHub, checks those keys against the current machine, and saves a copy holding only the entries that passed.

### Why Trim?

CCleaner evaluates every entry in winapp2.ini when it starts, and it is slow about it. Trimming typically removes around 90% of the file, so CCleaner starts faster. Trim can also fetch the latest database itself, eg. `winapp2ool -trim -d` gets you a copy fitted to your machine in one command.

---

# Table of Contents

**Using Trim**

1. [Requirements](#requirements)
2. [Quick Start](#quick-start)
3. [Common Uses](#common-uses)
4. [Menu Options](#menu-options)
5. [Command-Line Arguments](#command-line-arguments)
   - [Toggles](#toggles)
   - [File Selection](#file-selection)
6. [Includes and Excludes](#includes-and-excludes)
7. [Output Formatting](#output-formatting)
8. [Tips & Best Practices](#tips--best-practices)
9. [Troubleshooting](#troubleshooting)

**For Contributors**

10. [How Trimming Works](#how-trimming-works)
    - [Evaluation Order](#evaluation-order)
    - [DetectOS](#detectos)
    - [Detect](#detect)
    - [DetectFile](#detectfile)
    - [SpecialDetect](#specialdetect)
    - [Environment Variables](#environment-variables)
    - [VirtualStore Handling](#virtualstore-handling)

**Examples**

11. **[Examples](EXAMPLES.md)**: five worked examples with captured input and output

---

# Requirements

- A `winapp2.ini` file to trim, **or** an internet connection with downloading enabled (`-d`)

---

# Quick Start

### From a local file

1. Put your `winapp2.ini` in the same folder as winapp2ool
2. Open Trim from the winapp2ool main menu
3. Run. Trim checks every entry and saves the result

The trimmed file overwrites the input `winapp2.ini` unless you pick another path with **Choose save target** first.

### Downloading first

1. Open Trim
2. Enable **Toggle downloading** to use the latest winapp2.ini from GitHub
3. Run. Trim downloads and trims in one step

---

# Common Uses

| To | Run | Example |
|:-|:-|:-|
| Trim `winapp2.ini` in the current directory, overwriting it | `winapp2ool -trim` | [Anatomy of a Trim](EXAMPLES.md#example-1-anatomy-of-a-trim) |
| Trim to a new file and leave the original alone | `winapp2ool -trim -3f winapp2-trimmed.ini` | [Trimming the Full Database](EXAMPLES.md#example-2-trimming-the-full-database) |
| Download the latest winapp2.ini and trim it | `winapp2ool -trim -d` | [Downloading and Trimming](EXAMPLES.md#example-3-downloading-and-trimming-in-one-step) |
| Keep an entry Trim would remove, or remove one it would keep | `winapp2ool -trim -includes -excludes` | [Never-Trim and Always-Trim](EXAMPLES.md#example-4-never-trim-and-always-trim-overrides) |
| Download and trim from a script, with no output or prompts | `winapp2ool -trim -d -s` | |

In the menu, the same choices are toggles and file pickers set before **Run**; see [Menu Options](#menu-options). `-s` is one of winapp2ool's [global args](../../Readme.md#global-args).

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

---

# Includes and Excludes

Two optional files let you overrule detection; Trim reads both before it looks at any detection key.

The include list (`includes.ini` by default) keeps every entry it names, whether or not the entry's detection keys pass. Turn it on with **Toggle include list** in the menu, or `-includes` on the command line.

The exclude list (`excludes.ini` by default) removes every entry it names, even when the entry's detection keys pass. Turn it on with **Toggle exclude list** in the menu, or `-excludes` on the command line.

Both are plain ini files whose section names are the entries to match:

```ini
[Some Application *]
[Another Application *]
```

Trim ignores any keys in those sections, so you can leave them empty. See [Example 4](EXAMPLES.md#example-4-never-trim-and-always-trim-overrides).

---

# Output Formatting

Trim always writes a fully formatted winapp2.ini:

- Entries are sorted alphabetically and grouped with any winapp3.ini entries preceding all winapp2.ini entries
- The header comments at the top of the file are rewritten. The `; Version:` line comes from the input file (`; Version: 000000` if it had none), and `; # of entries:` lists the count after trimming
- **Comments are not preserved.** None of the input file's comments reach the output, and since the default save target is the input file, trimming an annotated file in place loses them
- Keys keep their order within each entry, except in entries that gain [VirtualStore keys](#virtualstore-handling), which are renumbered

---

# Tips & Best Practices

### Save to a Different File First

By default, Trim overwrites the input `winapp2.ini`, and [comments are not preserved](#output-formatting). The first time, save to a different file (e.g. `winapp2-trimmed.ini`) and check the result before you replace your working copy.

### Re-trim After Changes

A trimmed file matches the winapp2.ini from which it came and the software installed when you ran Trim. Re-trim after winapp2.ini updates or installing new software.

### CCleaner 7

[CC7Patcher](../cc7patcher/readme.md) can trim winapp2.ini itself before it patches the entries into CCleaner 7, so you don't need to run Trim first.

---

# Troubleshooting

| Message | Cause |
|:-|:-|
| "Internet connection lost! Please check your network connection and try again" | Downloading is enabled but no connection is available |
| File Chooser appears with header "winapp2.ini does not exist" | The input file was not found |
| "winapp2.ini was empty or not found" (red menu header) | The input file exists but contains no entries |
| "Failed to download winapp2.ini" (red menu header) | The download didn't complete; nothing is saved and the run exits 1 |
| "winapp2.ini can't be downloaded in offline mode" | `-d` was given with `-offline`; nothing is saved and the run exits 1 |
| "\<path\> contains a malformed environment variable and was ignored" (log only) | A `DetectFile` value has a broken `%Variable%`; the entry is kept |
| "Your key seems to be malformatted (bad root? ...)" (log only) | A `Detect` key uses a registry root other than `HKCU`/`HKLM`/`HKU`/`HKCR` |

| Symptom | Cause |
|:-|:-|
| An entry for installed software was removed | The `DetectFile`/`Detect` path doesn't match the actual installation location (e.g. a portable or non-standard install location). [How Trimming Works](#how-trimming-works) explains how each path is checked |
| An entry for software that isn't installed was kept | The `DetectFile`/`Detect` path exists on the current machine. If you previously had the software installed, it may have left this folder or registry key behind when it was uninstalled. Otherwise, consider filing an issue against the entry if you believe the detection criteria is incorrect. | 
| Download option is unavailable in the menu | winapp2ool is in offline mode |

---

# How Trimming Works

This section intended for people who write or review winapp2.ini entries, or who want to know why Trim kept or removed one. It covers technical details about how Trim reads each detection key. You do not need to understand this in order to trim winapp2.ini for your system.

## Evaluation Order

Trim checks each entry in this order, and stops at the first step that gives an answer.

1. Include list (if enabled): the entry is named → kept
2. Exclude list (if enabled): the entry is named → removed
3. `DetectOS`: present but not satisfied → removed
4. `Detect`: any of its registry paths exists → kept
5. `DetectFile`: any of its paths exists → kept
6. `SpecialDetect`: any value matches → kept
7. The entry's *only* detection key is a `DetectOS` that passed step 3 → kept
8. The entry has no detection keys at all → kept
9. Otherwise → removed

[Anatomy of a Trim](EXAMPLES.md#example-1-anatomy-of-a-trim) runs six entries through these steps and says which step decided each one.

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

- Trim also checks a path under `HKLM\Software` under `HKLM\SOFTWARE\WOW6432Node`, where 32-bit programs on 64-bit Windows write
- If a key can't be read because of permissions, Trim assumes it exists and keeps the entry

## DetectFile

`DetectFile` names a file or folder that Trim checks for existence.

- Wildcards (`*`) work anywhere in the path. Trim expands each wildcard segment against the folders on disk (the last segment against files too, eg. `%ProgramFiles%\App\*.exe`) and keeps the entry if any expansion exists
- Trim checks a path under `%ProgramFiles%` in both `Program Files` and `Program Files (x86)`. See [Environment Variables](#environment-variables)
- If a folder can't be read because of permissions, Trim assumes the target exists and keeps the entry

###### Note: Wildcards in DetectFiles are fully supported in winapp2ool, but applications consuming winapp2.ini have varying support. Avoid using them when possible.

## SpecialDetect

`SpecialDetect` is a deprecated CCleaner key. Trim still reads it so that very old winapp2.ini files work, and recognizes four values:

| Value | What is checked |
|:-|:-|
| `DET_CHROME` | A hardcoded list of ~26 Chromium-family browser paths and registry keys |
| `DET_MOZILLA` | `%AppData%\Mozilla\Firefox` |
| `DET_THUNDERBIRD` | `%AppData%\Thunderbird` |
| `DET_OPERA` | `%AppData%\Opera Software` |

If any of a value's paths or keys exists, the entry is kept.

###### Note: Many of the `DET_CHROME` browser identities belong to browsers that are no longer under active development and are left in-place for backwards compatibility only.

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

###### Note: These variables only work in CCleaner, and are no longer shipped in the official winapp2.ini. Support for them is left in-place for backwards compatibility only.

On 64-bit systems, `%ProgramFiles%` covers both `Program Files` and `Program Files (x86)`.

A `DetectFile` with a `%` that never closes (e.g. `%AppData\App`) counts as present, so the entry is kept, and Trim logs:

```
<path> contains a malformed environment variable and was ignored. The entry will be retained
```

## VirtualStore Handling

Windows quietly redirected pre-UAC programs' writes to protected locations from those locations and into the user's VirtualStore, and machines upgraded from older Windows versions still carry the leftovers. So for every entry that survives detection, Trim reads its `FileKey`, `RegKey` and `ExcludeKey` values and adds keys for the matching VirtualStore locations, but only where those locations exist:

| Key type | Original location | VirtualStore counterpart |
|:-|:-|:-|
| FileKey / ExcludeKey | `%ProgramFiles%` | `%LocalAppData%\VirtualStore\Program Files*` |
| FileKey / ExcludeKey | `%CommonAppData%` | `%LocalAppData%\VirtualStore\ProgramData` |
| FileKey / ExcludeKey | `%CommonProgramFiles%` | `%LocalAppData%\VirtualStore\Program Files*\Common Files` |
| FileKey / ExcludeKey / RegKey | `HKLM\Software` | `HKCU\Software\Classes\VirtualStore\MACHINE\SOFTWARE` |

The original location matches in any case (`HKLM\SOFTWARE` too). For a RegKey naming a value (`HKLM\Software\App|Setting`), Trim checks that the key holding the value exists in the VirtualStore. Most systems get no VirtualStore keys at all. When Trim adds some, it renumbers the entry's keys. See [Example 5](EXAMPLES.md#example-5-virtualstore-key-generation).
