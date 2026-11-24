# Trim

**Trim** is an optimizer for winapp2.ini. Every entry in winapp2.ini has detection criteria (a registry key or a file path) used to determine whether the software that it targets is installed. Trim evaluates that criteria against a particular machine, keeping only entries that target software currently installed on that machine, reducing startup time for CCleaner when winapp2.ini is installed.

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
7. [Notes](#notes)
8. [Troubleshooting](#troubleshooting)
9. **[Usage Examples](EXAMPLES.md)**

**For Contributors**

10. [How Trimming Works](#how-trimming-works)
    - [Evaluation Order](#evaluation-order)
    - [DetectOS](#detectos)
    - [Detect](#detect)
    - [DetectFile](#detectfile)
    - [SpecialDetect](#specialdetect)
    - [Environment Variables](#environment-variables)
    - [VirtualStore Handling](#virtualstore-handling)

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
| Toggle downloading | Turn downloading the latest winapp2.ini from GitHub on or off | Default: `Off`; unavailable in offline mode |
| Toggle include list | Turn the includes list on or off | Default: `Off` |
| Toggle exclude list | Turn the excludes list on or off | Default: `Off` |
| Choose winapp2.ini | Select a different local winapp2.ini to trim | Only shown when downloading is disabled |
| Choose save target | Select a different output file path | Default: overwrites the input file |
| Choose includes file | Select the includes file | Only shown when include list is enabled |
| Choose excludes file | Select the excludes file | Only shown when exclude list is enabled |
| Reset Settings | Restore all settings to their defaults | Only shown when settings have been changed |

---

# Command-Line Arguments

Invoke Trim from the command-line with `winapp2ool -trim` (`trim`, `-2`, and `2` also work)

Command-line runs always start from Trim's default settings.

### Toggles

| Arg | Effect |
|:-|:-|
| `-d` | Download the latest winapp2.ini from GitHub to use as the input  |
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

Trim supports overriding the result of the detection criteria using dedicated include and exclude lists. Matching entry names in winapp2.ini are kept or discarded irrespective of their detection criteria when using this feature. Turn them on with **Toggle include list** and **Toggle exclude list** in the menu or with `-includes` and `-excludes` from the command-line. 

Both are plain ini files whose section names are the entries to match:

```ini
[Some Application *]
[Another Application *]
```

Trim ignores any keys in those sections, so you can leave them empty. See [Example 4](EXAMPLES.md#example-4-never-trim-and-always-trim-overrides).

---

# Notes

### Re-trim After Changes

A trimmed file matches the winapp2.ini from which it came and the software installed when you ran Trim. Re-trim after winapp2.ini updates or installing new software.

### CCleaner 7

[CC7Patcher](../cc7patcher/readme.md) supports trimming winapp2.ini before patching CCleaner 7

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
| "Your key seems to be malformatted (bad root? ...)" (log only) | A `Detect` key uses a registry root other than `HKCC`/ `HKCU`/`HKLM`/`HKU`/`HKCR` |

| Symptom | Cause |
|:-|:-|
| An entry for installed software was removed | The detection criteria doesn't match the actual installation location; consider [opening an issue](https://github.com/MoscaDotTo/Winapp2/issues/new) with a corrected path  |
| An entry for software that isn't installed was kept | The detection criteria exists on the current machine; consider opening an issue against the entry if you believe the detection criteria is incorrect. | 
| Download option is unavailable in the menu | winapp2ool is in offline mode |

---

# How Trimming Works

###### This section covers technical details about Trim. You do not need to understand this in order to trim winapp2.ini for your system.

## Evaluation Order

1. Include list (if enabled): the entry is named → kept
2. Exclude list (if enabled): the entry is named → removed
3. `DetectOS`: present but not satisfied → removed
4. `Detect`: any of its registry paths exists → kept
5. `DetectFile`: any of its paths exists → kept
6. `SpecialDetect`: any value matches → kept
7. The entry's *only* detection key is a `DetectOS` that passed step 3 → kept
8. The entry has no detection keys at all → kept
9. Otherwise → removed

For a detailed example, see [Anatomy of a Trim](EXAMPLES.md#example-1-anatomy-of-a-trim) 

## DetectOS

`DetectOS` declares a Windows version requirement as a `major.minor` version number and takes one of three forms:

| Form | Meaning | Entry is kept when |
|:-|:-|:-|
| `DetectOS=6.1\|` | Minimum version | Windows version ≥ 6.1 |
| `DetectOS=\|6.1` | Maximum version | Windows version ≤ 6.1 |
| `DetectOS=6.1\|10.0` | Version range | 6.1 ≤ Windows version ≤ 10.0 |

A failed version check always removes the entry. If the check passes and there are no other detection keys, the entry is kept.

## Detect

`Detect` declares a registry path that Trim checks for existence.

Trim also checks keys in . If a key can't be read because of permissions, Trim assumes it exists and keeps the entry.

## DetectFile

`DetectFile` declares a file or folder that Trim checks for existence.

###### Note: Wildcards in DetectFiles are fully supported in winapp2ool, but applications consuming winapp2.ini have varying support. Avoid using them when possible.

## SpecialDetect

`SpecialDetect` is a deprecated CCleaner key. Trim still reads it so that very old winapp2.ini files work. Each holds a list of paths that may target multiple browsers. 

| Value | Browsers |
|:-|:-|
| `DET_CHROME` | CentBrowser, Chromium, CocCoc, Comodo Dragon, Epic Privacy Browser,Flock,Google Chrome,Iridium,Rockmelt,SRWare Iron,Slimjet,Torch,Vivaldi,Yandex 
| `DET_MOZILLA` | Mozilla Firefox |
| `DET_THUNDERBIRD` | Mozilla Thunderbird |
| `DET_OPERA` | Opera Browser  |

If any of a value's paths or keys exists, the entry is kept.

###### Note: Many of the `DET_CHROME` browsers are no longer under development and are left in-place for backwards compatibility only. SpecialDetect paths targeting browsers still under development may no longer be correct. 

## Environment Variables

### Variable Expansion 

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

On 64-bit systems, `%ProgramFiles%` covers both `Program Files` and `Program Files (x86)`, and `HKLM\Software` also covers `HKLM\SOFTWARE\WOW6432Node` If a file, folder, or regitry key cannot be read because of permissions, Trim assumes the target exists and keeps the entry

### Malformed Variables

A `DetectFile` with broken environment variable formatting is automatically kept and a warning is printed: 

```
<path> contains a malformed environment variable and was ignored. The entry will be retained
```

## VirtualStore Handling

Running some legacy software on modern Windows causes Windows to redirect that software's writes from protected locations to the user's VirtualStore. For every entry that is kept, Trim reads its `FileKey`, `RegKey` and `ExcludeKey` values and adds keys for the matching VirtualStore locations if they exist on the current system. 

| Original location | VirtualStore folder |
|:-|:-|
| `%ProgramFiles%` | `%LocalAppData%\VirtualStore\Program Files*` |
| `%CommonAppData%` | `%LocalAppData%\VirtualStore\ProgramData` |
| `%CommonProgramFiles%` | `%LocalAppData%\VirtualStore\Program Files*\Common Files` |
| `HKLM\Software` | `HKCU\Software\Classes\VirtualStore\MACHINE\SOFTWARE` |
