# Trim

**Trim** is a winapp2ool module that strips winapp2.ini down to the entries that apply to the machine on which it runs. Every entry carries detection criteria (a registry key, a file path, a Windows version requirement) saying whether the software it targets is actually installed. Trim checks those criteria against the current system and drops the entries whose software isn't there.

### What does Trim do?

Trim reads a winapp2.ini (off disk, or downloaded from GitHub), tests every entry against the machine it is running on, and saves a copy with only the entries that passed. 

### Why Trim?

- Performance: CCleaner evaluates every entry in winapp2.ini when it starts, and it is slow about it. Trimming typically removes around 90% of the file, so CCleaner starts faster
- Automation: `winapp2ool -trim -d` downloads the latest database and gives you a machine-specific copy in one command

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
    - [Entries Without Detection Keys](#entries-without-detection-keys)
    - [Environment Variables](#environment-variables)
    - [VirtualStore Handling](#virtualstore-handling)

**Examples**

11. **[Examples](EXAMPLES.md)**: five worked examples with captured input and output

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

Two optional override files let you overrule detection. Trim consults both before it looks at any detection key.

**Include list** (`includes.ini` by default): Name an entry here and it is always kept, whether or not its detection criteria are satisfied. Turn it on with **Toggle include list** in the menu, or `-includes` on the command line.

**Exclude list** (`excludes.ini` by default): Entries named here are always removed, even when their detection criteria pass. **Toggle exclude list** in the menu, `-excludes` on the command line.

Each file is a plain ini file where section names are the entry names to match:

```ini
[Some Application *]
[Another Application *]
```

Whatever keys those sections contain is ignored, so you can leave them empty. 

See [Example 4](EXAMPLES.md#example-4-never-trim-and-always-trim-overrides) 

---

# Output Formatting

Trim always writes its output as a fully formatted winapp2.ini file:

- Entries are sorted alphabetically and grouped into the winapp2.ini ordering for the selected flavor 
- The winapp2.ini preamble comments are regenerated at the top of the file: the `; Version:` line is carried over from the input file (or written as `; Version: 000000` if the input had none), and the `; # of entries:` line is updated to the post-trim count
- **Comments are not preserved.** Comments in the input file do not appear in the output. Since the default save target overwrites the input file, trimming an annotated custom file in place will lose your comments
- Key order within entries is unchanged, except that entries receiving [VirtualStore keys](#virtualstore-handling) have their keys renumbered

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
| "Failed to download winapp2.ini" (red menu header) | The download didn't complete; nothing is saved and the run exits 1 |
| "winapp2.ini can't be downloaded in offline mode" | `-d` was given with `-offline`; nothing is saved and the run exits 1 |
| "\<path\> contains a malformed environment variable and was ignored" (log only) | A `DetectFile` value has a broken `%Variable%`; the entry is retained |
| "Your key seems to be malformatted (bad root? ...)" (log only) | A `Detect` key uses a registry root other than `HKCU`/`HKLM`/`HKU`/`HKCR` |

| Symptom | Cause |
|:-|:-|
| An entry for installed software was removed | The `DetectFile`/`Detect` path doesn't match the actual installation location (e.g. a portable or non-standard install location). [How Trimming Works](#how-trimming-works) explains how each path is checked |
| Download option is unavailable in the menu | winapp2ool is in offline mode |

---

# How Trimming Works

This section is for people writing or reviewing winapp2.ini entries, or trying to understand why Trim kept or removed an entry. It explains how Trim reads each detection criteria, which decide whether an entry survives the trim. You are not expected to understand this in order to use Trim.

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

- Paths under `HKLM\Software` are also checked under `HKLM\SOFTWARE\WOW6432Node`, so 32-bit applications on 64-bit Windows are found 
- If a key can't be read because of permissions, Trim assumes it exists and keeps the entry

## DetectFile

`DetectFile` names a path that Trim checks for existence. It can point at either a file or a folder.

- Wildcards (`*`) work anywhere in the path. Trim expands each wildcard segment against the folders actually on disk (the last segment against files too, eg. `%ProgramFiles%\App\*.exe`), and keeps the entry if any expansion resolves to something that exists. 
- Paths under `%ProgramFiles%` are checked in both the native Program Files directory and the 32-bit `Program Files (x86)` directory. See [Environment Variables](#environment-variables)
- If a folder can't be read because of permissions, Trim assumes the target exists and keeps the entry

###### Note: Wildcards in DetectFiles are fully supported in winapp2ool, but applications consuming winapp2.ini have varying support. Avoid using them when possible.

## SpecialDetect

`SpecialDetect` is a deprecated CCleaner variable. Trim still handles it so that very old winapp2.ini files work. Four values are recognized:

| Value | What is checked |
|:-|:-|
| `DET_CHROME` | A hardcoded list of ~26 Chromium-family browser paths and registry keys |
| `DET_MOZILLA` | `%AppData%\Mozilla\Firefox` |
| `DET_THUNDERBIRD` | `%AppData%\Thunderbird` |
| `DET_OPERA` | `%AppData%\Opera Software` |

If any of the associated paths or keys exists on the system, the entry is kept.

###### Note: Many of the `DET_CHROME` browser identities belong to browsers that are no longer under active development and are left in-place for backwards compatibility only.

## Entries Without Detection Keys

Entries with no detection keys of any type are always kept

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

On 64-bit systems, `%ProgramFiles%` covers both the native Program Files directory and `Program Files (x86)`. `HKLM\Software` registry paths get the same treatment through `WOW6432Node`.

**Malformed variables:** a `DetectFile` whose `%` has no closing `%` (e.g. `%AppData\App`) counts as present, so the entry is kept. Trim notes it in the log:

```
<path> contains a malformed environment variable and was ignored. The entry will be retained
```

## VirtualStore Handling

Pre-UAC applications that wrote to protected locations had those writes quietly redirected into the user's VirtualStore, and machines upgraded from older Windows versions still carry the leftovers. So for every entry that survives detection, Trim reads the `FileKey`, `RegKey`, and `ExcludeKey` values and adds keys for the matching VirtualStore locations, but only where those locations exist:

| Key type | Original location | VirtualStore counterpart |
|:-|:-|:-|
| FileKey / ExcludeKey | `%ProgramFiles%` | `%LocalAppData%\VirtualStore\Program Files*` |
| FileKey / ExcludeKey | `%CommonAppData%` | `%LocalAppData%\VirtualStore\ProgramData` |
| FileKey / ExcludeKey | `%CommonProgramFiles%` | `%LocalAppData%\VirtualStore\Program Files*\Common Files` |
| FileKey / ExcludeKey / RegKey | `HKLM\Software` | `HKCU\Software\Classes\VirtualStore\MACHINE\SOFTWARE` |

The original location matches in any case (`HKLM\SOFTWARE` too). For a RegKey naming a value (`HKLM\Software\App|Setting`), Trim checks that the key holding the value exists in the VirtualStore. On most systems no VirtualStore keys are generated at all. When keys are added, the entry's keys are renumbered. See [Example 5](EXAMPLES.md#example-5-virtualstore-key-generation).
