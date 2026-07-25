# Downloader

**Downloader** is a winapp2ool module that fetches files from the winapp2 GitHub repository. It provides both a  menu and a scriptable command line interface for acquiring any flavor of winapp2.ini, the latest winapp2ool executable, and the project's supplementary files.

### What does Downloader do?

Downloader resolves a short name (`winapp2`, `readme`, `winapp3`, etc) to a URL on the winapp2 GitHub, downloads the file, and writes it to a directory you choose. For winapp2.ini specifically, the URL it picks depends on the active flavor in winapp2ool.

### Why Downloader?

- **Flavor targeting**: fetch the exact variant your cleaner expects without hunting through the repository
- **Self-updating**: winapp2ool can replace itself, keeping a backup of the version it replaced
- **Supplementary databases**: `winapp3.ini` and `Archived entries.ini`, ready to fold into a winapp2.ini with Transmute

---

# Table of Contents

1. [Requirements](#requirements)
2. [Quick Start](#quick-start)
3. [Menu Options](#menu-options)
   - [Main Menu](#main-menu)
   - [Advanced Downloads](#advanced-downloads)
4. [How Downloads Resolve](#how-downloads-resolve)
   - [Flavor Selection](#flavor-selection)
   - [Output File Names](#output-file-names)
   - [Where Files Come From](#where-files-come-from)
5. [Download Behavior](#download-behavior)
   - [The Save Directory](#the-save-directory)
   - [Existing Files](#existing-files)
   - [Updating Winapp2ool](#updating-winapp2ool)
6. [Command-Line Arguments](#command-line-arguments)
   - [File Selection](#file-selection)
   - [Path Parameters](#path-parameters)
   - [Flavor Flags](#flavor-flags)
   - [Global Flags](#global-flags)
   - [Examples](#examples)
7. [Tips & Best Practices](#tips--best-practices)
8. [Troubleshooting](#troubleshooting)
9. [Usage Examples](#usage-examples)
   - [Acquiring winapp2.ini](#acquiring-winapp2ini)
     - [Example 1: Downloading winapp2.ini](#example-1-downloading-winapp2ini)
     - [Example 2: Choosing a flavor without changing your settings](#example-2-choosing-a-flavor-without-changing-your-settings)
     - [Example 3: Downloading into a directory that doesn't exist yet](#example-3-downloading-into-a-directory-that-doesnt-exist-yet)
     - [Example 4: Renaming the download](#example-4-renaming-the-download)
   - [Scripting](#scripting)
     - [Example 5: Silent, scriptable downloads](#example-5-silent-scriptable-downloads)
     - [Example 6: Downloading over a file that already exists](#example-7-downloading-over-a-file-that-already-exists)
   - [Updating Winapp2ool](#updating-winapp2ool-1)
     - [Example 7: Updating winapp2ool in place](#example-8-updating-winapp2ool-in-place)
     - [Example 8: Downloading winapp2ool.exe without updating](#example-9-downloading-winapp2oolexe-without-updating)
   - [Advanced Downloads](#advanced-downloads-1)
     - [Example 9: winapp3.ini and Archived entries.ini](#example-10-winapp3ini-and-archived-entriesini)

---

# Requirements

- An active internet connection. Downloader is hidden from the main menu in offline mode
- .NET Framework 4.6 or later to download `winapp2ool.exe` (every other file works on 4.5)

---

# Quick Start

### Common Workflow

1. Open Downloader from the winapp2ool main menu
2. Select the file you want from the menu
3. The file is saved to the current save directory (default: current directory)

To save elsewhere, use **Change Save Directory** before downloading.

From the command line, one invocation downloads one file:

```
winapp2ool -download winapp2
```

---

# Menu Options

## Main Menu

| Option | Effect | Notes |
|:-|:-|:-|
| Winapp2.ini | Download the base (non-CCleaner) winapp2.ini | |
| CCleaner Winapp2.ini | Download the CCleaner flavor | |
| CCleaner 7 Winapp2.ini | Download the CCleaner 7 flavor | For installing into `ccleaner.ini`, see [CC7Patcher](../cc7patcher/readme.md) |
| BleachBit Winapp2.ini | Download the BleachBit flavor | |
| FluentCleaner Winapp2.ini | Download the FluentCleaner flavor | |
| System Ninja Winapp2.rules | Download the System Ninja flavor | Saved as `winapp2.rules` |
| Tron | Download the Tron flavor | |
| Winapp2ool | Download the latest winapp2ool.exe | Requires .NET 4.6+. Replaces the running executable when the save directory is the current directory — see [Updating Winapp2ool](#updating-winapp2ool) |
| ReadMe | Download the top-level winapp2ool readme | Saved as `readme.md` |
| Advanced | Open the Advanced Downloads sub-menu | See [Advanced Downloads](#advanced-downloads) |
| Change Save Directory | Select a new directory for all downloaded files | Default: current directory |
| Reset Settings | Restore the save directory to its default | Only shown when the save directory has been changed |

The seven winapp2.ini options are fixed to their flavors, they ignore the active flavor setting and always download what their label says. Only the command line resolves a download against the active flavor.

## Advanced Downloads

| Option | Effect | Notes |
|:-|:-|:-|
| Winapp3.ini | Download the extended/potentially unsafe entry database | Not recommended for general use |
| Archived entries.ini | Download entries for old or discontinued software | Not recommended for general use |

---

# How Downloads Resolve

## Flavor Selection

The flavor only matters for `winapp2`

The active flavor comes from winapp2ool's global settings (**Settings → Change Flavor** on the main menu), but any command-line [flavor flag](#flavor-flags) overrides it for that run. Command-line runs never save settings.

###### Note: the default flavor is **CCleaner**, not base. On a winapp2ool that has never had its flavor changed, `winapp2ool -download winapp2` fetches the CCleaner flavor. Pass `-base` (or `-ncc`) for the non-CCleaner build. 

## Output File Names

| Request | Active flavor | File written |
|:-|:-|:-|
| `winapp2` | Base | `winapp2.ini` |
| `winapp2` | CCleaner *(default)* | `winapp2.ini` |
| `winapp2` | BleachBit | `winapp2.ini` |
| `winapp2` | System Ninja | `winapp2.rules` |
| `winapp2` | Tron | `winapp2.ini` |
| `winapp2` | CCleaner 7 | `winapp2.ini` |
| `winapp2` | FluentCleaner | `winapp2.ini` |
| `winapp2ool` | *(not applicable)* | `winapp2ool.exe` |
| `readme` | *(not applicable)* | `readme.txt` from the CLI, `readme.md` from the menu |
| `winapp3` | *(not applicable)* | `winapp3.ini` |
| `archived` | *(not applicable)* | `Archived entries.ini` |

System Ninja is the only flavor that changes the extension. `-1f` overrides any of these names.

## Where Files Come From

Everything is fetched from `github.com/MoscaDotTo/Winapp2` on the `master` branch:

| File | Path in the repository |
|:-|:-|
| winapp2.ini (base) | `Non-CCleaner/Winapp2.ini` |
| winapp2.ini (CCleaner) | `Winapp2.ini` |
| winapp2.ini (BleachBit) | `Non-CCleaner/BleachBit/Winapp2.ini` |
| winapp2.rules (System Ninja) | `Non-CCleaner/SystemNinja/Winapp2.rules` |
| winapp2.ini (Tron) | `Non-CCleaner/Tron/Winapp2.ini` |
| winapp2.ini (CCleaner 7) | `Non-CCleaner/CCleaner7/Winapp2.ini` |
| winapp2.ini (FluentCleaner) | `Non-CCleaner/FluentCleaner/Winapp2.ini` |
| winapp2ool.exe | `winapp2ool/bin/Release/winapp2ool.exe` |
| readme | `winapp2ool/Readme.md` |
| winapp3.ini | `Winapp3/Winapp3.ini` |
| Archived entries.ini | `Winapp3/Archived entries.ini` |

With **Beta Participation** enabled (main menu → Settings), `winapp2ool.exe` is taken from the `Branch1` branch instead of `master`. No other file is affected by the beta setting.

---

# Download Behavior

## The Save Directory

The save directory defaults to the directory from which winapp2ool is running. If the download target directory does not exist, it is created, including any missing parent directories. 

## Existing Files

When the target file already exists and output is not suppressed, Downloader prompts before writing:

```
readme.txt already exists in the target directory.
Enter a new file name, or leave blank to overwrite the existing file:
```

Entering a name saves under that name; leaving it blank overwrites. This prompt appears on the command line too, where it will block a script waiting for input. Passing `-s` skips the prompt entirely and overwrites silently.

## Updating Winapp2ool

Downloading `winapp2ool.exe` into the directory from which winapp2ool is currently running does not simply write a file. Instead, winapp2ool:

1. Renames the running executable to `winapp2ool v<version>.exe.bak` in that directory
2. Writes the freshly downloaded executable as `winapp2ool.exe`
3. Launches the new executable and exits
---

# Command-Line Arguments

Downloader is command-line module **8**; `winapp2ool -download` and `winapp2ool 8` calls both work.

The module takes a single argument naming the file to download, plus optional path parameters. The file argument is required and must be valid, if it is in valid or missing, Downloader reports that no file was specified and exits with code 1.

One run downloads one file. To fetch several, run the command several times.

## File Selection

| Arg | File downloaded |
|:-|:-|
| `1` or `winapp2` | winapp2.ini, in the [active flavor](#flavor-selection) |
| `2` or `winapp2ool` | winapp2ool.exe |
| `3` or `readme` | readme.txt |
| `4` or `winapp3` | winapp3.ini |
| `5` or `archived` | Archived entries.ini |

## Path Parameters

| Arg | Effect | Default |
|:-|:-|:-|
| `-1d path` | Set the save directory. Created if it doesn't exist | Current directory |
| `-1f name` | Set the save filename | Set by the file argument |
| `-1f subdir\name` | Save under a subdirectory of the save directory | |

## Flavor Flags

Processed before the module runs, and applied to that invocation only.

| Flag | Alias | Flavor |
|:-|:-|:-|
| `-base` | `-ncc` | Base (non-CCleaner) |
| `-ccleaner` | `-cc` | CCleaner |
| `-bleachbit` | `-bb` | BleachBit |
| `-systemninja` | `-sn` | System Ninja |
| `-tron` | | Tron |
| `-ccleaner7` | `-cc7` | CCleaner 7 |
| `-fluentcleaner` | `-fc` | FluentCleaner |

## Global Flags

| Flag | Effect |
|:-|:-|
| `-s` | Suppress all output, skip the overwrite prompt, and exit when finished. Functionally required for unattended scripts |
| `-writelog` | Write `winapp2ool.log` to disk on exit |

###### Note: without `-s`, winapp2ool does not exit after the download.

## Examples

| Command | Effect |
|:-|:-|
| `winapp2ool -download winapp2` | Download winapp2.ini in the active flavor to the current directory |
| `winapp2ool -download winapp2 -cc` | Download the CCleaner flavor, whatever the active flavor is |
| `winapp2ool -download 1 -cc7 -1d "C:\Tools"` | Download the CCleaner 7 flavor to `C:\Tools` |
| `winapp2ool -download winapp2 -fc` | Download the FluentCleaner flavor |
| `winapp2ool -download winapp2 -sn -s` | Silently download `winapp2.rules` for System Ninja |
| `winapp2ool -download winapp2 -1f winapp2-backup.ini` | Download winapp2.ini under a different name |
| `winapp2ool -download winapp2ool -1d "C:\Tools"` | Download winapp2ool.exe to `C:\Tools` without self-updating |
| `winapp2ool -download winapp3` | Download winapp3.ini |

---

# Tips & Best Practices

### Scripts

- Always pass `-s`. It suppresses output, skips the overwrite prompt that would otherwise block on input, and makes winapp2ool exit instead of opening its menu
- The exit code reports argument validity, not download success: `1` when the file argument is missing or invalid, `0` otherwise
- Name the flavor explicitly on every call

### Self-Updating

- Downloading `winapp2ool.exe` into winapp2ool's own directory replaces the running executable and restarts it, leaving a `winapp2ool v<version>.exe.bak` behind
- The `.bak` files persist; one per version replaced.
- To keep a copy of the executable without updating, set the save directory somewhere else first

### Flavor Selection

- The menu always offers all seven flavors individually, regardless of the active flavor setting
- If you are unsure which flavor your cleaner wants, use the base (non-CCleaner) version 

### Winapp3 and Archived Entries

 `winapp3.ini` and `Archived entries.ini` are not part of the standard winapp2.ini distribution. They may contain entries that are incomplete, actively dangerous for your system, or written for software that no longer exists. Review them before use.

---

# Troubleshooting

| Message | Cause |
|:-|:-|
| `No file was specified for download` | No file argument was given. Exits with code 1 |
| `Unknown argument: {arg}` | The file argument isn't one of the five recognized names or numbers. Exits with code 1 |
| `Valid arguments are: winapp2, winapp2ool, readme, winapp3, archived` | Printed alongside either message above |
| `Download Failed.` / `Unable to download {name} to {dir}` | The download did not complete |
| `Download incomplete: {name}` (red menu header) | The menu's report of the same failure |
| `This option requires a newer version of the .NET Framework` | Downloading the executable needs .NET 4.6 or later |
| `Unable to download winapp2ool to the current directory, choose another directory before trying again` | Winapp2ool is running from the Windows temp folder, or .NET is out of date. Self-updating is disabled in both cases |

| Symptom | Cause |
|:-|:-|
| Downloader is missing from the main menu | Winapp2ool is in offline mode. |
| A script hangs after starting a download | The target file already exists and the overwrite prompt is waiting on input. Add `-s` |
| A script never returns, with no prompt | Without `-s`, winapp2ool opens its interactive menu after downloading. Add `-s` |

---

# Usage Examples

The transcripts below are taken from real runs. The current directory is shown as `C:\Tools\winapp2` throughout, and file sizes reflect winapp2.ini version `251109`.

## Acquiring winapp2.ini

### Example 1: Downloading winapp2.ini

**Context**

You want the current winapp2.ini database in the directory from which winapp2ool is currently running.

**Intent**

Download winapp2.ini, and see which flavor you actually get.

**Command**
```
winapp2ool -download winapp2
```

**Output**
```
Downloading winapp2.ini...
Download Complete.
Downloaded winapp2.ini to C:\Tools\winapp2
```

**Result**
```
C:\Tools\winapp2\
    winapp2.ini          (CCleaner flavor)
    winapp2ool.exe
```

**Explanation**
- No flavor flag was given, so the default flavor (CCleaner) was downloaded 
- The save directory defaults to the current directory
- Winapp2ool does not exit afterwards; it opens its main menu. See [Example 5](#example-5-silent-scriptable-downloads) for silent mode

**Notes**

To get the base (non-CCleaner) winapp2.ini instead:

```
winapp2ool -download winapp2 -base
```
```
Downloading winapp2.ini...
Download Complete.
Downloaded winapp2.ini to C:\Tools\winapp2
```

---

### Example 2: Choosing a flavor without changing your settings

**Context**

Your winapp2ool is configured for the base flavor because that's what you work with day to day. Today you need the CCleaner 7 database for a different machine, and the System Ninja rules for a third.

**Intent**

Fetch two other flavors without touching the saved flavor setting, and without visiting the settings menu twice.

**Commands**
```
winapp2ool -download winapp2 -cc7 -1d "C:\Tools\winapp2\ex2"
winapp2ool -download winapp2 -sn  -1d "C:\Tools\winapp2\ex2"
```

**Output**
```
Downloading winapp2.ini...
Download Complete.
Downloaded winapp2.ini to C:\Tools\winapp2\ex2

Downloading winapp2.rules...
Download Complete.
Downloaded winapp2.rules to C:\Tools\winapp2\ex2
```

**Result**
```
C:\Tools\winapp2\ex2\
    winapp2.ini          (CCleaner 7 flavor)
    winapp2.rules        (System Ninja flavor)
```

**Explanation**
- `-cc7` and `-sn` override the active flavor (base) 
- Command-line runs never write settings, so the saved flavor is still base afterwards

**Notes**

The System Ninja flavor is named `winapp2.rules` which is why both files can land in one directory without the second overwriting the first

---

### Example 3: Downloading into a directory that doesn't exist yet

**Context**

You keep dated snapshots of the database so you can diff releases against each other later.

**Intent**

Download the BleachBit flavor into a dated folder without creating the folder first.

**Command**
```
winapp2ool -download winapp2 -bb -1d "C:\Tools\winapp2\ex3\bleachbit\2026-07"
```

**Output**
```
Downloading winapp2.ini...
Download Complete.
Downloaded winapp2.ini to C:\Tools\winapp2\ex3\bleachbit\2026-07
```

**Result**
```
C:\Tools\winapp2\ex3\bleachbit\2026-07\
    winapp2.ini          (BleachBit flavor)
```

**Explanation**
- Neither `ex3`, `bleachbit`, nor `2026-07` existed beforehand; all three were created
- The BleachBit flavor of winapp2.ini is downloaded and placed in the newly created `C:\Tools\winapp2\ex3\bleachbit\2026-07`
---

### Example 4: Renaming the download

**Context**

You already have a `winapp2.ini` in your working directory that you have customized, and you want the upstream copy alongside it rather than over it.

**Intent**

Download winapp2.ini under a different name.

**Command**
```
winapp2ool -download 1 -1f winapp2-upstream.ini
```

**Output**
```
Downloading winapp2-upstream.ini...
Download Complete.
Downloaded winapp2-upstream.ini to C:\Tools\winapp2
```

**Result**
```
C:\Tools\winapp2\
    winapp2.ini                  (your customized copy)
    winapp2-upstream.ini         (newly downloaded upstream copy)
```

**Explanation**
- `-1f` overrides the default name that the file argument would otherwise choose

**Notes**

- `1` and `winapp2` are interchangeable
`-1f` also accepts a leading subdirectory: `-1f "\archive\winapp3.ini"` saves into `archive\` beneath the save directory, creating it if needed. See [Example 10](#example-10-winapp3ini-and-archived-entriesini)

---

## Scripting

### Example 5: Silent, scriptable downloads

**Context**

You want a scheduled task to refresh winapp2.ini overnight. The interactive behavior, the overwrite prompt, and the main menu that opens when the download finishes, would both hang it.

**Intent**

Download with no console output or requests for input

**Command**
```
winapp2ool -download winapp2 -s
```

**Output**
```
(nothing)
```

**Explanation**
- `-s` suppresses all console output
- The overwrite prompt is skipped and any existing file is overwritten without asking
- Winapp2ool exits instead of opening its menu, with exit code `0`

**Notes**

The exit code reports argument validity, not download success: a bad argument exits `1`, but a download that fails still exits `0`. In a script, verify the file afterwards:

```powershell
winapp2ool -download winapp2 -s -1d "C:\Tools\winapp2"
$f = "C:\Tools\winapp2\winapp2.ini"
if (-not (Test-Path $f) -or (Get-Item $f).Length -eq 0) { throw "winapp2.ini download failed" }
```

---

### Example 6: Downloading over a file that already exists

**Context**

You are working interactively and re-download the readme into a directory that already has one.

**Intent**

Keep the old copy by saving the new one under a different name.

**Command**
```
winapp2ool -download readme
```

**Output**
```
readme.txt already exists in the target directory.
Enter a new file name, or leave blank to overwrite the existing file: readme-2026-07.md
Downloading readme-2026-07.md...
Download Complete.
Downloaded readme-2026-07.md to C:\Tools\winapp2
```

**Result**
```
C:\Tools\winapp2\
    readme.txt              (original file)
    readme-2026-07.md       (the new download)
```

**Explanation**
- The prompt appears because `readme.txt` already exists
- Entering `readme-2026-07.md` saved to a new file; the original was left alone
- Leaving the prompt blank and pressing enter would have overwritten `readme.txt` instead
- This prompt appears on the command line as well as in the menu, and it blocks. `-s` skips it and always overwrites. See [Example 5](#example-5-silent-scriptable-downloads)

---

## Updating Winapp2ool

### Example 7: Updating winapp2ool in place

**Context**

Your copy of winapp2ool is out of date and you want the current release.

**Intent**

Replace the running executable with the latest build from GitHub.

**Starting state**
```
C:\Tools\winapp2\
    winapp2ool.exe       (428,032 bytes, v1.6.9701.23408)
```

**Command**
```
winapp2ool -download winapp2ool
```

**Output**
```
(nothing)
```

**Result**
```
C:\Tools\winapp2\
    winapp2ool.exe                          (the new release, now running)
    winapp2ool v1.6.9701.23408.exe.bak      (the starting version, closed)
```

**Explanation**
- The save directory was the directory from which winapp2ool was running , so winapp2ool performs an automatic update
- The running executable was renamed to `winapp2ool v<version>.exe.bak` before the new one took its place
- A new winapp2ool process was started and the old one exited

**Notes**

Self-updating is disabled when winapp2ool is running from the Windows temp folder, or when .NET is older than 4.6. 

With **Beta Participation** enabled, this same command installs the `Branch1` build instead of the `master` release.

---

### Example 8: Downloading winapp2ool.exe without updating

**Context**

You want to put a copy of winapp2ool on a network share for other machines, without disturbing the copy you're running.

**Intent**

Download the executable as an ordinary file.

**Command**
```
winapp2ool -download winapp2ool -1d "C:\Tools\winapp2\tools"
```

**Output**
```
Downloading winapp2ool.exe...
Download Complete.
Downloaded winapp2ool.exe to C:\Tools\winapp2\tools
```

**Result**
```
C:\Tools\winapp2\
    winapp2ool.exe                (still running)
    tools\
        winapp2ool.exe            (the downloaded release)
```

**Explanation**
- The save directory differs from the current directory, so no auto-update
- From the menu, the equivalent is **Change Save Directory** before choosing **Winapp2ool**

---

## Advanced Downloads

### Example 9: winapp3.ini and Archived entries.ini

**Context**

You maintain a machine running software that winapp2.ini no longer covers, and you want to fold the archived entries for it into your local database. Transmute's `-w` and `-a` preset flags look for `winapp3.ini` and `Archived Entries.ini` in the source directory.

**Intent**

Download both supplementary databases into an `archive\` subdirectory, then merge winapp3.ini into a local winapp2.ini.

**Commands**
```
winapp2ool -download winapp3  -1f "\archive\winapp3.ini"
winapp2ool -download archived -1f "\archive\Archived entries.ini"
winapp2ool -transmute -add -2d "C:\Tools\winapp2\archive" -2f winapp3.ini -3f winapp2-extended.ini
```

**Output**
```
Downloading winapp3.ini...
Download Complete.
Downloaded winapp3.ini to C:\Tools\winapp2\archive

Downloading Archived entries.ini...
Download Complete.
Downloaded Archived entries.ini to C:\Tools\winapp2\archive
```

**Result**
```
C:\Tools\winapp2\
    winapp2.ini                          (starting local copy)
    winapp2-extended.ini                 (transmute output)
    archive\
        winapp3.ini                        (downloaded)
        Archived entries.ini               (downloaded)
```

**Explanation**
- The `-1f "\subdir\name"` put both files in `archive\`, which was created by the first download if it didnt exist already 
- The Transmute step writes to `winapp2-extended.ini`, leaving the original winapp2.ini intact