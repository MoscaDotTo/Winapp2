# CC7Patcher

**CC7Patcher** is a winapp2ool module for end users that installs winapp2.ini entries into CCleaner 7's configuration file (`ccleaner.ini`). CCleaner 7 dropped support for loading a separate `winapp2.ini` file; instead, cleaning definitions are stored inside `ccleaner.ini` alongside CCleaner's own cleaning definitions, and requiring a modified entry format. CC7Patcher bridges this gap by patching the [CCleaner 7 flavor](#the-ccleaner-7-format) of winapp2.ini into `ccleaner.ini` using Transmute's Add mode.

### What does CC7Patcher do?

CC7Patcher performs these steps in order:

1. Downloads the CCleaner 7 flavor from GitHub (or accepts a local copy) 
2. Optionally trims winapp2.ini to keep only entries relevant to the current system
3. Prunes any winapp2 entries already in `ccleaner.ini` 
4. Adds the current winapp2.ini entries to `ccleaner.ini`

The result is a `ccleaner.ini` that contains both CCleaner's own cleaning definitions and the winapp2.ini entries, which CCleaner 7 will then use when cleaning. 

### Why CC7Patcher?

- CCleaner 7 compatibility: winapp2.ini can no longer be dropped next to the CCleaner executable
- One-step install: download and install in a single run
- Trim integration: optionally trim the database down to just the software on your machine for performance reasons
- Automation: fully scriptable from the command line 

---

# Table of Contents

**Using CC7Patcher**

1. [Requirements](#requirements)
2. [Quick Start](#quick-start)
3. [Common Uses](#common-uses)
4. [Menu Options](#menu-options)
5. [Command-Line Arguments](#command-line-arguments)
   - [Toggles](#toggles)
   - [File Selection](#file-selection)
6. [The CCleaner 7 Format](#the-ccleaner-7-format)
7. [How Patching Works](#how-patching-works)
8. [Tips & Best Practices](#tips--best-practices)
9. [Troubleshooting](#troubleshooting)

**Examples**

10. **[Examples](EXAMPLES.md)**: five worked examples with captured input and output

---

# Requirements

- CCleaner 7 installed with a `ccleaner.ini` file (typically at `%ProgramFiles%\Piriform\CCleaner 7`)
- An internet connection if downloading winapp2.ini, or a local **CCleaner 7 flavor** winapp2.ini file if not 

---

# Quick Start

It is suggested that you run winapp2ool from the same directory as the files on which you want to operate. 

### Default Workflow (download and patch)

1. Open CC7Patcher from the winapp2ool main menu
2. Use **Change ccleaner.ini** to point to your `ccleaner.ini` if necessary 
3. Run. CC7Patcher downloads the latest CCleaner 7 flavor of winapp2.ini and appends it into `ccleaner.ini`

### Local winapp2.ini Workflow

1. Open CC7Patcher
2. Toggle off **Download**. The **Change winapp2.ini** option will appear
3. Use **Change winapp2.ini** to select your local CCleaner 7 flavor winapp2.ini if necessary 
4. Use **Change ccleaner.ini** to select your ccleaner.ini if necessary 
5. Run, CC7Patcher appends your local CCleaner 7 flavor of winapp2.ini into `ccleaner.ini`

A successful run reports its progress with these lines (the trim lines appear only when Trim is enabled):

```
Trimming winapp2.ini...
Trimming complete: X entries remain
Patching ccleaner.ini with entries from winapp2.ini
Patched file saved to C:\Program Files\Piriform\CCleaner 7\ccleaner.ini
CCleaner 7 patching complete
```

---

# Common Uses

| To | Run | Example |
|:-|:-|:-|
| Download the latest CCleaner 7 flavor and patch `ccleaner.ini` in the current directory | `winapp2ool -cc7patcher` | [The default workflow (download and patch)](EXAMPLES.md#example-3-the-default-workflow-download-and-patch) |
| Trim winapp2.ini for this machine before patching | `winapp2ool -cc7patcher -trim` | [Trimming before patching](EXAMPLES.md#example-2-trimming-before-patching) |
| Patch from a local CCleaner 7 flavor winapp2.ini, saving the result separately | `winapp2ool -cc7patcher -nodownload -3f ccleaner-patched.ini` | [Patching a ccleaner.ini](EXAMPLES.md#example-1-patching-a-ccleanerini) |
| Download and patch into a separate file you can review before deploying it | `winapp2ool -cc7patcher -3f ccleaner-patched.ini` | [Review before deploying](EXAMPLES.md#example-4-review-before-deploying) |
| Patch a `ccleaner.ini` that already has winapp2.ini entries in it | `winapp2ool -cc7patcher -nodownload -2f ccleaner-patched.ini -3f ccleaner-repatched.ini` | [Re-patching in place](EXAMPLES.md#example-5-re-patching-in-place) |
| Download, trim, and patch `ccleaner.ini` in place in CCleaner's configuration directory | `winapp2ool -cc7patcher -trim -2d "%AppData%\CCleaner" -3d "%AppData%\CCleaner"` | |
| Trim before patching, but never remove the entries named in `keepers.ini` | `winapp2ool -cc7patcher -trim -includes -4f keepers.ini` | |

In the menu, the same choices are toggles and file pickers set before **Run**; see [Menu Options](#menu-options).

---

# Menu Options

| Option | Effect | Notes |
|:-|:-|:-|
| Run (default) | Install winapp2.ini for CCleaner 7 | |
| Toggle Trim | Enable or disable trimming winapp2.ini before patching | Default: `False` |
| Toggle include list | Enable or disable CC7Patcher's includes file for the trim | Default: `False`; only shown when Trim is enabled |
| Toggle exclude list | Enable or disable CC7Patcher's excludes file for the trim | Default: `False`; only shown when Trim is enabled |
| Toggle Download | Enable or disable downloading the latest CCleaner 7 flavor from GitHub | Default: `True`; unavailable in offline mode |
| Change winapp2.ini | Select the local winapp2.ini to use as input | Only shown when downloading is disabled |
| Change ccleaner.ini | Select the `ccleaner.ini` file to patch | Default: `ccleaner.ini` in the current directory |
| Change output file | Select where to save the patched `ccleaner.ini` | Default: `ccleaner.ini` in the current directory (overwrites the input when both are default) |
| Choose includes file | Select the file naming entries the trim always keeps | Default: `cc7-includes.ini`; only shown when Trim and the include list are enabled |
| Choose excludes file | Select the file naming entries the trim always removes | Default: `cc7-excludes.ini`; only shown when Trim and the exclude list are enabled |
| Reset Settings | Restore all settings to their defaults | Only shown when settings have been changed |

The menu also displays the currently selected winapp2.ini (shown as `Online` while downloading is enabled), ccleaner.ini, and output file paths.

---

# Command-Line Arguments

CC7Patcher supports command-line automation for scripting environments. The module is invoked as `winapp2ool -cc7patcher` (`cc7patcher`, `12`, and `-12` also work).

### Toggles

| Arg | Effect |
|:-|:-|
| `-nodownload` | Use a local winapp2.ini instead of downloading |
| `-trim` | Trim winapp2.ini before patching |
| `-includes` | When trimming, always keep the entries named in the includes file |
| `-excludes` | When trimming, always remove the entries named in the excludes file |

### File Selection

| Arg | Effect | Default |
|:-|:-|:-|
| `-1d path` | Set winapp2.ini directory | Current directory |
| `-1f name` | Set winapp2.ini file name | `winapp2.ini` |
| `-2d path` | Set ccleaner.ini directory | Current directory |
| `-2f name` | Set ccleaner.ini file name | `ccleaner.ini` |
| `-3d path` | Set output file directory | Current directory |
| `-3f name` | Set output file name | `ccleaner.ini` (overwrites input) |
| `-4d path` | Set includes file directory | Current directory |
| `-4f name` | Set includes file name | `cc7-includes.ini` |
| `-5d path` | Set excludes file directory | Current directory |
| `-5f name` | Set excludes file name | `cc7-excludes.ini` |

###### Note: Winapp2ool does not expand environment variables in path arguments. `%AppData%` in the examples below works because the Command Prompt (cmd) expands it before winapp2ool sees it. From PowerShell, use `$env:APPDATA` instead.

---

# The CCleaner 7 Format

CCleaner 7 entries are not the same as base winapp2.ini entries. Compared to the base format, each entry:

- Carries three additional keys: `ID=` (the entry's name), `Author=`, and `Tags=`
- Replaces the `LangSecRef=` / `Section=` categorization keys with a comma-delimited `Tags=` value (e.g. `Tags=Google,Chrome,Browser`)

```ini
[Windows System Assessment Tool *]
DetectFile=%WinDir%\Performance\WinSAT
FileKey1=%WinDir%\Performance\WinSAT|*|REMOVESELF
ID=Windows System Assessment Tool *
Author=Winapp2.ini Project
Tags=ccwindows
```

The winapp2.ini project publishes a ready-made [CCleaner 7 flavor](https://github.com/MoscaDotTo/Winapp2/tree/master/Non-CCleaner/CCleaner7) built in this format. If you provide a local file, it must be in this format, or entries will not be functional in CCleaner 7

---

# How Patching Works

CC7Patcher first prunes the target `ccleaner.ini` by removing any entry with a `Author=Winapp2.ini Project` key. It then uses [Transmute's](../transmute/readme.md) **Add** mode to add the winapp2.ini entries into the pruned file. This makes patching idempotent: running it again, or running it to update to a newer winapp2.ini, produces the same result as patching a clean `ccleaner.ini`. See [Example 5](EXAMPLES.md#example-5-re-patching-in-place).

### Notes

If you modify or create an entry for your system for use with CCleaner 7, remove the `Author=Winapp2.ini Project` key to prevent CC7Patcher from destroying it. 

---

# Tips & Best Practices

### Back Up ccleaner.ini First

The default output overwrites your input `ccleaner.ini`, and the overwrite also reorders the file and strips its comments. Back it up before running, or use **Change output file** (`-3f`) to write the patched result to a separate location for review first. See [Example 4](EXAMPLES.md#example-4-review-before-deploying).

### Include and Exclude Lists

The trim before patching works like [Trim's](../trim/README.md#includes-and-excludes) include and exclude lists, but CC7Patcher keeps its own: its own toggles, and its own files (`cc7-includes.ini` and `cc7-excludes.ini` by default).

### Re-patching After Updates

When winapp2.ini is updated, just run CC7Patcher again over your existing `ccleaner.ini`. There is no need to restore a clean copy first.

### CCleaner 7 updates overwrite ccleaner.ini
When CCleaner 7 updates, it overwrites its cleaning definitions. This has the side effect of unpatching winapp2.ini from `ccleaner.ini`. You will need to reinstall winapp2.ini after every CCleaner 7 update.

---

# Troubleshooting

| Message | Cause |
|:-|:-|
| "ccleaner.ini not found. Please select a valid file." | The target `ccleaner.ini` does not exist. The run aborts without writing anything |
| "Internet connection required to download winapp2.ini. Please check your connection." | Downloading is enabled but there is no connection. Toggle Download off and provide a local file, or restore your connection |
| "Failed to download winapp2.ini" | The download from GitHub failed. check your connection and try again |
| "winapp2.ini not found. Please select a valid file." | Downloading is disabled and the local winapp2.ini does not exist |
| "CCleaner.ini not found" | Shown by the menu when Run is selected while the target `ccleaner.ini` does not exist |
| Entries not appearing in CCleaner 7 | The patched file was not saved to the correct `ccleaner.ini`, or a non-CCleaner 7 flavor winapp2.ini was patched |
