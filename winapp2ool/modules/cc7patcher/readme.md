# CC7Patcher

**CC7Patcher** is a winapp2ool module for end users that installs winapp2.ini entries into CCleaner 7's configuration file (`ccleaner.ini`). CCleaner 7 no longer loads a separate `winapp2.ini`. Its cleaning definitions live inside `ccleaner.ini` instead, in a modified entry format. CC7Patcher uses Transmute's Add mode to put the [CCleaner 7 flavor](#the-ccleaner-7-format) of winapp2.ini into `ccleaner.ini`.

### What does CC7Patcher do?

CC7Patcher performs these steps in order:

1. Downloads the CCleaner 7 flavor from GitHub (or accepts a local copy)
2. Optionally trims winapp2.ini to keep only entries relevant to the current system
3. Prunes any winapp2 entries already in `ccleaner.ini`
4. Adds the current winapp2.ini entries to `ccleaner.ini`

The patched `ccleaner.ini` holds CCleaner's own cleaning definitions and the winapp2.ini entries side by side, and CCleaner 7 cleans with both.

### Why CC7Patcher?

CCleaner 7 ignores a winapp2.ini dropped next to its executable, so without CC7Patcher you'd be merging entries into `ccleaner.ini` by hand. CC7Patcher downloads and installs in one run. It can also trim the database down to the software on your machine first, which leaves CCleaner fewer entries to load, and everything it does can be scripted from the command line.

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

Run winapp2ool from the folder that holds the files you're working on. Every file setting defaults to the current directory.

### Default Workflow (download and patch)

1. Open CC7Patcher from the winapp2ool main menu
2. Use **Change ccleaner.ini** to point to your `ccleaner.ini` if necessary
3. Select **Run**. CC7Patcher downloads the latest CCleaner 7 flavor of winapp2.ini and adds it to `ccleaner.ini`

### Local winapp2.ini Workflow

1. Open CC7Patcher
2. Toggle off **Download**. The **Change winapp2.ini** option will appear
3. Use **Change winapp2.ini** to select your local CCleaner 7 flavor winapp2.ini if necessary
4. Use **Change ccleaner.ini** to select your ccleaner.ini if necessary
5. Select **Run**. CC7Patcher adds your local copy to `ccleaner.ini`

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
| Download, trim, and patch `ccleaner.ini` in place in CCleaner's configuration directory | `winapp2ool -cc7patcher -trim -2d "%ProgramFiles%\Piriform\CCleaner 7" -3d "%ProgramFiles%\Piriform\CCleaner 7"` | |
| Trim before patching, but never remove the entries named in `keepers.ini` | `winapp2ool -cc7patcher -trim -includes -4f keepers.ini` | |

In the menu, these are toggles and file pickers you set before **Run**. See [Menu Options](#menu-options).

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

On the command line, CC7Patcher is `winapp2ool -cc7patcher` (`cc7patcher`, `12`, and `-12` also work).

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

###### Note: Winapp2ool does not expand environment variables in path arguments. `%ProgramFiles%` in the commands on this page and in the [examples](EXAMPLES.md) works because the Command Prompt (cmd) expands it before winapp2ool sees it. From PowerShell, use `$env:ProgramFiles` instead.

---

# The CCleaner 7 Format

A CCleaner 7 entry differs from a base winapp2.ini entry in two ways:

- It carries three extra keys: `ID=` (the entry's name), `Author=`, and `Tags=`
- It has no `LangSecRef=` or `Section=`. A comma-delimited `Tags=` value (e.g. `Tags=Google,Chrome,Browser`) does their job

```ini
[Windows System Assessment Tool *]
DetectFile=%WinDir%\Performance\WinSAT
FileKey1=%WinDir%\Performance\WinSAT|*|REMOVESELF
ID=Windows System Assessment Tool *
Author=Winapp2.ini Project
Tags=ccwindows
```

The winapp2.ini project publishes a ready-made [CCleaner 7 flavor](https://github.com/MoscaDotTo/Winapp2/tree/master/Non-CCleaner/CCleaner7) in this format. A local file has to be in this format too, or its entries won't work in CCleaner 7.

---

# How Patching Works

CC7Patcher first prunes the target `ccleaner.ini`, removing every entry with an `Author=Winapp2.ini Project` key. Then it uses [Transmute's](../transmute/readme.md) **Add** mode to add the winapp2.ini entries to the pruned file. Because of the prune, patching an already patched `ccleaner.ini` gives the same result as patching a clean one, whether you're re-running the same winapp2.ini or updating to a newer one. See [Example 5](EXAMPLES.md#example-5-re-patching-in-place).

### Notes

If you edit an entry or write your own for CCleaner 7, take out its `Author=Winapp2.ini Project` key. Otherwise the next patch deletes it, and an edited entry goes back to our original.

---

# Tips & Best Practices

### Back Up ccleaner.ini First

By default the output overwrites your input `ccleaner.ini`, and the rewrite also sorts its sections and strips its comments. Back it up before running, or use **Change output file** (`-3f`) to write the patched result somewhere else and review it first. See [Example 4](EXAMPLES.md#example-4-review-before-deploying).

### Include and Exclude Lists

The trim before patching uses include and exclude lists the same way [Trim](../trim/README.md#includes-and-excludes) does, but CC7Patcher has its own toggles and its own files (`cc7-includes.ini` and `cc7-excludes.ini` by default).

### Re-patching After Updates

When winapp2.ini updates, run CC7Patcher again over your patched `ccleaner.ini`. You don't need to restore a clean copy first.

### CCleaner 7 updates overwrite ccleaner.ini

When CCleaner 7 updates, it rewrites its cleaning definitions in `ccleaner.ini`, and the winapp2.ini entries go with them. Run CC7Patcher again after every CCleaner 7 update.

---

# Troubleshooting

| Message | Cause |
|:-|:-|
| "ccleaner.ini not found. Please select a valid file." | The target `ccleaner.ini` does not exist. The run aborts without writing anything |
| "Internet connection required to download winapp2.ini. Please check your connection." | Downloading is enabled but there is no connection. Toggle Download off and provide a local file, or restore your connection |
| "Failed to download winapp2.ini" | The download from GitHub failed. Check your connection and try again |
| "winapp2.ini not found. Please select a valid file." | Downloading is disabled and the local winapp2.ini does not exist |
| "CCleaner.ini not found" | The menu shows this when you select Run and the target `ccleaner.ini` does not exist |
| Entries not appearing in CCleaner 7 | The patched file was not saved to the correct `ccleaner.ini`, or a non-CCleaner 7 flavor winapp2.ini was patched |
