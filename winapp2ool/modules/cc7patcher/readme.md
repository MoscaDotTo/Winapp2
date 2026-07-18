# CC7Patcher

**CC7Patcher** is a winapp2ool module that installs winapp2.ini entries into CCleaner 7's configuration file (`ccleaner.ini`). CCleaner 7 dropped support for loading a separate `winapp2.ini` file; instead, cleaning definitions are stored inside `ccleaner.ini` alongside CCleaner's own cleaning definitions, and requiring a modified entry format. CC7Patcher bridges this gap by patching the [CCleaner 7 flavor](#the-ccleaner-7-format) of winapp2.ini into `ccleaner.ini` using Transmute's Add mode.

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

1. [Requirements](#requirements)
2. [Quick Start](#quick-start)
3. [Menu Options](#menu-options)
4. [The CCleaner 7 Format](#the-ccleaner-7-format)
5. [How Patching Works](#how-patching-works)
6. [Command-Line Arguments](#command-line-arguments)
   - [Toggles](#toggles)
   - [File Selection](#file-selection)
   - [Examples](#examples)
7. [Tips & Best Practices](#tips--best-practices)
8. [Troubleshooting](#troubleshooting)
9. [Usage Examples](#usage-examples)
    - [Example 1: Patching a ccleaner.ini](#example-1-patching-a-ccleanerini)
    - [Example 2: Trimming before patching](#example-2-trimming-before-patching)
    - [Example 3: The default workflow (download and deploy)](#example-3-the-default-workflow-download-and-deploy)
    - [Example 4: Review before deploying](#example-4-review-before-deploying)
    - [Example 5: Re-patching in place](#example-5-re-patching-in-place)

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

# Menu Options

| Option | Effect | Notes |
|:-|:-|:-|
| Run (default) | Install winapp2.ini for CCleaner 7 | |
| Toggle Trim | Enable or disable trimming winapp2.ini before patching | Default: `False` |
| Toggle Download | Enable or disable downloading the latest CCleaner 7 flavor from GitHub | Default: `True`; unavailable in offline mode |
| Change winapp2.ini | Select the local winapp2.ini to use as input | Only shown when downloading is disabled |
| Change ccleaner.ini | Select the `ccleaner.ini` file to patch | Default: `ccleaner.ini` in the current directory |
| Change output file | Select where to save the patched `ccleaner.ini` | Default: `ccleaner.ini` in the current directory (overwrites the input when both are default) |
| Reset Settings | Restore all settings to their defaults | Only shown when settings have been changed |

The menu also displays the currently selected winapp2.ini (shown as `Online` while downloading is enabled), ccleaner.ini, and output file paths.

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

CC7Patcher first prunes the target `ccleaner.ini` by removing any entry with a `Author=Winapp2.ini Project` key. It then uses [Transmute's](../transmute/README.md) **Add** mode to add the winapp2.ini entries into the pruned file. This makes patching idempotent: running it again, or running it to update to a newer winapp2.ini, produces the same result as patching a clean `ccleaner.ini`. See [Example 5](#example-5-re-patching-in-place).

### Notes

If you modify or create an entry for your system for use with CCleaner 7, remove the `Author=Winapp2.ini Project` key to prevent CC7Patcher from destroying it. 

---

# Command-Line Arguments

CC7Patcher supports command-line automation for scripting environments. The module is invoked as `winapp2ool -cc7patcher` (`cc7patcher`, `12`, and `-12` also work).

### Toggles

| Arg | Effect |
|:-|:-|
| `-nodownload` | Use a local winapp2.ini instead of downloading |
| `-trim` | Trim winapp2.ini before patching |

### File Selection

| Arg | Effect | Default |
|:-|:-|:-|
| `-1d path` | Set winapp2.ini directory | Current directory |
| `-1f name` | Set winapp2.ini file name | `winapp2.ini` |
| `-2d path` | Set ccleaner.ini directory | Current directory |
| `-2f name` | Set ccleaner.ini file name | `ccleaner.ini` |
| `-3d path` | Set output file directory | Current directory |
| `-3f name` | Set output file name | `ccleaner.ini` (overwrites input) |

###### Note: Winapp2ool does not expand environment variables in path arguments. `%AppData%` in the examples below works because the Command Prompt (cmd) expands it before winapp2ool sees it. From PowerShell, use `$env:APPDATA` instead.

### Examples

| Command | Effect |
|:-|:-|
| `winapp2ool -cc7patcher` | Download the latest CCleaner 7 flavor and patch `ccleaner.ini` in the current directory |
| `winapp2ool -cc7patcher -trim` | Download, trim for the current system, then patch |
| `winapp2ool -cc7patcher -nodownload -3f ccleaner-patched.ini` | Patch using a local CCleaner 7 format winapp2.ini, saving the result separately |
| `winapp2ool -cc7patcher -trim -2d "%AppData%\CCleaner" -3d "%AppData%\CCleaner"` | Download, trim, and patch ccleaner.ini in place in CCleaner's configuration directory |

---

# Tips & Best Practices

### Back Up ccleaner.ini First

The default output overwrites your input `ccleaner.ini`, and the overwrite also reorders the file and strips its comments. Back it up before running, or use **Change output file** (`-3f`) to write the patched result to a separate location for review first. See [Example 4](#example-4-review-before-deploying).

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


---

# Usage Examples

The outputs below are taken verbatim from real CC7Patcher runs. The `winapp2.ini` used in Examples 1, 2, and 5 is this four-entry excerpt of the real CCleaner 7 flavor:

###### **Source file (`winapp2.ini`)**

```ini
[Google Chrome Autoplay Preferences *]
DetectFile=%LocalAppData%\Google\Chrome*
FileKey1=%LocalAppData%\Google\Chrome*\User Data\MEIPreload|*|REMOVESELF
ID=Google Chrome Autoplay Preferences *
Author=Winapp2.ini Project
Tags=Google,Chrome,Browser

[Opera GX Autoplay Preferences *]
DetectFile=%AppData%\Opera Software\Opera GX*
FileKey1=%AppData%\Opera Software\Opera GX*\_side_profiles\*\MEIPreload|*|REMOVESELF
FileKey2=%AppData%\Opera Software\Opera GX*\MEIPreload|*|REMOVESELF
ID=Opera GX Autoplay Preferences *
Author=Winapp2.ini Project
Tags=OperaGX,Browser

[Windows System Assessment Tool *]
DetectFile=%WinDir%\Performance\WinSAT
FileKey1=%WinDir%\Performance\WinSAT|*|REMOVESELF
ID=Windows System Assessment Tool *
Author=Winapp2.ini Project
Tags=ccwindows

[Windows Task Scheduler *]
Detect=HKCU\Software\Microsoft\Windows
FileKey1=%WinDir%\System32\Tasks_Migrated|*|REMOVESELF
ID=Windows Task Scheduler *
Author=Winapp2.ini Project
Tags=ccwindows
```

The `ccleaner.ini` being patched is this minimal illustration:

###### **Base file (`ccleaner.ini`)**

```ini
; CCleaner - Cleaning rules

[CCleaner Entry 1]
ID=CCleaner.Entry1
Author=CCleaner
Tags=ccapps
```

---

### Example 1: Patching a ccleaner.ini

**Context**

We have a local CCleaner 7 format winapp2.ini and want to install its entries into `ccleaner.ini`. To keep the original file untouched for now, we direct the output to a separate file.

**Intent**

We want to append every entry from winapp2.ini into `ccleaner.ini`, saving the result as `ccleaner-patched.ini`.

**Command**

```
winapp2ool -cc7patcher -nodownload -3f ccleaner-patched.ini
```

###### Note: Without `-3f`, the output would overwrite `ccleaner.ini` in place

**Output**

###### **Output file (`ccleaner-patched.ini`)**

```ini
[CCleaner Entry 1]
ID=CCleaner.Entry1
Author=CCleaner
Tags=ccapps

[Google Chrome Autoplay Preferences *]
DetectFile=%LocalAppData%\Google\Chrome*
FileKey1=%LocalAppData%\Google\Chrome*\User Data\MEIPreload|*|REMOVESELF
ID=Google Chrome Autoplay Preferences *
Author=Winapp2.ini Project
Tags=Google,Chrome,Browser

[Opera GX Autoplay Preferences *]
DetectFile=%AppData%\Opera Software\Opera GX*
FileKey1=%AppData%\Opera Software\Opera GX*\_side_profiles\*\MEIPreload|*|REMOVESELF
FileKey2=%AppData%\Opera Software\Opera GX*\MEIPreload|*|REMOVESELF
ID=Opera GX Autoplay Preferences *
Author=Winapp2.ini Project
Tags=OperaGX,Browser

[Windows System Assessment Tool *]
DetectFile=%WinDir%\Performance\WinSAT
FileKey1=%WinDir%\Performance\WinSAT|*|REMOVESELF
ID=Windows System Assessment Tool *
Author=Winapp2.ini Project
Tags=ccwindows

[Windows Task Scheduler *]
Detect=HKCU\Software\Microsoft\Windows
FileKey1=%WinDir%\System32\Tasks_Migrated|*|REMOVESELF
ID=Windows Task Scheduler *
Author=Winapp2.ini Project
Tags=ccwindows
```

**Explanation**

- All four winapp2.ini entries are appended into the file, exactly as written in the source
- The `[CCleaner Entry 1]` section and every key in it are untouched
- The output is sorted alphabetically by section name 
- The CCleaner 7 header comments are incidentally stripped from `ccleaner.ini`

---

### Example 2: Trimming before patching

**Context**

Most of the 3,700+ entries in winapp2.ini target software we don't have. There is no reason to make CCleaner 7 parse all of them at every startup. 

**Intent**

We want to patch `ccleaner.ini` with only the entries relevant to this machine. The demo machine has neither Google Chrome nor Opera GX installed.

**Command**

```
winapp2ool -cc7patcher -nodownload -trim -3f ccleaner-trimmed.ini
```

**Output**

The run reports `Trimming complete: 2 entries remain` before patching.

###### **Output file (`ccleaner-trimmed.ini`)**

```ini
[CCleaner Entry 1]
ID=CCleaner.Entry1
Author=CCleaner
Tags=ccapps

[Windows System Assessment Tool *]
DetectFile=%WinDir%\Performance\WinSAT
FileKey1=%WinDir%\Performance\WinSAT|*|REMOVESELF
ID=Windows System Assessment Tool *
Author=Winapp2.ini Project
Tags=ccwindows

[Windows Task Scheduler *]
Detect=HKCU\Software\Microsoft\Windows
FileKey1=%WinDir%\System32\Tasks_Migrated|*|REMOVESELF
ID=Windows Task Scheduler *
Author=Winapp2.ini Project
Tags=ccwindows
```

**Explanation**

- Trim evaluated each entry in winapp2.ini's detection criteria against the demo machine before patching
- `[Google Chrome Autoplay Preferences *]` and `[Opera GX Autoplay Preferences *]` were dropped because their `DetectFile` paths do not exist on the demo machine
- `[Windows System Assessment Tool *]` and `[Windows Task Scheduler *]` detection criteria was found on the dmeo machine. These entries are retained and then appended into `ccleaner.ini`

---

### Example 3: The default workflow (download and patch)

**Context**

Download the current CCleaner 7 flavor and install it directly into CCleaner's live configuration.

**Intent**

We want to patch the real `ccleaner.ini` in `%ProgramFiles%\Piriform\CCleaner 7`

**Command**

```
winapp2ool -cc7patcher -2d "%ProgramFiles%\Piriform\CCleaner 7" -3d "%ProgramFiles%\Piriform\CCleaner 7"
```

**Output**

The run appends all entries from the current CCleaner 7 flavor into `ccleaner.ini`.

###### Note: The full `ccleaner.ini` produced by this command is too large to display here. The console output has been provided instead. 

```
Patching ccleaner.ini with entries from winapp2.ini
Patched file saved to C:\Program Files\Piriform\CCleaner 7\ccleaner.ini
CCleaner 7 patching complete
```

**Explanation**

- `-2d`/`-3d` point the input and output at CCleaner's configuration directory
- After the run, the winapp2.ini entries appear in CCleaner 7's cleaning options alongside its built-in definitions

---

### Example 4: Review before deploying

**Context**

We are cautious about a tool rewriting CCleaner's live configuration and want to inspect the result before committing to it.

**Intent**

We want to produce the patched file in the working directory so that we can review it. 

**Commands**

```
copy "%ProgramFiles%\Piriform\CCleaner 7\ccleaner.ini" .
winapp2ool -cc7patcher -3f ccleaner-patched.ini
```

Then, after reviewing `ccleaner-patched.ini` in a text editor:

```
copy "%ProgramFiles%\Piriform\CCleaner 7\ccleaner.ini" "%ProgramFiles%\Piriform\CCleaner 7\ccleaner.ini.bak"
copy ccleaner-patched.ini "%ProgramFiles%\Piriform\CCleaner 7\ccleaner.ini"
```

**Explanation**

- `ccleaner.ini` is copied into the into the working directory from the CCleaner 7 directory
- `-3f ccleaner-patched.ini` writes the patched result to a separate file from the input for review 
- CCleaner 7's current unpatched `ccleaner.ini` is renamed to `ccleaner.ini.bak` as a backup 
- The newly patched `ccleaner-patched.ini` is copied into the CCleaner 7 directory and renamed to `ccleaner.ini`
---

### Example 5: Re-patching in place

**Context**

A new winapp2.ini version has been released, and we want to update our already-patched `ccleaner.ini` 

**Intent**

Show that re-patching is idempotent: CC7Patcher prunes the winapp2 entries the previous patch added, then re-adds the current set, so nothing is duplicated.

**Command**

```
winapp2ool -cc7patcher -nodownload -2f ccleaner-patched.ini -3f ccleaner-repatched.ini
```

**Output**

###### **Output file (`ccleaner-repatched.ini`)**

```ini
[Google Chrome Autoplay Preferences *]
DetectFile=%LocalAppData%\Google\Chrome*
FileKey1=%LocalAppData%\Google\Chrome*\User Data\MEIPreload|*|REMOVESELF
ID=Google Chrome Autoplay Preferences *
Author=Winapp2.ini Project
Tags=Google,Chrome,Browser
```

**Explanation**

- This entry was already in `ccleaner-patched.ini` 
- The prune pass removes it because it has an `Author=Winapp2.ini Project` key
- The Add pass put it back 
- If this version of winapp2.ini had dropped an entry, the old copy would have been deleted without being readded
- Any entry not written by the Winapp2.ini Project is left unmodified
