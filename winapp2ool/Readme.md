# Winapp2ool

Winapp2ool is a custom build toolchain and client side manager for winapp2.ini. It generates and assembles winapp2.ini from its source files, and creates the release and PR changelogs for the repo. It also provides several useful features to end users, such as showing custom changelogs, downloading files from the repo, or keeping winapp2.ini up to date and trimmed for individual machines.

Winapp2ool runs in a headless mode for CLI and automation use, or as a standard console application with a text based numbered menu system. It provides a suite of utilities for: 

**Using winapp2.ini**

* [Trim](modules/trim/README.md): Reduces the database to entries relevant to the current system
* [Diff](modules/diff/readme.md): Generates a semantic changelog between any two winapp2.ini versions that tracks additions and removals, follows entries through renames and merges, and tracks keys that were captured by others or moved between entries
* [Downloader](modules/download/readme.md): Downloads winapp2.ini and related files from GitHub
* [CC7Patcher](modules/cc7patcher/readme.md): Installs winapp2.ini into CCleaner 7
* [CCiniDebug](modules/ccdebug/readme.md): Removes stale winapp2.ini settings from CCleaner Classic

**Customizing winapp2.ini**

* [Transmute](modules/transmute/readme.md): Applies individual structured patches to ini files
* [Flavorizer](modules/transmute/Flavorizer/readme.md): Applies batches of structured patches to ini files
* [WinappDebug](modules/winappdebug/README.md): Performs static analysis on winapp2.ini and corrects style and syntax errors

**And building winapp2.ini**

* [Browser Builder](modules/browserbuilder/readme.md): Generates entries for every web browser from small templates
* [Entry Builder](modules/entrybuilder/readme.md): Generates winapp2.ini entries from an optional shorthand syntax 
* [Combine](modules/combine/readme.md): Merges folders (including subfolders) of ini files into a single file

### What is winapp2.ini?

Winapp2.ini is a community maintained database of declarative cleaning routines for Windows that works with FluentCleaner, BleachBit, CCleaner (*including CCleaner 7*), System Ninja, Avira System Speedup, R-Wipe&Clean, HDCleaner, and Reg Organizer. It has [its own readme](../README.md).

---

# Table of Contents

**Using winapp2ool**

1. [Requirements](#requirements)
2. [Installation](#installation)
3. [Quick Start](#quick-start)
4. [Common Uses](#common-uses)
5. [Menu Options](#menu-options)
6. [Command-Line Arguments](#command-line-arguments)
   - [Module Args](#module-args)
   - [Global Args](#global-args)
   - [Flavor Args](#flavor-args)
   - [File Selection Args](#file-selection-args)
7. [Usage Examples](#usage-examples)
   - [Example 1: First run (updating and trimming from the menu)](#example-1-first-run-updating-and-trimming-from-the-menu)
   - [Example 2: Scripted download](#example-2-scripted-download)
   - [Example 3: Downloading and trimming in one command](#example-3-downloading-and-trimming-in-one-command)
   - [Example 4: Choosing a Flavor](#example-4-choosing-a-flavor)
   - [Example 5: Fully offline maintenance chain](#example-5-fully-offline-maintenance-chain)
8. [Troubleshooting](#troubleshooting)
9. [Notes](#notes)

---

# Requirements

### Minimum

* Windows 7 SP1
* .NET Framework 4.8
* Administrator rights only when saving into a protected folder such as Program Files (see [Notes](#notes))

### Suggested

* Windows 10 or higher
* Network connection for full functionality

---

# Installation

Download the latest [winapp2ool.exe](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2ool.exe) and place it in the folder where you keep (or would like to keep) winapp2.ini.

### Updates

When a new version of winapp2ool is out, the main menu offers `Update Winapp2ool`. The update keeps the version it replaced beside the new one as `winapp2ool v<version>.exe.bak`. Remove the `.bak` to go back to it.

### Beta builds

Beta builds are occasionally made available while features are in development. They live on [Branch1](https://github.com/MoscaDotTo/Winapp2/tree/Branch1). To try one, turn on `Beta Participation` in the [winapp2ool settings](modules/maintool/readme.md), and winapp2ool will download and launch the latest beta if one is available.  

---

# Quick Start

Put winapp2ool.exe in an empty folder and run it. If you're online, winapp2ool sees that there's no winapp2.ini beside it and puts the download options right on the main menu:

```
 ╔══════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╗
 ║                                                 Update available for winapp2.ini                                             ║
 ╠══════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╣
 ║                                  Welcome to Winapp2ool! Check out the ReadMe on GitHub for help!                             ║
 ║                                                                                                                              ║
 ║                                                  Menu: Enter a number to select                                              ║
 ║                                                                                                                              ║
 ║ 0. Exit                            - Exit the application                                                                    ║
 ║                                                                                                                              ║
 ║ 1. WinappDebug                     - Scan for and correct style and syntax errors in winapp2.ini                             ║
 ║ 2. Trim                            - Optimize winapp2.ini for your system                                                    ║
 ║ 3. Transmute                       - Add, replace, or remove entire sections or individual keys from winapp2.ini             ║
 ║ 4. Diff                            - Generate a context-aware changelog between two winapp2.ini files                        ║
 ║ 5. CCiniDebug                      - Remove stale winapp2.ini configurations from ccleaner.ini                               ║
 ║ 6. Entry Lab                       - Generate winapp2.ini entries from templates                                             ║
 ║ 7. Combine                         - Join together a collection of ini files into one                                        ║
 ║ 8. CC7Patcher                      - Install winapp2.ini for CCleaner 7                                                      ║
 ║                                                                                                                              ║
 ║ 9. Downloader                      - Download files from the Winapp2 GitHub                                                  ║
 ║ 10. Settings                       - Manage Winapp2ool's settings                                                            ║
 ║                                                                                                                              ║
 ║                                            A new version of winapp2.ini is available!                                        ║
 ║                                                Current: v000000 (file not found)                                             ║
 ║                                                        Available: v251109                                                    ║
 ║                                                                                                                              ║
 ║ 11. Update Winapp2.ini             - Update your local copy of winapp2.ini                                                   ║
 ║ 12. Update & Trim                  - Download and trim the latest winapp2.ini                                                ║
 ║ 13. Show winapp2.ini changelog     - See the difference between your local file and the latest                               ║
 ║ 14. Show trimmed changelog         - See the difference between your trimmed local file and the latest                       ║
 ╚══════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╝

Enter a number, or leave blank to run the default:
```

###### Note: This menu was captured from a real first run in an empty folder. `Current: v000000 (file not found)` means there's no local winapp2.ini yet. Once you have one, the menu shows its version instead. The update options at the bottom only appear while an update is available.

From here, most people want one of these:

1. `Update Winapp2.ini` downloads the latest winapp2.ini
2. `Update & Trim` downloads it and trims it down to the entries that apply to your system
3. `Show winapp2.ini changelog` shows what changed between your local copy and the latest
4. `Update Winapp2ool` updates winapp2ool itself. It only appears when a new version is out

---

# Common Uses

| To | Run | Example |
|:-|:-|:-|
| Get a winapp2.ini trimmed for your machine, from the menu | Choose `Update & Trim` on the main menu | [First run](#example-1-first-run-updating-and-trimming-from-the-menu) |
| Download the latest winapp2.ini from a script, with no output or prompts | `winapp2ool download winapp2 -s` | [Scripted download](#example-2-scripted-download) |
| Download the latest winapp2.ini and trim it in one command | `winapp2ool -2 -d -s` | [Downloading and trimming in one command](#example-3-downloading-and-trimming-in-one-command) |
| Download a flavor other than base | `winapp2ool -bleachbit download winapp2 -s` | [Choosing a Flavor](#example-4-choosing-a-flavor) |
| Apply your own additions in `custom.ini` to winapp2.ini, then lint the result, without going online | `winapp2ool -transmute -add -2f custom.ini -3f winapp2.ini -s -offline`, then `winapp2ool -debug -c -1f winapp2.ini -3f winapp2.ini -s -offline` | [Fully offline maintenance chain](#example-5-fully-offline-maintenance-chain) |

Each module's readme has its own Common Uses table near the top. Every module is also a numbered option in the menu (see [Menu Options](#menu-options)). `-s` and `-offline` are described under [Global Args](#global-args), and `-bleachbit` under [Flavor Args](#flavor-args).

---

# Menu Options

Click a module's name to see its readme

| Option | Effect | Notes |
|:-|:-|:-|
| Exit | Exits the application | |
| [WinappDebug](modules/winappdebug/README.md) | Scan for and correct style and syntax errors in winapp2.ini | Enforces the winapp2.ini style and syntax guidelines |
| [Trim](modules/trim/README.md) | Optimize winapp2.ini for your system | Removes entries not relevant to your machine, reducing load times in tools like CCleaner |
| [Transmute](modules/transmute/readme.md) | Add, replace, or remove entire sections or individual keys from winapp2.ini | Patches whole sections or single keys. [Flavorizer](modules/transmute/Flavorizer/readme.md), which applies patches in batches, opens from this menu |
| [Diff](modules/diff/readme.md) | Generate a context-aware changelog between two winapp2.ini files | |
| [CCiniDebug](modules/ccdebug/readme.md) | Remove stale winapp2.ini configurations from ccleaner.ini | For CCleaner 6 and earlier only |
| [Entry Lab](modules/entrylab/readme.md) | Generate winapp2.ini entries from templates | Menu containing [Browser Builder](modules/browserbuilder/readme.md) and [Entry Builder](modules/entrybuilder/readme.md) |
| [Combine](modules/combine/readme.md) | Join together a collection of ini files into one | Includes subfolders |
| [CC7Patcher](modules/cc7patcher/readme.md) | Install winapp2.ini for CCleaner 7 | Appends winapp2.ini entries into CCleaner 7's ccleaner.ini |
| [Downloader](modules/download/readme.md) | Download files from the Winapp2 GitHub | **Unavailable in offline mode** |
| [Settings](modules/maintool/readme.md) | Manage Winapp2ool's settings | Application-level settings: saving settings to disk, beta participation, offline mode, the active [Flavor](#flavor-args), and log management |
| Update Winapp2.ini | Update your local copy of winapp2.ini | **Only shown while a winapp2.ini update is available; unavailable in offline mode** |
| Update & Trim | Download and trim the latest winapp2.ini | **Only shown while a winapp2.ini update is available; unavailable in offline mode** |
| Show winapp2.ini changelog | See the difference between your local file and the latest | **Only shown while a winapp2.ini update is available; unavailable in offline mode** |
| Show trimmed changelog | See the difference between your trimmed local file and the latest | **Only shown while a winapp2.ini update is available; unavailable in offline mode** |
| Update Winapp2ool | Get the latest Winapp2ool.exe | **Only shown while a winapp2ool update is available. Unavailable in offline mode, and when the installed .NET Framework is older than the version winapp2ool targets (so never in winapp2oolXP)** |
| Go online | Retry your internet connection | **Only available in offline mode** |

###### Note: The main menu also accepts the hidden commands `printlog` (print winapp2ool's internal log to the console) and `savelog` (write the log to disk). See [Log Management](modules/maintool/readme.md#log-management) in the settings readme.

---

# Command-Line Arguments

Command-line arguments ("args") let a script run winapp2ool without touching the menu. The args on this page work with every module. Each module's own args are documented in its readme.

Global and flavor args can go anywhere on the command line. Of the args left after those, the first one picks the module. You can name a module by number or by name, with or without a leading `-`, so `1`, `-1`, `debug`, and `-debug` all launch WinappDebug.

### Module Args

| Arg | Effect |
|:-|:-|
| `1` or `debug` | Launches [WinappDebug](modules/winappdebug/README.md) |
| `2` or `trim` | Launches [Trim](modules/trim/README.md) |
| `3` or `transmute` | Launches [Transmute](modules/transmute/readme.md) |
| `4` or `diff` | Launches [Diff](modules/diff/readme.md) |
| `5` or `ccdebug` | Launches [CCiniDebug](modules/ccdebug/readme.md) |
| `6` or `browserbuilder` | Launches [Browser Builder](modules/browserbuilder/readme.md) |
| `7` or `combine` | Launches [Combine](modules/combine/readme.md) |
| `8` or `download` | Launches [Downloader](modules/download/readme.md) |
| `9` or `flavorize` | Launches [Flavorizer](modules/transmute/Flavorizer/readme.md) |
| `11` or `entrybuilder` | Launches [Entry Builder](modules/entrybuilder/readme.md) |
| `12` or `cc7patcher` | Launches [CC7Patcher](modules/cc7patcher/readme.md) |

###### Note: These numbers are not the same as the numbers on the main menu.

### Global Args

| Arg | Effect | Notes |
|:-|:-|:-|
| `-s` | Enables silent mode, muting almost all output and prompts for input | Some exceptions and errors may not be shown when silent mode is enabled |
| `-offline` | Skips the network connection check at startup and runs in offline mode |  |
| `-autoupdate` | Checks for and applies a winapp2ool update before running the requested module | Installs only a newer build whose signature verifies. A failed update exits with code 1 in silent mode |
| `-writelog` | Writes winapp2ool's internal log to `winapp2ool.log` on exit | A run that exits with a nonzero code saves the log whether or not this arg is provided |

### Flavor Args

The active [Flavor](../README.md#what-are-flavors) determines which variant of winapp2.ini is downloaded by any module that downloads one (Downloader, Trim, Diff, CC7Patcher).

| Arg | Effect | Notes |
|:-|:-|:-|
| `-ccleaner` or `-cc` | Sets the Flavor to CCleaner | |
| `-base` or `-ncc` | Sets the Flavor to Base | Default |
| `-bleachbit` or `-bb` | Sets the Flavor to BleachBit | |
| `-systemninja` or `-sn` | Sets the Flavor to System Ninja | |
| `-tron` | Sets the Flavor to Tron | |
| `-ccleaner7` or `-cc7` | Sets the Flavor to CCleaner 7 | |
| `-fluentcleaner` or `-fc` | Sets the Flavor to FluentCleaner | |

### File Selection Args

Every module numbers the files with which it works:

| Arg | Effect | Notes |
|:-|:-|:-|
| `-1d`, `-2d`, ... | Defines a new directory (and optionally file name) for the module's respectively numbered file | Paths with spaces must be provided in quotes, eg. `-1d "C:\New Folder"` |
| `-1f`, `-2f`, ... | Defines a new file name for the module's respectively numbered file | Subdirectories can be given through the file name, eg. `-1f \subdir\winapp2.ini` |

###### Note: In most modules the "first file" (`-1d`/`-1f`) is the winapp2.ini being read and the "third file" (`-3d`/`-3f`) is the output file, but there are exceptions. Refer to a specific module's readme for its file numbering.

---

# Usage Examples

The outputs below were captured from real winapp2ool runs. Version numbers and entry counts are from the day and the machine they were recorded on.

## Example 1: First run (updating and trimming from the menu)

**Context**

You have just downloaded winapp2ool.exe into an empty folder and want a copy of winapp2.ini optimized for your machine.

**Steps**

1. Run winapp2ool.exe.
2. Because there is no local winapp2.ini, the main menu opens with the update options shown in [Quick Start](#quick-start)
3. Choose `Update & Trim`. Winapp2ool prints `Downloading & trimming, this may take a moment...`, downloads the latest winapp2.ini in the active Flavor (base unless you've changed it in [Settings](modules/maintool/readme.md)), removes the entries that do not apply to your system, and saves the result as `winapp2.ini` (`winapp2.rules` for System Ninja) in the current folder.
4. The update options disappear from the menu because your local copy is now current.

---

## Example 2: Scripted download

**Context**

A maintenance script needs the latest winapp2.ini without any user interaction.

**Command**

```
winapp2ool download winapp2 -s
```

**Output**

Nothing is printed to the console. `winapp2.ini` is downloaded from GitHub and saved in the current directory:

```ini
; Version: 251109
; # of entries: 3,715
```

**Explanation**

- `download` launches the Downloader module
- `winapp2` is the Downloader's file selector for winapp2.ini
- `-s` suppresses output and prompts, so the command can run unattended
- No Flavor arg is given, so the active Flavor is downloaded: base, unless you've changed it in [Settings](modules/maintool/readme.md). System Ninja's file is saved as `winapp2.rules`

---

## Example 3: Downloading and trimming in one command

**Context**

The same maintenance script should produce a winapp2.ini that's already trimmed for the machine on which it runs. This is a scripted version of `Update & Trim` from Example 1.

**Command**

```
winapp2ool -2 -d -s
```

**Output**

Nothing is printed to the console. `winapp2.ini` appears in the working directory:

```ini
; Version: 251109
; # of entries: 349
```

**Explanation**

- `-2` launches Trim
- `-d` tells Trim to download the file it trims
- The full database had 3,715 entries, and 349 of them applied to this machine

---

## Example 4: Choosing a Flavor

**Context**

A BleachBit 5 user wants the BleachBit flavor of winapp2.ini instead of the default base one.

**Command**

```
winapp2ool -bleachbit download winapp2 -s
```

**Output**

The BleachBit flavor of `winapp2.ini` is downloaded and saved to the current directory.

**Explanation**

- Flavor args are global, so `-bleachbit` works the same way with Trim, Diff, and CC7Patcher

---

## Example 5: Fully offline maintenance chain

**Context**

You keep your own additions to winapp2.ini in a `custom.ini` file. A script applies them to your local winapp2.ini and then fixes the style and syntax of the result. Your network connection is bad, and winapp2ool lags on startup while it tries to reach GitHub, so you do this offline.

**Files**

###### **Base file (`winapp2.ini`)**

```ini
[My App Logs *]
Section=Custom Entries
DetectFile=%AppData%\MyApp
FileKey1=%AppData%\MyApp\Logs|*.log
```

###### **Source file (`custom.ini`)**

```ini
; Personal additions maintained separately from the official winapp2.ini
[My App Logs *]
FileKey=%AppData%\MyApp\CrashDumps|*.dmp

[My Other App *]
Section=Custom Entries
DetectFile=%AppData%\MyOtherApp
FileKey1=%AppData%\MyOtherApp\Cache|*.tmp
```

**Commands**

```
winapp2ool -transmute -add -2f custom.ini -3f winapp2.ini -s -offline
winapp2ool -debug -c -1f winapp2.ini -3f winapp2.ini -s -offline
```

**Output**

###### **`winapp2.ini` after both commands**

```ini
[My App Logs *]
Section=Custom Entries
DetectFile=%AppData%\MyApp
FileKey1=%AppData%\MyApp\CrashDumps|*.dmp
FileKey2=%AppData%\MyApp\Logs|*.log

[My Other App *]
Section=Custom Entries
DetectFile=%AppData%\MyOtherApp
FileKey1=%AppData%\MyOtherApp\Cache|*.tmp
```

**Explanation**

- The first command adds the new key and the new entry from `custom.ini` to `winapp2.ini`. That's Transmute's Add mode, and `-3f` writes the result back over the base file
- The second command lints the result with WinappDebug, and `-c` saves its corrections. It numbers the `FileKey` that Transmute added and sorts the keys alphabetically
- `-s -offline` on both commands keeps the chain silent and skips the startup network check

---

# Troubleshooting

| Message | Cause |
|:-|:-|
| "Winapp2ool is currently in offline mode" | Winapp2ool couldn't reach GitHub at startup, or was launched with `-offline`. Everything that doesn't download something still works on local files. Choose `Go online` on the main menu to try again |
| "Winapp2ool was unable to establish a network connection. You are still in offline mode." | A `Go online` retry failed |
| "A new version of winapp2.ini is available!" | Not an error. Your local winapp2.ini is older than the latest release. `Current: v000000 (file not found)` means there's no winapp2.ini in winapp2ool's folder |
| "Your .NET Framework is out of date" | The installed .NET Framework is older than the version winapp2ool targets. Winapp2ool can't update itself or switch to the beta until you install a newer one |
| "Winapp2ool is unable to automatically update" | Winapp2ool is running from the `%temp%` folder, or the .NET Framework is out of date |
| "Invalid input. Please try again." | The menu received input it doesn't recognize |
| "Please report this error on GitHub. It will be saved to winapp2ool.log in the same folder as winapp2ool." | Something went wrong that winapp2ool didn't expect. The log it writes next to winapp2ool.exe has the details to include in a [GitHub issue](https://github.com/MoscaDotTo/Winapp2/issues) |

---

# Notes

### General

.NET Framework 4.8 is built into Windows 10 (version 1903 and later) and Windows 11.

By default, every module looks for its files in the same folder as winapp2ool.exe. Menus abbreviate the current directory as `..` in the paths they show.

Winapp2ool runs without administrator rights unless it needs them to save files. If the folder it runs from, or a folder passed with `-1d`, `-2d` and so on, is protected (such as `Program Files\CCleaner`), it restarts itself with administrator rights through a UAC prompt. If a save in the menu is refused, it offers the same restart. Declining the prompt leaves winapp2ool running without administrator rights, and saves to protected folders then fail with a message saying so.

Winapp2ool doesn't back up ini files before it changes them, so keep your own copy of anything you care about.

### Windows XP

Windows XP users should use winapp2oolXP. It can still download and trim the latest winapp2.ini, but it's no longer maintained, and we can't offer support for it.

This readme describes the modern state of Winapp2ool and not Winapp2oolXP, though most basic menus and functions are similar.
