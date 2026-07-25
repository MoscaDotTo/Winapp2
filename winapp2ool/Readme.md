# Winapp2ool

Winapp2ool is the tool that builds winapp2.ini, and a companion utility for the people who use it. The published winapp2.ini files (the base file, each of its [flavors](../README.md#what-are-flavors), and the changelog shipped alongside each one) are generated from source by winapp2ool rather than edited by hand. It is a menu-driven console application for end users, and can be driven run from a scripting environment for automation.

### What is winapp2.ini?

Winapp2.ini is a massive, community-driven database of declarative cleaning routines for Microsoft Windows, compatible with CCleaner, BleachBit, System Ninja, R-Wipe&Clean, HDCleaner, and FluentCleaner. Winapp2.ini has its own readme [here](../README.md).

### Why winapp2ool?

Winapp2ool was created to help automate otherwise complex or time consuming tasks for maintainers, contributors, and end users of Winapp2.ini. It provides several tools:

* [WinappDebug](modules/winappdebug/README.md): Performs static analysis and corrects style and syntax errors
* [Trim](modules/trim/README.md): Reduces the database to entries relevant to the current system
* [Transmute](modules/transmute/readme.md): Applies individual structured patches to ini files
* [Flavorizer](modules/transmute/Flavorizer/readme.md): Applies batches of structured patches to ini files
* [Diff](modules/diff/readme.md): Generates changelogs between any two Winapp2.ini versions, tracking additions, removals, renames, mergers, and key movements between entries using context-aware abstraction tracking
* [CCiniDebug](modules/ccdebug/readme.md): Removes stale Winapp2.ini configurations from the CCleaner settings
* [Entry Lab](modules/entrylab/readme.md): Hub for the entry generators: [Browser Builder](modules/browserbuilder/readme.md), [UWP Builder](modules/uwpbuilder/readme.md), and [Entry Builder](modules/entrybuilder/readme.md)
* [Combine](modules/combine/readme.md): Merges folders (including subfolders) of ini files into a single file
* [CC7Patcher](modules/cc7patcher/readme.md): Installs winapp2.ini entries into CCleaner 7's ccleaner.ini
* [Downloader](modules/download/readme.md): Downloads winapp2.ini and related files from GitHub

### How winapp2.ini gets built

The published winapp2.ini is not assembled by hand. Its sources live in [Assembler](../Assembler), and winapp2ool performs every stage that turns them into the published files. 

1. [Entry Builder](modules/entrybuilder/readme.md), [Browser Builder](modules/browserbuilder/readme.md), and [UWP Builder](modules/uwpbuilder/readme.md) expand the [base entry](../Assembler/EntryBuilder), [browser](../Assembler/BrowserBuilder), and [Microsoft Store app](../Assembler/UWP) sources into the committed [build artifacts](../Assembler/Entries)
2. [Combine](modules/combine/readme.md) joins those artifacts into one file in strict mode
3. [WinappDebug](modules/winappdebug/README.md) applies the style and syntax rules to the merged file and saves its corrections
4. [Flavorizer](modules/transmute/Flavorizer/readme.md) applies each flavor's ruleset to produce the CCleaner, CCleaner 7, BleachBit, FluentCleaner, Tron, and System Ninja variants
5. [Diff](modules/diff/readme.md) creates the changelog for the base file and for each flavor against the previously version

Because the artifacts under [Entries](../Assembler/Entries) are committed to the repository rather than existing only transiently during a build, the build can be checked against itself: `build winapp2.ps1 -Verify` regenerates all of them, byte-compares against what is committed, and exits nonzero on any drift. That check runs in GitHub Actions, alongside:

* Pull request verification: A pull request touching the build sources is built twice: master alone, then master with the pull request merged. A bot comment reports either the changelog the change produces or the stage at which it broke the build
* Scheduled builds: Sources are rebuilt daily. When the output changes, it is committed back and a release is cut with every flavor and changelog attached

---

# Table of Contents

1. [Requirements](#requirements)
2. [Installation](#installation)
3. [Quick Start](#quick-start)
4. [Menu Options](#menu-options)
5. [Command-Line Arguments](#command-line-arguments)
   - [Module Args](#module-args)
   - [Global Args](#global-args)
   - [Flavor Args](#flavor-args)
   - [File Selection Args](#file-selection-args)
6. [Usage Examples](#usage-examples)
   - [Example 1: First run — updating and trimming from the menu](#example-1-first-run--updating-and-trimming-from-the-menu)
   - [Example 2: Scripted download](#example-2-scripted-download)
   - [Example 3: Downloading and trimming in one command](#example-3-downloading-and-trimming-in-one-command)
   - [Example 4: Choosing a Flavor](#example-4-choosing-a-flavor)
   - [Example 5: Fully offline maintenance chain](#example-5-fully-offline-maintenance-chain)
7. [Troubleshooting](#troubleshooting)
8. [Notes](#notes)

---

# Requirements

### Minimum

* Windows Vista SP2
* .NET Framework 4.5
* Administrative permissions (see [Notes](#notes))

### Suggested

* Windows 7 or higher
* .NET Framework 4.6 or higher (for automatically updating the executable)
* Network connection for full functionality

---

# Installation

Download the latest winapp2ool.exe from the [Releases page](https://github.com/MoscaDotTo/Winapp2/releases/) or from the [Release directory](https://github.com/MoscaDotTo/Winapp2/tree/master/winapp2ool/bin/Release) and place it in the folder where you keep (or would like to keep) winapp2.ini.

### Updates

Winapp2ool will prompt you to update it from within the application whenever an update is available. The application will create a backup of itself with the `.bak` extension when doing this. Rename to `.exe` to restore the backed up version.

### Beta builds

Beta builds are occasionally made available to the public while features are in development. To access the beta build, open the [winapp2ool settings](modules/maintool/readme.md) and enable beta participation. The newest beta build will automatically be downloaded and launched. Beta participation requires .NET Framework 4.6 or higher. Beta builds live on [Branch1](https://github.com/MoscaDotTo/Winapp2/tree/Branch1).

---

# Quick Start

Place winapp2ool.exe in an empty folder and run it. On first launch with a network connection, winapp2ool detects that you have no local copy of winapp2.ini and presents download options directly on the main menu:

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

###### Note: This menu was captured from a real first run in an empty folder. `Current: v000000 (file not found)` reflects that no local winapp2.ini exists yet; once you have one, its version number is shown instead. The update options at the bottom only appear while an update is available.

From here, the most common tasks are one keypress away:

1. **Download the latest winapp2.ini:** Choose `Update Winapp2.ini`
2. **Optimize the latest winapp2.ini for your system:** Choose `Update & Trim`
3. **See the changelog between your local version and the latest:** Choose `Show winapp2.ini changelog`
4. **Update winapp2ool itself:** Choose `Update Winapp2ool` (shown when a new version of the tool is available)

---

# Menu Options

Click a linked option name to see the readme for that module.

| Option | Effect | Notes |
|:-|:-|:-|
| Exit | Exits the application | |
| [WinappDebug](modules/winappdebug/README.md) | Scan for and correct style and syntax errors in winapp2.ini | Enforces the winapp2.ini style and syntax guidelines |
| [Trim](modules/trim/README.md) | Optimize winapp2.ini for your system | Removes entries not relevant to your machine, reducing load times in tools like CCleaner |
| [Transmute](modules/transmute/readme.md) | Add, replace, or remove entire sections or individual keys from winapp2.ini | Patches ini files at whole-section or individual-key granularity. [Flavorizer](modules/transmute/Flavorizer/readme.md), which applies these patches in batches, is opened from this menu |
| [Diff](modules/diff/readme.md) | Generate a context-aware changelog between two winapp2.ini files | |
| [CCiniDebug](modules/ccdebug/readme.md) | Remove stale winapp2.ini configurations from ccleaner.ini | For CCleaner 6 and earlier only |
| [Entry Lab](modules/entrylab/readme.md) | Generate winapp2.ini entries from templates | Hub for [Browser Builder](modules/browserbuilder/readme.md), [UWP Builder](modules/uwpbuilder/readme.md), and [Entry Builder](modules/entrybuilder/readme.md) |
| [Combine](modules/combine/readme.md) | Join together a collection of ini files into one | Includes subfolders |
| [CC7Patcher](modules/cc7patcher/readme.md) | Install winapp2.ini for CCleaner 7 | Appends winapp2.ini entries into CCleaner 7's ccleaner.ini |
| [Downloader](modules/download/readme.md) | Download files from the Winapp2 GitHub | **Unavailable in offline mode** |
| [Settings](modules/maintool/readme.md) | Manage Winapp2ool's settings | Application-level settings: saving settings to disk, beta participation, offline mode, the active [Flavor](#flavor-args), and log management |
| Update Winapp2.ini | Update your local copy of winapp2.ini | **Only shown while a winapp2.ini update is available; unavailable in offline mode** |
| Update & Trim | Download and trim the latest winapp2.ini | **Only shown while a winapp2.ini update is available; unavailable in offline mode** |
| Show winapp2.ini changelog | See the difference between your local file and the latest | **Only shown while a winapp2.ini update is available; unavailable in offline mode** |
| Show trimmed changelog | See the difference between your trimmed local file and the latest | **Only shown while a winapp2.ini update is available; unavailable in offline mode** |
| Update Winapp2ool | Get the latest Winapp2ool.exe | **Only shown while a winapp2ool update is available. Unavailable in offline mode and on machines with .NET Framework 4.5 or lower (ie. Winapp2oolXP)** |
| Go online | Retry your internet connection | **Only available in offline mode** |

###### Note: The main menu also accepts the hidden commands `printlog` (print winapp2ool's internal log to the console) and `savelog` (write the log to disk). See [Log Management](modules/maintool/readme.md#log-management) in the settings readme.

---

# Command-Line Arguments

Winapp2ool supports command line arguments ("args"). These allow Winapp2ool to be used from a scripting environment (such as a shell script) without having to interact with the UI. There are several top level args which apply settings globally, and then there are tool specific args which are defined in each tool's own readme.

The first argument provided should always refer to the module you would like to use, as below. Modules can be selected by number or by name, with or without a leading `-` — `1`, `-1`, `debug`, and `-debug` are all equivalent.

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
| `10` or `uwpbuilder` | Launches [UWP Builder](modules/uwpbuilder/readme.md) |
| `11` or `entrybuilder` | Launches [Entry Builder](modules/entrybuilder/readme.md) |
| `12` or `cc7patcher` | Launches [CC7Patcher](modules/cc7patcher/readme.md) |

###### Note: These numbers are not the same as the numbers on the main menu. 

### Global Args

| Arg | Effect | Notes |
|:-|:-|:-|
| `-s` | Enables silent mode, muting almost all output and prompts for input | Some exceptions and errors may not be shown when silent mode is enabled |
| `-offline` | Skips the network connection check at startup and runs in offline mode |  |
| `-autoupdate` | Checks for and applies a winapp2ool update before running the requested module | Requires .NET Framework 4.6 or higher |
| `-writelog` | Writes winapp2ool's internal log to `winapp2ool.log` on exit | A run that exits with a nonzero code saves the log whether or not this arg is provided |

### Flavor Args

The active [Flavor](../README.md#what-are-flavors) determines which variant of winapp2.ini is downloaded by any module that downloads one (Downloader, Trim, Diff, CC7Patcher).

| Arg | Effect | Notes |
|:-|:-|:-|
| `-ccleaner` or `-cc` | Sets the Flavor to CCleaner | Default |
| `-ncc` or `-base` | Sets the Flavor to Non-CCleaner (base) | |
| `-bleachbit` or `-bb` | Sets the Flavor to BleachBit | |
| `-systemninja` or `-sn` | Sets the Flavor to System Ninja | |
| `-tron` | Sets the Flavor to Tron | |
| `-ccleaner7` or `-cc7` | Sets the Flavor to CCleaner 7 | |
| `-fluentcleaner` or `-fc` | Sets the Flavor to FluentCleaner | |

### File Selection Args

Every module numbers the files it works with. Two argument types override where a numbered file lives:

| Arg | Effect | Notes |
|:-|:-|:-|
| `-1d`, `-2d`, ... | Defines a new directory (and optionally file name) for the module's respectively numbered file | Paths with spaces must be provided in quotes, eg. `-1d "C:\New Folder"` |
| `-1f`, `-2f`, ... | Defines a new file name for the module's respectively numbered file | Subdirectories can be given through the file name, eg. `-1f \subdir\winapp2.ini` |

###### Note: In most modules the "first file" (`-1d`/`-1f`) is the winapp2.ini being read and the "third file" (`-3d`/`-3f`) is the output file, but there are exceptions. Refer to a specific module's readme for its file numbering.

---

# Usage Examples

The outputs below were captured from real winapp2ool runs. Winapp2.ini version numbers and entry counts reflect the day and the machine on which the examples were recorded. 

## Example 1: First run (updating and trimming from the menu)

**Context**

You have just downloaded winapp2ool.exe into an empty folder and want a copy of winapp2.ini optimized for your machine.

**Steps**

1. Run winapp2ool.exe.
2. Because there is no local winapp2.ini, the main menu opens with the update options shown in [Quick Start](#quick-start)
3. Choose `Update & Trim`. Winapp2ool prints `Downloading & trimming, this may take a moment...`, downloads the latest CCleaner flavor winapp2.ini, removes the entries that do not apply to your system, and saves the result as `winapp2.ini` in the current folder.
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

Nothing is printed to the console. `winapp2.ini` is downloaded from github and saved in the current directory

```ini
; Version: 251109
; # of entries: 3,715
```

**Explanation**

- `download` launches the Downloader module
- `winapp2` is the Downloader's file selector for winapp2.ini
- `-s` suppresses all output and prompts, allowing the command to run unattended
- No Flavor arg is given, so the default CCleaner Flavor is downloaded

---

## Example 3: Downloading and trimming in one command

**Context**

The same maintenance script should instead produce a winapp2.ini already optimized for the machine it runs on. This is the scripted equivalent of Example 1's `Update & Trim`.

**Command**

```
winapp2ool -2 -d -s
```

**Output**

Nothing is printed to the console. `winapp2.ini` appears in the working directory. 

```ini
; Version: 251109
; # of entries: 349
```

**Explanation**

- `-2` launches Trim 
- `-d` is Trim's flag for downloading the file to trim 
- The full database contained 3,715 entries; 349 survived trimming on this machine. 

---

## Example 4: Choosing a Flavor

**Context**

A BleachBit user wants the BleachBit variant of winapp2.ini instead of the default CCleaner one.

**Command**

```
winapp2ool -bleachbit download winapp2 -s
```

**Output**

The BleachBit flavor of `winapp2.ini` is downloaded and saved to the current directory 

**Explanation**

- Flavor args are global: they affect any module that downloads winapp2.ini

---

## Example 5: Fully offline maintenance chain

**Context**

You maintain personal additions to winapp2.ini in a `custom.ini` file. A script applies them to your local winapp2.ini and then corrects the style and syntax of the result. Your network connection is bad, and you notice that winapp2ool lags on startup while it tries to connect to GitHub. 

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

- The first command adds the new key and the new entry from `custom.ini` into `winapp2.ini` (Transmute Add mode, writing back over the base file via `-3f`)
- The second command lints the result with WinappDebug (`-c` saves its corrections): the unnumbered `FileKey` added by Transmute has been renumbered and the keys alphabetized
- `-s -offline` on both commands makes the chain silent and skips the startup network check. 

---

# Troubleshooting

| Message | Cause |
|:-|:-|
| "Winapp2ool is currently in offline mode" | Winapp2ool could not reach GitHub at startup (or was launched with `-offline`). All functions not directly related to downloading still work on local files. Use `Go online` from the main menu to retry the connection |
| "Winapp2ool was unable to establish a network connection. You are still in offline mode." | A `Go online` retry failed |
| "A new version of winapp2.ini is available!" | Informational: your local winapp2.ini is older than the latest release. `Current: v000000 (file not found)` means no winapp2.ini exists in winapp2ool's folder |
| "Your .NET Framework is out of date" | The machine has .NET Framework 4.5 or lower.  |
| "Winapp2ool is unable to automatically update" | Winapp2ool was launched from the `%tmp%` directory or the .NET Framework is out of date |
| "Invalid input. Please try again." | The menu received input it doesn't recognize |
| "Please report this error on GitHub. It will be saved to winapp2ool.log in the same folder as winapp2ool." | An unexpected error occurred; the log written next to winapp2ool.exe contains the details to include in a [GitHub issue](https://github.com/MoscaDotTo/Winapp2/issues) |

---

# Notes

### General

.NET Framework 4.5 (or newer) comes pre-installed by default on Windows 8 and newer.

By default, each tool in the application assumes that local files it is looking for are in the same folder as the executable. File paths displayed in menus abbreviate the current directory as `..`

Winapp2ool performs queries against protected system areas such as the Program Files and Windows directories and may return invalid results if run without administrative permissions.

Winapp2ool does not perform any automatic backup of ini files before modifying them. 

### Windows XP

Windows XP users should use winapp2oolXP. Winapp2oolXP is no longer maintained and no longer receives updates, but retains the ability to download and trim the latest winapp2.ini for users on that platform. We no longer able to provide application support or updates for users of Winapp2oolXP.
