# Downloader

**Downloader** is a winapp2ool module for end users that fetches files from the winapp2 GitHub repository (this one!). It gets you any flavor of winapp2.ini, the latest winapp2ool, or one of the project's supplementary files, from the menu or from a script.

### What does Downloader do?

Downloader resolves a short name (`winapp2`, `readme`, `winapp3`, etc) to a URL on the winapp2 GitHub, downloads the file, and writes it to a directory you choose. For winapp2.ini specifically, the URL it picks depends on the active flavor in winapp2ool.

### Why Downloader?

There are several flavors of winapp2.ini, and Downloader knows where each one lives, so you don't have to hunt through the repository for the right file. It is also how winapp2ool updates itself, keeping a backup of the version it replaced. And it fetches `winapp3.ini` and `Archived entries.ini`, which you can fold into a winapp2.ini with Transmute.

---

# Table of Contents

**Using Downloader**

1. [Requirements](#requirements)
2. [Quick Start](#quick-start)
3. [Common Uses](#common-uses)
4. [Menu Options](#menu-options)
   - [Main Menu](#main-menu)
   - [Advanced Downloads](#advanced-downloads)
5. [Command-Line Arguments](#command-line-arguments)
   - [File Selection](#file-selection)
   - [Path Parameters](#path-parameters)
   - [Flavor Flags](#flavor-flags)
   - [Global Flags](#global-flags)
6. [How Downloads Resolve](#how-downloads-resolve)
   - [Flavor Selection](#flavor-selection)
   - [Output File Names](#output-file-names)
7. [Download Behavior](#download-behavior)
   - [The Save Directory](#the-save-directory)
   - [Existing Files](#existing-files)
   - [Updating Winapp2ool](#updating-winapp2ool)
8. [Tips & Best Practices](#tips--best-practices)
9. [Troubleshooting](#troubleshooting)

**For Contributors**

10. [Where Files Come From](#where-files-come-from)

**Examples**

11. **[Examples](EXAMPLES.md)**: nine worked examples with captured input and output

---

# Requirements

- An active internet connection. Downloader is hidden from the main menu in offline mode
- .NET Framework 4.8 to download `winapp2ool.exe`. Every other file downloads on any .NET that runs winapp2ool

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

# Common Uses

| To | Run | Example |
|:-|:-|:-|
| Download winapp2.ini in the active flavor to the current directory | `winapp2ool -download winapp2` | [Downloading winapp2.ini](EXAMPLES.md#example-1-downloading-winapp2ini) |
| Download one flavor without changing the active flavor | `winapp2ool -download winapp2 -cc` | [Choosing a flavor without changing your settings](EXAMPLES.md#example-2-choosing-a-flavor-without-changing-your-settings) |
| Download into another directory, created if it doesn't exist | `winapp2ool -download 1 -cc7 -1d "C:\Tools"` | [Downloading into a directory that doesn't exist yet](EXAMPLES.md#example-3-downloading-into-a-directory-that-doesnt-exist-yet) |
| Save the download under a different name | `winapp2ool -download winapp2 -1f winapp2-backup.ini` | [Renaming the download](EXAMPLES.md#example-4-renaming-the-download) |
| Download from a script, with no output or prompts | `winapp2ool -download winapp2 -sn -s` | [Silent, scriptable downloads](EXAMPLES.md#example-5-silent-scriptable-downloads) |
| Update winapp2ool to the latest release | `winapp2ool -download winapp2ool` | [Updating winapp2ool in place](EXAMPLES.md#example-7-updating-winapp2ool-in-place) |
| Download winapp2ool.exe without updating the copy you're running | `winapp2ool -download winapp2ool -1d "C:\Tools"` | [Downloading winapp2ool.exe without updating](EXAMPLES.md#example-8-downloading-winapp2oolexe-without-updating) |

In the menu, each file is its own option and **Change Save Directory** sets where it is saved; see [Menu Options](#menu-options). `-s` is described under [Global Flags](#global-flags).

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
| Winapp2ool | Download the latest winapp2ool.exe | Requires .NET 4.8. Replaces the running executable when the save directory is the current directory. See [Updating Winapp2ool](#updating-winapp2ool) |
| ReadMe | Download the top-level winapp2ool readme | Saved as `readme.md` |
| Advanced | Open the Advanced Downloads sub-menu | See [Advanced Downloads](#advanced-downloads) |
| Change Save Directory | Select a new directory for all downloaded files | Default: current directory |
| Reset Settings | Restore the save directory to its default | Only shown when the save directory has been changed |

The seven winapp2.ini options always download the flavor their label names and ignore the active flavor setting. Only the command line goes by the active flavor.

## Advanced Downloads

| Option | Effect | Notes |
|:-|:-|:-|
| Winapp3.ini | Download the extended/potentially unsafe entry database | Not recommended for general use |
| Archived entries.ini | Download entries for old or discontinued software | Not recommended for general use |

---

# Command-Line Arguments

Downloader is command-line module **8**, so `winapp2ool -download` and `winapp2ool 8` both work.

The module takes one argument naming the file to download, plus optional path parameters. The file argument is required. If it is missing or isn't one of the names below, Downloader lists the valid arguments and exits with code 1.

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

A flavor flag applies to the run it's on and never changes the saved flavor.

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
| `-s` | Suppress all output, skip the overwrite prompt, and exit when finished. Any script that runs unattended needs it |
| `-writelog` | Write `winapp2ool.log` to disk on exit |

###### Note: without `-s`, winapp2ool does not exit after the download.

---

# How Downloads Resolve

## Flavor Selection

The flavor only matters for `winapp2`.

The active flavor comes from winapp2ool's global settings (**Settings → Change Flavor** on the main menu), but any command-line [flavor flag](#flavor-flags) overrides it for that run. Command-line runs never save settings.

###### Note: the default flavor is base. On a winapp2ool that has never had its flavor changed, `winapp2ool -download winapp2` fetches the base flavor. Pass `-cc` (or `-ccleaner`) for the CCleaner flavor.

## Output File Names

| Request | Active flavor | File written |
|:-|:-|:-|
| `winapp2` | Base *(default)* | `winapp2.ini` |
| `winapp2` | CCleaner | `winapp2.ini` |
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

---

# Download Behavior

## The Save Directory

The save directory defaults to the directory from which winapp2ool runs. If the save directory doesn't exist, Downloader creates it, along with any missing parent directories.

## Existing Files

When the target file already exists and output is not suppressed, Downloader prompts before writing:

```
readme.txt already exists in the target directory.
Enter a new file name, or leave blank to overwrite the existing file:
```

Entering a name saves under that name; leaving it blank overwrites. This prompt appears on the command line too, where it will block a script waiting for input. Passing `-s` skips the prompt entirely and overwrites silently.

## Updating Winapp2ool

Downloading `winapp2ool.exe` into the directory winapp2ool is running from doesn't just save a file. Winapp2ool updates itself instead:

1. Downloads the new executable and its signature file, `winapp2ool.exe.sig`, without writing either to disk
2. Checks the signature against the public keys built into winapp2ool, and discards the download if it doesn't match
3. Writes the new executable beside the running one as `winapp2ool.exe.new`, then confirms it is winapp2ool and a newer version than the one running
4. Renames the running executable to `winapp2ool v<version>.exe.bak`, puts the new one in its place, launches it with the same arguments, and exits

If any step fails, the running executable stays where it is and winapp2ool says why. In silent mode a failed update exits with code 1.

An update never installs an older version, even one with a valid signature. Since winapp2ool 1.8, that means turning off **Beta Participation** won't roll a beta build back to an older release.

---

# Tips & Best Practices

### Scripts

- Always pass `-s`. It suppresses output, skips the overwrite prompt that would otherwise block on input, and makes winapp2ool exit instead of opening its menu
- The exit code is `1` when the file argument is missing or invalid or the download fails, and `0` otherwise
- Name the flavor explicitly on every call

### Self-Updating

- Downloading `winapp2ool.exe` into winapp2ool's own directory replaces the running executable and restarts it, leaving a `winapp2ool v<version>.exe.bak` behind. It only does so when the download's signature verifies and its version is newer
- The `.bak` files pile up, one per version replaced. Winapp2ool never deletes these.
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
| `This option requires a newer version of the .NET Framework` | Downloading the executable needs .NET 4.8 |
| `Unable to download winapp2ool to the current directory, choose another directory before trying again` | Winapp2ool is running from the Windows temp folder, or .NET is out of date. Self-updating is disabled in both cases |

| Symptom | Cause |
|:-|:-|
| Downloader is missing from the main menu | Winapp2ool is in offline mode |
| A script hangs after starting a download | The target file already exists and the overwrite prompt is waiting on input. Add `-s` |
| A script never returns, with no prompt | Without `-s`, winapp2ool opens its interactive menu after downloading. Add `-s` |

---

# Where Files Come From

This section is for people who maintain the winapp2 repository, or who want to know which file in it each download comes from. You don't need it to use Downloader.

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
