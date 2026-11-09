# Combine

**Combine** is a winapp2ool build-chain module for winapp2.ini maintainers that merges every ini file in a target directory (subdirectories included) into one output file. It's the assembly stage of the winapp2.ini build, but it works on any collection of ini files. Most end users won't need it. The exception is merging a folder of your own entries into one `custom.ini` for [Transmute](../transmute/readme.md) (see [Example 7](EXAMPLES.md#example-7-building-a-customini-for-transmute)).

### What does Combine do?

Combine finds every `.ini` file under the target directory, reads them in order, and writes their sections into a single output file. When two files define a section with the same name, Combine reports the collision and merges the two sections' keys, leaving out any key that exactly repeats one already there. With [strict name checking](#strict-name-checking) on, a collision fails the run instead and nothing is written.

### Why Combine?

The winapp2.ini build generates its entries into separate artifact files, one per letter for the base entries plus one each for the browser and UWP generators, and Combine is the step that folds them back into one database. It reports every section name that more than one file defines, naming the file that defined it first and each file that defined it again. In strict mode those collisions fail the build, which is what the pipeline wants, since its inputs are supposed to be disjoint.

Outside the build, it lets you keep a large ini collection split into files small enough to work in, and join them whenever something wants a single file.

---

# Table of Contents

**Using Combine**

1. [Requirements](#requirements)
2. [Quick Start](#quick-start)
3. [Common Uses](#common-uses)
4. [Menu Options](#menu-options)
5. [Command-Line Arguments](#command-line-arguments)
   - [File Selection](#file-selection)
   - [Toggles](#toggles)
6. [How Combining Works](#how-combining-works)
   - [File Discovery](#file-discovery)
   - [Duplicate Section Handling](#duplicate-section-handling)
   - [Cross-File Name Collisions](#cross-file-name-collisions)
   - [Strict Name Checking](#strict-name-checking)
   - [Output Formatting](#output-formatting)
   - [Run Summary and Log](#run-summary-and-log)
7. [Tips & Best Practices](#tips--best-practices)
8. [Troubleshooting](#troubleshooting)

**Examples**

9. **[Examples](EXAMPLES.md)**: ten worked examples with captured input and output

---

# Requirements

- A target directory containing one or more `.ini` files

---

# Quick Start

### Common Workflow

1. Open Combine from the winapp2ool main menu (option 7)
2. Use **Change the target directory** to pick the directory holding your ini files
3. If you want the output somewhere other than `combined.ini` in the current directory, use **Change the save file location**
4. Choose **Run**. Every `.ini` file in the directory and its subdirectories is merged into the output file

Or from the command line:

```
winapp2ool -combine -1d "C:\path\to\your\files"
```

---

# Common Uses

| To | Run | Example |
|:-|:-|:-|
| Combine every ini file in the current directory into `combined.ini` | `winapp2ool -combine` | |
| Combine a folder of ini files into `combined.ini` in the current directory | `winapp2ool -combine -1d "C:\ini\source"` | |
| Combine a folder and save the result under another name in another folder | `winapp2ool -combine -1d "C:\ini\source" -3d "C:\ini\output" -3f merged.ini` | [Combining Files with Distinct Sections](EXAMPLES.md#example-1-combining-files-with-distinct-sections) |
| Merge a folder of your own entries into a `custom.ini` for Transmute | `winapp2ool -combine -1d C:\ini\case7\customs -3d C:\ini\case7 -3f custom.ini` | [Building a custom.ini for Transmute](EXAMPLES.md#example-7-building-a-customini-for-transmute) |
| Fail the run with exit code 1 when two files define the same section | `winapp2ool -combine -strict -1d "C:\ini\source" -3f merged.ini` | [A Strict Run That Fails the Build](EXAMPLES.md#example-6-a-strict-run-that-fails-the-build) |
| Run the winapp2.ini build's merge stage | `winapp2ool -s -offline -combine -strict -1d \Entries -3f Winapp2.ini` | [Merging the winapp2.ini Build Artifacts](EXAMPLES.md#example-9-merging-the-winapp2ini-build-artifacts) |

In the menu, the same choices are **Change the target directory**, **Change the save file location** and **Toggle Strict Mode**, set before **Run**; see [Menu Options](#menu-options). `-s` and `-offline` are winapp2ool [global args](../../Readme.md#global-args).

---

# Menu Options

| Option | Effect | Notes |
|:-|:-|:-|
| Run (default) | Merge all ini files in the target directory into the output file | |
| Change the target directory | Pick the directory to scan for ini files | Default: the current directory |
| Change the save file location | Pick the output file's folder and name | Default: `combined.ini` in the current directory |
| Toggle Strict Mode | Turn [strict name checking](#strict-name-checking) on or off | When on, a section name found in more than one input file fails the run instead of merging. Default: off |
| Log Viewer | Show the detailed log from the last Combine run | Only appears once Combine has run in the current session |
| Reset Settings | Restore all settings to their defaults | Only appears when a setting has been changed |

---

# Command-Line Arguments

Combine takes a target directory and an output file location. It is CLI module **7**, called as `-combine` or `-7`.

Command line runs always start from Combine's default settings, whatever you've saved in the menu.

## File Selection

| Arg | Effect | Default |
|:-|:-|:-|
| `-1d path` | Set the target directory | Current directory |
| `-3d path` | Set the output directory | Current directory |
| `-3f name` | Set the output file name | `combined.ini` |

###### Note: A leading `\` (eg. `-1d \Entries`) names a subfolder of the current directory.

## Toggles

| Arg | Effect |
|:-|:-|
| `-strict` | Turn on [strict name checking](#strict-name-checking): when a section name appears in more than one input file, report every collision and exit with code 1 instead of saving the combined output |

---

# How Combining Works

## File Discovery

Combine searches the target directory and all of its subdirectories for files with the `.ini` extension, then sorts them alphabetically by full path. Because the whole path is compared, a file in a subdirectory can sort ahead of a file in the target directory itself (see [Example 2](EXAMPLES.md#example-2-subdirectories-and-repeated-runs)). Messages name each file by its path below the target directory (`A.ini`, `UWP\A.ini`), so two files with the same name in different folders are never confused.

If the output file lives inside the target directory, Combine skips it. It still counts toward `Found N files`, but it contributes nothing and isn't counted in the run statistics.

A file with no sections is skipped too. Only the [log](#run-summary-and-log) mentions it; the on-screen summary doesn't. A file that can't be read gets a warning and fails the run, but the other files are still combined and the output is still saved (see [Example 10](EXAMPLES.md#example-10-unreadable-and-empty-inputs)).

## Duplicate Section Handling

When two or more files contain a section with the same name, Combine merges their keys into the first copy:

- A key is skipped if the section already has a key with the same name and the same value, ignoring case
- A key with the same name but a different value isn't a duplicate, so both are kept
- Nothing is renumbered or reordered

The merged section ends up with every distinct key from every file that defined it. See [Example 3](EXAMPLES.md#example-3-merging-two-copies-of-the-same-entry).

## Cross-File Name Collisions

Whenever a section name appears in more than one input file, Combine lists it on screen and in the log:

```
 ║ 2 section names appeared in more than one input file:                                                              ║
 ║   [Anki Media Cache *] first defined in A.ini, contributed again by Browsers\browsers.ini, UWP\uwp.ini             ║
 ║   [Audacity Temp Files *] first defined in A.ini, contributed again by UWP\uwp.ini                                 ║
```

Names are compared ignoring case, so `[Foo *]` and `[foo *]` collide. Without strict checking this report is informational. Whether a collision is a problem depends on whether your files were meant to overlap.

## Strict Name Checking

Strict name checking (**Toggle Strict Mode** in the menu, `-strict` on the command line, off by default) turns name collisions into an error:

- Every collision is listed, as above
- `Strict name checking is enabled - {file} will not be saved` follows the list
- Combine doesn't write the output file
- The exit code is set to **1**. A silent run exits with it and writes the global log to `winapp2ool.log`

Strict mode doesn't stop at the first collision, so one run shows you all of them.

## Output Formatting

The output is a plain ini file:

- Sections appear in the order they were first seen across the sorted input files
- Keys in a merged section appear in the order they were added
- Comments in the input files aren't carried over

If the output has no sections, Combine doesn't write it.

## Run Summary and Log

Each run prints a summary:

```
 ╔════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╗
 ║                                        Combining files from C:\ini\case1\customs                                   ║
 ╠════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╣
 ║ Found 2 files with ini extension in C:\ini\case1\customs                                                           ║
 ║ Processed: C.ini (2 sections)                                                                                      ║
 ║ Processed: O.ini (1 section)                                                                                       ║
 ║                                                                                                                    ║
 ║                                    Combined 2 files into custom.ini with 3 sections                                ║
 ╚════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╝
 ╔════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╗
 ║                                                  Combination complete!                                             ║
 ╚════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╝
Press any key to return to the menu.
```

The **Log Viewer** menu option shows the run's part of the winapp2ool log, which has detail the summary leaves out:

| Log message | Meaning |
|:-|:-|
| `Added new section: [name] (N keys)` | A section was seen for the first time and copied into the output |
| `Added {key} to [{section}]` | A key was merged into an existing section of the same name |
| `Skipped duplicate key in [{section}]: {key}` | A key matched an existing key's name and value and was left out |
| `Skipping file with no sections: {file}` | A file was found but had no sections |
| `Output file found in target directory, skipping: {path}` | The output file lives inside the target directory and was left out |
| `N section names appeared in more than one input file:` | Followed by one line per colliding name |
| `Strict name checking is enabled - {file} will not be saved` | A strict run found at least one collision |
| `Processed N files, M contained combinable sections` | Run statistics. `N` leaves out the skipped output file; `M` leaves out files with no sections |
| `Saving {file}` | The output was written to disk |

---

# Tips & Best Practices

### Output File Location

If the output file is inside the target directory, Combine skips it, so running the same command again is safe ([Example 2](EXAMPLES.md#example-2-subdirectories-and-repeated-runs)). Keeping the output outside the target directory is still cleaner, since then there's nothing to skip.

### Alphabetical Processing Order

Files are processed in alphabetical order by full path. When two files define the same section and you want one file's keys to come first, name or place it so it sorts first. Later files only add the keys the section doesn't already have. See [Example 4](EXAMPLES.md#example-4-controlling-key-order-with-file-names).

### winapp2.ini Outputs

Combine doesn't renumber keys or apply winapp2.ini styling. If the result is meant to be a winapp2.ini file, run WinappDebug on it afterward (`winapp2ool -debug -c -1f yourfile.ini -3f yourfile.ini`) to fix numbering, key order and style. See [Example 8](EXAMPLES.md#example-8-correcting-syntax-after-combining).

### Checking the Result

After a run, the Log Viewer shows which files were processed, which sections and keys each one contributed, and which duplicate keys were skipped.

---

# Troubleshooting

| Message / Symptom | Cause |
|:-|:-|
| "Target directory not found. Please select a valid directory." | The target directory doesn't exist. Shown in red above the menu; the run does nothing |
| "N section names appeared in more than one input file:" | Normal output. It only matters if your inputs were supposed to be disjoint |
| "Strict name checking is enabled - {file} will not be saved" | A strict run found a collision. Nothing was written and the exit code is 1 |
| "No valid sections found to combine - {file} will not be saved" | No file in the target directory had a section Combine could read, or there were no `.ini` files at all (`Found 0 files...`), so there's no output file |
| "Error processing file: {path}" followed by "Check the winapp2ool log for more information: {path}" | Combine couldn't read that file, often because another program had it open. The other files are still combined and the output is saved, but the exit code is 1, and the global log is saved to disk straight away |
| "{file} was not saved" | Combine couldn't write the output file. Check that the folder exists and that you can write to it. In the menu, winapp2ool offers to restart with administrator rights when the folder is protected |
| A scripted run fails with exit code 1 and no obvious error | Silent mode prints nothing, so the cause is in `winapp2ool.log`, which is written whenever a silent run exits nonzero. Usually it's a strict-mode collision; an unreadable input or an output that couldn't be written does the same |
