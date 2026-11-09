# Combine

**Combine** is a winapp2ool build-chain module for winapp2.ini maintainers that merges all ini files from a target directory (and all of its subdirectories) into a single output file. It is the winapp2.ini build pipeline's assembly stage, but it works standalone on any collection of ini files. Most end users won't need it. The exception is merging a folder of your own entries into one `custom.ini` for [Transmute](../transmute/readme.md) (see [Example 7](EXAMPLES.md#example-7-building-a-customini-for-transmute)).

### What does Combine do?

Combine recursively scans a target directory for files with the `.ini` extension, reads each one, and merges their sections into a single output file. When two source files contain a section with the same name, Combine reports the collision and merges their keys together, skipping any key that is an exact duplicate of one already present. With [strict name checking](#strict-name-checking) enabled, name collisions instead fail the run without writing anything.

### Why Combine?

- Build pipeline assembly: winapp2.ini's thousands of entries are generated into dozens of separate artifact files; Combine folds them into one database
- Collision detection: every section name defined by more than one input file is reported, naming the file that defined it first and every file that redefined it
- Build safety: strict name checking turns name collisions into a failed build
- Consolidating customizations: merge a folder of personal entry files into a single `custom.ini` for use with [Transmute](../transmute/readme.md)
- Organized maintenance: keep large ini collections split into manageable files 

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

1. Launch Combine from the winapp2ool main menu (option 7)
2. Use **Change the target directory** to select the directory containing your ini files
3. Optionally use **Change the save file location** to set a different output path
4. Run. all `.ini` files in the directory and its subdirectories will be merged into the output file

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
| Change the target directory | Select the directory to scan for ini files | Default: `current directory` |
| Change the save file location | Select the output file path and name | Default: `combined.ini` in the current directory |
| Strict Mode | Toggle [strict name checking](#strict-name-checking). | Fails the run instead of merging when a section name appears in more than one input file. Default: `False` |
| Log Viewer | Show the detailed log from the last Combine run | Only available after Combine has been run at least once during the current session |
| Reset Settings | Restore all settings to their defaults | Only shown when settings have been changed |

---

# Command-Line Arguments

Combine accepts file and directory parameters for specifying the target directory and output file location. It is CLI module **7**, invocable as `-combine` or `-7`.

Command line runs always begin from Combine's default settings.

## File Selection

| Arg | Effect | Default |
|:-|:-|:-|
| `-1d path` | Set the target directory | Current directory |
| `-3d path` | Set the output directory | Current directory |
| `-3f name` | Set the output file name | `combined.ini` |

###### Note: A leading `\` (eg. `-1d \Entries`) targets a subfolder of the current directory. 

## Toggles

| Arg | Effect |
|:-|:-|
| `-strict` | Enable [strict name checking](#strict-name-checking): when a section name appears in more than one input file, report every collision and exit with code 1 instead of saving the combined output |

---

# How Combining Works

## File Discovery

Combine searches the target directory and all of its subdirectories for files with the `.ini` extension. Files are always sorted alphabetically by their full path before processing. Because the entire path is compared, a file inside a subdirectory can sort ahead of a file in the target directory itself (see [Example 2](EXAMPLES.md#example-2-subdirectories-and-repeated-runs)).

If the configured output file is located inside the target directory, it is automatically skipped. A skipped output file still appears in the `Found N files` count, but does nothing and is excluded from the run statistics.

Files that contain no sections are skipped and excluded from the run statistics; this is noted in the [log](#run-summary-and-log) rather than the on-screen summary. Unreadable files produce a warning (see [Example 10](EXAMPLES.md#example-10-unreadable-and-empty-inputs)).

## Duplicate Section Handling

When two or more source files contain a section with the same name, Combine merges their keys:

- A key is skipped if another key with the same name and also the same value is already present in the output section
- Keys with the same name but a different value are not considered duplicates and are both included
- No keys are renumbered or reordered during combination

The merged section in the output contains the union of all unique keys from every source file that shared that section name. See [Example 3](EXAMPLES.md#example-3-merging-two-copies-of-the-same-entry) 

## Cross-File Name Collisions

Whenever a section name appears in more than one input file, Combine logs it on screen:

```
 ║                                2 section name(s) appeared in more than one input file:                             ║
 ║                  [Anki Media Cache *] first defined in A.ini, contributed again by browsers.ini, uwp.ini           ║
 ║                       [Audacity Temp Files *] first defined in A.ini, contributed again by uwp.ini                 ║
```

Whether a collision is a problem or not is situational and these logs are purely informational.

## Strict Name Checking

Strict name checking (**Strict Mode** in the menu, `-strict` on the command line, `off` by default) turns name collisions into an error:

- Every collision is logged exactly as above
- `Strict name checking is enabled - {file} will not be saved` is logged too
- No output file is written by Combine
- winapp2ool exits with error code **1** and saves a `winapp2ool.log`

## Output Formatting

The output is written as a plain ini file:

- Sections appear in the order they were first encountered across the alphabetically-sorted source files 
- Keys within a merged section appear in the order they were added
- Comments are not preserved. Comments from the source files will *not* be carried over into the output file.

If the final output contains no sections, nothing is written to disk.

## Run Summary and Log

A run prints a summary to the screen as it works:

```
 ╔════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╗
 ║                                        Combining files from C:\ini\case1\customs                                   ║
 ╠════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╣
 ║ Found 2 files with ini extension in C:\ini\case1\customs                                                           ║
 ║ Processed: C.ini (2 sections)                                                                                      ║
 ║ Processed: O.ini (1 sections)                                                                                      ║
 ║                                                                                                                    ║
 ║                                    Combined 2 files into custom.ini with 3 sections                                ║
 ╚════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╝
 ╔════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╗
 ║                                                  Combination complete!                                             ║
 ╚════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╝
Press any key to return to the menu.
```


The **Log Viewer** menu option displays information from the winapp2ool log that the on-screen summary omits:

| Log message | Meaning |
|:-|:-|
| `Added new section: [name] (N keys)` | A section was seen for the first time and copied into the output |
| `Added {key} to [{section}]` | A key was merged into an existing section of the same name |
| `Skipped duplicate key in [{section}]: {key}` | A key matched an existing key's name and value and was not added |
| `Skipping file with no sections: {file}` | A file was found but contained nothing |
| `Output file found in target directory, skipping: {path}` | The configured output file lives inside the target directory and was excluded |
| `N section name(s) appeared in more than one input file:` | Followed by one line per colliding name |
| `Strict name checking is enabled - {file} will not be saved` | A strict run found at least one collision |
| `Processed N files, M contained combinable sections` | Run statistic. `N` excludes the skipped output file; `M` excludes files with no sections |
| `Saving {file}` | The output was written to disk |

---

# Tips & Best Practices

### Output File Location

If your output file is inside the target directory, Combine detects and skips it automatically, so repeated runs are safe ([Example 2](EXAMPLES.md#example-2-subdirectories-and-repeated-runs)). Placing the output file outside the target directory is cleaner and avoids any ambiguity.

### Alphabetical Processing Order

Files are processed in alphabetical order by full path. When two files have a duplicate section and you want a specific file's keys to take precedence, name or locate the files such that the preferred one sorts first; later files contribute only their novel keys. See [Example 4](EXAMPLES.md#example-4-controlling-key-order-with-file-names).

### winapp2.ini Outputs

Combine does not renumber keys or apply winapp2.ini styling. If the combined result is a winapp2.ini file, run WinappDebug on it afterward (`winapp2ool -debug -c -1f yourfile.ini -3f yourfile.ini`) to normalize numbering, ordering, and style. See [Example 8](EXAMPLES.md#example-8-correcting-syntax-after-combining).

### Checking the Result

Use the Log Viewer after a run to see exactly which files were processed, which sections and keys each contributed, and whether any duplicate keys were skipped. 

---

# Troubleshooting

| Message / Symptom | Cause |
|:-|:-|
| "Target directory not found. Please select a valid directory." | The configured target directory does not exist. Shown in red atop the menu; the run does nothing |
| "N section name(s) appeared in more than one input file:" | Normal output. Informational unless your inputs were supposed to be disjoint |
| "Strict name checking is enabled - {file} will not be saved" | A `-strict` run found a collision. Nothing was written and the exit code is 1. |
| "No valid sections found to combine - {file} will not be saved" | No file in the target directory contained a parseable section; no output file is written |
| "Error processing file: {path}" followed by "Check the winapp2ool log for more information: {path}" | A file failed while to read. The remaining files are still processed, and the global log is saved to disk automatically |
| A scripted run fails with exit code 1 and no obvious error | A strict-mode collision. In silent mode nothing is printed; check `winapp2ool.log`, which is written automatically whenever the exit code is nonzero |
| No output file created, no warning shown | The target directory contained no `.ini` files at all (`Found 0 files...`) |
