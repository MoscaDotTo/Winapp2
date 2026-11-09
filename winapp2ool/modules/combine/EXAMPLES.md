# Combine Examples

Worked examples for [Combine](readme.md). The [Common Uses](readme.md#common-uses) table in the readme lists the commands; these show what each one does, with real input and output.

# Table of Contents

1. [Basics](#basics)
   - [Example 1: Combining Files with Distinct Sections](#example-1-combining-files-with-distinct-sections)
   - [Example 2: Subdirectories and Repeated Runs](#example-2-subdirectories-and-repeated-runs)
2. [Merging Duplicate Sections](#merging-duplicate-sections)
   - [Example 3: Merging Two Copies of the Same Entry](#example-3-merging-two-copies-of-the-same-entry)
   - [Example 4: Controlling Key Order with File Names](#example-4-controlling-key-order-with-file-names)
3. [Strict Name Checking](#strict-name-checking)
   - [Example 5: A Clean Strict Run](#example-5-a-clean-strict-run)
   - [Example 6: A Strict Run That Fails the Build](#example-6-a-strict-run-that-fails-the-build)
4. [Winapp2.ini Workflows](#winapp2ini-workflows)
   - [Example 7: Building a custom.ini for Transmute](#example-7-building-a-customini-for-transmute)
   - [Example 8: Correcting Syntax After Combining](#example-8-correcting-syntax-after-combining)
   - [Example 9: Merging the winapp2.ini Build Artifacts](#example-9-merging-the-winapp2ini-build-artifacts)
5. [Diagnostics](#diagnostics)
   - [Example 10: Unreadable and Empty Inputs](#example-10-unreadable-and-empty-inputs)

---

The inputs, console output, and output files below are lifted verbatim from real Combine runs. 

## Basics

### Example 1: Combining Files with Distinct Sections

**Context**

You maintain your personal winapp2.ini entries the way the project maintains its own: split across per-letter files so each one stays small and easy to find things in. Your cleaner only reads one file, so they have to be joined before use.

**Intent**

We want to merge every file in `C:\ini\case1\customs` into a single `custom.ini` one directory up.

**Files**

###### **Target directory file (`customs\C.ini`)**
```ini
[Calibre Cover Cache *]
Section=My Custom Entries
DetectFile=%AppData%\calibre
FileKey1=%AppData%\calibre\caches|*|RECURSE

[Cemu Shader Cache *]
Section=My Custom Entries
DetectFile=%SystemDrive%\Emulation\Cemu
FileKey1=%SystemDrive%\Emulation\Cemu\shaderCache\transferable|*.bin
```

###### **Target directory file (`customs\O.ini`)**
```ini
[Obsidian Workspace Cache *]
Section=My Custom Entries
DetectFile=%AppData%\obsidian
FileKey1=%AppData%\obsidian\Cache|*|RECURSE
```

**Command**
```
winapp2ool -combine -1d C:\ini\case1\customs -3d C:\ini\case1 -3f custom.ini
```

**Output**

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
```

###### **Output file (`custom.ini`)**
```ini
[Calibre Cover Cache *]
Section=My Custom Entries
DetectFile=%AppData%\calibre
FileKey1=%AppData%\calibre\caches|*|RECURSE

[Cemu Shader Cache *]
Section=My Custom Entries
DetectFile=%SystemDrive%\Emulation\Cemu
FileKey1=%SystemDrive%\Emulation\Cemu\shaderCache\transferable|*.bin

[Obsidian Workspace Cache *]
Section=My Custom Entries
DetectFile=%AppData%\obsidian
FileKey1=%AppData%\obsidian\Cache|*|RECURSE
```

**Explanation**
- The target directory is `C:\ini\case1\customs`
- The output file is `C:\ini\case1\custom.ini`
- `C.ini` sorts before `O.ini`, so its sections appear first

---

### Example 2: Subdirectories and Repeated Runs

**Context**

Combine recurses into subdirectories, and the output file may live inside the target directory. We want to  demonstrate that a second run does not re-combine the first run's output.

**Intent**

We want to combine a collection directory (including its `archive` subfolder) into a `combined.ini` inside that same directory, and be able to re-run the combination safely.

**Files**

###### **Target directory file (`collection\current.ini`)**
```ini
[Foobar2000 Playlist Backups *]
Section=My Custom Entries
DetectFile=%AppData%\foobar2000-v2
FileKey1=%AppData%\foobar2000-v2\playlists-v2|*.bak
```

###### **Target directory file (`collection\archive\retired.ini`)**
```ini
[Winamp Cache *]
Section=My Custom Entries
DetectFile=%AppData%\Winamp
FileKey1=%AppData%\Winamp\Plugins\cache|*
```

**Command**
```
winapp2ool -combine -1d C:\ini\case2\collection -3d C:\ini\case2\collection -3f combined.ini
```

**Output**

First run:
```
 ║ Found 2 files with ini extension in C:\ini\case2\collection                                                        ║
 ║ Processed: retired.ini (1 sections)                                                                                ║
 ║ Processed: current.ini (1 sections)                                                                                ║
 ║                                                                                                                    ║
 ║                                   Combined 2 files into combined.ini with 2 sections                               ║
```

Second run of the identical command:
```
 ║ Found 3 files with ini extension in C:\ini\case2\collection                                                        ║
 ║ Processed: retired.ini (1 sections)                                                                                ║
 ║ Processed: current.ini (1 sections)                                                                                ║
 ║                                                                                                                    ║
 ║                                   Combined 2 files into combined.ini with 2 sections                               ║
```

###### **Output file (`collection\combined.ini`)**
```ini
[Winamp Cache *]
Section=My Custom Entries
DetectFile=%AppData%\Winamp
FileKey1=%AppData%\Winamp\Plugins\cache|*

[Foobar2000 Playlist Backups *]
Section=My Custom Entries
DetectFile=%AppData%\foobar2000-v2
FileKey1=%AppData%\foobar2000-v2\playlists-v2|*.bak
```

**Explanation**
- The `archive` subdirectory is searched automatically as part of the recursive call
- `[Winamp Cache *]` appears *first* because files sort by full path: `collection\archive\retired.ini` sorts before `collection\current.ini`
- On the second run, `combined.ini` itself is found (3 files) but recognized as the output file and skipped
- The skip is recorded in the log: `Output file found in target directory, skipping: C:\ini\case2\collection\combined.ini`

---

## Merging Duplicate Sections

### Example 3: Merging Two Copies of the Same Entry

**Context**

You wrote a custom entry for Obsidian, then later wrote a second, more thorough version in a different file without remembering the first. Both describe `[Obsidian Workspace Cache *]`. 

This example demonstrates all three duplicate-handling rules at once: a new key is added, an exact duplicate differing only in casing is skipped, and a same-name/different-value key is kept alongside the original

**Intent**

We want one `[Obsidian Workspace Cache *]` section containing every unique key from both files.

**Files**

###### **Target directory file (`customs\my_customs.ini`)**
```ini
[Obsidian Workspace Cache *]
Section=My Custom Entries
DetectFile=%AppData%\obsidian
FileKey1=%AppData%\obsidian\Cache|*|RECURSE
```

###### **Target directory file (`customs\second_opinion.ini`)**
```ini
[Obsidian Workspace Cache *]
Section=My Custom Entries
detectfile=%appdata%\obsidian
FileKey1=%AppData%\obsidian\Code Cache|*|RECURSE
FileKey2=%AppData%\obsidian\GPUCache|*|RECURSE
```

**Command**
```
winapp2ool -combine -1d C:\ini\case3\customs -3d C:\ini\case3 -3f combined.ini
```

**Output**

```
 ╔════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╗
 ║                                        Combining files from C:\ini\case3\customs                                   ║
 ╠════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╣
 ║ Found 2 files with ini extension in C:\ini\case3\customs                                                           ║
 ║ Processed: my_customs.ini (1 sections)                                                                             ║
 ║ Processed: second_opinion.ini (1 sections)                                                                         ║
 ║                                1 section name(s) appeared in more than one input file:                             ║
 ║          [Obsidian Workspace Cache *] first defined in my_customs.ini, contributed again by second_opinion.ini     ║
 ║                                                                                                                    ║
 ║                                   Combined 2 files into combined.ini with 1 sections                               ║
 ╚════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╝
```

###### **Output file (`combined.ini`)**
```ini
[Obsidian Workspace Cache *]
Section=My Custom Entries
DetectFile=%AppData%\obsidian
FileKey1=%AppData%\obsidian\Cache|*|RECURSE
FileKey1=%AppData%\obsidian\Code Cache|*|RECURSE
FileKey2=%AppData%\obsidian\GPUCache|*|RECURSE
```

**Explanation**
- `my_customs.ini` sorts first, so its copy of `[Obsidian Workspace Cache *]` is used as the base in the output section
- `Section=My Custom Entries` from the second file is an exact duplicate and is skipped
- `detectfile=%appdata%\obsidian` is also skipped. Duplicate matching is case-insensitive.
- `FileKey1=...\Code Cache|*|RECURSE` shares a name with the existing `FileKey1` but has a different value, so both are kept
- `FileKey2=...\GPUCache|*|RECURSE` is new and is added
- The collision log appears even though this merge is exactly what we wanted; it is informational here, not an error

**Notes**

Two keys named `FileKey1` now exist in the output. [Example 8](#example-8-correcting-syntax-after-combining) shows the WinappDebug pass that fixes this. 

The per-key detail behind this merge is visible in the Log Viewer:

```
Added new section: [Obsidian Workspace Cache *] (3 keys)
Processed: my_customs.ini (1 sections)
Skipped duplicate key in [Obsidian Workspace Cache *]: Section
Skipped duplicate key in [Obsidian Workspace Cache *]: detectfile
Added FileKey1 to [Obsidian Workspace Cache *]
Added FileKey2 to [Obsidian Workspace Cache *]
Processed: second_opinion.ini (1 sections)
```

---

### Example 4: Controlling Key Order with File Names

**Context**

When duplicate sections carry same-name/different-value keys, both are kept, but are ordered as they were encountered. Since files are processed alphabetically by full path, you can force a preferred file's keys to lead by naming it to sort first.

**Intent**

We want the key from our overrides file to appear ahead of the keys from the bulk file within the shared `[HandBrake Activity Logs *]` section.

**Files**

###### **Target directory file (`precedence\1_overrides.ini`)**
```ini
[HandBrake Activity Logs *]
FileKey1=%AppData%\HandBrake\logs|*.txt|REMOVESELF
```

###### **Target directory file (`precedence\handbrake.ini`)**
```ini
[HandBrake Activity Logs *]
Section=My Custom Entries
DetectFile=%AppData%\HandBrake
FileKey1=%AppData%\HandBrake\logs|*.log
```

**Command**
```
winapp2ool -combine -1d C:\ini\case4\precedence -3d C:\ini\case4 -3f combined.ini
```

**Output**

```
 ║ Found 2 files with ini extension in C:\ini\case4\precedence                                                        ║
 ║ Processed: 1_overrides.ini (1 sections)                                                                            ║
 ║ Processed: handbrake.ini (1 sections)                                                                              ║
 ║                                1 section name(s) appeared in more than one input file:                             ║
 ║             [HandBrake Activity Logs *] first defined in 1_overrides.ini, contributed again by handbrake.ini       ║
 ║                                                                                                                    ║
 ║                                   Combined 2 files into combined.ini with 1 sections                               ║
```

###### **Output file (`combined.ini`)**
```ini
[HandBrake Activity Logs *]
FileKey1=%AppData%\HandBrake\logs|*.txt|REMOVESELF
Section=My Custom Entries
DetectFile=%AppData%\HandBrake
FileKey1=%AppData%\HandBrake\logs|*.log
```

**Explanation**
- The `1_` prefix makes `1_overrides.ini` sort before `handbrake.ini`, so its `FileKey1` is part of the base section and appears first
- The bulk file's `Section`, `DetectFile`, and different-valued `FileKey1` are appended after it, in their original order

---

## Strict Name Checking

### Example 5: A Clean Strict Run

**Context**

Strict mode is invisible when nothing goes wrong, we want to know what success looks like. 

**Intent**

We want to confirm that the per-letter customs from [Example 1](#example-1-combining-files-with-distinct-sections) define no overlapping entry names.

**Command**
```
winapp2ool -combine -strict -1d C:\ini\case1\customs -3d C:\ini\case1 -3f custom.ini
```

**Output**

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
```

The same run in silent mode exits with code `0`:

```
winapp2ool -s -offline -combine -strict -1d C:\ini\case1\customs -3d C:\ini\case1 -3f custom.ini
```

**Explanation**
- The output is identical to the non-strict run in [Example 1](#example-1-combining-files-with-distinct-sections) - Because no section name is defined twice, no collision log is produced and the file is saved normally

---

### Example 6: A Strict Run That Fails the Build

**Context**

This is what strict mode exists for. Three staged artifact files are supposed to describe disjoint sets of entries, but a maintainer has added `[Anki Media Cache *]` in three of them and `[Audacity Temp Files *]` in two. Without strict mode these would silently merge into single entries carrying keys from every contributor.

**Intent**

We want the build to fail, with enough detail to find the duplicates, rather than build a potentially malformatted winapp2.ini

**Files**

###### **Target directory file (`Entries\A.ini`)**
```ini
[Anki Media Cache *]
Section=Games
DetectFile=%AppData%\Anki2
FileKey1=%AppData%\Anki2\*\media.trash|*|REMOVESELF

[Audacity Temp Files *]
Section=Multimedia
DetectFile=%AppData%\audacity
FileKey1=%LocalAppData%\Temp\audacity_temp|*|RECURSE
```

###### **Target directory file (`Entries\Browsers\browsers.ini`)**
```ini
[Anki Media Cache *]
Section=Games
DetectFile=%AppData%\Anki2
FileKey1=%AppData%\Anki2\*\media.trash|*|REMOVESELF

[Vivaldi Cache *]
Section=Vivaldi Web Browser
DetectFile=%LocalAppData%\Vivaldi\User Data
FileKey1=%LocalAppData%\Vivaldi\User Data\*\Cache|*|RECURSE
```

###### **Target directory file (`Entries\UWP\uwp.ini`)**
```ini
[Anki Media Cache *]
Section=Games
DetectFile=%AppData%\Anki2
FileKey1=%AppData%\Anki2\*\media.trash|*|REMOVESELF

[Audacity Temp Files *]
Section=Multimedia
DetectFile=%AppData%\audacity
FileKey1=%LocalAppData%\Temp\audacity_temp|*|RECURSE
```

**Command**
```
winapp2ool -combine -strict -1d C:\ini\case6\Entries -3d C:\ini\case6 -3f Winapp2.ini
```

**Output**

```
 ╔════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╗
 ║                                        Combining files from C:\ini\case6\Entries                                   ║
 ╠════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╣
 ║ Found 3 files with ini extension in C:\ini\case6\Entries                                                           ║
 ║ Processed: A.ini (2 sections)                                                                                      ║
 ║ Processed: browsers.ini (2 sections)                                                                               ║
 ║ Processed: uwp.ini (2 sections)                                                                                    ║
 ║                                2 section name(s) appeared in more than one input file:                             ║
 ║                  [Anki Media Cache *] first defined in A.ini, contributed again by browsers.ini, uwp.ini           ║
 ║                       [Audacity Temp Files *] first defined in A.ini, contributed again by uwp.ini                 ║
 ║                            Strict name checking is enabled - Winapp2.ini will not be saved                         ║
 ║                                                                                                                    ║
 ╚════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╝
```

No `Winapp2.ini` is written, and the process exits with code **1**.

**Explanation**
- Both colliding names are logged: strict mode does not stop at the first collision
- The count is **2** because two distinct section *names* collided, across three collision events
- `[Anki Media Cache *]` lists two later contributors; `A.ini` defined it first only because it sorts first

**Notes**

Dropping `-strict` from the same command produces the identical collision log, then saves a 3-section file in which each duplicated entry has been merged. This is what strict mode is meant to prevent

In silent mode (`-s`) nothing is printed at all; the failure shows only as the nonzero exit code. The global log (containing the collision log) is written to `winapp2ool.log` automatically whenever the exit code is nonzero.

---

## Winapp2.ini Workflows

### Example 7: Building a custom.ini for Transmute

**Context**

[Transmute](../transmute/readme.md) applies one source file to one base file. If your customizations are spread across several files, Combine is the step that turns them into the single source file Transmute expects.

**Intent**

We want to fold two customization files into one `custom.ini`, then use it to patch a winapp2.ini: adding a key to an entry that already exists and adding an entry that does not.

**Files**

###### **Target directory file (`customs\handbrake_extra.ini`)**
```ini
[HandBrake Activity Logs *]
FileKey=%AppData%\HandBrake\logs|*.txt|REMOVESELF
```

###### **Target directory file (`customs\obsidian.ini`)**
```ini
[Obsidian Workspace Cache *]
Section=My Custom Entries
DetectFile=%AppData%\obsidian
FileKey1=%AppData%\obsidian\Cache|*|RECURSE
```

###### **Base file (`base\winapp2.ini`)**
```ini
[Anki Media Cache *]
Section=Games
DetectFile=%AppData%\Anki2
FileKey1=%AppData%\Anki2\*\media.trash|*|REMOVESELF

[HandBrake Activity Logs *]
Section=Multimedia
DetectFile=%AppData%\HandBrake
FileKey1=%AppData%\HandBrake\logs|*.log
```

**Commands**
```
winapp2ool -combine -1d C:\ini\case7\customs -3d C:\ini\case7 -3f custom.ini
winapp2ool -transmute -add -1d C:\ini\case7\base -1f winapp2.ini -2d C:\ini\case7 -2f custom.ini -3d C:\ini\case7 -3f winapp2-transmuted.ini
```

**Output**

Step 1: combination:
```
 ║ Found 2 files with ini extension in C:\ini\case7\customs                                                           ║
 ║ Processed: handbrake_extra.ini (1 sections)                                                                        ║
 ║ Processed: obsidian.ini (1 sections)                                                                               ║
 ║                                                                                                                    ║
 ║                                    Combined 2 files into custom.ini with 2 sections                                ║
```

###### **Intermediate file (`custom.ini`)**
```ini
[HandBrake Activity Logs *]
FileKey=%AppData%\HandBrake\logs|*.txt|REMOVESELF

[Obsidian Workspace Cache *]
Section=My Custom Entries
DetectFile=%AppData%\obsidian
FileKey1=%AppData%\obsidian\Cache|*|RECURSE
```

Step 2: transmutation:
```
 ║ Transmutator: AddByKey - ByName                                                                                    ║
 ║ Adding keys to HandBrake Activity Logs *                                                                           ║
 ║   += Added key: FileKey=%AppData%\HandBrake\logs|*.txt|REMOVESELF                                                  ║
 ║ + Added new section: Obsidian Workspace Cache *                                                                   ║
```

###### **Output file (`winapp2-transmuted.ini`)**
```ini
[Anki Media Cache *]
Section=Games
DetectFile=%AppData%\Anki2
FileKey1=%AppData%\Anki2\*\media.trash|*|REMOVESELF

[HandBrake Activity Logs *]
Section=Multimedia
DetectFile=%AppData%\HandBrake
FileKey1=%AppData%\HandBrake\logs|*.log
FileKey=%AppData%\HandBrake\logs|*.txt|REMOVESELF

[Obsidian Workspace Cache *]
Section=My Custom Entries
DetectFile=%AppData%\obsidian
FileKey1=%AppData%\obsidian\Cache|*|RECURSE
```

**Explanation**
- Combine produces the single source file
- Transmute then matches the customizations against an existing base file, adding `FileKey` to the entry that already existed and inserting the entry that did not
- Transmute applies a winapp2.ini formatting pass  

**Notes**

Combine and Transmute Add differ in an important way. Combine merges all of its inputs symmetrically and skips exact duplicate keys; Transmute Add applies a source onto a base and does not check for duplicate keys. 

---

### Example 8: Correcting Syntax After Combining

**Context**

Continuing from [Example 4](#example-4-controlling-key-order-with-file-names), whose output is a valid ini file but an invalid winapp2.ini entry: the keys are out of order and `FileKey1` is defined twice. 

**Intent**

We want to normalize the combined output into valid winapp2.ini style.

**Files**

###### **Input file (`combined.ini`, produced by Example 4)**
```ini
[HandBrake Activity Logs *]
FileKey1=%AppData%\HandBrake\logs|*.txt|REMOVESELF
Section=My Custom Entries
DetectFile=%AppData%\HandBrake
FileKey1=%AppData%\HandBrake\logs|*.log
```

**Commands**
```
winapp2ool -combine -1d C:\ini\case4\precedence -3d C:\ini\case8 -3f combined.ini
winapp2ool -debug -c -1d C:\ini\case8 -1f combined.ini -3d C:\ini\case8 -3f combined.ini
```

**Output**

```
 ╔════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╗
 ║                                                  Linting winapp2.ini                                               ║
 ╠                                                                                                                    ╣
 ║ Error in [HandBrake Activity Logs *]:                                                                              ║
 ║ FileKey entry is incorrectly numbered.                                                                             ║
 ║ Expected: FileKey2                                                                                                 ║
 ║ Found:    FileKey1                                                                                                 ║
 ║                                                                                                                    ║
 ║ Error in [HandBrake Activity Logs *]:                                                                              ║
 ║ FileKey alphabetization                                                                                            ║
 ║ FileKey1=%AppData%\HandBrake\logs|*.txt|REMOVESELF appears to be out of place                                      ║
 ║ Expected position: 2                                                                                               ║
 ║                                                                                                                    ║
 ╠                                                                                                                    ╣
 ║                                                     Lint Complete!                                                 ║
 ╠                                                                                                                    ╣
 ║                                                     Entry count: 1                                                 ║
 ║                                              2 possible errors detected.                                           ║
 ║                                      combined.ini saved with any corrections made                                  ║
 ║                                                                                                                    ║
 ╚════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╝
```

###### **Output file (`combined.ini`)**
```ini
[HandBrake Activity Logs *]
Section=My Custom Entries
DetectFile=%AppData%\HandBrake
FileKey1=%AppData%\HandBrake\logs|*.log
FileKey2=%AppData%\HandBrake\logs|*.txt|REMOVESELF
```

**Explanation**
- The `-c` flag tells WinappDebug to write its corrections rather than only reporting them
- The duplicated `FileKey1` is renumbered to `FileKey2`, resolving the collision Combine left behind
- The keys are regrouped into winapp2.ini order
- The FileKeys are alphabetized by value, which is why the `*.log` key ends up first

---

### Example 9: Merging the winapp2.ini Build Artifacts

**Context**

This is Combine's production job. The winapp2.ini build pipeline stages all of its generated artifacts under the Winapp2 repository's `Assembler\Entries\` directory: the 27 per-letter base files (`#.ini`, `A.ini` … `Z.ini`, written by EntryBuilder), `Browsers\browsers.ini` (written by BrowserBuilder), and `UWP\uwp.ini` (written by UWPBuilder). A single strict-mode Combine merges all of them into the base winapp2.ini in one pass. Strict name checking guarantees no entry name is defined by more than one artifact.

**Intent**

We want to reproduce the build pipeline's merge stage.

**Command**
```
winapp2ool -s -offline -combine -strict -1d \Entries -3f Winapp2.ini
```

###### Note: This is the invocation from `Assembler\build winapp2.ps1`, verbatim, run with `Assembler` as the working directory. `-s` (silent) and `-offline` are global winapp2ool flags used by the pipeline for unattended runs; `-1d \Entries` targets the `Entries` subfolder of the working directory.

**Output**

```
 ╔════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╗
 ║                       Combining files from C:\Users\hazel\source\repos\Winapp2\Assembler\Entries                   ║
 ╠════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╣
 ║ Found 29 files with ini extension in C:\Users\hazel\source\repos\Winapp2\Assembler\Entries                         ║
 ║ Processed: #.ini (37 sections)                                                                                     ║
 ║ Processed: A.ini (365 sections)                                                                                    ║
 ║ Processed: B.ini (72 sections)                                                                                     ║
 ║ Processed: browsers.ini (1240 sections)                                                                            ║
 ║ Processed: C.ini (165 sections)                                                                                    ║
 ║ Processed: D.ini (148 sections)                                                                                    ║
 ║ Processed: E.ini (87 sections)                                                                                     ║
 ║ Processed: F.ini (92 sections)                                                                                     ║
 ║ Processed: G.ini (81 sections)                                                                                     ║
 ║ Processed: H.ini (61 sections)                                                                                     ║
 ║ Processed: I.ini (98 sections)                                                                                     ║
 ║ Processed: J.ini (34 sections)                                                                                     ║
 ║ Processed: K.ini (31 sections)                                                                                     ║
 ║ Processed: L.ini (73 sections)                                                                                     ║
 ║ Processed: M.ini (219 sections)                                                                                    ║
 ║ Processed: N.ini (101 sections)                                                                                    ║
 ║ Processed: O.ini (39 sections)                                                                                     ║
 ║ Processed: P.ini (157 sections)                                                                                    ║
 ║ Processed: Q.ini (17 sections)                                                                                     ║
 ║ Processed: R.ini (97 sections)                                                                                     ║
 ║ Processed: S.ini (252 sections)                                                                                    ║
 ║ Processed: T.ini (127 sections)                                                                                    ║
 ║ Processed: U.ini (29 sections)                                                                                     ║
 ║ Processed: uwp.ini (187 sections)                                                                                  ║
 ║ Processed: V.ini (51 sections)                                                                                     ║
 ║ Processed: W.ini (143 sections)                                                                                    ║
 ║ Processed: X.ini (22 sections)                                                                                     ║
 ║ Processed: Y.ini (4 sections)                                                                                      ║
 ║ Processed: Z.ini (11 sections)                                                                                     ║
 ║                                                                                                                    ║
 ║                                  Combined 29 files into Winapp2.ini with 4040 sections                             ║
 ╚════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╝
```

**Explanation**
- All 29 artifact files are discovered recursively and processed in alphabetical order by full path 
- The staged artifacts have no overlapping section names, so no collision log appears. 
- The output is saved in the working directory (`Assembler\Winapp2.ini`) 
- In silent mode nothing is printed to the screen

**Notes**

The files under `Assembler\Entries\` are generated build artifacts

---

## Diagnostics

### Example 10: Unreadable and Empty Inputs

**Context**

Not every file in a directory is usable. This example puts three files in front of Combine: one holding only comments, one valid entry file, and one that is locked open by another program at the moment of the run.

**Intent**

We want to see how a bad input is reported and confirm that it does not cost us the rest of the run.

**Files**

###### **Target directory file (`sources\a_notes.ini`)**
```ini
; Scratch notes for entries I still need to write
; - Anki: check whether media.trash is safe to purge
; - Cemu: shader cache rebuild time is significant, maybe skip
```

###### **Target directory file (`sources\b_good.ini`)**
```ini
[Anki Media Cache *]
Section=My Custom Entries
DetectFile=%AppData%\Anki2
FileKey1=%AppData%\Anki2\*\media.trash|*|REMOVESELF
```

###### **Target directory file (`sources\c_locked.ini`)**: held open by another process during the run
```ini
[Cemu Shader Cache *]
Section=My Custom Entries
DetectFile=%SystemDrive%\Emulation\Cemu
FileKey1=%SystemDrive%\Emulation\Cemu\shaderCache\transferable|*.bin
```

**Command**
```
winapp2ool -combine -1d C:\ini\case10\sources -3d C:\ini\case10 -3f combined.ini
```

**Output**

```
 ╔════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╗
 ║                                       Combining files from C:\ini\case10\sources                                   ║
 ╠════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╣
 ║ Found 3 files with ini extension in C:\ini\case10\sources                                                          ║
 ║ Processed: b_good.ini (1 sections)                                                                                 ║
 ║                                Error processing file: C:\ini\case10\sources\c_locked.ini                           ║
 ║                      Check the winapp2ool log for more information: C:\ini\case10\winapp2ool.log                   ║
 ║                                                                                                                    ║
 ║                                   Combined 1 files into combined.ini with 1 sections                               ║
 ╚════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╝
```

###### **Output file (`combined.ini`)**
```ini
[Anki Media Cache *]
Section=My Custom Entries
DetectFile=%AppData%\Anki2
FileKey1=%AppData%\Anki2\*\media.trash|*|REMOVESELF
```

**Explanation**
- `a_notes.ini` contains no sections and is silently skipped. Only the log records it: `Skipping file with no sections: a_notes.ini`
- `c_locked.ini` could not be read, which produces the two-line error and an automatic save of the global log to disk
- The run continues regardless and still writes the output
- `Found 3` but `Combined 1`: the found count is every `.ini` on disk, while the combined count is only files that actually contributed sections

**Notes**

A directory in which nothing is parseable produces no output file at all:

```
 ║ Found 1 files with ini extension in C:\ini\case10\nosections                                                       ║
 ║                        No valid sections found to combine - empty-result.ini will not be saved                     ║
```
