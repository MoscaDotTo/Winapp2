# Combine Examples

Worked examples for [Combine](readme.md). The readme's [Common Uses](readme.md#common-uses) table lists the commands; these show what each one does, with real input and output.

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

The inputs, console output and output files below come from real Combine runs. Some console captures show only the lines that matter to the example.

## Basics

### Example 1: Combining Files with Distinct Sections

**Context**

You keep your own winapp2.ini entries the way the project keeps its own, split into per-letter files so each one stays small and easy to search. Your cleaner reads only one file, so they have to be joined before you can use them.

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
- `C.ini` sorts before `O.ini`, so its sections come first

---

### Example 2: Subdirectories and Repeated Runs

**Context**

Combine searches subdirectories too, and the output file can live inside the target directory. This example shows that a second run doesn't fold the first run's output back in.

**Intent**

We want to combine a collection directory, `archive` subfolder included, into a `combined.ini` inside that same directory, and be able to run the same command again safely.

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
- The `archive` subdirectory is searched without being asked for
- `[Winamp Cache *]` comes *first* because files sort by full path, and `collection\archive\retired.ini` sorts before `collection\current.ini`
- On the second run Combine finds `combined.ini` too (3 files), recognizes it as the output file and skips it
- The log records the skip: `Output file found in target directory, skipping: C:\ini\case2\collection\combined.ini`

---

## Merging Duplicate Sections

### Example 3: Merging Two Copies of the Same Entry

**Context**

You wrote a custom entry for Obsidian, then later wrote a second, more thorough one in a different file, having forgotten the first. Both describe `[Obsidian Workspace Cache *]`.

This example shows all three duplicate rules at once: a new key is added, a duplicate that differs only in case is skipped, and a key with the same name but a different value is kept beside the original.

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
- `my_customs.ini` sorts first, so its copy of `[Obsidian Workspace Cache *]` is the one the other file's keys are merged into
- `Section=My Custom Entries` from the second file is an exact duplicate, so it's skipped
- `detectfile=%appdata%\obsidian` is skipped too, because duplicates are matched ignoring case
- `FileKey1=...\Code Cache|*|RECURSE` has the same name as the existing `FileKey1` but a different value, so both are kept
- `FileKey2=...\GPUCache|*|RECURSE` is new, so it's added
- The collision report still appears, even though this merge is exactly what we wanted. Without strict mode it's only information

**Notes**

The output now has two keys named `FileKey1`. [Example 8](#example-8-correcting-syntax-after-combining) shows the WinappDebug pass that fixes that.

The Log Viewer shows the merge key by key:

```
Processing: my_customs.ini (1 sections)
  Added new section: [Obsidian Workspace Cache *] (3 keys)
Processed: my_customs.ini (1 sections)
Processing: second_opinion.ini (1 sections)
  Skipped duplicate key in [Obsidian Workspace Cache *]: Section
  Skipped duplicate key in [Obsidian Workspace Cache *]: detectfile
  Added FileKey1 to [Obsidian Workspace Cache *]
  Added FileKey2 to [Obsidian Workspace Cache *]
Processed: second_opinion.ini (1 sections)
```

---

### Example 4: Controlling Key Order with File Names

**Context**

When two copies of a section have keys with the same name but different values, both keys are kept, in the order Combine met them. Files are processed alphabetically by full path, so you can put one file's keys first by giving it a name that sorts first.

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
- The `1_` prefix makes `1_overrides.ini` sort before `handbrake.ini`, so its `FileKey1` is in the section first
- The bulk file's `Section`, `DetectFile` and its own `FileKey1` follow, in their original order

---

## Strict Name Checking

### Example 5: A Clean Strict Run

**Context**

When nothing goes wrong, strict mode doesn't show itself at all, so it helps to know what a passing run looks like.

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
- The output matches the non-strict run in [Example 1](#example-1-combining-files-with-distinct-sections). No section name is defined twice, so there's no collision report and the file is saved as usual

---

### Example 6: A Strict Run That Fails the Build

**Context**

This is the case strict mode is for. Three staged artifact files are supposed to hold separate sets of entries, but `[Anki Media Cache *]` has ended up in all three and `[Audacity Temp Files *]` in two. Without strict mode each would be merged into one entry carrying keys from every file, and the build would carry on.

**Intent**

We want the build to fail, with enough detail to find the duplicates, rather than produce a winapp2.ini with merged entries nobody wrote.

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

Combine doesn't write `Winapp2.ini`, and the process exits with code **1**.

**Explanation**
- Both colliding names are listed, because strict mode doesn't stop at the first collision
- The count is **2** because it counts colliding section *names*. There were three repeat definitions
- `[Anki Media Cache *]` lists two later files. `A.ini` counts as the first definition only because it sorts first

**Notes**

Without `-strict`, the same command prints the same collision report and then saves a 3-section file with each duplicated entry merged. That's the outcome strict mode is there to stop.

In silent mode (`-s`) nothing is printed, and the failure shows only in the exit code. A silent run that exits nonzero writes the global log, collision report included, to `winapp2ool.log`.

---

## Winapp2.ini Workflows

### Example 7: Building a custom.ini for Transmute

**Context**

[Transmute](../transmute/readme.md) applies one source file to one base file. If your customizations are spread across several files, Combine turns them into the single source file Transmute needs.

**Intent**

We want to join two customization files into one `custom.ini`, then use it to patch a winapp2.ini: adding a key to an entry that already exists, and adding an entry that doesn't.

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

Step 1, Combine:
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

Step 2, Transmute:
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
- Transmute matches it against the base file, adding `FileKey` to the entry that already existed and adding the entry that didn't
- Transmute saves the result as a winapp2.ini file, with the standard header (left out above), but it doesn't renumber keys. The added `FileKey` stays unnumbered until a WinappDebug pass like [Example 8](#example-8-correcting-syntax-after-combining)'s

**Notes**

Combine and Transmute's Add mode treat duplicates differently. Combine merges all its inputs as equals and skips exact duplicate keys. Transmute Add applies a source onto a base and doesn't check for duplicate keys at all.

---

### Example 8: Correcting Syntax After Combining

**Context**

This picks up from [Example 4](#example-4-controlling-key-order-with-file-names), whose output is a valid ini file but not a valid winapp2.ini entry: the keys are out of order and there are two `FileKey1`s.

**Intent**

We want to turn the combined output into valid winapp2.ini.

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
- `-c` tells WinappDebug to save its corrections instead of only reporting them
- The second `FileKey1` becomes `FileKey2`
- The keys are put back in winapp2.ini order
- FileKeys are sorted by value, which is why the `*.log` key ends up first
- The saved file also starts with the standard winapp2.ini header, left out above

---

### Example 9: Merging the winapp2.ini Build Artifacts

**Context**

This is Combine's day job. The winapp2.ini build puts its generated artifacts under the Winapp2 repository's `Assembler\Entries\` directory: the 27 per-letter base files (`#.ini`, `A.ini` … `Z.ini`, written by EntryBuilder), `Browsers\browsers.ini` (written by BrowserBuilder) and `UWP\uwp.ini` (written by UWPBuilder). One strict-mode Combine merges them all into the base winapp2.ini, and strict checking makes sure no entry name comes from more than one artifact.

**Intent**

We want to reproduce the build pipeline's merge stage.

**Command**
```
winapp2ool -s -offline -combine -strict -1d \Entries -3f Winapp2.ini
```

###### Note: This is the command `Assembler\build winapp2.ps1` runs, word for word, with `Assembler` as the working directory. `-s` (silent) and `-offline` are global winapp2ool flags the build uses for unattended runs, and `-1d \Entries` names the `Entries` subfolder of the working directory.

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
- Combine finds all 29 artifact files across the subfolders and processes them in alphabetical order by full path
- No two artifacts share a section name, so there's no collision report
- The output is saved in the working directory, as `Assembler\Winapp2.ini`
- The build itself runs in silent mode, so it prints none of this. The capture above is the same command without `-s`
- The section counts are from one day's build and change with every build

**Notes**

The files under `Assembler\Entries\` are generated. Edit their sources, never the files themselves.

---

## Diagnostics

### Example 10: Unreadable and Empty Inputs

**Context**

Not every file in a directory is usable. This example gives Combine three files: one holding only comments, one valid entry file, and one that another program has open and locked during the run.

**Intent**

We want to see how a bad input is reported, and check that it doesn't cost us the rest of the run.

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
- `a_notes.ini` has no sections, so it's skipped without a word on screen. Only the log records it: `Skipping file with no sections: a_notes.ini`
- `c_locked.ini` couldn't be read. That prints the two-line error and saves the global log to disk straight away
- The run carries on and still writes the output, but it exits with code **1**, so a script running it sees the failure
- `Found 3` but `Combined 1`: the found count is every `.ini` on disk, and the combined count is only the files that contributed sections

**Notes**

A directory where no file has a section produces no output file at all:

```
 ║ Found 1 files with ini extension in C:\ini\case10\nosections                                                       ║
 ║                        No valid sections found to combine - empty-result.ini will not be saved                     ║
```
