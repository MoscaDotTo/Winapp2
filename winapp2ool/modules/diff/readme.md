# Diff

**Diff** is a winapp2ool module that generates a semantic, context-aware changelog between two versions of winapp2.ini. Unlike a standard [diff](https://en.wikipedia.org/wiki/Diff), **Diff** understands the structure and syntax of winapp2.ini entries and categorizes every change by what actually happened: entries added, modified, renamed, merged into other entries, or removed without replacement, including detecting when keys are captured by abstractions such as wildcards.

### What does Diff do?

**Diff** compares two winapp2.ini files and classifies every entry into one of three top-level categories:

- **Added:** present in the new file and not in the old file
- **Modified:** present in both files with the same name but changed in some way (excluding Mergers and Renames)
- **Removed:** present in the old file but not in the new file

Removed entries are further classified:

- **Renamed:** the entry no longer exists by its old name, but its content was found mostly unchanged in another entry in the new file
- **Merged:** the entry no longer exists, but its content was absorbed into another entry that is substantially different from the old version
- **Removed without replacement:** the entry is gone and its content does not appear in any other entry in the new file

By default, **Diff** compares a local winapp2.ini against the latest version available on GitHub, and provides an option to trim the downloaded file for the current system before comparison.

###### The downloaded file is flavor-specific, and the default flavor is CCleaner. See [Flavors](#flavors).

### Why Diff?

- **Coherent, semantic changelogs:** Diff understands winapp2.ini structure and syntax; it classifies changes by what actually happened, not which lines changed, and ignores irrelevant text formatting changes.
- **CI/CD integration:** Automated changelog generation via command-line args and log saving; Diff generates the official changelogs for the winapp2.ini project.
- **Noise reduction:** the optional trim filters the diff to changes relevant to installed software only, handy if you keep your local copy trimmed.

---

# Table of Contents

1. [Requirements](#requirements)
2. [Quick Start](#quick-start)
   - [Flavors](#flavors)
3. [Menu Options](#menu-options)
4. [Diff Categories](#diff-categories)
   - [Added Entries](#added-entries)
     - [Novel entries](#novel-entries)
     - [Entries consolidating removed entries](#entries-consolidating-removed-entries)
   - [Modified Entries](#modified-entries)
     - [Simple modification](#simple-modification)
     - [Modification by merger](#modification-by-merger)
     - [Renames with key changes](#renames-with-key-changes)
     - [Cross-Entry Key Movements](#cross-entry-key-movements)
   - [Removed Entries](#removed-entries)
     - [Renamed](#renamed)
     - [Merged](#merged)
     - [Removed Without Replacement](#removed-without-replacement)
5. [How Classification Works](#how-classification-works)
   - [Value Normalization](#value-normalization)
   - [Candidate Selection](#candidate-selection)
   - [Matching Basis](#matching-basis)
   - [Vague Paths](#vague-paths)
   - [Classification Rules](#classification-rules)
6. [Output and Log](#output-and-log)
   - [Output Structure](#output-structure)
   - [Console Output vs. Saved Log](#console-output-vs-saved-log)
   - [Sample Output](#sample-output)
   - [Interpreting the Summary](#interpreting-the-summary)
   - [Source Entry Key Status Report](#source-entry-key-status-report)
7. [Command-Line Arguments](#command-line-arguments)
   - [Toggles](#toggles)
   - [File Selection](#file-selection)
   - [Flavor Selection](#flavor-selection)
   - [Examples](#examples)
8. [Tips & Best Practices](#tips--best-practices)
   - [Controlling Scope (Trimming)](#trimming)
   - [CI/CD and Scripting](#scripting)
9. [Troubleshooting](#troubleshooting)
10. [Usage Examples](#usage-examples)
    - [Example 1: Default Workflow](#example-1-default-workflow)
    - [Example 2: Full Database Comparison](#example-2-full-database-comparison)
    - [Example 3: Comparing Archived Versions](#example-3-comparing-archived-versions)
    - [Example 4: Verbose Mode](#example-4-verbose-mode)
    - [Example 5: Saving the Log](#example-5-saving-the-log)
    - [Example 6: Diffing a Non-Default Flavor](#example-6-diffing-a-non-default-flavor)
    - [Example 7: Automated Changelogs (CI/CD)](#example-7-automated-changelogs-cicd)

---

# Requirements

- A local winapp2.ini file to use as the "older" version.
- A second winapp2.ini against which to compare; either a local file or downloaded on the fly from GitHub.
- Both files should be the same [flavor](#flavors).  

---

# Quick Start

### Configuration (Optional)

Before running Diff, you can toggle several settings from the Diff menu:

- **Remote Diffing:** Downloading and diffing against GitHub is enabled by default. Select `Toggle diffing against GitHub` to disable.
- **Verbose Mode:** Select `Toggle verbose mode` to print the full text of each changed entry alongside the key-change summary.
- **Trimming:** When diffing against GitHub, trimming is enabled by default. Select `Toggle remote file trim` to disable. In offline mode, both downloading and trimming default to off.
- **Logging:** Saving the diff log (which includes the Source Entry Key Status Report) is disabled by default. Select `Toggle log saving` to enable.
- Use `Choose older/local file` to specify an older/local `winapp2.ini` in a different directory or with a different name.
- Use `Choose newer file` to specify a newer local `winapp2.ini` when diffing against GitHub is disabled.

The flavor is not a Diff setting: it lives in winapp2ool's Global Settings menu (`Change Flavor`) and applies to every module that downloads winapp2.ini.

### Running Diff against GitHub

This is the default setting.

1. From the winapp2ool main menu, select `Diff`.
2. (Optional) Select `Choose older/local file` if your local `winapp2.ini` is in a different directory or has a different name.
3. Select `Run (default)`, or just press `Enter`.
4. Diff downloads the latest winapp2.ini of the current flavor from GitHub, trims it for the current system, and displays the results.

### Running Diff against local files

Use this to compare two specific versions already on your machine.

1. From the winapp2ool main menu, select `Diff`.
2. Select `Toggle diffing against GitHub`.
3. Select `Choose older/local file` if your older copy of `winapp2.ini` is in a different directory or has a different name.
4. Select `Choose newer file` to select the newer copy of `winapp2.ini`.
5. Select `Run`.

### Running Diff from the winapp2ool main menu

When a new winapp2.ini is available, the main menu offers `Show winapp2.ini changelog` and `Show trimmed changelog`. Both diff the `winapp2.ini` in the current directory against the latest file of the current flavor, and the second trims the download first. They follow the Diff module's other settings (verbose modeand log saving) but never change them.

## Flavors

winapp2.ini is published in several **flavors** which are variants designed to suit individual applications. When Diff downloads the "newer" file it resolves the current flavor setting to pick which published file to fetch, exactly as the Download module does.

The default flavor is CCleaner, not base.`winapp2ool -diff` without a flavor argument compares your local file against the *CCleaner* build. Diffing between different versions can be noisy and is not recommended. 

Set the flavor for a single run with a CLI flag (see [Flavor Selection](#flavor-selection)), or persistently via the main menu's `Global Settings` → `Change Flavor`. The Diff menu does not display the current flavor.

---

# Menu Options

| Option | Effect | Notes |
|:-|:-|:-|
| Run (default) | Compare the two files and display the changelog | A "newer" file must be selected or downloading must be enabled. Pressing `Enter` on an empty prompt runs it too |
| Toggle diffing against GitHub | Enable or disable downloading the latest winapp2.ini as the "newer" file | Default: `Enabled`. Unavailable in offline mode |
| Toggle remote file trim | Enable or disable trimming the downloaded file for the current system before diffing | Default: `Enabled`; only available when downloading |
| Toggle log saving | Enable or disable saving the diff output to disk | Default: `Disabled` |
| Toggle verbose mode | Print the full text of changed entries alongside each change | Default: `Disabled`; CLI: `-verbose` |
| Choose older/local file | Select the "old" version of winapp2.ini | Default: `winapp2.ini` in the current directory |
| Choose newer file | Select the "new" version of winapp2.ini | Only shown when downloading is disabled |
| Choose save target | Select where to save the diff output | Only shown when log saving is enabled; default: `diff.txt` |
| Log Viewer | Show the most recent diff output | Only available after Diff has been run at least once during the current session |
| Reset Settings | Restore all settings to their defaults | Only shown when settings have been changed |

---

# Diff Categories

###### Note: The output excerpts in this section are taken verbatim from the saved log of a real Diff between base-flavor database versions 220510 (May 2022) and 251109 (November 2025). To reproduce them, extract `Non-CCleaner/Winapp2.ini` at commit `f5ea7371` and at tag `v251109`, then diff the two locally.

## Added Entries

An entry is **Added** if it exists in the new file and has no entry of the same name in the old file. Added entries render in two separate sections of the output: entries which consolidate content from removed entries appear first (under `Added entries containing merged content:`), followed by entirely novel entries (under `Added entries:`). Each section reports its own total count.

### Novel entries

Entirely new entries with no connection to any removed entry appear as a single line:

```
     .NET Platform * has been added

     360 Secure Browser Autofill Data & Search Engine Preferences * has been added
```

### Entries consolidating removed entries

When a new entry contains keys from one or more removed entries, its header is followed by a `This entry contains content originating from N removed entries:` line listing the source entries. Diff then sorts the keys of the new entry and its sources into these categories:

| Category | Type | Annotation | Conditions |
|:-|:-|:-|:-|
| **Novel** | Added keys | `(novel)` in an `N keys added:` block | key is new content with no equivalent in any source entry |
| **Moved in** | Added keys | `(from [Other *])` in an `N keys added:` block | key [moved](#cross-entry-key-movements) here from an entry that still exists |
| **Carried over** | Carried-over keys | `(from [Source *])` in an `N keys carried over from merged entries:` block | key is carried over verbatim from the source entry |
| **Dropped** | Removed keys | Appears in an `N keys from merged entries not in this entry:` block | the key existed in a source entry merged into this entry alone, and this entry doesn't have it, unchanged or captured. A split source's lost keys are listed under the split. |
| **Captured** | Modified keys | `(from [Source *])` on the old key | the key does not appear verbatim but is replaced by another key that covers it — typically a wildcard path that absorbs one or more specific old paths |

A key that several source entries share is credited to each of them in name order, eg. `(from [Source A *], [Source B *])`. Past three sources, Diff names the first three and counts the rest: `(from [Source A *], [Source B *], [Source C *] and 4 others)`.

**Example 1: Novel and carried keys (Abelssoft PCFresh):**

**Old (v220510):**
```ini
[Abelssoft PCFresh Backups *]
LangSecRef=3024
DetectFile=%LocalAppData%\Abelssoft\PCFresh
FileKey1=%LocalAppData%\Abelssoft\PCFresh\Backup|*.*
```

**New (v251109):**
```ini
[Abelssoft PCFresh *]
LangSecRef=3024
DetectFile=%LocalAppData%\Abelssoft\PCFresh
FileKey1=%LocalAppData%\Abelssoft\PCFresh|*.log
FileKey2=%LocalAppData%\Abelssoft\PCFresh\Backup|*
```

**Diff output:**
```
   Abelssoft PCFresh * has been added
   This entry contains content originating from 1 removed entry:
     - Abelssoft PCFresh Backups *

   1 key added:

       Added 1 FileKey
             FileKey1=%LocalAppData%\Abelssoft\PCFresh|*.log (novel)

   3 keys carried over from merged entries:

       Carried over 1 LangSecRef
       Carried over 1 DetectFile
       Carried over 1 FileKey
             LangSecRef=3024 (from [Abelssoft PCFresh Backups *])
             DetectFile=%LocalAppData%\Abelssoft\PCFresh (from [Abelssoft PCFresh Backups *])
             FileKey2=%LocalAppData%\Abelssoft\PCFresh\Backup|* (from [Abelssoft PCFresh Backups *])
```

**Explanation:**

`FileKey1` is novel; no equivalent exists in `[Abelssoft PCFresh Backups *]`.

`LangSecRef`, `DetectFile`, and `FileKey2` were carried over unchanged, so they're listed in their own group, with attribution to their source entry.

Note that `FileKey2`'s pattern changed from `*.*` to `*`, and that the old key is printed as `|*`. Diff rewrites deprecated values in the old file before comparing, so the key counts as carried over rather than modified. See [Value Normalization](#value-normalization).

**Example 2: Captured keys (EFSoftware):**

**Old (v220510):**

```ini
[EF Duplicate Files Manager *]
LangSecRef=3024
DetectFile=%AppData%\EFSoftware\DFMCache
FileKey1=%AppData%\EFSoftware|DFMCache

[EF Duplicate MP3 Finder *]
LangSecRef=3024
DetectFile=%AppData%\EFSoftware\MP3Cache
FileKey1=%AppData%\EFSoftware|MP3Cache
```

**New (v251109):**

```ini
[EFSoftware  *]
LangSecRef=3024
DetectFile=%AppData%\EFSoftware
FileKey1=%AppData%\EFSoftware|*Cache
```

**Diff output:**

```
   EFSoftware  * has been added
   This entry contains content originating from 2 removed entries:
     - EF Duplicate Files Manager *
     - EF Duplicate MP3 Finder *

   1 key carried over from merged entries:

       Carried over 1 LangSecRef
             LangSecRef=3024 (from [EF Duplicate Files Manager *], [EF Duplicate MP3 Finder *])

   2 keys capturing content from merged entries:

       Modified 1 Detection criteria
       Modified 1 FileKey

       Detection criteria modified, replacing 2 old keys
              + New: DetectFile=%AppData%\EFSoftware
              - Old: DetectFile=%AppData%\EFSoftware\DFMCache (from [EF Duplicate Files Manager *])
              - Old: DetectFile=%AppData%\EFSoftware\MP3Cache (from [EF Duplicate MP3 Finder *])
       FileKey1 has been modified, replacing 2 old keys
              + New: FileKey1=%AppData%\EFSoftware|*Cache
              - Old: FileKey1=%AppData%\EFSoftware|DFMCache (from [EF Duplicate Files Manager *])
              - Old: FileKey1=%AppData%\EFSoftware|MP3Cache (from [EF Duplicate MP3 Finder *])
```

**Explanation:**

`LangSecRef` is carried over and credited to both source entries.

The `DetectFile` and `FileKey1` in the new entry each cover both old paths with a single pattern; `%AppData%\EFSoftware` is a parent path, and `*Cache` matches both old parameters. Both are reported as captured (modified) keys replacing 2 old keys each.

**Example 3: Dropped keys (Cook'n):**

Three related Cook'n entries were consolidated into a single `Cook'n *` entry, but the `Warning=` key from `Cook'n Dups *` was not carried over or replaced.

**Old (v220510):**

```ini
[Cook'n Cache *]
LangSecRef=3021
Detect=HKLM\Software\Microsoft\Windows\CurrentVersion\Uninstall\Cook'n
FileKey1=%AppData%\Mozilla\eclipse\Cache|*.*

[Cook'n Dups *]
LangSecRef=3021
Detect=HKLM\Software\Microsoft\Windows\CurrentVersion\Uninstall\Cook'n
Warning=This removes identical duplicates of the huge Getting Started Guide.
FileKey1=%LocalAppData%\DVO\Cook'n10App\plugins|Getting Started Guide.rtf|RECURSE

[Cook'n Logs *]
LangSecRef=3021
Detect=HKLM\Software\Microsoft\Windows\CurrentVersion\Uninstall\Cook'n
FileKey1=%Documents%\Cook'n*|*.log|RECURSE
FileKey2=%LocalAppData%\DVO\Cook'n*App|*.log|RECURSE
```

**New (v251109):**

```ini
[Cook'n *]
LangSecRef=3021
Detect=HKLM\Software\Microsoft\Windows\CurrentVersion\Uninstall\Cook'n
FileKey1=%AppData%\Mozilla\eclipse\Cache|*
FileKey2=%LocalAppData%\DVO\Cook'n*App|*.log|RECURSE
FileKey3=%LocalAppData%\DVO\Cook'n10App\plugins|Getting Started Guide.rtf|RECURSE
FileKey4=%UserProfile%\Documents\Cook'n*|*.log|RECURSE
```

**Diff output:**

```
   Cook'n * has been added
   This entry contains content originating from 3 removed entries:
     - Cook'n Cache *
     - Cook'n Dups *
     - Cook'n Logs *

   6 keys carried over from merged entries:

       Carried over 1 LangSecRef
       Carried over 1 Detect
       Carried over 4 FileKeys
             LangSecRef=3021 (from [Cook'n Cache *], [Cook'n Dups *], [Cook'n Logs *])
             Detect=HKLM\Software\Microsoft\Windows\CurrentVersion\Uninstall\Cook'n (from [Cook'n Cache *], [Cook'n Dups *], [Cook'n Logs *])
             FileKey1=%AppData%\Mozilla\eclipse\Cache|* (from [Cook'n Cache *])
             FileKey2=%LocalAppData%\DVO\Cook'n*App|*.log|RECURSE (from [Cook'n Logs *])
             FileKey3=%LocalAppData%\DVO\Cook'n10App\plugins|Getting Started Guide.rtf|RECURSE (from [Cook'n Dups *])
             FileKey4=%UserProfile%\Documents\Cook'n*|*.log|RECURSE (from [Cook'n Logs *])

   1 key from merged entries not in this entry:

       Dropped 1 Warning
             Warning=This removes identical duplicates of the huge Getting Started Guide. (from [Cook'n Dups *])
```

**Explanation:**

Every `FileKey` from all three source entries was carried over, attributed to its originating entry.

Keys that all three source entries shared (`LangSecRef`, `Detect`) are credited to all three.

The `Warning=` key from `Cook'n Dups *` was dropped, and is listed under the dropped heading with attribution to its originating entry.

The change from `%Documents%` to `%UserProfile%\Documents` is a deprecated-path rewrite, so `FileKey4` matches `Cook'n Logs *`'s `FileKey1` despite the textual difference. See [Value Normalization](#value-normalization).

---

## Modified Entries

An entry is **Modified** if it exists in both files with the same name but has changed. Modifications include added keys, removed keys, and changed key values or parameters. Diff itemizes the specific key-level changes for each modified entry.

Keys annotated `(novel)` are new content with no detected equivalent in the old file. Keys annotated `(from [Entry Name *])` were sourced from a specific entry: the old version of the same entry, a removed entry that was merged in, or an entry the key [moved](#cross-entry-key-movements) from.

Modified output itemizes five kinds of key-level change:

- **Added keys**: present in the new entry with no detected equivalent in the old entry
- **Removed keys**: present in the old entry with no detected equivalent in the new entry or anywhere else
- **Modified keys**: a key whose value changed, shown as `X has been modified, replacing N old keys`; in the simple case one old key was updated to a new value; in more complex cases a single new key uses a wildcard or semicolon-joined pattern to capture multiple old keys
- **Moved keys**: keys that left this entry for another one, listed under `N keys moved to other entries:` as `FileKey2=... moved to [Other *]`. See [Cross-Entry Key Movements](#cross-entry-key-movements)
- **Moved-in keys**: keys this entry took from another entry that still exists, listed under `N keys moved in from other entries:` as `FileKey2=... (from [Other *])`. An entry that also absorbed removed entries lists these among its added keys instead, with the same credit. 
 ###### Note: A moved-in key that the block shows as an update or capture of a merged source's key credits only the source, so the key is also listed under `N keys moved in from other entries:`

Each kind has its own list, under a heading with its total (`13 keys added:`, `2 keys removed:`, `7 keys modified:`), followed by a count line for each key type (`Added 2 DetectFiles`). An entry whose only change is a moved key is still itemized.

Changes to detection keys (`Detect`, `DetectFile`) are grouped and reported together as `Detection criteria`.

Modified entries are displayed in two separate output sections: entries which absorbed content from removed entries appear first (each followed by a note itemizing the removed entries against which the changes are measured), and entries with ordinary modifications appear later under the `Modified entries:` header.

### Simple modification

**Example: Modified keys and detection criteria (Facebook):**

**Old (v220510):**

```ini
[Facebook *]
LangSecRef=3031
Detect=HKCU\Software\Classes\Local Settings\Software\Microsoft\Windows\CurrentVersion\AppModel\SystemAppData\Facebook.Facebook_8xx8rvfyw5nnt
FileKey1=%LocalAppData%\Packages\Facebook.Facebook_*\AC\INet*|*.*|RECURSE
FileKey2=%LocalAppData%\Packages\Facebook.Facebook_*\AC\Microsoft\CryptnetUrlCache\*|*.*|RECURSE
FileKey3=%LocalAppData%\Packages\Facebook.Facebook_*\AC\Temp|*.*|RECURSE
```

**New (v251109):**

```ini
[Facebook *]
LangSecRef=3022
DetectFile=%LocalAppData%\Packages\Facebook.Facebook_*
FileKey1=%LocalAppData%\Packages\Facebook.Facebook_*\AC|*|RECURSE
```

**Diff output:**

```
     Facebook * has been modified

     3 keys modified:

       Modified 1 LangSecRef
       Modified 1 FileKey
       Modified 1 Detection criteria

       LangSecRef has been modified, replacing 1 old key
              + New: LangSecRef=3022
              - Old: LangSecRef=3031
       FileKey1 has been modified, replacing 3 old keys
              + New: FileKey1=%LocalAppData%\Packages\Facebook.Facebook_*\AC|*|RECURSE
              - Old: FileKey1=%LocalAppData%\Packages\Facebook.Facebook_*\AC\INet*|*|RECURSE
              - Old: FileKey2=%LocalAppData%\Packages\Facebook.Facebook_*\AC\Microsoft\CryptnetUrlCache\*|*|RECURSE
              - Old: FileKey3=%LocalAppData%\Packages\Facebook.Facebook_*\AC\Temp|*|RECURSE
       Detection criteria modified
              + New: DetectFile=%LocalAppData%\Packages\Facebook.Facebook_*
              - Old: Detect=HKCU\Software\Classes\Local Settings\Software\Microsoft\Windows\CurrentVersion\AppModel\SystemAppData\Facebook.Facebook_8xx8rvfyw5nnt
```

**Explanation:**

The updated `LangSecRef` is identified as modified.

The updated `FileKey1` is identified as modified, replacing 3 old keys. The new `\AC|*|RECURSE` path captures the three separate subdirectory paths from the old entry.

The replacement of `Detect` with `DetectFile` is paired into a single `Detection criteria modified` notification.

### Modification by merger

When entries from the old file are merged into a modified entry, their absorbed keys appear alongside the entry's own changes. The `(from [Entry Name *])` annotation on each old key identifies the old entry or entries from which it came. Keys labeled with the entry's own name came from its prior version; keys labeled with a different name came from a removed entry that was consolidated. Keys the entry's prior version had and lost are listed as removed; keys that only a merged entry had and that didn't survive are listed separately, under `N keys from merged entries not in this entry:`, as in an added entry's merger block. When a removed entry was split across several entries, its keys aren't listed as dropped here. A key that went to one of the others is counted on a closing line that names where it went, eg. `2 other keys from merged entries went to [Windows Font Cache *]`. A key that no entry kept is listed once, in the removed entry's own section under [Merged](#merged), and counted on a second closing line, eg. `173 other keys that no entry kept are listed under [Browser Metrics *], [Crash Reports *], [Download History *] and 8 others`. Each such block opens with the same `This entry contains content originating from N removed entries:` list as an added entry's, right under its header.

In this example, two version-specific log path entries were merged into a single entry using a wildcard to cover both:

**Old (v220510):**

```ini
[Age of Wonders II *]
Section=Games
DetectFile=%ProgramFiles%\Age of Wonders II
FileKey1=%ProgramFiles%\Age of Wonders II|*aow2Log.txt
FileKey2=%ProgramFiles%\Age of Wonders II\Resource\FX\Spell|*.lnk

[Age of Wonders: Shadow Magic *]
Section=Games
DetectFile=%ProgramFiles%\Age of Wonders II
FileKey1=%ProgramFiles%\Age of Wonders Shadow Magic|*aowsmLog.txt
FileKey2=%ProgramFiles%\Age of Wonders Shadow Magic\Resource\Maps|*.bak
```

**New (v251109):**

```ini
[Age of Wonders II *]
Section=Games
DetectFile=%ProgramFiles%\Age of Wonders II
FileKey1=%ProgramFiles%\Age of Wonders II\Resource\FX\Spell|*.lnk
FileKey2=%ProgramFiles%\Age of Wonders Shadow Magic\Resource\Maps|*.bak
FileKey3=%ProgramFiles%\Age of Wonders*|*Log.txt
```

**Diff output:**

```
     Age of Wonders: Shadow Magic * has been merged into Age of Wonders II *
[...]

     Age of Wonders II * has been modified
     This entry contains content originating from 1 removed entry:
       - Age of Wonders: Shadow Magic *

     1 key carried over from merged entries:

       Carried over 1 FileKey
             FileKey2=%ProgramFiles%\Age of Wonders Shadow Magic\Resource\Maps|*.bak (from [Age of Wonders: Shadow Magic *])

     1 key modified:

       Modified 1 FileKey

       FileKey3 has been modified, replacing 2 old keys
              + New: FileKey3=%ProgramFiles%\Age of Wonders*|*Log.txt
              - Old: FileKey1=%ProgramFiles%\Age of Wonders Shadow Magic|*aowsmLog.txt (from [Age of Wonders: Shadow Magic *])
              - Old: FileKey1=%ProgramFiles%\Age of Wonders II|*aow2Log.txt (from [Age of Wonders II *])
```

**Explanation:**

The merger itself is reported in the merged-entries list

The key-level detail appears in the modified-with-merged-content section.

Shadow Magic's `Resource\Maps` key survives unchanged as `FileKey2`, so it is listed as carried over and credited to the removed entry.

The `FileKey3` wildcard captures both games' log paths into a single key, one from the removed entry and one from this entry's own prior version.

### Renames with key changes

When an entry is renamed and also has other key changes, it does not appear in the Renamed list (which contains only name-only renames). Instead, it appears in a dedicated `Minor changes to renamed entries:` section, displayed as `Old Name * has been renamed to New Name *` alongside its key changes.

In this example, `IDLE Recent Files *` was renamed to `IDLE *` and simultaneously recategorized from `3021` to `3024`:

**Old (v220510):**

```ini
[IDLE Recent Files *]
LangSecRef=3021
Detect=HKCU\Software\Python
FileKey1=%UserProfile%\.idlerc|recent-files.lst
```

**New (v251109):**

```ini
[IDLE *]
LangSecRef=3024
Detect=HKCU\Software\Python
FileKey1=%UserProfile%\.idlerc|recent-files.lst
```

**Diff output:**

```
     IDLE Recent Files * has been renamed to IDLE *

     1 key modified:

       Modified 1 LangSecRef

       LangSecRef has been modified, replacing 1 old key
              + New: LangSecRef=3024
              - Old: LangSecRef=3021
```

**Explanation:**

The `FileKey` and `Detect` content of `[IDLE Recent Files *]` was found unchanged in `[IDLE *]`, so the entry is classified as renamed rather than removed.

`LangSecRef`'s value also changed, so the entry appears in the `Minor changes to renamed entries:` section with its key diffs, rather than in the name-only rename list.

### Cross-Entry Key Movements

A key that leaves an entry present in both files and lands in any other entry has moved. Moves are reported in a dedicated `Cross-Entry key movements:` section, grouped by the entry the keys left, each group headed with its total, eg. `MP3Gain * moved 2 keys:`. The receiving entry can be modified, renamed, or added, so splitting part of an entry out into a new one shows up here. Unlike with a merger, the source entry still exists.

The section is omitted entirely when no movements are detected.

**Diff output:**

```
 Cross-Entry key movements:
[...]

   CleanMyPC Registry Cleaner * moved 1 key:
       -> FileKey2=%WinDir%\$regcmp$|*|REMOVESELF moved to [Windows Subsystems *]
[...]

   MP3Gain * moved 2 keys:
       -> RegKey1=HKCU\Software\VB and VBA Program Settings\MP3GainAnalysis\StartUp|AddFilesPath moved to [MP3Gain StartUp *]
       -> RegKey2=HKCU\Software\VB and VBA Program Settings\MP3GainAnalysis\StartUp|AddFolderPath moved to [MP3Gain StartUp *]
[...]

   Soulseek * moved 5 keys:
       -> FileKey1=%UserProfile%\Documents\Soulseek Chat Logs\Private|*.txt moved to [Soulseek Chat History *]
       -> FileKey2=%LocalAppData%\Soulseek Chat Logs\*|*.log moved to [Soulseek Chat History *]
       -> FileKey3=%LocalAppData%\SoulseekQt\Soulseek Chat Logs|*.log|RECURSE moved to [Soulseek Chat History *]
       -> FileKey5=%ProgramFiles%\Soulseek*\Private_Chat_Logs|*.txt moved to [Soulseek Chat History *]
       -> FileKey6=%ProgramFiles%\Soulseek*\Room_Chat_Logs|*.txt moved to [Soulseek Chat History *]
[...]

   Windows Subsystems * moved 13 keys:
       -> FileKey1=%ProgramData%\Microsoft\PlayReady|*.hds moved to [Microsoft PlayReady *]
       -> FileKey2=%ProgramData%\Microsoft\PlayReady\Cache|* moved to [Microsoft PlayReady *]
       -> FileKey6=%ProgramData%\Microsoft\Windows\DRM|*.log moved to [Microsoft PlayReady *]
       -> FileKey7=%ProgramData%\Microsoft\Windows\DRM\Cache|*|RECURSE moved to [Microsoft PlayReady *]
       -> FileKey8=%ProgramData%\Microsoft\Windows\DRM\PreUpgrade|*.log moved to [Microsoft PlayReady *]
       -> FileKey11=%LocalAppData%\Microsoft\Windows\PRICache|*|RECURSE moved to [Universal Windows Platform *]
       -> FileKey14=%SystemDrive%\spoolerlogs|spooler.xml moved to [Windows Printing *]
       -> FileKey17=%WinDir%\System32\spool|spooler.xml moved to [Windows Printing *]
       -> RegKey1=HKCU\Software\Microsoft\Direct3D\MostRecentApplication moved to [Microsoft DirectX *]
       -> RegKey2=HKLM\Software\Microsoft\Direct3D\MostRecentApplication moved to [Microsoft DirectX *]
       -> RegKey3=HKLM\Software\Microsoft\DirectDraw\MostRecentApplication moved to [Microsoft DirectX *]
       -> RegKey5=HKLM\Software\Wow6432Node\Microsoft\Direct3D\MostRecentApplication moved to [Microsoft DirectX *]
       -> RegKey6=HKLM\Software\Wow6432Node\Microsoft\DirectDraw\MostRecentApplication moved to [Microsoft DirectX *]
[...]

 100 keys moved between entries (41 source entries)
```

**Explanation:**

`[CleanMyPC Registry Cleaner *]` gave its `%WinDir%\$regcmp$` key to `[Windows Subsystems *]`, an entry that already existed.

`[MP3Gain *]` and `[Soulseek *]` were each split: their recent-file and chat-log keys went into new entries of their own. The last two Soulseek keys landed in one wildcard key, `%ProgramFiles%\Soulseek*\*Chat_Logs|*.txt`, rather than identical ones.

`[Windows Subsystems *]` gave keys to four entries. `[Microsoft PlayReady *]` is a renamed `[Media Play Ready Client *]`. The other three are new, and each also absorbed removed entries, so their merger output credits these keys `(from [Windows Subsystems *])`.

Every source entry still exists in the new file, which is what distinguishes a movement from a merger.

Each move also appears in the itemized output of both entries, under `Moved` in the source and `Moved in` in the receiving entry (see [Modified Entries](#modified-entries)):

```
     CleanMyPC Registry Cleaner * has been modified

     1 key moved to other entries:

       Moved 1 FileKey
             FileKey2=%WinDir%\$regcmp$|*|REMOVESELF moved to [Windows Subsystems *]
```

## Removed Entries

An entry is **Removed** if it exists in the old file but not in the new file. Removed entries fall into one of three sub-categories.

### Renamed

The entry no longer exists by its old name, but another entry in the new file contains its content mostly unchanged. Diff reports both the old and new names.

The Renamed list contains only renames with no key changes. Renames with accompanying key changes appear separately in the `Minor changes to renamed entries:` section.

**Sample output:**

```
     3delite Filesystem Dialogs * has been renamed to 3delite Filesystem Dialogs Library *

     4Sync Extras * has been renamed to 4Sync *

     Tracks Eraser Pro * has been renamed to Acesoft Tracks Eraser Pro *
```

The list closes with a count: `N entries renamed (name-only changes)`.

### Merged

The entry no longer exists, but its content was absorbed into another entry in the new file that is substantially different from the old version. Diff reports the name of the absorbing entry. Mergers are further categorized in the summary report:

- **Modified by merger**: the absorbing entry existed in the old file and was modified to include the removed entry's content
- **Added with merger**: the absorbing entry is itself new and contains content from one or more removed entries

**Sample output (merged into an existing entry):**

```
     .NET Application History * has been merged into .NET Framework *
[...]

     Adobe Reader XI * has been merged into Adobe Acrobat Reader *
[...]

     Common Language Runtime * has been merged into .NET Framework *
```

**Sample output (merged into a new entry):**

```
     Abelssoft GoogleClean * has been merged into Abelssoft GClean *
```

In the added-entries output, the entry's header is followed by `This entry contains content originating from 1 removed entry:` and the source's name.

**Example with before/after context:**

Two companion entries (`Clover Bookmarks Backups *` and `Clover Session *`) were consolidated into the main `Clover *` entry. Each contributed a single FileKey, which the new entry combines into a single pattern.

**Old (v220510):**

```ini
[Clover *]
LangSecRef=3024
Detect=HKCU\Software\Clover
FileKey1=%LocalAppData%\Clover\User Data\Default\JumpListIcons*|*tmp
FileKey2=%LocalAppData%\Clover\User Data\temp|*.*|RECURSE

[Clover Bookmarks Backups *]
LangSecRef=3024
Detect=HKCU\Software\Clover
FileKey1=%LocalAppData%\Clover\User Data\Default|Bookmarks.bak

[Clover Session *]
LangSecRef=3024
Detect=HKCU\Software\Clover
FileKey1=%LocalAppData%\Clover\User Data\Default|Current *;Last *
```

**New (v251109):**

```ini
[Clover *]
LangSecRef=3024
Detect=HKCU\Software\Clover
FileKey1=%LocalAppData%\Clover\User Data\Default|Bookmarks.bak;Current *;Last *
FileKey2=%LocalAppData%\Clover\User Data\Default\JumpListIcons*|*tmp
FileKey3=%LocalAppData%\Clover\User Data\temp|*|RECURSE
```

**Diff output:**

```
     Clover Bookmarks Backups * has been merged into Clover *

     Clover Session * has been merged into Clover *
[...]

     Clover * has been modified
     This entry contains content originating from 2 removed entries:
       - Clover Bookmarks Backups *
       - Clover Session *

     1 key modified:

       Modified 1 FileKey

       FileKey1 has been modified, replacing 2 old keys
              + New: FileKey1=%LocalAppData%\Clover\User Data\Default|Bookmarks.bak;Current *;Last *
              - Old: FileKey1=%LocalAppData%\Clover\User Data\Default|Bookmarks.bak (from [Clover Bookmarks Backups *])
              - Old: FileKey1=%LocalAppData%\Clover\User Data\Default|Current *;Last * (from [Clover Session *])
```

The two removed entries each contributed a FileKey, and the new version of `Clover *` combines their values into a single `FileKey1`.

Because `Clover *` is a *modified* entry rather than an added one, its source entries do **not** appear in the [Source Entry Key Status Report](#source-entry-key-status-report).

**Sample output (split across multiple entries):**

```
     Adobe Reader DC * has been split/merged into 2 entries
       • Adobe Acrobat *
       • Adobe Acrobat Reader *
```

A split/merged result means the old entry's content was distributed across more than one new entry, i.e. different keys matched in different places. The keys of a split entry that none of those entries kept, unchanged or captured, are listed once, here, under `N keys not kept by any of these entries:`. The merger blocks of the entries it went into don't list them as dropped; each ends with a line pointing back here instead (see [Modification by merger](#modification-by-merger)).

**Example with before/after context (large split):** The following shows how `Alternate Services *`, a single legacy entry covering nine different Firefox-based browsers, was split into six browser-specific entries. Each browser that received a dedicated entry now appears as a separate absorbing target, and the keys for the three browsers that got no entry are listed under the split.

**Old (v220510):**

```ini
[Alternate Services *]
LangSecRef=3026
Detect1=HKCU\Software\ArtistScope\ArtisBrowser
Detect2=HKCU\Software\Classes\Local Settings\Software\Microsoft\Windows\CurrentVersion\AppModel\SystemAppData\Mozilla.Firefox_n80bbvh6b1yt2
Detect3=HKCU\Software\LibreWolf
Detect4=HKLM\Software\ComodoGroup\IceDragon
Detect5=HKLM\Software\FlashPeak\SlimBrowser
Detect6=HKLM\Software\Mozilla\Basilisk
Detect7=HKLM\Software\Mozilla\Pale Moon
Detect8=HKLM\Software\Mozilla\SeaMonkey
Detect9=HKLM\Software\Mozilla\Waterfox
DetectFile=%AppData%\Mozilla\Firefox
FileKey1=%AppData%\ArtistScope\ArtisBrowser\Profiles\*|AlternateServices.txt
FileKey2=%AppData%\Comodo\IceDragon\Profiles\*|AlternateServices.txt
FileKey3=%AppData%\FlashPeak\SlimBrowser\Profiles\*|AlternateServices.txt
FileKey4=%AppData%\LibreWolf\Profiles\*|AlternateServices.txt
FileKey5=%AppData%\Moonchild Productions\Basilisk\Profiles\*|AlternateServices.txt
FileKey6=%AppData%\Moonchild Productions\Pale Moon\Profiles\*|AlternateServices.txt
FileKey7=%AppData%\Mozilla\Firefox\Profiles\*|AlternateServices.txt
FileKey8=%AppData%\Mozilla\SeaMonkey\Profiles\*|AlternateServices.txt
FileKey9=%AppData%\Waterfox\Profiles\*|AlternateServices.txt
FileKey10=%LocalAppData%\Packages\Mozilla.Firefox_*\LocalCache\Roaming\Mozilla\Firefox\Profiles\*|AlternateServices.txt
```

**Diff output (220510→251109):**

```
     Alternate Services * has been split/merged into 6 entries
       • ArtisBrowser Caches *
       • LibreWolf Caches *
       • Mozilla Firefox Caches *
       • Pale Moon Caches *
       • SeaMonkey Caches *
       • Waterfox Caches *
   
     14 keys not kept by any of these entries:

       Lost 1 LangSecRef
       Lost 9 Detects
       Lost 1 DetectFile
       Lost 3 FileKeys
             LangSecRef=3026
             Detect1=HKCU\Software\ArtistScope\ArtisBrowser
             Detect2=HKCU\Software\Classes\Local Settings\Software\Microsoft\Windows\CurrentVersion\AppModel\SystemAppData\Mozilla.Firefox_n80bbvh6b1yt2
             Detect3=HKCU\Software\LibreWolf
             Detect4=HKLM\Software\ComodoGroup\IceDragon
             Detect5=HKLM\Software\FlashPeak\SlimBrowser
             Detect6=HKLM\Software\Mozilla\Basilisk
             Detect7=HKLM\Software\Mozilla\Pale Moon
             Detect8=HKLM\Software\Mozilla\SeaMonkey
             Detect9=HKLM\Software\Mozilla\Waterfox
             DetectFile=%AppData%\Mozilla\Firefox
             FileKey2=%AppData%\Comodo\IceDragon\Profiles\*|AlternateServices.txt
             FileKey3=%AppData%\FlashPeak\SlimBrowser\Profiles\*|AlternateServices.txt
             FileKey5=%AppData%\Moonchild Productions\Basilisk\Profiles\*|AlternateServices.txt
```

The new entries detect their browsers with other keys (`LibreWolf Caches *`, for one, uses `DetectFile1=%AppData%\librewolf\Profiles`), so none of the old Detects survive verbatim. They are categorized by `Section=` rather than `LangSecRef=3026`, so the old category key is lost too.

Six of the nine browser FileKeys were matched in six new per-browser entries.

The remaining three (IceDragon, SlimBrowser, Basilisk) were not matched in any new entry.

### Removed Without Replacement

The entry is gone and its content does not appear in any other entry in the new file. These are listed individually under the `Entries removed without replacement:` header, closed by a count.

**Sample output:**

```
 Entries removed without replacement:

     2K Launcher * has been removed

     337 Wallpaper * has been removed

     3D Builder * has been removed
[...]

 - 526 removed without replacement
```

---

# How Classification Works

This section describes how Diff decides whether a removed entry was renamed, merged, or removed without replacement.

## Value Normalization

Before any comparison happens, Diff rewrites a set of deprecated values in the old file to their modern equivalents. This prevents a database-wide style change from being reported as thousands of modifications.

| Deprecated | Rewritten to |
|:-|:-|
| `*.*` | `*` |
| `%CommonAppData%` | `%ProgramData%` |
| `%LocalLowAppData%` | `%UserProfile%\AppData\LocalLow` |
| `%Documents%` | `%UserProfile%\Documents` |
| `%Pictures%` | `%UserProfile%\Pictures` |
| `%Music%` | `%UserProfile%\Music` |
| `%Videos%` | `%UserProfile%\Videos` |

## Candidate Selection

When an entry is absent from the new file, Diff does not compare it against every entry in the new file. It first gathers a candidate pool using four heuristics:

- name and browser-reference similarity against the added/modified snapshot
- an exact key-value reverse index
- a path-root reverse index
- a wildcard-prefix reverse index

The resulting names are then filtered to those that are eligible: entries that were added to the new file, or that exist in both files with modifications. Only entries satisfying both steps are considered potential rename or merge targets.

## Matching Basis

Diff scores candidates by comparing `FileKey` and `RegKey` values between the old entry and each candidate. Other key types are used to identify candidate entries but do not contribute to the match score. Key comparison uses wildcard pattern matching to capture path abstractions.

## Vague Paths

A set of locations is considered too generic to establish a content match on its own. Matching on them alone would link unrelated entries that happen to touch the same common directory. These include `%AppData%`, `%LocalAppData%`, `%UserProfile%`, `%Documents%`, `%WinDir%`, `%WinDir%\System32`, `%SystemDrive%`, `%Public%`, `%Pictures%`, `%Music%`, `%Video%`, `%UserProfile%\Desktop`, and several broad registry roots such as `HKCU\Software\Microsoft\Windows` and `HKLM\Software\Microsoft\Windows`.

## Classification Rules

| Outcome | Criteria |
|:-|:-|
| **Renamed** | The candidate's name is new to the database (an added entry, not a modified one); all FileKeys and RegKeys from the old entry are matched in the candidate; the match counts are identical; no wildcard reduction occurred (a key gaining patterns is still a rename, with the change itemized) |
| **Merged** | At least one FileKey or RegKey from the old entry is matched in the candidate, but the rename criteria are not met. This includes the case where every key matched but the candidate already existed in the old file (the old entry was absorbed into it) |
| **Removed without replacement** | No FileKeys or RegKeys from the old entry are matched in any candidate|


---

# Output and Log

## Output Structure

A diff run produces output in the following order. Each section opens with a descriptive header line and closes with a count line.

| # | Section | Header line | Closing count line |
|:-|:-|:-|:-|
| 1 | Run header | `Beginning Diff`, `Diff: version XXXXXX -> version XXXXXX` | |
| 2 | Browser categories new to the database, then any no longer in it | `Web Browser additions`, `Web Browser removals` | `N web browsers added`, `N web browsers removed` |
| 3 | Entries removed without replacement | `Entries removed without replacement:` | `- N removed without replacement` |
| 4 | Name-only renames, old name → new name | `Entries renamed without other changes:` | `N entries renamed (name-only changes)` |
| 5 | Mergers, including split/merged entries | `Entries merged or split into other entries:` | `N entries merged or split into other entries` |
| 6 | Key diffs for renamed entries | `Minor changes to renamed entries:` | `Minor changes to N renamed entries` |
| 7 | Modified entries that absorbed merged content, with per-entry source notes | `Modified entries incorporating merged content:` | `N modified entries incorporating merged content` |
| 8 | Keys relocated from entries that still exist into other entries | `Cross-Entry key movements:` | `N keys moved between entries (M source entries)` |
| 9 | Ordinary modifications | `Modified entries:` | `N modified entries` |
| 10 | Consolidating additions with per-key tracking | `Added entries containing merged content:` | `N added entries consolidating removed content` |
| 11 | Novel additions | `Added entries:` | `N novel entries added` |
| 12 | Per-key capture status for merged sources | `SOURCE ENTRY KEY STATUS REPORT:` | `KEY STATUS SUMMARY:` table |
| 13 | Timing and aggregate counts | `Total diff time:`, `Diff Summary` | `Diff complete` |

Each browser list in section 2 is omitted when it would be empty, and so is section 8 when no cross-entry movements were detected. Section 12 is written only to the saved log; use log saving (`-savelog`) or the Log Viewer to see it. A file without a `; Version:` line reads as `version not given` in the run header.

When a run finds nothing at all, the Diff Summary is the single line `No changes detected across N entries`, counting the new file's entries. The log still carries an empty `Entries removed without replacement:` section closed by `- 0 removed without replacement`.

## Console Output vs. Saved Log

The interactive console and the saved log render the same data in two different arrangements. The table above describes the **saved log**.

In the console:

- The count line is printed as a leading header above its list, not below it.
- The descriptive header lines (`Entries removed without replacement:`, `Modified entries:`, etc) do not appear at all
- Split/merge blocks gain a `Merged into:` label above the bullet list
- The Source Entry Key Status Report is not shown at all.

## Sample Output

The following is a trimmed extract of the saved log for a real diff run comparing base-flavor winapp2.ini version 220510 (May 2022) against version 251109 (November 2025). `[...]` marks sections cut for brevity.

```
Beginning Diff
Diff: version 220510 -> version 251109

 Web Browser additions

   .360 Secure Browser Web Browser

   AOL Shield Web Browser
[...]

 49 web browsers added

 Entries removed without replacement:

     2K Launcher * has been removed

     337 Wallpaper * has been removed
[...]

 - 526 removed without replacement

 Entries renamed without other changes:

     3delite Filesystem Dialogs * has been renamed to 3delite Filesystem Dialogs Library *

     4Sync Extras * has been renamed to 4Sync *
[...]

 112 entries renamed (name-only changes)

 Entries merged or split into other entries:

     .NET Application History * has been merged into .NET Framework *

     .NET Framework Isolated Storage * has been merged into .NET Framework *

     3D Viewer * has been merged into Microsoft 3D Viewer *
[...]

 547 entries merged or split into other entries

 Minor changes to renamed entries:

     Abelssoft SSD Fresh Backups * has been renamed to Abelssoft SSD Fresh *

     1 key modified:

       Modified 1 FileKey

       FileKey1 has been modified, replacing 1 old key
              + New: FileKey1=%LocalAppData%\Abelssoft\SSD Fresh|backup.*;*.log
              - Old: FileKey1=%LocalAppData%\Abelssoft\SSD Fresh|backup.*
[...]

 Minor changes to 42 renamed entries

 Modified entries incorporating merged content:

     .NET Framework * has been modified
     This entry contains content originating from 5 removed entries:
       - .NET Application History *
       - .NET Framework Isolated Storage *
       - Common Language Runtime *
       - Font Cache *
       - Presentation Foundation *
[...]

 135 modified entries incorporating merged content

 Cross-Entry key movements:

   Adobe Acrobat Distiller * moved 1 key:
       -> FileKey1=%AppData%\Adobe\Acrobat\Distiller*|*.log moved to [Adobe *]
[...]

 100 keys moved between entries (41 source entries)

 Modified entries:

     //N.P.P.D. RUSH// - The Milk of Ultra Violet * has been modified

     1 key modified:

       Modified 1 Detection criteria

       Detection criteria modified
              + New: DetectFile=%LocalAppData%\NPPDRUSH
              - Old: Detect=HKCU\Software\Valve\Steam\Apps\270090
[...]

 325 modified entries

 Added entries containing merged content:

   Abelssoft GClean * has been added
   This entry contains content originating from 1 removed entry:
     - Abelssoft GoogleClean *

   1 key added:

       Added 1 FileKey
             FileKey2=%LocalAppData%\Abelssoft\GClean\Backups|* (novel)

   3 keys carried over from merged entries:

       Carried over 1 LangSecRef
       Carried over 2 FileKeys
             LangSecRef=3024 (from [Abelssoft GoogleClean *])
[...]

 316 added entries consolidating removed content

 Added entries:

     .NET Platform * has been added

     360 Secure Browser Autofill Data & Search Engine Preferences * has been added
[...]

 1215 novel entries added
 SOURCE ENTRY KEY STATUS REPORT:
[...]

 Total diff time: 00:00:03.2057302

 Diff Summary
   Net entry count change: 3257 → 3715 (+458)
    + 49 new browsers added
   Modified entries: 460
    + 486 added keys across 196 entries
    - 344 removed keys without replacement across 163 entries
    ~ 388 updated keys replaced 603 old keys across 220 entries
    ~ 100 keys moved from 41 entries into 24 entries
       ~ 2 keys moved into 1 modified entry
       & 5 keys moved into 1 renamed entry
       + 93 keys moved into 22 added entries
    + 135 entries also received merged content from removed entries (see merged entries below)
       = 118 entries contain 365 keys carried over unchanged from merged sources
   Removed entries: 1227
    @ 547 removed entries have been merged into other entries
       @ 175 merged into 135 modified entries
       + 382 merged into 316 added entries
       - 77 split entries lost 1481 keys that no entry kept
    & 154 removed entries have been renamed
       = 112 are name-only changes (no key differences)
       + 25 added keys across 12 entries
       - 23 removed keys across 21 entries
       ~ 66 updated keys replaced 78 old keys across 30 entries
    - 526 entries have been removed without replacement
   Added entries: 1685
    @ 316 entries consolidate content from 382 removed entries
       + 174 entries contain 1078 novel keys (not from merged sources)
       = 222 entries contain 1246 keys carried over unchanged from merged sources
       ~ 258 entries contain 935 keys capturing 1278 removed keys
       - 84 entries dropped 259 keys from merged sources
    + 1215 novel entries (without merged content)
    & 154 added entries are renamed versions of removed entries and may contain other minor changes

Diff complete
```

## Interpreting the Summary

The **Diff Summary** at the end of each run condenses all classifications into aggregate counts.

**`Net entry count change: OLD → NEW (+N)`**

The entry counts of the old and new files, and the difference between them.

**`+ N new browsers added`**

The number of new browser categories detected among the added entries (itemized in the `Web Browser additions` section at the top of the output).

**`- N browsers removed`**

The number of browser categories in the old file that no longer appear in the new one (itemized in the `Web Browser removals` section). Like the line above, it is shown only when nonzero.

---

**`Modified entries: N`**
Entries present in both files (by name) that changed in any way.

- `+ N added keys across N entries`: new keys added to existing entries. Keys moved in from another entry are counted on the moved line instead, and keys taken over unchanged from merged removed entries on the carried-over line
- `- N removed keys without replacement across N entries`: keys the entry's own old version had, with no equivalent in the new entry. Keys that only a merged removed entry had are not counted here
- `~ N updated keys replaced M old keys across N entries`: keys whose values changed; one new key captured one or more old keys via wildcard match or content consolidation
- `~ N keys moved from X entries into Y entries`: cross-entry key movements; only shown when movements were detected. The X entries are always modified or renamed ones. A moved key isn't also counted as removed. Bullets split the moves by the kind of entry they landed in, each shown only when nonzero:
  - `~ N keys moved into N modified entries`: entries present in both files under the same name
  - `& N keys moved into N renamed entries`
  - `+ N keys moved into N added entries`: new entries, whether novel or consolidating removed entries. Splitting part of an entry out into a new one shows up here
- `+ N entries also received merged content from removed entries`: how many of the modified entries absorbed removed entries; one bullet, shown only when nonzero:
  - `= N entries contain N keys carried over unchanged from merged sources`: keys the entries took in verbatim from a merged source and that their own old version didn't have, as their `carried over from merged entries` groups list them

---

**`Removed entries: N`**
Entries present in the old file but absent from the new file.

- `@ N removed entries have been merged into other entries`: total mergers, categorized:
  - `@ N merged into N modified entries`: absorbed by entries that already existed in the old file
  - `+ N merged into N added entries`: absorbed by entries new to the new file
  - `- N split entries lost N keys that no entry kept`: keys of entries split across several new entries that none of them has, unchanged or captured, as each split entry's `not kept by any of these entries` list shows them
- `& N removed entries have been renamed`: classified as renames; bullets counting how many were name-only vs. included minor changes
- `- N entries have been removed without replacement`: no matching content found in the new file

> The two merger sub-counts can add up to more than their headline. An entry that was split across both a modified target and an added target is counted in each. In the sample above, 175 + 382 = 557 against a headline of 547, meaning 10 removed entries contributed content to both kinds of target. The same applies to the added-entries bullets, where one entry can appear in several categories at once.

---

**`Added entries: N`**
Entries present in the new file but absent from the old file.

- `@ N entries consolidate content from N removed entries`: new entries that absorbed at least one removed entry; bullets:
  - `+ N entries contain N novel keys`: keys with no equivalent in any source entry, leaving out keys moved in from an entry that still exists
  - `= N entries contain N keys carried over unchanged`: keys transferred verbatim from source entries, as the entries' `carried over from merged entries` groups list them
  - `~ N entries contain N keys capturing N removed keys`: keys that matched and replaced one or more old keys, and the old keys they replaced, as the entries' blocks list them
  - `- N entries dropped N keys from merged sources`: the keys the entries' `not in this entry` blocks list. A split source's lost keys are counted under `Removed entries` instead
- `+ N novel entries (without merged content)`: entirely new entries with no connection to any removed entry
- `& N added entries are renamed versions of removed entries and may contain other minor changes`: the same rename set as the removed-side `& N removed entries have been renamed` line; the two counts always match

---

When **log saving** is enabled, the diff output is also written to disk (`diff.txt` by default). Use **Log Viewer** in the menu to review the most recent run's output without re-running.

---

## Source Entry Key Status Report

When any mergers into *added* entries are present in the diff, a **Source Entry Key Status Report** is written to the saved log between the Added entries section and the Diff Summary. It does not appear in the interactive console output. Use log saving (`-savelog`) or the Log Viewer to see it.

> **Scope:** the report covers only source entries absorbed by **added** entries. Entries merged into *modified* entries, the `@ N merged into N modified entries` line of the summary, are not included. In the sample run, that means 382 of the 547 merged entries appear here and 175 do not.

The report shows, for each covered source entry, the full capture status of every key: whether each key was found in the absorbing entry (✓ captured) or absent from it (✗ dropped). Keys are tagged by role:

- `[DELETION]`: `FileKey` or `RegKey`
- `[DETECTION]`: `Detect` or `DetectFile`
- `[CATEGORY]`: `Section` or `LangSecRef`
- `[OTHER]`: `Warning` and any other key type

Each entry block shows a header with total captured and dropped counts, followed by one line per key. The report closes with a `KEY STATUS SUMMARY:` table aggregating counts and capture rates across all four categories.

**Example (Cook'n consolidation):**

```
 [Cook'n Cache *] - 3 captured, 0 dropped
     ✓ [CATEGORY] LangSecRef=3021
     ✓ [DETECTION] Detect=HKLM\Software\Microsoft\Windows\CurrentVersion\Uninstall\Cook'n
     ✓ [DELETION] FileKey1=%AppData%\Mozilla\eclipse\Cache|*

 [Cook'n Dups *] - 3 captured, 1 dropped
     ✓ [CATEGORY] LangSecRef=3021
     ✓ [DETECTION] Detect=HKLM\Software\Microsoft\Windows\CurrentVersion\Uninstall\Cook'n
     ✗ [OTHER] Warning=This removes identical duplicates of the huge Getting Started Guide.
     ✓ [DELETION] FileKey1=%LocalAppData%\DVO\Cook'n10App\plugins|Getting Started Guide.rtf|RECURSE

 [Cook'n Logs *] - 4 captured, 0 dropped
     ✓ [CATEGORY] LangSecRef=3021
     ✓ [DETECTION] Detect=HKLM\Software\Microsoft\Windows\CurrentVersion\Uninstall\Cook'n
     ✓ [DELETION] FileKey1=%UserProfile%\Documents\Cook'n*|*.log|RECURSE
     ✓ [DELETION] FileKey2=%LocalAppData%\DVO\Cook'n*App|*.log|RECURSE
```

The `Warning=` key from `Cook'n Dups *` is the only key across all three source entries that did not survive, shown here as ✗ `[OTHER]`.

The same key appeared in the `1 key from merged entries not in this entry:` block of the `Cook'n *` Added section output.


After all per-entry blocks, a summary is shown:

```
 KEY STATUS SUMMARY:                 Total  Captured  Dropped  Capture Rate
 DELETION  (FileKey/RegKey):          2511      1825      686         72.7%
 DETECTION (Detect/DetectFile):       1251       273      978         21.8%
 CATEGORY  (Section/LangSecRef):       382       164      218         42.9%
 OTHER     (Warning etc.):              92        18       74         19.6%
 All keys:                            4236      2280     1956         53.8%
```

---

# Command-Line Arguments

Diff's CLI resets all Diff settings to their defaults before applying arguments, so a command-line run never inherits settings saved from the menu. `-d` and `-donttrim` always turn their setting off, so `-offline -d` still compares two local files. `-savelog` and `-verbose` flip their setting from its **default** value, not from whatever is on disk.

## Toggles

| Arg | Effect |
|:-|:-|
| `-d` | Disable downloading (compare two local files) |
| `-donttrim` | Disable trimming the downloaded file (downloading remains active) |
| `-savelog` | Enable saving the diff output to disk |
| `-verbose` | Print the full text of changed entries alongside each change |

## File Selection

| Arg | Effect | Default |
|:-|:-|:-|
| `-1d path` | Set the older/local file directory | Current directory |
| `-1f name` | Set the older/local file name | `winapp2.ini` |
| `-2d path` | Set the newer file directory (local mode only) | Current directory |
| `-2f name` | Set the newer file name (local mode only) | None |
| `-3d path` | Set the log save directory | Current directory |
| `-3f name` | Set the log save file name | `diff.txt` |
| `-4d path`, `-summaryd path` | Set the outcome summary directory | Current directory |
| `-4f name`, `-summaryf name` | Set the outcome summary file name. saved only when named | None (no summary written) |

With `-d` and no `-2f`, there is no newer file to against which to diff and the run exits silently without producing any output. Always pair `-d` with `-2f`.

The summary file is saved if and only if it is assigned a name. It is intended for the build script.

## Flavor Selection

These are global winapp2ool arguments, not Diff-specific ones. They select which published build of winapp2.ini gets downloaded as the "newer" file, and have no effect when `-d` is used. They override the saved flavor for the current run only.

| Arg | Flavor |
|:-|:-|
| `-base`, `-ncc` | Base (non-CCleaner) |
| `-ccleaner`, `-cc` | CCleaner (**default**) |
| `-bleachbit`, `-bb` | BleachBit |
| `-systemninja`, `-sn` | System Ninja |
| `-tron` | Tron |
| `-ccleaner7`, `-cc7` | CCleaner 7 |
| `-fluentcleaner`, `-fc` | FluentCleaner |

## Examples

| Command | Effect |
|:-|:-|
| `winapp2ool -diff` | Diff local `winapp2.ini` against the latest CCleaner-flavor build from GitHub (trimmed) |
| `winapp2ool -diff -base` | Diff local `winapp2.ini` against the latest base (non-CCleaner) build from GitHub (trimmed) |
| `winapp2ool -diff -donttrim` | Diff against GitHub without trimming first |
| `winapp2ool -diff -savelog` | Diff against GitHub and save the output to `diff.txt` |
| `winapp2ool -diff -verbose` | Diff against GitHub, printing full entry text alongside each change |
| `winapp2ool -diff -d -1f old.ini -2f new.ini` | Compare two specifically named local files |
| `winapp2ool -diff -d -savelog -2f new.ini` | Compare two local files and save the output |
| `winapp2ool -diff -d -2f new.ini -4f summary.txt` | Compare two local files and write only the machine-readable outcome |

---

# Tips & Best Practices

### Trimming

Trimming (enabled by default when downloading) runs Trim on the downloaded winapp2.ini before diffing, removing entries for software not installed on the current system.

This produces a diff showing only changes relevant to installed software, but it is only meaningful when your local copy is already trimmed. Diffing a trimmed vs untrimmed file is noisy.

### Scripting

Diff can be driven fully from the command line in silent mode for automated changelog generation:

```
winapp2ool -s -diff -donttrim -savelog -3f changelog.txt
```

In silent mode nothing is written to the console; the log file is the only output. Diff is used this way in the winapp2.ini build pipeline, which generates a changelog for every flavor it publishes.

Always invoke the flavor explicitly in any automated script, the CLI resets Diff's own settings but reads the flavor from the saved global configuration unless a flavor flag is passed.

### Reading the outcome from a script

The changelog is written for people. The summary is written for machines. `-4f` writes the same run's counts as `key=value` lines:

```
winapp2ool -s -offline -diff -savelog -1f old.ini -2f new.ini -3f diff.txt -4f summary.txt
```

```
haschanges=false
oldentrycount=4058
newentrycount=4058
addedentries=0
removedentries=0
modifiedentries=0
renamedentries=0
mergedentries=0
addedkeys=0
removedkeys=0
updatedkeys=0
movedkeys=0
```

Every value is an integer or `true`/`false`, so PowerShell can read the whole file with `ConvertFrom-StringData (Get-Content summary.txt -Raw)`. `haschanges` is true when any of the counts is nonzero or the entry counts differ.

`haschanges=false` does not mean the two files are identical. It means that Diff didn't detect a semantic difference about which it cares.

---

# Troubleshooting

| Symptom | Cause |
|:-|:-|
| `Run (default)` is displayed in red, and selecting it shows "Please select a file against which to diff" | No "newer" file is selected and downloading is disabled |
| A CLI run with `-d` produces no output at all | No `-2f` was given, so there is no newer file; the run exits silently |
| "winapp2.ini was empty or not found" | The older/local file (or the selected newer file) does not exist or contains no entries |
| Download fails | No internet connection, GitHub is unavailable, or the selected flavor has not been published yet |
| An old key in the log does not match the archived file's text | Deprecated values are rewritten before comparison. See [Value Normalization](#value-normalization) |
| Log Viewer is not available | Diff has not been run yet during the current session |
| "Diff complete, but diff.txt was not saved" | Log saving is on, but the save target couldn't be written. Pick a writable folder with `Choose save target` |

---

# Usage Examples

###### Note: All outputs below are lifted from real tool runs. Examples 2–5 compare base-flavor archives extracted from this repository (`Non-CCleaner/Winapp2.ini` at commit `f5ea7371` for v220510, and at tags `v251002` / `v251109`). Example 1 is a live trimmed run and depends on the software installed on the machine performing the diff. `[...]` marks lines cut for brevity.

## Example 1: Default Workflow

**Context:** Your local `winapp2.ini` is a trimmed copy from May 2022 (version 220510) and you want to see what has changed since.

**Command:**

```
winapp2ool -diff -base
```

**Output:**

```
Beginning Diff
Diff: version 220510 -> version 260915

 Web Browser additions

   Chromium Web Browser

   Google Chrome Web Browser

   Microsoft Edge Web Browser

   Mozilla Firefox Web Browser

   Opera Web Browser

 5 web browsers added

 Entries removed without replacement:

     Accounts Control * has been removed

     ActiveSync * has been removed

     ART * has been removed
[...]

 - 167 removed without replacement

 Entries renamed without other changes:

     MS Visual Studio Backups * has been renamed to Microsoft Visual Studio Backups *

     MS Visual Studio Installation Packages * has been renamed to Microsoft Visual Studio Installation Packages *
[...]

     RegEdit * has been renamed to Windows Registry Editor *

     Recent Colors History * has been renamed to Windows Shell - Theme Colors History *

 10 entries renamed (name-only changes)

 Entries merged or split into other entries:

     .NET Framework Isolated Storage * has been merged into .NET Framework *

     Active Setup Temp Folder * has been merged into Windows Subsystems *
[...]

 205 entries merged or split into other entries
[...]

 Cross-Entry key movements:

   ASUS * moved 1 key:
       -> FileKey2=%ProgramData%\ASUS\*|*.log|RECURSE moved to [ASUS USB Charger Plus *]
[...]

 35 keys moved between entries (6 source entries)

 Modified entries:
[...]

     Discord * has been modified

     5 keys added:

       Added 5 FileKeys
             FileKey1=%AppData%\BetterDiscord|*.log
             FileKey2=%AppData%\BetterDiscord\temp|*|REMOVESELF
             FileKey6=%AppData%\Discord*\Crashpad\reports|*
             FileKey8=%AppData%\Discord*\module_data\crashlogs|*
             FileKey12=%LocalAppData%\Discord*\packages|*.nupkg

     1 key removed:

       Removed 1 FileKey
             FileKey10=%LocalAppData%\SquirrelTemp|*|REMOVESELF

     2 keys modified:

       Modified 2 FileKeys
[...]

 53 modified entries

 Added entries containing merged content:
[...]

   Chromium Caches * has been added
   This entry contains content originating from 13 removed entries:
     - Application Cache Extras *
     - Blob Storage *
     - Budget Database *
     - Code Cache *
[...]

 Total diff time: 00:00:00.6979621

 Diff Summary
   Net entry count change: 492 → 432 (-60)
    + 5 new browsers added
   Modified entries: 68
    + 317 added keys across 52 entries
    - 39 removed keys without replacement across 26 entries
    ~ 108 updated keys replaced 168 old keys across 41 entries
    ~ 35 keys moved from 6 entries into 8 entries
       + 35 keys moved into 8 added entries
    + 15 entries also received merged content from removed entries (see merged entries below)
       = 14 entries contain 55 keys carried over unchanged from merged sources
   Removed entries: 399
    @ 205 removed entries have been merged into other entries
       @ 31 merged into 15 modified entries
       + 181 merged into 102 added entries
       - 49 split entries lost 1379 keys that no entry kept
    & 27 removed entries have been renamed
       = 10 are name-only changes (no key differences)
       + 21 added keys across 4 entries
       - 10 removed keys across 10 entries
       ~ 26 updated keys replaced 31 old keys across 11 entries
    - 167 entries have been removed without replacement
   Added entries: 339
    @ 102 entries consolidate content from 181 removed entries
       + 73 entries contain 641 novel keys (not from merged sources)
       = 77 entries contain 491 keys carried over unchanged from merged sources
       ~ 82 entries contain 356 keys capturing 509 removed keys
       - 45 entries dropped 395 keys from merged sources
    + 210 novel entries (without merged content)
    & 27 added entries are renamed versions of removed entries and may contain other minor changes

Diff complete
```

**Explanation:**

- `-base` selects the base flavor 
- Diff downloaded the latest winapp2.ini and trimmed it for the current system before comparing
- Only entries for software detected on this machine appear in the output
- The header names the version that was current on GitHub at the time of the run, here 260915. The trimmed file keeps its version line
- `Net entry count change: 492 → 432 (-60)`: the new file has fewer *total* entries for this machine
- Many entries were consolidated: 205 of the 399 removed entries were mergers
- All 35 moved keys landed in added entries, eg. `AVG *` gave 19 keys to the new `AVG AntiVirus *`

**Notes:** The exact output and counts depend on what software is installed. Running with `-donttrim` shows the full database diff regardless of installed software; see Example 2.

---

## Example 2: Full Database Comparison

**Context:** You want to see all changes across the entire database, not just entries relevant to your system.

**Command:**

```
winapp2ool -diff -d -1f winapp2-220510.ini -2f winapp2-251109.ini
```

**Output:**

```
Beginning Diff
Diff: version 220510 -> version 251109

 Web Browser additions

   .360 Secure Browser Web Browser
[...]

 49 web browsers added

 Entries removed without replacement:

     2K Launcher * has been removed
[...]

 - 526 removed without replacement

 Entries renamed without other changes:

     3delite Filesystem Dialogs * has been renamed to 3delite Filesystem Dialogs Library *
[...]

 112 entries renamed (name-only changes)

 Entries merged or split into other entries:
[...]

     Abelssoft GoogleClean * has been merged into Abelssoft GClean *
[...]

     Adobe Reader DC * has been split/merged into 2 entries
       • Adobe Acrobat *
       • Adobe Acrobat Reader *
[...]

 547 entries merged or split into other entries
[...]

 Total diff time: 00:00:03.2057302

 Diff Summary
   Net entry count change: 3257 → 3715 (+458)
    + 49 new browsers added
   Modified entries: 460
    + 486 added keys across 196 entries
    - 344 removed keys without replacement across 163 entries
    ~ 388 updated keys replaced 603 old keys across 220 entries
    ~ 100 keys moved from 41 entries into 24 entries
       ~ 2 keys moved into 1 modified entry
       & 5 keys moved into 1 renamed entry
       + 93 keys moved into 22 added entries
    + 135 entries also received merged content from removed entries (see merged entries below)
       = 118 entries contain 365 keys carried over unchanged from merged sources
   Removed entries: 1227
    @ 547 removed entries have been merged into other entries
       @ 175 merged into 135 modified entries
       + 382 merged into 316 added entries
       - 77 split entries lost 1481 keys that no entry kept
    & 154 removed entries have been renamed
       = 112 are name-only changes (no key differences)
       + 25 added keys across 12 entries
       - 23 removed keys across 21 entries
       ~ 66 updated keys replaced 78 old keys across 30 entries
    - 526 entries have been removed without replacement
   Added entries: 1685
    @ 316 entries consolidate content from 382 removed entries
       + 174 entries contain 1078 novel keys (not from merged sources)
       = 222 entries contain 1246 keys carried over unchanged from merged sources
       ~ 258 entries contain 935 keys capturing 1278 removed keys
       - 84 entries dropped 259 keys from merged sources
    + 1215 novel entries (without merged content)
    & 154 added entries are renamed versions of removed entries and may contain other minor changes

Diff complete
```

**Explanation:**

- Across three and a half years (May 2022 -> November 2025), 1227 entries were removed (526 without replacement, 154 renamed, and 547 merged into other entries), while 1685 new entries were added
- `Net entry count change: 3257 → 3715 (+458)`: the new database has 458 more entries than the old one

---

## Example 3: Comparing Archived Versions

**Context:** You keep an archive of winapp2.ini releases and want a month-over-month changelog between two of them, without downloading anything.

**Command:**

```
winapp2ool -diff -d -1d C:\archive -1f winapp2-251002.ini -2d C:\archive -2f winapp2-251109.ini
```

**Output:**

```
Beginning Diff
Diff: version 251002 -> version 251109

 Web Browser additions

   BriskBard Web Browser

   Ghost Web Browser

   Helium Web Browser

   Lunascape Phoebe Web Browser

   Midori Web Browser

   Mullvad Web Browser

   Norton Neo Web Browser

   Wavebox Web Browser

 8 web browsers added

 Entries removed without replacement:

     Abylon Protection Manager * has been removed
[...]

 Total diff time: 00:00:00.2155657

 Diff Summary
   Net entry count change: 3519 → 3715 (+196)
    + 8 new browsers added
   Modified entries: 81
    + 150 added keys across 55 entries
    - 242 removed keys without replacement across 17 entries
    ~ 53 updated keys replaced 54 old keys across 33 entries
    ~ 18 keys moved from 6 entries into 7 entries
       ~ 1 key moved into 1 modified entry
       + 17 keys moved into 6 added entries
    + 1 entry also received merged content from removed entries (see merged entries below)
       = 1 entry contains 1 key carried over unchanged from merged sources
   Removed entries: 313
    @ 17 removed entries have been merged into other entries
       @ 1 merged into 1 modified entry
       + 17 merged into 34 added entries
       - 4 split entries lost 29 keys that no entry kept
    & 88 removed entries have been renamed
       = 83 are name-only changes (no key differences)
       + 13 added keys across 2 entries
       ~ 5 updated keys replaced 5 old keys across 4 entries
    - 208 entries have been removed without replacement
   Added entries: 509
    @ 34 entries consolidate content from 17 removed entries
       + 29 entries contain 225 novel keys (not from merged sources)
       = 34 entries contain 132 keys carried over unchanged from merged sources
       ~ 23 entries contain 39 keys capturing 39 removed keys
       - 3 entries dropped 5 keys from merged sources
    + 387 novel entries (without merged content)
    & 88 added entries are renamed versions of removed entries and may contain other minor changes

Diff complete
```

**Explanation:**

- `-d` disables downloading; both files must exist locally
- `-1d`/`-1f` set the "old" file's directory and name; `-2d`/`-2f` set the "new" file's. Here, both live in `C:\archive` rather than the working directory
- The rename block omits a `- N removed keys` line entirely, because no renamed entry lost a key

---

## Example 4: Verbose Mode

**Context:** The default output shows only the changed keys of each entry. You want the full text of the affected entries in context.

**Command:**

```
winapp2ool -diff -d -1f winapp2-251002.ini -2f winapp2-251109.ini -verbose
```

**Output:**

A renamed entry prints both its old and new full text:

```
     360 Secure Browser Autofill Data * has been renamed to 360 Secure Browser Autofill Data & Search Engine Preferences *


       Old entry:
           [360 Secure Browser Autofill Data *]
           Section=.360 Secure Browser Web Browser
           DetectFile=%AppData%\360se6\User Data
           FileKey1=%AppData%\360se6\User Data\*|*Web Data
           FileKey2=%AppData%\360se6\User Data\*\AutoFill*|*|REMOVESELF
           FileKey3=%AppData%\360se6\User Data\AutoFill*|*|REMOVESELF


   Renamed entry:
           [360 Secure Browser Autofill Data & Search Engine Preferences *]
           Section=.360 Secure Browser Web Browser
           DetectFile=%AppData%\360se6\User Data
           FileKey1=%AppData%\360se6\User Data\*|*Web Data
           FileKey2=%AppData%\360se6\User Data\*\AutoFill*|*|REMOVESELF
           FileKey3=%AppData%\360se6\User Data\AutoFill*|*|REMOVESELF
```

A modified entry prints its new full text beneath the header, followed by the key diff:

```
     4Team Sync2 * has been modified
           [4Team Sync2 *]
           LangSecRef=3022
           DetectFile=%AppData%\4Team\Sync2
           FileKey1=%AppData%\4Team\Sync2|*.log;*.err
           FileKey2=%AppData%\4Team\Sync2\Logs|*
           FileKey3=%LocalAppData%\4Team\Sync2Cloud|*.log


     2 keys added:

       Added 2 FileKeys
             FileKey2=%AppData%\4Team\Sync2\Logs|*
             FileKey3=%LocalAppData%\4Team\Sync2Cloud|*.log

     1 key modified:

       Modified 1 FileKey

       FileKey1 has been modified, replacing 1 old key
              + New: FileKey1=%AppData%\4Team\Sync2|*.log;*.err
              - Old: FileKey1=%AppData%\4Team\Sync2|*.log|RECURSE
```

**Explanation:**

- `-verbose` (or the `Toggle verbose mode` menu option) prints the full text of each changed entry alongside its change line: removed entries show their old text, renamed and split/merged entries show both old and new text, and modified entries show their new text
- Verbose output is substantially larger

---

## Example 5: Saving the Log

**Context:** You want to archive the diff output so you can read the Source Entry Key Status Report.

**Command:**

```
winapp2ool -diff -d -1f winapp2-220510.ini -2f winapp2-251109.ini -savelog
```

**Output:** The same console output as Example 2. When the run completes, the output is written to `diff.txt` in the current directory, and the Diff menu header reports:

```
diff.txt saved
```

Opening `diff.txt`, the Source Entry Key Status Report appears between the Added entries section and the Diff Summary:

```
 1215 novel entries added
 SOURCE ENTRY KEY STATUS REPORT:
 ✓ = Captured, ✗ = Dropped

 [Abelssoft GoogleClean *] - 5 captured, 0 dropped
     ✓ [CATEGORY] LangSecRef=3024
     ✓ [DETECTION] DetectFile1=%LocalAppData%\Abelssoft\GClean
     ✓ [DETECTION] DetectFile2=%LocalAppData%\Abelssoft\GoogleClean
     ✓ [DELETION] FileKey1=%LocalAppData%\Abelssoft\GClean|*.log
     ✓ [DELETION] FileKey2=%LocalAppData%\Abelssoft\GoogleClean\log|*|REMOVESELF

 [Abelssoft PCFresh Backups *] - 3 captured, 0 dropped
     ✓ [CATEGORY] LangSecRef=3024
     ✓ [DETECTION] DetectFile=%LocalAppData%\Abelssoft\PCFresh
     ✓ [DELETION] FileKey1=%LocalAppData%\Abelssoft\PCFresh\Backup|*
[...]

 KEY STATUS SUMMARY:                 Total  Captured  Dropped  Capture Rate
 DELETION  (FileKey/RegKey):          2511      1825      686         72.7%
 DETECTION (Detect/DetectFile):       1251       273      978         21.8%
 CATEGORY  (Section/LangSecRef):       382       164      218         42.9%
 OTHER     (Warning etc.):              92        18       74         19.6%
 All keys:                            4236      2280     1956         53.8%
```

**Explanation:**

- `-savelog` enables log saving; output is written to `diff.txt` in the current directory after the run completes
- The saved file contains the full diff output including the Source Entry Key Status Report and the summary block
- Use **Log Viewer** in the Diff menu to review the most recent run's output without re-running

**Notes:**

- To save to a different file name: `winapp2ool -diff -savelog -3f changelog.txt`
- To save to a different directory: `winapp2ool -diff -savelog -3d C:\logs`

---

## Example 6: Diffing a Non-Default Flavor

**Context:** You maintain a BleachBit installation and want to know what changed in the BleachBit build of winapp2.ini since your copy.

**Command:**

```
winapp2ool -diff -bb -donttrim
```

**Explanation:**

- `-bb` selects the BleachBit flavor for the download
- `-donttrim` disables the trim step

---

## Example 7: Automated Changelogs (CI/CD)

**Context:** You want a build pipeline or scheduled task to generate a changelog with no interactive console session.

**Command:**

```
winapp2ool -s -diff -base -donttrim -savelog -3f changelog.txt
```

**Output:**

Nothing. In silent mode (`-s`), nothing is output to the console.

The full diff log, identical in structure to the saved log in Example 5, is written to `changelog.txt`.

**Explanation:**

- `-s` suppresses all interactive output and prompts
- `-base` specifies the base flavor
- `-savelog -3f changelog.txt` writes the complete changelog to `changelog.txt`
- Diff fetches the latest winapp2.ini from GitHub automatically
- `-donttrim` disables the trim step

**Notes:** To diff two local files instead:

```
winapp2ool -s -diff -d -1f winapp2-old.ini -2f winapp2.ini -savelog -3f changelog.txt
```
