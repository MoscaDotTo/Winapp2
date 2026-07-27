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
   - [Ignored Key Types](#ignored-key-types)
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
- **Trimming:** When diffing against GitHub, trimming is enabled by default. Select `Toggle remote file trim` to disable.
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
| Toggle log saving | Enable or disable saving the diff output to disk | Default: `False` |
| Toggle verbose mode | Print the full text of changed entries alongside each change | Default: `Disabled`; CLI: `-verbose` |
| Choose older/local file | Select the "old" version of winapp2.ini | Default: `winapp2.ini` in the current directory |
| Choose newer file | Select the "new" version of winapp2.ini | Only shown when downloading is disabled |
| Choose save target | Select where to save the diff output | Only shown when log saving is enabled; default: `diff.txt` |
| Log Viewer | Show the most recent diff output | Only available after Diff has been run at least once during the current session |
| Reset Settings | Restore all settings to their defaults | Only shown when settings have been changed |

---

# Diff Categories

###### Note: The output excerpts in this section are taken verbatim from the saved log of a real Diff between base-flavor database versions 220510 (May 2022) and 251109 (November 2025). To reproduce them, extract `Non-CCleaner/Winapp2.ini` at commit `f5ea7371` and at tag `v251109`, then diff the two locally. Because removal processing runs in parallel, list ordering and source attribution for keys shared between multiple entries can vary between runs.

## Added Entries

An entry is **Added** if it exists in the new file and has no entry of the same name in the old file. Added entries render in two separate sections of the output: entries which consolidate content from removed entries appear first (under `Added entries containing merged content:`), followed by entirely novel entries (under `Added entries:`). Each section reports its own total count.

### Novel entries

Entirely new entries with no connection to any removed entry appear as a single line:

```
     .NET Platform * has been added

     360 Secure Browser Autofill Data & Search Engine Preferences * has been added
```

### Entries consolidating removed entries

When a new entry contains keys from one or more removed entries, it is annotated with a count and followed by a `Merged from:` block listing the source entries. The content of those source entries is then tracked across four categories for each key:

| Category | Type | Annotation | Conditions |
|:-|:-|:-|:-|
| **Carried over** | Added keys | `(from [Source *])` | key is carried over verbatim from the source entry |
| **Novel** | Added keys | `(novel)` | key is new content with no equivalent in any source entry |
| **Dropped** | Removed keys | Appears in an `N keys from merged entries not in this entry:` block | the key existed in a source entry but is not captured by any new or modified key in the absorbing entry |
| **Captured** | Modified keys | `(from [Source *])` on the old key | the key does not appear verbatim but is replaced by another key that covers it — typically a wildcard path that absorbs one or more specific old paths |

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
   Abelssoft PCFresh * has been added (consolidating 1 removed entry)
   Merged from:
     • Abelssoft PCFresh Backups *

   4 keys added or carried over from merged sources:

       Added 2 FileKeys
       Added 1 LangSecRef
       Added 1 DetectFile
             FileKey1=%LocalAppData%\Abelssoft\PCFresh|*.log (novel)
             LangSecRef=3024 (from [Abelssoft PCFresh Backups *])
             DetectFile=%LocalAppData%\Abelssoft\PCFresh (from [Abelssoft PCFresh Backups *])
             FileKey2=%LocalAppData%\Abelssoft\PCFresh\Backup|* (from [Abelssoft PCFresh Backups *])
```

**Explanation:**

`FileKey1` is novel; no equivalent exists in `[Abelssoft PCFresh Backups *]`.

`LangSecRef`, `DetectFile`, and `FileKey2` were carried over unchanged and are listed with attribution to their source entry.

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
   EFSoftware  * has been added (consolidating 2 removed entries)
   Merged from:
     • EF Duplicate MP3 Finder *
     • EF Duplicate Files Manager *

   1 key added or carried over from merged sources:

       Added 1 LangSecRef
             LangSecRef=3024 (from [EF Duplicate MP3 Finder *])

   2 keys capturing content from merged entries

       Modified 1 Detection criteria
       Modified 1 FileKey

       Detection criteria modified, replacing 2 old keys
              + New: DetectFile=%AppData%\EFSoftware
              - Old: DetectFile=%AppData%\EFSoftware\MP3Cache (from [EF Duplicate MP3 Finder *])
              - Old: DetectFile=%AppData%\EFSoftware\DFMCache (from [EF Duplicate Files Manager *])
       FileKey1 has been modified, replacing 2 old keys
              + New: FileKey1=%AppData%\EFSoftware|*Cache
              - Old: FileKey1=%AppData%\EFSoftware|MP3Cache (from [EF Duplicate MP3 Finder *])
              - Old: FileKey1=%AppData%\EFSoftware|DFMCache (from [EF Duplicate Files Manager *])
```

**Explanation:**

`LangSecRef` is carried over and attributed to `[EF Duplicate MP3 Finder *]`. Diff attributes a key shared by several sources to whichever source entry it processed first, which can vary between runs.

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
   Cook'n * has been added (consolidating 3 removed entries)
   Merged from:
     • Cook'n Cache *
     • Cook'n Dups *
     • Cook'n Logs *

   6 keys added or carried over from merged sources:

       Added 1 LangSecRef
       Added 1 Detect
       Added 4 FileKeys
             LangSecRef=3021 (from [Cook'n Cache *])
             Detect=HKLM\Software\Microsoft\Windows\CurrentVersion\Uninstall\Cook'n (from [Cook'n Cache *])
             FileKey1=%AppData%\Mozilla\eclipse\Cache|* (from [Cook'n Cache *])
             FileKey2=%LocalAppData%\DVO\Cook'n*App|*.log|RECURSE (from [Cook'n Logs *])
             FileKey3=%LocalAppData%\DVO\Cook'n10App\plugins|Getting Started Guide.rtf|RECURSE (from [Cook'n Dups *])
             FileKey4=%UserProfile%\Documents\Cook'n*|*.log|RECURSE (from [Cook'n Logs *])

   1 key from merged entries not in this entry:

       Removed 1 Warning
             Warning=This removes identical duplicates of the huge Getting Started Guide. (from [Cook'n Dups *])
```

**Explanation:**

Every `FileKey` from all three source entries was carried over, attributed to its originating entry.

For keys shared between multiple entries (`LangSecRef`, `Detect`), Diff attributes to whichever source entry it processes first

The `Warning=` key from `Cook'n Dups *` was dropped, and is listed as removed with attribution to its originating entry.

The change from `%Documents%` to `%UserProfile%\Documents` is a deprecated-path rewrite, so `FileKey4` matches `Cook'n Logs *`'s `FileKey1` despite the textual difference. See [Value Normalization](#value-normalization).

---

## Modified Entries

An entry is **Modified** if it exists in both files with the same name but has changed. Modifications include added keys, removed keys, and changed key values or parameters. Diff itemizes the specific key-level changes for each modified entry.

Keys annotated `(novel)` are new content with no detected equivalent in the old file. Keys annotated `(from [Entry Name *])` were sourced from a specific entry, either the old version of the same entry or a removed entry that was merged in.

Modified output itemizes three kinds of key-level change:

- **Added keys**: present in the new entry with no detected equivalent in the old entry
- **Removed keys**: present in the old entry with no detected equivalent in the new entry
- **Modified keys**: a key whose value changed, shown as `X has been modified, replacing N old keys`; in the simple case one old key was updated to a new value; in more complex cases a single new key uses a wildcard or semicolon-joined pattern to capture multiple old keys

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

When entries from the old file are merged into a modified entry, their absorbed keys appear alongside the entry's own changes. The `(from [Entry Name *])` annotation on each old key identifies from which old entry it originated. Keys labeled with the entry's own name came from its prior version; keys labeled with a different name came from a removed entry that was consolidated. Each such block is followed by a note itemizing the removed entries against which the changes were measured.

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

     Age of Wonders II * has been modified

       Modified 1 FileKey

       FileKey3 has been modified, replacing 2 old keys
              + New: FileKey3=%ProgramFiles%\Age of Wonders*|*Log.txt
              - Old: FileKey1=%ProgramFiles%\Age of Wonders Shadow Magic|*aowsmLog.txt (from [Age of Wonders: Shadow Magic *])
              - Old: FileKey1=%ProgramFiles%\Age of Wonders II|*aow2Log.txt (from [Age of Wonders II *])

     The above changes are measured against the following removed/old entries
     Age of Wonders: Shadow Magic *
```

**Explanation:**

The merger itself is reported in the merged-entries list

The key-level detail appears in the modified-with-merged-content section.

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

       Modified 1 LangSecRef

       LangSecRef has been modified, replacing 1 old key
              + New: LangSecRef=3024
              - Old: LangSecRef=3021
```

**Explanation:**

The `FileKey` and `Detect` content of `[IDLE Recent Files *]` was found unchanged in `[IDLE *]`, so the entry is classified as renamed rather than removed.

`LangSecRef`'s value also changed, so the entry appears in the `Minor changes to renamed entries:` section with its key diffs, rather than in the name-only rename list.

### Cross-Entry Key Movements

Keys that shifted between two entries that both continue to exist are reported in a dedicated `Cross-Entry key movements:` section. Unlike mergers (where a removed entry's content lands in a new or modified entry), these are reassignments of keys from one extant entry to another.

The section is omitted entirely when no movements are detected.

**Diff output:**

```
 Cross-Entry key movements:

   CleanMyPC Registry Cleaner *
       -> FileKey2=%WinDir%\$regcmp$|*|REMOVESELF moved to [Windows Subsystems *]

   iExpert Registry Clean Expert *
       -> FileKey3=%WinDir%\$regcmp$|*|REMOVESELF moved to [Windows Subsystems *]

   Windows Subsystems *
       -> FileKey1=%ProgramData%\Microsoft\PlayReady|*.hds moved to [Microsoft PlayReady *]
       -> FileKey2=%ProgramData%\Microsoft\PlayReady\Cache|* moved to [Microsoft PlayReady *]
       -> FileKey6=%ProgramData%\Microsoft\Windows\DRM|*.log moved to [Microsoft PlayReady *]
       -> FileKey7=%ProgramData%\Microsoft\Windows\DRM\Cache|*|RECURSE moved to [Microsoft PlayReady *]
       -> FileKey8=%ProgramData%\Microsoft\Windows\DRM\PreUpgrade|*.log moved to [Microsoft PlayReady *]

 7 keys moved between entries (3 source entries)
```

**Explanation:**

Two entries each contributed an identical `%WinDir%\$regcmp$|*|REMOVESELF` key to `[Windows Subsystems *]`.

`[Windows Subsystems *]` in turn gave five PlayReady/DRM paths to the `[Microsoft PlayReady *]` entry.

All three source entries continue to exist in the new file, which is what distinguishes a movement from a merger.

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

     Common Language Runtime * has been merged into .NET Framework *

     Adobe Reader XI * has been merged into Adobe Acrobat Reader *
```

**Sample output (merged into a new entry):**

```
     Abelssoft GoogleClean * has been merged into Abelssoft GClean *
```

In the added-entries output, an annotation is included: `Abelssoft GClean * has been added (consolidating 1 removed entry)`.

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

     Clover * has been modified

       Modified 1 FileKey

       FileKey1 has been modified, replacing 2 old keys
              + New: FileKey1=%LocalAppData%\Clover\User Data\Default|Bookmarks.bak;Current *;Last *
              - Old: FileKey1=%LocalAppData%\Clover\User Data\Default|Bookmarks.bak (from [Clover Bookmarks Backups *])
              - Old: FileKey1=%LocalAppData%\Clover\User Data\Default|Current *;Last * (from [Clover Session *])

     The above changes are measured against the following removed/old entries
     Clover Bookmarks Backups *
     Clover Session *
```

The two removed entries each contributed a FileKey, and the new version of `Clover *` combines their values into a single `FileKey1`.

Because `Clover *` is a *modified* entry rather than an added one, its source entries do **not** appear in the [Source Entry Key Status Report](#source-entry-key-status-report).

**Sample output (split across multiple entries):**

```
     Adobe Reader DC * has been split/merged into 2 entries
       • Adobe Acrobat *
       • Adobe Acrobat Reader *
```

A split/merged result means the old entry's content was distributed across more than one new entry, i.e. different keys matched in different places.

**Example with before/after context (large split):** The following shows how `Alternate Services *`, a single legacy entry covering nine different Firefox-based browsers, was split into six browser-specific entries. Each browser that received a dedicated entry now appears as a separate absorbing target.

**Old (v220510):**

```ini
[Alternate Services *]
LangSecRef=3026
Detect1=HKCU\Software\ArtistScope\ArtisBrowser
Detect2=HKCU\Software\Classes\Local Settings\...\Mozilla.Firefox_n80bbvh6b1yt2
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
```

**Diff output:**

```
     Alternate Services * has been split/merged into 6 entries
       • ArtisBrowser Caches *
       • LibreWolf Caches *
       • Pale Moon Caches *
       • Mozilla Firefox Caches *
       • SeaMonkey Caches *
       • Waterfox Caches *
```

Six of the nine browser FileKeys were matched in six new per-browser entries.

The remaining three (IceDragon, SlimBrowser, Basilisk) were not matched in any new entry.

### Removed Without Replacement

The entry is gone and its content does not appear in any other entry in the new file. These are listed individually under the `Entry removals:` header, closed by a count.

**Sample output:**

```
 Entry removals:

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
| **Renamed** | The candidate's name is new to the database (an added entry, not a modified one); all FileKeys and RegKeys from the old entry are matched in the candidate; the match counts are identical; no wildcard reduction or parameter expansion occurred |
| **Merged** | At least one FileKey or RegKey from the old entry is matched in the candidate, but the rename criteria are not met. This includes the case where every key matched but the candidate already existed in the old file (the old entry was absorbed into it) |
| **Removed without replacement** | No FileKeys or RegKeys from the old entry are matched in any candidate|


---

# Output and Log

## Output Structure

A diff run produces output in the following order. Only some sections have a descriptive header line, sections 4, 5 and 7 are identified solely by the count line that closes them.

| # | Section | Header line | Closing count line |
|:-|:-|:-|:-|
| 1 | Run header | `Beginning Diff`, `Diff:  version XXXXXX ->  version XXXXXX` | |
| 2 | Browser categories new to the database | `Web Browser additions` | `N web browsers added` |
| 3 | Entries removed without replacement | `Entry removals:` | `- N removed without replacement` |
| 4 | Name-only renames, old name → new name | *(none)* | `N entries renamed (name-only changes)` |
| 5 | Mergers, including split/merged entries | *(none)* | `N entries merged or split into other entries` |
| 6 | Key diffs for renamed entries | `Minor changes to renamed entries:` | `Minor changes to N renamed entries` |
| 7 | Modified entries that absorbed merged content, with per-entry source notes | *(none)* | `N modified entries incorporating merged content` |
| 8 | Keys relocated between entries that both still exist | `Cross-Entry key movements:` | `N keys moved between entries (M source entries)` |
| 9 | Ordinary modifications | `Modified entries:` | `N modified entries` |
| 10 | Consolidating additions with per-key tracking | `Added entries containing merged content:` | `N added entries consolidating removed content` |
| 11 | Novel additions | `Added entries:` | `N novel entries added` |
| 12 | Per-key capture status for merged sources | `SOURCE ENTRY KEY STATUS REPORT:` | `KEY STATUS SUMMARY:` table |
| 13 | Timing and aggregate counts | `Total diff time:`, `Diff Summary` | `Diff complete` |

Section 8 is omitted when no cross-entry movements were detected. Section 12 is written only to the saved log; use log saving (`-savelog`) or the Log Viewer to see it.

## Console Output vs. Saved Log

The interactive console and the saved log render the same data in two different arrangements. The table above describes the **saved log**.

In the console:

- The count line is printed as a leading header above its list, not below it.
- The descriptive header lines (`Entry removals:`, `Modified entries:`, etc) do not appear at all
- Split/merge blocks gain a `Merged into:` label above the bullet list
- The Source Entry Key Status Report is not shown at all.

## Sample Output

The following is a trimmed extract of the saved log for a real diff run comparing base-flavor winapp2.ini version 220510 (May 2022) against version 251109 (November 2025). `[...]` marks sections cut for brevity.

```
Beginning Diff
Diff:  version 220510 ->  version 251109

 Web Browser additions

   .360 Secure Browser Web Browser

   AOL Shield Web Browser
[...]

 49 web browsers added

 Entry removals:

     2K Launcher * has been removed

     337 Wallpaper * has been removed
[...]

 - 526 removed without replacement

     3delite Filesystem Dialogs * has been renamed to 3delite Filesystem Dialogs Library *

     4Sync Extras * has been renamed to 4Sync *
[...]

 112 entries renamed (name-only changes)


     .NET Application History * has been merged into .NET Framework *

     .NET Framework Isolated Storage * has been merged into .NET Framework *

     3D Viewer * has been merged into Microsoft 3D Viewer *
[...]

 548 entries merged or split into other entries

 Minor changes to renamed entries:

     Avast TuneUp * has been renamed to Avast Cleanup *

       Added 1 Detect
       Added 1 DetectFile
       Added 10 FileKeys
             Detect2=HKLM\Software\AVAST Software\Tuneup
[...]

 Minor changes to 41 renamed entries

     Age of Wonders II * has been modified

       Modified 1 FileKey
[...]

 135 modified entries incorporating merged content

 Cross-Entry key movements:

   CleanMyPC Registry Cleaner *
       -> FileKey2=%WinDir%\$regcmp$|*|REMOVESELF moved to [Windows Subsystems *]
[...]

 7 keys moved between entries (3 source entries)

 Modified entries:

     //N.P.P.D. RUSH// - The Milk of Ultra Violet * has been modified

       Modified 1 Detection criteria

       Detection criteria modified
              + New: DetectFile=%LocalAppData%\NPPDRUSH
              - Old: Detect=HKCU\Software\Valve\Steam\Apps\270090
[...]

 323 modified entries

 Added entries containing merged content:

   Abelssoft GClean * has been added (consolidating 1 removed entry)
   Merged from:
     • Abelssoft GoogleClean *

   4 keys added or carried over from merged sources:

       Added 3 FileKeys
[...]

 317 added entries consolidating removed content

 Added entries:

     .NET Platform * has been added

     360 Secure Browser Autofill Data & Search Engine Preferences * has been added
[...]

 1215 novel entries added
 SOURCE ENTRY KEY STATUS REPORT:
[...]

 Total diff time: 00:00:02.1611092

 Diff Summary
   Net entry count change: 3257 → 3715 (+458)
    + 49 new browsers added
   Modified entries: 460
    + 878 added keys across 276 entries
    - 430 removed keys without replacement across 191 entries
    ~ 354 updated keys replaced 503 old keys across 211 entries
    ~ 7 keys moved from 3 entries into 2 entries
    + 135 entries also received merged content from removed entries (see merged entries below)
   Removed entries: 1227
    @ 548 removed entries have been merged into other entries
       @ 175 merged into 135 modified entries
       + 383 merged into 317 added entries
    & 153 removed entries have been renamed
       = 112 are name-only changes (no key differences)
       + 25 added keys across 12 entries
       - 23 removed keys across 21 entries
       ~ 65 updated keys replaced 77 old keys across 29 entries
    - 526 entries have been removed without replacement
   Added entries: 1685
    @ 317 entries consolidate content from 383 removed entries
       + 175 entries contain 1096 novel keys (not from merged sources)
       = 139 entries contain 856 keys carried over unchanged from merged sources
       ~ 259 entries contain 936 keys capturing 1826 removed keys
       - 261 entries dropped 686 keys from merged sources
    + 1215 novel entries (without merged content)
    & 153 added entries are renamed versions of removed entries and may contain other minor changes

Diff complete
```

## Interpreting the Summary

The **Diff Summary** at the end of each run condenses all classifications into aggregate counts.

**`Net entry count change: OLD → NEW (+N)`**

The entry counts of the old and new files, and the difference between them.

**`+ N new browsers added`**

The number of new browser categories detected among the added entries (itemized in the `Web Browser additions` section at the top of the output).

---

**`Modified entries: N`**
Entries present in both files (by name) that changed in any way.

- `+ N added keys across N entries`: new keys added to existing entries; includes keys absorbed from merged removed entries
- `- N removed keys without replacement across N entries`: keys deleted with no equivalent in the new entry
- `~ N updated keys replaced M old keys across N entries`: keys whose values changed; one new key captured one or more old keys via wildcard match or content consolidation
- `~ N keys moved from X entries into Y entries`: cross-entry key movements; only shown when movements were detected
- `+ N entries also received merged content from removed entries`: how many of the modified entries absorbed removed entries

---

**`Removed entries: N`**
Entries present in the old file but absent from the new file.

- `@ N removed entries have been merged into other entries`: total mergers, categorized:
  - `@ N merged into N modified entries`: absorbed by entries that already existed in the old file
  - `+ N merged into N added entries`: absorbed by entries new to the new file
- `& N removed entries have been renamed`: classified as renames; bullets counting how many were name-only vs. included minor changes
- `- N entries have been removed without replacement`: no matching content found in the new file

> The two merger sub-counts can add up to more than their headline. An entry that was split across both a modified target and an added target is counted in each. In the sample above, 175 + 383 = 558 against a headline of 548, meaning 10 removed entries contributed content to both kinds of target. The same applies to the added-entries bullets, where one entry can appear in several categories at once.

---

**`Added entries: N`**
Entries present in the new file but absent from the old file.

- `@ N entries consolidate content from N removed entries`: new entries that absorbed at least one removed entry; bullets:
  - `+ N entries contain N novel keys`: keys with no equivalent in any source entry
  - `= N entries contain N keys carried over unchanged`: keys transferred verbatim from source entries
  - `~ N entries contain N keys capturing N removed keys`: keys that matched and replaced one or more old keys
  - `- N entries dropped N keys from merged sources`: keys from source entries that did not survive into the absorbing entry
- `+ N novel entries (without merged content)`: entirely new entries with no connection to any removed entry
- `& N added entries are renamed versions of removed entries and may contain other minor changes`: the same rename set as the removed-side `& N removed entries have been renamed` line; the two counts always match

---

When **log saving** is enabled, the diff output is also written to disk (`diff.txt` by default). Use **Log Viewer** in the menu to review the most recent run's output without re-running.

---

## Source Entry Key Status Report

When any mergers into *added* entries are present in the diff, a **Source Entry Key Status Report** is written to the saved log between the Added entries section and the Diff Summary. It does not appear in the interactive console output. Use log saving (`-savelog`) or the Log Viewer to see it.

> **Scope:** the report covers only source entries absorbed by **added** entries. Entries merged into *modified* entries, the `@ N merged into N modified entries` line of the summary, are not included. In the sample run, that means 383 of the 548 merged entries appear here and 175 do not.

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
 DELETION  (FileKey/RegKey):          2512      1826      686         72.7%
 DETECTION (Detect/DetectFile):       1252       274      978         21.9%
 CATEGORY  (Section/LangSecRef):       383       165      218         43.1%
 OTHER     (Warning etc.):              92        18       74         19.6%
 All keys:                            4239      2283     1956         53.9%
```

---

# Command-Line Arguments

Diff's CLI resets all Diff settings to their defaults before applying arguments, so a command-line run never inherits settings saved from the menu. The toggles below flip their setting from its **default** value, not from whatever is on disk.

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

With `-d` and no `-2f`, there is no newer file to against which to diff and the run exits silently without producing any output. Always pair `-d` with `-2f`.

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

---

# Troubleshooting

| Symptom | Cause |
|:-|:-|
| `Run (default)` is displayed in red, and selecting it shows "Please select a file against which to diff" | No "newer" file is selected and downloading is disabled |
| A CLI run with `-d` produces no output at all | No `-2f` was given, so there is no newer file; the run exits silently |
| "winapp2.ini was empty or not found" | The older/local file (or the selected newer file) does not exist or contains no entries |
| Download fails | No internet connection, GitHub is unavailable, or the selected flavor has not been published yet |
| An old key in the log does not match the archived file's text | Deprecated values are rewritten before comparison. See [Value Normalization](#value-normalization) |
| Diff produces slightly different results on successive runs | Removal processing runs in parallel; ordering and shared-key attribution can vary between runs |
| Log Viewer is not available | Diff has not been run yet during the current session |

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
Diff:  version 220510 ->  251109

 Web Browser additions

   Chromium Web Browser

   Microsoft Edge Web Browser

   Mozilla Firefox Web Browser

   Opera Web Browser

 4 web browsers added

 Entry removals:

     Boinc * has been removed

     Bookmarks Backup * has been removed

     Brave Blob Storage * has been removed
[...]

 - 157 removed without replacement

     RegEdit * has been renamed to Windows Registry Editor *

     Recent Colors History * has been renamed to Windows Shell - Theme Colors History *

 10 entries renamed (name-only changes)


     .NET Framework Isolated Storage * has been merged into .NET Framework *

     Accounts Control * has been merged into Windows Security *

     ActiveSync * has been merged into Microsoft Outlook *
[...]

 203 entries merged or split into other entries
[...]

 Modified entries:

     Discord * has been modified

       Added 5 FileKeys
             FileKey1=%AppData%\BetterDiscord|*.log
             FileKey2=%AppData%\BetterDiscord\temp|*|REMOVESELF
             FileKey6=%AppData%\Discord*\Crashpad\reports|*
             FileKey8=%AppData%\Discord*\module_data\crashlogs|*
             FileKey12=%LocalAppData%\Discord*\packages|*.nupkg

       Removed 1 FileKey
             FileKey10=%LocalAppData%\SquirrelTemp|*|REMOVESELF

       Modified 2 FileKeys
[...]

 43 modified entries

 Added entries containing merged content:

   Chromium Caches * has been added (consolidating 14 removed entries)
   Merged from:
     • Download History *
     • Blob Storage *
     • Application Cache Extras *
[...]

 Total diff time: 00:00:00.4186843

 Diff Summary
   Net entry count change: 471 → 350 (-121)
    + 4 new browsers added
   Modified entries: 57
    + 239 added keys across 46 entries
    - 63 removed keys without replacement across 23 entries
    ~ 60 updated keys replaced 101 old keys across 32 entries
    + 14 entries also received merged content from removed entries (see merged entries below)
   Removed entries: 389
    @ 203 removed entries have been merged into other entries
       @ 31 merged into 14 modified entries
       + 178 merged into 82 added entries
    & 29 removed entries have been renamed
       = 10 are name-only changes (no key differences)
       + 25 added keys across 7 entries
       - 13 removed keys across 13 entries
       ~ 33 updated keys replaced 40 old keys across 13 entries
    - 157 entries have been removed without replacement
   Added entries: 268
    @ 82 entries consolidate content from 178 removed entries
       + 59 entries contain 428 novel keys (not from merged sources)
       = 50 entries contain 432 keys carried over unchanged from merged sources
       ~ 62 entries contain 236 keys capturing 695 removed keys
       - 65 entries dropped 1012 keys from merged sources
    + 157 novel entries (without merged content)
    & 29 added entries are renamed versions of removed entries and may contain other minor changes

Diff complete
```

**Explanation:**

- `-base` selects the base flavor 
- Diff downloaded the latest winapp2.ini and trimmed it for the current system before comparing
- Only entries for software detected on this machine appear in the output
- `Net entry count change: 471 → 350 (-121)`: the new file has fewer *total* entries for this machine
- Many entries were consolidated: 203 of the 389 removed entries were mergers
- No `Cross-Entry key movements:` section appears, because no movements were detected in this subset

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
Diff:  version 220510 ->  version 251109

 Web Browser additions

   .360 Secure Browser Web Browser
[...]

 49 web browsers added

 Entry removals:

     2K Launcher * has been removed
[...]

 - 526 removed without replacement

     3delite Filesystem Dialogs * has been renamed to 3delite Filesystem Dialogs Library *
[...]

 112 entries renamed (name-only changes)


     Abelssoft GoogleClean * has been merged into Abelssoft GClean *

     Adobe Reader DC * has been split/merged into 2 entries
       • Adobe Acrobat *
       • Adobe Acrobat Reader *
[...]

 548 entries merged or split into other entries
[...]

 Total diff time: 00:00:02.1611092

 Diff Summary
   Net entry count change: 3257 → 3715 (+458)
    + 49 new browsers added
   Modified entries: 460
    + 878 added keys across 276 entries
    - 430 removed keys without replacement across 191 entries
    ~ 354 updated keys replaced 503 old keys across 211 entries
    ~ 7 keys moved from 3 entries into 2 entries
    + 135 entries also received merged content from removed entries (see merged entries below)
   Removed entries: 1227
    @ 548 removed entries have been merged into other entries
       @ 175 merged into 135 modified entries
       + 383 merged into 317 added entries
    & 153 removed entries have been renamed
       = 112 are name-only changes (no key differences)
       + 25 added keys across 12 entries
       - 23 removed keys across 21 entries
       ~ 65 updated keys replaced 77 old keys across 29 entries
    - 526 entries have been removed without replacement
   Added entries: 1685
    @ 317 entries consolidate content from 383 removed entries
       + 175 entries contain 1096 novel keys (not from merged sources)
       = 139 entries contain 856 keys carried over unchanged from merged sources
       ~ 259 entries contain 936 keys capturing 1826 removed keys
       - 261 entries dropped 686 keys from merged sources
    + 1215 novel entries (without merged content)
    & 153 added entries are renamed versions of removed entries and may contain other minor changes

Diff complete
```

**Explanation:**

- Across three and a half years (May 2022 -> November 2025), 1227 entries were removed (526 without replacement, 153 renamed, and 548 merged into other entries), while 1685 new entries were added
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
Diff:  version 251002 ->  version 251109

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

 Entry removals:

     Abylon Protection Manager * has been removed
[...]

 Diff Summary
   Net entry count change: 3519 → 3715 (+196)
    + 8 new browsers added
   Modified entries: 81
    + 151 added keys across 57 entries
    - 259 removed keys without replacement across 21 entries
    ~ 53 updated keys replaced 54 old keys across 33 entries
    ~ 1 key moved from 1 entry into 1 entry
    + 1 entry also received merged content from removed entries (see merged entries below)
   Removed entries: 313
    @ 18 removed entries have been merged into other entries
       @ 1 merged into 1 modified entry
       + 18 merged into 35 added entries
    & 87 removed entries have been renamed
       = 83 are name-only changes (no key differences)
       + 13 added keys across 2 entries
       ~ 4 updated keys replaced 4 old keys across 3 entries
    - 208 entries have been removed without replacement
   Added entries: 509
    @ 35 entries consolidate content from 18 removed entries
       + 29 entries contain 227 novel keys (not from merged sources)
       = 29 entries contain 117 keys carried over unchanged from merged sources
       ~ 24 entries contain 40 keys capturing 103 removed keys
       - 23 entries dropped 27 keys from merged sources
    + 387 novel entries (without merged content)
    & 87 added entries are renamed versions of removed entries and may contain other minor changes

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


       Added 2 FileKeys
             FileKey2=%AppData%\4Team\Sync2\Logs|*
             FileKey3=%LocalAppData%\4Team\Sync2Cloud|*.log

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
 DELETION  (FileKey/RegKey):          2512      1826      686         72.7%
 DETECTION (Detect/DetectFile):       1252       274      978         21.9%
 CATEGORY  (Section/LangSecRef):       383       165      218         43.1%
 OTHER     (Warning etc.):              92        18       74         19.6%
 All keys:                            4239      2283     1956         53.9%
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
