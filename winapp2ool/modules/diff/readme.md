# Diff

**Diff** is a winapp2ool module for end users that generates a semantic, context-aware changelog between two versions of winapp2.ini. Unlike a standard [diff](https://en.wikipedia.org/wiki/Diff), **Diff** understands the structure and syntax of winapp2.ini entries and categorizes every change by what actually happened: entries added, modified, renamed, merged into other entries, or removed without replacement, including detecting when keys are moved between entries or are captured by abstractions such as wildcards or recursive deletions.

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

**Using Diff**

1. [Requirements](#requirements)
2. [Quick Start](#quick-start)
   - [Flavors](#flavors)
3. [Common Uses](#common-uses)
4. [Menu Options](#menu-options)
5. [Command-Line Arguments](#command-line-arguments)
   - [Toggles](#toggles)
   - [File Selection](#file-selection)
   - [Flavor Selection](#flavor-selection)
6. [Diff Categories](#diff-categories)
7. [Output and Log](#output-and-log)
   - [Output Structure](#output-structure)
   - [Console Output vs. Saved Log](#console-output-vs-saved-log)
   - [Sample Output](OUTPUT.md#sample-output)
   - [Interpreting the Summary](#interpreting-the-summary)
8. [Tips & Best Practices](#tips--best-practices)
   - [Controlling Scope (Trimming)](#trimming)
   - [CI/CD and Scripting](#scripting)
9. [Troubleshooting](#troubleshooting)

**For Contributors**

10. [How Classification Works](#how-classification-works)
    - [Value Normalization](#value-normalization)
    - [Candidate Selection](#candidate-selection)
    - [Matching Basis](#matching-basis)
    - [Vague Paths](#vague-paths)
    - [Classification Rules](#classification-rules)
11. [Source Entry Key Status Report](#source-entry-key-status-report)

**Examples**

12. **[Examples](EXAMPLES.md)**: seven worked examples with captured input and output

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

# Common Uses

| To | Run | Example |
|:-|:-|:-|
| See what changed between your winapp2.ini and the latest CCleaner-flavor release, trimmed for this machine | `winapp2ool -diff` | |
| Do the same against the latest base (non-CCleaner) release | `winapp2ool -diff -base` | [Default Workflow](EXAMPLES.md#example-1-default-workflow) |
| Compare against another flavor's latest release without trimming it | `winapp2ool -diff -bb -donttrim` | [Diffing a Non-Default Flavor](EXAMPLES.md#example-6-diffing-a-non-default-flavor) |
| Print the full text of each changed entry alongside its changes | `winapp2ool -diff -verbose` | [Verbose Mode](EXAMPLES.md#example-4-verbose-mode) |
| Save the changelog to `diff.txt` | `winapp2ool -diff -savelog` | [Saving the Log](EXAMPLES.md#example-5-saving-the-log) |
| Compare two local files | `winapp2ool -diff -d -1f old.ini -2f new.ini` | [Full Database Comparison](EXAMPLES.md#example-2-full-database-comparison) |
| Write a changelog from a script, with no output or prompts | `winapp2ool -s -diff -base -donttrim -savelog -3f changelog.txt` | [Automated Changelogs (CI/CD)](EXAMPLES.md#example-7-automated-changelogs-cicd) |

In the menu, the same choices are toggles and file pickers set before **Run**; see [Menu Options](#menu-options). The flavor is set in winapp2ool's global settings, as [Flavors](#flavors) explains. `-s` is one of winapp2ool's [global args](../../Readme.md#global-args).

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

---

# Diff Categories

[Reading Diff's Output](OUTPUT.md) explains each of these categories with real samples from Diff's changelog.

| Category | Meaning |
|:-|:-|
| **Added** | |
| [Novel entries](OUTPUT.md#novel-entries) | Entirely new entries with no connection to any removed entry |
| [Entries consolidating removed entries](OUTPUT.md#entries-consolidating-removed-entries) | A new entry contains keys from one or more removed entries |
| **Modified** | |
| [Simple modification](OUTPUT.md#simple-modification) | An entry exists in both files with the same name but has changed |
| [Modification by merger](OUTPUT.md#modification-by-merger) | Entries from the old file are merged into a modified entry; their absorbed keys appear alongside the entry's own changes |
| [Renames with key changes](OUTPUT.md#renames-with-key-changes) | An entry is renamed and also has other key changes, so it does not appear in the Renamed list (which contains only name-only renames) |
| [Cross-Entry Key Movements](OUTPUT.md#cross-entry-key-movements) | A key that leaves an entry present in both files and lands in any other entry has moved |
| **Removed** | |
| [Renamed](OUTPUT.md#renamed) | The entry no longer exists by its old name, but another entry in the new file contains its content mostly unchanged |
| [Merged](OUTPUT.md#merged) | The entry no longer exists, but its content was absorbed into another entry in the new file that is substantially different from the old version |
| [Removed Without Replacement](OUTPUT.md#removed-without-replacement) | The entry is gone and its content does not appear in any other entry in the new file |

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

For a full saved log from a real run, see [Sample Output](OUTPUT.md#sample-output).

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

# How Classification Works

These sections are for people who write or review winapp2.ini entries, or who want to know why Diff classified a change the way it did. You don't need them to read a changelog.

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

# Source Entry Key Status Report

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
