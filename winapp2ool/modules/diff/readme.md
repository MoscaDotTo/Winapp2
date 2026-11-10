# Diff

**Diff** is a winapp2ool module for end users that creates a human readable semantic changelog between any two versions of winapp2.ini, making sense of how cleaning rules are changed between versions beyond just their raw text. Unlike a standard [diff](https://en.wikipedia.org/wiki/Diff), **Diff** understands the winapp2.ini syntax, so it can tell when an entry was added, modified, renamed, merged into another entry, or removed without replacement, and tracks when individual keys are modified, move from one entry to another, or get absorbed by a wildcard or a recursive deletion. 

### What does Diff do?

Diff compares two winapp2.ini files and sorts each entry into one of three categories:

- **Added:** in the new file but not the old one
- **Modified:** in both files under the same name, but changed in some way (mergers and renames *aren't* counted here)
- **Removed:** in the old file but not the new one

Removed entries are then split three ways:

- **Renamed:** the old name is gone, but another entry in the new file holds its content mostly unchanged
- **Merged:** the entry is gone, but its content went into another entry that is substantially different from the old version
- **Removed without replacement:** the entry is gone and none of its content turns up anywhere in the new file

By default, Diff compares your local winapp2.ini against the latest one on GitHub, trimmed for the current machine.

###### The downloaded file is flavor-specific, and the default flavor is base. See [Flavors](#flavors).

### Why Diff?

A raw text diff of winapp2.ini contains a lot of noise which is irrelevant to understanding the impact of the changes made to the file. Diff ignores textual and formatting changes and captures abstractions not possible with a simple text diff, then reports the ways in which entries have been changed at a higher level instead. If you keep your local copy trimmed, the optional trim limits the changelog to software you actually have installed.

Diff is also an important winapp2.ini build tool; it is used by the build script to produce the automated changelogs between sequential official winapp2.ini releases.

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
   - [Trimming](#trimming)
   - [Scripting](#scripting)
   - [Reading the outcome from a script](#reading-the-outcome-from-a-script)
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

- A local winapp2.ini to use as the "older" version.
- A second winapp2.ini against which to compare it, either a local file or one Diff downloads from GitHub.
- Typically, both files should be the same [flavor](#flavors).

---

# Quick Start

### Configuration (Optional)

The Diff menu has a few settings you can change before running it:

- `Toggle diffing against GitHub` is on by default. Turn it off to compare two local files.
- `Toggle verbose mode` prints the full text of each changed entry alongside its list of key changes.
- `Toggle remote file trim` is on by default when diffing against GitHub. In offline mode, downloading and trimming both start off.
- `Toggle log saving` is off by default. Turn it on to save the diff log, which is the only place the Source Entry Key Status Report appears.
- `Choose older/local file` points Diff at an older or local `winapp2.ini` in a different folder or under a different name.
- `Choose newer file` picks the newer local `winapp2.ini` when diffing against GitHub is off.

The flavor isn't a Diff setting. It lives in winapp2ool's Global Settings menu (`Change Flavor`) and applies to every module that downloads winapp2.ini.

### Running Diff against GitHub

1. From the winapp2ool main menu, select `Diff`.
2. (Optional) Select `Choose older/local file` if your local `winapp2.ini` is in a different folder or has a different name.
3. Select `Run (default)`, or just press `Enter`.
4. Diff downloads the latest winapp2.ini of the current flavor from GitHub, trims it for this machine, and shows the results.

### Running Diff against local files

Use this to compare two versions you already have.

1. From the winapp2ool main menu, select `Diff`.
2. Select `Toggle diffing against GitHub`.
3. Select `Choose older/local file` if your older copy of `winapp2.ini` is in a different folder or has a different name.
4. Select `Choose newer file` and pick the newer copy of `winapp2.ini`.
5. Select `Run`.

### Running Diff from the winapp2ool main menu

When a new winapp2.ini is available, the main menu offers `Show winapp2.ini changelog` and `Show trimmed changelog`. Both compare the `winapp2.ini` in the current folder against the latest file of the current flavor, and the second trims the download first. They use the Diff module's other settings (verbose mode and log saving) but never change them.

## Flavors

winapp2.ini is published in several **flavors**, each one tailored to the program that reads it. Diff downloads whichever flavor is currently selected as the "newer" file, the same way the Download module does.

The default flavor is base, so `winapp2ool -diff` with no flavor argument compares your local file against the base build. If your local file is the CCleaner flavor, pass `-cc`. Comparing two different flavors is noisy and not recommended, so make sure your local file's flavor matches.

Set the flavor for a single run with a command-line flag (see [Flavor Selection](#flavor-selection)), or permanently from the main menu's `Global Settings` → `Change Flavor`. The Diff menu doesn't show the current flavor.

---

# Common Uses

| To | Run | Example |
|:-|:-|:-|
| See what changed between your winapp2.ini and the latest base release, trimmed for this machine | `winapp2ool -diff` | |
| Do the same against the latest base release, naming the flavor as a script should | `winapp2ool -diff -base` | [Default Workflow](EXAMPLES.md#example-1-default-workflow) |
| Compare against another flavor's latest release without trimming it | `winapp2ool -diff -bb -donttrim` | [Diffing a Non-Default Flavor](EXAMPLES.md#example-6-diffing-a-non-default-flavor) |
| Print the full text of each changed entry alongside its changes | `winapp2ool -diff -verbose` | [Verbose Mode](EXAMPLES.md#example-4-verbose-mode) |
| Save the changelog to `diff.txt` | `winapp2ool -diff -savelog` | [Saving the Log](EXAMPLES.md#example-5-saving-the-log) |
| Compare two local files | `winapp2ool -diff -d -1f old.ini -2f new.ini` | [Full Database Comparison](EXAMPLES.md#example-2-full-database-comparison) |
| Write a changelog from a script, with no output or prompts | `winapp2ool -s -diff -base -donttrim -savelog -3f changelog.txt` | [Automated Changelogs (CI/CD)](EXAMPLES.md#example-7-automated-changelogs-cicd) |

In the menu, the same choices are toggles and file pickers you set before **Run**; see [Menu Options](#menu-options). The flavor is set in winapp2ool's global settings, as [Flavors](#flavors) explains. `-s` is one of winapp2ool's [global args](../../Readme.md#global-args).

---

# Menu Options

| Option | Effect | Notes |
|:-|:-|:-|
| Run (default) | Compare the two files and display the changelog | Needs a "newer" file selected, or downloading turned on. Pressing `Enter` on an empty prompt runs it too |
| Toggle diffing against GitHub | Turn downloading the latest winapp2.ini as the "newer" file on or off | Default: `Enabled`. Unavailable in offline mode |
| Toggle remote file trim | Turn trimming the downloaded file for this machine on or off | Default: `Enabled`; only available when downloading |
| Toggle log saving | Turn saving the diff output to disk on or off | Default: `Disabled` |
| Toggle verbose mode | Print the full text of changed entries alongside each change | Default: `Disabled`; CLI: `-verbose` |
| Choose older/local file | Select the "old" version of winapp2.ini | Default: `winapp2.ini` in the current folder |
| Choose newer file | Select the "new" version of winapp2.ini | Only shown when downloading is off |
| Choose save target | Select where to save the diff output | Only shown when log saving is on; default: `diff.txt` |
| Log Viewer | Show the most recent diff output | Only available once Diff has run this session |
| Reset Settings | Restore all settings to their defaults | Only shown when settings have been changed |

---

# Command-Line Arguments

Diff's CLI resets every Diff setting to its default before applying arguments, so a command-line run never inherits settings saved from the menu. `-d` and `-donttrim` always turn their setting off, so `-offline -d` still compares two local files. `-savelog` and `-verbose` flip their setting from its **default** value, not from whatever is saved on disk.

## Toggles

| Arg | Effect |
|:-|:-|
| `-d` | Disable downloading (compare two local files) |
| `-donttrim` | Disable trimming the downloaded file (downloading stays on) |
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
| `-4f name`, `-summaryf name` | Set the outcome summary file name | None (no summary written) |

With `-d` and no `-2f` there's nothing to compare against, so the run exits silently without any output. Always pair `-d` with `-2f`.

Diff writes the outcome summary if and only if `-4f` gives it a name. It's meant for the build script and release automation; see [Reading the outcome from a script](#reading-the-outcome-from-a-script).

## Flavor Selection

These are global winapp2ool arguments, not Diff's own. They pick which published flavor of winapp2.ini Diff downloads as the "newer" file, for the current run only, and do nothing when `-d` is used.

| Arg | Flavor |
|:-|:-|
| `-base`, `-ncc` | Base (non-CCleaner) (**default**) |
| `-ccleaner`, `-cc` | CCleaner |
| `-bleachbit`, `-bb` | BleachBit |
| `-systemninja`, `-sn` | System Ninja |
| `-tron` | Tron |
| `-ccleaner7`, `-cc7` | CCleaner 7 |
| `-fluentcleaner`, `-fc` | FluentCleaner |

---

# Diff Categories

[Reading Diff's Output](OUTPUT.md) walks through each of these categories with real samples from Diff's changelog.

| Category | Meaning |
|:-|:-|
| **Added** | |
| [Novel entries](OUTPUT.md#novel-entries) | Entirely new entries with no connection to any removed entry |
| [Entries consolidating removed entries](OUTPUT.md#entries-consolidating-removed-entries) | A new entry that holds keys from one or more removed entries |
| **Modified** | |
| [Simple modification](OUTPUT.md#simple-modification) | An entry in both files under the same name that has changed |
| [Modification by merger](OUTPUT.md#modification-by-merger) | Entries from the old file merged into a modified entry; the keys it took in are listed alongside the entry's own changes |
| [Renames with key changes](OUTPUT.md#renames-with-key-changes) | An entry that was renamed and also changed in other ways. These stay out of the Renamed list, which holds name-only renames |
| [Cross-entry key movements](OUTPUT.md#cross-entry-key-movements) | A key that leaves an entry present in both files and lands in any other entry has moved |
| **Removed** | |
| [Renamed](OUTPUT.md#renamed) | The old name is gone, but another entry in the new file holds its content mostly unchanged |
| [Merged](OUTPUT.md#merged) | The entry is gone, but its content went into another entry in the new file that is substantially different from the old version |
| [Removed without replacement](OUTPUT.md#removed-without-replacement) | The entry is gone and none of its content turns up anywhere in the new file |

---

# Output and Log

## Output Structure

A diff run writes its sections in this order. Each one opens with a header line and closes with a count line.

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

Each browser list in section 2 is left out when it would be empty, and so is section 8 when nothing moved between entries. Section 12 is only written to the saved log; use log saving (`-savelog`) or the Log Viewer to see it. A file without a `; Version:` line reads as `version not given` in the run header.

When a run finds nothing at all, the Diff Summary is the single line `No changes detected across N entries`, counting the new file's entries. The log still carries an empty `Entries removed without replacement:` section closed by `- 0 removed without replacement`.

## Console Output vs. Saved Log

The console and the saved log show the same data, arranged differently. The table above describes the **saved log**.

In the console:

- The count line is printed above its list as a header, not below it.
- The descriptive header lines (`Entries removed without replacement:`, `Modified entries:`, etc.) don't appear.
- Split/merge blocks get a `Merged into:` label above the bullet list.
- The Source Entry Key Status Report isn't shown.

For a full saved log from a real run, see [Sample Output](OUTPUT.md#sample-output).

## Interpreting the Summary

The **Diff Summary** at the end of each run boils the classifications down to counts.

**`Net entry count change: OLD → NEW (+N)`**

The entry counts of the old and new files, and the difference between them.

**`+ N new browsers added`**

How many new browser categories turned up among the added entries (itemized in the `Web Browser additions` section at the top of the output).

**`- N browsers removed`**

How many browser categories in the old file no longer appear in the new one (itemized in the `Web Browser removals` section). Both browser lines are shown only when nonzero.

---

**`Modified entries: N`**
Entries in both files (by name) that changed in any way.

- `+ N added keys across N entries`: new keys added to existing entries. Keys moved in from another entry are counted on the moved line instead, and keys taken over unchanged from merged removed entries on the carried-over line
- `- N removed keys without replacement across N entries`: keys the entry's own old version had, with no equivalent in the new entry. Keys that only a merged removed entry had aren't counted here
- `~ N updated keys replaced M old keys across N entries`: keys whose values changed; one new key can capture one or more old keys through recursion, a wildcard match or by consolidating their content
- `~ N keys moved from X entries into Y entries`: cross-entry key movements, shown only when there were any. The X entries are always modified or renamed ones. A moved key isn't also counted as removed. Bullets split the moves by the kind of entry they landed in, each shown only when nonzero:
  - `~ N keys moved into N modified entries`: entries in both files under the same name
  - `& N keys moved into N renamed entries`
  - `+ N keys moved into N added entries`: new entries, whether novel or consolidating removed entries. Splitting part of an entry out into a new one shows up here
- `+ N entries also received merged content from removed entries`: how many of the modified entries absorbed removed entries; one bullet, shown only when nonzero:
  - `= N entries contain N keys carried over unchanged from merged sources`: keys the entries took in verbatim from a merged source that their own old version didn't have, as their `carried over from merged entries` groups list them

---

**`Removed entries: N`**
Entries in the old file but not the new one.

- `@ N removed entries have been merged into other entries`: total mergers, broken down as:
  - `@ N merged into N modified entries`: absorbed by entries that already existed in the old file
  - `+ N merged into N added entries`: absorbed by entries new to the new file
  - `- N split entries lost N keys that no entry kept`: keys of entries split across several new entries that none of them kept, unchanged or captured, as each split entry's `not kept by any of these entries` list shows them
- `& N removed entries have been renamed`: classified as renames; bullets count how many were name-only and how many came with minor changes
- `- N entries have been removed without replacement`: no matching content found in the new file

> The two merger sub-counts can add up to more than their headline, because an entry split across both a modified target and an added target is counted in each. In the [sample run](OUTPUT.md#sample-output), 175 + 382 = 557 against a headline of 547, so 10 removed entries gave content to both kinds of target. The added-entries bullets work the same way: one entry can show up in several of them.

---

**`Added entries: N`**
Entries in the new file but not the old one.

- `@ N entries consolidate content from N removed entries`: new entries that absorbed at least one removed entry; bullets:
  - `+ N entries contain N novel keys`: keys with no equivalent in any source entry, leaving out keys moved in from an entry that still exists
  - `= N entries contain N keys carried over unchanged`: keys taken verbatim from source entries, as the entries' `carried over from merged entries` groups list them
  - `~ N entries contain N keys capturing N removed keys`: keys that matched and replaced one or more old keys, and the old keys they replaced, as the entries' blocks list them
  - `- N entries dropped N keys from merged sources`: the keys the entries' `not in this entry` blocks list. A split source's lost keys are counted under `Removed entries` instead
- `+ N novel entries (without merged content)`: entirely new entries with no connection to any removed entry
- `& N added entries are renamed versions of removed entries and may contain other minor changes`: the same set of renames as the removed side's `& N removed entries have been renamed` line, so the two counts always match

---

With **log saving** on, Diff also writes its output to disk (`diff.txt` by default). The menu's **Log Viewer** shows the most recent run's output without running it again.

---

# Tips & Best Practices

### Trimming

When Diff downloads the newer file, it runs Trim on it first by default, dropping the entries for software this machine doesn't have. That keeps the changelog to software you have installed, but it only makes sense if your local copy is trimmed too. Compare a trimmed file against an untrimmed one and the real changes get buried under every entry Trim dropped.

### Scripting

In silent mode, Diff runs entirely from the command line:

```
winapp2ool -s -diff -base -donttrim -savelog -3f changelog.txt
```

Nothing is printed to the console, so the log file is the only output. The winapp2.ini build runs Diff this way to write the changelog for each flavor it publishes.

Always pass a flavor flag in a script. The CLI resets Diff's own settings, but it still reads the flavor from the saved global settings unless you give it one.

### Reading the outcome from a script

The changelog is meant for people to read. For a script, `-4f` writes the same run's counts as `key=value` lines:

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

`haschanges=false` doesn't mean the two files are identical, only that Diff found no differences about which it cares.

---

# Troubleshooting

| Symptom | Cause |
|:-|:-|
| `Run (default)` is displayed in red, and selecting it shows "Please select a file against which to diff" | No "newer" file is selected and downloading is off |
| A CLI run with `-d` produces no output at all | No `-2f` was given, so there is no newer file; the run exits silently |
| "winapp2.ini was empty or not found" | The older/local file (or the selected newer file) doesn't exist or has no entries |
| Download fails | No internet connection, GitHub is unavailable, or the selected flavor hasn't been published yet |
| An old key in the log doesn't match the archived file's text | Diff rewrites deprecated values before comparing. See [Value Normalization](#value-normalization) |
| Log Viewer is not available | Diff hasn't run yet this session |
| "Diff complete, but diff.txt was not saved" | Log saving is on, but the save target couldn't be written. Pick a writable folder with `Choose save target` |

---

# How Classification Works

This section is for people who write or review winapp2.ini entries, or who want to know why Diff classified a change the way it did. You don't need to understand it in order to read a changelog as an end user. It covers how Diff decides whether a removed entry was renamed, merged, or removed without replacement.

## Value Normalization

Before comparing anything, Diff rewrites a set of deprecated values in both files to their modern equivalents, so that a database-wide style change doesn't show up as thousands of modifications.

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

When an entry is missing from the new file, Diff doesn't compare it against every entry in the new file. It first gathers a pool of candidates four ways:

- name and browser-reference similarity against the added/modified snapshot
- an exact key-value reverse index
- a path-root reverse index
- a wildcard-prefix reverse index

It then keeps only the eligible names: entries added to the new file, or entries in both files that were modified. Only those are considered as rename or merge targets.

## Matching Basis

Diff scores each candidate by comparing the old entry's `FileKey` and `RegKey` values against the candidate's. Other key types help find candidates but don't count toward the score. Keys are compared with wildcard matching, so a key the candidate covers with a broader path or pattern still counts as a match.

## Vague Paths

Some locations are too generic to prove a content match on their own. Matching on them alone would link unrelated entries that happen to touch the same common folder. These include `%AppData%`, `%LocalAppData%`, `%UserProfile%`, `%Documents%`, `%WinDir%`, `%WinDir%\System32`, `%SystemDrive%`, `%Public%`, `%Pictures%`, `%Music%`, `%Video%`, `%UserProfile%\Desktop`, and several broad registry roots such as `HKCU\Software\Microsoft\Windows` and `HKLM\Software\Microsoft\Windows`.

## Classification Rules

| Outcome | Criteria |
|:-|:-|
| **Renamed** | The candidate's name is new to the database (an added entry, not a modified one); all FileKeys and RegKeys from the old entry are matched in the candidate; the match counts are identical; no wildcard reduction occurred (a key gaining patterns is still a rename, with the change itemized) |
| **Merged** | At least one FileKey or RegKey from the old entry is matched in the candidate, but the rename criteria aren't met. This includes the case where every key matched but the candidate already existed in the old file (the old entry was absorbed into it) |
| **Removed without replacement** | No FileKeys or RegKeys from the old entry are matched in any candidate |

---

# Source Entry Key Status Report

When the diff has any mergers into *added* entries, a **Source Entry Key Status Report** is written to the saved log between the Added entries section and the Diff Summary. It never appears in the console. Use log saving (`-savelog`) or the Log Viewer to see it.

> **Scope:** the report covers only source entries absorbed by **added** entries. Entries merged into *modified* entries (the summary's `@ N merged into N modified entries` line) aren't included. In the [sample run](OUTPUT.md#sample-output), that means 382 of the 547 merged entries appear here and 175 don't.

For each source entry it covers, the report shows what happened to every key: either the absorbing entry has it (✓ captured) or it doesn't (✗ dropped). Each key is tagged by role:

- `[DELETION]`: `FileKey` or `RegKey`
- `[DETECTION]`: `Detect` or `DetectFile`
- `[CATEGORY]`: `Section` or `LangSecRef`
- `[OTHER]`: `Warning` and any other key type

Each entry's block opens with its captured and dropped totals, followed by one line per key. The report closes with a `KEY STATUS SUMMARY:` table of counts and capture rates for all four roles.

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

The `Warning=` key from `Cook'n Dups *` is the only key of the three source entries that didn't survive, so it shows as ✗ `[OTHER]`. The same key appears in the `1 key from merged entries not in this entry:` block of the `Cook'n *` entry in the Added section.

After the per-entry blocks comes the summary:

```
 KEY STATUS SUMMARY:                 Total  Captured  Dropped  Capture Rate
 DELETION  (FileKey/RegKey):          2511      1825      686         72.7%
 DETECTION (Detect/DetectFile):       1251       273      978         21.8%
 CATEGORY  (Section/LangSecRef):       382       164      218         42.9%
 OTHER     (Warning etc.):              92        18       74         19.6%
 All keys:                            4236      2280     1956         53.8%
```
