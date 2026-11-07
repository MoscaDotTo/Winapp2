# Flavorizer

**Flavorizer** is a [Transmute](../readme.md) sub-module for anyone building their own flavor of winapp2.ini. It applies a whole set of changes to an ini file in one run. You give it up to six correction files, one for each kind of change, and it applies them in a fixed order. It does the work of several Transmute passes in a single command.

In the menu, you open Flavorizer from Transmute's menu. On the command line it's a module of its own, `winapp2ool -flavorize`. See [Command-Line Arguments](#command-line-arguments).

Flavorizer is both an end-user module for people maintaining their own set of patches while keeping winapp2.ini up to date, and an important build tool for winapp2.ini. If you just want to use winapp2.ini with a cleaner, you do not need to use Flavorizer or read this page.

### What is a flavor?

A flavor is a set of changes that adapts the base ini file for a more specific use. For winapp2.ini, each flavor is a version of the database for a particular cleaning tool, meant to drop in where that tool's own copy of winapp2.ini goes.

### Why Flavorizer?

One Flavorizer run does what would otherwise take up to six Transmute runs, and it always applies them in the same order, so the same files give the same output every time. Every correction file is optional, and a stage with no file changes nothing. If you name your files by its conventions, Flavorizer can also find them in a folder for you (see [Auto-Detection](#auto-detection)).

---

# Table of Contents

**Using Flavorizer**

1. [Requirements](#requirements)
2. [Quick Start](#quick-start)
3. [Common Uses](#common-uses)
4. [Menu Options](#menu-options)
5. [Command-Line Arguments](#command-line-arguments)
   - [Toggles](#toggles)
   - [File Selection](#file-selection)
6. [Flavor Files](#flavor-files)
   - [Operation Order](#operation-order)
   - [How Matching Works](#how-matching-works)
   - [Stage 1: Section Removal](#stage-1-section-removal)
   - [Stage 2: Key Name Removal](#stage-2-key-name-removal)
   - [Stage 3: Key Value Removal](#stage-3-key-value-removal)
   - [Stage 4: Section Replacement](#stage-4-section-replacement)
   - [Stage 5: Key Replacement](#stage-5-key-replacement)
   - [Stage 6: Additions](#stage-6-additions)
7. [Auto-Detection](#auto-detection)
8. [Source File Format](#source-file-format)
9. [Tips & Best Practices](#tips--best-practices)
10. [Troubleshooting](#troubleshooting)

**Examples**

11. **[Examples](EXAMPLES.md)**: four worked examples with captured input and output

---


# Requirements

- A base ini file to change
- At least one correction file

Strictly speaking, every correction file is optional. If you assign none, or none of the ones you assign exist, Flavorizer warns *"No correction files specified - output will be identical to input"* and the run finishes with output that matches the input.

---

# Quick Start

### Common Workflow

1. Work out the changes you want and sort them by kind (removals, replacements, additions)
2. Write a correction file for each kind you need
3. Pick your base file and save target
4. Assign each correction file to its slot
5. Run Flavorizer

*For a single change, Transmute on its own is simpler.*

---

# Common Uses

| To | Run | Example |
|:-|:-|:-|
| Remove the sections named in `section_removals.ini` from `winapp2.ini`, saving to `winapp2-flavorized.ini` | `winapp2ool -flavorize -3f section_removals.ini` | [A Single-Stage Flavor](EXAMPLES.md#example-1-a-single-stage-flavor) |
| Remove sections from `winapp2.ini` and add new content in one run, saving to `winapp2-flavorized.ini` | `winapp2ool -flavorize -3f section_removals.ini -8f additions.ini` | |
| Apply value removals and additions to `browsers.ini`, saving to `browsers-cc.ini` | `winapp2ool -flavorize -1f browsers.ini -2f browsers-cc.ini -5f value_removals.ini -8f additions.ini` | [A Multi-Stage Flavor](EXAMPLES.md#example-2-a-multi-stage-flavor) |
| Apply a key replacement file of global key mapping rules to `browsers.ini`, saving to `browsers-cc.ini` | `winapp2ool -flavorize -1f browsers.ini -2f browsers-cc.ini -7f cc_key_replacements.ini` | [Global Operations in Flavor Files](EXAMPLES.md#example-4-global-operations-in-flavor-files) |
| Auto-detect every correction file in a flavor directory | `winapp2ool -flavorize -autodetect -9d C:\Flavors\CCleaner` | [Auto-Detecting a Flavor](EXAMPLES.md#example-3-auto-detecting-a-flavor) |
| Auto-detect correction files in the current directory and save without winapp2.ini formatting | `winapp2ool -flavorize -autodetect -nowinapp` | |
| Build a flavor from a script, with no output or prompts | `winapp2ool -s -offline -flavorize -autodetect -2f winapp2-ccleaner-flavor.ini -9d \CCleaner` | [Auto-Detecting a Flavor](EXAMPLES.md#example-3-auto-detecting-a-flavor) |

In the menu, each correction file has its own option, such as **Change additions file**, or **Auto detect Flavor** fills them all from the target directory; see [Menu Options](#menu-options). `-s` and `-offline` are winapp2ool [global args](../../../Readme.md#global-args).

---


# Menu Options

| Option | Effect | Notes |
|:-|:-|:-|
| Run (default) | Apply all assigned correction files to the base file | Requires a base file |
| Auto detect Flavor | Scan the target directory for correction files by name | See [Auto-Detection](#auto-detection) |
| winapp2.ini formatting \* | Toggle formatting the output as winapp2.ini | Default: `True` |
| Change base file | Select the file to be transformed | Default: `winapp2.ini` |
| Change save target | Select the output location | Default: `winapp2-flavorized.ini` |
| Change target directory | Set the directory used by Auto detect Flavor | Default: current directory |
| Change section removal file | Assign the Stage 1 correction file | |
| Change key name removal file | Assign the Stage 2 correction file | |
| Change key value removal file | Assign the Stage 3 correction file | |
| Change section replacement file | Assign the Stage 4 correction file | |
| Change key replacement file | Assign the Stage 5 correction file | |
| Change additions file | Assign the Stage 6 correction file | |

\* With winapp2.ini formatting on, the output gets the winapp2.ini header comments and its entries are sorted the way winapp2.ini sorts them. With it off, Flavorizer writes a plain ini with its sections in alphabetical order. Either way, Flavorizer does **not** run WinappDebug on its output, so if you're working on winapp2.ini, lint the result with WinappDebug afterward.

Below the options, the menu lists each assigned correction file, green when the file exists and red when it doesn't. A red file's stage will change nothing when you Run, so check its name and directory.

---

# Command-Line Arguments

On the command line, Flavorizer is its own module rather than a Transmute mode:

```
winapp2ool -flavorize [options]
```

### Toggles

| Arg | Effect | Notes |
|:-|:-|:-|
| `-nowinapp` | Save output without winapp2.ini formatting | Formatting is enabled by default |
| `-autodetect` | Automatically detect correction files in the target directory | See [Auto-Detection](#auto-detection) |

A [flavor flag](../../../Readme.md#flavor-args) such as `-cc` names that flavor in the output's `; Flavor:` header line, replacing whatever flavor the base file named. Without one, and in every menu run, the output keeps the base file's flavor. This only applies with winapp2.ini formatting on.

### File Selection

Each file slot has a number for the `-Nd` (directory) and `-Nf` (file name) arguments:

| Slot | Stage | File | Default |
|:-|:-|:-|:-|
| 1 |  | Base file | `winapp2.ini` in current directory |
| 2 |  | Save target | `winapp2-flavorized.ini` in current directory |
| 3 | 1 | Section removal file | None |
| 4 | 2 | Key name removal file | None |
| 5 | 3 | Key value removal file | None |
| 6 | 4 | Section replacement file | None |
| 7 | 5 | Key replacement file | None |
| 8 | 6 | Additions file | None |
| 9 |  | Auto-detect target directory | Current directory (directory only, use `-9d`) |

| Arg | Effect |
|:-|:-|
| `-Nd path` | Set directory for file slot N |
| `-Nf name` | Set file name for file slot N |
| `-Nf subdir\name` | Set file name within a subdirectory of its path |

With `-autodetect`, Flavorizer reads only slots 1, 2 and 9 from the command line. Auto-detection fills the rest.

---

# Flavor Files

## Operation Order

Flavorizer always applies correction files in this order, whichever ones you provide:

| Stage | Operation | CLI slot | What it does |
|:-|:-|:-|:-|
| 1 | Section Removal | `-3f` | Removes entire sections from the base file |
| 2 | Key Name Removal | `-4f` | Removes individual keys matched by name |
| 3 | Key Value Removal | `-5f` | Removes individual keys matched by type and value |
| 4 | Section Replacement | `-6f` | Replaces entire sections in the base file |
| 5 | Key Replacement | `-7f` | Replaces individual key values within sections |
| 6 | Additions | `-8f` | Adds new sections and keys to the base file |

A stage with no file changes nothing, though it still prints its progress line (see [Troubleshooting](#troubleshooting)). The order matters. Removals run before additions, for example, so removing a key and adding its replacement leaves exactly one copy of the new value.

Every stage also understands Transmute's [global operations](../readme.md#global-operations). A `[*]` section applies its keys to every section of the base file under that stage's mode, and a `[*Name: scaffold]` section does the same for only the base sections whose name and content match. Stages 1 and 4 work on whole sections and refuse both with a warning. The Key Replacement stage also reads `[*Map: label]` key mapping rules. In each file, global sections run before named ones. See [Example 4](EXAMPLES.md#example-4-global-operations-in-flavor-files).

## How Matching Works

All matching ignores case: section names, key names, key types and key values, the same as in Transmute. Each stage pairs the sections in its correction file with base sections by name, and warns about anything the base file doesn't have. Transmute's [How Matching Works](../readme.md#how-matching-works) has the full matching table, including the rule that a criterion applies to *every* matching key in a base section.

---

## Stage 1: Section Removal

Removes whole sections from the base file, using Transmute's **Remove BySection** mode. CLI slot: `-3f`.

### Behavior
- A section in both the correction file and the base file is removed from the base file
- Sections the base file doesn't have are ignored with a warning
- Only section names matter. Any keys under them are ignored

### When to use
- Removing entries that don't work at all in the target tool
- Pruning generated entries that are incomplete or don't apply

---

## Stage 2: Key Name Removal

Removes single keys from base sections, matched by key name, using Transmute's **Remove ByKey ByName** mode. CLI slot: `-4f`.

### Behavior
- A key whose name matches a key in the same base section is removed
- Values are ignored and only names count, though writing the values anyway makes the file easier to read later
- Keys the base section doesn't have are ignored with a warning

### When to use
- Removing unnumbered keys (eg. `Section`, `LangSecRef`) from many entries at once
- Removing a numbered key when you know its exact number

---

## Stage 3: Key Value Removal

Removes single keys from base sections, matched by key type (the name without its number) and value, using Transmute's **Remove ByKey ByValue** mode. CLI slot: `-5f`.

### Behavior
- Each key is compared with the base keys by KeyType (its name with the trailing digits stripped) and value
- The number doesn't count, so `FileKey=` and `FileKey1=` both match any `FileKey` with that value
- Keys the base section doesn't have are ignored with a warning

### When to use
- Removing numbered keys (eg. `FileKey`, `DetectFile`) whose numbers differ between files or shift as entries change
- Removing a known value without tracking its current number

---

## Stage 4: Section Replacement

Replaces whole sections in the base file with the correction file's versions, using Transmute's **Replace BySection** mode. CLI slot: `-6f`.

### Behavior
- A section in both files replaces the base section entirely. Nothing of the old section is kept
- Sections the base file doesn't have are ignored with a warning

### When to use
- Fixing an entry that needs so many changes it's easier to rewrite than to patch

---

## Stage 5: Key Replacement

Replaces single key values inside sections, using Transmute's **Replace ByKey** mode. CLI slot: `-7f`.

### Behavior
- A key whose name matches a key in the same base section gets the correction file's value
- Only the value changes. The base key's name, casing included, stays as it was
- Keys the base section doesn't have are ignored with a warning

### When to use
- Fixing specific values that don't work in the target tool, without rewriting the whole section

### Key mapping rules

This is the only stage that reads `[*Map: label]` sections. These are global rules that match keys anywhere in the base file by KeyType and Value and replace the whole key, *Name included*. The CCleaner flavor has one per browser, turning each generated browser category into CCleaner's (`Section=Google Chrome Web Browser` → `LangSecRef=3029`) in place of hundreds of per-entry sections. Transmute's [Key Mapping Rules](../readme.md#key-mapping-rules-map) has the details, and [Example 4](EXAMPLES.md#example-4-global-operations-in-flavor-files) shows them in a flavor.

---

## Stage 6: Additions

Adds sections and keys to the base file, using Transmute's **Add** mode. CLI slot: `-8f`.

### Behavior
- A section the base file doesn't have is added as written
- A section the base file does have gets the correction file's keys added to it
- Existing keys are never changed or removed
- Keys aren't renumbered. For winapp2.ini, run WinappDebug afterward to fix the numbering

### When to use
- Adding keys the target tool needs that the base file lacks (eg. `LangSecRef` on entries that use `Section`)
- Adding entries that only the flavor has

---

# Auto-Detection

Flavorizer can find and assign correction files from a directory by looking for these names inside each file's name:

| File name contains | Assigned to |
|:-|:-|
| `section_removals.ini` | Section removal file (Stage 1) |
| `name_removals.ini` | Key name removal file (Stage 2) |
| `value_removals.ini` | Key value removal file (Stage 3) |
| `section_replacements.ini` | Section replacement file (Stage 4) |
| `key_replacements.ini` | Key replacement file (Stage 5) |
| `additions.ini` | Additions file (Stage 6) |

It's a substring match, so `cc_section_removals.ini` counts too. That's how the winapp2.ini project names the files in its flavor directories. Flavorizer looks only at the directory itself, not its subfolders, and if more than one file matches a slot it takes the first.

The target directory defaults to the current directory. Point it elsewhere with **Change target directory** in the menu or `-9d` on the command line.

Auto-detection doesn't assign the base file or the save target. You set those yourself.

###### Note: In the menu, the files Auto detect Flavor assigns are saved with your module settings and stay assigned in later sessions. Command-line runs never save settings, so `-autodetect` assigns them for that run only.

---

# Source File Format

Correction files are ordinary ini files:

```ini
[Section Name]
Key1=Value1
Key2=Value2

[Another Section]
Key1=Value1
```

- Comments in correction files don't carry into the output
- Stage 1 (Section Removal) reads only the section headers. Keys under them are ignored
- Stage 2 (Key Name Removal) reads only key names. Values are ignored
- Stage 3 (Key Value Removal) ignores the numbers in key names, so `FileKey=value` and `FileKey99=value` mean the same thing

---

# Tips & Best Practices

### Safety First

- Work on copies of files you care about
- Flavorizer makes no backup, so saving over the base file can't be undone unless you made one yourself
- Check the *"Applying N correction file(s)"* line at the start of each run. If N is lower than the number of files you assigned, one of them isn't on disk (a typo, or a folder that moved) and its stage will change nothing
- Read the log to confirm each stage did what you expected

### Organize Your Flavor Files

- Keep one file per stage. A file can't mix kinds of operation
- Name your files by the [Auto-Detection](#auto-detection) conventions and `-autodetect` will find them for you
- Keep each flavor's files together in a directory of their own

### Stage Ordering Matters

- To change a key's value, you can remove the key (Stage 2 or 3) and add the new one (Stage 6) instead of replacing it (Stage 5). That sidesteps numbering conflicts for numbered keys
- Additions run last, so nothing in an earlier stage can touch the keys Stage 6 adds

### When to Use Flavorizer vs Transmute Directly

- Use Flavorizer when you need more than one kind of change in one run
- Use Transmute on its own for a single change, or when you want to control the order of several passes yourself

---

# Troubleshooting

| Message | Cause |
|:-|:-|
| "You must select a base file" | Run was selected from the menu without a base file assigned |
| "No correction files specified - output will be identical to input" | None of the six correction slots holds a file that exists on disk. The run still finishes, and the output matches the input |
| "Applying N correction file(s)" shows fewer files than you assigned | One or more assigned files aren't on disk. Check the spelling and the directory. Those stages change nothing |
| "Target section not found in base file: [section] - no changes applied" | A correction file names a section the base file doesn't have (Stages 1–5) |
| "Removal target not found: {key} not found in {section}" | A key listed for removal isn't in that base section (Stages 2 and 3) |
| "Replacement target not found: {key} not found in {section}" | A key listed for replacement isn't in that base section (Stage 5) |
| "{file} was not saved" | Flavorizer couldn't write the save target. The run still ends with *"Flavorization completed successfully"*, so look for this line above it |

During a run, every stage prints its progress line in order, whether or not it has a file: *"Removing sections"*, *"Removing keys by name"*, *"Removing keys by value"*, *"Replacing sections"*, *"Replacing keys by name"*, *"Adding keys and sections"*. A heading with nothing under it is normal for a stage with no file.

Warnings from `[*]`, `[*Map:]` and `[*Name:]` sections (refusals, malformed rules, stale rules) come from Transmute. Its [Troubleshooting](../readme.md#troubleshooting) table lists them all.
