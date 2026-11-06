# Flavorizer

**Flavorizer** is a [Transmute](../readme.md) sub-module for anyone building their own flavor of winapp2.ini. It applies a complete set of transformations to an ini file in a single operation. Rather than running multiple individual Transmute passes, you provide up to six correction files (one per operation type) and Flavorizer applies them in a fixed, deterministic order. Think of it as a batch orchestrator for Transmute.

Interactively, Flavorizer is opened from the Transmute menu. On the command line it is its own module, invoked with `winapp2ool -flavorize`. See [Command-Line Arguments](#command-line-arguments).

### What is a flavor?

A "flavor" is a collection of modifications that adapts a base ini file for a specific use case. In the context of winapp2.ini, flavors are used to produce variants of the database for different cleaning tools. Flavors are intended as drop-in replacements for the winapp2.ini shipped with their target application. 

### Why Flavorizer?

- **Single operation**: Apply up to six classes of transformation in one run instead of six separate Transmute invocations
- **Deterministic order**: Operations are always applied in the same sequence, so results are predictable and repeatable
- **Selective**: Any correction file can be omitted; stages without a file make no changes
- **Auto-detection**: Flavorizer can automatically discover a set of correction files in a directory by name convention, reducing manual configuration

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

- A base ini file to be transformed
- At least one correction file providing the modifications

All correction files are optional. If none are provided (or none of the provided files exist on disk), Flavorizer warns (*"No correction files specified - output will be identical to input"*) and the run completes with output identical to the input.

---

# Quick Start

### Common Workflow

1. Identify the transformations you want to apply and group them by type (removals, replacements, additions)
2. Create a correction file for each type you need
3. Select your base file and save target
4. Assign your correction files to their appropriate slots
5. Run Flavorizer

*To apply a single isolated transformation, use Transmute directly instead.*

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

In the menu, each correction file is assigned with its own option, such as **Change additions file**, or filled from the target directory by **Auto detect Flavor**; see [Menu Options](#menu-options). `-s` and `-offline` are winapp2ool [global args](../../../Readme.md#global-args).

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

\* winapp2.ini formatting means respecting the winapp2.ini ordering of sections and leading comments/information in the output file. Flavorizer does **not** run WinappDebug on its output. If you are working with winapp2.ini, you should separately run your output through WinappDebug to ensure style/syntax correctness.

Below the options, the menu lists each assigned correction file, colored green when the file exists on disk and red otherwise. A red entry means that stage will make no changes when you Run, check the name and directory.

---

# Command-Line Arguments

Flavorizer supports command-line automation for scripting environments. On the command line, Flavorizer is invoked directly as its own module (not through Transmute):

```
winapp2ool -flavorize [options]
```

### Toggles

| Arg | Effect | Notes |
|:-|:-|:-|
| `-nowinapp` | Save output without winapp2.ini formatting | Formatting is enabled by default |
| `-autodetect` | Automatically detect correction files in the target directory | See [Auto-Detection](#auto-detection) |

### File Selection

Each file slot has a corresponding index for use with the `-Nd` (directory) and `-Nf` (file name) argument pattern:

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

When using `-autodetect`, only slots 1, 2, and 9 are read from the command line; the remaining slots are filled by auto-detection.

---

# Flavor Files

## Operation Order

Flavorizer always applies its correction files in the following fixed order, regardless of which files are provided:

| Stage | Operation | CLI slot | What it does |
|:-|:-|:-|:-|
| 1 | Section Removal | `-3f` | Removes entire sections from the base file |
| 2 | Key Name Removal | `-4f` | Removes individual keys matched by name |
| 3 | Key Value Removal | `-5f` | Removes individual keys matched by type and value |
| 4 | Section Replacement | `-6f` | Replaces entire sections in the base file |
| 5 | Key Replacement | `-7f` | Replaces individual key values within sections |
| 6 | Additions | `-8f` | Adds new sections and keys to the base file |

Stages for which no file has been provided make no changes (each stage still prints its progress line, see [Troubleshooting](#troubleshooting)). The order is significant: for example, removing a key before adding its replacement ensures the result contains exactly one copy of the new value.

Every stage also honors Transmute's [global operations](../readme.md#global-operations): a `[*]` section in a correction file applies its keys to every section of the base file under that stage's mode, the Key Replacement stage additionally recognizes `[*Map: label]` key mapping rules, and every stage recognizes `[*Name: scaffold]` name-filtered rules that apply their changes to only the base sections whose name and content match. Global sections are processed before the named sections in the same file. See [Example 4](EXAMPLES.md#example-4-global-operations-in-flavor-files).

## How Matching Works

**All matching is case-insensitive**: section names, key names, key types, and key values alike, exactly as in Transmute. Each stage matches sections in its correction file to sections in the base file by name; content the base file doesn't contain is ignored with a warning. See [How Matching Works](../readme.md#how-matching-works) in the Transmute readme for the full matching table, including the rule that a criterion applies to *every* matching key in a base section.

---

## Stage 1: Section Removal

Removes entire sections from the base file. This uses Transmute's **Remove BySection** mode. CLI slot: `-3f`.

### Behavior
- Sections present in the correction file that are also present in the base file are removed entirely
- Sections in the correction file not found in the base file are ignored with a warning
- Key values in this file are ignored; only section names matter

### When to use
- Removing entries that are entirely unsupported in the target context
- Pruning incomplete or non-viable generated entries

---

## Stage 2: Key Name Removal

Removes individual keys from sections in the base file, matched by key name. This uses Transmute's **Remove ByKey ByName** mode. CLI slot: `-4f`.

### Behavior
- Keys in the correction file whose name matches a key in the corresponding base section are removed
- Section and key name matching is case-insensitive
- Key values in this file are ignored; only names matter (providing values anyway keeps the file self-documenting)
- Keys in the correction file not found in the base section are ignored with a warning

### When to use
- Removing unnumbered keys (e.g. `Section`, `LangSecRef`) from many entries at once
- Removing a numbered key when you know its exact number

---

## Stage 3: Key Value Removal

Removes individual keys from sections in the base file, matched by key type (name without numbers) and value. This uses Transmute's **Remove ByKey ByValue** mode. CLI slot: `-5f`.

### Behavior
- Keys in the correction file are matched against base keys by comparing their KeyType (name stripped of trailing digits) and value
- Section name, key type, and value matching is case-insensitive
- Numbers in key names are ignored for matching purposes; you can write `FileKey=` or `FileKey1=` and both will match any `FileKey` with the given value
- Keys in the correction file not found in the base section are ignored with a warning

### When to use
- Removing numbered keys (e.g. `FileKey`, `DetectFile`) whose exact numbers may differ between files or may change as entries are updated
- Removing a specific known value without needing to track its current number

---

## Stage 4: Section Replacement

Replaces entire sections in the base file with the versions provided in the correction file. This uses Transmute's **Replace BySection** mode. CLI slot: `-6f`.

### Behavior
- Sections in the correction file that are also found in the base file replace the base section entirely
- Sections in the correction file not found in the base file are ignored with a warning
- No content from the original base section is preserved

### When to use
- Providing a substantially revised version of an entry that is easier to rewrite than to patch incrementally
- Correcting entries that require a large number of changes

---

## Stage 5: Key Replacement

Replaces individual key values within sections. This uses Transmute's **Replace ByKey** mode. CLI slot: `-7f`.

### Behavior
- Keys in the correction file whose name matches a key in the corresponding base section have their value replaced
- Only the key's Value is replaced, the base file's key name (including its casing) is untouched
- Section and key name matching is case-insensitive
- Keys in the correction file not found in the base section are ignored with a warning

### When to use
- Correcting specific key values that are incompatible with the target context without rewriting the entire section
- Updating a known key to a new value as part of a flavor

### Key mapping rules

This stage is the one that recognizes `[*Map: label]` sections: global rules that match keys anywhere in the base file by KeyType and Value and replace the whole key, *including its Name*. The CCleaner flavor uses these to convert generated browser categories (`Section=Google Chrome Web Browser` → `LangSecRef=3029`) across every entry with ~14 rules instead of hundreds of per-entry cohort sections. See [Key Mapping Rules](../readme.md#key-mapping-rules-map) in the Transmute readme for the full semantics, and [Example 4](EXAMPLES.md#example-4-global-operations-in-flavor-files) below.

---

## Stage 6: Additions

Adds new sections and keys to the base file. This uses Transmute's **Add** mode. CLI slot: `-8f`.

### Behavior
- Sections in the correction file not found in the base file are added to the base file as written
- Sections in the correction file that already exist in the base file have the provided keys added to them
- No existing keys are modified or removed
- Keys are not renumbered; use WinappDebug afterward to normalize numbering (winapp2.ini only)

### When to use
- Adding context-specific keys missing from the base file (e.g. adding `LangSecRef` to entries that use `Section`)
- Adding entirely new entries that exist only in the target flavor

---

# Auto-Detection

Flavorizer can automatically discover and assign correction files from a directory by looking for files whose names contain these standard substrings:

| File name contains | Assigned to |
|:-|:-|
| `section_removals.ini` | Section removal file (Stage 1) |
| `name_removals.ini` | Key name removal file (Stage 2) |
| `value_removals.ini` | Key value removal file (Stage 3) |
| `section_replacements.ini` | Section replacement file (Stage 4) |
| `key_replacements.ini` | Key replacement file (Stage 5) |
| `additions.ini` | Additions file (Stage 6) |

Auto-detection uses a substring match, so files like `cc_section_removals.ini` will also be found: this is how the winapp2.ini project's flavor directories are named. The first match found for each slot is used.

The target directory for auto-detection defaults to the current directory. Use **Change target directory** from the menu or `-9d` on the command line to point it elsewhere.

Auto-detection does **not** assign the base file or save target, those must still be configured separately.

###### Note: When run from the menu, the assignments made by Auto detect Flavor are saved with your module settings and persist across sessions. Command-line runs never save settings, so `-autodetect` binds the files for that run only.

---

# Source File Format

All correction files use standard ini format:

```ini
[Section Name]
Key1=Value1
Key2=Value2

[Another Section]
Key1=Value1
```

- Comments in your correction files will **not** appear in the output
- For Stage 1 (Section Removal), keys within sections are ignored, only the section header is required
- For Stage 2 (Key Name Removal), values are ignored, only key names are used for matching
- For Stage 3 (Key Value Removal), numbers in key names are ignored, `FileKey=value` and `FileKey99=value` are equivalent

---

# Tips & Best Practices

### Safety First

- Always test on copies of important files
- Saving to the base file is non-reversible without a backup; Flavorizer does not create one
- Check the *"Applying N correction file(s)"* line at the start of each run: if N is lower than the number of files you assigned, one of them doesn't exist on disk (a typo or a moved directory) and its stage will make no changes
- Review the detailed output log to verify each stage applied correctly

### Organize Your Flavor Files

- Keep one file per stage. Mixing operation types in a single file is not supported
- Name your files using the [Auto-Detection](#auto-detection) conventions so you can use `-autodetect` without manual configuration
- Store all correction files for a flavor together in a dedicated directory

### Stage Ordering Matters

- If you need to replace a key with a different value, consider removing it first (Stage 2 or 3) and then adding the new value (Stage 6) rather than replacing it directly (Stage 5). This avoids numbering conflicts for numbered keys
- Additions are always applied last, so keys added in Stage 6 will not be affected by earlier removal or replacement stages

### When to Use Flavorizer vs Transmute Directly

- Use Flavorizer when you need more than one type of transformation in a single operation
- Use Transmute directly when you need only one transformation, or when you need fine-grained control over the order of multiple passes

---

# Troubleshooting

| Message | Cause |
|:-|:-|
| "You must select a base file" | Run was selected from the menu without a base file assigned |
| "No correction files specified - output will be identical to input" | None of the six correction slots holds a file that exists on disk. The run still completes; the output matches the input |
| "Applying N correction file(s)" shows fewer files than you assigned | One or more assigned files don't exist on disk - check the spelling and directory. The affected stages make no changes |
| "Target section not found in base file: [section] - no changes applied" | A correction file references a section not present in the base file (Stages 1–5) |
| "Removal target not found: {key} not found in {section}" | A key listed for removal doesn't exist in the base section (Stages 2 and 3) |
| "Replacement target not found: {key} not found in {section}" | A key listed for replacement doesn't exist in the base section (Stage 5) |

During a run, every stage prints its progress line in order: *"Removing sections"*, *"Removing keys by name"*, *"Removing keys by value"*, *"Replacing sections"*, *"Replacing keys by name"*, *"Adding keys and sections"*  whether or not a file is assigned to it. A stage header with no changes beneath it is normal for unassigned stages.

Warnings produced by `[*]` and `[*Map:]` sections (refusals, malformed rules, stale rules) are Transmute's. See its [Troubleshooting](../readme.md#troubleshooting) table for the full list.
