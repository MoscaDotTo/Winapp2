# Transmute

**Transmute** is a winapp2ool module for anyone customizing winapp2.ini (or another ini file). It changes one ini file using instructions from another, through three operations: Add, Replace and Remove. Each works on whole sections or on individual keys, so you can patch winapp2.ini as broadly or as narrowly as you need.

Transmute replaced the Merge module. If you have an existing Merge workflow, see [Migrating from Merge](#migrating-from-merge) at the bottom of this page.

Transmute is both an end user module and also a winapp2.ini build tool. If you only want to use winapp2.ini with a cleaner, you likely do not need to read this page or use Transmute.

### What does Transmute do?
Transmute modifies one ini file (the *base* file) using the contents of another (the *source* file). It can add new content, replace what's there, or remove what you don't want, down to a single key.

### Why Transmute?
Keep your own changes to winapp2.ini in a file of their own and apply them again whenever winapp2.ini updates, instead of editing the file by hand each time. Transmute runs from the command line as easily as from the menu, so it fits in scripts. It's also how the winapp2.ini build makes each flavor (CCleaner, BleachBit and the rest), through [Flavorizer](Flavorizer/readme.md).

---

# Table of Contents

**Using Transmute**

1. [Requirements](#requirements)
2. [Quick Start](#quick-start)
3. [Common Uses](#common-uses)
4. [Menu Options](#menu-options)
5. [Command-Line Arguments](#command-line-arguments)
   - [Primary Mode](#primary-mode)
   - [Sub-Modes (Replace/Remove)](#sub-modes-replaceremove)
   - [Key Removal Modes](#key-removal-modes)
   - [Preset Source Files](#preset-source-files)
   - [Toggles](#toggles)
   - [File Selection](#file-selection)
6. [Transmute (Primary) Modes](#transmute-primary-modes)
   - [How Matching Works](#how-matching-works)
   - [Add Mode](#add-mode)
   - [Replace Mode](#replace-mode)
   - [Remove Mode](#remove-mode)
7. [Global Operations](#global-operations)
   - [The [*] Section](#the--section)
   - [Key Mapping Rules ([*Map:])](#key-mapping-rules-map)
   - [Name-Filtered Rules ([*Name:])](#name-filtered-rules-name)
8. [Output Formatting](#output-formatting)
9. [Flavorization](#flavorization)
10. [Source Files](#source-files)
11. [Tips & Best Practices](#tips--best-practices)
    - [Safety First](#safety-first)
    - [Effective Source Files](#effective-source-files)
    - [Mode Selection](#mode-selection)
12. [Troubleshooting](#troubleshooting)
13. [Migrating from Merge](#migrating-from-merge)
    - [New content](#new-content)
    - [Replacement content](#replacement-content)
    - [Removals](#removals)

**Examples**

14. **[Examples](EXAMPLES.md)**: nine worked examples with captured input and output

---

# Requirements
- A base ini file you want to modify
- A source ini file holding the changes

---

# Quick Start

### Common Workflow
1. Put your base file (usually `winapp2.ini`) in the same folder as winapp2ool
2. Write a source file holding the changes you want
3. Open Transmute, choose your source file, and pick a mode
4. Select **Run**. Transmute saves the result to `winapp2-transmuted.ini`
5. Adjust and run again until the output is what you want

*If you're applying several transmutations in a row, look at [Flavorizer](Flavorizer/readme.md), which runs them all in one go.*

---

# Common Uses

| To | Run | Example |
|:-|:-|:-|
| Add the entries and keys in `custom.ini` to `winapp2.ini`, saving to `winapp2-transmuted.ini` | `winapp2ool -transmute -add -2f custom.ini` | [Adding New Sections and Keys](EXAMPLES.md#example-1-adding-new-sections-and-keys) |
| Replace key values in `browsers.ini`, overwriting it | `winapp2ool -transmute -replace -bykey -1f browsers.ini -2f browser_key_replacements.ini -3f browsers.ini` | [Replacing Key Values](EXAMPLES.md#example-2-replacing-key-values) |
| Replace whole sections in `winapp2.ini` with the ones in `custom.ini`, saving to `winapp2-transmuted.ini` | `winapp2ool -transmute -replace -bysection -c` | [Replacing Entire Sections](EXAMPLES.md#example-3-replacing-entire-sections) |
| Remove the sections named in `section_removals.ini` from `winapp2.ini`, saving to `cleaned.ini` | `winapp2ool -transmute -remove -bysection -2f section_removals.ini -3f cleaned.ini` | [Removing Entire Sections](EXAMPLES.md#example-4-removing-entire-sections) |
| Remove keys by name from `winapp2.ini`, overwriting it | `winapp2ool -transmute -remove -bykey -byname -2f key_name_removals.ini -3f winapp2.ini` | [Remove Keys By Name](EXAMPLES.md#example-5-remove-keys-by-name) |
| Remove keys by value from `winapp2.ini`, saving to `winapp2-transmuted.ini` | `winapp2ool -transmute -remove -bykey -byvalue -2f key_value_removals.ini` | [Removing Keys By Value](EXAMPLES.md#example-6-removing-keys-by-value) |

In the menu, you make the same choices with the mode options and file choosers before you select **Run**. See [Menu Options](#menu-options). `-c` is one of the [preset source files](#preset-source-files).

---


# Menu Options

| Option | Effect | Notes |
|:-|:-|:-|
| Run (default) | Apply the transmutation with the current settings | Pressing Enter on an empty line also runs it. Refused until you've chosen a source file |
| Open Flavorizer | Open the Flavorizer sub-module | Runs a series of transmutations in a fixed order |
| custom | Set the source file name to `custom.ini` | Preset |
| winapp3.ini | Set the source file name to `winapp3.ini` | Preset |
| Archived Entries | Set the source file name to `Archived Entries.ini` | Preset |
| browsers.ini | Set the source file name to `browsers.ini` | Preset |
| uwp.ini | Set the source file name to `uwp.ini` | Preset |
| Change base file | Choose the file to modify | Default: `winapp2.ini` |
| Change source file | Choose the file holding the changes | |
| Change save target | Choose where the output is saved | Default: `winapp2-transmuted.ini` |
| Toggle Syntax | Turn winapp2.ini formatting of the output on or off \* | Default: `On` |
| Toggle Global sections | Turn on or off treating `[*]`, `[*Map:]` and `[*Name:]` source sections as [global operations](#global-operations) | Default: `On` |
| Change Transmute mode | Cycle through Add → Replace → Remove | Default: `Add` |
| Change Replace mode | Switch Replace between BySection and ByKey | Only shown in `Replace` mode |
| Change Remove mode | Switch Remove between BySection and ByKey | Only shown in `Remove` mode |
| Change Key Removal Mode | Switch between ByName and ByValue | Only shown in `Remove` ByKey mode |
| Reset Settings | Restore all settings to their defaults | Only shown when settings have been changed |

\* [Output Formatting](#output-formatting) describes what winapp2.ini formatting does. Transmute does **not** run WinappDebug on its output, so if you're working with winapp2.ini, run the output through WinappDebug yourself to fix its style and syntax.

---

# Command-Line Arguments
On the command line, Transmute is `winapp2ool -transmute` (`transmute`, `-3` and `3` also work).

**A source file is required.** If you don't name one with `-2f` or a preset flag, Transmute prints an error and exits without changing anything.

You can leave out the arguments marked Default, since Transmute assumes them anyway, though we recommend including them so your scripts say what they do.

Passing a default argument does nothing, so when two arguments conflict, the one that isn't the default wins. For example, `-bykey -bysection` gives you By Section.

### Primary Mode

|Arg|Effect|Notes
|:-|:-|:-
| -add     | Add mode     | Default
| -replace | Replace mode |
| -remove  | Remove mode  |

### Sub-Modes (Replace/Remove)

|Arg|Effect|Notes
|:-|:-|:-
| -bysection | Operate on entire sections |
| -bykey     | Operate on individual keys | Default

### Key Removal Modes
|Arg|Effect|Notes
|:-|:-|:-
| -byname  | Remove keys by exact name match     | Default
| -byvalue | Remove keys by type and value match |

### Preset Source Files
Each of these sets the source file name to a predefined file. Most of them can be downloaded from GitHub.

|Arg|Effect|Description
|:-|:-|:-
| -c | Use "custom.ini"           | The suggested name for a file of your own additions
| -w | Use "winapp3.ini"          | winapp2.ini entries that may break applications. Use at your own risk!
| -a | Use "Archived Entries.ini" | winapp2.ini entries removed because the applications they target are no longer available
| -b | Use "browsers.ini"         | BrowserBuilder's default output file
| -u | Use "uwp.ini"              | UWPBuilder's default output file

###### Note: Source files *must* be local. Transmute doesn't download anything.

### Toggles
|Arg|Effect|
|:-|:-
| -dontlint | Save the output without winapp2.ini formatting: sections in alphabetical order, with no header (see [Output Formatting](#output-formatting))
| -noglobal | Treat `[*]`, `[*Map:]` and `[*Name:]` source sections as ordinary section names instead of [global operations](#global-operations)

### File Selection
| Arg | Effect | Default Value
|:-|:-|:-
| -1d path        | Set the base file directory         | Current directory
| -1f name        | Set the base file name              | winapp2.ini
| -1f subdir\name | Use a file in a subfolder of the base file directory |
| -2d path        | Set the source file directory       | Current directory
| -2f name        | Set the source file name            | None
| -2f subdir\name | Use a file in a subfolder of the source file directory |
| -3d path        | Set the output directory            | Current directory
| -3f name        | Set the output file name            | winapp2-transmuted.ini
| -3f subdir\name | Save to a subfolder of the output directory |

###### Note: The output goes to `winapp2-transmuted.ini` unless you name another file, so your base file stays as it was. To modify the base file in place, pass its name to `-3f`.

---

# Transmute (Primary) Modes

## How Matching Works

**All matching in Transmute ignores case**: section names, key names, key types and key values alike.

| Operation | Source file provides | Matched on | When the base file has no match |
|:-|:-|:-|:-|
| Add | Sections and keys | Section name | The section is added as new |
| Replace BySection | Entire replacement sections | Section name | Ignored with a warning |
| Replace ByKey | Keys carrying replacement values | Section name, then key Name | Ignored with a warning |
| Remove BySection | Section headers (no keys needed) | Section name | Ignored with a warning |
| Remove ByKey (ByName) | Key names (values are ignored) | Section name, then key Name | Ignored with a warning |
| Remove ByKey (ByValue) | Key values | Section name, then KeyType (the name without its number) and Value | Ignored with a warning |

Each criterion in the source file applies to **every** matching key in the base section. If the base section has several keys with the same name (ByName), or the same KeyType and Value (ByValue), all of them are removed or replaced.

## Add Mode

Adds content from the source file to the base file.

### Sub-modes
None

### Behavior

- A source section the base file doesn't have is added as written
- A source section the base file already has gets the source's keys added to it as written
- Add doesn't check for duplicate keys
- No existing key is changed or removed
- Keys are *not* renumbered. For winapp2.ini, WinappDebug fixes the numbering afterwards
- Keys go on the end of the matched section. With Syntax on (the default), they're then regrouped by type when the file is saved (see [Output Formatting](#output-formatting))

### When to use
- Keeping your own customizations to winapp2.ini while still taking its updates
- Merging several configuration files into one
- Adding keys to existing entries (eg. a FileKey for a location the entry misses)
- Adding whole new entries that only exist in your configuration

---

## Replace Mode

Overwrites content in the base file with content from the source file.

### Sub-modes

|Sub Mode|Effect|Notes
|:-|:-|:-|
| BySection | Replaces entire sections by their name       | Section names are matched ignoring case |
| By Key    | Replaces individual key values by their name | Default                                 |

### Behavior

#### By Section
- A source section the base file has replaces the base file's section entirely, as written
- A source section the base file doesn't have is ignored

##### When to use
- Swapping in a heavily revised entry that's easier to rewrite than to patch key by key
- Keeping a local version of an entry (eg. one adapted to your install) while taking updates to the rest of the file

#### By Key
- Each key in a source section gives a new *value* for the key of the same name in the matching base section
- Only the Value changes. The base file's key name, casing included, stays as it was
- Source sections and keys the base file doesn't have are ignored

##### When to use
- Correcting specific values in generated entries without rewriting the whole section
- Updating a known key to a new value as part of keeping a configuration current
- For numbered keys whose numbers may shift as entries change, removing the old value (ByValue) and adding the new one is often more reliable

---

## Remove Mode
Removes content from the base file wherever it matches the source file.

### Sub-modes

|Sub Mode|Effect|Notes
|:-|:-|:-
|By Section | Removes entire sections from the base file by name | Section names are matched ignoring case
|By Key     | Removes individual keys from sections in the base file | Default

### Key Removal Sub Modes
|Sub Mode|Effect|Notes
|:-|:-|:-
|By Name  | Removes keys from the base file by their Name                 | Default. Values in the source file are ignored
|By Value | Removes keys from the base file by their Name and Value | The number in the key name is ignored

### Behavior

#### By Section
- A source section the base file has is removed from it
- A source section the base file doesn't have is ignored
- Any keys in the source section are ignored

##### When to use
- Removing entries that don't apply at all in your target, while taking updates to the rest of the file
- Pruning generated entries that are incomplete or can't work

#### By Key - By Name
- A key in a source section removes every key with the same name from the matching base section
- Values in the source file are ignored. Only the name matters
- Source keys the base file doesn't have are ignored

##### When to use
- Removing unnumbered keys (eg. `Section`, `LangSecRef`) when you know their exact names
- Removing a numbered key when you know its exact number

#### By Key - By Value
- A key in a source section removes every key in the matching base section with the same name (numbers ignored) and the same value
- Source keys the base file doesn't have are ignored

##### When to use
- Removing numbered keys (eg. `FileKey`, `DetectFile`) whose numbers may differ between files or change as entries are updated
- Removing a value you know without having to track which number its key has today

---

# Global Operations

A global operation works across every section of the base file, instead of on one section matched by name. Transmute applies them before any named sections in the same source file, in the order `[*Map:]` → `[*]` → `[*Name:]`, so a named section can refine what a global one did.

The Global sections toggle (`-noglobal` on the command line) controls all three. With it off, `[*]`, `[*Map: ...]` and `[*Name: ...]` are treated as ordinary named sections.

## The [*] Section

A source section named exactly `[*]` applies its keys to every section of the base file under the current mode:

| Mode | What `[*]` does |
|:-|:-|
| Add | Add each key to every base section that doesn't already have a key of that name |
| Replace ByKey | Set the value of the key with that Name in every section that has one |
| Remove ByKey (ByName) | Remove keys with that Name from every section |
| Remove ByKey (ByValue) | Remove keys with that KeyType and Value from every section |
| Remove BySection | **Refused** with a warning, since it would remove every section |
| Replace BySection | **Refused** with a warning, since there's no sensible way to replace every section with one |

###### Note: Add mode refuses numbered keys in `[*]`.

### The %EntryName% token

A key value in a `[*]` section can contain the token `%EntryName%` (case doesn't matter). As each key is applied, Transmute replaces the token with the receiving section's name, trailing ` *` included, so one global key gives every section its own value. The token only works inside `[*]` sections. Named sections keep it as literal text.

##### Example

```ini
; Remove Edge's DetectFile from every entry that carries it (Remove - ByKey - ByValue)
[*]
DetectFile=%LocalAppData%\Microsoft\Edge*
```

```ini
; Give every entry a unique ID and a shared Author (Add mode)
; In [Adobe Flash Player *], ID=%EntryName% becomes ID=Adobe Flash Player *
[*]
ID=%EntryName%
Author=Winapp2.ini Project
```

## Key Mapping Rules ([*Map:])

A `[*Map: <label>]` section is a key mapping rule. It finds keys anywhere in the base file by their KeyType and Value and replaces the whole key, *Name included*. No other Transmute operation can change a key's Name.

Key mapping rules only work in **Replace ByKey** mode. In any other mode Transmute skips them with a warning.

Each rule section holds:

| Key | Meaning | Notes |
|:-|:-|:-|
| `Match=<Name>=<Value>` | What to match | Repeatable as `Match1=`, `Match2=`, ... to map many keys to one. KeyTypes are compared with their numbers stripped, so `Match=FileKey=...` matches any `FileKeyN` with that value. A Value of exactly `*` matches any value of that KeyType |
| `Replace=<Name>=<Value>` | The full replacement key line | Repeatable as `Replace1=`, `Replace2=`, ... to map one key to many. The first replacement takes the matched key's place, and the rest go right after it |

##### Example

```ini
; Convert the generated browser category into CCleaner's built-in one
[*Map: Chrome]
Match=Section=Google Chrome Web Browser
Replace=LangSecRef=3029

; Many-to-one: collapse several categories into a single tag
[*Map: CC7 apps tag]
Match1=LangSecRef=3005
Match2=LangSecRef=3021
Match3=LangSecRef=3022
Replace=Tags=ccapps

; One-to-many: de-abstract a wildcard detection into hardcoded variants
[*Map: Brave DetectFile]
Match=DetectFile=%LocalAppData%\BraveSoftware\Brave-Browser*
Replace1=DetectFile1=%LocalAppData%\BraveSoftware\Brave-Browser
Replace2=DetectFile2=%LocalAppData%\BraveSoftware\Brave-Browser-Beta
Replace3=DetectFile3=%LocalAppData%\BraveSoftware\Brave-Browser-Nightly

; Wildcard fallback: specific rules first, the wildcard catches everything else.
; Because rules are first-match-wins in file order, place the fallback LAST
[*Map: Chrome tag]
Match=LangSecRef=3029
Replace=Tags=Google,Chrome,Browser

[*Map: Category fallback]
Match1=LangSecRef=*
Match2=Section=*
Replace=Tags=ccapps
```

### Rule behavior

- Rules run in one pass, in file order, and the first rule to match a key wins. Each base key is checked against its *original* value only, so a key an earlier rule replaced is never matched again by a later rule in the same run, and keys a one-to-many rule inserts aren't checked at all
- Numbering the replacement keys is up to you. They're written exactly as given, as in Add mode, though the [default output formatting](#output-formatting) renumbers keys anyway
- A rule that matches nothing anywhere in the base file gets a warning. That often means the rule is stale and the values it targets are gone
- A malformed rule is skipped with a warning and the other rules still run. A rule is malformed if it has no `Match=` keys, no `Replace=` keys, a key that isn't `Match` or `Replace`, or a `Match=`/`Replace=` value that isn't itself a `Name=Value` pair

### Example use cases
- The CCleaner flavor keeps its browser category mapping (`Section=` → `LangSecRef=`) as about 14 rules, where it used to need hundreds of per-entry sections kept in step across two files
- The System Ninja flavor, whose cleaner doesn't support wildcards in detection, turns wildcard `DetectFile` paths into hardcoded variants with one rule per wildcard. That replaced a remove file and an add file with a section per affected entry, and new entries with the same wildcards are covered automatically
- The CCleaner 7 flavor turns every entry's `LangSecRef`/`Section` category key into its `Tags=` value: specific rules map each known category, and a wildcard fallback rule (placed last) sends everything else to `Tags=ccapps`. With a `[*]` section adding `ID=%EntryName%` and `Author=`, the whole CCleaner 7 format conversion is written as data

## Name-Filtered Rules ([*Name:])

A `[*Name: <scaffold>]` section applies its keys under the current mode like `[*]` does, but only to the base sections it selects. `[*]` applies everywhere and `[*Map:]` matches on key content alone, while `[*Name:]` selects by the entry's name, plus key content if you ask for it.

A base section is selected when both of these hold:

1. **Name suffix**: the section name ends with `" <scaffold> *"` (ignoring case), where `<scaffold>` is the label after `*Name:`. Transmute adds the ` *` that ends every winapp2.ini entry name, so a label can't match partway through a longer name: `[*Name: Web Browsing Session]` selects `[Firefox Web Browsing Session *]` but **not** `[Firefox Web Browsing Session Backups *]`.
2. **Key content**: an optional set of `Match=` keys. With none, the name alone decides. Otherwise the section needs at least one key that satisfies any one of the `Match=` criteria.

Each rule section holds:

| Key | Meaning | Notes |
|:-|:-|:-|
| `Match=<Name>=<Value>` | Key content to match | Optional, repeatable as `Match1=`, `Match2=`, .... KeyTypes are compared with their numbers stripped. The value can contain wildcards: `Section=* Web Browser` matches `AVG Secure Web Browser` |
| *(any other key)* | Content applied to each selected section | In `Add` mode, a key the section already has isn't added again |

##### Value-glob vs `[*Map:]`

`[*Map:]` only understands a `*` that is the whole value: either the value is a wildcard or it has none. `[*Name:]` understands a `*` **anywhere** in the value (`* Web Browser`, `Brave*`), which is what lets it match messy real Section values like `.360 Secure Browser Web Browser` and `AVG Secure Web Browser`. A bare `*` still matches anything in both.

##### Refusals

As with `[*]`, **Remove BySection** and **Replace BySection** refuse `[*Name:]` sections, and `Add` mode refuses numbered keys in them.

##### Example

```ini
; FluentCleaner: mark the cookies scaffold of every browser Default=False (Add mode).
; Selects any entry named "... Web Browsing Cookies *" that is a browser, and skips
; the key where it already exists. New BrowserBuilder browsers are covered automatically.
[*Name: Web Browsing Cookies]
Match1=Section=* Web Browser
Match2=Section=* Email Client
Default=False
```

A rule that selects no sections gets a warning.

---

# Output Formatting

Transmute writes its output in one of two formats, set by the Syntax toggle (`-dontlint` on the command line):

### Syntax enabled (default)

The output is written as a winapp2.ini file:

- Entries are sorted alphabetically. Only winapp3.ini's sections (`Language Files`, `Dangerous Long`, `Dangerous`) appear in labeled groups
- The keys in each entry are regrouped into winapp2.ini key order (categorization, then detection, then the cleaning keys), keeping their order within each group
- The winapp2.ini header comments (version line, entry count, the flavor if the base file names one, and the project and license links) are written fresh at the top of the file

Because of the regrouping, a key that Add mode put at the end of an entry may not be at the end in the output. [Example 7](EXAMPLES.md#example-7-chaining-operations-together-1) shows this happening.

If your base file isn't a winapp2.ini file, turn Syntax off.

### Syntax disabled (`-dontlint`)

The output is written as a plain ini file with its sections in **alphabetical order** by name. Neither format keeps the base file's original section order.

### Comments are not preserved

Comments from the base file and the source file are *not* carried into the output. If you overwrite your base file, you lose its comments.

---

# Flavorization

Flavorization applies a set of transformations to make a specialized variant of an ini file. The operations run in this order:

1. Section Removal
2. Key Name Removal
3. Key Value Removal
4. Section Replacement
5. Key Replacement
6. Section and Key Additions

The [Flavorizer readme](Flavorizer/readme.md) covers the sub-module that runs this process.

### Example use cases

- Making variants of winapp2.ini (eg. a CCleaner or BleachBit flavor)
- Making platform-specific configurations (eg. a Windows XP flavor)
- Fixing entries the generators get wrong (eg. BrowserBuilder output)

---

# Source Files

Your source file must be a standard ini file:

```ini
[Section Name]
Key1=Value1
Key2=Value2

[Another Section]
Key1=Value1
```

Comments in your source files are fine for your own reference, but they won't appear in the output. See [Output Formatting](#output-formatting).

# Tips & Best Practices

### Safety First

- Test on copies of important files
- The output goes to `winapp2-transmuted.ini` and your base file is left alone unless you point the save target at it. If you do, you can't undo the overwrite without a backup, and winapp2ool doesn't make one
- Comments from either file don't make it into the output

### Effective Source Files

- Keep each source file to one focused set of changes
- Comment your source files so you remember why each change is there
- Section and key names are matched ignoring case
- Turn Syntax off (`-dontlint`) for files that aren't winapp2.ini

### Mode Selection

- Add is for supplements and extensions
- Replace is for updates and corrections
- Remove is for cleanup
- Each sub-mode's "When to use" list under [Transmute (Primary) Modes](#transmute-primary-modes) helps you choose within a mode

---

# Troubleshooting
|Error Message|Cause
|:-|:-
|"You must select a valid source file"|You selected **Run** in the menu without choosing a source file
|"Transmute requires a source file..."|No `-2f` or preset flag was given on the command line
|"{file} was not saved"|Transmute couldn't write the save target. Check that the folder exists and that you can write to it
|"Target section not found in base file: [section] - no changes applied"|The source file names a section the base file doesn't have (Replace and Remove only)
|"Replacement target not found: {key} not found in {section}"|The source key doesn't exist in the base section (Replace ByKey)
|"Removal target not found: {key} not found in {section}"|The source key doesn't exist in the base section (Remove ByKey). {key} is the key name in ByName mode, or the KeyType=Value pair in ByValue mode
|"[\*] cannot be used with Remove BySection (this would remove every section) - skipping"|The source file has a `[*]` section and the mode is Remove BySection
|"[\*] cannot be used with Replace BySection (a global section replacement is incoherent) - skipping"|The source file has a `[*]` section and the mode is Replace BySection
|"[\*] Refusing to add numbered key {key} to every section - global adds must be unnumbered"|A `[*]` section has a numbered key (eg. `FileKey1`) in Add mode
|"[\*Map:] rules are only applied in Replace ByKey mode - skipping..."|The source file has `[*Map:]` sections and the mode isn't Replace ByKey
|"[Map: {label}] is malformed and will be skipped: {reason}"|The rule has no `Match=` or no `Replace=` key, has a key that isn't `Match` or `Replace`, or has a value that isn't a `Name=Value` pair
|"[Map: {label}] matched nothing in {file} - the rule may be stale"|No key in the base file matched any of the rule's criteria. Check whether the values the rule targets still exist
|"[\*Name:] cannot be used with Remove BySection ..." / "... Replace BySection ..."|The source file has `[*Name:]` sections and the mode is Remove or Replace BySection
|"[\*Name: {label}] selected no sections in {file} - the rule may be stale"|No base section ended with the rule's label or passed its `Match=` criteria

---

# Migrating from Merge

### What happened to Merge?

As Merge grew, calling its output the result of a "merger" stopped making sense, so we now call it a Transmutation. In spirit this is still the Merge module, but there are some important technical differences in how Transmute does its job.

Merge *always* applied additions, and its "Add & Remove" and "Add & Replace" modes gave it only limited ways to resolve conflicts. Transmute makes much more granular changes, and can still produce the same output.

To migrate, sort your changes into groups by what they do under Transmute's modes.

### New content
This is the most common case. Content you're adding (eg. custom entries, or keys you wrote for your system) can all go in one file. It works much like additions did in Merge, except you can now add individual keys to existing entries instead of providing a whole replacement section. Apply additions with the Transmute mode set to Add.

### Replacement content
Put sections you want to replace in winapp2.ini in one file, and keys you want to replace within individual sections in another. Apply replacements with the Transmute mode set to Replace. Replace defaults to By Key, so set it to By Section to replace whole sections.

### Removals
Put sections you want to remove from winapp2.ini entirely in one file, and keys you want to remove from individual sections in another. Likewise, keep keys you want to remove by value in a different file from keys you want to remove by name. When removing whole sections, you don't need to include any keys. Apply removals with the Transmute mode set to Remove. Remove defaults to By Key, so set it to By Section to remove whole sections. Key removal defaults to By Name, so set it to By Value to remove keys by their value.

##### This guide is a general framework for deciding where each change goes. For the commands themselves, see the [Examples](EXAMPLES.md).
