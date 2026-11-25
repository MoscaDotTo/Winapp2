# Transmute

Transmute is a patching tool for ini files. It changes one ini file (the *base* file) using instructions stored in another (the *source* file) through three operations: Add, Replace and Remove. Each works on whole sections or on individual keys. 

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
   - [Add Mode](#add-mode)
   - [Replace Mode](#replace-mode)
   - [Remove Mode](#remove-mode)
7. [Global Operations](#global-operations)
   - [The [*] Section](#the--section)
   - [Key Mapping Rules ([*Map:])](#key-mapping-rules-map)
   - [Name-Filtered Rules ([*Name:])](#name-filtered-rules-name)
8. [Output Formatting](#output-formatting)
9. [Flavorization](#flavorization)
10. [Troubleshooting](#troubleshooting)
11. [Migrating from Merge](#migrating-from-merge)
    - [New content](#adding-content)
    - [Replacement content](#replacement-content)
    - [Removals](#removals)
12. **[Usage Examples](EXAMPLES.md)**

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

*If you're applying several transmutations in a row, see [Flavorizer](Flavorizer/readme.md)*

---

# Common Uses

| To | Run | Example |
|:-|:-|:-|
| Add the entries and keys in `custom.ini` to `winapp2.ini`, saving as `winapp2-transmuted.ini` | `winapp2ool -transmute -add -2f custom.ini` | [Adding New Sections and Keys](EXAMPLES.md#example-1-adding-new-sections-and-keys) |
| Replace key values in `browsers.ini`, overwriting it | `winapp2ool -transmute -replace -bykey -1f browsers.ini -2f browser_key_replacements.ini -3f browsers.ini` | [Replacing Key Values](EXAMPLES.md#example-2-replacing-key-values) |
| Replace whole sections in `winapp2.ini` with the ones in `custom.ini`, saving as `winapp2-transmuted.ini` | `winapp2ool -transmute -replace -bysection -c` | [Replacing Entire Sections](EXAMPLES.md#example-3-replacing-entire-sections) |
| Remove the sections declared in `section_removals.ini` from `winapp2.ini`, saving as `cleaned.ini` | `winapp2ool -transmute -remove -bysection -2f section_removals.ini -3f cleaned.ini` | [Removing Entire Sections](EXAMPLES.md#example-4-removing-entire-sections) |
| Remove keys by name from `winapp2.ini`, overwriting it | `winapp2ool -transmute -remove -bykey -byname -2f key_name_removals.ini -3f winapp2.ini` | [Remove Keys By Name](EXAMPLES.md#example-5-remove-keys-by-name) |
| Remove keys by value from `winapp2.ini`, saving as `winapp2-transmuted.ini` | `winapp2ool -transmute -remove -bykey -byvalue -2f key_value_removals.ini` | [Removing Keys By Value](EXAMPLES.md#example-6-removing-keys-by-value) |

In the menu, you make the same choices with the mode options and file choosers before you select **Run**. See [Menu Options](#menu-options). `-c` is one of the [preset source files](#preset-source-files).

---

# Menu Options

| Option | Effect | Notes |
|:-|:-|:-|
| Run (default) | Apply the transmutation with the current settings | Pressing Enter on an empty line also runs if a source file is set |
| Open Flavorizer | Open the [Flavorizer](Flavorizer/readme.md) | Runs a series of transmutations in a fixed order |
| custom | Set the source file name to `custom.ini` | Preset |
| winapp3.ini | Set the source file name to `winapp3.ini` | Preset |
| Archived Entries | Set the source file name to `archived.ini` | Preset |
| browsers.ini | Set the source file name to `browsers.ini` | Preset |
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

\* [Output Formatting](#output-formatting) describes what **Toggle Syntax** does. Transmute does **not** run WinappDebug on its output.

---

# Command-Line Arguments
On the command-line, Transmute is `winapp2ool -transmute` (`transmute`, `-3` and `3` also work).

**A source file is required.** If you don't set one with `-2f` or a preset flag, Transmute prints an error and exits without making any changes.

Arguments marked default can be omitted.

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
| -w | Use "winapp3.ini"          | Aggressive entries that may break applications or take a very long time to scan. Use at your own risk!
| -a | Use "archived.ini"         | winapp2.ini entries removed because the applications they target are no longer available
| -b | Use "browsers.ini"         | BrowserBuilder's default output file

###### Note: Source files *must* be local. Transmute doesn't download anything.

### Toggles
|Arg|Effect|
|:-|:-
| -dontlint | Save the output sections in alphabetical order, with no header (see [Output Formatting](#output-formatting))
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

| Operation | Source file provides | Matched on | When the base file has no match |
|:-|:-|:-|:-|
| Add | Sections and keys | Section name | The section is added |
| Replace BySection | Entire replacement sections | Section name | Ignored with a warning |
| Replace ByKey | Keys carrying replacement values | Section name, then key Name | Ignored with a warning |
| Remove BySection | Section headers (no keys needed) | Section name | Ignored with a warning |
| Remove ByKey (ByName) | Key names (values are ignored) | Section name, then key Name | Ignored with a warning |
| Remove ByKey (ByValue) | Key values | Section name, then KeyType (the name without its number) and Value | Ignored with a warning |

## Add Mode

Adds content from the source file to the base file.

### Sub-modes
None

### Behavior

- A source section the base file doesn't have is added 
- A source section the base file already has gets the source's keys added to it 
- Add doesn't check for duplicate keys
- No existing key is changed or removed
- Numbered keys are *not* renumbered. 
- Keys go on the end of the matched section.

 With Syntax on (the default), keys are regrouped by type when the file is saved (see [Output Formatting](#output-formatting))

---

## Replace Mode

Overwrites content in the base file with content from the source file.

### Sub-modes

|Sub Mode|Effect|Notes
|:-|:-|:-|
| BySection | Replaces entire sections by their name       |  |
| By Key    | Replaces individual key values by their name | Default                                 |

### Behavior

#### By Section
- Sections declared in the source file replace ones by the same name in the base file 
- Sections not in the base file are ignored

#### By Key
- Keys declared in sections in the source file replace their counterparts in like-named sections in the base file 
- Only the Value changes 
- Source sections and keys the base file doesn't have are ignored

---

## Remove Mode
Removes content from the base file wherever it matches the source file.

### Sub-modes

|Sub Mode|Effect|Notes
|:-|:-|:-
|By Section | Removes entire sections from the base file by name | 
|By Key     | Removes individual keys from sections in the base file | Default

### Key Removal Sub Modes
|Sub Mode|Effect|Notes
|:-|:-|:-
|By Name  | Removes keys from the base file by their Name                 | Default. Values in the source file are ignored
|By Value | Removes keys from the base file by their Name and Value | Any numbers in the key name are ignored

### Behavior

#### By Section
- Any section declared in both files is removed from the base file 
- Sections not declared in the base file are ignored

#### By Key - By Name
- A key in a source section removes every key with the same name from the matching base section
- Values in the source file are ignored. Only the name matters
- Source keys the base file doesn't have are ignored

#### By Key - By Value
- A key in a source section removes every key in the matching base section with the same name (numbers ignored) and the same value
- Source keys the base file doesn't have are ignored

---

# Global Operations

A global operation works across every section of the base file. Transmute applies them before any named sections in the same source file, in the order `[*Map:]` → `[*]` → `[*Name:]`

The Global sections toggle (`-noglobal` on the command-line) treats these as ordinary named sections.

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

### The %EntryName% variable

`[*]` sections can access the name of each entry with an `%EntryName%` variable. 

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

A `[*Map: <label>]` section is a key mapping rule. It finds keys anywhere in the base file by their KeyType and Value and replaces the whole key. 

Key mapping rules only work in **Replace ByKey** mode. In any other mode Transmute skips them with a warning.

Each rule section holds at least one `Match` and `Replace` key pairing 

Rules run in the order they're declared in the source file. Malformed rules and rules with no match in the base file produce a warning. 

### Match

`Match` keys have the format `Match=<Name>=<Value>` where `<Name>=<Value>` is the exact text you want to match. A `<Value>` of `*` matches any value for that KeyType, but wildcards are not supported otherwise. Repeatable as `Match1=, Match2=, ...` against a single `Replace`. `

### Replace

`Replace` keys have the format `Replace=<Name>=<Value>` where `<Name>=<Valu>` is the exact text with which you want to replace any matches. Repeatable as `Replace1=, Replace2=, ...` against a single `Match` in which case `Replace1=` replaces the match value and the remaining values are added. 

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

## Name-Filtered Rules ([*Name:])

`[Name: <scaffold>]` rules apply their keys to any base section matching the scaffold. Base sections match the scaffold when their section name ends with `<scaffold> *`. Omit the `*` at the end of winapp2.ini entry names from your `<scaffold>`. Include `Match=` keys to apply the rule only when a particular key is present. Supports wildcards. Repeatable with `Match1=, Mattch2= ...`. Numbers in KeyTypes are ignored when matching.

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

A rule that selects no sections produces a warning.

---

# Output Formatting

Transmute writes its output in one of two formats, set by the Syntax toggle (`-dontlint` on the command line). Neither format keeps the base file's original section order, and comments from either file do not appear in the output.  

When writing output as a winapp2.ini file:

- Entries are sorted alphabetically. Winapp3.ini's sections (`Language Files`, `Dangerous Long`, `Dangerous`) appear in labeled groups under the header comments but above the mainline winapp2.ini entries 
- The keys in each entry are regrouped into winapp2.ini key order but not sorted or renumbered  within those groups
- The winapp2.ini header comments are written fresh at the top of the file

Otherwise, the output is written with sections the sections as-is and in alphabetical order  

---

# Flavorization

Flavorization applies a set of transmutations in a specific order:

1. Section Removal
2. Key Name Removal
3. Key Value Removal
4. Section Replacement
5. Key Replacement
6. Section and Key Additions

The Flavorizer has [its own readme](Flavorizer/readme.md) 

---

# Troubleshooting
|Error Message|Cause
|:-|:-
|"You must select a valid source file"|You selected **Run** in the menu without choosing a source file
|"Transmute requires a source file..."|No `-2f` or preset flag was given on the command line
|"{file} was not saved"|Transmute couldn't write the save target. Check that the folder exists and that you can write to it
|"Target section not found in base file: [section] - no changes applied"|The source file names a section the base file doesn't have (Replace and Remove only)
|"Replacement target not found: {key} not found in {section}"|The source key doesn't exist in the base section (Replace ByKey)
|"Removal target not found: {key} not found in {section}"|The source key doesn't exist in the base section (Remove ByKey). 
|"[\*] cannot be used with Remove BySection (this would remove every section) - skipping"|The source file has a `[*]` section and the mode is Remove BySection
|"[\*] cannot be used with Replace BySection (a global section replacement is incoherent) - skipping"|The source file has a `[*]` section and the mode is Replace BySection
|"[\*] Refusing to add numbered key {key} to every section - global adds must be unnumbered"|A `[*]` section has a numbered key (eg. `FileKey1`) in Add mode
|"[\*Map:] rules are only applied in Replace ByKey mode - skipping..."|The source file has `[*Map:]` sections but the mode isn't Replace ByKey
|"[Map: {label}] is malformed and will be skipped: {reason}"|The rule has no `Match=` or no `Replace=` key, has a key that isn't `Match` or `Replace`, or has a value that isn't a `Name=Value` pair
|"[Map: {label}] matched nothing in {file} - the rule may be stale"|No key in the base file matched any of the rule's criteria.
|"[\*Name:] cannot be used with Remove BySection ..." / "... Replace BySection ..."|The source file has `[*Name:]` sections and the mode is Remove or Replace BySection
|"[\*Name: {label}] selected no sections in {file} - the rule may be stale"|No base section ended with the rule's label or passed its `Match=` criteria

---

# Migrating from Merge

As Merge grew, calling its output the result of a "merger" stopped making sense, so we now call it a Transmutation. In spirit this is still the Merge module, but there are some important technical differences in how Transmute does its job.

Merge *always* applied additions, and its "Add & Remove" and "Add & Replace" modes gave it only limited ways to resolve conflicts. Transmute makes much more granular changes, and can still produce the same output.

To migrate, sort your changes into groups by what they do under Transmute's modes.

### Adding content 

Content you're adding (eg. custom entries, or keys you wrote for your system) can all go in one file. It works much like additions did in Merge, except you can now add individual keys to existing entries instead of providing a whole replacement section. Apply additions with the Transmute mode set to Add.

### Replacement content

Put sections you want to replace in winapp2.ini in one file, and keys you want to replace within individual sections in another. Apply replacements with the Transmute mode set to Replace. Replace defaults to By Key.

### Removals

Put entries you want to remove from winapp2.ini entirely in one file, and keys you want to remove from individual entries in another. Likewise, keep keys you want to remove by value in a different file from keys you want to remove by name. When removing entire sections, you don't need to include any keys. Apply removals with the Transmute mode set to Remove. Remove defaults to By Key - By Name

###### This guide is a general conceptual framework for deciding where each change goes. For the commands, see [Examples](EXAMPLES.md).
