# Transmute 

**Transmute** is a winapp2ool module for anyone customizing winapp2.ini (or another ini file). It provides precise control over modifying ini files through three primary operations: Add, Replace, and Remove, with concise sub-modes where relevant. It enables users to apply targeted changes to configuration files at their chosen granularity (whole sections or individual keys) akin to "patching" winapp2.ini. 

Transmute is the successor to the Merge module. If you have an existing Merge workflow, see [Migrating from Merge](#migrating-from-merge) at the bottom of this page.

### What does Transmute do?
Transmute allows you to modify one ini file (the "base" file) using instructions from another ini file (the "source" file). Think of it as a merge tool that can add new content, replace existing content, or remove unwanted content with precision.

### Why Transmute?
- Batch Operations: Apply complex modifications across entire files automatically
- Flavorization: Create specialized variants of winapp2.ini for different use cases 
- Automation: Perfect for scripting and maintaining configuration variants
- Concise modifications: Straightforward single operation modifications with clear sub-modes 
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
- A base ini file you wish to modify 
- A source ini file providing the modifications

---

# Quick Start

### Common Workflow
1.	Create your base file 
2.	Create your source file(s) with your desired changes
3.	Choose your transmutation mode and options
4.	Run Transmute to apply the changes
5.  Reconfigure and run as many times as necessary to achieve your desired output 

*If you are applying multiple transmutations in succession, you may wish to investigate the Flavorizer*

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

In the menu, the same choices are the mode options and file pickers set before **Run**; see [Menu Options](#menu-options). `-c` is one of the [preset source files](#preset-source-files).

---


# Menu Options

| Option | Effect | Notes |
|:-|:-|:-|
| Run (default) | Apply the current transmutation settings | |
| Open Flavorizer | Open the Flavorizer sub-module | Provides a UI for chain operations |
| Removed Entries | Set the source file name to `Removed Entries.ini` | Preset shortcut |
| custom | Set the source file name to `custom.ini` | Preset shortcut |
| winapp3.ini | Set the source file name to `winapp3.ini` | Preset shortcut |
| browsers.ini | Set the source file name to `browsers.ini` | Preset shortcut |
| uwp.ini | Set the source file name to `uwp.ini` | Preset shortcut |
| Change base file | Select the file to be modified | Default: `winapp2.ini` |
| Change source file | Select the file containing modifications | |
| Change save target | Select the output location | Default: `winapp2-transmuted.ini` |
| Toggle Syntax | Toggle formatting the output as winapp2.ini \* | Default: `True` |
| Toggle Global sections | Toggle treating `[*]`, `[*Map:]`, and `[*Name:]` source sections as [global operations](#global-operations) | Default: `True` |
| Change Transmute mode | Cycle through Add → Replace → Remove | Default: `Add` |
| Change Replace mode | Switch the Replace mode between BySection ↔ ByKey | Only visible in `Replace` mode |
| Change Remove mode | Switch the Remove mode between BySection ↔ ByKey | Only visible in `Remove` mode |
| Change Key Removal Mode | Switch between ByName ↔ ByValue | Only visible in `Remove` ByKey mode |
| Reset Settings | Restore all settings to their defaults | Only shown when settings have been changed |

\* See [Output Formatting](#output-formatting) for exactly what winapp2.ini formatting entails. Transmute does **not** run WinappDebug on its output. If you are working with winapp2.ini, you should separately run your transmute output through WinappDebug to ensure style/syntax correctness.  

---

# Command-Line Arguments
Transmute supports command-line automation for scripting environments.

**A source file is required.** If no source file is provided via `-2f` or a preset flag, Transmute prints an error and exits without doing anything.

Arguments labeled "Default" are assumed by default and can be optionally omitted when invoking, though this is not recommended. 

Arguments affecting default settings are ignored. Eg. If you specify `-bykey`, the default sub mode, it is ignored. If you specify both `-bykey` and also `-bysection`, the `-bykey` will be ignored and the resulting sub mode will be `By Section`

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
Sets the source file name to one of the pre-defined defaults, most of which are available from GitHub

|Arg|Effect|Description
|:-|:-|:-
| -r | Use "Removed Entries.ini"  | Contains winapp2.ini entries removed because CCleaner incorporated their coverage natively
| -c | Use "custom.ini"           | Suggested name for user additions files
| -w | Use "winapp3.ini"          | Contains winapp2.ini entries which may potentially break applications, use at your own  risk!
| -a | Use "Archived Entries.ini" | Contains winapp2.ini entries removed because the applications they target are no longer available 
| -b | Use "browsers.ini"         | The default output file from Browser Builder
| -u | Use "uwp.ini"              | The default output file from UWP Builder

###### Note: Source files *must* be local. Transmute does not directly download any files.

### Toggles
|Arg|Effect|
|:-|:-
| -dontlint | Save the output without winapp2.ini formatting. Sections are written in alphabetical order with no preamble (see [Output Formatting](#output-formatting))
| -noglobal | Treat `[*]`, `[*Map:]`, and `[*Name:]` source sections as ordinary section names instead of [global operations](#global-operations)

### File Selection
| Arg | Effect | Default Value
|:-|:-|:-
| -1d path        | Set base file path                                      | Current Directory
| -1f name        | Set base file name                                      | winapp2.ini
| -1f subdir\name | Set base file name within a subfolder of its path       |
| -2d path        | Set source file path                                    | Current Directory            
| -2f name        | Set source file name                                    | None 
| -2f subdir\name | Set the source file name within a subfolder of its path | 
| -3d path        | Set output file path                                    | Current Directory  
| -3f name        | Set output file name                                    | winapp2-transmuted.ini
| -3f subdir\name | Set the output file name within a subfolder of its path |

###### Note: By default the output is saved to `winapp2-transmuted.ini` and the base file is left untouched. To modify the base file in place, pass the base file's name to `-3f` explicitly.

---

# Transmute (Primary) Modes

## How Matching Works

**All matching in Transmute is case-insensitive** — section names, key names, key types, and key values alike.

| Operation | Source file provides | Matched on | When no match exists in the base file |
|:-|:-|:-|:-|
| Add | Sections and keys | Section name | Section is added as new |
| Replace BySection | Entire replacement sections | Section name | Ignored with a warning |
| Replace ByKey | Keys carrying replacement values | Section name, then key Name | Ignored with a warning |
| Remove BySection | Section headers (keys unnecessary) | Section name | Ignored with a warning |
| Remove ByKey (ByName) | Key names (values are ignored) | Section name, then key Name | Ignored with a warning |
| Remove ByKey (ByValue) | Key values | Section name, then KeyType (name without numbers) + Value | Ignored with a warning |

A criterion from the source file applies to **every** matching key in the base section: if the base section contains multiple keys sharing the same name (ByName) or the same KeyType and Value pair (ByValue), all of them are removed or replaced.

## Add Mode

Adds content from the source file to the base file.

### Sub-modes 
None

### Behavior

- Sections from the source file not found in the base file are added to the base file as written
- Sections from the source file which are found in the base file have the keys from the source file added to the base file as written 
- Does not avoid creating duplicate keys
- No existing keys are modified or removed
- Keys are *not* renumbered; normalize later with WinappDebug if needed (winapp2.ini only)
- Keys are appended to the end of the matched section; when Syntax is enabled (the default), keys are then regrouped by type on save (see [Output Formatting](#output-formatting))

### When to use
- Maintaining a set of customizations to winapp2.ini while also keeping it up to date 
- Merging multiple configuration files into one
- Adding keys to existing entries (eg. a FileKey covering a location the entry misses)
- Adding entirely new entries which exist only in your configuration

---

## Replace Mode

Overwrites existing content in the base file with content from the source file.

### Sub-modes

|Sub Mode|Effect|Notes
|:-|:-|:-|
| BySection | Replaces entire sections by their name       | Section name matches are case-insensitive |
| By Key    | Replaces individual key values by their name | Default                                   |

### Behavior 

#### By Section
- Sections from the source file which are found in the base file replace entirely the section in the base file as they are written  
- Sections from the source file not found in the base file are ignored 

##### When to use
- Providing a substantially revised version of an entry that is easier to rewrite than to patch key by key
- Maintaining a local version of an entry (eg. one adapted to your installation) while keeping the rest of the file up to date

#### By Key
- Keys from sections in the source file which are found in the base file provide replacement *values* for keys of the same name in the base file
- Only the key's Value is replaced — the base file's key name (including its casing) is untouched
- Sections and keys from the source file not found in the base file are ignored 

##### When to use
- Correcting specific key values in generated entries without rewriting the entire section
- Updating a known key to a new value as part of keeping a configuration current
- For numbered keys whose numbering may shift as entries are updated, it can be more reliable to Remove the old value (ByValue) and Add the new one instead

---

## Remove Mode
Removes content from the base file based on matches in the source file.

### Sub-modes

|Sub Mode|Effect|Notes
|:-|:-|:-
|By Section | Removes entire sections from the base file by their name | Section name matches are case-insensitive
|By Key     | Removes individual keys from sections in the base file   | Default 

### Key Removal Sub Modes
|Sub Mode|Effect|Notes
|:-|:-|:-
|By Name  | Removes keys from the base file by their Name                 | Default. Values provided in the source file are ignored
|By Value | Removes keys from the base file based on their Name and Value | Value matches ignore numbers in the key name

### Behavior

#### By Section
- Sections from the source file which are found in the base file are removed 
- Sections from the source file not found in the base file are ignored 
- Key values are ignored in this mode 

##### When to use
- Removing entries which are entirely unsupported in the target context while keeping the rest of the file up to date
- Pruning incomplete or non-viable generated entries

#### By Key - By Name
- Keys from sections in the source file which are found in the base file are removed if they have the same name 
- The values provided for keys in the source file are ignored — matching is by name alone
- Keys from the source file not found in the base file are ignored

##### When to use
- Removing unnumbered keys (eg. `Section`, `LangSecRef`) when you know their exact names
- Removing a numbered key when you know its exact number

#### By Key - By Value
- Keys from sections in the source file which are found in the base file are removed if they have the same name (ignoring numbers) and value 
- Keys from the source file not found in the base file are ignored 

##### When to use
- Removing numbered keys (eg. `FileKey`, `DetectFile`) whose exact numbers may differ between files or change as entries are updated
- Removing a specific known value without needing to track its current key number

---

# Global Operations

Transmute is capable of performing operations over every section in the base file rather than to a single section matched by name. These operations are processed before any named sections in the same source file, in the order `[*Map:]` → `[*]` → `[*Name:]`. This ordering means specific per-section operations can refine the result of global ones.  

All three are controlled by the Global sections toggle (`-noglobal` via the command line). When disabled, `[*]`, `[*Map: ...]`, and `[*Name: ...]` are treated as ordinary named sections. 

## The [*] Section

A source section literally named `[*]` applies its keys to every section of the base file under the current mode:

| Mode | Global meaning |
|:-|:-|
| Add | Add each key to every base section |
| Replace ByKey | Set the value of keys matching Name in every section that has them |
| Remove ByKey (ByName) | Remove keys matching Name from every section |
| Remove ByKey (ByValue) | Remove keys matching KeyType + Value from every section |
| Remove BySection | **Refused** with a warning: this would remove every section |
| Replace BySection | **Refused** with a warning: a global section replacement is incoherent |

###### Note: Numbered keys will be rejected in Add mode. 

### The %EntryName% token

A key value in a `[*]` section may contain the token `%EntryName%` (case-insensitive). As the key is applied, the token is replaced with each receiving section's name (including the trailing ` *` of winapp2.ini entries) so a single global key produces a distinct value per section. The token is only interpreted inside `[*]` sections; named-section operations leave it as a literal value.

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

`[*Map: <label>]` sections define key mapping rules: match a key anywhere in the base file by its KeyType and Value, and replace the whole key *including its Name*. This is the only Transmute operation that can change a key's Name.

Key mapping rules are recognized **only in Replace ByKey mode**. Under any other mode they are skipped with a warning

Each rule section contains:

| Key | Meaning | Notes |
|:-|:-|:-|
| `Match=<Name>=<Value>` | A match criteria | Repeatable as `Match1=`, `Match2=`, ... for many-to-one mappings. KeyTypes are compared with numbers stripped from both sides, so `Match=FileKey=...` matches any `FileKeyN` with that value. A Value of exactly `*` is a wildcard matching any value of that KeyType |
| `Replace=<Name>=<Value>` | The full replacement key line | Repeatable as `Replace1=`, `Replace2=`, ... for one-to-many mappings. The first replacement takes the matched key's ordinal position; any remaining replacements are inserted immediately after it in file order |

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

- Rules are applied in a single pass in file order with first-match-wins semantics: each base key is evaluated against its *original* value only, so a key replaced by an earlier rule is never re-matched by a later rule in the same run, and keys inserted by a one-to-many rule are never evaluated at all
- Replacement key numbering is the rule author's responsibility — replacement keys are written exactly as given, consistent with Add mode. Note that the [default output formatting](#output-formatting) renumbers keys anyway
- A rule which matches nothing anywhere in the base file emits a warning. This can be a sign that the rule has gone stale and the values it targets no longer exist in the base file
- A malformed rule is skipped with a warning while the remaining rules still apply. A rule is malformed if it has no `Match=` keys, no `Replace=` keys, an unrecognized key, or a `Match=`/`Replace=` value which is not itself a `Name=Value` pair

### Example use cases
- Maintaining the browser category mapping (`Section=` → `LangSecRef=`) for the CCleaner flavor as ~14 rules instead of hundreds of per-entry cohort sections that must be kept in lockstep across two files
- De-abstracting wildcard `DetectFile` paths into hardcoded variants for the System Ninja flavor (which does not support wildcards in detection) as one rule per wildcard, instead of a remove + add lockstep spanning two files and one section per affected entry. New entries carrying the same wildcards are covered automatically
- Converting every entry's `LangSecRef`/`Section` category key into its CCleaner 7 `Tags=` value: specific rules map each known category and a wildcard fallback rule (placed last) funnels every remaining category into `Tags=ccapps`. Combined with a `[*]` section adding `ID=%EntryName%` and `Author=`, this expresses the entire CCleaner 7 format conversion as data

## Name-Filtered Rules ([*Name:])

A `[*Name: <scaffold>]` section applies its keys under the current mode like `[*]`, but only to the base sections it selects instead of every section. Where `[*]` is unconditional and `[*Map:]` matches by key content alone, `[*Name:]` selects by the entry name plus optional key content. 

A base section is selected when both conditions hold:

1. **Anchored name suffix**: the section name ends with `" <scaffold> *"` (case-insensitive), where `<scaffold>` is the label after `*Name:`. The engine appends the ` *` for winapp2.ini entries to prevent entry name substring collision: `[*Name: Web Browsing Session]` selects `[Firefox Web Browsing Session *]` but **not** `[Firefox Web Browsing Session Backups *]`.
2. **Key Content**: an optional set of `Match=` keys. With none, the name suffix alone decides. Otherwise the section must contain at least one key satisfying any a `Match=` criteria. 

Each rule section contains:

| Key | Meaning | Notes |
|:-|:-|:-|
| `Match=<Name>=<Value>` | Key content matching criteria | Optional, repeatable as `Match1=`, `Match2=`, ..., KeyTypes compared with numbers stripped. The value supports wildcards; eg. `Section=* Web Browser` matches `AVG Secure Web Browser`.  |
| *(any other key)* | Content applied to each selected section | In `Add` mode, duplicate keys will not be re-added |

##### Value-glob vs `[*Map:]`

`[*Map:]` only understands a whole-value `*` (the entire value is a wildcard or nothing is). `[*Name:]` predicates understand a **partial** glob anywhere in the value (`* Web Browser`, `Brave*`), which is why it can match the messy real Section values (`.360 Secure Browser Web Browser`, `AVG Secure Web Browser`). A bare `*` still reduces to "match anything", so the two are compatible.

##### Refusals

Mirroring `[*]`: **Remove BySection** and **Replace BySection** do not support this mapping and numbered keys are not permitted in `Add` mode 

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

A rule which selects no sections emits a warning. 

---

# Output Formatting

Transmute writes its output in one of two formats, controlled by the Syntax toggle (`-dontlint` via the command line):

### Syntax enabled (default)

The output is written as a winapp2.ini file:

- Entries are sorted alphabetically and grouped into the standard winapp2.ini category sections
- Keys within each entry are regrouped into winapp2.ini key order (categorization → detection → deletion routines), preserving their order within each group
- The standard winapp2.ini preamble comments (version line, entry count, license and project links) are regenerated at the top of the file

Because of the key regrouping, a key added by Add mode may not appear at the end of its entry in the output. See [Example 7](EXAMPLES.md#example-7-chaining-operations-together-1) for this in action.

If your base file is not a winapp2.ini file, disable Syntax.

### Syntax disabled (`-dontlint`)

The output is written as a plain ini file with sections in **alphabetical order** by name. The original section order of the base file is not preserved in either format.

### Comments are not preserved

Comments from both the base file and the source file will *not* be carried over into the output file. If you are overwriting your base file, you will lose any comments in it. 

---

# Flavorization

Flavorization applies a set of transformations to create specialized variants of ini files. Operations are applied in this order:

1.	Section Removal
2.	Key Name Removal
3.	Key Value Removal
4.	Section Replacement
5.	Key Replacement
6.	Section and Key Additions

See the [Flavorizer readme](Flavorizer/readme.md) for the dedicated sub-module which orchestrates this process.

### Example use cases

-	Creating winapp2.ini variants (eg. a CCleaner or BleachBit flavor)
-	Generating platform-specific configurations (eg. a Windows XP flavor)
-	Automated quality control corrections (eg. Correcting entries generated by winapp2ool)

---

# Source Files

Your source file must follow standard ini format:

```ini
[Section Name]
Key1=Value1
Key2=Value2

[Another Section]
Key1=Value1
``` 

Comments are welcome in your source files for your own reference, but they will not appear in the output — see [Output Formatting](#output-formatting).

# Tips & Best Practices

### Safety First

- Always test on copies of important files
- By default, output is saved to `winapp2-transmuted.ini` and your base file is untouched. If you point the save target at the base file, overwriting it is non-reversible without a backup. Winapp2ool does not create one
- Comments from either file do not get carried into the output

### Effective Source Files

- Keep source files focused on specific changes
- Comment your source files for future reference
- Matching is case-insensitive for both section and key names
- Disable the Syntax toggle (`-dontlint`) when working with files other than winapp2.ini

### Mode Selection

- Use Add for supplements and extensions
- Use Replace for updates and corrections
- Use Remove for cleanup and simplification
- Each sub-mode's "When to use" list under [Transmute (Primary) Modes](#transmute-primary-modes) gives finer-grained guidance for choosing within a mode

---

# Troubleshooting
|Error Message|Cause
|:-|:-
|"Target section not found in base file: [section] - no changes applied"|Source file references a section not in base file (only affects Replace/Remove)
|"Replacement target not found: {key} not found in {section}"|Source key doesn't exist in the base section (Replace ByKey mode)
|"Removal target not found: {key} not found in {section}"|Source key doesn't exist in the base section (Remove ByKey mode). {key} is the key name in ByName mode, or the KeyType=Value pair in ByValue mode
|"Transmute requires a source file..."|No `-2f` or preset flag was provided via the command line
|"[\*] cannot be used with Remove BySection (this would remove every section) - skipping"|A `[*]` section was provided while in Remove BySection mode
|"[\*] cannot be used with Replace BySection (a global section replacement is incoherent) - skipping"|A `[*]` section was provided while in Replace BySection mode
|"[\*] Refusing to add numbered key {key} to every section - global adds must be unnumbered"|A `[*]` section contains a numbered key (eg. `FileKey1`) in Add mode
|"[\*Map:] rules are only applied in Replace ByKey mode - skipping..."|`[*Map:]` sections were provided outside Replace ByKey mode
|"[Map: {label}] is malformed and will be skipped: {reason}"|The rule is missing `Match=` or `Replace=`, has more than one `Replace=`, contains an unrecognized key, or a value isn't a `Name=Value` pair
|"[Map: {label}] matched nothing in {file} - the rule may be stale"|No key in the base file matched any of the rule's criteria — check whether the values the rule targets still exist

---

# Migrating from Merge

### What happened to Merge?

As the functionality of Merge evolved, it no longer felt appropriate to refer to its output as the result of a "merger" necessarily. We now consider the resulting output a Transmutation. Nevertheless, this is still spiritually the Merge module. However, there are some important technical differences in how Transmute completes its task.

Merge fundamentally differed from Transmute in that Merge *always* applied additions, but with a much more limited capacity for conflict resolution between its "Add & Remove" and "Add & Replace" modes. Transmute makes much more granular changes to the files but can achieve the same output. 

To migrate to Transmute, you'll need to reconfigure your set of changes into categories by their effects under the new Transmute modes 

### New content
 This is the most common use case. Content you are adding (eg. custom entries or keys you wrote for your system) can all be placed together in one file. This functions mostly similarly to the way additions worked in Merge, with the new feature of being able to add individual keys to existing entries rather than requiring you to provide an entire section replacement. Apply additions by setting the Transmute mode to Add. 

### Replacement content
 Place any sections you want to have replace entries in winapp2.ini in a separate file from keys you want to replace within individual sections. Apply replacements by first setting the Transmute mode to Replace. The default Replace mode is By Key. Set the Replace mode to By Section to replace entire sections. 

### Removals
 Place any sections you want to remove entirely from winapp2.ini into a separate file from keys you want to remove from within individual sections. Likewise, place any keys you want to remove by their value in a separate file from keys you want to remove by their name. When removing entire sections, you need not provide any keys. Apply removals by first setting the Transmute mode to Remove. The default Remove mode is By Key. Set the Remove mode to By Section to remove entire sections. The default Key Removal mode is By Name. Set the Key Removal mode to By Value to remove keys by their value. 

##### This guide is provided as a general framework for decision making. For technical guidance on the commands required to migrate to Transmute from Merge, see the [Usage Examples](EXAMPLES.md) above
