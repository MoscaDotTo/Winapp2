# Transmute 

**Transmute** is a winapp2ool module that provides precise control over modifying ini files through three primary operations: Add, Replace, and Remove, with concise sub-modes where relevant. It enables users to apply targeted changes to configuration files at their chosen granularity (whole sections or individual keys) akin to "patching" winapp2.ini. 

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
1. [Requirements](#requirements)
2. [Quick Start](#quick-start)
3. [Menu Options](#menu-options)
4. [Transmute (Primary) Modes](#transmute-primary-modes)
   - [How Matching Works](#how-matching-works)
   - [Add Mode](#add-mode)
   - [Replace Mode](#replace-mode)
   - [Remove Mode](#remove-mode)
5. [Global Operations](#global-operations)
   - [The [*] Section](#the--section)
   - [Key Mapping Rules ([*Map:])](#key-mapping-rules-map)
6. [Output Formatting](#output-formatting)
7. [Flavorization](#flavorization)
8. [Command-Line Arguments](#command-line-arguments)
   - [Primary Mode](#primary-mode)
   - [Sub-Modes (Replace/Remove)](#sub-modes-replaceremove)
   - [Key Removal Modes](#key-removal-modes)
   - [Preset Source Files](#preset-source-files)
   - [Toggles](#toggles)
   - [File Selection](#file-selection)
   - [Examples](#examples)
9. [Source Files](#source-files)
10. [Tips & Best Practices](#tips--best-practices)
    - [Safety First](#safety-first)
    - [Effective Source Files](#effective-source-files)
    - [Mode Selection](#mode-selection)
11. [Troubleshooting](#troubleshooting)
12. [Usage Examples](#usage-examples)
    - [Adding Content](#adding-content)
      - [Example 1: Adding New Sections and Keys](#example-1-adding-new-sections-and-keys)
    - [Replacements](#replacements)
      - [Example 2: Replacing Key Values](#example-2-replacing-key-values)
      - [Example 3: Replacing Entire Sections](#example-3-replacing-entire-sections)
    - [Removals](#removals)
      - [Example 4: Removing Entire Sections](#example-4-removing-entire-sections)
      - [Example 5: Remove Keys By Name](#example-5-remove-keys-by-name)
      - [Example 6: Removing Keys By Value](#example-6-removing-keys-by-value)
    - [Advanced](#advanced)
      - [Example 7: Chaining operations together 1](#example-7-chaining-operations-together-1)
      - [Example 8: Chaining operations together 2](#example-8-chaining-operations-together-2)
      - [Example 9: Correcting syntax](#example-9-correcting-syntax)
13. [Migrating from Merge](#migrating-from-merge)
    - [New content](#new-content)
    - [Replacement content](#replacement-content)
    - [Removals](#removals-1)

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
| Toggle Global sections | Toggle treating `[*]` and `[*Map:]` source sections as [global operations](#global-operations) | Default: `True` |
| Change Transmute mode | Cycle through Add → Replace → Remove | Default: `Add` |
| Change Replace mode | Switch the Replace mode between BySection ↔ ByKey | Only visible in `Replace` mode |
| Change Remove mode | Switch the Remove mode between BySection ↔ ByKey | Only visible in `Remove` mode |
| Change Key Removal Mode | Switch between ByName ↔ ByValue | Only visible in `Remove` ByKey mode |
| Reset Settings | Restore all settings to their defaults | Only shown when settings have been changed |

\* See [Output Formatting](#output-formatting) for exactly what winapp2.ini formatting entails. Transmute does **not** run WinappDebug on its output. If you are working with winapp2.ini, you should separately run your transmute output through WinappDebug to ensure style/syntax correctness.  

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

Transmute is capable of performing operations over every section in the base file rather than to a single section matched by name. These operations are processed before any named sections in the same source file, and `[*Map:]` rules are applied before the `[*]` section. This ordering means specific per-section operations can refine the result of global ones.  

Both operations are controlled by the Global sections toggle (`-noglobal` via the command line). When disabled, `[*]` and `[*Map: ...]` are treated as ordinary named sections. 

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

##### Example

```ini
; Remove Edge's DetectFile from every entry that carries it (Remove - ByKey - ByValue)
[*]
DetectFile=%LocalAppData%\Microsoft\Edge*
```

## Key Mapping Rules ([*Map:])

`[*Map: <label>]` sections define key mapping rules: match a key anywhere in the base file by its KeyType and Value, and replace the whole key *including its Name*. This is the only Transmute operation that can change a key's Name.

Key mapping rules are recognized **only in Replace ByKey mode**. Under any other mode they are skipped with a warning

Each rule section contains:

| Key | Meaning | Notes |
|:-|:-|:-|
| `Match=<Name>=<Value>` | A match criteria | Repeatable as `Match1=`, `Match2=`, ... for many-to-one mappings. KeyTypes are compared with numbers stripped from both sides, so `Match=FileKey=...` matches any `FileKeyN` with that value |
| `Replace=<Name>=<Value>` | The full replacement key line | Exactly one per rule |

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
```

### Rule behavior

- Rules are applied in a single pass in file order with first-match-wins semantics: each base key is evaluated against its *original* value only, so a key replaced by an earlier rule is never re-matched by a later rule in the same run
- A rule which matches nothing anywhere in the base file emits a warning. This can be a sign that the rule has gone stale and the values it targets no longer exist in the base file
- A malformed rule is skipped with a warning while the remaining rules still apply. A rule is malformed if it has no `Match=` keys, a missing or duplicate `Replace=` key, an unrecognized key, or a `Match=`/`Replace=` value which is not itself a `Name=Value` pair

### Example use cases
- Maintaining the browser category mapping (`Section=` → `LangSecRef=`) for the CCleaner flavor as ~14 rules instead of hundreds of per-entry cohort sections that must be kept in lockstep across two files
- Stripping a detection key from every generated entry of a browser family in the System Ninja flavor, so new entries added by BrowserBuilder are covered automatically

---

# Output Formatting

Transmute writes its output in one of two formats, controlled by the Syntax toggle (`-dontlint` via the command line):

### Syntax enabled (default)

The output is written as a winapp2.ini file:

- Entries are sorted alphabetically and grouped into the standard winapp2.ini category sections
- Keys within each entry are regrouped into winapp2.ini key order (categorization → detection → deletion routines), preserving their order within each group
- The standard winapp2.ini preamble comments (version line, entry count, license and project links) are regenerated at the top of the file

Because of the key regrouping, a key added by Add mode may not appear at the end of its entry in the output. See [Example 7](#example-7-chaining-operations-together-1) for this in action.

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
| -noglobal | Treat `[*]` and `[*Map:]` source sections as ordinary section names instead of [global operations](#global-operations)

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

### Examples

|Command|Effect|
 |:-|:-| 
|winapp2ool -transmute -add -2f custom.ini|Add entries and keys from custom.ini to winapp2.ini, saving the result to winapp2-transmuted.ini
|winapp2ool -transmute -replace -bysection -c|Replace entire sections in winapp2.ini with ones from custom.ini, saving the result to winapp2-transmuted.ini
|winapp2ool -transmute -remove -bykey -byvalue -2f key_value_removals.ini |Remove keys from winapp2.ini sections defined by value in key_value_removals.ini, saving the result to winapp2-transmuted.ini
|winapp2ool -transmute -remove -bysection -2f section_removals.ini -3f cleaned.ini | Remove sections from winapp2.ini defined in section_removals.ini and save the result  to cleaned.ini

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

# Usage Examples

To drive some of our examples, we'll take a look at some of the work done by winapp2ool to apply corrections to the output of the Browser Builder. These files can be found [here](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/BrowserBuilder), but relevant lines of code will be provided on this page.  

###### Note: When Syntax is enabled (the default), Transmute regenerates the standard winapp2.ini preamble comments at the top of the output file. The example outputs below omit this preamble for brevity.

## Adding Content

### Example 1: Adding New Sections and Keys

**Context** 

Browser Builder creates the browser entries using a scaffold framework. While this simplifies the overwhelming amount of work involved in maintaining web browser entries, it fails to cover some of the nuances of individual browsers. 

**Intent**

We want to increase the coverage of 360 Secure Browser by adding a `FileKey` to the generated `[360 Secure Browser Bookmarked Websites *]` entry. Likewise, we want to add an entry, `[360 Secure Browser Web Browsing History Backups *]`, to provide additional coverage for a feature unique to this browser alongside the standard Browser Builder output. 

**Files**
###### **Base file (`browsers.ini`)**
```ini
[360 Secure Browser Bookmarked Websites *]
Section=.360 Secure Browser Web Browser
DetectFile=%AppData%\360se6\User Data
FileKey1=%AppData%\360se6\User Data\*|bookmarks;BookmarkMergedSurfaceOrdering
FileKey2=%AppData%\360se6\User Data\*\power_bookmarks|*|REMOVESELF
```

###### **Source file (`browser_additions.ini`)**

```ini
; This key will be added to the generated entry of the same name in the base file to provide better coverage
[360 Secure Browser Bookmarked Websites *]
FileKey=%AppData%\360se6\User Data\*|360Bookmarks*

; This entire entry will be added to the base file 
[360 Secure Browser Web Browsing History Backups *]
Section=.360 Secure Browser Web Browser
DetectFile=%AppData%\360se6\User Data
FileKey1=%AppData%\360se6\User Data\*\HisDailyBackup|*|REMOVESELF
```

###### Note: *Unlike* replacements and removals, entry and key additions can be provided in just a single file 

**Command**
```
winapp2ool -transmute -add -1f browsers.ini -2f browser_additions.ini -3f browsers.ini 
```
###### Note: `Add` is the default transmute mode and technically the `-add` argument could be omitted here but is provided for the utmost clarity

**Output**

###### **Output file (`browsers.ini`) after transmutation**

```ini
[360 Secure Browser Bookmarked Websites *]
Section=.360 Secure Browser Web Browser
DetectFile=%AppData%\360se6\User Data
FileKey1=%AppData%\360se6\User Data\*|bookmarks;BookmarkMergedSurfaceOrdering
FileKey2=%AppData%\360se6\User Data\*\power_bookmarks|*|REMOVESELF
FileKey=%AppData%\360se6\User Data\*|360Bookmarks*

[360 Secure Browser Web Browsing History Backups *]
Section=.360 Secure Browser Web Browser
DetectFile=%AppData%\360se6\User Data
FileKey1=%AppData%\360se6\User Data\*\HisDailyBackup|*|REMOVESELF
```

**Explanation**
- The base file is browsers.ini
- The source file is browser_additions.ini
- The output file is browsers.ini (overwriting the base file)
- `[360 Secure Browser Web Browsing History Backups *]` is added to the base file as defined in the source file 
- `[360 Secure Browser Bookmarked Websites *]` in the base file has `FileKey` added with value `%AppData%\360se6\User Data\*|360Bookmarks*` from the source file
- Sections in the base file not defined in the source file remain unchanged  

**Notes**

Transmute does not do any work to ensure correct key numbering on its own, it adds keys as written (eg. `FileKey` above ). Winapp2.ini files should be run through WinappDebug to correct their syntax.  

---

## Replacements

### Example 2: Replacing Key Values

**Context** 

Browser Builder produces some `DetectFile` keys which are compatible with winapp2ool but *not* with CCleaner because CCleaner does not support wildcards in parent paths to the `DetectFile`. This means that a key such as `DetectFile2=%LocalAppData%\Packages\Mozilla.Firefox_*\LocalCache\Roaming\Mozilla\Firefox\Profiles` will not be correctly interpreted by CCleaner. We know that our generated Mozilla Firefox entries always contain this `DetectFile2` and want to replace it with a value compatible with CCleaner.

**Intent**

We want to replace the value of `DetectFile2` in `[Mozilla Firefox Autocomplete History *] ` after it is generated with `%LocalAppData%\Packages\Mozilla.Firefox_*` which will be correctly interpreted by CCleaner 

**Files**

###### **Base file (`browsers.ini`)**
```ini
[Mozilla Firefox Autocomplete History *]
Section=Mozilla Firefox Web Browser
DetectFile1=%AppData%\Mozilla\Firefox\Profiles
DetectFile2=%LocalAppData%\Packages\Mozilla.Firefox_*\LocalCache\Roaming\Mozilla\Firefox\Profiles
FileKey1=%AppData%\Mozilla\Firefox\Profiles\*|formhistory*
FileKey2=%LocalAppData%\Packages\Mozilla.Firefox_*\LocalCache\Roaming\Mozilla\Firefox\Profiles\*|formhistory*
```

###### **Source file (`browser_key_replacements.ini`)**
```ini
[Mozilla Firefox Autocomplete History *]
DetectFile2=%LocalAppData%\Packages\Mozilla.Firefox_*
```

**Command**
```
winapp2ool -transmute -replace -bykey -1f browsers.ini -2f browser_key_replacements.ini -3f browsers.ini
```
###### Note: `By Key` is the default replace mode and technically the `-bykey` argument could be omitted here but is provided for the utmost clarity

**Output**

###### **Output file (`browsers.ini`) after transmutation**

```ini
[Mozilla Firefox Autocomplete History *]
Section=Mozilla Firefox Web Browser
DetectFile1=%AppData%\Mozilla\Firefox\Profiles
DetectFile2=%LocalAppData%\Packages\Mozilla.Firefox_*
FileKey1=%AppData%\Mozilla\Firefox\Profiles\*|formhistory*
FileKey2=%LocalAppData%\Packages\Mozilla.Firefox_*\LocalCache\Roaming\Mozilla\Firefox\Profiles\*|formhistory*
```

**Explanation**
- The base file is browsers.ini
- The source file is browser_key_replacements.ini
- The output file is browsers.ini (overwriting the base file)
- `DetectFile2` in `[Mozilla Firefox Autocomplete History *]` in the base file has its value replaced with the value of `DetectFile2` from `[Mozilla Firefox Autocomplete History *]` in the source file 
- Sections in the base file not defined in the source file remain unchanged  

**Notes**

 This mode replaces key *values* by matching the key's `Name`, which is the entire text to the left of the `=`. The base file's key name itself is untouched. It may produce more consistent results for numbered keys to replace a key's value by first removing it and then adding a new key with the desired value. 

---

### Example 3: Replacing Entire Sections

**Context**

You have installed a game, *And Yet It Moves*, from GOG. The winapp2.ini entry only targets the Steam version. You have written a suitable replacement. 

**Intent**

We want to replace the winapp2.ini version of `[And Yet It Moves *]` with our custom copy we maintain separately

**Files**

###### **Base file (`winapp2.ini`)**
```ini
[And Yet It Moves *]
Section=Games
Detect=HKCU\Software\Valve\Steam\Apps\18700
FileKey1=%AppData%\Broken Rules\And Yet It Moves Steam|console.log
```
###### **Source file (`section_replacements.ini`)**
```ini
[And Yet It Moves *]
Section=GOG Games
DetectFile=%SystemDrive%\GOG\Broken Rules\And Yet It Moves
FileKey1=%SystemDrive%\GOG\Broken Rules\And Yet It Moves|console.log
```

**Command**
```
winapp2ool -transmute -replace -bysection -2f section_replacements.ini -3f winapp2.ini
```
###### Note: Without `-3f winapp2.ini`, the output would be saved to `winapp2-transmuted.ini` and the base file would be left untouched

**Output**

###### **Output file (`winapp2.ini`) after transmutation**

```ini
[And Yet It Moves *]
Section=GOG Games
DetectFile=%SystemDrive%\GOG\Broken Rules\And Yet It Moves
FileKey1=%SystemDrive%\GOG\Broken Rules\And Yet It Moves|console.log
```

**Explanation**
- The base file is winapp2.ini (default)
- The source file is section_replacements.ini
- The output file is winapp2.ini (overwriting the base file)
- `[And Yet It Moves *]` in the base file is entirely replaced by `[And Yet It Moves *]` from the source file 
- Sections in the base file not defined in the source file remain unchanged  

---

## Removals

### Example 4: Removing Entire Sections

**Context**

In the first version, entries generated by Browser Builder could be incomplete or target features not implemented in a particular browser. Rather than ship them targeting nothing, we chose to prune them from the set of generated entries before combining them into winapp2.ini. Arc implements its pinned tabs storage as a part of a JSON file shared with persistent configuration which winapp2.ini doesn't support cleaning non-destructively.   

###### Note: In the current version of winapp2ool, this entry is no longer generated because it lacks the full corpus of data required for it. This example is left here for its use in the readme, but no longer reflects the current state of the browser builder output.

**Intent**

We want to remove `[Arc Pinned Tabs *]` from browsers.ini because there is nothing we can do to make it viable. 

**Files**

###### **Base file (browsers.ini)**
```ini
; This entry is generated without a working RegKey because Arc's pinned tabs live in a shared JSON file
[Arc Pinned Tabs *]
Section=Arc Web Browser
DetectFile=%LocalAppData%\Packages\TheBrowserCompany.Arc_ttt1ap7aakyb4\LocalCache\Local\Arc\User Data
RegKey1=

[Arc Privacy Sandbox *]
Section=Arc Web Browser
DetectFile=%LocalAppData%\Packages\TheBrowserCompany.Arc_ttt1ap7aakyb4\LocalCache\Local\Arc\User Data
FileKey1=%LocalAppData%\Packages\TheBrowserCompany.Arc_ttt1ap7aakyb4\LocalCache\Local\Arc\User Data|*first_party_sets*
FileKey2=%LocalAppData%\Packages\TheBrowserCompany.Arc_ttt1ap7aakyb4\LocalCache\Local\Arc\User Data\*|BrowsingTopics*;Conversions*;InterestGroups;MediaDeviceSalts;SharedStorage*;PrivateAggregation*
FileKey3=%LocalAppData%\Packages\TheBrowserCompany.Arc_ttt1ap7aakyb4\LocalCache\Local\Arc\User Data\*\Network|Trust Tokens*
FileKey4=%LocalAppData%\Packages\TheBrowserCompany.Arc_ttt1ap7aakyb4\LocalCache\Local\Arc\User Data\CookieReadinessList|*|REMOVESELF
FileKey5=%LocalAppData%\Packages\TheBrowserCompany.Arc_ttt1ap7aakyb4\LocalCache\Local\Arc\User Data\FirstPartySetsPreloaded|*|REMOVESELF
FileKey6=%LocalAppData%\Packages\TheBrowserCompany.Arc_ttt1ap7aakyb4\LocalCache\Local\Arc\User Data\PrivacySandboxAttestationsPreloaded|*|REMOVESELF
FileKey7=%LocalAppData%\Packages\TheBrowserCompany.Arc_ttt1ap7aakyb4\LocalCache\Local\Arc\User Data\ProbabilisticRevealTokenRegistry|*|REMOVESELF
FileKey8=%LocalAppData%\Packages\TheBrowserCompany.Arc_ttt1ap7aakyb4\LocalCache\Local\Arc\User Data\TrustTokenKeyCommitments|*|REMOVESELF
```

###### **Source file (`browser_section_removals.ini`)**

```ini
; Arc Browser stores pinned tabs in the sidebar configuration (JSON)
[Arc Pinned Tabs *]
```

###### Note: There is no need to include keys when removing entire sections.  

**Command**
```
winapp2ool -transmute -remove -bysection -1f browsers.ini -2f browser_section_removals.ini -3f browsers.ini
```

**Output** 

###### **Output file (`browsers.ini`) after transmutation**

```ini
[Arc Privacy Sandbox *]
Section=Arc Web Browser
DetectFile=%LocalAppData%\Packages\TheBrowserCompany.Arc_ttt1ap7aakyb4\LocalCache\Local\Arc\User Data
FileKey1=%LocalAppData%\Packages\TheBrowserCompany.Arc_ttt1ap7aakyb4\LocalCache\Local\Arc\User Data|*first_party_sets*
FileKey2=%LocalAppData%\Packages\TheBrowserCompany.Arc_ttt1ap7aakyb4\LocalCache\Local\Arc\User Data\*|BrowsingTopics*;Conversions*;InterestGroups;MediaDeviceSalts;SharedStorage*;PrivateAggregation*
FileKey3=%LocalAppData%\Packages\TheBrowserCompany.Arc_ttt1ap7aakyb4\LocalCache\Local\Arc\User Data\*\Network|Trust Tokens*
FileKey4=%LocalAppData%\Packages\TheBrowserCompany.Arc_ttt1ap7aakyb4\LocalCache\Local\Arc\User Data\CookieReadinessList|*|REMOVESELF
FileKey5=%LocalAppData%\Packages\TheBrowserCompany.Arc_ttt1ap7aakyb4\LocalCache\Local\Arc\User Data\FirstPartySetsPreloaded|*|REMOVESELF
FileKey6=%LocalAppData%\Packages\TheBrowserCompany.Arc_ttt1ap7aakyb4\LocalCache\Local\Arc\User Data\PrivacySandboxAttestationsPreloaded|*|REMOVESELF
FileKey7=%LocalAppData%\Packages\TheBrowserCompany.Arc_ttt1ap7aakyb4\LocalCache\Local\Arc\User Data\ProbabilisticRevealTokenRegistry|*|REMOVESELF
FileKey8=%LocalAppData%\Packages\TheBrowserCompany.Arc_ttt1ap7aakyb4\LocalCache\Local\Arc\User Data\TrustTokenKeyCommitments|*|REMOVESELF
```

**Explanation**

- The base file is browsers.ini
- The source file is browser_section_removals.ini
- The output file is browsers.ini (overwriting the base file)
- `[Arc Pinned Tabs *]` is removed from the base file 
- Sections in the base file not defined in the source file remain unchanged  

---

### Example 5: Remove Keys By Name

**Context**

You are developing a tool which implements winapp2.ini. You are aware that `Section` is a categorizer, but your tool doesn't use these categories. As such, you want to remove the Section key from a particular entry, `[And Yet It Moves *]`, before passing it into your tool. 

**Intent**

We want to remove the `Section` key from `[And Yet It Moves *]`

**Files**

###### **Base file (`winapp2.ini`)**
```ini
[And Yet It Moves *]
Section=Games
Detect=HKCU\Software\Valve\Steam\Apps\18700
FileKey1=%AppData%\Broken Rules\And Yet It Moves Steam|console.log
```

##### **Source file (`key_name_removals.ini`)**
```ini
[And Yet It Moves *]
Section=Games
```

###### Note: In ByName mode the value provided in the source file is ignored. `Section=` would remove the key just the same. Providing the value anyway keeps the source file self-documenting.

**Command**
```
winapp2ool -transmute -remove -bykey -byname -2f key_name_removals.ini -3f winapp2.ini
```

**Output**

###### **Output file (`winapp2.ini`) after transmutation**

```ini
[And Yet It Moves *]
Detect=HKCU\Software\Valve\Steam\Apps\18700
FileKey1=%AppData%\Broken Rules\And Yet It Moves Steam|console.log
```

**Explanation**
- The base file is winapp2.ini
- The source file is key_name_removals.ini
- The output file is winapp2.ini (overwriting the base file)
- `Section=Games` is removed from `[And Yet It Moves *]` in the base file 
- Sections in the base file not defined in the source file remain unchanged 

---

### Example 6: Removing Keys By Value

**Context**

When Browser Builder generates entries for DuckDuckGo Browser, it produces two `DetectFile` keys, neither of which are compatible with CCleaner due to containing wildcards in a parent path, eg. `DetectFile=%LocalAppData%\Packages\*DuckDuckGo*Browser*\LocalState\EBWebView`. Rather than remove one and replace the value of the other, we choose to address this by adding a value which captures both paths. The first step of this process is removing the incompatible keys.

**Intent**

We want to remove the incompatible `DetectFile` values without having to know which key number (eg. the **1** in `DetectFile1`) is associated with the value.

**Files**

###### **Base file (`browsers.ini`)**
```ini
; The DetectFiles in this entry won't work with CCleaner
[DuckDuckGo Autofill Data *]
Section=DuckDuckGo Web Browser
DetectFile1=%LocalAppData%\Packages\*DuckDuckGo*Browser*\LocalState\EBWebView
DetectFile2=%LocalAppData%\Packages\*DuckDuckGo*Browser*\LocalState\internalEnvironment\EBWebView
FileKey1=%LocalAppData%\Packages\*DuckDuckGo*Browser*\LocalState\EBWebView\*|*Web Data
FileKey2=%LocalAppData%\Packages\*DuckDuckGo*Browser*\LocalState\EBWebView\*\AutoFill*|*|REMOVESELF
FileKey3=%LocalAppData%\Packages\*DuckDuckGo*Browser*\LocalState\EBWebView\AutoFill*|*|REMOVESELF
FileKey4=%LocalAppData%\Packages\*DuckDuckGo*Browser*\LocalState\internalEnvironment\EBWebView\*|*Web Data
FileKey5=%LocalAppData%\Packages\*DuckDuckGo*Browser*\LocalState\internalEnvironment\EBWebView\*\AutoFill*|*|REMOVESELF
FileKey6=%LocalAppData%\Packages\*DuckDuckGo*Browser*\LocalState\internalEnvironment\EBWebView\AutoFill*|*|REMOVESELF
```

###### **Source file (`browser_value_removals.ini`)**
```ini
; We don't want these values! We'll provide replacements in a separate file 
[DuckDuckGo Autofill Data *]
DetectFile=%LocalAppData%\Packages\*DuckDuckGo*Browser*\LocalState\EBWebView
DetectFile=%LocalAppData%\Packages\*DuckDuckGo*Browser*\LocalState\internalEnvironment\EBWebView
```

**Command**
```
winapp2ool -transmute -remove -bykey -byvalue -1f browsers.ini -2f browser_value_removals.ini -3f browsers.ini
```

**Output**

###### **Output file (`browsers.ini`) after transmutation**
```ini
[DuckDuckGo Autofill Data *]
Section=DuckDuckGo Web Browser
FileKey1=%LocalAppData%\Packages\*DuckDuckGo*Browser*\LocalState\EBWebView\*|*Web Data
FileKey2=%LocalAppData%\Packages\*DuckDuckGo*Browser*\LocalState\EBWebView\*\AutoFill*|*|REMOVESELF
FileKey3=%LocalAppData%\Packages\*DuckDuckGo*Browser*\LocalState\EBWebView\AutoFill*|*|REMOVESELF
FileKey4=%LocalAppData%\Packages\*DuckDuckGo*Browser*\LocalState\internalEnvironment\EBWebView\*|*Web Data
FileKey5=%LocalAppData%\Packages\*DuckDuckGo*Browser*\LocalState\internalEnvironment\EBWebView\*\AutoFill*|*|REMOVESELF
FileKey6=%LocalAppData%\Packages\*DuckDuckGo*Browser*\LocalState\internalEnvironment\EBWebView\AutoFill*|*|REMOVESELF
```

**Explanation**
- The base file is browsers.ini
- The source file is browser_value_removals.ini
- The output file is browsers.ini (overwriting the base file)
- Any `DetectFile` key with a value provided from `[DuckDuckGo Autofill Data *]` in the source file is removed from `[DuckDuckGo Autofill Data *]` in the base file
- Sections in the base file not defined in the source file remain unchanged 

---

## Advanced

### Example 7: Chaining operations together 1

**Context**

Continuing from **Example 6**, lets complete the task of both removing an unwanted value and replacing it with a new one. The entry as it is at the end of the last example possesses no detection criteria, we want to add one. 

**Intent**

We want to add a functional `DetectFile` to replace the two `DetectFile` keys we remove in **Example 6.**

**Files**

###### **Base file (`browsers.ini`)**
```ini
; The DetectFiles in this entry won't work with CCleaner
[DuckDuckGo Autofill Data *]
Section=DuckDuckGo Web Browser
DetectFile1=%LocalAppData%\Packages\*DuckDuckGo*Browser*\LocalState\EBWebView
DetectFile2=%LocalAppData%\Packages\*DuckDuckGo*Browser*\LocalState\internalEnvironment\EBWebView
FileKey1=%LocalAppData%\Packages\*DuckDuckGo*Browser*\LocalState\EBWebView\*|*Web Data
FileKey2=%LocalAppData%\Packages\*DuckDuckGo*Browser*\LocalState\EBWebView\*\AutoFill*|*|REMOVESELF
FileKey3=%LocalAppData%\Packages\*DuckDuckGo*Browser*\LocalState\EBWebView\AutoFill*|*|REMOVESELF
FileKey4=%LocalAppData%\Packages\*DuckDuckGo*Browser*\LocalState\internalEnvironment\EBWebView\*|*Web Data
FileKey5=%LocalAppData%\Packages\*DuckDuckGo*Browser*\LocalState\internalEnvironment\EBWebView\*\AutoFill*|*|REMOVESELF
FileKey6=%LocalAppData%\Packages\*DuckDuckGo*Browser*\LocalState\internalEnvironment\EBWebView\AutoFill*|*|REMOVESELF
```

###### **Source file (`browser_value_removals.ini`)**
```ini
; We don't want these values! We'll provide replacements in browser_additions.ini 
[DuckDuckGo Autofill Data *]
DetectFile=%LocalAppData%\Packages\*DuckDuckGo*Browser*\LocalState\EBWebView
DetectFile=%LocalAppData%\Packages\*DuckDuckGo*Browser*\LocalState\internalEnvironment\EBWebView
```

###### **Source file(`browser_additions.ini`)**
```ini
; This one value replaces the two we remove in browser_value_removals.ini in our final output 
[DuckDuckGo Autofill Data *]
DetectFile=%LocalAppData%\Packages\*DuckDuckGo*Browser*
```

**Commands**
```
winapp2ool -transmute -remove -bykey -byvalue -1f browsers.ini -2f browser_value_removals.ini -3f browsers.ini
winapp2ool -transmute -add -1f browsers.ini -2f browser_additions.ini -3f browsers.ini 
```

**Output**

###### Output file (`browsers.ini`) after transmutation
```ini
[DuckDuckGo Autofill Data *]
Section=DuckDuckGo Web Browser
DetectFile=%LocalAppData%\Packages\*DuckDuckGo*Browser*
FileKey1=%LocalAppData%\Packages\*DuckDuckGo*Browser*\LocalState\EBWebView\*|*Web Data
FileKey2=%LocalAppData%\Packages\*DuckDuckGo*Browser*\LocalState\EBWebView\*\AutoFill*|*|REMOVESELF
FileKey3=%LocalAppData%\Packages\*DuckDuckGo*Browser*\LocalState\EBWebView\AutoFill*|*|REMOVESELF
FileKey4=%LocalAppData%\Packages\*DuckDuckGo*Browser*\LocalState\internalEnvironment\EBWebView\*|*Web Data
FileKey5=%LocalAppData%\Packages\*DuckDuckGo*Browser*\LocalState\internalEnvironment\EBWebView\*\AutoFill*|*|REMOVESELF
FileKey6=%LocalAppData%\Packages\*DuckDuckGo*Browser*\LocalState\internalEnvironment\EBWebView\AutoFill*|*|REMOVESELF
```

**Explanation**

Two separate transmutations are conducted, each overwriting browsers.ini so that the second builds on the first:

| Step | Mode | Source file | Effect |
|:-|:-|:-|:-|
| 1 | Remove ByKey ByValue | browser_value_removals.ini | Both incompatible `DetectFile` keys are removed from `[DuckDuckGo Autofill Data *]` |
| 2 | Add | browser_additions.ini | `DetectFile=%LocalAppData%\Packages\*DuckDuckGo*Browser*` is added to `[DuckDuckGo Autofill Data *]` |

- Sections in the base file not defined in the source files remain unchanged 

**Notes**

The added `DetectFile` appears alongside the entry's other detection keys rather than at the end of the entry: Add mode appends keys to the end of the section, but saving with Syntax enabled regroups keys by type (see [Output Formatting](#output-formatting)).

---

### Example 8: Chaining operations together 2

**Context**

As part of the switch to non-ccleaner as default, winapp2.ini recently declared separate sections for each of the web browsers. This conflicts with the old configuration which placed them, at least for CCleaner, together with the CCleaner entries for the same browser. We want to undo this as part of creating a CCleaner flavor of winapp2.ini 

**Intent**

We want to remove the `Section` keys from a particular web browser and add a `LangSecRef` pointing to the appropriate CCleaner section. 

**Files**

There are 22 generated Vivaldi entries in the base file; three are shown here. The remaining entries follow the same pattern and receive the same treatment.

###### **Base file (`browsers.ini`)**
```ini
[Vivaldi Autofill Data *]
Section=Vivaldi Web Browser
DetectFile=%LocalAppData%\Vivaldi\User Data
FileKey1=%LocalAppData%\Vivaldi\User Data\*|*Web Data
FileKey2=%LocalAppData%\Vivaldi\User Data\*\AutoFill*|*|REMOVESELF
FileKey3=%LocalAppData%\Vivaldi\User Data\AutoFill*|*|REMOVESELF

[Vivaldi Autoplay Preferences *]
Section=Vivaldi Web Browser
DetectFile=%LocalAppData%\Vivaldi\User Data
FileKey1=%LocalAppData%\Vivaldi\User Data\MEIPreload|*|REMOVESELF

[Vivaldi Bookmark Backups *]
Section=Vivaldi Web Browser
DetectFile=%LocalAppData%\Vivaldi\User Data
FileKey1=%LocalAppData%\Vivaldi\User Data\*|Bookmarks.bak

; 19 further Vivaldi entries, all carrying Section=Vivaldi Web Browser
```

###### **Source file (`browser_value_removals.ini`)**
```ini
[Vivaldi Autofill Data *]
Section=Vivaldi Web Browser

[Vivaldi Autoplay Preferences *]
Section=Vivaldi Web Browser

[Vivaldi Bookmark Backups *]
Section=Vivaldi Web Browser

; one section per remaining Vivaldi entry, each listing the same Section key to remove 
```

###### **Source file(`browser_additions.ini`)**
```ini
[Vivaldi Autofill Data *]
LangSecRef=3033

[Vivaldi Autoplay Preferences *]
LangSecRef=3033

[Vivaldi Bookmark Backups *]
LangSecRef=3033

; one section per remaining Vivaldi entry, each adding the same LangSecRef
```

**Commands**
```
winapp2ool -transmute -remove -bykey -byvalue -1f browsers.ini -2f browser_value_removals.ini -3f browsers.ini
winapp2ool -transmute -add -1f browsers.ini -2f browser_additions.ini -3f browsers.ini 
```

###### Note: `By Key` is the default remove mode and technically the `-bykey` argument could be omitted here but is provided for the utmost clarity

**Output**

###### Output file (`browsers.ini`) after transmutation
```ini
[Vivaldi Autofill Data *]
LangSecRef=3033
DetectFile=%LocalAppData%\Vivaldi\User Data
FileKey1=%LocalAppData%\Vivaldi\User Data\*|*Web Data
FileKey2=%LocalAppData%\Vivaldi\User Data\*\AutoFill*|*|REMOVESELF
FileKey3=%LocalAppData%\Vivaldi\User Data\AutoFill*|*|REMOVESELF

[Vivaldi Autoplay Preferences *]
LangSecRef=3033
DetectFile=%LocalAppData%\Vivaldi\User Data
FileKey1=%LocalAppData%\Vivaldi\User Data\MEIPreload|*|REMOVESELF

[Vivaldi Bookmark Backups *]
LangSecRef=3033
DetectFile=%LocalAppData%\Vivaldi\User Data
FileKey1=%LocalAppData%\Vivaldi\User Data\*|Bookmarks.bak

; ... 19 further Vivaldi entries, likewise now carrying LangSecRef=3033 ...
```

**Explanation**

Two separate transmutations are conducted, each overwriting browsers.ini so that the second builds on the first:

| Step | Mode | Source file | Effect |
|:-|:-|:-|:-|
| 1 | Remove ByKey ByValue | browser_value_removals.ini | The `Section=Vivaldi Web Browser` key is removed from each listed entry |
| 2 | Add | browser_additions.ini | `LangSecRef=3033` is added to each listed entry |

- Sections in the base file not defined in the source files remain unchanged 

---

### Example 9: Correcting syntax

**Context** 

Lets revisit [Example 1](#example-1-adding-new-sections-and-keys), specifically the malformatted winapp2.ini formatting produced by the Add operation. We want to ensure that any changes we make to a winapp2.ini file produce a new winapp2.ini with valid formatting. 


**Intent**

We want to increase the coverage of 360 Secure Browser by adding a `FileKey` to the generated `[360 Secure Browser Bookmarked Websites *]` entry. We then want to correct the syntax such that it is correctly interpreted 

**Files**
###### **Base file (`browsers.ini`)**
```ini
[360 Secure Browser Bookmarked Websites *]
Section=.360 Secure Browser Web Browser
DetectFile=%AppData%\360se6\User Data
FileKey1=%AppData%\360se6\User Data\*|bookmarks;BookmarkMergedSurfaceOrdering
FileKey2=%AppData%\360se6\User Data\*\power_bookmarks|*|REMOVESELF
```

###### **Source file (`browser_additions.ini`)**

```ini
[360 Secure Browser Bookmarked Websites *]
FileKey=%AppData%\360se6\User Data\*|360Bookmarks*
``` 

**Command**
```
winapp2ool -transmute -add -1f browsers.ini -2f browser_additions.ini -3f browsers.ini 
winapp2ool -debug -c -1f browsers.ini -3f browsers.ini
```
###### Note: WinappDebug does not by default perform the optimization shown below, you must manually enable the "Optimizations" lint rule in the WinappDebug Scan Settings first

**Output**

###### **Output file (`browsers.ini`) after transmutation and linting**

```ini
[360 Secure Browser Bookmarked Websites *]
Section=.360 Secure Browser Web Browser
DetectFile=%AppData%\360se6\User Data
FileKey1=%AppData%\360se6\User Data\*|360Bookmarks*;bookmarks;BookmarkMergedSurfaceOrdering
FileKey2=%AppData%\360se6\User Data\*\power_bookmarks|*|REMOVESELF
```

**Explanation**
- The base file is browsers.ini
- The source file is browser_additions.ini
- The output file is browsers.ini (overwriting the base file)
- `[360 Secure Browser Bookmarked Websites *]` in the base file has `FileKey` added with value `%AppData%\360se6\User Data\*|360Bookmarks*` from the source file
- Sections in the base file not defined in the source file remain unchanged  
- WinappDebug is invoked
- The input file is browsers.ini 
- The output file is browsers.ini (overwriting the input file)
- WinappDebug detects that the added `FileKey` points to the same location as the existing `FileKey1` and merges its parameters into the existing key and removes the now-unneeded `FileKey` which was just added
- The style and syntax of the entry is corrected if there are any additional issues

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

##### This guide is provided as a general framework for decision making. For technical guidance on the commands required to migrate to Transmute from Merge, see the [Usage Examples](#usage-examples) above
