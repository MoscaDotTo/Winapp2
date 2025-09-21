# Transmute 

**Transmute** is a winapp2ool module that provides precise control over modifying ini files through three primary operations: Add, Replace, and Remove, with concise sub-modes where relevant. It enables users to apply targeted changes to configuration files with granular conflict resolution, akin to "patching" winapp2.ini. 

### What happened to Merge?

As the functionality of Merge evolved, it no longer felt appropriate to refer to its output as the result of a "merger" necessarily. We now consider the resulting output a Transmutation. Nevertheless, this is still spiritually the Merge module. However, there are some fundamental differences in how Transmute completes its task. See the [Migrating From Merge](#migrating) section below for guidance on adjusting your patch files to be compatible with Transmute. 

### What does Transmute do?
Transmute allows you to modify one ini file (the "base" file) using instructions from another ini file (the "source" file). Think of it as a sophisticated merge tool that can add new content, replace existing content, or remove unwanted content with precision.

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
   - [Case Sensitivity](#case-sensitivity)
   - [Add Mode](#add-mode)
   - [Replace Mode](#replace-mode)
   - [Remove Mode](#remove-mode)
5. [Flavorization](#flavorization)
6. [Command-Line Arguments](#command-line-arguments)
   - [Primary Mode](#primary-mode)
   - [Sub-Modes (Replace/Remove)](#sub-modes-replace-remove)
   - [Key Removal Modes](#key-removal-modes)
   - [Preset Source Files](#preset-source-files)
   - [Toggles](#toggles)
   - [File Selection](#file-selection)
   - [Examples](#examples)
7. [Source Files](#source-files)
8. [Tips & Best Practices](#tips--best-practices)
   - [Safety First](#safety-first)
   - [Effective Source Files](#effective-source-files)
   - [Mode Selection](#mode-selection)
   - [Key Removal Strategy](#key-removal-strategy)
9. [Troubleshooting](#troubleshooting)
10. [Migrating from Merge](#migrating-from-merge)
    - [New content](#new-content)
    - [Replacement content](#replacement-content)
    - [Removals](#removals)
11. [Usage Examples](#usage-examples)
    - [Adding Content](#adding-content)
      - [Example 1: Adding New Sections and Keys](#example-1-adding-new-sections-and-keys)
    - [Replacements](#replacements)
      - [Example 2: Replacing Key Values](#example-2-replacing-key-values)
      - [Example 3: Replacing Entire Sections](#example-3-replacing-entire-sections)
    - [Removals](#removals-1)
      - [Example 4: Removing Entire Sections](#example-4-removing-entire-sections)
      - [Example 5: Remove Keys By Name](#example-5-remove-keys-by-name)
      - [Example 6: Removing Keys By Value](#example-6-removing-keys-by-value)
    - [Advanced](#advanced)
      - [Example 7: Chaining operations together 1](#example-7-chaining-operations-together-1)
      - [Example 8: Chaining operations together 2](#example-8-chaining-operations-together-2)
      - [Example 9: Correcting syntax](#example-9-correcting-syntax)

---

# [Requirements](#requirements)
- A base ini file you wish to modify 
- A source ini file providing the modifications

---

# [Quick Start](#quick-start)

### Common Workflow
1.	Create your base file 
2.	Create your source file(s) with your desired changes
3.	Choose your transmutation mode and options
4.	Run Transmute to apply the changes
5.  Reconfigure and run as many times as necessary to achieve your desired output 

*If you are applying multiple transmutations in succession, you may wish to investigate the Flavorizer*

---

# [Menu Options](#menu-options)

|Option|Effect|Notes|
|:-|:-|:-| 
| Run (default)           | Apply the current transmutation settings              |                                     |
| Open Flavorizer         | Opens the Flavorizer sub-modules                      | Provides a UI for chain operations  |
| Change Transmute mode   | Cycles through Add → Replace → Remove                 | Default: `Add`                      |
| Change Replace mode     | Switches the Replace mode between BySection ↔ ByKey   | Only visible in `Replace` mode      |
| Change Remove mode      | Switches the Remove mode between BySection ↔ ByKey    | Only visible in `Remove` mode       |
| Change Key Removal Mode | Switches between ByName ↔ ByValue                     | Only visible in `Remove` ByKey mode |  
| Toggle Syntax           | Toggles formatting the output as winapp2.ini \*       | Default: `True`                     |
| Choose base file        | Select the file to be modified                        | Default: `winapp2.ini`              | 
| Choose source file      | Select the file containing modifications              |                                     |
| Choose save target      | Select output location                                | Default: `winapp2.ini`              |

\* winapp2.ini formatting means respecting the winapp2.ini ordering of sections and leading comments/information within in the output file. Transmute does **not** run WinappDebug on its output. If you are working with winapp2.ini, you should separately run your transmute output through WinappDebug to ensure style/syntax correctness.  

---

# [Transmute (Primary) Modes](#primary)

## [Case Sensitivity](#case-sensitivity)

- Section names: case-sensitive
- Replace ByKey: key name matching is case-insensitive
- Remove ByKey (ByName): key name matching is case-insensitive
- Remove ByKey (ByValue): matches KeyType (name without numbers) + Value, case-insensitive

## [Add Mode](#add-mode) 

Adds content from the source file to the base file.

### Sub-modes 
None

### Behavior

- Sections from the source file not found in the base file are added to the base file as written
- Sections from the source file which are found in the base file have the keys from the source file added to the base file as written 
- Does not avoid creating duplicate keys
- No existing keys are modified or removed
- Keys are *not* renumbered; normalize later with WinappDebug if needed (winapp2.ini only)

### Example use cases
- Maintaining a set of customizations to winapp2.ini while also keeping it up to date 
- Merging multiple configuration files into one

#

## [Replace Mode](#replace-mode)

Overwrites existing content in the base file with content from the source file.

### Sub-modes

|Sub Mode|Effect|Notes
|:-|:-|:-|
| BySection | Replaces entire sections by their name       | Section name matches are case sensitive |
| By Key    | Replaces individual key values by their name | Default                                 |

### Behavior 

#### By Section
- Sections from the source file which are found in the base file replace entirely the section in the base file as they are written  
- Sections from the source file not found in the base file are ignored 

#### By Key
- Keys from sections in the source file which are found in the base file replace entirely keys of the same name in the base file as they are written 
- Sections and keys from the source file not found in the base file are ignored 

### Example use cases
 
- Correcting errors in generated entries
- Supporting local configurations in winapp2.ini while also keeping it up to date 

#

## [Remove Mode](#remove-mode)
Removes content from the base file based on matches in the source file.

### Sub-modes

|Sub Mode|Effect|Notes
|:-|:-|:-
|By Section | Removes entire sections from the base file by their name | Section name matches are case sensitive
|By Key     | Removes individual keys from sections in the base file   | Default 

### Key Removal Sub Modes
|Sub Mode|Effect|Notes
|:-|:-|:-
|By Name  | Removes keys from the base file by their Name                 | Default
|By Value | Removes keys from the base file based on their Name and Value | Value matches ignore numbers in the key name

### Behavior

#### By Section
- Sections from the source file which are found in the base file are removed 
- Sections from the source file not found in the base file are ignored 
- Key values are ignored in this mode 

#### By Key - By Name
- Keys from sections in the source file which are found in the base file are removed if they have the same name 
- Keys from the source file not found in the base file are ignored

#### By Key - By Value
- Keys from sections in the source file which are found in the base file are removed if they have the same name (ignoring numbers) and value 
- Keys from the source file not found in the base file are ignored 

### Example use cases
- Automatically Removing unwanted content from winapp2.ini while also keeping it up to date 
- Automatically removing generated entries which are non-viable for a purpose 

---

# [Flavorization](flavorization)

Flavorization applies a comprehensive set of transformations to create specialized variants of ini files. Operations are applied in this order:

1.	Section Removal
2.	Key Name Removal
3.	Key Value Removal
4.	Section Replacement
5.	Key Replacement
6.	Section and Key Additions

### Example use cases

-	Creating winapp2.ini variants (eg. a CCleaner or BleachBit flavor)
-	Generating platform-specific configurations (eg. a Windows XP flavor)
-	Automated quality control corrections (eg. Correcting entries generated by winapp2ool)

---

# [Command-Line Arguments](#command-line-arguments)
Transmute supports command-line automation for scripting environments.

Arguments labeled "Default" are assumed by default and can be optionally omitted when invoking, though this is not recommended. 

Arguments affecting default settings are ignored. Eg. If you specify `-bykey`, the default sub mode, it is ignored. If you specify both `-bykey` and also `-bysection`, the `-bykey` will be ignored and the resulting sub mode will be `By Section`

### [Primary Mode](#primary-mode)

|Arg|Effect|Notes
|:-|:-|:-
| -add     | Add mode     | Default
| -replace | Replace mode |
| -remove  | Remove mode  |

### [Sub-Modes (Replace/Remove)](#sub-modes-replace-remove)

|Arg|Effect|Notes
|:-|:-|:-
| -bysection | Operate on entire sections |
| -bykey     | Operate on individual keys | Default

### [Key Removal Modes](#key-removal-modes)
|Arg|Effect|Notes
|:-|:-|:-
| -byname  | Remove keys by exact name match     | Default
| -byvalue | Remove keys by type and value match |

### [Preset Source Files](#preset-source-files)
Sets the source file name to one of the pre-defined defaults, most of which are available from GitHub

### [Toggles](toggles)
|Arg|Effect|
|:-|:-
| -dontlint | Save the output without winapp2.ini formatting (section ordering and pre-amble comments)

|Arg|Effect|Description
|:-|:-|:-
| -r | Use "Removed Entries.ini"  | Contains entries removed from winapp2.ini to create the non-CCleaner winapp2.ini*
| -c | Use "Custom.ini"           | Suggested name for user additions files
| -w | Use "winapp3.ini"          | Contains winapp2.ini entries which may potentially break applications, use at your own  risk!
| -a | Use "Archived Entries.ini" | Contains winapp2.ini entries removed because the applications they target are no longer available 
| -b | Use "Browsers.ini"         | The default output file from Browser Builder

\*The non-ccleaner winapp2.ini will be deprecated soon and treated as the default, support for this file will be removed in a later update 

###### Note: Source files *must* be local. Transmute does not directly download any files.

### [File Selection](#file-selection)
| Arg | Effect | Default Value
|:-|:-|:-
| -1d path        | Set base file path                                      | Current Directory
| -1f name        | Set base file name                                      | winapp2.ini
| -1f subdir\name | Set base file name within a subfolder of its path       |
| -2d path        | Set source file path                                    | Current Directory            
| -2f name        | Set source file name                                    | None 
| -2f subdir\name | Set the source file name within a subfolder of its path | 
| -3d path        | Set output file path                                    | Current Directory  
| -3f name        | Set output file name                                    | winapp2.ini
| -3f subdir\name | Set the output file name within a subfolder of its path |

#

### [Examples](#examples)

|Command|Effect|
 |:-|:-| 
|winapp2ool -transmute -add -2f custom.ini|Add entries and keys from custom.ini to winapp2.ini
|winapp2ool -transmute -replace -bysection -c|Replace entire sections in winapp2.ini with ones from Custom.ini
|winapp2ool -transmute -remove -bykey -byvalue -2f key_value_removals.ini |Remove keys from winapp2.ini sections defined by value in key_value_removals.ini
|winapp2ool -transmute -remove -bysection -2f section_removals.ini -3f cleaned.ini | Remove sections from winapp2.ini defined in section_removals.ini and save the result  to cleaned.ini

---

# [Source Files](#source-files)

Your source file must follow standard ini format:

```ini
[Section Name]
Key1=Value1
Key2=Value2

[Another Section]
Key1=Value1
``` 

Comments included in your source file will not be placed in the base file.

## [Tips & Best Practices](#tips--best-practices)

### [Safety First](#safety-first)

- Always test on copies of important files
- Overwriting the base file is non-reversible without a backup, Winapp2ool does not create one 
- Review the detailed output log to ensure your changes applied correctly and in the way you intended 

### [Effective Source Files](#effective-source-files)

- Keep source files focused on specific changes
- Comment your source files for future reference. Comments will ***not*** be placed in the output file. 
- It is important to note that Transmute operations are case sensitive when assessing the names of sections

### [Mode Selection](#mode-selection)

- Use Add for supplements and extensions
- Use Replace for updates and corrections
- Use Remove for cleanup and simplification
- It may be easier to remove a key and add a new one with a different value than it is to replace a key by value under certain circumstances

### [Key Removal Strategy](#key-removal-strategy)

- Use ByName for unnumbered keys or when you know exact names
- Use ByValue for numbered keys where numbering might vary 

---

# [Troubleshooting](#troubleshooting)
|Error Message|Cause
|:-|:-
|"Target section not found in the base file: [section]"|Source file references a section not in base file (only affects Replace/Remove)
|"Replacement target not found in [section]"|Source key doesn't exist in the base section (Replace mode)
|"Removal target not found: {key} not found in [section]"|Source key doesn't exist in the base section (Remove mode)
---
# [Migrating from Merge](#migrating-from-merge)

Merge fundamentally differed from Transmute in that Merge *always* applied additions, but with a much more limited capacity for conflict resolution between its "Add & Remove" and "Add & Replace" modes. Transmute makes much more granular changes to the files but can achieve the same output. 

To migrate to Transmute, you'll need to reconfigure your set of changes into categories by their effects under the new Transmute modes 

### [New content](#new-content)
 Content you are adding (eg. custom entries or keys you wrote for your system) can all be placed together in one file. This functions mostly similarly to the way additions worked in Merge, with the new feature of being able to add individual keys to existing entries rather than requiring you to provide an entire section replacement. Apply additions by setting the Transmute mode to Add.

### [Replacement content](#replacement-content)
 Place any sections you want to have replace entries in winapp2.ini in a separate file from keys you want to replace within individual sections. Apply replacements by first setting the Transmute mode to Replace. The default Replace mode is By Key. Set the Replace mode to By Section to replace entire sections. 

### [Removals](#removals)
 Place any sections you want to remove entirely from winapp2.ini into a separate file from keys you want to remove from within individual sections. Likewise, place any keys you want to remove by their value in a separate file from keys you want to remove by their name. When removing entire sections, you need not provide any keys. Apply removals by first setting the Transmute mode to Remove. The default Remove mode is By Key. Set the Remove mode to By Section to remove entire sections. The default Key Removal mode is By Name. Set the Key Removal mode to By Value to remove keys by their value. 

##### This guide is provided as a general framework for decision making. For technical guidance on the commands required to migrate to Transmute from Merge, see Usage Examples below
 
---

 # [Usage Examples](#usage-examples)

To drive some of our examples, we'll take a look at some of the work done by winapp2ool to apply corrections to the output of the Browser Builder. These files can be found [here](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/BrowserBuilder), but relevant lines of code will be provided on this page.  

## [Adding Content](#adding-content)

### [Example 1: Adding New Sections and Keys](#example-1-adding-new-sections-and-keys)

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
; This key will be added to the generated entry of the same name (case sensitive) in the base file to provide better coverage
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
- `[360 Secure Browser Web Browsing History Backups]` is added to the base file as defined in the source file 
- `[360 Secure Browser Bookmarked Websites]` in the base file has `FileKey` added with value `%AppData%\360se6\User Data\*|360Bookmarks*` from the source file
- Sections in the base file not defined in the source file remain unchanged  

**Notes**

Transmute does not do any work to ensure correct key numbering on its own, it adds keys as written (eg. `FileKey` above ). Winapp2.ini files should be run through WinappDebug to correct their syntax.  

---

## [Replacements](#replacements)

### [Example 2: Replacing Key Values](#example-2-replacing-key-values)

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

 This mode replaces keys by their `Name` which is the entire value to the left of the `=`. It may produce more consistent results for numbered keys to replace a key's value by first removing it and then adding a new key with the desired value. 

#

### [Example 3: Replacing Entire Sections](#example-3-replacing-entire-sections)

**Context**

Suppose you have Some Application installed which is supported by winapp2.ini but for which you require substantial changes. You have made those changes, but you want to keep the rest of winapp2.ini up to date.

**Intent**

We want to replace the winapp2.ini version of `[Some Application]` with our custom copy we maintain separately

**Files**

###### **Base file (`winapp2.ini`)**
```ini
[Some Application]
LangSecRef=3021
DetectFile=%AppData%\Some Application
FileKey1=%AppData%\Some Application\*Cache*|*|REMOVESELF
FileKey2=%LocalAppData%\Some Application\Logs|*.log
```
###### **Source file (`section_replacements.ini`)**
```ini
[Some Application]
Section=My Application
DetectFile=%SystemDrive%\Some Application
FileKey1=%SystemDrive%\Some Application\*Cache*|*|REMOVESELF
FileKey2=%SystemDrive%\Some Application\Logs|*.log
```

**Command**
```
winapp2ool -transmute -replace -bysection -2f section_replacements.ini
```

**Output**

###### **Output file (`winapp2.ini`) after transmutation**

```ini
[Some Application]
Section=My Application
DetectFile=%SystemDrive%\Some Application
FileKey1=%SystemDrive%\Some Application\*Cache*|*|REMOVESELF
FileKey2=%SystemDrive%\Some Application\Logs|*.log
```

**Explanation**
- The base file is winapp2.ini (default)
- The source file is section_replacements.ini
- The output file is winapp2.ini (default, overwriting the base file)
- `[Some Application]` in the base file is entirely replaced by `[Some Application]` from the source file 
- Sections in the base file not defined in the source file remain unchanged  

---

## [Removals](#removals-1)

### [Example 4: Removing Entire Sections](#example-4-removing-entire-sections)

**Context**

Some entries generated by Browser Builder are incomplete or target features not implemented in a particular browser. Rather than ship them targeting nothing, we'd like to prune them from the set of generated entries before combining them into winapp2.ini. Arc implements its pinned tabs storage as a part of a JSON file shared with persistent configuration which winapp2.ini doesn't support cleaning non-destructively.   

**Intent**

We want to remove `[Arc Pinned Tabs *]` from browsers.ini because there is nothing we can do to make it viable. 

**Files**

###### **Base file (browsers.ini)**
```ini
; This entry is generated without a RegKey because the RegistryRoot is not defined in chromium.ini 
[Arc Pinned Tabs *]
Section=NAVER Whale Web Browser
DetectFile=%LocalAppData%\Naver\Naver Whale\User Data
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

#

### [Example 5: Remove Keys By Name(#example-5-remove-keys-by-name)

**Context**

You are developing a tool which implements winapp2.ini. You are aware that `LangSecRef` is a categorizer, but your tool doesn't use these categories. As such, you want to remove the LangSecRef key from a particular entry before passing it into your tool. 

**Intent**

We want to remove the `LangSecRef` key from `[Some Application]`

**Files**

###### **Base file (`winapp2.ini`)**
```ini
[Some Application]
LangSecRef=3021
DetectFile=%AppData%\Some Application
FileKey1=%AppData%\Some Application\*cache*|*|REMOVESELF
```

##### **Source file (`key_name_removals.ini`)**
```ini
[Some Application]
LangSecRef=3021
```

**Command**
```
winapp2ool -transmute -remove -bykey -byname -2f key_name_removals.ini
```

**Output**

###### **Output file (`winapp2.ini`) after transmutation**

```ini
[Some Application]
DetectFile=%AppData%\Some Application
FileKey1=%AppData%\Some Application\*cache*|*|REMOVESELF
```

**Explanation**
- The base file is winapp2.ini
- The source file is key_name_removals.ini
- The output file is winapp2.ini (overwriting the base file)
- `LangSecRef=3021` is removed from the base file 
- Sections in the base file not defined in the source file remain unchanged 

#

### [Example 6: Removing Keys By Value](#example-6-removing-keys-by-value)

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
- Any `DetectFile` key with a value provided from `[DuckDuckGo Autofill Data]` in the source file is removed from `[DuckDuckGo Autofill Data]` in the base file
- Sections in the base file not defined in the source file remain unchanged 

---

## [Advanced](#advanced)

### [Example 7: Chaining operations together 1](#example-7-chaining-operations-together-1)

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
- Two separate transmutations are conducted
- For the first transmutation, the base file is browsers.ini
- For the first transmutation, the source file is browser_value_removals.ini
- For the first transmutation, the output file is browsers.ini (overwriting the base file)
- Any `DetectFile` key with a value provided from `[DuckDuckGo Autofill Data]` in the source file is removed from `[DuckDuckGo Autofill Data]` in the base file
- Sections in the base file not defined in the source file remain unchanged 
- The output is saved to browsers.ini, overwriting the initial base file before the second transmutation  
- For the second transmutation, the base file is browsers.ini (having been modified once already)
- For the second transmutation, the source file is browser_additions.ini 
- For the second transmutation, the output file is browsers.ini (overwriting the base file)
- The key `DetectFile=%LocalAppData%\Packages\*DuckDuckGo*Browser*` from `[DuckDuckGo Autofill Data]` in the source file is added to `[DuckDuckGo Autofill Data]` in the base file 
- Sections in the base file not defined in the source file remain unchanged 

### [Example 8: Chaining operations together 2](#example-8-chaining-operations-together-2)

**Context**

As part of the switch to non-ccleaner as default, winapp2.ini recently declared separate sections for each of the web browsers. This conflicts with the old configuration which placed them, at least for CCleaner, together with the CCleaner entries for the same browser. We want to undo this as part of creating a CCleaner flavor of winapp2.ini 

**Intent**

We want to remove the `Section` keys from a particular web browser and add a `LangSecRef` pointing to the appropriate CCleaner section. 

**Files**

###### **Base file (`browsers.ini`)**
```ini
; We're going to use just one browser as an example 
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

[Vivaldi Bookmark Favicons *]
Section=Vivaldi Web Browser
DetectFile=%LocalAppData%\Vivaldi\User Data
FileKey1=%LocalAppData%\Vivaldi\User Data\*|favicons*

[Vivaldi Bookmarked Websites *]
Section=Vivaldi Web Browser
DetectFile=%LocalAppData%\Vivaldi\User Data
FileKey1=%LocalAppData%\Vivaldi\User Data\*|bookmarks;BookmarkMergedSurfaceOrdering
FileKey2=%LocalAppData%\Vivaldi\User Data\*\power_bookmarks|*|REMOVESELF

[Vivaldi Caches *]
Section=Vivaldi Web Browser
DetectFile=%LocalAppData%\Vivaldi\User Data
FileKey1=%LocalAppData%\Vivaldi\User Data|*-journal|RECURSE
FileKey2=%LocalAppData%\Vivaldi\User Data|Module Info Cache
FileKey3=%LocalAppData%\Vivaldi\User Data\*|*.ldb;CURRENT;LOCK;MANIFEST-*;ServerCertificate
FileKey4=%LocalAppData%\Vivaldi\User Data\*\*Cache*|*|REMOVESELF
FileKey5=%LocalAppData%\Vivaldi\User Data\*\blob_storage|*|REMOVESELF
FileKey6=%LocalAppData%\Vivaldi\User Data\*\BudgetDatabase|*|REMOVESELF
FileKey7=%LocalAppData%\Vivaldi\User Data\*\Download Service|*|REMOVESELF
FileKey8=%LocalAppData%\Vivaldi\User Data\*\File System|*|REMOVESELF
FileKey9=%LocalAppData%\Vivaldi\User Data\*\GCM Store|*|REMOVESELF
FileKey10=%LocalAppData%\Vivaldi\User Data\*\JumpListIcons*|*|REMOVESELF
FileKey11=%LocalAppData%\Vivaldi\User Data\*\optimization*|*|REMOVESELF
FileKey12=%LocalAppData%\Vivaldi\User Data\*\Platform Notifications|*|REMOVESELF
FileKey13=%LocalAppData%\Vivaldi\User Data\*\Service Worker|*|REMOVESELF
FileKey14=%LocalAppData%\Vivaldi\User Data\*\Shared Dictionary\cache|*|REMOVESELF
FileKey15=%LocalAppData%\Vivaldi\User Data\*\Storage\ext\*\def\*cache*|*|REMOVESELF
FileKey16=%LocalAppData%\Vivaldi\User Data\*\Storage\ext\*\def\Platform Notifications|*|REMOVESELF
FileKey17=%LocalAppData%\Vivaldi\User Data\*\Sync Data|*|REMOVESELF
FileKey18=%LocalAppData%\Vivaldi\User Data\*Cache*|*|REMOVESELF
FileKey19=%LocalAppData%\Vivaldi\User Data\Avatars|*|REMOVESELF
FileKey20=%LocalAppData%\Vivaldi\User Data\Optimization*|*|REMOVESELF

[Vivaldi Default App Handlers *]
Section=Vivaldi Web Browser
DetectFile=%LocalAppData%\Vivaldi\User Data
FileKey1=%LocalAppData%\Vivaldi\User Data\*|PreferredApps

[Vivaldi Download History *]
Section=Vivaldi Web Browser
DetectFile=%LocalAppData%\Vivaldi\User Data
FileKey1=%LocalAppData%\Vivaldi\User Data\*|DownloadMetadata
FileKey2=%LocalAppData%\Vivaldi\User Data\*\shared_proto_db|*|REMOVESELF

[Vivaldi DRM Data *]
Section=Vivaldi Web Browser
DetectFile=%LocalAppData%\Vivaldi\User Data
FileKey1=%LocalAppData%\Vivaldi\User Data\*\MediaFoundationCdmStore|*|REMOVESELF
FileKey2=%LocalAppData%\Vivaldi\User Data\WidevineCdm|*|REMOVESELF

[Vivaldi Extension Cookies *]
Section=Vivaldi Web Browser
DetectFile=%LocalAppData%\Vivaldi\User Data
FileKey1=%LocalAppData%\Vivaldi\User Data\*|Extension Cookies

[Vivaldi Pinned Tabs *]
Section=Vivaldi Web Browser
DetectFile=%LocalAppData%\Vivaldi\User Data
RegKey1=HKCU\Software\Vivaldi\PreferenceMACs\Default|pinned_tabs

[Vivaldi Privacy Sandbox *]
Section=Vivaldi Web Browser
DetectFile=%LocalAppData%\Vivaldi\User Data
FileKey1=%LocalAppData%\Vivaldi\User Data|*first_party_sets*
FileKey2=%LocalAppData%\Vivaldi\User Data\*|BrowsingTopics*;Conversions*;InterestGroups;MediaDeviceSalts;SharedStorage*;PrivateAggregation*
FileKey3=%LocalAppData%\Vivaldi\User Data\*\Network|Trust Tokens*
FileKey4=%LocalAppData%\Vivaldi\User Data\CookieReadinessList|*|REMOVESELF
FileKey5=%LocalAppData%\Vivaldi\User Data\FirstPartySetsPreloaded|*|REMOVESELF
FileKey6=%LocalAppData%\Vivaldi\User Data\PrivacySandboxAttestationsPreloaded|*|REMOVESELF
FileKey7=%LocalAppData%\Vivaldi\User Data\ProbabilisticRevealTokenRegistry|*|REMOVESELF
FileKey8=%LocalAppData%\Vivaldi\User Data\TrustTokenKeyCommitments|*|REMOVESELF

[Vivaldi Progressive Web Apps *]
Section=Vivaldi Web Browser
DetectFile=%LocalAppData%\Vivaldi\User Data
FileKey1=%LocalAppData%\Vivaldi\User Data\*\Web Applications|*|REMOVESELF

[Vivaldi Saved Usernames & Passwords *]
Section=Vivaldi Web Browser
DetectFile=%LocalAppData%\Vivaldi\User Data
FileKey1=%LocalAppData%\Vivaldi\User Data\*|Login Data*

[Vivaldi Security & Threat Detection *]
Section=Vivaldi Web Browser
DetectFile=%LocalAppData%\Vivaldi\User Data
FileKey1=%LocalAppData%\Vivaldi\User Data\*|DIPS*
FileKey2=%LocalAppData%\Vivaldi\User Data\*\ClientCertificates|*|REMOVESELF
FileKey3=%LocalAppData%\Vivaldi\User Data\*\Network|TransportSecurity*
FileKey4=%LocalAppData%\Vivaldi\User Data\*\Safe Browsing Network|*|REMOVESELF
FileKey5=%LocalAppData%\Vivaldi\User Data\CertificateRevocation|*|REMOVESELF
FileKey6=%LocalAppData%\Vivaldi\User Data\Crowd Deny|*|REMOVESELF
FileKey7=%LocalAppData%\Vivaldi\User Data\PKIMetadata|*|REMOVESELF
FileKey8=%LocalAppData%\Vivaldi\User Data\Safe Browsing|*|REMOVESELF
FileKey9=%LocalAppData%\Vivaldi\User Data\SafetyTips|*|REMOVESELF
FileKey10=%LocalAppData%\Vivaldi\User Data\SSLErrorAssistant|*|REMOVESELF
FileKey11=%LocalAppData%\Vivaldi\User Data\Subresource Filter|*|REMOVESELF
FileKey12=%LocalAppData%\Vivaldi\User Data\ZxcvbnData|*|REMOVESELF

[Vivaldi Shopping Insights *]
Section=Vivaldi Web Browser
DetectFile=%LocalAppData%\Vivaldi\User Data
FileKey1=%LocalAppData%\Vivaldi\User Data\*\chrome_cart_db|*|REMOVESELF
FileKey2=%LocalAppData%\Vivaldi\User Data\*\commerce_subscription_db|*|REMOVESELF
FileKey3=%LocalAppData%\Vivaldi\User Data\*\coupon_db|*|REMOVESELF
FileKey4=%LocalAppData%\Vivaldi\User Data\*\discount*_db|*|REMOVESELF
FileKey5=%LocalAppData%\Vivaldi\User Data\*\parcel_tracking_db|*|REMOVESELF

[Vivaldi Telemetry *]
Section=Vivaldi Web Browser
DetectFile=%LocalAppData%\Vivaldi\User Data
FileKey1=%LocalAppData%\Vivaldi\Application|debug.log
FileKey2=%LocalAppData%\Vivaldi\Application\SetupMetrics|*|REMOVESELF
FileKey3=%LocalAppData%\Vivaldi\User Data|*.pma;LOG;LOG.old|RECURSE
FileKey4=%LocalAppData%\Vivaldi\User Data|*_shutdown_ms.txt;*.log;Last Browser;Last Version;Breadcrumbs;BrowsingTopics*
FileKey5=%LocalAppData%\Vivaldi\User Data\*|*.log
FileKey6=%LocalAppData%\Vivaldi\User Data\*\DataSharing|*|REMOVESELF
FileKey7=%LocalAppData%\Vivaldi\User Data\*\Feature Engagement Tracker|*|REMOVESELF
FileKey8=%LocalAppData%\Vivaldi\User Data\*\Network|Network Persistent State*;Reporting and NEL*;SCT Auditing Pending Reports*
FileKey9=%LocalAppData%\Vivaldi\User Data\*\PersistentOriginTrials|*|REMOVESELF
FileKey10=%LocalAppData%\Vivaldi\User Data\*\Site Characteristics Database|*|REMOVESELF
FileKey11=%LocalAppData%\Vivaldi\User Data\*\VideoDecodeStats|*|REMOVESELF
FileKey12=%LocalAppData%\Vivaldi\User Data\*\WebRTC Logs|*|REMOVESELF
FileKey13=%LocalAppData%\Vivaldi\User Data\*\WebrtcVideoStats|*|REMOVESELF
FileKey14=%LocalAppData%\Vivaldi\User Data\*BrowserMetrics|*|REMOVESELF
FileKey15=%LocalAppData%\Vivaldi\User Data\Crash Reports|*|REMOVESELF
FileKey16=%LocalAppData%\Vivaldi\User Data\Crashpad|*|REMOVESELF
FileKey17=%LocalAppData%\Vivaldi\User Data\Local Traces|*|REMOVESELF
FileKey18=%LocalAppData%\Vivaldi\User Data\OriginTrials|*|REMOVESELF
FileKey19=%LocalAppData%\Vivaldi\User Data\Stability|*|REMOVESELF
FileKey20=%ProgramFiles%\Vivaldi\Application|debug.log
FileKey21=%ProgramFiles%\Vivaldi\Application\SetupMetrics|*|REMOVESELF
RegKey1=HKCU\Software\Vivaldi\BlBeacon|failed_count
RegKey2=HKCU\Software\Vivaldi\BlBeacon|state
RegKey3=HKCU\Software\Vivaldi\Installer|LastChecked
RegKey4=HKCU\Software\Vivaldi\StabilityMetrics

[Vivaldi Updates *]
Section=Vivaldi Web Browser
DetectFile=%LocalAppData%\Vivaldi\User Data
FileKey1=%LocalAppData%\Vivaldi\Application\*\Installer|*|REMOVESELF
FileKey2=%LocalAppData%\Vivaldi\Application\*\Temp|*|REMOVESELF
FileKey3=%LocalAppData%\Vivaldi\Temp|*|REMOVESELF
FileKey4=%ProgramFiles%\Vivaldi\Application\*\Installer|*|REMOVESELF
FileKey5=%ProgramFiles%\Vivaldi\Application\*\Temp|*|REMOVESELF
FileKey6=%ProgramFiles%\Vivaldi\Temp|*|REMOVESELF

[Vivaldi Web Browsing Cookies *]
Section=Vivaldi Web Browser
DetectFile=%LocalAppData%\Vivaldi\User Data
FileKey1=%LocalAppData%\Vivaldi\User Data\*\Network|Cookies*;Device Bound Sessions*

[Vivaldi Web Browsing History *]
Section=Vivaldi Web Browser
DetectFile=%LocalAppData%\Vivaldi\User Data
FileKey1=%LocalAppData%\Vivaldi\User Data\*|History*;Visited Links*;Top Sites*;Network Action Predictor*;shortcuts*

[Vivaldi Web Browsing Session *]
Section=Vivaldi Web Browser
DetectFile=%LocalAppData%\Vivaldi\User Data
FileKey1=%LocalAppData%\Vivaldi\User Data\*\Extension State|*|REMOVESELF
FileKey2=%LocalAppData%\Vivaldi\User Data\*\Sessions|*|REMOVESELF

[Vivaldi Web Storage *]
Section=Vivaldi Web Browser
DetectFile=%LocalAppData%\Vivaldi\User Data
FileKey1=%LocalAppData%\Vivaldi\User Data\*\IndexedDB|*|REMOVESELF
FileKey2=%LocalAppData%\Vivaldi\User Data\*\Local Storage\*\IndexedDB\https*|*
FileKey3=%LocalAppData%\Vivaldi\User Data\*\Local Storage\*\Session Storage|*
FileKey4=%LocalAppData%\Vivaldi\User Data\*\Local Storage\*\Storage\ext\*\def\Session Storage|*
FileKey5=%LocalAppData%\Vivaldi\User Data\*\Local Storage\*\WebStorage|*|RECURSE
FileKey6=%LocalAppData%\Vivaldi\User Data\*\Local Storage\leveldb|*|REMOVESELF
FileKey7=%LocalAppData%\Vivaldi\User Data\*\Session Storage|*|REMOVESELF
FileKey8=%LocalAppData%\Vivaldi\User Data\*\WebStorage|*|REMOVESELF

```

###### **Source file (`browser_value_removals.ini`)**
```ini
[Vivaldi Autofill Data *]
Section=Vivaldi Web Browser

[Vivaldi Autoplay Preferences *]
Section=Vivaldi Web Browser

[Vivaldi Bookmark Backups *]
Section=Vivaldi Web Browser

[Vivaldi Bookmark Favicons *]
Section=Vivaldi Web Browser

[Vivaldi Bookmarked Websites *]
Section=Vivaldi Web Browser

[Vivaldi Caches *]
Section=Vivaldi Web Browser

[Vivaldi Default App Handlers *]
Section=Vivaldi Web Browser

[Vivaldi Download History *]
Section=Vivaldi Web Browser

[Vivaldi DRM Data *]
Section=Vivaldi Web Browser

[Vivaldi Extension Cookies *]
Section=Vivaldi Web Browser

[Vivaldi Pinned Tabs *]
Section=Vivaldi Web Browser

[Vivaldi Privacy Sandbox *]
Section=Vivaldi Web Browser

[Vivaldi Progressive Web Apps *]
Section=Vivaldi Web Browser

[Vivaldi Saved Usernames & Passwords *]
Section=Vivaldi Web Browser

[Vivaldi Security & Threat Detection *]
Section=Vivaldi Web Browser

[Vivaldi Shopping Insights *]
Section=Vivaldi Web Browser

[Vivaldi Telemetry *]
Section=Vivaldi Web Browser

[Vivaldi Updates *]
Section=Vivaldi Web Browser

[Vivaldi Web Browsing Cookies *]
Section=Vivaldi Web Browser

[Vivaldi Web Browsing History *]
Section=Vivaldi Web Browser

[Vivaldi Web Browsing Session *]
Section=Vivaldi Web Browser

[Vivaldi Web Storage *]
Section=Vivaldi Web Browser
```

###### **Source file(`browser_additions.ini`)**
```ini
[Vivaldi Autofill Data *]
LangSecRef=3033

[Vivaldi Autoplay Preferences *]
LangSecRef=3033

[Vivaldi Bookmark Backups *]
LangSecRef=3033

[Vivaldi Bookmark Favicons *]
LangSecRef=3033

[Vivaldi Bookmarked Websites *]
LangSecRef=3033

[Vivaldi Caches *]
LangSecRef=3033

[Vivaldi Default App Handlers *]
LangSecRef=3033

[Vivaldi Download History *]
LangSecRef=3033

[Vivaldi DRM Data *]
LangSecRef=3033

[Vivaldi Extension Cookies *]
LangSecRef=3033

[Vivaldi Pinned Tabs *]
LangSecRef=3033

[Vivaldi Privacy Sandbox *]
LangSecRef=3033

[Vivaldi Progressive Web Apps *]
LangSecRef=3033

[Vivaldi Saved Usernames & Passwords *]
LangSecRef=3033

[Vivaldi Security & Threat Detection *]
LangSecRef=3033

[Vivaldi Shopping Insights *]
LangSecRef=3033

[Vivaldi Telemetry *]
LangSecRef=3033

[Vivaldi Updates *]
LangSecRef=3033

[Vivaldi Web Browsing Cookies *]
LangSecRef=3033

[Vivaldi Web Browsing History *]
LangSecRef=3033

[Vivaldi Web Browsing Session *]
LangSecRef=3033

[Vivaldi Web Storage *]
LangSecRef=3033
```

**Commands**
```
winapp2ool -transmute -remove -bykey -byvalue -1f browsers.ini -2f browser_value_removals.ini -3f browsers.ini
winapp2ool -transmute -add -1f browsers.ini -2f browser_additions.ini -3f browsers.ini 
```

###### Note: `By Key` is the default remove mode and technically the `-bykey` argument could be omitted here but is provided for the utmost clarity 

###### Note

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

[Vivaldi Bookmark Favicons *]
LangSecRef=3033
DetectFile=%LocalAppData%\Vivaldi\User Data
FileKey1=%LocalAppData%\Vivaldi\User Data\*|favicons*

[Vivaldi Bookmarked Websites *]
LangSecRef=3033
DetectFile=%LocalAppData%\Vivaldi\User Data
FileKey1=%LocalAppData%\Vivaldi\User Data\*|bookmarks;BookmarkMergedSurfaceOrdering
FileKey2=%LocalAppData%\Vivaldi\User Data\*\power_bookmarks|*|REMOVESELF

[Vivaldi Caches *]
LangSecRef=3033
DetectFile=%LocalAppData%\Vivaldi\User Data
FileKey1=%LocalAppData%\Vivaldi\User Data|*-journal|RECURSE
FileKey2=%LocalAppData%\Vivaldi\User Data|Module Info Cache
FileKey3=%LocalAppData%\Vivaldi\User Data\*|*.ldb;CURRENT;LOCK;MANIFEST-*;ServerCertificate
FileKey4=%LocalAppData%\Vivaldi\User Data\*\*Cache*|*|REMOVESELF
FileKey5=%LocalAppData%\Vivaldi\User Data\*\blob_storage|*|REMOVESELF
FileKey6=%LocalAppData%\Vivaldi\User Data\*\BudgetDatabase|*|REMOVESELF
FileKey7=%LocalAppData%\Vivaldi\User Data\*\Download Service|*|REMOVESELF
FileKey8=%LocalAppData%\Vivaldi\User Data\*\File System|*|REMOVESELF
FileKey9=%LocalAppData%\Vivaldi\User Data\*\GCM Store|*|REMOVESELF
FileKey10=%LocalAppData%\Vivaldi\User Data\*\JumpListIcons*|*|REMOVESELF
FileKey11=%LocalAppData%\Vivaldi\User Data\*\optimization*|*|REMOVESELF
FileKey12=%LocalAppData%\Vivaldi\User Data\*\Platform Notifications|*|REMOVESELF
FileKey13=%LocalAppData%\Vivaldi\User Data\*\Service Worker|*|REMOVESELF
FileKey14=%LocalAppData%\Vivaldi\User Data\*\Shared Dictionary\cache|*|REMOVESELF
FileKey15=%LocalAppData%\Vivaldi\User Data\*\Storage\ext\*\def\*cache*|*|REMOVESELF
FileKey16=%LocalAppData%\Vivaldi\User Data\*\Storage\ext\*\def\Platform Notifications|*|REMOVESELF
FileKey17=%LocalAppData%\Vivaldi\User Data\*\Sync Data|*|REMOVESELF
FileKey18=%LocalAppData%\Vivaldi\User Data\*Cache*|*|REMOVESELF
FileKey19=%LocalAppData%\Vivaldi\User Data\Avatars|*|REMOVESELF
FileKey20=%LocalAppData%\Vivaldi\User Data\Optimization*|*|REMOVESELF

[Vivaldi Default App Handlers *]
LangSecRef=3033
DetectFile=%LocalAppData%\Vivaldi\User Data
FileKey1=%LocalAppData%\Vivaldi\User Data\*|PreferredApps

[Vivaldi Download History *]
LangSecRef=3033
DetectFile=%LocalAppData%\Vivaldi\User Data
FileKey1=%LocalAppData%\Vivaldi\User Data\*|DownloadMetadata
FileKey2=%LocalAppData%\Vivaldi\User Data\*\shared_proto_db|*|REMOVESELF

[Vivaldi DRM Data *]
LangSecRef=3033
DetectFile=%LocalAppData%\Vivaldi\User Data
FileKey1=%LocalAppData%\Vivaldi\User Data\*\MediaFoundationCdmStore|*|REMOVESELF
FileKey2=%LocalAppData%\Vivaldi\User Data\WidevineCdm|*|REMOVESELF

[Vivaldi Extension Cookies *]
LangSecRef=3033
DetectFile=%LocalAppData%\Vivaldi\User Data
FileKey1=%LocalAppData%\Vivaldi\User Data\*|Extension Cookies

[Vivaldi Pinned Tabs *]
LangSecRef=3033
DetectFile=%LocalAppData%\Vivaldi\User Data
RegKey1=HKCU\Software\Vivaldi\PreferenceMACs\Default|pinned_tabs

[Vivaldi Privacy Sandbox *]
LangSecRef=3033
DetectFile=%LocalAppData%\Vivaldi\User Data
FileKey1=%LocalAppData%\Vivaldi\User Data|*first_party_sets*
FileKey2=%LocalAppData%\Vivaldi\User Data\*|BrowsingTopics*;Conversions*;InterestGroups;MediaDeviceSalts;SharedStorage*;PrivateAggregation*
FileKey3=%LocalAppData%\Vivaldi\User Data\*\Network|Trust Tokens*
FileKey4=%LocalAppData%\Vivaldi\User Data\CookieReadinessList|*|REMOVESELF
FileKey5=%LocalAppData%\Vivaldi\User Data\FirstPartySetsPreloaded|*|REMOVESELF
FileKey6=%LocalAppData%\Vivaldi\User Data\PrivacySandboxAttestationsPreloaded|*|REMOVESELF
FileKey7=%LocalAppData%\Vivaldi\User Data\ProbabilisticRevealTokenRegistry|*|REMOVESELF
FileKey8=%LocalAppData%\Vivaldi\User Data\TrustTokenKeyCommitments|*|REMOVESELF

[Vivaldi Progressive Web Apps *]
LangSecRef=3033
DetectFile=%LocalAppData%\Vivaldi\User Data
FileKey1=%LocalAppData%\Vivaldi\User Data\*\Web Applications|*|REMOVESELF

[Vivaldi Saved Usernames & Passwords *]
LangSecRef=3033
DetectFile=%LocalAppData%\Vivaldi\User Data
FileKey1=%LocalAppData%\Vivaldi\User Data\*|Login Data*

[Vivaldi Security & Threat Detection *]
LangSecRef=3033
DetectFile=%LocalAppData%\Vivaldi\User Data
FileKey1=%LocalAppData%\Vivaldi\User Data\*|DIPS*
FileKey2=%LocalAppData%\Vivaldi\User Data\*\ClientCertificates|*|REMOVESELF
FileKey3=%LocalAppData%\Vivaldi\User Data\*\Network|TransportSecurity*
FileKey4=%LocalAppData%\Vivaldi\User Data\*\Safe Browsing Network|*|REMOVESELF
FileKey5=%LocalAppData%\Vivaldi\User Data\CertificateRevocation|*|REMOVESELF
FileKey6=%LocalAppData%\Vivaldi\User Data\Crowd Deny|*|REMOVESELF
FileKey7=%LocalAppData%\Vivaldi\User Data\PKIMetadata|*|REMOVESELF
FileKey8=%LocalAppData%\Vivaldi\User Data\Safe Browsing|*|REMOVESELF
FileKey9=%LocalAppData%\Vivaldi\User Data\SafetyTips|*|REMOVESELF
FileKey10=%LocalAppData%\Vivaldi\User Data\SSLErrorAssistant|*|REMOVESELF
FileKey11=%LocalAppData%\Vivaldi\User Data\Subresource Filter|*|REMOVESELF
FileKey12=%LocalAppData%\Vivaldi\User Data\ZxcvbnData|*|REMOVESELF

[Vivaldi Shopping Insights *]
LangSecRef=3033
DetectFile=%LocalAppData%\Vivaldi\User Data
FileKey1=%LocalAppData%\Vivaldi\User Data\*\chrome_cart_db|*|REMOVESELF
FileKey2=%LocalAppData%\Vivaldi\User Data\*\commerce_subscription_db|*|REMOVESELF
FileKey3=%LocalAppData%\Vivaldi\User Data\*\coupon_db|*|REMOVESELF
FileKey4=%LocalAppData%\Vivaldi\User Data\*\discount*_db|*|REMOVESELF
FileKey5=%LocalAppData%\Vivaldi\User Data\*\parcel_tracking_db|*|REMOVESELF

[Vivaldi Telemetry *]
LangSecRef=3033
DetectFile=%LocalAppData%\Vivaldi\User Data
FileKey1=%LocalAppData%\Vivaldi\Application|debug.log
FileKey2=%LocalAppData%\Vivaldi\Application\SetupMetrics|*|REMOVESELF
FileKey3=%LocalAppData%\Vivaldi\User Data|*.pma;LOG;LOG.old|RECURSE
FileKey4=%LocalAppData%\Vivaldi\User Data|*_shutdown_ms.txt;*.log;Last Browser;Last Version;Breadcrumbs;BrowsingTopics*
FileKey5=%LocalAppData%\Vivaldi\User Data\*|*.log
FileKey6=%LocalAppData%\Vivaldi\User Data\*\DataSharing|*|REMOVESELF
FileKey7=%LocalAppData%\Vivaldi\User Data\*\Feature Engagement Tracker|*|REMOVESELF
FileKey8=%LocalAppData%\Vivaldi\User Data\*\Network|Network Persistent State*;Reporting and NEL*;SCT Auditing Pending Reports*
FileKey9=%LocalAppData%\Vivaldi\User Data\*\PersistentOriginTrials|*|REMOVESELF
FileKey10=%LocalAppData%\Vivaldi\User Data\*\Site Characteristics Database|*|REMOVESELF
FileKey11=%LocalAppData%\Vivaldi\User Data\*\VideoDecodeStats|*|REMOVESELF
FileKey12=%LocalAppData%\Vivaldi\User Data\*\WebRTC Logs|*|REMOVESELF
FileKey13=%LocalAppData%\Vivaldi\User Data\*\WebrtcVideoStats|*|REMOVESELF
FileKey14=%LocalAppData%\Vivaldi\User Data\*BrowserMetrics|*|REMOVESELF
FileKey15=%LocalAppData%\Vivaldi\User Data\Crash Reports|*|REMOVESELF
FileKey16=%LocalAppData%\Vivaldi\User Data\Crashpad|*|REMOVESELF
FileKey17=%LocalAppData%\Vivaldi\User Data\Local Traces|*|REMOVESELF
FileKey18=%LocalAppData%\Vivaldi\User Data\OriginTrials|*|REMOVESELF
FileKey19=%LocalAppData%\Vivaldi\User Data\Stability|*|REMOVESELF
FileKey20=%ProgramFiles%\Vivaldi\Application|debug.log
FileKey21=%ProgramFiles%\Vivaldi\Application\SetupMetrics|*|REMOVESELF
RegKey1=HKCU\Software\Vivaldi\BlBeacon|failed_count
RegKey2=HKCU\Software\Vivaldi\BlBeacon|state
RegKey3=HKCU\Software\Vivaldi\Installer|LastChecked
RegKey4=HKCU\Software\Vivaldi\StabilityMetrics

[Vivaldi Updates *]
LangSecRef=3033
DetectFile=%LocalAppData%\Vivaldi\User Data
FileKey1=%LocalAppData%\Vivaldi\Application\*\Installer|*|REMOVESELF
FileKey2=%LocalAppData%\Vivaldi\Application\*\Temp|*|REMOVESELF
FileKey3=%LocalAppData%\Vivaldi\Temp|*|REMOVESELF
FileKey4=%ProgramFiles%\Vivaldi\Application\*\Installer|*|REMOVESELF
FileKey5=%ProgramFiles%\Vivaldi\Application\*\Temp|*|REMOVESELF
FileKey6=%ProgramFiles%\Vivaldi\Temp|*|REMOVESELF

[Vivaldi Web Browsing Cookies *]
LangSecRef=3033
DetectFile=%LocalAppData%\Vivaldi\User Data
FileKey1=%LocalAppData%\Vivaldi\User Data\*\Network|Cookies*;Device Bound Sessions*

[Vivaldi Web Browsing History *]
LangSecRef=3033
DetectFile=%LocalAppData%\Vivaldi\User Data
FileKey1=%LocalAppData%\Vivaldi\User Data\*|History*;Visited Links*;Top Sites*;Network Action Predictor*;shortcuts*

[Vivaldi Web Browsing Session *]
LangSecRef=3033
DetectFile=%LocalAppData%\Vivaldi\User Data
FileKey1=%LocalAppData%\Vivaldi\User Data\*\Extension State|*|REMOVESELF
FileKey2=%LocalAppData%\Vivaldi\User Data\*\Sessions|*|REMOVESELF

[Vivaldi Web Storage *]
LangSecRef=3033
DetectFile=%LocalAppData%\Vivaldi\User Data
FileKey1=%LocalAppData%\Vivaldi\User Data\*\IndexedDB|*|REMOVESELF
FileKey2=%LocalAppData%\Vivaldi\User Data\*\Local Storage\*\IndexedDB\https*|*
FileKey3=%LocalAppData%\Vivaldi\User Data\*\Local Storage\*\Session Storage|*
FileKey4=%LocalAppData%\Vivaldi\User Data\*\Local Storage\*\Storage\ext\*\def\Session Storage|*
FileKey5=%LocalAppData%\Vivaldi\User Data\*\Local Storage\*\WebStorage|*|RECURSE
FileKey6=%LocalAppData%\Vivaldi\User Data\*\Local Storage\leveldb|*|REMOVESELF
FileKey7=%LocalAppData%\Vivaldi\User Data\*\Session Storage|*|REMOVESELF
FileKey8=%LocalAppData%\Vivaldi\User Data\*\WebStorage|*|REMOVESELF
```

**Explanation**
- Two separate transmutations are conducted
- For the first transmutation, the base file is browsers.ini
- For the first transmutation, the source file is browser_value_removals.ini
- For the first transmutation, the output file is browsers.ini (overwriting the base file)
- The `Section` key is removed from each section in the base file
- Sections in the base file not defined in the source file remain unchanged 
- The output is saved to browsers.ini, overwriting the initial base file before the second transmutation  
- For the second transmutation, the base file is browsers.ini (having been modified once already)
- For the second transmutation, the source file is browser_additions.ini 
- For the second transmutation, the output file is browsers.ini (overwriting the base file)
- The key `LangSecRef=3033` is added to each section in the base file that is defined in the source file
- Sections in the base file not defined in the source file remain unchanged 

### [Example 9: Correcting syntax](#example-9-correcting-syntax)

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
- `[360 Secure Browser Bookmarked Websites]` in the base file has `FileKey` added with value `%AppData%\360se6\User Data\*|360Bookmarks*` from the source file
- Sections in the base file not defined in the source file remain unchanged  
- WinappDebug is invoked
- The input file is browsers.ini 
- The output file is browsers.ini (overwriting the input file)
- WinappDebug detects that the added `FileKey` points to the same location as the existing `FileKey1` and merges its parameters into the existing key and removes the now-unneeded `FileKey` which was just added
- The style and syntax of the entry is corrected if there are any additional issues
