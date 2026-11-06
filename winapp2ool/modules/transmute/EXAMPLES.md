# Transmute Examples

Worked examples for [Transmute](readme.md). The [Common Uses](readme.md#common-uses) table in the readme lists the commands; these show what each one does, with real input and output.

# Table of Contents

1. [Adding Content](#adding-content)
   - [Example 1: Adding New Sections and Keys](#example-1-adding-new-sections-and-keys)
2. [Replacements](#replacements)
   - [Example 2: Replacing Key Values](#example-2-replacing-key-values)
   - [Example 3: Replacing Entire Sections](#example-3-replacing-entire-sections)
3. [Removals](#removals)
   - [Example 4: Removing Entire Sections](#example-4-removing-entire-sections)
   - [Example 5: Remove Keys By Name](#example-5-remove-keys-by-name)
   - [Example 6: Removing Keys By Value](#example-6-removing-keys-by-value)
4. [Advanced](#advanced)
   - [Example 7: Chaining operations together 1](#example-7-chaining-operations-together-1)
   - [Example 8: Chaining operations together 2](#example-8-chaining-operations-together-2)
   - [Example 9: Correcting syntax](#example-9-correcting-syntax)

---

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

The added `DetectFile` appears alongside the entry's other detection keys rather than at the end of the entry: Add mode appends keys to the end of the section, but saving with Syntax enabled regroups keys by type (see [Output Formatting](readme.md#output-formatting)).

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
winapp2ool -debug -c -opti -1f browsers.ini -3f browsers.ini
```
###### Note: WinappDebug does not perform the optimization shown below by default. The `-opti` flag enables its experimental Optimizations rule (FileKey merger) for this run. See the [WinappDebug readme](../winappdebug/README.md) for details

**Output**

###### **Output file (`browsers.ini`) after transmutation and linting**

```ini
[360 Secure Browser Bookmarked Websites *]
Section=.360 Secure Browser Web Browser
DetectFile=%AppData%\360se6\User Data
FileKey1=%AppData%\360se6\User Data\*|360Bookmarks*;BookmarkMergedSurfaceOrdering;bookmarks
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
- The merged parameter list is sorted alphabetically, like any other FileKey parameter repair
- The style and syntax of the entry is corrected if there are any additional issues
