# Transmute Examples

Worked examples for [Transmute](readme.md). The [Common Uses](readme.md#common-uses) table in the readme lists the commands. These show what each one does, with real input and output.

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

Several of these examples come from the corrections winapp2ool applies to BrowserBuilder's output. Those files live in [Assembler/Sources/BrowserBuilder](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/Sources/BrowserBuilder), but each example shows the lines it needs.

###### Note: With Syntax on (the default), Transmute writes the winapp2.ini header comments at the top of the output file. The outputs below leave the header out to save space.

## Adding Content

### Example 1: Adding New Sections and Keys

**Context**

BrowserBuilder builds every browser entry from a shared set of scaffolds. That saves an enormous amount of work, but the scaffolds can't know every browser's quirks.

**Intent**

We want better coverage for 360 Secure Browser. That means adding a `FileKey` to the generated `[360 Secure Browser Bookmarked Websites *]` entry, and adding a new entry, `[360 Secure Browser Web Browsing History Backups *]`, for a feature only this browser has.

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

###### Note: *Unlike* replacements and removals, new entries and new keys can share one file

**Command**
```
winapp2ool -transmute -add -1f browsers.ini -2f browser_additions.ini -3f browsers.ini 
```
###### Note: `Add` is the default mode, so you could leave out `-add`. It's here so the command says what it does

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
- `[360 Secure Browser Web Browsing History Backups *]` is added to the base file as the source file defines it
- `[360 Secure Browser Bookmarked Websites *]` in the base file has `FileKey` added with value `%AppData%\360se6\User Data\*|360Bookmarks*` from the source file
- Sections in the base file not defined in the source file remain unchanged

**Notes**

Transmute doesn't fix key numbering. It adds keys exactly as written, like the unnumbered `FileKey` above. Run winapp2.ini files through WinappDebug afterwards to fix their syntax ([Example 9](#example-9-correcting-syntax) does this).

---

## Replacements

### Example 2: Replacing Key Values

**Context**

BrowserBuilder produces some `DetectFile` keys that work in winapp2ool but *not* in CCleaner, because CCleaner doesn't support a wildcard in the folders above the one a `DetectFile` points to. CCleaner misreads a key like `DetectFile2=%LocalAppData%\Packages\Mozilla.Firefox_*\LocalCache\Roaming\Mozilla\Firefox\Profiles`. Our generated Firefox entries always contain this `DetectFile2`, so we want to replace it with a value CCleaner can read.

**Intent**

We want to replace the generated value of `DetectFile2` in `[Mozilla Firefox Autocomplete History *]` with `%LocalAppData%\Packages\Mozilla.Firefox_*`, which CCleaner reads correctly.

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
###### Note: `By Key` is the default Replace mode, so you could leave out `-bykey`. It's here so the command says what it does

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
- `DetectFile2` in the base file's `[Mozilla Firefox Autocomplete History *]` takes the value of `DetectFile2` from the same section in the source file
- Sections in the base file not defined in the source file remain unchanged

**Notes**

This mode replaces key *values* by matching the key's `Name`, which is everything to the left of the `=`. The base file's key name stays as it was. For numbered keys, removing the old key and adding a new one with the value you want often gives more consistent results ([Example 7](#example-7-chaining-operations-together-1)).

---

### Example 3: Replacing Entire Sections

**Context**

You installed the game *And Yet It Moves* from GOG, but the winapp2.ini entry only covers the Steam version. You've written a replacement that covers yours.

**Intent**

We want to replace winapp2.ini's `[And Yet It Moves *]` with the copy we keep separately.

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
###### Note: Without `-3f winapp2.ini`, the output would go to `winapp2-transmuted.ini` and the base file would stay as it was

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
- The base file's `[And Yet It Moves *]` is replaced entirely by the source file's
- Sections in the base file not defined in the source file remain unchanged

---

## Removals

### Example 4: Removing Entire Sections

**Context**

In BrowserBuilder's first version, a generated entry could be incomplete, or target a feature a particular browser doesn't have. Rather than ship entries that clean nothing, we pruned them from the generated set before combining it into winapp2.ini. Arc keeps its pinned tabs in a JSON file that also holds its settings, and winapp2.ini can't clean part of a file without deleting the rest.

###### Note: BrowserBuilder no longer generates this entry, since it never had the data the entry needs. We kept the example because it still shows how the mode works, but it doesn't match BrowserBuilder's current output.

**Intent**

We want to remove `[Arc Pinned Tabs *]` from browsers.ini, because nothing we could add would make it work.

**Files**

###### **Base file (`browsers.ini`)**
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

###### Note: You don't need to include any keys when removing whole sections.

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

You're writing a tool that reads winapp2.ini. `Section` only sorts entries into categories, and your tool doesn't use categories, so you want to remove the `Section` key from `[And Yet It Moves *]` before your tool reads it.

**Intent**

We want to remove the `Section` key from `[And Yet It Moves *]`.

**Files**

###### **Base file (`winapp2.ini`)**
```ini
[And Yet It Moves *]
Section=Games
Detect=HKCU\Software\Valve\Steam\Apps\18700
FileKey1=%AppData%\Broken Rules\And Yet It Moves Steam|console.log
```

###### **Source file (`key_name_removals.ini`)**
```ini
[And Yet It Moves *]
Section=Games
```

###### Note: ByName mode ignores the value in the source file, so `Section=` would remove the key just the same. Writing the value anyway makes the source file easier to read later.

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

BrowserBuilder's DuckDuckGo Browser entries each get two `DetectFile` keys, and neither works in CCleaner, because each has a wildcard in a parent folder (eg. `DetectFile=%LocalAppData%\Packages\*DuckDuckGo*Browser*\LocalState\EBWebView`). Rather than remove one and replace the other's value, we'll remove both and add one value that covers both paths. This example is the first half: removing the keys CCleaner can't read.

**Intent**

We want to remove the two `DetectFile` values without having to know which number (eg. the **1** in `DetectFile1`) each key has.

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
- Every `DetectFile` in the base file's `[DuckDuckGo Autofill Data *]` whose value appears in the same section of the source file is removed
- Sections in the base file not defined in the source file remain unchanged

---

## Advanced

### Example 7: Chaining operations together 1

**Context**

[Example 6](#example-6-removing-keys-by-value) left the entry with no detection at all. This example does the whole job, removing the two old values and adding the new one.

**Intent**

We want one working `DetectFile` in place of the two that Example 6 removes.

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

###### **Source file (`browser_additions.ini`)**
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

This runs two transmutations, each overwriting browsers.ini, so the second builds on the first:

| Step | Mode | Source file | Effect |
|:-|:-|:-|:-|
| 1 | Remove ByKey ByValue | browser_value_removals.ini | Both incompatible `DetectFile` keys are removed from `[DuckDuckGo Autofill Data *]` |
| 2 | Add | browser_additions.ini | `DetectFile=%LocalAppData%\Packages\*DuckDuckGo*Browser*` is added to `[DuckDuckGo Autofill Data *]` |

- Sections in the base file not defined in the source files remain unchanged

**Notes**

The new `DetectFile` sits with the entry's other detection keys instead of at the end. Add mode puts keys at the end of the section, but saving with Syntax on regroups keys by type (see [Output Formatting](readme.md#output-formatting)).

---

### Example 8: Chaining operations together 2

**Context**

When winapp2.ini made its non-CCleaner format the default, each web browser got a `Section` of its own. CCleaner used to show those entries alongside its own entries for the same browser, so the CCleaner flavor of winapp2.ini needs to put them back there.

**Intent**

We want to remove one browser's `Section` keys and give each entry a `LangSecRef` pointing at that browser's CCleaner section.

**Files**

The base file has 22 generated Vivaldi entries. Three are shown here, and the rest follow the same pattern and get the same treatment.

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

###### **Source file (`browser_additions.ini`)**
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

###### Note: `By Key` is the default Remove mode, so you could leave out `-bykey`. It's here so the command says what it does

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

This runs two transmutations, each overwriting browsers.ini, so the second builds on the first:

| Step | Mode | Source file | Effect |
|:-|:-|:-|:-|
| 1 | Remove ByKey ByValue | browser_value_removals.ini | The `Section=Vivaldi Web Browser` key is removed from each listed entry |
| 2 | Add | browser_additions.ini | `LangSecRef=3033` is added to each listed entry |

- Sections in the base file not defined in the source files remain unchanged

**Notes**

The two files need a section for every Vivaldi entry, and they have to be kept in step whenever an entry is added or renamed. The CCleaner flavor now does this with one [key mapping rule](readme.md#key-mapping-rules-map) per browser instead, which covers new entries automatically. This example is retained for demonstration purposes.

---

### Example 9: Correcting syntax

**Context**

[Example 1](#example-1-adding-new-sections-and-keys) left an unnumbered `FileKey` at the end of `[360 Secure Browser Bookmarked Websites *]`. Any change we make to a winapp2.ini file should leave it valid winapp2.ini, so this time we lint the result.

**Intent**

We want to add a `FileKey` to the generated `[360 Secure Browser Bookmarked Websites *]` entry, then fix the syntax so cleaners read the entry correctly.

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

**Commands**
```
winapp2ool -transmute -add -1f browsers.ini -2f browser_additions.ini -3f browsers.ini 
winapp2ool -debug -c -opti -1f browsers.ini -3f browsers.ini
```
###### Note: WinappDebug doesn't merge FileKeys unless you ask. `-opti` turns on its experimental Optimizations rule (the FileKey merger) for this run. See the [WinappDebug readme](../winappdebug/README.md) for details

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
- WinappDebug then runs on browsers.ini, overwriting it with the repaired file
- It sees that the new `FileKey` points at the same folder as `FileKey1`, merges the new key's patterns into `FileKey1`, and removes the new key
- The merged pattern list is sorted alphabetically, as in any other FileKey pattern repair
- Any other style or syntax problems in the entry are fixed too
