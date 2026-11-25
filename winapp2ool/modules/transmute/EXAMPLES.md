# Transmute Examples

With Syntax on (the default), Transmute writes the winapp2.ini header comments at the top of the output file. The outputs below leave the header out.

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
5. [Global Operations](#global-operations)
   - [Example 10: Removing keys from every entry](#example-10-removing-keys-from-every-entry)
   - [Example 11: Mapping one key to another everywhere](#example-11-mapping-one-key-to-another-everywhere)
   - [Example 12: Adding keys to entries by name](#example-12-adding-keys-to-entries-by-name)
   - [Example 13: Mapping one key to several](#example-13-mapping-one-key-to-several)

---

## Adding Content

### Example 1: Adding New Sections and Keys

**Context**

BrowserBuilder builds every browser entry from a shared set of scaffolds. This covers the parts of every browser that are the same, but some browsers require some additional support to cover all of their features. 

**Intent**

We want better coverage for 360 Secure Browser. Here, that requires adding a `FileKey` to the generated `[360 Secure Browser Bookmarked Websites *]` entry, and adding a new entry, `[360 Secure Browser Web Browsing History Backups *]`, for a feature only this browser has.

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

**Command**
```
winapp2ool -transmute -add -1f browsers.ini -2f browser_additions.ini -3f browsers.ini 
```

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
- `[360 Secure Browser Web Browsing History Backups *]` is added to the base file
- `[360 Secure Browser Bookmarked Websites *]` in the base file has `FileKey=%AppData%\360se6\User Data\*|360Bookmarks* ` added to it 

**Notes**

Transmute doesn't fix key numbering. It adds keys exactly as written, like the unnumbered `FileKey` above. Run winapp2.ini files through WinappDebug afterwards to fix their syntax See [Example 9](#example-9-correcting-syntax).

---

## Replacements

### Example 2: Replacing Key Values

**Context**

BrowserBuilder produces some `DetectFile` keys that work in winapp2ool but *not* in CCleaner, because CCleaner doesn't support nested wildcards such as `DetectFile2=%LocalAppData%\Packages\Mozilla.Firefox_*\LocalCache\Roaming\Mozilla\Firefox\Profiles`.

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
- `DetectFile2` in the base file's `[Mozilla Firefox Autocomplete History *]` takes the value of `%LocalAppData%\Packages\Mozilla.Firefox_*`

---

### Example 3: Replacing Entire Sections

**Context**

You installed the game *And Yet It Moves* from GOG, but the official winapp2.ini entry only covers the Steam version. You've written a replacement that covers yours.

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

---

## Removals

### Example 4: Removing Entire Sections

**Context**

In an older version of BrowserBuilder, a generated entry could be incomplete, or target a feature a particular browser doesn't have. Rather than ship entries that clean nothing, we pruned them from the generated set before combining it into winapp2.ini. Arc kept its pinned tabs in a JSON file that also holds its settings, and winapp2.ini can't clean part of a file without deleting the rest.

**Intent**

We want to remove `[Arc Pinned Tabs *]` from browsers.ini

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

---

### Example 5: Remove Keys By Name

**Context**

You're writing a tool that reads winapp2.ini, and your tool doesn't use categories, so you want to remove the `Section` key from `[And Yet It Moves *]` before your tool reads it.

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

---

### Example 6: Removing Keys By Value

**Context**

BrowserBuilder's DuckDuckGo Browser entries each get two `DetectFile` keys, and neither works in CCleaner, because each has a wildcard in a parent folder (eg. `DetectFile=%LocalAppData%\Packages\*DuckDuckGo*Browser*\LocalState\EBWebView`). Rather than remove one and replace the other's value, we'll remove both and add one value that covers both paths.

**Intent**

We want to remove the two `DetectFile` keys 

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
- Every `DetectFile` in the base file's `[DuckDuckGo Autofill Data *]` is removed

---

## Advanced

### Example 7: Chaining operations together 1

**Context**

[Example 6](#example-6-removing-keys-by-value) gave us an entry with no detection criteria.  

**Intent**

We want one working `DetectFile` in place of the two that Example 6 removed.

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
- The base file is browsers.ini
- The source file is browser_value_removals.ini
- The output file is browsers.ini (overwriting the base file)
- Every `DetectFile` in the base file's `[DuckDuckGo Autofill Data *]` is removed
- The second source file is browser_additions.ini 
- `DetectFile=%LocalAppData%\Packages\*DuckDuckGo*Browser*` is added to `[DuckDuckGo Autofill Data *]` in the base file 

---

### Example 8: Chaining operations together 2

**Context**

CCleaner has special categories for some web browsers. To better integrate into CCleaner's UI, the CCleaner flavor of winapp2.ini recategorizes those browsers into their CCleaner categories. 

**Intent**

We want to remove one browser's `Section` keys and give each entry a `LangSecRef` pointing at that browser's CCleaner section.

**Files**

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
```

###### **Source file (`browser_value_removals.ini`)**
```ini
[Vivaldi Autofill Data *]
Section=Vivaldi Web Browser

[Vivaldi Autoplay Preferences *]
Section=Vivaldi Web Browser

[Vivaldi Bookmark Backups *]
Section=Vivaldi Web Browser
```

###### **Source file (`browser_additions.ini`)**
```ini
[Vivaldi Autofill Data *]
LangSecRef=3033

[Vivaldi Autoplay Preferences *]
LangSecRef=3033

[Vivaldi Bookmark Backups *]
LangSecRef=3033
```

**Commands**
```
winapp2ool -transmute -remove -bykey -byvalue -1f browsers.ini -2f browser_value_removals.ini -3f browsers.ini
winapp2ool -transmute -add -1f browsers.ini -2f browser_additions.ini -3f browsers.ini 
```

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
```

**Explanation**
- The base file is browsers.ini
- The first source file is browser_value_removals.ini
- The output file is browsers.ini (overwriting the base file)
- Every `Section` key is removed from each listed entry 
- The second source file is browser_additions.ini
- The output file is browsers.ini (overwriting the base file again)
- `LangSecRef=3033` is added to each listed entry 

This example is provided for demonstration purposes. See [Example 11](#example-11-mapping-one-key-to-another-everywhere) for doing this at scale.

---

### Example 9: Correcting syntax

**Context**

[Example 1](#example-1-adding-new-sections-and-keys) left an unnumbered `FileKey` at the end of `[360 Secure Browser Bookmarked Websites *]`. Any change we make to a winapp2.ini file should leave it valid winapp2.ini.

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
###### Note: `-opti` turns on FileKey merger for this run. See the [WinappDebug readme](../winappdebug/README.md) for details

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
- `FileKey=%AppData%\360se6\User Data\*|360Bookmarks*` is added to `[360 Secure Browser Bookmarked Websites *]` in the base file
- WinappDebug then runs on browsers.ini
- The output file for WinappDebug is browsers.ini (overwriting the input)
- The newly added `FileKey` is merged with the existing `FileKey1` to create  `FileKey1=%AppData%\360se6\User Data\*|BookmarkMergedSurfaceOrdering;bookmarks;360Bookmarks*`
- The extra `FileKey=` is removed as a result of the merger 
- The merged pattern list is sorted alphabetically, creating `FileKey1=%AppData%\360se6\User Data\*|BookmarkMergedSurfaceOrdering;bookmarks;360Bookmarks`

---

## Global Operations

### Example 10: Removing keys from every entry

**Context**

[Example 5](#example-5-remove-keys-by-name) removed the `Section` key from one entry for a tool that doesn't use categories. That tool doesn't want a category on *any* entry. 

**Intent**

We want to remove every `LangSecRef` and `Section` key from winapp2.ini.

**Files**

###### **Base file (`winapp2.ini`)**
```ini
[Adobe Flash Player *]
LangSecRef=3021
DetectFile=%AppData%\Macromedia\Flash Player
FileKey1=%AppData%\Macromedia\Flash Player\#SharedObjects|*|REMOVESELF

[And Yet It Moves *]
Section=Games
Detect=HKCU\Software\Valve\Steam\Apps\18700
FileKey1=%AppData%\Broken Rules\And Yet It Moves Steam|console.log
```

###### **Source file (`key_name_removals.ini`)**
```ini
; Remove ByName ignores values, so these can be left empty
[*]
LangSecRef=
Section=
```

**Command**
```
winapp2ool -transmute -remove -bykey -byname -2f key_name_removals.ini -3f winapp2.ini
```

**Output**

###### **Output file (`winapp2.ini`) after transmutation**

```ini
[Adobe Flash Player *]
DetectFile=%AppData%\Macromedia\Flash Player
FileKey1=%AppData%\Macromedia\Flash Player\#SharedObjects|*|REMOVESELF

[And Yet It Moves *]
Detect=HKCU\Software\Valve\Steam\Apps\18700
FileKey1=%AppData%\Broken Rules\And Yet It Moves Steam|console.log
```

**Explanation**
- The base file is winapp2.ini
- The source file is key_name_removals.ini
- The output file is winapp2.ini (overwriting the base file)
- Every `LangSecRef` and every `Section` key is removed from the base file

---

### Example 11: Mapping one key to another everywhere

**Context**

[Example 8](#example-8-chaining-operations-together-2) moved Vivaldi's entries into CCleaner's Vivaldi section with two source files that list several Vivaldi entries by name. A key mapping rule matches keys by value instead of by entry name, so one rule covers every Vivaldi entry, including ones added later.

**Intent**

We want to replace every `Section=Vivaldi Web Browser` key with `LangSecRef=3033`

**Files**

###### **Base file (`browsers.ini`)**
```ini
[Brave Bookmark Backups *]
Section=Brave Web Browser
DetectFile=%LocalAppData%\BraveSoftware\Brave-Browser*\User Data
FileKey1=%LocalAppData%\BraveSoftware\Brave-Browser*\User Data\*|Bookmarks.bak

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

```

###### **Source file (`browser_key_replacements.ini`)**
```ini
[*Map: Vivaldi]
Match=Section=Vivaldi Web Browser
Replace=LangSecRef=3033
```

**Command**
```
winapp2ool -transmute -replace -bykey -1f browsers.ini -2f browser_key_replacements.ini -3f browsers.ini
```

**Output**

###### **Output file (`browsers.ini`) after transmutation**
```ini
[Brave Bookmark Backups *]
Section=Brave Web Browser
DetectFile=%LocalAppData%\BraveSoftware\Brave-Browser*\User Data
FileKey1=%LocalAppData%\BraveSoftware\Brave-Browser*\User Data\*|Bookmarks.bak

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

```

**Explanation**
- The base file is browsers.ini
- The source file is browser_key_replacements.ini
- The output file is browsers.ini (overwriting the base file)
- Every `Section=Vivaldi Web Browser` key in the base file is replaced with `LangSecRef=3033`

---

### Example 12: Adding keys to entries by name

**Context**

FluentCleaner turns on every entry that has no `Default` key, so the FluentCleaner of winapp2.ini disables some entries by inserting a `Default=False` key into them. 

**Intent**

We want to add `Default=False` to every browser's `Web Browsing Session` entry

**Files**

Three entries are shown. The base file has one `Web Browsing Session` entry for each browser.

###### **Base file (`browsers.ini`)**
```ini
[Thunderbird Web Browsing Session *]
Section=Thunderbird Email Client
DetectFile1=%AppData%\Thunderbird\Profiles
DetectFile2=%LocalAppData%\Packages\MozillaThunderbird.MozillaThunderbirdRelease_h5892qc0xkpca\LocalCache\Roaming\Thunderbird\Profiles
DetectFile3=%LocalAppData%\Packages\MozillaThunderbird.MZLA_h5892qc0xkpca\LocalCache\Roaming\Thunderbird\Profiles
DetectFile4=%LocalAppData%\Packages\MozillaThunderbird.ThunderbirdBeta_h5892qc0xkpca\LocalCache\Roaming\Thunderbird\Profiles
FileKey1=%AppData%\Thunderbird\Profiles\*|session*.json*
FileKey2=%LocalAppData%\Packages\MozillaThunderbird.MozillaThunderbirdRelease_h5892qc0xkpca\LocalCache\Roaming\Thunderbird\Profiles\*|session*.json*
FileKey3=%LocalAppData%\Packages\MozillaThunderbird.MZLA_h5892qc0xkpca\LocalCache\Roaming\Thunderbird\Profiles\*|session*.json*
FileKey4=%LocalAppData%\Packages\MozillaThunderbird.ThunderbirdBeta_h5892qc0xkpca\LocalCache\Roaming\Thunderbird\Profiles\*|session*.json*

[Thunderbird Web Browsing Session Backups *]
Section=Thunderbird Email Client
DetectFile1=%AppData%\Thunderbird\Profiles
DetectFile2=%LocalAppData%\Packages\MozillaThunderbird.MozillaThunderbirdRelease_h5892qc0xkpca\LocalCache\Roaming\Thunderbird\Profiles
DetectFile3=%LocalAppData%\Packages\MozillaThunderbird.MZLA_h5892qc0xkpca\LocalCache\Roaming\Thunderbird\Profiles
DetectFile4=%LocalAppData%\Packages\MozillaThunderbird.ThunderbirdBeta_h5892qc0xkpca\LocalCache\Roaming\Thunderbird\Profiles
FileKey1=%AppData%\Thunderbird\Profiles\*\sessionstore-backups|*
FileKey2=%LocalAppData%\Packages\MozillaThunderbird.MozillaThunderbirdRelease_h5892qc0xkpca\LocalCache\Roaming\Thunderbird\Profiles\*\sessionstore-backups|*
FileKey3=%LocalAppData%\Packages\MozillaThunderbird.MZLA_h5892qc0xkpca\LocalCache\Roaming\Thunderbird\Profiles\*\sessionstore-backups|*
FileKey4=%LocalAppData%\Packages\MozillaThunderbird.ThunderbirdBeta_h5892qc0xkpca\LocalCache\Roaming\Thunderbird\Profiles\*\sessionstore-backups|*

[Vivaldi Web Browsing Session *]
Section=Vivaldi Web Browser
DetectFile=%LocalAppData%\Vivaldi\User Data
FileKey1=%LocalAppData%\Vivaldi\User Data\*\Extension State|*|REMOVESELF
FileKey2=%LocalAppData%\Vivaldi\User Data\*\Sessions|*|REMOVESELF
```

###### **Source file (`fc_additions.ini`)**
```ini
; Select entries whose name ends with " Web Browsing Session *"
; and that belong to a browser, an email client, or an internet suite
[*Name: Web Browsing Session]
Match1=Section=* Web Browser
Match2=Section=* Email Client
Match3=Section=* Internet Suite
Match4=Section=* ArtisBrowser
Default=False
```

**Command**
```
winapp2ool -transmute -add -1f browsers.ini -2f fc_additions.ini -3f browsers.ini
```

**Output**

###### **Output file (`browsers.ini`) after transmutation**
```ini
[Thunderbird Web Browsing Session *]
Section=Thunderbird Email Client
DetectFile1=%AppData%\Thunderbird\Profiles
DetectFile2=%LocalAppData%\Packages\MozillaThunderbird.MozillaThunderbirdRelease_h5892qc0xkpca\LocalCache\Roaming\Thunderbird\Profiles
DetectFile3=%LocalAppData%\Packages\MozillaThunderbird.MZLA_h5892qc0xkpca\LocalCache\Roaming\Thunderbird\Profiles
DetectFile4=%LocalAppData%\Packages\MozillaThunderbird.ThunderbirdBeta_h5892qc0xkpca\LocalCache\Roaming\Thunderbird\Profiles
Default=False
FileKey1=%AppData%\Thunderbird\Profiles\*|session*.json*
FileKey2=%LocalAppData%\Packages\MozillaThunderbird.MozillaThunderbirdRelease_h5892qc0xkpca\LocalCache\Roaming\Thunderbird\Profiles\*|session*.json*
FileKey3=%LocalAppData%\Packages\MozillaThunderbird.MZLA_h5892qc0xkpca\LocalCache\Roaming\Thunderbird\Profiles\*|session*.json*
FileKey4=%LocalAppData%\Packages\MozillaThunderbird.ThunderbirdBeta_h5892qc0xkpca\LocalCache\Roaming\Thunderbird\Profiles\*|session*.json*

[Thunderbird Web Browsing Session Backups *]
Section=Thunderbird Email Client
DetectFile1=%AppData%\Thunderbird\Profiles
DetectFile2=%LocalAppData%\Packages\MozillaThunderbird.MozillaThunderbirdRelease_h5892qc0xkpca\LocalCache\Roaming\Thunderbird\Profiles
DetectFile3=%LocalAppData%\Packages\MozillaThunderbird.MZLA_h5892qc0xkpca\LocalCache\Roaming\Thunderbird\Profiles
DetectFile4=%LocalAppData%\Packages\MozillaThunderbird.ThunderbirdBeta_h5892qc0xkpca\LocalCache\Roaming\Thunderbird\Profiles
FileKey1=%AppData%\Thunderbird\Profiles\*\sessionstore-backups|*
FileKey2=%LocalAppData%\Packages\MozillaThunderbird.MozillaThunderbirdRelease_h5892qc0xkpca\LocalCache\Roaming\Thunderbird\Profiles\*\sessionstore-backups|*
FileKey3=%LocalAppData%\Packages\MozillaThunderbird.MZLA_h5892qc0xkpca\LocalCache\Roaming\Thunderbird\Profiles\*\sessionstore-backups|*
FileKey4=%LocalAppData%\Packages\MozillaThunderbird.ThunderbirdBeta_h5892qc0xkpca\LocalCache\Roaming\Thunderbird\Profiles\*\sessionstore-backups|*

[Vivaldi Web Browsing Session *]
Section=Vivaldi Web Browser
DetectFile=%LocalAppData%\Vivaldi\User Data
Default=False
FileKey1=%LocalAppData%\Vivaldi\User Data\*\Extension State|*|REMOVESELF
FileKey2=%LocalAppData%\Vivaldi\User Data\*\Sessions|*|REMOVESELF
```

**Explanation**
- The base file is browsers.ini
- The source file is fc_additions.ini
- The output file is browsers.ini (overwriting the base file)
- `[*Name: Web Browsing Session]` selects each section whose name ends with ` Web Browsing Session *` and also has a key matching one of its `Match` keys
- `[Thunderbird Web Browsing Session *]` matches `Section=* Email Client` and `[Vivaldi Web Browsing Session *]` matches `Section=* Web Browser`, so each gets `Default=False`
- `[Thunderbird Web Browsing Session Backups *]` ends with ` Session Backups *`, so this rule doesn't select it

---

### Example 13: Mapping one key to several

**Context**

[Example 7](#example-7-chaining-operations-together-1) left `[DuckDuckGo Autofill Data *]` with a single wildcard `DetectFile`. System Ninja doesn't support wildcards in `DetectFile`, so its flavor needs one `DetectFile` for each DuckDuckGo package.

**Intent**

We want to replace every `DetectFile=%LocalAppData%\Packages\*DuckDuckGo*Browser*` with two `DetectFile` keys, one per package

**Files**

###### **Base file (`browsers.ini`)**
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

###### **Source file (`sn_key_replacements.ini`)**
```ini
[*Map: DuckDuckGo DetectFile]
Match=DetectFile=%LocalAppData%\Packages\*DuckDuckGo*Browser*
Replace1=DetectFile1=%LocalAppData%\Packages\63909DuckDuckGoInc.DuckDuckGoPrivateBrowser_qzdjx70tn762j
Replace2=DetectFile2=%LocalAppData%\Packages\DuckDuckGo.DesktopBrowser_ya2fgkz3nks94
```

**Command**
```
winapp2ool -transmute -replace -bykey -1f browsers.ini -2f sn_key_replacements.ini -3f browsers.ini
```

**Output**

###### **Output file (`browsers.ini`) after transmutation**
```ini
[DuckDuckGo Autofill Data *]
Section=DuckDuckGo Web Browser
DetectFile1=%LocalAppData%\Packages\63909DuckDuckGoInc.DuckDuckGoPrivateBrowser_qzdjx70tn762j
DetectFile2=%LocalAppData%\Packages\DuckDuckGo.DesktopBrowser_ya2fgkz3nks94
FileKey1=%LocalAppData%\Packages\*DuckDuckGo*Browser*\LocalState\EBWebView\*|*Web Data
FileKey2=%LocalAppData%\Packages\*DuckDuckGo*Browser*\LocalState\EBWebView\*\AutoFill*|*|REMOVESELF
FileKey3=%LocalAppData%\Packages\*DuckDuckGo*Browser*\LocalState\EBWebView\AutoFill*|*|REMOVESELF
FileKey4=%LocalAppData%\Packages\*DuckDuckGo*Browser*\LocalState\internalEnvironment\EBWebView\*|*Web Data
FileKey5=%LocalAppData%\Packages\*DuckDuckGo*Browser*\LocalState\internalEnvironment\EBWebView\*\AutoFill*|*|REMOVESELF
FileKey6=%LocalAppData%\Packages\*DuckDuckGo*Browser*\LocalState\internalEnvironment\EBWebView\AutoFill*|*|REMOVESELF
```

**Explanation**
- The base file is browsers.ini
- The source file is sn_key_replacements.ini
- The output file is browsers.ini (overwriting the base file)
- `[*Map: DuckDuckGo DetectFile]` finds every `DetectFile` whose value is exactly `%LocalAppData%\Packages\*DuckDuckGo*Browser*`
- `Replace1` takes the matched key's place, and `Replace2` is inserted after it

**Notes**

The `Match` value is compared as plain text: the `*`s in it are part of the value, not treated as wildcards.  
