# BrowserBuilder Examples

Worked examples for [BrowserBuilder](readme.md). The [Common Uses](readme.md#common-uses) table in the readme lists the commands; these show what each one does, with real input and output.

# Table of Contents

1. [Example 1: Generating an Entry](#example-1-generating-an-entry)
2. [Example 2: The Cross-Product](#example-2-the-cross-product)
3. [Example 3: Multiple UserDataPaths](#example-3-multiple-userdatapaths)
4. [Example 4: TruncateDetect](#example-4-truncatedetect)
5. [Example 5: Registry Scaffolds and RequiresRegistryRoot](#example-5-registry-scaffolds-and-requiresregistryroot)
6. [Example 6: UserDataParent and the ProgramFiles Mirror](#example-6-userdataparent-and-the-programfiles-mirror)
7. [Example 7: LocalUserDataPath (Gecko)](#example-7-localuserdatapath-gecko)
8. [Example 8: Generating Entries for a Portable Browser](#example-8-generating-entries-for-a-portable-browser)

---

The examples below use the real rulesets from the winapp2.ini project, found [here](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/BrowserBuilder). Relevant excerpts are provided on this page, and every output shown is the module's real output. Unless stated otherwise, the source directory contains only the rulesets (no flavor files), and the [generated comment header](readme.md#output-format) is omitted from outputs for brevity.

## Example 1: Generating an Entry

**Context**

The simplest possible case: one browser crossed with one scaffold. Chromium keeps backup copies of the bookmarks file in each profile, and we want an entry that cleans them.

**Intent**

We want to generate a `Bookmark Backups` entry for Chromium.

**Files**

###### **Ruleset (`chromium.ini`), relevant sections**

```ini
[BrowserInfo: Chromium]
Section=Chromium Web Browser
UserDataPath=%LocalAppData%\Chromium\User Data
RegistryRoot=HKCU\Software\Chromium

[EntryScaffold: Bookmark Backups]
FileKeyBase=<UserDataPath>\*|Bookmarks.bak
```

**Command**

```
winapp2ool -browserbuilder -1d C:\rulesets
```

**Output**

###### **Output file (`browsers.ini`), relevant entry**

```ini
[Chromium Bookmark Backups *]
Section=Chromium Web Browser
DetectFile=%LocalAppData%\Chromium\User Data
FileKey1=%LocalAppData%\Chromium\User Data\*|Bookmarks.bak
```

**Explanation**

- The entry name is the browser name + the scaffold name + ` *`: `Chromium` + `Bookmark Backups` → `[Chromium Bookmark Backups *]`
- `Section=` is copied from the `BrowserInfo`
- `DetectFile=` is the browser's `UserDataPath`
- The `FileKeyBase` template becomes `FileKey1` with `<UserDataPath>` substituted
- The `RegistryRoot` key goes unused here. This scaffold has no `RegKeyBase` templates

---

## Example 2: The Cross-Product

**Context**

The point of BrowserBuilder is that browsers and templates multiply: every scaffold is applied to every browser in the same ruleset. Adding one browser adds a full family of entries; adding one scaffold extends every browser at once.

**Intent**

We want to see what two browsers × two scaffolds produce.

**Files**

###### **Ruleset (`chromium.ini`)**

```ini
[BrowserInfo: CCleaner Browser]
Section=CCleaner Web Browser
UserDataPath=%LocalAppData%\CCleaner Browser\User Data
RegistryRoot=HKCU\Software\Piriform\Browser

[BrowserInfo: Chromium]
Section=Chromium Web Browser
UserDataPath=%LocalAppData%\Chromium\User Data
RegistryRoot=HKCU\Software\Chromium

[EntryScaffold: Autoplay Preferences]
FileKeyBase=<UserDataPath>\MEIPreload|*|REMOVESELF

[EntryScaffold: Bookmark Backups]
FileKeyBase=<UserDataPath>\*|Bookmarks.bak
```

**Command**

```
winapp2ool -browserbuilder -1d C:\rulesets
```

**Output**

###### **Output file (`browsers.ini`)**

```ini
[CCleaner Browser Autoplay Preferences *]
Section=CCleaner Web Browser
DetectFile=%LocalAppData%\CCleaner Browser\User Data
FileKey1=%LocalAppData%\CCleaner Browser\User Data\MEIPreload|*|REMOVESELF

[CCleaner Browser Bookmark Backups *]
Section=CCleaner Web Browser
DetectFile=%LocalAppData%\CCleaner Browser\User Data
FileKey1=%LocalAppData%\CCleaner Browser\User Data\*|Bookmarks.bak

[Chromium Autoplay Preferences *]
Section=Chromium Web Browser
DetectFile=%LocalAppData%\Chromium\User Data
FileKey1=%LocalAppData%\Chromium\User Data\MEIPreload|*|REMOVESELF

[Chromium Bookmark Backups *]
Section=Chromium Web Browser
DetectFile=%LocalAppData%\Chromium\User Data
FileKey1=%LocalAppData%\Chromium\User Data\*|Bookmarks.bak
```

**Explanation**

- 2 browsers × 2 scaffolds = 4 entries, each personalized with its browser's paths and category

---

## Example 3: Multiple UserDataPaths

**Context**

Mozilla Firefox can be installed standalone (profiles under `%AppData%`) or from the Microsoft Store (profiles under the package's `LocalCache`). Both locations need coverage, and a machine may have either.

**Intent**

We want one entry per scaffold that detects and cleans both install types.

**Files**

###### **Ruleset (`gecko.ini`)**

```ini
[BrowserInfo: Mozilla Firefox]
Section=Mozilla Firefox Web Browser
UserDataPath=%AppData%\Mozilla\Firefox\Profiles
UserDataPath=%LocalAppData%\Packages\Mozilla.Firefox_n80bbvh6b1yt2\LocalCache\Roaming\Mozilla\Firefox\Profiles

[EntryScaffold: Autocomplete History]
FileKeyBase=<UserDataPath>\*|formhistory*
```

**Command**

```
winapp2ool -browserbuilder -1d C:\rulesets
```

**Output**

###### **Output file (`browsers.ini`)**

```ini
[Mozilla Firefox Autocomplete History *]
Section=Mozilla Firefox Web Browser
DetectFile1=%AppData%\Mozilla\Firefox\Profiles
DetectFile2=%LocalAppData%\Packages\Mozilla.Firefox_n80bbvh6b1yt2\LocalCache\Roaming\Mozilla\Firefox\Profiles
FileKey1=%AppData%\Mozilla\Firefox\Profiles\*|formhistory*
FileKey2=%LocalAppData%\Packages\Mozilla.Firefox_n80bbvh6b1yt2\LocalCache\Roaming\Mozilla\Firefox\Profiles\*|formhistory*
```

**Explanation**

- With multiple `UserDataPath` keys, the `DetectFile`s become numbered: `DetectFile1`, `DetectFile2`, one per path
- Every `FileKeyBase` template is expanded once per path, so one template line produced two FileKeys

---

## Example 4: TruncateDetect

**Context**

Brave's stable, beta, and nightly channels store data under `Brave-Browser`, `Brave-Browser-Beta`, and `Brave-Browser-Nightly` respectively, which the ruleset covers with a single wildcard: `%LocalAppData%\BraveSoftware\Brave-Browser*\User Data`. But some cleaning tools reject a wildcard in a `DetectFile`'s parent path, so detecting on the full `UserDataPath` would break.

**Intent**

We want the generated `DetectFile` to point at the wildcarded parent directory instead of the `User Data` folder inside it.

**Files**

###### **Ruleset (`chromium.ini`)**

```ini
[BrowserInfo: Brave]
Section=Brave Web Browser
TruncateDetect=True
UserDataPath=%LocalAppData%\BraveSoftware\Brave-Browser*\User Data
RegistryRoot=HKCU\Software\BraveSoftware\Brave-Browser
RegistryRoot=HKCU\Software\BraveSoftware\Brave-Browser-Beta
RegistryRoot=HKCU\Software\BraveSoftware\Brave-Browser-Nightly

[EntryScaffold: Bookmark Backups]
FileKeyBase=<UserDataPath>\*|Bookmarks.bak
```

**Command**

```
winapp2ool -browserbuilder -1d C:\rulesets
```

**Output**

###### **Output file (`browsers.ini`)**

```ini
[Brave Bookmark Backups *]
Section=Brave Web Browser
DetectFile=%LocalAppData%\BraveSoftware\Brave-Browser*
FileKey1=%LocalAppData%\BraveSoftware\Brave-Browser*\User Data\*|Bookmarks.bak
```

**Explanation**

- The `DetectFile` now points at the parent path, `Brave-Browser*`, rather than `Brave-Browser*\User Data`

---

## Example 5: Registry Scaffolds and RequiresRegistryRoot

**Context**

Chromium browsers record pinned-tab state in the registry under the browser vendor's root key. Not every browser in the ruleset has a known registry root and generating a pinned-tabs entry with no keys in it would produce a broken entry.

**Intent**

We want a `Pinned Tabs` entry for every browser with a known registry root, and no entry at all for the rest.

**Files**

###### **Ruleset (`chromium.ini`)**

```ini
[BrowserInfo: Arc]
Section=Arc Web Browser
UserDataPath=%LocalAppData%\Packages\TheBrowserCompany.Arc_ttt1ap7aakyb4\LocalCache\Local\Arc\User Data

[BrowserInfo: Brave]
Section=Brave Web Browser
TruncateDetect=True
UserDataPath=%LocalAppData%\BraveSoftware\Brave-Browser*\User Data
RegistryRoot=HKCU\Software\BraveSoftware\Brave-Browser
RegistryRoot=HKCU\Software\BraveSoftware\Brave-Browser-Beta
RegistryRoot=HKCU\Software\BraveSoftware\Brave-Browser-Nightly

[EntryScaffold: Pinned Tabs]
RequiresRegistryRoot=True
RegKeyBase=<RegistryRoot>\PreferenceMACs\Default|pinned_tabs
```

**Command**

```
winapp2ool -browserbuilder -1d C:\rulesets
```

**Output**

###### **Output file (`browsers.ini`)**

```ini
[Brave Pinned Tabs *]
Section=Brave Web Browser
DetectFile=%LocalAppData%\BraveSoftware\Brave-Browser*
RegKey1=HKCU\Software\BraveSoftware\Brave-Browser\PreferenceMACs\Default|pinned_tabs
RegKey2=HKCU\Software\BraveSoftware\Brave-Browser-Beta\PreferenceMACs\Default|pinned_tabs
RegKey3=HKCU\Software\BraveSoftware\Brave-Browser-Nightly\PreferenceMACs\Default|pinned_tabs
```

There is no `[Arc Pinned Tabs *]` in the output. During the run, the module reports:

```
Skipping Arc for Pinned Tabs: scaffold requires RegistryRoot
```

**Explanation**

- One `RegKey` is generated per `RegistryRoot` per `RegKeyBase` template: Brave's three roots × one template = `RegKey1`-`RegKey3`
- `RequiresRegistryRoot` excludes browsers without a `RegistryRoot` from this scaffold entirely, with a warning naming the browser and scaffold

---

## Example 6: UserDataParent and the ProgramFiles Mirror

**Context**

Some browser files live *outside* the user data directory (e.g. installer logs, setup metrics). When a Chromium browser is installed "for all users", those application files live under `%ProgramFiles%` instead of the per-user location.

**Intent**

We want the Telemetry entry to clean the application directory for both per-user and all-users installations, without writing both variants into the scaffold.

**Files**

###### **Ruleset (`chromium.ini`)**

```ini
[BrowserInfo: Chromium]
Section=Chromium Web Browser
UserDataPath=%LocalAppData%\Chromium\User Data
RegistryRoot=HKCU\Software\Chromium

[EntryScaffold: Telemetry]
FileKeyBase=<UserDataParent>\Application|debug.log
FileKeyBase=<UserDataParent>\Application\SetupMetrics|*|REMOVESELF
```

**Command**

```
winapp2ool -browserbuilder -1d C:\rulesets
```

**Output**

###### **Output file (`browsers.ini`), trimmed to match above**

```ini
[Chromium Telemetry *]
Section=Chromium Web Browser
DetectFile=%LocalAppData%\Chromium\User Data
FileKey1=%LocalAppData%\Chromium\Application|debug.log
FileKey2=%LocalAppData%\Chromium\Application\SetupMetrics|*|REMOVESELF
; ... 
FileKey20=%ProgramFiles%\Chromium\Application|debug.log
FileKey21=%ProgramFiles%\Chromium\Application\SetupMetrics|*|REMOVESELF
```

**Explanation**

- `<UserDataParent>` is the parent of `UserDataPath`. Here: `%LocalAppData%\Chromium`
- Each `<UserDataParent>` template produced two `FileKeys`: the substituted path, and a mirror with the user-profile variable swapped for `%ProgramFiles%`
- The mirror is Chromium-only; Gecko browsers get just the direct substitution

---

## Example 7: LocalUserDataPath (Gecko)

**Context**

Gecko browsers split their profile data: settings and history roam under `%AppData%`, while caches and other machine-local data sit in a mirrored path under `%LocalAppData%`. Cache-cleaning templates need to target the local mirror, but `BrowserInfo` sections only declare the roaming path.

**Intent**

We want the Caches scaffold to clean the `%LocalAppData%` mirror without every browser declaring a second path.

**Files**

###### **Ruleset (`gecko.ini`)**

```ini
[BrowserInfo: Waterfox]
Section=Waterfox Web Browser
UserDataPath=%AppData%\Waterfox\Profiles

[EntryScaffold: Caches]
FileKeyBase=<LocalUserDataPath>\*\*cache*|*|REMOVESELF
FileKeyBase=<LocalUserDataPath>\*\chrome_debugger_profile|*|REMOVESELF
FileKeyBase=<LocalUserDataPath>\*\thumbnails|*|REMOVESELF
FileKeyBase=<UserDataPath>\*|*.corrupt|RECURSE
FileKeyBase=<UserDataPath>\*|AlternateServices.txt;notificationstore.json;parent.lock;serviceworker.txt;webappsstore.sqlite;cert9.db;ClientAuthRememberList.txt;SiteSecurityServiceState*
FileKeyBase=<UserDataPath>\*\notificationstore|*
FileKeyBase=<UserDataPath>\*\security_state|*
FileKeyBase=<UserDataPath>\*\shader-cache|*
FileKeyBase=<UserDataPath>\*\storage\temporary|*|RECURSE
```

**Command**

```
winapp2ool -browserbuilder -1d C:\rulesets
```

**Output**

###### **Output file (`browsers.ini`)**

```ini
[Waterfox Caches *]
Section=Waterfox Web Browser
DetectFile=%AppData%\Waterfox\Profiles
FileKey1=%AppData%\Waterfox\Profiles\*|*.corrupt|RECURSE
FileKey2=%AppData%\Waterfox\Profiles\*|AlternateServices.txt;cert9.db;ClientAuthRememberList.txt;notificationstore.json;parent.lock;serviceworker.txt;SiteSecurityServiceState*;webappsstore.sqlite
FileKey3=%AppData%\Waterfox\Profiles\*\notificationstore|*
FileKey4=%AppData%\Waterfox\Profiles\*\security_state|*
FileKey5=%AppData%\Waterfox\Profiles\*\shader-cache|*
FileKey6=%AppData%\Waterfox\Profiles\*\storage\temporary|*|RECURSE
FileKey7=%LocalAppData%\Waterfox\Profiles\*\*cache*|*|REMOVESELF
FileKey8=%LocalAppData%\Waterfox\Profiles\*\chrome_debugger_profile|*|REMOVESELF
FileKey9=%LocalAppData%\Waterfox\Profiles\*\thumbnails|*|REMOVESELF
```

**Explanation**

- `<LocalUserDataPath>` took the declared `UserDataPath` and swapped `%AppData%` → `%LocalAppData%`, producing FileKey7-9
- The `DetectFile` still only points at the roaming path
- The keys and their parameters have been reordered relative to their templates as a result of BrowserBuilder running its output through WinappDebug

---

## Example 8: Generating Entries for a Portable Browser

**Context**

You run Google Chrome Portable from `D:\PortableApps`. The distributed winapp2.ini doesn't cover this: its Chrome entries detect and clean the standard install locations. BrowserBuilder can generate entries for *your* install path, and because scaffolds are just text, you can copy whichever ones you care about from the [project's chromium.ini](https://github.com/MoscaDotTo/Winapp2/blob/master/Assembler/BrowserBuilder/chromium.ini).

**Intent**

We want history and cookie cleaning entries for a portable Chrome, then we want them merged into our local winapp2.ini.

**Files**

###### **Ruleset (`C:\myrules\chromium.ini`)**

```ini
[BrowserInfo: Chrome Portable]
Section=Chrome Portable Web Browser
UserDataPath=D:\PortableApps\GoogleChromePortable\Data\profile

[EntryScaffold: Web Browsing History]
FileKeyBase=<UserDataPath>\*|History*;Visited Links*;Top Sites*;Network Action Predictor*;shortcuts*

[EntryScaffold: Web Browsing Cookies]
FileKeyBase=<UserDataPath>\*\Network|Cookies*;Device Bound Sessions*
```

**Commands**

```
winapp2ool -browserbuilder -1d C:\myrules -2d C:\myrules
winapp2ool -transmute -add -1f winapp2.ini -2d C:\myrules -2f browsers.ini -3f winapp2.ini
```

**Output**

###### **Generated file (`C:\myrules\browsers.ini`) after the first command**

```ini
[Chrome Portable Web Browsing Cookies *]
Section=Chrome Portable Web Browser
DetectFile=D:\PortableApps\GoogleChromePortable\Data\profile
FileKey1=D:\PortableApps\GoogleChromePortable\Data\profile\*\Network|Cookies*;Device Bound Sessions*

[Chrome Portable Web Browsing History *]
Section=Chrome Portable Web Browser
DetectFile=D:\PortableApps\GoogleChromePortable\Data\profile
FileKey1=D:\PortableApps\GoogleChromePortable\Data\profile\*|History*;Network Action Predictor*;shortcuts*;Top Sites*;Visited Links*
```

**Explanation**

Two steps are chained:

| Step | Module | Effect |
|:-|:-|:-|
| 1 | BrowserBuilder | Generates the two entries above from the hand-written ruleset |
| 2 | Transmute (Add) | Merges the generated entries into your local `winapp2.ini` in place |

###### Notes

- `UserDataPath` and `RegistryRoot` values accept absolute paths and also environment variables
- Because your additions live in the ruleset, you can regenerate and re-merge after each winapp2.ini update
