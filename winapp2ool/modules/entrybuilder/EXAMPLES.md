# EntryBuilder Examples

Worked examples for [EntryBuilder](readme.md). The [Common Uses](readme.md#common-uses) table in the readme lists the commands; these show what each one does, with real input and output.

# Table of Contents

1. [Pass-Through and Variables](#pass-through-and-variables)
   - [Example 1: A pass-through entry with a single-value variable](#example-1-a-pass-through-entry-with-a-single-value-variable)
   - [Example 2: Automatic detection from the reserved Root variable](#example-2-automatic-detection-from-the-reserved-root-variable)
   - [Example 3: List-variable fan-out](#example-3-list-variable-fan-out)
   - [Example 4: Co-variance of repeated variables](#example-4-co-variance-of-repeated-variables)
   - [Example 5: Inline lists](#example-5-inline-lists)
   - [Example 6: Nested references](#example-6-nested-references)
   - [Example 7: Literal angle brackets in registry keys](#example-7-literal-angle-brackets-in-registry-keys)
2. [Scaffold Families](#scaffold-families)
   - [Example 8: Default WebView scaffolds](#example-8-default-webview-scaffolds)
   - [Example 9: Everything except All with exclusions](#example-9-everything-except-all-with-exclusions)
   - [Example 10: QtWebEngine across several profiles](#example-10-qtwebengine-across-several-profiles)
   - [Example 11: An Electron app with an updater cache](#example-11-an-electron-app-with-an-updater-cache)
   - [Example 12: An Electron app with session partitions](#example-12-an-electron-app-with-session-partitions)
3. [Housekeeping](#housekeeping)
   - [Example 13: Skipping a defunct entry](#example-13-skipping-a-defunct-entry)

---

Most of these examples are simplified versions of entries in [`Assembler/Sources/EntryBuilder`](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/Sources/EntryBuilder), and a few apps are made up. Each one works the same way: put the source snippet in an `.ini` file (each example names the letter file it would live in), run

```
winapp2ool -entrybuilder -1d <folder containing the source file>
```

and open `entrybuilder.ini` in the current directory. The scaffold examples (8 to 12) also need `-3d` pointing at the shared catalogs, as the commands in Examples 8 to 11 show.

###### Note: These outputs show the keys as EntryBuilder generates them, before WinappDebug's normalization pass sorts and renumbers each key family, and without the file header. The set of keys is exactly what you'll get. Long fan-outs are cut short with `; ...` comments.

## Pass-Through and Variables

### Example 1: A pass-through entry with a single-value variable

**Context**

A simple game with one log file. The entry is already plain winapp2 syntax; the only annoyance is writing the same long path twice.

**Intent**

We want the standard entry with the path written once. We name the variable `DiskRoot`, *not* `Root`, because we're writing detection by hand and don't need [reserved-`Root` inference](readme.md#the-reserved-root-variable). Here it would do no harm, but pick the name on purpose.

**Files**

###### **Source (`A.ini`)**

```ini
[Annie's Millions *]
DiskRoot=%AppData%\PoBros\Annies Millions
Section=Games
DetectFile=<DiskRoot>
FileKey=<DiskRoot>|logfile.txt
```

**Command**

```
winapp2ool -entrybuilder -1d ..\..\Assembler\Sources\EntryBuilder
```

###### Note: `-1d` points at the directory containing the source files; `entrybuilder.ini` is written to the current directory

**Output**

```ini
[Annie's Millions *]
Section=Games
DetectFile=%AppData%\PoBros\Annies Millions
FileKey1=%AppData%\PoBros\Annies Millions|logfile.txt
```

**Explanation**

- `DiskRoot=` isn't a key EntryBuilder knows, so it becomes a single-value variable
- `<DiskRoot>` expands in both keys; a single-value variable produces no fan-out
- The un-numbered `FileKey=` is renumbered to `FileKey1`; the single `DetectFile` stays bare
- `DiskRoot=` itself is consumed and does not appear in the output

---

### Example 2: Automatic detection from the reserved Root variable

**Context**

Most entries anchor both their detection and their cleaning targets on one path. Declaring that path as `Root=` lets EntryBuilder write the detection key for you.

**Intent**

We want a complete entry for a game whose logs live under its install folder, without writing a detection key.

**Files**

###### **Source (`A.ini`)**

```ini
[A Hat in Time *]
Root=%ProgramFiles%\Steam\steamapps\common\HatinTime\HatinTimeGame

Section=Games
FileKeyBase=<Root>\Logs|*
```

**Output**

```ini
[A Hat in Time *]
Section=Games
DetectFile=%ProgramFiles%\Steam\steamapps\common\HatinTime\HatinTimeGame
FileKey1=%ProgramFiles%\Steam\steamapps\common\HatinTime\HatinTimeGame\Logs|*
```

**Explanation**

- `Root=` is the reserved variable: EntryBuilder classifies `%ProgramFiles%\Steam\steamapps\common\HatinTime\HatinTimeGame` as a filesystem path and injects `DetectFile=<Root>`, which then expands normally
- A registry-hive value (`Root=HKCU\Software\Valve\Steam\Apps\253230`) would have produced `Detect=` instead
- A multi-valued `Root=%AppData%\Foo,%LocalAppData%\Foo` would fan the injected token into `DetectFile1` / `DetectFile2`

**Notes**

The log records the inference: `[A Hat in Time *] inferred DetectFile=<Root>`. If the entry already contained a hand-written `DetectFile=<Root>`, nothing would be added (inference never duplicates).

---

### Example 3: List-variable fan-out

**Context**

Microsoft Office entries cover many old versions, each with its own registry key. Written by hand, that's dozens of nearly identical lines, and every one of them needs touching when a version comes out.

**Intent**

We want one `Detect` per Office version and one `RegKey` per (version, MRU-slot) combination, declared once.

**Files**

###### **Source (`M.ini`)**

```ini
[Microsoft Outlook *]
Versions=9.0,10.0,11.0,12.0,14.0,15.0,16.0
MRUNum=1,2,3
RegRoot=HKCU\Software\Microsoft\Office
LangSecRef=3024
Detect=<RegRoot>\<Versions>\Outlook
RegKeyBase=<RegRoot>\<Versions>\Outlook\Office Finder|<MRUNum>
```

**Output**

```ini
[Microsoft Outlook *]
LangSecRef=3024
Detect1=HKCU\Software\Microsoft\Office\9.0\Outlook
Detect2=HKCU\Software\Microsoft\Office\10.0\Outlook
; ... Detect3 – Detect6 ...
Detect7=HKCU\Software\Microsoft\Office\16.0\Outlook
RegKey1=HKCU\Software\Microsoft\Office\9.0\Outlook\Office Finder|1
RegKey2=HKCU\Software\Microsoft\Office\9.0\Outlook\Office Finder|2
RegKey3=HKCU\Software\Microsoft\Office\9.0\Outlook\Office Finder|3
RegKey4=HKCU\Software\Microsoft\Office\10.0\Outlook\Office Finder|1
; ... RegKey5 – RegKey20, version-major / MRU-minor ...
RegKey21=HKCU\Software\Microsoft\Office\16.0\Outlook\Office Finder|3
```

**Explanation**

- `Detect` references one 7-value variable → 7 keys, numbered `Detect1`–`Detect7`
- `RegKeyBase` references two variables → 7 × 3 = 21 keys; the leftmost reference (`<Versions>`) varies slowest
- `<RegRoot>` is single-valued, so it abbreviates without multiplying
- Supporting the next Office version means adding one item to `Versions=`, and both key families grow with it

**Notes**

WinappDebug sorts numbers by value, so `..\10.0` comes after `..\9.0` and the version order shown here survives normalization. The note at the top about pre-normalization output only matters when declaration order differs from sort order. [Example 4](#example-4-co-variance-of-repeated-variables)'s `Stable,Beta` keys, for instance, come back renumbered with `Beta` first.

---

### Example 4: Co-variance of repeated variables

**Context**

Some applications put the same version number in a path twice, once as a folder name and once inside a file name. The two always agree: version 1.0's folder never holds version 2.0's log.

**Intent**

We want one cache key per (version, channel) combination, and one log key per version, where the version in the file name always matches the version in the folder.

**Files**

###### **Source (`H.ini`)**

```ini
[Hazel Vault *]
Versions=1.0,2.0
Channels=Stable,Beta
Root=%AppData%\Hazel Vault

LangSecRef=3021
FileKeyBase=<Root>\<Versions>\<Channels>\Cache|*|RECURSE
FileKeyBase=<Root>\<Versions>|app-<Versions>.log
```

**Output**

```ini
[Hazel Vault *]
LangSecRef=3021
DetectFile=%AppData%\Hazel Vault
FileKey1=%AppData%\Hazel Vault\1.0\Stable\Cache|*|RECURSE
FileKey2=%AppData%\Hazel Vault\1.0\Beta\Cache|*|RECURSE
FileKey3=%AppData%\Hazel Vault\2.0\Stable\Cache|*|RECURSE
FileKey4=%AppData%\Hazel Vault\2.0\Beta\Cache|*|RECURSE
FileKey5=%AppData%\Hazel Vault\1.0|app-1.0.log
FileKey6=%AppData%\Hazel Vault\2.0|app-2.0.log
```

**Explanation**

- The first `FileKeyBase` references two *different* multi-value variables, so they cross-multiply: 2 × 2 = 4 keys (`<Root>` is single-valued and just abbreviates)
- The second `FileKeyBase` references `<Versions>` *twice*, so the two uses co-vary: 2 keys, not 4. Both take the same value in each output key, so there's no `\1.0|app-2.0.log`
- A repeated *declared* variable always co-varies. The same *inline* list repeated in one key does too, but logs an advisory (see [Inline Lists](readme.md#inline-lists))

---

### Example 5: Inline lists

**Context**

7-Zip keeps an archive history value under two sibling registry keys. A named variable for a two-item list we use once is more typing than it's worth.

**Intent**

We want one `RegKey` per sibling key from a single line, and we'll take detection for free while we're at it.

**Files**

###### **Source (`#.ini`)**

```ini
[7-Zip *]
LangSecRef=3024
Root=HKCU\Software\7-Zip
RegKeyBase=<Root>\<Compression,Extraction>|ArcHistory
```

**Output**

```ini
[7-Zip *]
LangSecRef=3024
Detect=HKCU\Software\7-Zip
RegKey1=HKCU\Software\7-Zip\Compression|ArcHistory
RegKey2=HKCU\Software\7-Zip\Extraction|ArcHistory
```

**Explanation**

- `<Compression,Extraction>` has a comma in it, so it's an inline list: two values, no declaration needed
- `Root=` starts with a registry hive here, so the inferred detection is a `Detect`, not a `DetectFile`
- Writing the *same* inline list twice in one key makes the two copies co-vary and logs an advisory. Declare a named variable when you want to say outright whether they co-vary

---

### Example 6: Nested references

**Context**

ArcGIS ArcMap's registry layout changed across versions: the oldest releases used a bare `ArcMap` key, later ones `Desktop<version>\ArcMap`. The set of key names is itself version-derived.

**Intent**

We want the version list declared once, the derived key-name list built from it, and every dependent key fanned across the result.

**Files**

###### **Source (`A.ini`)**

```ini
[ArcGIS ArcMap *]
Versions=10.0,10.1
ArcMapKey=ArcMap,Desktop<Versions>\ArcMap

LangSecRef=3021
Detect=HKCU\Software\ESRI\<ArcMapKey>
RegKeyBase=HKCU\Software\ESRI\<ArcMapKey>\Recent File List
```

**Output**

```ini
[ArcGIS ArcMap *]
LangSecRef=3021
Detect1=HKCU\Software\ESRI\ArcMap
Detect2=HKCU\Software\ESRI\Desktop10.0\ArcMap
Detect3=HKCU\Software\ESRI\Desktop10.1\ArcMap
RegKey1=HKCU\Software\ESRI\ArcMap\Recent File List
RegKey2=HKCU\Software\ESRI\Desktop10.0\ArcMap\Recent File List
RegKey3=HKCU\Software\ESRI\Desktop10.1\ArcMap\Recent File List
```

**Explanation**

- `ArcMapKey`'s value list contains `<Versions>`, so it resolves first: the 2-value `Versions` expands the second list item, yielding the 3-value list `(ArcMap, Desktop10.0\ArcMap, Desktop10.1\ArcMap)`
- `Detect` and `RegKeyBase` then fan out across those three plain values
- A cycle (`A` referencing `B` referencing `A`) would leave both variables as authored, with a warning naming the participants

---

### Example 7: Literal angle brackets in registry keys

**Context**

Unlike file paths, registry key and value names may legitimately contain `<` and `>`. Some applications store values with names like `<default>`.

**Intent**

We want to clean a registry value named `<default>`, and to see why EntryBuilder lets it through instead of treating it as a broken variable reference.

**Files**

###### **Source (`H.ini`)**

```ini
[Hazel's Silly Little System Utility *]
Root=HKCU\Software\Hazel\HSLSU

LangSecRef=3024
RegKeyBase=<Root>\Recent|<default>
```

**Output**

```ini
[Hazel's Silly Little System Utility *]
LangSecRef=3024
Detect=HKCU\Software\Hazel\HSLSU
RegKey1=HKCU\Software\Hazel\HSLSU\Recent|<default>
```

**Explanation**

- `<Root>` matches a declared variable and expands; `<default>` matches nothing
- Because `RegKey` is a registry-domain key, the undeclared token is emitted **as literal text**, with a log-only advisory: `Undeclared variable <default> in registry-domain key [Hazel's Silly Little System Utility *].RegKey; passed through as literal`
- The same token in a `FileKeyBase` would instead **drop the key** with a menu-visible warning, since `<` / `>` can never appear in a Windows path

**Notes**

This is why the unreferenced-variable backstop matters: if you declare `Versions=` but typo every reference as `<Versoin>`, the registry domain would silently emit the literal. The `Variable 'Versions=' declared ... but never referenced` warning is what surfaces the mistake.

---

## Scaffold Families

###### Note: Unlike the earlier examples, the scaffold examples show `entrybuilder.ini` exactly as written, after WinappDebug's normalization pass, so FileKeys are alphabetized and same-path keys are merged. Only the file header is omitted.

### Example 8: Default WebView scaffolds

**Context**

Every app that embeds Chromium leaves the same cache and telemetry files behind, and the shared catalog holds those patterns so no entry has to. We'll use Discord's data folder to show the WebView catalog. Discord is really an Electron app, so a real entry would use `ElectronRoot=` instead, as in [Example 11](#example-11-an-electron-app-with-an-updater-cache).

**Intent**

We want baseline cache and telemetry cleaning for Discord's Chromium data folder from a four-line definition.

**Files**

###### **Source (`D.ini`)**

```ini
[Discord *]
WebViewRoot=%AppData%\discord
LangSecRef=3023
DetectFile=%AppData%\discord
```

**Command**

```
winapp2ool -entrybuilder -1d ..\..\Assembler\Sources\EntryBuilder -3d ..\..\Assembler\Sources\Scaffolds
```

###### Note: `-3d` points file slot 3 (the scaffold directory) at the shared scaffolds folder, from which every family's catalog is loaded

**Output**

```ini
[Discord *]
LangSecRef=3023
DetectFile=%AppData%\discord
FileKey1=%AppData%\discord|*.log;*_shutdown_ms.txt;Breadcrumbs;BrowsingTopics*;Last Browser;Last Version
FileKey2=%AppData%\discord|*.pma;LOG;LOG.old;*-journal|RECURSE
FileKey3=%AppData%\discord\*|*.ldb;CURRENT;LOCK;MANIFEST-*;ServerCertificate;*.log
FileKey4=%AppData%\discord\*\*Cache*|*|REMOVESELF
FileKey5=%AppData%\discord\*\blob_storage|*|REMOVESELF
; ... FileKey6 - FileKey20: the remaining profile-scoped Caches and Telemetry keys ...
FileKey21=%AppData%\discord\*BrowserMetrics|*|REMOVESELF
FileKey22=%AppData%\discord\*Cache*|*|REMOVESELF
FileKey23=%AppData%\discord\Crashpad|*|REMOVESELF
FileKey24=%AppData%\discord\Local Traces|*|REMOVESELF
FileKey25=%AppData%\discord\Optimization*|*|REMOVESELF
FileKey26=%AppData%\discord\OriginTrials|*|REMOVESELF
```

**Explanation**

- There's no `WebViewScaffolds=` key, so the default set applies: `Caches` and `Telemetry`
- Every `FileKeyBase=` template in each selected scaffold is copied in with `<WebViewRoot>` replaced by `%AppData%\discord`
- The lint pass merged templates that share a path: FileKey2 and FileKey3 each combine a Caches pattern list with a Telemetry one
- The keys, and their count (26 here), follow the catalog. When the catalog gains a pattern, every entry that selects the scaffold gets it on the next build without a source edit

**Notes**

A second `WebViewRoot=` line would double the scaffold output: one full set per root.

---

### Example 9: Everything except All with exclusions

**Context**

Adobe Photoshop's embedded WebView draws plugin UI. Most of what it leaves behind is safe to remove, but clearing its web storage and saved logins would lose plugin state people care about.

**Intent**

We want every current catalog scaffold **except** `WebStorage` and `LoginData`, plus one hand-written FileKey for Photoshop's own logs.

**Files**

###### **Source (`A.ini`)**

```ini
[Adobe Photoshop *]
WebViewRoot=%AppData%\Adobe\UXP\PluginsStorage\PHSP\26\Shared\EBWebView
WebViewScaffolds=All
ExcludeWebViewScaffolds=WebStorage,LoginData
LangSecRef=3023

DetectFile=%AppData%\Adobe\Adobe Photoshop *
FileKey=%AppData%\Adobe\Adobe Photoshop *\Logs|*
```

**Command**

```
winapp2ool -entrybuilder -1d ..\..\Assembler\Sources\EntryBuilder -3d ..\..\Assembler\Sources\Scaffolds
```

**Output**

```ini
[Adobe Photoshop *]
LangSecRef=3023
DetectFile=%AppData%\Adobe\Adobe Photoshop *
FileKey1=%AppData%\Adobe\Adobe Photoshop *\Logs|*
FileKey2=%AppData%\Adobe\UXP\PluginsStorage\PHSP\26\Shared\EBWebView|*.log;*_shutdown_ms.txt;Breadcrumbs;BrowsingTopics*;Last Browser;Last Version;*first_party_sets*
FileKey3=%AppData%\Adobe\UXP\PluginsStorage\PHSP\26\Shared\EBWebView|*.pma;LOG;LOG.old;*-journal|RECURSE
FileKey4=%AppData%\Adobe\UXP\PluginsStorage\PHSP\26\Shared\EBWebView\*|*.ldb;CURRENT;LOCK;MANIFEST-*;ServerCertificate;*.log;*Web Data;BrowsingTopics*;Conversions*;InterestGroups;MediaDeviceSalts;PrivateAggregation*;SharedStorage*;DIPS*;favicons*;History*;Network Action Predictor*;Top Sites*;Visited Links*;PreferredApps
FileKey5=%AppData%\Adobe\UXP\PluginsStorage\PHSP\26\Shared\EBWebView\*\*Cache*|*|REMOVESELF
FileKey6=%AppData%\Adobe\UXP\PluginsStorage\PHSP\26\Shared\EBWebView\*\AutoFill*|*|REMOVESELF
; ... FileKey7 - FileKey14 ...
FileKey15=%AppData%\Adobe\UXP\PluginsStorage\PHSP\26\Shared\EBWebView\*\Extension State|*|REMOVESELF
; ... FileKey16 - FileKey19 ...
FileKey20=%AppData%\Adobe\UXP\PluginsStorage\PHSP\26\Shared\EBWebView\*\Network|Cookies*;Device Bound Sessions*;Network Persistent State*;Reporting and NEL*;SCT Auditing Pending Reports*;TransportSecurity*;Trust Tokens*
; ... FileKey21 - FileKey26 ...
FileKey27=%AppData%\Adobe\UXP\PluginsStorage\PHSP\26\Shared\EBWebView\*\Sessions|*|REMOVESELF
; ... FileKey28 - FileKey55: the rest of the profile-scoped keys, then the root-level folders ...
FileKey56=%AppData%\Adobe\UXP\PluginsStorage\PHSP\26\Shared\EBWebView\ZxcvbnData|*|REMOVESELF
```

**Explanation**

- `All` expands to every scaffold outside the legacy tier, and the log records it: `WebViewScaffolds=All in [Adobe Photoshop *]; expanded to 17 scaffold(s) from catalog, 8 legacy withheld`
- The two exclusions subtract from that expansion, leaving 15 scaffolds that produce 55 FileKeys after the lint pass merges same-path templates. That includes the host-risk sets `All` deliberately opts into: `WebCookies` (the `Cookies*;Device Bound Sessions*` patterns merged into FileKey20), `WebHistory` (the `History*` through `Visited Links*` patterns merged into FileKey4), and `WebSession` (FileKey15 and FileKey27)
- `All` with `Exclude...` is the intended way to write this and doesn't warn. An explicit list other than `All` with exclusions does
- The hand-written `FileKey=` goes through the same normalization as the scaffold keys, so it sorts to FileKey1
- A legacy scaffold would have to be named to be included, e.g. `WebViewScaffolds=All,LegacyTelemetry`

---

### Example 10: QtWebEngine across several profiles

**Context**

calibre embeds QtWebEngine, and unusually it runs three profiles: the standard `Default`, plus `OffTheRecord` and `viewer-lookup` for the ebook viewer. Each profile is just another root, and an inline list can declare all three at once.

**Intent**

We want the default QtWebEngine scaffolds applied to every profile, plus calibre's own cache directory, without repeating the long path three times.

**Files**

###### **Source (`C.ini`)**

```ini
[Calibre *]
QtRoot=%LocalAppData%\calibre-ebook.com\calibre\QtWebEngine
QtWebEngineRoot=<QtRoot>\<Default,OffTheRecord,viewer-lookup>
LangSecRef=3023
DetectFile=%LocalAppData%\calibre-ebook.com
FileKeyBase=%LocalAppData%\calibre-cache|*|RECURSE
```

**Command**

```
winapp2ool -entrybuilder -1d ..\..\Assembler\Sources\EntryBuilder -3d ..\..\Assembler\Sources\Scaffolds
```

###### Note: `-3d` points file slot 3 (the scaffold directory) at the shared scaffolds folder; `qtwebengine.ini` is discovered inside it by section header

**Output**

```ini
[Calibre *]
LangSecRef=3023
DetectFile=%LocalAppData%\calibre-ebook.com
FileKey1=%LocalAppData%\calibre-cache|*|RECURSE
FileKey2=%LocalAppData%\calibre-ebook.com\calibre\QtWebEngine\Default|*-journal;*.log;*.old;LOG;LOG.old;Network Persistent State;Origin Bound Certs|RECURSE
FileKey3=%LocalAppData%\calibre-ebook.com\calibre\QtWebEngine\Default|QuotaManager*;Visited Links
FileKey4=%LocalAppData%\calibre-ebook.com\calibre\QtWebEngine\Default\*Cache*|*|REMOVESELF
FileKey5=%LocalAppData%\calibre-ebook.com\calibre\QtWebEngine\Default\blob_storage|*|REMOVESELF
FileKey6=%LocalAppData%\calibre-ebook.com\calibre\QtWebEngine\Default\Platform Notifications|*|REMOVESELF
FileKey7=%LocalAppData%\calibre-ebook.com\calibre\QtWebEngine\Default\Service Worker\*Cache*|*|REMOVESELF
FileKey8=%LocalAppData%\calibre-ebook.com\calibre\QtWebEngine\Default\Shared Dictionary\cache|*|REMOVESELF
FileKey9=%LocalAppData%\calibre-ebook.com\calibre\QtWebEngine\OffTheRecord|*-journal;*.log;*.old;LOG;LOG.old;Network Persistent State;Origin Bound Certs|RECURSE
FileKey10=%LocalAppData%\calibre-ebook.com\calibre\QtWebEngine\OffTheRecord|QuotaManager*;Visited Links
FileKey11=%LocalAppData%\calibre-ebook.com\calibre\QtWebEngine\OffTheRecord\*Cache*|*|REMOVESELF
FileKey12=%LocalAppData%\calibre-ebook.com\calibre\QtWebEngine\OffTheRecord\blob_storage|*|REMOVESELF
FileKey13=%LocalAppData%\calibre-ebook.com\calibre\QtWebEngine\OffTheRecord\Platform Notifications|*|REMOVESELF
FileKey14=%LocalAppData%\calibre-ebook.com\calibre\QtWebEngine\OffTheRecord\Service Worker\*Cache*|*|REMOVESELF
FileKey15=%LocalAppData%\calibre-ebook.com\calibre\QtWebEngine\OffTheRecord\Shared Dictionary\cache|*|REMOVESELF
FileKey16=%LocalAppData%\calibre-ebook.com\calibre\QtWebEngine\viewer-lookup|*-journal;*.log;*.old;LOG;LOG.old;Network Persistent State;Origin Bound Certs|RECURSE
FileKey17=%LocalAppData%\calibre-ebook.com\calibre\QtWebEngine\viewer-lookup|QuotaManager*;Visited Links
FileKey18=%LocalAppData%\calibre-ebook.com\calibre\QtWebEngine\viewer-lookup\*Cache*|*|REMOVESELF
FileKey19=%LocalAppData%\calibre-ebook.com\calibre\QtWebEngine\viewer-lookup\blob_storage|*|REMOVESELF
FileKey20=%LocalAppData%\calibre-ebook.com\calibre\QtWebEngine\viewer-lookup\Platform Notifications|*|REMOVESELF
FileKey21=%LocalAppData%\calibre-ebook.com\calibre\QtWebEngine\viewer-lookup\Service Worker\*Cache*|*|REMOVESELF
FileKey22=%LocalAppData%\calibre-ebook.com\calibre\QtWebEngine\viewer-lookup\Shared Dictionary\cache|*|REMOVESELF
```

**Explanation**

- One `QtWebEngineRoot=` line produced three roots. Root substitution happens before token expansion, so the `<Default,OffTheRecord,viewer-lookup>` inline list inside the root value fans out normally and `<QtRoot>` resolves in the same pass
- Each root got keys for all the default scaffolds (`Caches`, `StorageQuota`, `Telemetry`, `VisitedLinks`)
- FileKey3, 10 and 17 each hold two scaffolds' patterns, merged by WinappDebug's optimization pass
- The `<QtWebEngineCacheRoot>` template in `Caches` was dropped, since the entry declares no `QtWebEngineCacheRoot=`; calibre's own cache is covered by the hand-written FileKey1 instead
- `Favicons`, `PrivacySandbox`, `Security` and the host-risk `WebCookies`, `WebHistory`, `WebSession`, `WebStorage` would require `QtWebEngineScaffolds=All` or an explicit list. The legacy `LegacyTelemetry` and `LegacyWebStorage` must be named even beside `All`

---

### Example 11: An Electron app with an updater cache

**Context**

Signal Desktop is a typical Electron app: its `userData` folder is `%AppData%\Signal`, and electron-updater keeps downloaded installers in a sibling folder under `%LocalAppData%`. The hand-written form of this entry was seven FileKeys of Chromium boilerplate.

**Intent**

We want the default Electron scaffolds for both the app data and the updater cache, plus one target of Signal's own that the catalog doesn't cover.

**Files**

###### **Source (`S.ini`)**

```ini
[Signal Messenger *]
LangSecRef=3022
DetectFile=%AppData%\Signal
ElectronRoot=%AppData%\Signal
ElectronUpdaterRoot=%LocalAppData%\Signal-Desktop-Updater
FileKeyBase=<ElectronRoot>\Temp|*
```

**Command**

```
winapp2ool -entrybuilder -1d ..\..\Assembler\Sources\EntryBuilder -3d ..\..\Assembler\Sources\Scaffolds
```

###### Note: `-3d` points file slot 3 (the scaffold directory) at the shared scaffolds folder; `electron.ini` is discovered inside it by section header

**Output**

```ini
[Signal Messenger *]
LangSecRef=3022
DetectFile=%AppData%\Signal
FileKey1=%AppData%\Signal|*.log;Network Persistent State*;Origin Bound Certs;Reporting and NEL*;Visited Links*;QuotaManager*
FileKey2=%AppData%\Signal|*.old;LOG;LOG.old;*-journal|RECURSE
FileKey3=%AppData%\Signal\*\*Logs|*|RECURSE
FileKey4=%AppData%\Signal\*Cache*|*|REMOVESELF
FileKey5=%AppData%\Signal\*Logs|*|RECURSE
FileKey6=%AppData%\Signal\blob_storage|*|REMOVESELF
FileKey7=%AppData%\Signal\Crashpad|*|REMOVESELF
FileKey8=%AppData%\Signal\Network|Network Persistent State*;Reporting and NEL*;SCT Auditing Pending Reports*
FileKey9=%AppData%\Signal\sentry|*|REMOVESELF
FileKey10=%AppData%\Signal\Service Worker\*Cache*|*|REMOVESELF
FileKey11=%AppData%\Signal\Shared Dictionary\cache|*|REMOVESELF
FileKey12=%AppData%\Signal\shared_proto_db|*|REMOVESELF
FileKey13=%AppData%\Signal\Temp|*
FileKey14=%AppData%\Signal\VideoDecodeStats|*|REMOVESELF
FileKey15=%AppData%\Signal\webrtc_event_logs|*|REMOVESELF
FileKey16=%AppData%\Signal\WebStorage|QuotaManager*
FileKey17=%LocalAppData%\Signal-Desktop-Updater|*.7z;*.exe
FileKey18=%LocalAppData%\Signal-Desktop-Updater\pending|*|REMOVESELF
```

**Explanation**

- Five source lines produced eighteen FileKeys, from the default set `AppLogs`, `Caches`, `StorageQuota`, `Telemetry`, `UpdaterCache`
- FileKey1 is the merged product of several scaffolds' root-level pattern lists
- FileKey17 and 18 came from `<ElectronUpdaterRoot>`. Without `ElectronUpdaterRoot=`, those two templates would have been dropped
- The placeholder in `FileKeyBase=` is `<ElectronRoot>`. EntryBuilder refuses the old `%ElectronRoot%` spelling and fails the run

---

### Example 12: An Electron app with session partitions

**Context**

Notion uses session partitions, so alongside its own profile it keeps a complete Chromium profile for each partition under `%AppData%\Notion\partitions\`, and every one of them needs cleaning.

**Intent**

We want cache and log cleaning for the main profile and every partition.

**Files**

###### **Source (`N.ini`)**

```ini
[Notion *]
Root=%AppData%\Notion
LangSecRef=3021
ElectronRoot1=<Root>
ElectronRoot2=<Root>\partitions\*
ElectronUpdaterRoot=%LocalAppData%\notion-updater
ElectronScaffolds=Caches,AppLogs,UpdaterCache
```

**Output**

```ini
[Notion *]
LangSecRef=3021
DetectFile=%AppData%\Notion
FileKey1=%AppData%\Notion|*.log
FileKey2=%AppData%\Notion|*-journal|RECURSE
FileKey3=%AppData%\Notion\*\*Logs|*|RECURSE
FileKey4=%AppData%\Notion\*Cache*|*|REMOVESELF
FileKey5=%AppData%\Notion\*Logs|*|RECURSE
FileKey6=%AppData%\Notion\blob_storage|*|REMOVESELF
FileKey7=%AppData%\Notion\partitions\*|*.log
FileKey8=%AppData%\Notion\partitions\*|*-journal|RECURSE
FileKey9=%AppData%\Notion\partitions\*\*\*Logs|*|RECURSE
FileKey10=%AppData%\Notion\partitions\*\*Cache*|*|REMOVESELF
FileKey11=%AppData%\Notion\partitions\*\*Logs|*|RECURSE
FileKey12=%AppData%\Notion\partitions\*\blob_storage|*|REMOVESELF
FileKey13=%AppData%\Notion\partitions\*\Service Worker\*Cache*|*|REMOVESELF
FileKey14=%AppData%\Notion\partitions\*\Shared Dictionary\cache|*|REMOVESELF
FileKey15=%AppData%\Notion\partitions\*\shared_proto_db|*|REMOVESELF
FileKey16=%AppData%\Notion\Service Worker\*Cache*|*|REMOVESELF
FileKey17=%AppData%\Notion\Shared Dictionary\cache|*|REMOVESELF
FileKey18=%AppData%\Notion\shared_proto_db|*|REMOVESELF
FileKey19=%LocalAppData%\notion-updater|*.7z;*.exe
FileKey20=%LocalAppData%\notion-updater\pending|*|REMOVESELF
```

**Explanation**

- Each of the two numbered root keys got its own copy of the selected scaffolds
- `DetectFile=%AppData%\Notion` was inferred from the reserved `Root` variable (see [The Reserved Root Variable](readme.md#the-reserved-root-variable)); `ElectronRoot1=<Root>` reuses it so the path appears once in the source
- `UpdaterCache` had to be named explicitly because `ElectronScaffolds=` replaces the default set rather than adding to it

---

## Housekeeping

### Example 13: Skipping a defunct entry

**Context**

An application is discontinued, but deleting its definition would throw away the research in it, and the app might come back or get forked.

**Intent**

We want the entry omitted from the output while keeping its definition in source.

**Files**

###### **Source (`D.ini`)**

```ini
[Discontinued App *]
Skip=
WebViewRoot=%AppData%\DiscontinuedApp
LangSecRef=3023
DetectFile=%AppData%\DiscontinuedApp
```

**Output**

The entry does not appear in `entrybuilder.ini`. Re-enabling it later is a one-line deletion.

**Explanation**

- `Skip=` leaves the entry out whatever its value, so even `Skip=False` skips
- The run statistics count skipped entries, so a skipped entry still shows up in the totals
