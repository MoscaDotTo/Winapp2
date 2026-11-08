# UWPBuilder Examples

Worked examples for [UWPBuilder](readme.md). The [Common Uses](readme.md#common-uses) table in the readme lists the commands; these show what each one does, with real input and output.

# Table of Contents

1. [Example 1: A Minimal Entry](#example-1-a-minimal-entry)
2. [Example 2: App-Specific FileKeys](#example-2-app-specific-filekeys)
3. [Example 3: Hybrid UWP + win32 Applications](#example-3-hybrid-uwp--win32-applications)
4. [Example 4: Multi-Package Applications](#example-4-multi-package-applications)
5. [Example 5: WebView Scaffolds (Default Selection)](#example-5-webview-scaffolds-default-selection)
6. [Example 6: WebView Scaffolds (All With Exclusions)](#example-6-webview-scaffolds-all-with-exclusions)
7. [Example 7: Secondary Entries With SkipUWPFileKeys](#example-7-secondary-entries-with-skipuwpfilekeys)
8. [Example 8: Validation Warnings and Skips](#example-8-validation-warnings-and-skips)

---

All examples below were generated with real runs of UWPBuilder against the scaffold template and catalogs shipped in `Assembler/`. The source files shown are AppInfo sections; the command for every example is the standard invocation:

```
winapp2ool -uwpbuilder -1d \UWP -2f uwp.ini -3d \Scaffolds
```

###### Note: The generated header comment block (see [Output Formatting](readme.md#output-formatting)) is omitted from the outputs below for brevity.

### Example 1: A Minimal Entry

**Context**

The vast majority of UWP applications store nothing cleanable outside the standard package layout. For these apps, the scaffold baseline is the entire entry.

**Intent**

We want a complete winapp2.ini entry for ASUS Quick Launch by declaring only its package and category.

**Files**

###### **AppInfo section (`AppInfo\A.ini`)**
```ini
[ASUS Quick Launch *]
Package=B9ECED6F.ASUSQuickLaunch_qmba6cd70vzyy
LangSecRef=3024
```

**Output**

```ini
[ASUS Quick Launch *]
LangSecRef=3024
DetectFile=%LocalAppData%\Packages\B9ECED6F.ASUSQuickLaunch_qmba6cd70vzyy
FileKey1=%LocalAppData%\Packages\B9ECED6F.ASUSQuickLaunch_qmba6cd70vzyy\AC|*|RECURSE
FileKey2=%LocalAppData%\Packages\B9ECED6F.ASUSQuickLaunch_qmba6cd70vzyy\Settings|*.log*
FileKey3=%LocalAppData%\Packages\B9ECED6F.ASUSQuickLaunch_qmba6cd70vzyy\SystemAppData\Helium|*.log*
FileKey4=%LocalAppData%\Packages\B9ECED6F.ASUSQuickLaunch_qmba6cd70vzyy\TempState|*|REMOVESELF
```

**Explanation**
- `DetectFile` is generated from the scaffold's `DetectFileBase=<Package>` template
- The four FileKeys are the scaffold's `FileKeyBase=` templates with `<Package>` expanded
- The category key passes through as written
- Nothing in this entry was written by hand except the package name and category

---

### Example 2: App-Specific FileKeys

**Context**

Apple Devices keeps logs in several locations under its package that the baseline scaffold doesn't cover.

**Intent**

We want to extend the baseline with three additional package-relative FileKeys.

**Files**

###### **AppInfo section (`AppInfo\A.ini`)**
```ini
[Apple Devices *]
Package=AppleInc.AppleDevices_nzyj5cx40ttqa
LangSecRef=3024
FileKeyBase=<Package>\LocalCache\Local\Logs|*
FileKeyBase=<Package>\LocalCache\Roaming\Apple Computer\Logs|*
FileKeyBase=<Package>\LocalState\Logs|*
```

**Output**

```ini
[Apple Devices *]
LangSecRef=3024
DetectFile=%LocalAppData%\Packages\AppleInc.AppleDevices_nzyj5cx40ttqa
FileKey1=%LocalAppData%\Packages\AppleInc.AppleDevices_nzyj5cx40ttqa\AC|*|RECURSE
FileKey2=%LocalAppData%\Packages\AppleInc.AppleDevices_nzyj5cx40ttqa\LocalCache\Local\Logs|*
FileKey3=%LocalAppData%\Packages\AppleInc.AppleDevices_nzyj5cx40ttqa\LocalCache\Roaming\Apple Computer\Logs|*
FileKey4=%LocalAppData%\Packages\AppleInc.AppleDevices_nzyj5cx40ttqa\LocalState\Logs|*
FileKey5=%LocalAppData%\Packages\AppleInc.AppleDevices_nzyj5cx40ttqa\Settings|*.log*
FileKey6=%LocalAppData%\Packages\AppleInc.AppleDevices_nzyj5cx40ttqa\SystemAppData\Helium|*.log*
FileKey7=%LocalAppData%\Packages\AppleInc.AppleDevices_nzyj5cx40ttqa\TempState|*|REMOVESELF
```

**Explanation**
- The three app-specific keys are interleaved with the scaffold baseline: the lint pass sorts FileKeys alphabetically by value, so source ordering is lost
- All seven keys were renumbered sequentially after sorting

---

### Example 3: Hybrid UWP + win32 Applications

**Context**

Amazon Music ships as both a UWP package and a win32 installation. A single entry can clean both.

**Intent**

We want package-relative keys (`FileKeyBase=` with `<Package>`) for the UWP install, verbatim win32 keys (`FileKey=`), and a registry detection for the win32 version alongside the generated package detection.

**Files**

###### **AppInfo section (`AppInfo\A.ini`)**
```ini
[Amazon Music *]
Detect=HKCU\Software\Amazon\Amazon Music
LangSecRef=3023
Package=AmazonMobileLLC.AmazonMusic_kc6t79cpj4tp0
FileKeyBase=<Package>\LocalCache\Local\Amazon Music\Data\App Cache\*Cache*|RECURSE
FileKeyBase=<Package>\LocalCache\Local\Amazon Music\Data\Artwork Cache|*.jpg;*.png|RECURSE
FileKeyBase=<Package>\LocalCache\Local\Amazon Music\Data\Hammer Cache|*
FileKeyBase=<Package>\LocalCache\Local\Amazon Music\Data\S*Cache|*
FileKeyBase=<Package>\LocalCache\Local\Amazon Music\Logs|*
FileKeyBase=<Package>\LocalCache\Local\Amazon Music\User Data|*.pma
FileKeyBase=<Package>\LocalCache\Local\Amazon Music\User Data\Crashpad\Reports|*
FileKey=%LocalAppData%\Amazon Music\Crash Dumps|*
FileKey=%LocalAppData%\Amazon Music\data\App Cache|*-journal;ChromeDWriteFontCache;data_*;f_*;index
FileKey=%LocalAppData%\Amazon Music\data\App Cache\*Cache|*|RECURSE
FileKey=%LocalAppData%\Amazon Music\data\Artwork Cache|*.jpg;*.png|RECURSE
FileKey=%LocalAppData%\Amazon Music\data\Hammer Cache|*
FileKey=%LocalAppData%\Amazon Music\data\S*Cache|*
FileKey=%LocalAppData%\Amazon Music\Logs|*
FileKey=%LocalAppData%\Amazon Music\User Data|*.pma
FileKey=%LocalAppData%\Amazon Music\User Data\Crashpad\Reports|*
FileKey=%UserProfile%\.amu\updates|*.exe
FileKey=%UserProfile%\Documents\Amazon Music Importer\Logs|*|RECURSE
```

**Output**

```ini
[Amazon Music *]
LangSecRef=3023
Detect=HKCU\Software\Amazon\Amazon Music
DetectFile=%LocalAppData%\Packages\AmazonMobileLLC.AmazonMusic_kc6t79cpj4tp0
FileKey1=%LocalAppData%\Amazon Music\Crash Dumps|*
FileKey2=%LocalAppData%\Amazon Music\data\App Cache|*-journal;ChromeDWriteFontCache;data_*;f_*;index
FileKey3=%LocalAppData%\Amazon Music\data\App Cache\*Cache|*|RECURSE
FileKey4=%LocalAppData%\Amazon Music\data\Artwork Cache|*.jpg;*.png|RECURSE
FileKey5=%LocalAppData%\Amazon Music\data\Hammer Cache|*
FileKey6=%LocalAppData%\Amazon Music\data\S*Cache|*
FileKey7=%LocalAppData%\Amazon Music\Logs|*
FileKey8=%LocalAppData%\Amazon Music\User Data|*.pma
FileKey9=%LocalAppData%\Amazon Music\User Data\Crashpad\Reports|*
FileKey10=%LocalAppData%\Packages\AmazonMobileLLC.AmazonMusic_kc6t79cpj4tp0\AC|*|RECURSE
FileKey11=%LocalAppData%\Packages\AmazonMobileLLC.AmazonMusic_kc6t79cpj4tp0\LocalCache\Local\Amazon Music\Data\App Cache\*Cache*|RECURSE
FileKey12=%LocalAppData%\Packages\AmazonMobileLLC.AmazonMusic_kc6t79cpj4tp0\LocalCache\Local\Amazon Music\Data\Artwork Cache|*.jpg;*.png|RECURSE
FileKey13=%LocalAppData%\Packages\AmazonMobileLLC.AmazonMusic_kc6t79cpj4tp0\LocalCache\Local\Amazon Music\Data\Hammer Cache|*
FileKey14=%LocalAppData%\Packages\AmazonMobileLLC.AmazonMusic_kc6t79cpj4tp0\LocalCache\Local\Amazon Music\Data\S*Cache|*
FileKey15=%LocalAppData%\Packages\AmazonMobileLLC.AmazonMusic_kc6t79cpj4tp0\LocalCache\Local\Amazon Music\Logs|*
FileKey16=%LocalAppData%\Packages\AmazonMobileLLC.AmazonMusic_kc6t79cpj4tp0\LocalCache\Local\Amazon Music\User Data|*.pma
FileKey17=%LocalAppData%\Packages\AmazonMobileLLC.AmazonMusic_kc6t79cpj4tp0\LocalCache\Local\Amazon Music\User Data\Crashpad\Reports|*
FileKey18=%LocalAppData%\Packages\AmazonMobileLLC.AmazonMusic_kc6t79cpj4tp0\Settings|*.log*
FileKey19=%LocalAppData%\Packages\AmazonMobileLLC.AmazonMusic_kc6t79cpj4tp0\SystemAppData\Helium|*.log*
FileKey20=%LocalAppData%\Packages\AmazonMobileLLC.AmazonMusic_kc6t79cpj4tp0\TempState|*|REMOVESELF
FileKey21=%UserProfile%\.amu\updates|*.exe
FileKey22=%UserProfile%\Documents\Amazon Music Importer\Logs|*|RECURSE
```

**Explanation**
- `Detect=` passes through verbatim, providing win32 detection; the package `DetectFile` is generated alongside it
- The win32 `FileKey=` values contain no `<Package>` variable and pass through verbatim
- One entry covers both installations

---

### Example 4: Multi-Package Applications

**Context**

Cortana has shipped under two different package identities. Some data locations exist under both packages, others only under the newer one.

**Intent**

We want one entry covering both packages: shared keys should expand for every package (`<Package>`), package-specific keys should target only one (`<Package[2]>`).

**Files**

###### **AppInfo section (`AppInfo\C.ini`)**
```ini
[Cortana *]
Package1=Microsoft.549981C3F5F10_8wekyb3d8bbwe
Package2=Microsoft.Cortana_8wekyb3d8bbwe
LangSecRef=3025
FileKeyBase=<Package>\LocalCache|*|RECURSE
FileKeyBase=<Package[2]>\LocalState|*|RECURSE
ExcludeKey1=FILE|<Package[2]>\LocalState\DeviceSearchCache\|SettingsCache.txt
```

**Output**

```ini
[Cortana *]
LangSecRef=3025
DetectFile1=%LocalAppData%\Packages\Microsoft.549981C3F5F10_8wekyb3d8bbwe
DetectFile2=%LocalAppData%\Packages\Microsoft.Cortana_8wekyb3d8bbwe
FileKey1=%LocalAppData%\Packages\Microsoft.549981C3F5F10_8wekyb3d8bbwe\AC|*|RECURSE
FileKey2=%LocalAppData%\Packages\Microsoft.549981C3F5F10_8wekyb3d8bbwe\LocalCache|*|RECURSE
FileKey3=%LocalAppData%\Packages\Microsoft.549981C3F5F10_8wekyb3d8bbwe\Settings|*.log*
FileKey4=%LocalAppData%\Packages\Microsoft.549981C3F5F10_8wekyb3d8bbwe\SystemAppData\Helium|*.log*
FileKey5=%LocalAppData%\Packages\Microsoft.549981C3F5F10_8wekyb3d8bbwe\TempState|*|REMOVESELF
FileKey6=%LocalAppData%\Packages\Microsoft.Cortana_8wekyb3d8bbwe\AC|*|RECURSE
FileKey7=%LocalAppData%\Packages\Microsoft.Cortana_8wekyb3d8bbwe\LocalCache|*|RECURSE
FileKey8=%LocalAppData%\Packages\Microsoft.Cortana_8wekyb3d8bbwe\LocalState|*|RECURSE
FileKey9=%LocalAppData%\Packages\Microsoft.Cortana_8wekyb3d8bbwe\Settings|*.log*
FileKey10=%LocalAppData%\Packages\Microsoft.Cortana_8wekyb3d8bbwe\SystemAppData\Helium|*.log*
FileKey11=%LocalAppData%\Packages\Microsoft.Cortana_8wekyb3d8bbwe\TempState|*|REMOVESELF
ExcludeKey1=FILE|%LocalAppData%\Packages\Microsoft.Cortana_8wekyb3d8bbwe\LocalState\DeviceSearchCache\|SettingsCache.txt
```

**Explanation**
- One `DetectFile` was generated per package from the scaffold's `DetectFileBase=<Package>`
- The scaffold baseline and the `<Package>` FileKey expanded once per package: keys 1–5 for `Package1`, keys 6–7 and 9–11 for `Package2`
- The `<Package[2]>` keys (FileKey8's `LocalState`, and the ExcludeKey) expanded exactly once, targeting only the second package

---

### Example 5: WebView Scaffolds (Default Selection)

**Context**

GoldKey Acellus is an education app that embeds WebView2 and keeps its browser data under `LocalState\EBWebView`.

**Intent**

We want the entry to draw the default scaffold selection (Caches + Telemetry) from the shared catalog by declaring only the WebView root.

**Files**

###### **AppInfo section**
```ini
[GoldKey Acellus *]
Package=GoldKeyCorporation.Acellus_v2rn50m687qry
WebViewRoot=<Package>\LocalState\EBWebView
LangSecRef=3021
```

###### Note: The entry for this app shipped in the winapp2.ini project selects `WebViewScaffolds=All` with exclusions — that variant is [Example 6](#example-6-webview-scaffolds-all-with-exclusions). This example omits the selector keys to demonstrate the default behavior.

**Output**

```ini
[GoldKey Acellus *]
LangSecRef=3021
DetectFile=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry
FileKey1=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\AC|*|RECURSE
FileKey2=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView|*.log;*_shutdown_ms.txt;Breadcrumbs;BrowsingTopics*;Last Browser;Last Version
FileKey3=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView|*.pma;LOG;LOG.old;*-journal|RECURSE
FileKey4=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView\*|*.ldb;CURRENT;LOCK;MANIFEST-*;ServerCertificate;*.log
FileKey5=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView\*\*Cache*|*|REMOVESELF
FileKey6=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView\*\blob_storage|*|REMOVESELF
FileKey7=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView\*\BudgetDatabase|*|REMOVESELF
FileKey8=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView\*\Download Service|*|REMOVESELF
FileKey9=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView\*\Feature Engagement Tracker|*|REMOVESELF
FileKey10=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView\*\File System|*|REMOVESELF
FileKey11=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView\*\GCM Store|*|REMOVESELF
FileKey12=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView\*\JumpListIcons*|*|REMOVESELF
FileKey13=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView\*\Network|Network Persistent State*;Reporting and NEL*;SCT Auditing Pending Reports*
FileKey14=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView\*\optimization*|*|REMOVESELF
FileKey15=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView\*\PersistentOriginTrials|*|REMOVESELF
FileKey16=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView\*\Platform Notifications|*|REMOVESELF
FileKey17=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView\*\Service Worker|*|REMOVESELF
FileKey18=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView\*\Shared Dictionary\cache|*|REMOVESELF
FileKey19=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView\*\Site Characteristics Database|*|REMOVESELF
FileKey20=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView\*\VideoDecodeStats|*|REMOVESELF
FileKey21=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView\*\WebrtcVideoStats|*|REMOVESELF
FileKey22=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView\*BrowserMetrics|*|REMOVESELF
FileKey23=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView\*Cache*|*|REMOVESELF
FileKey24=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView\Crashpad|*|REMOVESELF
FileKey25=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView\Local Traces|*|REMOVESELF
FileKey26=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView\Optimization*|*|REMOVESELF
FileKey27=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView\OriginTrials|*|REMOVESELF
FileKey28=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\Settings|*.log*
FileKey29=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\SystemAppData\Helium|*.log*
FileKey30=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\TempState|*|REMOVESELF
```

**Explanation**
- Declaring `WebViewRoot=` opted the entry into scaffold emission
- With no `WebViewScaffolds=` key, the default Caches + Telemetry selection applied
- The catalog templates had `<WebViewRoot>` substituted with the package-expanded root path
- 26 WebView FileKeys were generated from a single declaration line
- The lint pass merged scaffold templates targeting the same folder (e.g. FileKey4 combines a Caches template and a Telemetry template both targeting `EBWebView\*`)

---

### Example 6: WebView Scaffolds (All With Exclusions)

**Context**

For a dedicated single-purpose app like Acellus, the browsing data inside the WebView (history, cookies, sessions) can be safely deleted, but wiping `WebStorage` or `LoginData` would log the user out or destroy locally-stored user data.

**Intent**

We want every scaffold outside the legacy tier *except* the two that remove primary user state. This is the variant of this entry actually shipped by the winapp2.ini project.

**Files**

###### **AppInfo section (`AppInfo\G.ini`)**
```ini
[GoldKey Acellus *]
Package=GoldKeyCorporation.Acellus_v2rn50m687qry
WebViewRoot=<Package>\LocalState\EBWebView
WebViewScaffolds=All
ExcludeWebViewScaffolds=WebStorage,LoginData
LangSecRef=3021
```

**Output**

The full output carries **59 FileKeys**; the excerpt below shows the shape, with the middle trimmed for brevity:

```ini
[GoldKey Acellus *]
LangSecRef=3021
DetectFile=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry
FileKey1=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\AC|*|RECURSE
FileKey2=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView|*.log;*_shutdown_ms.txt;Breadcrumbs;BrowsingTopics*;Last Browser;Last Version;*first_party_sets*
FileKey3=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView|*.pma;LOG;LOG.old;*-journal|RECURSE
FileKey4=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\LocalState\EBWebView\*|*.ldb;CURRENT;LOCK;MANIFEST-*;ServerCertificate;*.log;*Web Data;BrowsingTopics*;Conversions*;InterestGroups;MediaDeviceSalts;PrivateAggregation*;SharedStorage*;DIPS*;favicons*;History*;Network Action Predictor*;Top Sites*;Visited Links*;PreferredApps

; ... FileKey5 through FileKey56: the remaining Autofill, Autoplay, Caches, DownloadHistory, DRMData,
; PrivacySandbox, Security, Shopping, Sync Data, Telemetry, WebCookies and WebSession scaffold keys ...

FileKey57=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\Settings|*.log*
FileKey58=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\SystemAppData\Helium|*.log*
FileKey59=%LocalAppData%\Packages\GoldKeyCorporation.Acellus_v2rn50m687qry\TempState|*|REMOVESELF
```

**Explanation**
- `WebViewScaffolds=All` expanded to every scaffold outside the legacy tier, logged as `WebViewScaffolds=All in [GoldKey Acellus *]; expanded to 17 scaffold(s) from catalog, 8 legacy withheld`. `ExcludeWebViewScaffolds=WebStorage,LoginData` then removed those two from the selection
- Compare FileKey4 with Example 5's FileKey4: the same `EBWebView\*` folder now aggregates parameters from many more scaffolds, merged into one key by the lint pass
- No legacy scaffold was emitted. Adding one requires naming it: `WebViewScaffolds=All,LegacyTelemetry`

---

### Example 7: Secondary Entries With SkipUWPFileKeys

**Context**

Spotify's downloaded songs occupy significant disk space, but users won't necessarily expect a cleaner to remove their offline music. This warrants a *separate* entry so the two cleaning surfaces can be toggled independently, but a second entry for the same package would duplicate the scaffold baseline.

**Intent**

We want a primary `[Spotify *]` entry with full coverage, and a secondary `[Spotify - Downloaded Songs *]` entry that shares the package but carries only its own keys.

**Files**

###### **AppInfo sections (`AppInfo\S.ini`)**
```ini
[Spotify *]
Package=SpotifyAB.SpotifyMusic_zpdnekdrzrea0
LangSecRef=3023
DetectFile=%AppData%\Spotify
FileKeyBase=<Package>\LocalCache\Spotify|*|REMOVESELF
FileKeyBase=<Package>\LocalState\Users\*|ad-state-storage.bnk;context_player_state_restore;cached
FileKey=%AppData%\Spotify|*.log
FileKey=%AppData%\Spotify\users\*|ad-state-storage.bnk;frequency.pb;local-files.bnk;local_sync_state;watch-sources.bnk
FileKey=%LocalAppData%\Spotify|*|REMOVESELF
ExcludeKey1=PATH|%LocalAppData%\Spotify\Storage\*\|*

[Spotify - Downloaded Songs *]
Package=SpotifyAB.SpotifyMusic_zpdnekdrzrea0
SkipUWPFileKeys=True
LangSecRef=3023
DetectFile=%AppData%\Spotify
FileKeyBase=<Package>\LocalState\Spotify\Storage|*|REMOVESELF
FileKey=%LocalAppData%\Spotify\Storage|*|REMOVESELF
```

**Output**

```ini
[Spotify - Downloaded Songs *]
LangSecRef=3023
DetectFile1=%AppData%\Spotify
DetectFile2=%LocalAppData%\Packages\SpotifyAB.SpotifyMusic_zpdnekdrzrea0
FileKey1=%LocalAppData%\Packages\SpotifyAB.SpotifyMusic_zpdnekdrzrea0\LocalState\Spotify\Storage|*|REMOVESELF
FileKey2=%LocalAppData%\Spotify\Storage|*|REMOVESELF

[Spotify *]
LangSecRef=3023
DetectFile1=%AppData%\Spotify
DetectFile2=%LocalAppData%\Packages\SpotifyAB.SpotifyMusic_zpdnekdrzrea0
FileKey1=%AppData%\Spotify|*.log
FileKey2=%AppData%\Spotify\users\*|ad-state-storage.bnk;frequency.pb;local-files.bnk;local_sync_state;watch-sources.bnk
FileKey3=%LocalAppData%\Packages\SpotifyAB.SpotifyMusic_zpdnekdrzrea0\AC|*|RECURSE
FileKey4=%LocalAppData%\Packages\SpotifyAB.SpotifyMusic_zpdnekdrzrea0\LocalCache\Spotify|*|REMOVESELF
FileKey5=%LocalAppData%\Packages\SpotifyAB.SpotifyMusic_zpdnekdrzrea0\LocalState\Users\*|ad-state-storage.bnk;cached;context_player_state_restore
FileKey6=%LocalAppData%\Packages\SpotifyAB.SpotifyMusic_zpdnekdrzrea0\Settings|*.log*
FileKey7=%LocalAppData%\Packages\SpotifyAB.SpotifyMusic_zpdnekdrzrea0\SystemAppData\Helium|*.log*
FileKey8=%LocalAppData%\Packages\SpotifyAB.SpotifyMusic_zpdnekdrzrea0\TempState|*|REMOVESELF
FileKey9=%LocalAppData%\Spotify|*|REMOVESELF
ExcludeKey1=PATH|%LocalAppData%\Spotify\Storage\*\|*
```

**Explanation**
- `SkipUWPFileKeys` suppressed only the four scaffold FileKeys on the secondary entry; its detection was still generated (both entries carry the same two DetectFiles)
- The primary entry keeps full baseline coverage but excludes the `Storage` folder, which belongs to the secondary entry.
- Note the entries are alphabetized in the output, and FileKey5's parameter list was alphabetized by the lint pass

---

### Example 8: Validation Warnings and Skips

**Context**

AppInfo sections are validated before generation. It's worth seeing what rejection actually looks like, and what happens to entries that only *partially* misbehave.

**Intent**

We run three deliberately flawed sections through the builder: one with no `Package=`, one declaring both category keys plus an unsupported key, and one carrying `Skip=False`.

**Files**

###### **AppInfo section (`AppInfo\src.ini`)**
```ini
[Hazel Notes *]
LangSecRef=3021
FileKeyBase=<Package>\LocalState\Logs|*

[Hazel Weather *]
Package=Hazel.Weather_123abc
LangSecRef=3021
Section=Hazel Apps
Author=Hazel

[Hazel Legacy *]
Package=Hazel.Legacy_123abc
LangSecRef=3021
Skip=False
```

**Warnings**

The run reports:

```
No Package key in [Hazel Notes *], skipping
Unexpected key type in [Hazel Weather *]: Author
Both LangSecRef and Section present in [Hazel Weather *], using LangSecRef
```

**Output**

```ini
[Hazel Weather *]
LangSecRef=3021
DetectFile=%LocalAppData%\Packages\Hazel.Weather_123abc
FileKey1=%LocalAppData%\Packages\Hazel.Weather_123abc\AC|*|RECURSE
FileKey2=%LocalAppData%\Packages\Hazel.Weather_123abc\Settings|*.log*
FileKey3=%LocalAppData%\Packages\Hazel.Weather_123abc\SystemAppData\Helium|*.log*
FileKey4=%LocalAppData%\Packages\Hazel.Weather_123abc\TempState|*|REMOVESELF
```

**Explanation**
- `[Hazel Notes *]` was skipped entirely because it has no `Package=` key
- `[Hazel Weather *]` was generated with warnings: the conflicting `Section=` was discarded in favor of `LangSecRef`, and the unsupported `Author=` key was dropped entirely
- `[Hazel Legacy *]` was skipped **silently**: `Skip` is presence-based, so `Skip=False` skips just like `Skip=True`
- Only one of the three sections produced an entry
