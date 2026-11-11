# UWPBuilder Examples

Worked examples for [UWPBuilder](readme.md), each with real input and output. The [Common Uses](readme.md#common-uses) table in the readme lists the commands.

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

Every example below is a real run of UWPBuilder against the scaffold template and catalogs in `Assembler/`. Each shows an AppInfo section and what UWPBuilder made from it, and they all use the same command:

```
winapp2ool -uwpbuilder -1d \Sources\UWPBuilder -2f uwp.ini -3d \Sources\Scaffolds
```

###### Note: The outputs below leave out the generated header (see [Output Formatting](readme.md#output-formatting)).

### Example 1: A Minimal Entry

**Context**

Most UWP apps keep nothing worth cleaning outside the standard package folders. For those apps, the scaffold baseline is the whole entry.

**Intent**

We want a complete entry for ASUS Quick Launch from nothing more than its package and category.

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
- The package name and category are the only parts written by hand

---

### Example 2: App-Specific FileKeys

**Context**

Apple Devices keeps logs in several locations under its package that the baseline scaffold doesn't cover.

**Intent**

We want to add three FileKeys inside the package to the baseline.

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
- The three app-specific keys end up mixed in with the baseline, because the lint pass sorts FileKeys by value and the source order is lost
- All seven keys were renumbered after sorting

---

### Example 3: Hybrid UWP + win32 Applications

**Context**

Amazon Music ships as both a UWP package and a win32 install. One entry can clean both.

**Intent**

We want `<Package>` keys (`FileKeyBase=`) for the UWP install, win32 keys passed through as written (`FileKey=`), and registry detection for the win32 version beside the generated package detection.

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
- `Detect=` passes through as written and detects the win32 install. The package `DetectFile` is generated beside it
- The win32 `FileKey=` values have no `<Package>` variable, so they pass through unchanged
- One entry covers both installs

---

### Example 4: Multi-Package Applications

**Context**

Cortana has shipped under two package names. Some of its data lives under both packages, and some only under the newer one.

**Intent**

We want one entry for both packages, with shared keys expanded for every package (`<Package>`) and package-specific keys aimed at just one (`<Package[2]>`).

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
- The scaffold's `DetectFileBase=<Package>` produced one `DetectFile` per package
- The baseline and the `<Package>` FileKey expanded once per package: keys 1–5 for `Package1`, keys 6–7 and 9–11 for `Package2`
- The `<Package[2]>` keys (FileKey8's `LocalState` and the ExcludeKey) expanded once, for the second package only

---

### Example 5: WebView Scaffolds (Default Selection)

**Context**

GoldKey Acellus is an education app that embeds WebView2 and keeps its browser data under `LocalState\EBWebView`.

**Intent**

We want the entry to get the default scaffold selection (Caches and Telemetry) from the shared catalog, declaring nothing but the WebView root.

**Files**

###### **AppInfo section**
```ini
[GoldKey Acellus *]
Package=GoldKeyCorporation.Acellus_v2rn50m687qry
WebViewRoot=<Package>\LocalState\EBWebView
LangSecRef=3021
```

###### Note: The entry the winapp2.ini project ships for this app selects `WebViewScaffolds=All` with exclusions, which is [Example 6](#example-6-webview-scaffolds-all-with-exclusions). This one leaves the selection keys out to show the default.

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
- Declaring `WebViewRoot=` opted the entry into WebView scaffolds
- With no `WebViewScaffolds=` key, it got the default Caches and Telemetry selection
- Each catalog template had `<WebViewRoot>` replaced with the root path, `<Package>` already expanded
- One line of source produced 26 WebView FileKeys
- The lint pass merged templates aimed at the same folder: FileKey4 combines a Caches template and a Telemetry template that both target `EBWebView\*`

---

### Example 6: WebView Scaffolds (All With Exclusions)

**Context**

In a single-purpose app like Acellus, the browsing data inside the WebView (history, cookies, sessions) is safe to delete, but clearing `WebStorage` or `LoginData` would sign the user out or destroy data kept on the machine.

**Intent**

We want every scaffold outside the legacy tier *except* the two that remove the user's own data. This is the version of the entry the winapp2.ini project ships.

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

The full output has **59 FileKeys**. The excerpt below shows its shape, with the middle cut:

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
- `WebViewScaffolds=All` selected every scaffold outside the legacy tier, logged as `WebViewScaffolds=All in [GoldKey Acellus *]; expanded to 17 scaffold(s) from catalog, 8 legacy withheld`. `ExcludeWebViewScaffolds=WebStorage,LoginData` then took those two out
- Compare FileKey4 with Example 5's: the same `EBWebView\*` folder now collects patterns from many more scaffolds, merged into one key by the lint pass
- No legacy scaffold was generated. To add one, name it: `WebViewScaffolds=All,LegacyTelemetry`

---

### Example 7: Secondary Entries With SkipUWPFileKeys

**Context**

Spotify's downloaded songs take up a lot of disk space, but people don't necessarily expect a cleaner to delete their offline music. That calls for a *separate* entry, so each can be turned on and off by itself, but a second entry for the same package would repeat the scaffold baseline.

**Intent**

We want a main `[Spotify *]` entry with full coverage, and a second `[Spotify - Downloaded Songs *]` entry that shares the package but carries only its own keys.

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
- `SkipUWPFileKeys` left out only the four baseline FileKeys on the second entry. Its detection was still generated, so both entries carry the same two DetectFiles
- The main entry keeps the full baseline but excludes the `Storage` folder, which belongs to the second entry
- The entries come out in alphabetical order, and the lint pass sorted FileKey5's pattern list

---

### Example 8: Validation Warnings and Skips

**Context**

UWPBuilder checks each AppInfo section before generating it. This example shows what a rejected section looks like, and what happens to one that only *partly* misbehaves.

**Intent**

We run three deliberately broken sections through the builder: one with no `Package=`, one with both category keys and a key UWPBuilder doesn't know, and one with `Skip=False`.

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
Both LangSecRef and Section present in [Hazel Weather *], using LangSecRef
Variable 'Author=' declared in [Hazel Weather *] but never referenced by any <Author> token; possible typo
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
- `[Hazel Weather *]` was generated with warnings. The conflicting `Section=` was thrown away in favor of `LangSecRef`. `Author=` isn't a key UWPBuilder knows, so it declared a variable, and since nothing uses `<Author>`, the unused-variable check flagged it as a possible typo. It doesn't appear in the output
- `[Hazel Legacy *]` was skipped **without a warning**: `Skip` takes effect whenever it's present, so `Skip=False` skips just like `Skip=True`
- Only one of the three sections became an entry
