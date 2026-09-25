# Scaffolds

### What is this?

This folder contains the shared cleaning-pattern catalogs consumed by winapp2ool's [UWPBuilder](https://github.com/MoscaDotTo/Winapp2/blob/master/winapp2ool/modules/uwpbuilder/readme.md) and [EntryBuilder](https://github.com/MoscaDotTo/Winapp2/blob/master/winapp2ool/modules/entrybuilder/readme.md) modules. When a source entry in either module declares an embedded browser data folder, the generator expands the entry's selected scaffolds from these catalogs into concrete FileKeys such that one curated catalog of Chromium cleaning patterns covers every app that needs them. 

### What is a scaffold?

A scaffold is a named group of FileKey templates. Each catalog section defines one scaffold:

```ini
[WebViewScaffold: WebCookies]
FileKeyBase=%WebViewRoot%\*\Network|Cookies*;Device Bound Sessions*
```

The section name format is `[WebViewScaffold: Name]` in `webview.ini`, `[QtWebEngineScaffold: Name]` in `qtwebengine.ini`, and `[ElectronScaffold: Name]` in `electron.ini`. Each `FileKeyBase=` line is a FileKey template; the family placeholder is substituted by the consuming module once per root the entry declares.

The three families each define a root differently, and getting this wrong is the most common way to write a scaffold that silently matches nothing:

| Family | Placeholder | What the root names |
| :- | :- | :- |
| WebView2 | `%WebViewRoot%` | The folder containing the profiles (`...\EBWebView`). Profile-scoped templates use `%WebViewRoot%\*\`.
| QtWebEngine | `%QtWebEngineRoot%` | One profile directory, segment included (`...\QtWebEngine\Default`) |
| Electron | `%ElectronRoot%` | One profile directory, which is the app's `userData` folder (`%AppData%\Signal`). |

Two families have a second placeholder, declared separately because it can't be inferred from the first:

* `%ElectronUpdaterRoot%`: the electron-updater download cache (`%LocalAppData%\signal-updater`)
* `%QtWebEngineCacheRoot%`: the folder holding the profile's HTTP `Cache\`, which Qt keeps apart from the profile (`%LocalAppData%\VideoKeeper\cache\QtWebEngine\Default`)

A template whose placeholder has no declared root is dropped.

### How do entries select scaffolds?

An entry opts a family in by declaring its root key; without it, no scaffold keys are emitted for that family:

| Consumer     | Opt-in key                          | Selection keys                                                                                    |
| :-           | :-                                  | :-                                                                                                |
| UWPBuilder   | `WebViewPath=` / `QtWebEnginePath=` / `QtWebEngineCachePath=` / `ElectronRoot=` / `ElectronUpdaterRoot=` | `WebViewScaffolds=`, `QtWebEngineScaffolds=`, `ElectronScaffolds=`, each with a matching `Exclude...Scaffolds=` |
| EntryBuilder | `WebViewRoot=` / `QtWebEngineRoot=` / `QtWebEngineCacheRoot=` / `ElectronRoot=` / `ElectronUpdaterRoot=` | The same selection key names as UWPBuilder |

Root keys are repeatable. To declare several roots, number them:

```ini
ElectronRoot1=%AppData%\Notion
ElectronRoot2=%AppData%\Notion\partitions\*
```

The selection contract is identical in every module and every family:

* With no selection keys, an opted-in entry receives that family's default set (below)
* `...Scaffolds=` **replaces** the default set with the listed scaffolds
* `Exclude...Scaffolds=` **subtracts** from the selected set
* The sentinel `All` (case-insensitive) expands to the entire catalog **except the legacy tier**. `All` is reserved and cannot be used as a scaffold name
* `...Scaffolds=All` + `Exclude...Scaffolds=X,Y` to delete "everything except X and Y" 
* A legacy scaffold is generated only when named, alone or beside `All`: `WebViewScaffolds=All,LegacyTelemetry`

The host-risk scaffolds (`WebCookies`, `WebStorage`, `WebHistory`, `WebSession`, `LoginData`) remove data an application may treat as primary user state, such as logged-in sessions and saved passwords. They are never in a default set, but `All` includes them, so an entry that selects `All` excludes the ones it must keep.

### The legacy tier

A section carrying `Tier=Legacy` stays in the catalog but is withheld from `All`:

```ini
[WebViewScaffold: LegacyWebHistory]
Tier=Legacy
FileKeyBase=%WebViewRoot%\*|shortcuts*
```

It holds patterns that modern hosts no longer write, but that older hosts may.

### Catalog contents

Defaults in **bold**, legacy tier in *italics*

`webview.ini` currently defines: Autofill, Autoplay, BookmarkFavicons, **Caches**, DefaultApps, DownloadHistory, DRMData, PrivacySandbox, LoginData, Security, Shopping, Sync Data, **Telemetry**, WebCookies, WebHistory, WebSession, WebStorage, *LegacyCaches*, *LegacyDownloadHistory*, *LegacyDRMData*, *LegacyExtensionCookies*, *LegacyProgressiveWebApps*, *LegacyStorageQuota*, *LegacyTelemetry*, *LegacyWebHistory*.

`qtwebengine.ini` currently defines: **Caches**, Favicons, PrivacySandbox, Security, **StorageQuota**, **Telemetry**, **VisitedLinks**, WebCookies, WebHistory, WebSession, WebStorage, *LegacyTelemetry*, *LegacyWebStorage*.

`electron.ini` currently defines: **AppLogs**, **Caches**, MediaDRM, PrivacySandbox, Security, **StorageQuota**, **Telemetry**, TempFiles, **UpdaterCache**, WebCookies, WebStorage.

Note that the QtWebEngine default set (bolded above) is wider than the WebView one. `StorageQuota` and `VisitedLinks` are separate default-on scaffolds there rather than members of the host-risk `WebStorage` / `WebHistory` sets, because the hand-written QtWebEngine entries the catalog replaces treated both as routine cleaning. The catalog header explains the reasoning for each tier decision.

### Notes for contributors

* Adding a section to the catalog enables it in both modules automatically and it will automatically be included in scaffolds invoking `All`, unless it carries `Tier=Legacy`
* A catalog's engine family is read from its section headers, not its filename, and both modules are pointed at this whole folder rather than at individual files. A section header whose family no module consumes triggers a warning.
* Renaming a scaffold breaks every source entry that selects or excludes it by name

# Files

| Name                                                                                                                              | Description                                                                          |
| :-                                                                                                                                | :-                                                                                   |
| [webview.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/Scaffolds/webview.ini)             | The scaffold catalog for embedded WebView2 / EBWebView data folders                  |
| [qtwebengine.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/Scaffolds/qtwebengine.ini)     | The scaffold catalog for embedded QtWebEngine data folders                           |
| [electron.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/Scaffolds/electron.ini)           | The scaffold catalog for Electron application data folders                            |
