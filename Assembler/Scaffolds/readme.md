# Scaffolds

### What is this?

This folder contains the shared cleaning-pattern catalogs consumed by winapp2ool's [UWPBuilder](https://github.com/MoscaDotTo/Winapp2/blob/master/winapp2ool/modules/uwpbuilder/readme.md) and [EntryBuilder](https://github.com/MoscaDotTo/Winapp2/blob/master/winapp2ool/modules/entrybuilder/readme.md) modules. When a source entry in either module declares an embedded browser data folder, the generator expands the entry's selected scaffolds from these catalogs into concrete FileKeys such that one curated catalog of Chromium cleaning patterns covers every app that needs them. 

### What is a scaffold?

A scaffold is a named group of FileKey templates. Each catalog section defines one scaffold:

```ini
[WebViewScaffold: WebCookies]
FileKeyBase=%WebViewRoot%\Default\Network|Cookies*;Device Bound Sessions*
```

The section name format is `[WebViewScaffold: Name]` in `webview.ini`, `[QtWebEngineScaffold: Name]` in `qtwebengine.ini`, and `[ElectronScaffold: Name]` in `electron.ini`. Each `FileKeyBase=` line is a FileKey template; the family placeholder is substituted by the consuming module once per root the entry declares.

The three families each define a root differently, and getting this wrong is the most common way to write a scaffold that silently matches nothing:

| Family | Placeholder | What the root names |
| :- | :- | :- |
| WebView2 | `%WebViewRoot%` | The folder containing the profile (`...\EBWebView`). These templates bake the `\Default\` segment in themselves |
| QtWebEngine | `%QtWebEngineRoot%` | One profile directory, segment included (`...\QtWebEngine\Default`) |
| Electron | `%ElectronRoot%` | One profile directory, which is the app's `userData` folder (`%AppData%\Signal`). |

Electron is also the only family with a second placeholder: `%ElectronUpdaterRoot%`, the electron-updater download cache. It is declared separately, and cannot be inferred.A template whose placeholder has no declared root is dropped.

### How do entries select scaffolds?

An entry opts a family in by declaring its root key; without it, no scaffold keys are emitted for that family:

| Consumer     | Opt-in key                          | Selection keys                                                                                    |
| :-           | :-                                  | :-                                                                                                |
| UWPBuilder   | `WebViewPath=` / `QtWebEnginePath=` / `ElectronRoot=` / `ElectronUpdaterRoot=` | `WebViewScaffolds=`, `QtWebEngineScaffolds=`, `ElectronScaffolds=`, each with a matching `Exclude...Scaffolds=` |
| EntryBuilder | `WebViewRoot=` / `QtWebEngineRoot=` / `ElectronRoot=` / `ElectronUpdaterRoot=` | The same selection key names as UWPBuilder |

Root keys are repeatable. To declare several roots, number them:

```ini
ElectronRoot1=%AppData%\Notion
ElectronRoot2=%AppData%\Notion\partitions\*
```

The selection contract is identical in every module and every family:

* With no selection keys, an opted-in entry receives that family's default set (below)
* `...Scaffolds=` **replaces** the default set with the listed scaffolds
* `Exclude...Scaffolds=` **subtracts** from the selected set
* The sentinel `All` (case-insensitive) expands to the entire catalog. `All` is reserved and cannot be used as a scaffold name
* `...Scaffolds=All` + `Exclude...Scaffolds=X,Y` to delete "everything except X and Y" 

Scaffolds outside the default set exist in tiers (documented in each catalog's header comment). The host-risk tier (`cookies`, `web storage`, `history`, `sessions`, `saved logins`) removes data an application may treat as primary user state (logged-in sessions, saved passwords). These are never included by default.

### Catalog contents

Defaults in **bold**

`webview.ini` currently defines: Autofill, Autoplay, BookmarkBackups, BookmarkFavicons, **Caches**, DefaultApps, DownloadHistory, DRMData, ExtensionCookies, ProgressiveWebApps, PrivacySandbox, LoginData, Security, Shopping, StorageQuota, **Telemetry**, WebCookies, WebHistory, WebSession, WebStorage.

`qtwebengine.ini` currently defines: **Caches**, **StorageQuota**, **Telemetry**, **VisitedLinks**, WebCookies, WebHistory, WebSession, WebStorage.

`electron.ini` currently defines: **AppLogs**, **Caches**, MediaDRM, PrivacySandbox, Security, **StorageQuota**, **Telemetry**, TempFiles, **UpdaterCache**, WebCookies, WebStorage.

Note that the QtWebEngine default set (bolded above) is wider than the WebView one. `StorageQuota` and `VisitedLinks` are separate default-on scaffolds there rather than members of the host-risk `WebStorage` / `WebHistory` sets, because the hand-written QtWebEngine entries the catalog replaces treated both as routine cleaning. The catalog header explains the reasoning for each tier decision.

### Notes for contributors

* Adding a section to the catalog enables it in both modules automatically and it will automatically be included in scaffolds invoking `All`
* A catalog's engine family is read from its section headers, not its filename, and both modules are pointed at this whole folder rather than at individual files. A section header whose family no module consumes triggers a warning.
* Renaming a scaffold breaks every source entry that selects or excludes it by name

# Files

| Name                                                                                                                              | Description                                                                          |
| :-                                                                                                                                | :-                                                                                   |
| [webview.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/Scaffolds/webview.ini)             | The scaffold catalog for embedded WebView2 / EBWebView data folders                  |
| [qtwebengine.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/Scaffolds/qtwebengine.ini)     | The scaffold catalog for embedded QtWebEngine data folders                           |
| [electron.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/Scaffolds/electron.ini)           | The scaffold catalog for Electron application data folders                            |
