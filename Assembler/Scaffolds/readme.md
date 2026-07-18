# Scaffolds

### What is this?

This folder contains the shared cleaning-pattern catalogs consumed by winapp2ool's [UWPBuilder](https://github.com/MoscaDotTo/Winapp2/blob/master/winapp2ool/modules/uwpbuilder/readme.md) and [EntryBuilder](https://github.com/MoscaDotTo/Winapp2/blob/master/winapp2ool/modules/entrybuilder/readme.md) modules. When a source entry in either module declares an embedded browser data folder, the generator expands the entry's selected scaffolds from these catalogs into concrete FileKeys such that one curated catalog of Chromium cleaning patterns covers every app that needs them. 

### What is a scaffold?

A scaffold is a named group of FileKey templates. Each catalog section defines one scaffold:

```ini
[WebViewScaffold: WebCookies]
FileKeyBase=%WebViewRoot%\Default\Network|Cookies*;Device Bound Sessions*
```

The section name format is `[WebViewScaffold: Name]` in `webview.ini` and `[QtWebEngineScaffold: Name]` in `qtwebengine.ini`. Each `FileKeyBase=` line is a FileKey template; the family placeholder (`%WebViewRoot%` or `%QtWebEngineRoot%`) is substituted by the consuming module once per data-folder root the entry declares.

### How do entries select scaffolds?

An entry opts a family in by declaring its root key; without it, no scaffold keys are emitted for that family:

| Consumer     | Opt-in key                          | Selection keys                                                                                    |
| :-           | :-                                  | :-                                                                                                |
| UWPBuilder   | `WebViewPath=` / `QtWebEnginePath=` | `WebViewScaffolds=` / `ExcludeWebViewScaffolds=`, `QtWebEngineScaffolds=` / `ExcludeQtWebEngineScaffolds=` |
| EntryBuilder | `WebViewRoot=` / `QtWebEngineRoot=` | Same selection key names as UWPBuilder                                                            |

The selection contract is identical in both modules and both families:

* With no selection keys, an opted-in entry receives the default set: **Caches** and **Telemetry**
* `...Scaffolds=` **replaces** the default set with the listed scaffolds
* `Exclude...Scaffolds=` **subtracts** from the selected set
* The sentinel `All` (case-insensitive) expands to the entire catalog. `All` is reserved and cannot be used as a scaffold name
* `...Scaffolds=All` + `Exclude...Scaffolds=X,Y` to delete "everything except X and Y" 

Scaffolds outside the default set exist in tiers (documented in each catalog's header comment). The host-risk tier (`cookies`, `web storage`, `history`, `sessions`, `saved logins`) removes data an application may treat as primary user state (logged-in sessions, saved passwords). These are never included by default.

### Catalog contents

Defaults in **bold**

`webview.ini` currently defines: Autofill, Autoplay, BookmarkBackups, BookmarkFavicons, **Caches**, DefaultApps, DownloadHistory, DRMData, ExtensionCookies, ProgressiveWebApps, PrivacySandbox, LoginData, Security, Shopping, StorageQuota, **Telemetry**, WebCookies, WebHistory, WebSession, WebStorage.

`qtwebengine.ini` currently defines: **Caches**, **Telemetry**, WebCookies, WebHistory, WebSession, WebStorage.

### Notes for contributors

* Adding a section to the catalog enables it in both modules automatically and it will automatically be included in scaffolds invoking `All`
* Renaming a scaffold breaks every source entry that selects or excludes it by name

# Files

| Name                                                                                                                              | Description                                                                          |
| :-                                                                                                                                | :-                                                                                   |
| [webview.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/Scaffolds/webview.ini)             | The scaffold catalog for embedded WebView2 / EBWebView data folders                  |
| [qtwebengine.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/Scaffolds/qtwebengine.ini)     | The scaffold catalog for embedded QtWebEngine data folders                           |
