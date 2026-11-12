# Scaffolds

### What is this?

This folder holds the shared cleaning-pattern catalogs used by winapp2ool's [UWPBuilder](https://github.com/MoscaDotTo/Winapp2/blob/master/winapp2ool/modules/uwpbuilder/readme.md) and [EntryBuilder](https://github.com/MoscaDotTo/Winapp2/blob/master/winapp2ool/modules/entrybuilder/readme.md) modules. When an entry in either one declares an embedded browser data folder, the generator turns the scaffolds that entry selects into real FileKeys. That way one curated catalog of Chromium cleaning patterns covers every app that embeds Chromium.

### What is a scaffold?

A scaffold is a named group of FileKey templates. Each catalog section defines one:

```ini
[WebViewScaffold: WebCookies]
FileKeyBase=<WebViewRoot>\*\Network|Cookies*;Device Bound Sessions*
```

Sections are named `[WebViewScaffold: Name]` in `webview.ini`, `[QtWebEngineScaffold: Name]` in `qtwebengine.ini`, and `[ElectronScaffold: Name]` in `electron.ini`. Each `FileKeyBase=` line is a FileKey template, and the module reading it fills in the family's placeholder once for each root the entry declares.

Each family means something different by a root, and getting that wrong is the most common way to write a scaffold that quietly matches nothing:

| Family | Placeholder | What the root names |
| :- | :- | :- |
| WebView2 | `<WebViewRoot>` | The folder holding the profiles (`...\EBWebView`). Profile-level templates use `<WebViewRoot>\*\` |
| QtWebEngine | `<QtWebEngineRoot>` | One profile folder, profile name included (`...\QtWebEngine\Default`) |
| Electron | `<ElectronRoot>` | One profile folder, which is the app's `userData` folder (`%AppData%\Signal`) |

Two families have a second placeholder, declared separately because it can't be worked out from the first:

* `<ElectronUpdaterRoot>`: the electron-updater download cache (`%LocalAppData%\signal-updater`)
* `<QtWebEngineCacheRoot>`: the folder holding the profile's HTTP `Cache\`, which Qt keeps apart from the profile (`%LocalAppData%\VideoKeeper\cache\QtWebEngine\Default`)

A template whose placeholder has no declared root is dropped.

Placeholders match in any case. They're spelled like the builders' `<Name>` variables, but they're filled in as plain text before variables expand, so a `<Package>` or `<Root>` inside a declared root expands along with the rest of the template. The old `%WebViewRoot%` spelling and its siblings (and UWPBuilder's old `%Package%`) are refused. A template holding one, even with its closing `%` missing, is dropped with a warning naming the new spelling, and the build fails:

```
`%ElectronRoot%` in [ElectronScaffold: Caches].FileKeyBase is now spelled `<ElectronRoot>`; template dropped
```

### How do entries select scaffolds?

An entry turns a family on by declaring its root key. Without one, that family adds nothing to the entry:

| Consumer     | Opt-in key                          | Selection keys                                                                                    |
| :-           | :-                                  | :-                                                                                                |
| UWPBuilder   | `WebViewPath=` / `QtWebEnginePath=` / `QtWebEngineCachePath=` / `ElectronRoot=` / `ElectronUpdaterRoot=` | `WebViewScaffolds=`, `QtWebEngineScaffolds=`, `ElectronScaffolds=`, each with a matching `Exclude...Scaffolds=` |
| EntryBuilder | `WebViewRoot=` / `QtWebEngineRoot=` / `QtWebEngineCacheRoot=` / `ElectronRoot=` / `ElectronUpdaterRoot=` | The same selection key names as UWPBuilder |

Root keys can repeat. To declare several roots, number them:

```ini
ElectronRoot1=%AppData%\Notion
ElectronRoot2=%AppData%\Notion\partitions\*
```

Selection works the same way in both modules and every family:

* With no selection keys, an entry that turned a family on gets that family's default set (below)
* `...Scaffolds=` replaces the default set with the scaffolds it lists. An empty `...Scaffolds=` selects none
* `Exclude...Scaffolds=` takes scaffolds out of whatever is selected
* `All` (in any case) means the whole catalog except the legacy tier. No scaffold can be named `All`
* `...Scaffolds=All` with `Exclude...Scaffolds=X,Y` gives you everything except X and Y
* A legacy scaffold is generated only when named, alone or beside `All`: `WebViewScaffolds=All,LegacyTelemetry`

The host-risk scaffolds (`WebCookies`, `WebStorage`, `WebHistory`, `WebSession`, `LoginData`) delete data an app may treat as the user's own, such as logged-in sessions and saved passwords. No default set includes them, but `All` does, so an entry that selects `All` has to exclude the ones its app needs to keep.

### The legacy tier

A section carrying `Tier=Legacy` stays in the catalog but is left out of `All`:

```ini
[WebViewScaffold: LegacyWebHistory]
Tier=Legacy
FileKeyBase=<WebViewRoot>\*|shortcuts*
```

The legacy tier holds patterns that current hosts no longer write but older ones may.

### Catalog contents

Defaults in **bold**, legacy tier in *italics*

`webview.ini` currently defines: Autofill, Autoplay, BookmarkFavicons, **Caches**, DefaultApps, DownloadHistory, DRMData, PrivacySandbox, LoginData, Security, Shopping, Sync Data, **Telemetry**, WebCookies, WebHistory, WebSession, WebStorage, *LegacyCaches*, *LegacyDownloadHistory*, *LegacyDRMData*, *LegacyExtensionCookies*, *LegacyProgressiveWebApps*, *LegacyStorageQuota*, *LegacyTelemetry*, *LegacyWebHistory*.

`qtwebengine.ini` currently defines: **Caches**, Favicons, PrivacySandbox, Security, **StorageQuota**, **Telemetry**, **VisitedLinks**, WebCookies, WebHistory, WebSession, WebStorage, *LegacyTelemetry*, *LegacyWebStorage*.

`electron.ini` currently defines: **AppLogs**, **Caches**, MediaDRM, PrivacySandbox, Security, **StorageQuota**, **Telemetry**, TempFiles, **UpdaterCache**, WebCookies, WebStorage.

QtWebEngine's default set is wider than WebView's. `StorageQuota` and `VisitedLinks` are their own default-on scaffolds there instead of being part of the host-risk `WebStorage` and `WebHistory`, because the hand-written QtWebEngine entries the catalog replaced treated both as routine cleaning. The catalog's header explains each tier decision.

### Notes for contributors

* A new catalog section reaches both modules on the next build, and `All` includes it unless it carries `Tier=Legacy`
* A catalog's engine family comes from its section headers, not its file name, and both modules read this whole folder rather than single files. A section header naming a family neither module reads gets a warning
* Renaming a scaffold breaks every entry that selects or excludes it by name

# Files

| Name                                                                                                                              | Description                                                                          |
| :-                                                                                                                                | :-                                                                                   |
| [webview.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/Sources/Scaffolds/webview.ini)             | The scaffold catalog for embedded WebView2 / EBWebView data folders                  |
| [qtwebengine.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/Sources/Scaffolds/qtwebengine.ini)     | The scaffold catalog for embedded QtWebEngine data folders                           |
| [electron.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/Sources/Scaffolds/electron.ini)           | The scaffold catalog for Electron application data folders                           |
