# UWPBuilder Source Files

### What is this?

This folder contains the entry scaffold (`UWP.ini`) and the AppInfo definitions (`AppInfo\*.ini`) used by winapp2ool's [UWPBuilder](https://github.com/MoscaDotTo/Winapp2/blob/master/winapp2ool/modules/uwpbuilder/readme.md) module to generate winapp2.ini entries for Universal Windows Platform (Microsoft Store) applications. During the build, the AppInfo files are combined, each defined application receives the scaffold's baseline detection and cleaning keys, and the output (`uwp.ini`) is merged into the base Winapp2.ini alongside the other generated content.

### Why are UWP entries generated?

Every UWP application stores its data under `%LocalAppData%\Packages\<PackageFolderName>`, and most of that layout is identical from app to app (`\AC\`, `\TempState\`, and so on). Rather than hand-writing the same FileKeys with a different verbose package path hundreds of times, UWPBuilder applies a shared scaffold to every application: a complete entry can be defined in as little as two lines, and changing the template in `UWP.ini` changes every generated entry on the next build.

### Example

This AppInfo section:

```ini
[BreeZip *]
Package=3138AweZip.AweZip_ffd303wmbhcjt
LangSecRef=3021
```

generates this winapp2.ini entry:

```ini
[BreeZip *]
LangSecRef=3021
DetectFile=%LocalAppData%\Packages\3138AweZip.AweZip_ffd303wmbhcjt
FileKey1=%LocalAppData%\Packages\3138AweZip.AweZip_ffd303wmbhcjt\AC|*|RECURSE
FileKey2=%LocalAppData%\Packages\3138AweZip.AweZip_ffd303wmbhcjt\Settings|*.log*
FileKey3=%LocalAppData%\Packages\3138AweZip.AweZip_ffd303wmbhcjt\SystemAppData\Helium|*.log*
FileKey4=%LocalAppData%\Packages\3138AweZip.AweZip_ffd303wmbhcjt\TempState|*|REMOVESELF
```

The `DetectFile` and the four baseline FileKeys come from the `[EntryScaffold: UWP App]` section of `UWP.ini`, with `%Package%` expanded to the declared package folder. App-specific keys (additional FileKeys, RegKeys, win32 paths for hybrid apps) are layered on top and pass through with the same `%Package%` substitution.

### Embedded browser data

An AppInfo section may declare `WebViewPath=` (an embedded WebView2/EBWebView data folder) or `QtWebEnginePath=` (an embedded QtWebEngine data folder) to draw additional curated FileKeys from the shared catalogs in [Scaffolds](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/Scaffolds), selected per entry with `WebViewScaffolds=` / `ExcludeWebViewScaffolds=` (and their QtWebEngine equivalents). See the [Scaffolds readme](https://github.com/MoscaDotTo/Winapp2/blob/master/Assembler/Scaffolds/readme.md).

The complete UWPBuilder DSL is documented in the [UWPBuilder readme](https://github.com/MoscaDotTo/Winapp2/blob/master/winapp2ool/modules/uwpbuilder/readme.md).

# Files

| Name                                                                                                            | Description                                                                                              |
| :-                                                                                                              | :-                                                                                                       |
| [UWP.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/UWP/UWP.ini)         | The entry scaffold: baseline DetectFile and FileKey templates applied to every generated application     |
| [AppInfo](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/UWP/AppInfo)                              | 27 alphabetically-split definition files, one section per application                                    |
