# UWPBuilder Source Files

### What is this?

This folder holds the entry scaffold (`UWP.ini`) and the app definitions (`AppInfo\*.ini`) that winapp2ool's [UWPBuilder](https://github.com/MoscaDotTo/Winapp2/blob/master/winapp2ool/modules/uwpbuilder/readme.md) module turns into winapp2.ini entries for Universal Windows Platform (Microsoft Store) apps. During the build, every app defined here gets the scaffold's baseline detection and cleaning keys, and the result is written to [`Assembler/Generated/UWPBuilder/`](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/Generated/UWPBuilder) as 27 per-letter files, one per AppInfo file. Those are merged into the base winapp2.ini with the rest of the generated content.

### Why are UWP entries generated?

Every UWP app keeps its data under `%LocalAppData%\Packages\<PackageFolderName>`, and most of what's in there is laid out the same way from app to app (`\AC\`, `\TempState\`, and so on). Instead of writing out the same FileKeys with a different long package path for every app, UWPBuilder applies one scaffold to all of them. A complete entry can be as short as two lines, and a change to `UWP.ini` reaches every generated entry on the next build.

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

The `DetectFile` and the four baseline FileKeys come from the `[EntryScaffold: UWP App]` section of `UWP.ini`, with `<Package>` replaced by the app's package folder. Keys specific to the app (more FileKeys, RegKeys, or the desktop paths of an app that also has a win32 version) are added after them, with the same `<Package>` substitution. In an app that declares several packages, `<Package[2]>` picks the second. The old `%Package%` spelling is refused: the key is dropped and the build fails.

### Embedded browser data

An AppInfo section can declare `WebViewPath=` (an embedded WebView2/EBWebView data folder), `QtWebEnginePath=` (an embedded QtWebEngine data folder) or `ElectronRoot=` (an Electron `userData` folder) to get more FileKeys from the shared catalogs in [Scaffolds](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/Sources/Scaffolds). `WebViewScaffolds=` and `ExcludeWebViewScaffolds=` (and their QtWebEngine and Electron counterparts) choose which ones. See the [Scaffolds readme](https://github.com/MoscaDotTo/Winapp2/blob/master/Assembler/Sources/Scaffolds/readme.md).

The [UWPBuilder readme](https://github.com/MoscaDotTo/Winapp2/blob/master/winapp2ool/modules/uwpbuilder/readme.md) documents the rest of what an AppInfo section can do.

# Files

| Name | Description |
| :- | :- |
| [UWP.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/Sources/UWPBuilder/UWP.ini) | The entry scaffold: the DetectFile and FileKey templates every generated app gets |
| [AppInfo](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/Sources/UWPBuilder/AppInfo) | 27 definition files split by first letter, one section per app |
