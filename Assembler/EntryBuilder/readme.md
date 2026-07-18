# EntryBuilder Source Files

### What is this?

This folder contains the alphabetically-split source files used by winapp2ool's [EntryBuilder](https://github.com/MoscaDotTo/Winapp2/blob/master/winapp2ool/modules/entrybuilder/readme.md) module to generate winapp2.ini entries from a shorthand DSL. During the build, all 27 files are combined in-memory, the shorthand is expanded into standard winapp2.ini syntax, and the output (`entrybuilder.ini`) is merged into the base Winapp2.ini alongside the other generated content.

### Why a shorthand?

EntryBuilder is a preprocessor for winapp2.ini: anything that is already valid winapp2 syntax passes through unchanged, and a small set of winapp2ool-exclusive keys expand into standard form at generation time. Repetitive entries can be written once, declaratively:

* **List variables**: declare `Versions=10.0,11.0,12.0` once and reference it from any key as `<Versions>`; each referencing key fans out into one output key per value. Inline lists (`<a,b,c>` written directly inside a key value) fan out the same way without a declaration
* **Automatic detection**: an entry which declares the reserved variable `Root=` receives a `Detect` (registry root) or `DetectFile` (filesystem root) generated from it automatically
* **Scaffold families**: an entry which declares `WebViewRoot=` or `QtWebEngineRoot=` receives baseline cleaning FileKeys for its embedded browser data drawn from the shared catalogs in [Scaffolds](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/Scaffolds) 

### Example

This source section:

```ini
[A-PDF Watermark *]
Root=HKCU\Software\A-PDF\Watermark

LangSecRef=3021
RegKeyBase=<Root>\Setting|<hdIndir,hdoutDir,hloadpdfpath,hSaveallpdfpath,ptlastprinter>
```

generates this winapp2.ini entry:

```ini
[A-PDF Watermark *]
LangSecRef=3021
Detect=HKCU\Software\A-PDF\Watermark
RegKey1=HKCU\Software\A-PDF\Watermark\Setting|hdIndir
RegKey2=HKCU\Software\A-PDF\Watermark\Setting|hdoutDir
RegKey3=HKCU\Software\A-PDF\Watermark\Setting|hloadpdfpath
RegKey4=HKCU\Software\A-PDF\Watermark\Setting|hSaveallpdfpath
RegKey5=HKCU\Software\A-PDF\Watermark\Setting|ptlastprinter
```

The `Detect` is inferred from `Root`, and the inline list fans the single `RegKeyBase` template out into one `RegKey` per value name.

The complete DSL is documented in the [EntryBuilder readme](https://github.com/MoscaDotTo/Winapp2/blob/master/winapp2ool/modules/entrybuilder/readme.md).

### Which entries live here?

Base entries which have been migrated to the EntryBuilder shorthand. 

# Files

| File                     | Description                                                                     |
| :-                       | :-                                                                              |
| `#.ini`                  | Contains entries for applications whose names begin with a number or symbol     |
| `A.ini` through `Z.ini`  | Contains entries for applications whose names begin with the respective letter  |
