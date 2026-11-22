# EntryBuilder Source Files

### What is this?

This folder holds the sources for every base entry in winapp2.ini, which means every entry that isn't a generated browser entry, Microsoft Store apps included. They're split by first letter into 27 files. Changes to base entries go here; see the [contributor guidelines](https://github.com/MoscaDotTo/Winapp2/blob/master/CONTRIBUTING.md).

During the build, winapp2ool's [EntryBuilder](https://github.com/MoscaDotTo/Winapp2/blob/master/winapp2ool/modules/entrybuilder/readme.md) module reads all 27 files, expands any shorthand into standard winapp2.ini syntax, normalizes the result, and writes it out as 27 per-letter files in [`Assembler/Generated/`](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/Generated). Those are merged into the base winapp2.ini with the rest of the generated content. Comments you write here stay in the sources, but they don't reach the generated files.

**You don't need to learn the shorthand to contribute.** Standard winapp2.ini syntax works here and passes through generation unchanged.

### Why a shorthand?

EntryBuilder is a preprocessor for winapp2.ini. Valid winapp2 syntax passes through untouched, and a few keys only winapp2ool understands expand into standard syntax when the entry is generated. That lets a repetitive entry be written once:

* `Versions=10.0,11.0,12.0` declares a list variable, and any key that uses `<Versions>` becomes one key per value. An inline list (`<a,b,c>` written straight into a key's value) does the same without a declaration, and `<Versions[2]>` picks one value by its position
* An entry that declares `Root=` gets a `Detect` (for a registry root) or `DetectFile` (for a filesystem root) made from it
* An entry that declares `WebViewRoot=`, `QtWebEngineRoot=` or `ElectronRoot=` gets baseline FileKeys for its embedded browser data from the shared catalogs in [Scaffolds](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/Sources/Scaffolds)
* An entry that declares `Package=` is a Microsoft Store app: it gets a `DetectFile` for its folder under `%LocalAppData%\Packages`, a `<Package>` variable for its own keys, and the baseline FileKeys every Store app gets. A hybrid app keeps its win32 keys in the same entry. The [EntryBuilder readme](https://github.com/MoscaDotTo/Winapp2/blob/master/winapp2ool/modules/entrybuilder/readme.md#store-and-hybrid-apps) explains how

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

The `Detect` comes from `Root`, and the inline list turns the one `RegKeyBase` template into a `RegKey` for each value name.

The [EntryBuilder readme](https://github.com/MoscaDotTo/Winapp2/blob/master/winapp2ool/modules/entrybuilder/readme.md) documents the whole shorthand.

### Which entries live here?

Every winapp2.ini entry that isn't a browser entry, Microsoft Store apps included. Most are written in plain winapp2.ini syntax, and some have been converted to the shorthand.

Archived entries live here too, marked with an `Archive=` line.

### Archived entries

An entry with an `Archive=` key goes to the archive instead of winapp2.ini. EntryBuilder writes it to [`Assembler/Archived/EntryBuilder/archived.ini`](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/Archived), and the build publishes the archive as `archived.ini` on every release. The value is the reason it was archived. A bare `Archive=` archives too.

We put the line directly under the section header, followed by a blank line:

```ini
[LEGO Universe *]
Archive=Discontinued (2012)

Section=Games
DetectFile=%LocalAppData%\LEGO Software\LEGO Universe
FileKey1=%LocalAppData%\LEGO Software\LEGO Universe\Log Files|*
```

To retire an entry, add the line. To bring it back, delete it. An archived entry is built and checked like a live one, it stays in its letter's file in alphabetical order, and its name can't match a live entry's. The [EntryBuilder readme](https://github.com/MoscaDotTo/Winapp2/blob/master/winapp2ool/modules/entrybuilder/readme.md#archive) has the details.

# Files

| File | Description |
| :- | :- |
| `#.ini` | Entries for applications whose names start with a number or symbol |
| `A.ini` through `Z.ini` | Entries for applications whose names start with that letter |

A pull request that puts an entry in the wrong letter's file fails the [generated artifact guard](https://github.com/MoscaDotTo/Winapp2/blob/master/CONTRIBUTING.md#the-generated-artifact-guard), which names the file it belongs in.
