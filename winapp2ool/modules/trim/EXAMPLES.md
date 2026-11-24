# Trim Examples

###### Trim produces results that are fit to individual machines; you may not be able to reproduce the examples locally.

# Table of Contents

1. [Example 1: Anatomy of a Trim](#example-1-anatomy-of-a-trim)
2. [Example 2: Trimming the Full Database](#example-2-trimming-the-full-database)
3. [Example 3: Downloading and Trimming in One Step](#example-3-downloading-and-trimming-in-one-step)
4. [Example 4: Never-Trim and Always-Trim Overrides](#example-4-never-trim-and-always-trim-overrides)
5. [Example 5: VirtualStore Key Generation](#example-5-virtualstore-key-generation)

--- 

## Example 1: Anatomy of a Trim

**Context**

We have a file containing various detection criteria: `Detect`, `DetectFile` and `DetectOS`, and an entry with none of them. 

This computer has 7-Zip, Firefox, Steam, and Windows Calculator installed.

**Intent**

We want to trim this file in place and see which entries are kept

**Files**

###### **Input file (`winapp2.ini`)**

```ini
[7-Zip ZS *]
LangSecRef=3024
Detect=HKCU\Software\7-Zip-Zstandard
RegKey1=HKCU\Software\7-Zip-Zstandard\Compression|ArcHistory
RegKey2=HKCU\Software\7-Zip-Zstandard\Extraction|PathHistory
RegKey3=HKCU\Software\7-Zip-Zstandard\FM|CopyHistory
RegKey4=HKCU\Software\7-Zip-Zstandard\FM|FolderHistory

[Mozilla Firefox Autofill Data *]
LangSecRef=3026
DetectFile1=%AppData%\Mozilla\Firefox\Profiles
DetectFile2=%LocalAppData%\Packages\Mozilla.Firefox_*
FileKey1=%AppData%\Mozilla\Firefox\Profiles\*|autofill-profiles.json
FileKey2=%LocalAppData%\Packages\Mozilla.Firefox_*\LocalCache\Roaming\Mozilla\Firefox\Profiles\*|autofill-profiles.json

[My Custom Cleanup *]
Section=Custom Entries
FileKey1=%UserProfile%\Downloads|*.partial

[Steam Packages *]
Section=Games
Detect=HKCU\Software\Valve\Steam
FileKey1=%ProgramFiles%\Steam\package|*.zip.*

[Windows Calculator *]
LangSecRef=3025
DetectFile=%LocalAppData%\Packages\Microsoft.WindowsCalculator_*
FileKey1=%LocalAppData%\Packages\Microsoft.WindowsCalculator_*\AC|*|RECURSE
FileKey2=%LocalAppData%\Packages\Microsoft.WindowsCalculator_*\LocalCache|*|RECURSE
FileKey3=%LocalAppData%\Packages\Microsoft.WindowsCalculator_*\LocalState\Cache|*|RECURSE
FileKey4=%LocalAppData%\Packages\Microsoft.WindowsCalculator_*\Settings|*.log*

[Windows XP Era Application *]
Section=Legacy Applications
DetectOS=|5.2
FileKey1=%WinDir%\Prefetch|*.pf
```

**Command**

```
winapp2ool -trim
```

**Output**

###### **Console**

```
 ╔════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╗
 ║                                                      Trim Complete                                                 ║
 ╠════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╣
 ║ Initial entry count: 6                                                                                             ║
 ║ Trimmed entry count: 4                                                                                             ║
 ║ 2 entries trimmed from winapp2.ini (33%)                                                                           ║
 ╠════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╣
 ║                                          Press any key to return to the menu.                                      ║
 ╚════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╝
```

###### **Output file (`winapp2.ini`) after trimming**

```ini
[Mozilla Firefox Autofill Data *]
LangSecRef=3026
DetectFile1=%AppData%\Mozilla\Firefox\Profiles
DetectFile2=%LocalAppData%\Packages\Mozilla.Firefox_*
FileKey1=%AppData%\Mozilla\Firefox\Profiles\*|autofill-profiles.json
FileKey2=%LocalAppData%\Packages\Mozilla.Firefox_*\LocalCache\Roaming\Mozilla\Firefox\Profiles\*|autofill-profiles.json

[My Custom Cleanup *]
Section=Custom Entries
FileKey1=%UserProfile%\Downloads|*.partial

[Steam Packages *]
Section=Games
Detect=HKCU\Software\Valve\Steam
FileKey1=%ProgramFiles%\Steam\package|*.zip.*

[Windows Calculator *]
LangSecRef=3025
DetectFile=%LocalAppData%\Packages\Microsoft.WindowsCalculator_*
FileKey1=%LocalAppData%\Packages\Microsoft.WindowsCalculator_*\AC|*|RECURSE
FileKey2=%LocalAppData%\Packages\Microsoft.WindowsCalculator_*\LocalCache|*|RECURSE
FileKey3=%LocalAppData%\Packages\Microsoft.WindowsCalculator_*\LocalState\Cache|*|RECURSE
FileKey4=%LocalAppData%\Packages\Microsoft.WindowsCalculator_*\Settings|*.log*
```

**Explanation**

• `[7-Zip ZS *]` is removed, `HKCU\Software\7-Zip-Zstandard` does not exist on this machine. 7-Zip itself is installed, but this entry targets for the Zstandard fork

• `[Mozilla Firefox Autofill Data *]` is kept, `%AppData%\Mozilla\Firefox\Profiles` exists 

• `[My Custom Cleanup *]` is kept; an entry with no detection criteria is always kept


• `[Steam Packages *]` is kept, `HKCU\Software\Valve\Steam` exists


• `[Windows Calculator *]` is kept, the wildcard `DetectFile` expanded to the real package folder `Microsoft.WindowsCalculator_8wekyb3d8bbwe`

• `[Windows XP Era Application *]` is removed, `DetectOS=\|5.2` requires Windows ≤ 5.2 (XP/Server 2003); this machine has 10.0 (11)

---

## Example 2: Trimming the Full Database

**Context**

A freshly downloaded winapp2.ini is untrimmed.

**Intent**

We want to trim the full winapp2.ini into a separate file and leave the original alone.

**Command**

```
winapp2ool -trim -3f winapp2-trimmed.ini
```

**Output**

###### **Console**

```
 ╔════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╗
 ║                                                      Trim Complete                                                 ║
 ╠════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╣
 ║ Initial entry count: 3715                                                                                          ║
 ║ Trimmed entry count: 349                                                                                           ║
 ║ 3366 entries trimmed from winapp2.ini (91%)                                                                        ║
 ╠════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╣
 ║                                          Press any key to return to the menu.                                      ║
 ╚════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╝
```

###### **Output file (`winapp2-trimmed.ini`), first lines**

```ini
; Version: 251109
; # of entries: 349
;
; Winapp2.ini is fully licensed under the CC-BY-SA-4.0 license agreement. ...
```

**Explanation**

- The input is `winapp2.ini` in the current directory, the default
- Trim writes `winapp2-trimmed.ini` and leaves the input alone
- Trim removed 91% of the database. The 349 entries left are the ones that passed detection on this machine
- The `; Version: 251109` line comes from the input file, and `; # of entries:` shows the count after trimming

---

## Example 3: Downloading and Trimming in One Step

**Context**

We only ever want a trimmed file, so we don't need to keep a copy of the complete database on disk. 

**Intent**

We want to download the latest winapp2.ini from GitHub and trim it to fit the current machine. 

**Command**

```
winapp2ool -trim -d -3f winapp2-trimmed.ini
```

**Output**

###### **Console**

```
 ╔════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╗
 ║                                                      Trim Complete                                                 ║
 ╠════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╣
 ║ Initial entry count: 3715                                                                                          ║
 ║ Trimmed entry count: 349                                                                                           ║
 ║ 3366 entries trimmed from winapp2.ini (91%)                                                                        ║
 ╚════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╝
```

**Explanation**

- `-d` downloads the latest database and trims it without needing a local file to be present first
- `-d` ignores the `-1f`/`-1d` input file settings
- The output is `winapp2-trimmed.ini`

---

## Example 4: Never-Trim and Always-Trim Overrides

**Context**

Continuing from [Example 1](#example-1-anatomy-of-a-trim): suppose we plan to install 7-Zip Zstandard soon, so we want its entry to be kept even though it isn't detected yet. Likewise, we never want `[Steam Packages *]`, even though Steam is installed.

**Intent**

We want to keep `[7-Zip ZS *]` and drop `[Steam Packages *]`, irrespective of whether or not they're installed on the current system 

**Files**

Using the same input file as Example 1, plus:

###### **Includes file (`includes.ini`)**

```ini
[7-Zip ZS *]
```

###### **Excludes file (`excludes.ini`)**

```ini
[Steam Packages *]
```

**Command**

```
winapp2ool -trim -includes -excludes -3f winapp2-trimmed.ini
```

**Output**

###### **Entries in `winapp2-trimmed.ini`**

```ini
[Mozilla Firefox Autofill Data *]
[7-Zip ZS *]
[My Custom Cleanup *]
[Windows Calculator *]
```

**Explanation**

- `[7-Zip ZS *]` is now kept: the include list overrides its failing `Detect`
- `[Steam Packages *]` is now removed: the exclude list overrides its passing `Detect`
- `[Windows XP Era Application *]` is still removed as it was in example 1

**Notes**

Trim checks both lists before it looks at any detection key. If an entry appears in both files, the include list wins and the entry is always kept. 

---

## Example 5: VirtualStore Key Generation

**Context**

Under some circumstances, Windows may redirect a program's writes to into the VirtualStore. This computer has no such leftovers, so for this example we created an empty `%LocalAppData%\VirtualStore\Program Files (x86)\Steam\package` folder before running Trim

**Intent**

We want surviving entries to clean their VirtualStore copies too, but only where those copies exist.

**Files**

The same input file as [Example 1](#example-1-anatomy-of-a-trim); the relevant entry:

```ini
[Steam Packages *]
Section=Games
Detect=HKCU\Software\Valve\Steam
FileKey1=%ProgramFiles%\Steam\package|*.zip.*
```

**Command**

```
winapp2ool -trim -3f winapp2-trimmed.ini
```

**Output**

###### **The entry in `winapp2-trimmed.ini` after trimming**

```ini
[Steam Packages *]
Section=Games
Detect=HKCU\Software\Valve\Steam
FileKey1=%LocalAppData%\VirtualStore\Program Files*\Steam\package|*.zip.*
FileKey2=%ProgramFiles%\Steam\package|*.zip.*
```

**Explanation**

- `[Steam Packages *]` passed detection, so Trim looked for VirtualStore counterparts of its FileKeys
- `%ProgramFiles%\Steam\package` maps to `%LocalAppData%\VirtualStore\Program Files*\Steam\package`. That folder existed during this run, so Trim added a key for it and renumbered the entry's FileKeys
- Without that folder, the entry would passe through unchanged, as it did in Example 1
