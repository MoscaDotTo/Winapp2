# WinappDebug Examples

Worked examples for [WinappDebug](README.md). The readme's [Common Uses](README.md#common-uses) table lists the commands, and these examples show what each one does with real input and output.

Every file WinappDebug saves starts with winapp2.ini's header comment. The output files below leave it out, except in [Example 16](#example-16-scripted-in-place-repair).

# Table of Contents

1. [Scanning](#scanning)
   - [Example 1: Scanning a file for errors](#example-1-scanning-a-file-for-errors)
   - [Example 2: Auditing a single category](#example-2-auditing-a-single-category)
   - [Example 3: Reporting without repairing](#example-3-reporting-without-repairing)
2. [Repairing](#repairing)
   - [Example 4: Repairing a file and saving the corrections](#example-4-repairing-a-file-and-saving-the-corrections)
   - [Example 5: Renumbering and re-alphabetization](#example-5-renumbering-and-re-alphabetization)
   - [Example 6: Structural key repairs](#example-6-structural-key-repairs)
   - [Example 7: Casing, duplicates, and unneeded numbering](#example-7-casing-duplicates-and-unneeded-numbering)
   - [Example 8: Flag and ExcludeKey repairs](#example-8-flag-and-excludekey-repairs)
   - [Example 9: Path and registry validity](#example-9-path-and-registry-validity)
3. [Errors WinappDebug will not repair](#errors-winappdebug-will-not-repair)
   - [Example 10: Errors that cannot be repaired automatically](#example-10-errors-that-cannot-be-repaired-automatically)
4. [Flavors and Default keys](#flavors-and-default-keys)
   - [Example 11: Flavor-aware linting (System Ninja)](#example-11-flavor-aware-linting-system-ninja)
   - [Example 12: Flavor-aware linting (BleachBit)](#example-12-flavor-aware-linting-bleachbit)
   - [Example 13: Auditing Default key values](#example-13-auditing-default-key-values)
   - [Example 14: Preserving Default keys with -keepdefaults](#example-14-preserving-default-keys-with--keepdefaults)
5. [Advanced](#advanced)
   - [Example 15: Merging redundant FileKeys (Optimizations)](#example-15-merging-redundant-filekeys-optimizations)
   - [Example 16: Scripted in-place repair](#example-16-scripted-in-place-repair)

---

Some examples below reuse scenarios from the [Transmute examples](../transmute/EXAMPLES.md), where Transmute corrects BrowserBuilder output as part of the winapp2.ini build.

## Scanning

### Example 1: Scanning a file for errors

**Context**

We have a small winapp2.ini with a few common style mistakes: a forward slash in a path, a gap in the FileKey numbering, and `Default` keys, which current winapp2.ini style leaves out.

**Intent**

We want to see every error WinappDebug finds, without changing any file.

**Files**

###### **Input file (`winapp2.ini`)**

```ini
[7-Zip *]
LangSecRef=3021
Detect=HKCU\Software\7-Zip
FileKey1=%AppData%/7-Zip|*.tmp
Default=False

[Notepad++ *]
LangSecRef=3021
Detect=HKCU\Software\Notepad++
FileKey1=%AppData%\Notepad++|*.log
FileKey3=%AppData%\Notepad++\backup|*.*
Default=True
```

**Command**
```
winapp2ool -debug
```

**Output**

The report is rendered to the console:

```
 ╔════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╗
 ║                                                  Linting winapp2.ini                                               ║
 ╠                                                                                                                    ╣
 ║ Error in [7-Zip *]:                                                                                                ║
 ║ Forward slash (/) detected in lieu of backslash (\).                                                               ║
 ║ Key: FileKey1=%AppData%/7-Zip|*.tmp                                                                                ║
 ║                                                                                                                    ║
 ║ Error in [7-Zip *]:                                                                                                ║
 ║ Entry has a Default key where there should be none                                                                 ║
 ║                                                                                                                    ║
 ║ Error in [Notepad++ *]:                                                                                            ║
 ║ FileKey entry is incorrectly numbered.                                                                             ║
 ║ Expected: FileKey2                                                                                                 ║
 ║ Found:    FileKey3                                                                                                 ║
 ║                                                                                                                    ║
 ║ Error in [Notepad++ *]:                                                                                            ║
 ║ Entry has a Default key where there should be none                                                                 ║
 ║                                                                                                                    ║
 ╠                                                                                                                    ╣
 ║                                                     Lint Complete!                                                 ║
 ╠                                                                                                                    ╣
 ║                                                     Entry count: 2                                                 ║
 ║                                              4 possible errors detected.                                           ║
 ║                                                                                                                    ║
 ║                                          Press any key to return to the menu.                                      ║
 ╚════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╝
```

###### Note: The other examples show just the report lines, without the frame around them.

**Explanation**
- WinappDebug lints `winapp2.ini`, its default input
- Both `Default` keys are flagged. Current style has no `Default` keys at all, so `Default=False` is as wrong as `Default=True`
- The numbering error names the key it expected and the key it found
- Saving is off, so nothing is written to disk

---

### Example 2: Auditing a single category

**Context**

We want to know whether the file system and registry paths in a copy of winapp2.ini are well formed.

**Intent**

We want a report on path problems and nothing else, with no changes made to the entries.

**Files**

###### **Input file (`winapp2.ini`)**

```ini
[HandBrake *]
LangSecRef=3023
Detect=HKCU\Software\HandBrake
DetectFile1=%LocalAppData%\HandBrake\
DetectFile2=%LocalAppData%\HandBrake\*\logs\*
FileKey1=%LocalAppData%\HandBrake\logs|*.txt
FileKey3=%LocalAppData%\HandBrake\cache|*.tmp
RegKey1=Software\HandBrake\Recent
```

**Steps**

`alloff` is typed at the Scan Settings menu. It isn't a command-line flag:

1. Open **WinappDebug** from the main menu
2. Select **Toggle Saving** to turn saving on
3. Select **Scan Settings**
4. Type `alloff`. Every scan and every repair is now off
5. Select item **13** (*Path Validity*) under **Scan Options** to turn that scan back on by itself
6. Go back to the WinappDebug menu with `0`, then select **Run**

**Output**

```
Error in [HandBrake *]:
Nested wildcard found in DetectFile
Key: DetectFile2=%LocalAppData%\HandBrake\*\logs\*

Error in [HandBrake *]:
Invalid registry path detected.
Key: RegKey1=Software\HandBrake\Recent

2 possible errors detected.
```

###### **Output file (`winapp2-debugged.ini`)**

```ini
[HandBrake *]
LangSecRef=3023
Detect=HKCU\Software\HandBrake
DetectFile1=%LocalAppData%\HandBrake\
DetectFile2=%LocalAppData%\HandBrake\*\logs\*
FileKey1=%LocalAppData%\HandBrake\logs|*.txt
FileKey3=%LocalAppData%\HandBrake\cache|*.tmp
RegKey1=Software\HandBrake\Recent
```

**Explanation**
- `RegKey1` is the Path Validity finding
- The nested wildcard check is *(always on)*, so it reports anyway
- The trailing backslash on `DetectFile1` (Slashes) and `FileKey3`'s numbering (Improper Numbering) go unreported, because their categories are off
- Both findings are report-only, so turning Path Validity's repair on as well wouldn't have fixed either one
- Saving was on only to show that the saved copy matches the input

---

### Example 3: Reporting without repairing

**Context**

We're reviewing a contribution and want WinappDebug's full report, but we don't trust its slash handling on this particular file. We want to look at the paths ourselves before anything rewrites them.

**Intent**

We want every category to report as usual, with only the Slashes repair turned off.

**Files**

The input file is the same as in [Example 1](#example-1-scanning-a-file-for-errors).

**Steps**

1. Open **WinappDebug** from the main menu
2. Select **Toggle Saving** to turn saving on
3. Select **Scan Settings**
4. Select item **21** (*Slashes*) under **Repair Options** to turn that repair off
5. Go back with `0`, then select **Run**

**Output**

```
Error in [7-Zip *]:
Forward slash (/) detected in lieu of backslash (\).
Key: FileKey1=%AppData%/7-Zip|*.tmp

Error in [7-Zip *]:
Missing backslash (\) after %EnvironmentVariable%.
Key: FileKey1=%AppData%/7-Zip|*.tmp

Error in [7-Zip *]:
Entry has a Default key where there should be none

Error in [Notepad++ *]:
FileKey entry is incorrectly numbered.
Expected: FileKey2
Found:    FileKey3

Error in [Notepad++ *]:
Entry has a Default key where there should be none

5 possible errors detected.
```

###### **Output file (`winapp2-debugged.ini`)**

```ini
[7-Zip *]
LangSecRef=3021
Detect=HKCU\Software\7-Zip
FileKey1=%AppData%/7-Zip|*.tmp

[Notepad++ *]
LangSecRef=3021
Detect=HKCU\Software\Notepad++
FileKey1=%AppData%\Notepad++|*.log
FileKey2=%AppData%\Notepad++\backup|*.*
```

**Explanation**
- The forward slash is still reported, but not repaired
- Every other repair ran as usual: both `Default` keys are gone and `FileKey3` became `FileKey2`
- The same file gives 5 errors here and 4 in [Example 1](#example-1-scanning-a-file-for-errors). The note below explains the extra one

**Notes**

`Missing backslash (\) after %EnvironmentVariable%.` never fires in Example 1, because by the time that check runs, the Slashes repair has already turned `%AppData%/` into `%AppData%\`.

---

## Repairing

### Example 4: Repairing a file and saving the corrections

**Context**

Picking up from [Example 1](#example-1-scanning-a-file-for-errors): we've read the report and want the corrections written to disk.

**Intent**

We want WinappDebug to repair all four errors and save the result.

**Files**

The input file is the same as in Example 1.

**Command**
```
winapp2ool -debug -c
```

**Output**

The console report matches Example 1, plus one more summary line:

```
winapp2-debugged.ini saved with any corrections made
```

###### **Output file (`winapp2-debugged.ini`)**

```ini
[7-Zip *]
LangSecRef=3021
Detect=HKCU\Software\7-Zip
FileKey1=%AppData%\7-Zip|*.tmp

[Notepad++ *]
LangSecRef=3021
Detect=HKCU\Software\Notepad++
FileKey1=%AppData%\Notepad++|*.log
FileKey2=%AppData%\Notepad++\backup|*.*
```

**Explanation**
- The forward slash becomes a backslash
- `FileKey3` becomes `FileKey2`
- Both `Default` keys are gone

---

### Example 5: Renumbering and re-alphabetization

**Context**

After some hand editing, an entry's FileKeys are out of order and the entries themselves aren't alphabetical anymore.

**Intent**

We want WinappDebug to sort and renumber the keys, and sort the entries.

**Files**

###### **Input file (`winapp2.ini`)**

```ini
[Speccy *]
LangSecRef=3025
Detect=HKCU\Software\Piriform\Speccy
FileKey2=%LocalAppData%\Programs\Speccy|*.log
FileKey1=%AppData%\Speccy|*.xml

[CCleaner *]
LangSecRef=3025
Detect=HKCU\Software\Piriform\CCleaner
FileKey1=%ProgramFiles%\CCleaner|*.log
```

**Command**
```
winapp2ool -debug -c
```

**Output**

```
Error in [Speccy *]:
FileKey entry is incorrectly numbered.
Expected: FileKey1
Found:    FileKey2

Error in [Speccy *]:
FileKey entry is incorrectly numbered.
Expected: FileKey2
Found:    FileKey1

Error in [Speccy *]:
FileKey alphabetization
FileKey1=%LocalAppData%\Programs\Speccy|*.log appears to be out of place
Expected position: 2

Error in Speccy *:
Entry alphabetization
Speccy * appears to be out of place
Expected position: 2

4 possible errors detected.
```

###### **Output file (`winapp2-debugged.ini`)**

```ini
[CCleaner *]
LangSecRef=3025
Detect=HKCU\Software\Piriform\CCleaner
FileKey1=%ProgramFiles%\CCleaner|*.log

[Speccy *]
LangSecRef=3025
Detect=HKCU\Software\Piriform\Speccy
FileKey1=%AppData%\Speccy|*.xml
FileKey2=%LocalAppData%\Programs\Speccy|*.log
```

**Explanation**
- WinappDebug sorts the two FileKeys by value and renumbers them to match their new positions
- `[Speccy *]` moves after `[CCleaner *]`

**Notes**

WinappDebug repairs errors as it finds them, so later messages can refer to keys it has already repaired. The alphabetization error calls the key `FileKey1`, the name the renumbering gave it.

---

### Example 6: Structural key repairs

**Context**

A hand-written entry has three broken keys: one with stray whitespace, one with no `=` at all, and one ending in a semicolon.

**Intent**

We want all three fixed.

**Files**

###### **Input file (`winapp2.ini`)**

```ini
[Greenshot *]
LangSecRef=3021
Detect=HKCU\Software\Greenshot
FileKey1=%AppData%\Greenshot|*.log;
FileKey2%AppData%\Greenshot\thumbnails|*.png
FileKey3= %LocalAppData%\Greenshot|*.tmp
```

**Command**
```
winapp2ool -debug -c
```

**Output**

```
Error in [Greenshot *]:
Detected unwanted whitespace in iniKey
Key: FileKey3= %LocalAppData%\Greenshot|*.tmp

Error in [Greenshot *]:
Missing '=' detected and repaired in key.
Key: FileKey2=%AppData%\Greenshot\thumbnails|*.png

Error in [Greenshot *]:
Trailing semicolon (;).
Key: FileKey1=%AppData%\Greenshot|*.log;

Error in [Greenshot *]:
FileKey entry is incorrectly numbered.
Expected: FileKey2
Found:    FileKey3

Error in [Greenshot *]:
FileKey entry is incorrectly numbered.
Expected: FileKey3
Found:    FileKey2

Error in [Greenshot *]:
FileKey alphabetization
FileKey2=%LocalAppData%\Greenshot|*.tmp appears to be out of place
Expected position: 3

6 possible errors detected.
```

###### **Output file (`winapp2-debugged.ini`)**

```ini
[Greenshot *]
LangSecRef=3021
Detect=HKCU\Software\Greenshot
FileKey1=%AppData%\Greenshot|*.log
FileKey2=%AppData%\Greenshot\thumbnails|*.png
FileKey3=%LocalAppData%\Greenshot|*.tmp
```

**Explanation**
- `FileKey2` gets its missing `=` back
- The whitespace and the trailing semicolon are stripped
- The keys are sorted and renumbered

---

### Example 7: Casing, duplicates, and unneeded numbering

**Context**

An entry has picked up four separate problems from hand editing: a stray second `LangSecRef`, a `Detect1` in an entry with only one Detect, a lowercase environment variable, and a FileKey that duplicates another key's value.

**Intent**

We want to see what four of the repairs that delete or normalize keys do to one entry in a single pass.

**Files**

###### **Input file (`winapp2.ini`)**

```ini
[Foobar2000 *]
LangSecRef=3023
LangSecRef=3024
Detect1=HKCU\Software\foobar2000
FileKey1=%appdata%\foobar2000|*.log
FileKey2=%AppData%\foobar2000\cache|*.tmp
FileKey3=%AppData%\foobar2000|*.log
```

**Command**
```
winapp2ool -debug -c
```

**Output**

```
Error in [Foobar2000 *]:
Multiple LangSecRef detected.
Key: LangSecRef=3024

Error in [Foobar2000 *]:
Detected unnecessary numbering.
Expected: Detect
Found:    Detect1

Error in [Foobar2000 *]:
AppData has a casing error.
Key: FileKey1=%appdata%\foobar2000|*.log

Error in [Foobar2000 *]:
Duplicate key value found
Key:            FileKey3=%AppData%\foobar2000|*.log
Duplicates:     FileKey1=%AppData%\foobar2000|*.log

4 possible errors detected.
```

###### **Output file (`winapp2-debugged.ini`)**

```ini
[Foobar2000 *]
LangSecRef=3023
Detect=HKCU\Software\foobar2000
FileKey1=%AppData%\foobar2000|*.log
FileKey2=%AppData%\foobar2000\cache|*.tmp
```

**Explanation**
- `LangSecRef=3024` is deleted. An entry can only have one `LangSecRef`, and the repair keeps the first
- `Detect1` becomes `Detect`, since the entry has only one Detect
- `%appdata%` becomes `%AppData%`
- `FileKey3` is deleted because it duplicates `FileKey1`, once the casing repair has fixed `FileKey1`

---

### Example 8: Flag and ExcludeKey repairs

**Context**

A hand-written entry gets flags wrong in nearly every way it can: a `RECURSE` with no pipe before it, an ExcludeKey with no flag at all, one with its flag run into the path, one missing the backslash before its pattern pipe, and one with a flag it shouldn't have.

**Intent**

We want to see which of these the Flags and Path Validity repairs can fix, and which need a person.

**Files**

###### **Input file (`winapp2.ini`)**

```ini
[Free Download Manager *]
LangSecRef=3021
Detect=HKCU\Software\FreeDownloadManager
FileKey1=%AppData%\FreeDownloadManager\logs|*.logRECURSE
ExcludeKey1=%AppData%\FreeDownloadManager\sessions|*.dat
ExcludeKey2=FILE%AppData%\FreeDownloadManager\queue|*.dat
ExcludeKey3=FILE|%AppData%\FreeDownloadManager\cache|*.tmp
ExcludeKey4=PATH|%AppData%\FreeDownloadManager\plugins\|*.*|RECURSE
```

**Command**
```
winapp2ool -debug -c
```

**Output**

```
Error in [Free Download Manager *]:
Missing pipe (|) before RECURSE.
Key: FileKey1=%AppData%\FreeDownloadManager\logs|*.logRECURSE

Error in [Free Download Manager *]:
No valid exclude flag (FILE, PATH, or REG) found in ExcludeKey.
Key: ExcludeKey1=%AppData%\FreeDownloadManager\sessions|*.dat

Error in [Free Download Manager *]:
Missing pipe (|) after ExcludeKey flag
Key: ExcludeKey2=FILE%AppData%\FreeDownloadManager\queue|*.dat

Error in [Free Download Manager *]:
Missing backslash (\) before pipe (|) in ExcludeKey.
Key: ExcludeKey2=FILE|%AppData%\FreeDownloadManager\queue|*.dat

Error in [Free Download Manager *]:
Missing backslash (\) before pipe (|) in ExcludeKey.
Key: ExcludeKey3=FILE|%AppData%\FreeDownloadManager\cache|*.tmp

Error in [Free Download Manager *]:
ExcludeKey has too many flags
Key: ExcludeKey4=PATH|%AppData%\FreeDownloadManager\plugins\|*.*|RECURSE

Error in [Free Download Manager *]:
ExcludeKey alphabetization
ExcludeKey2=FILE|%AppData%\FreeDownloadManager\queue\|*.dat appears to be out of place
Expected position: 3

7 possible errors detected.
```

###### **Output file (`winapp2-debugged.ini`)**

```ini
[Free Download Manager *]
LangSecRef=3021
Detect=HKCU\Software\FreeDownloadManager
FileKey1=%AppData%\FreeDownloadManager\logs|*.log|RECURSE
ExcludeKey1=%AppData%\FreeDownloadManager\sessions|*.dat
ExcludeKey2=FILE|%AppData%\FreeDownloadManager\cache\|*.tmp
ExcludeKey3=FILE|%AppData%\FreeDownloadManager\queue\|*.dat
ExcludeKey4=PATH|%AppData%\FreeDownloadManager\plugins\|*.*|RECURSE
```

**Explanation**
- `FileKey1` gets its missing pipe: `*.logRECURSE` becomes `*.log|RECURSE`
- `ExcludeKey1` has no flag, and WinappDebug can't guess one, so it stays as it is
- `ExcludeKey2` is repaired twice. First its flag gets a pipe (`FILE%AppData%` becomes `FILE|%AppData%`), and then it turns out to be missing the backslash before its pattern pipe as well
- `ExcludeKey3` gets the same missing backslash
- `ExcludeKey4` has both a `PATH` flag and a `RECURSE` flag. The extra flag is reported but left alone
- After the repairs the ExcludeKeys are sorted again, so `queue` and `cache` swap places and get renumbered

---

### Example 9: Path and registry validity

**Context**

An entry mixes good and bad paths: a lowercase registry root, a `DetectFile` with a trailing backslash, a `DetectFile` with a wildcard in the middle of its path, and a `RegKey` with no hive.

**Intent**

We want to see which path problems are mechanical and which mean the path itself is wrong.

**Files**

###### **Input file (`winapp2.ini`)**

```ini
[HandBrake *]
LangSecRef=3023
Detect=hkcu\Software\HandBrake
DetectFile1=%LocalAppData%\HandBrake\
DetectFile2=%LocalAppData%\HandBrake\*\logs\*
FileKey1=%LocalAppData%\HandBrake\logs|*.txt
RegKey1=Software\HandBrake\Recent
```

**Command**
```
winapp2ool -debug -c
```

**Output**

```
Error in [HandBrake *]:
Incorrect registry root casing.
Key: Detect=hkcu\Software\HandBrake

Error in [HandBrake *]:
Trailing backslash (\) found in DetectFile
Key: DetectFile1=%LocalAppData%\HandBrake\

Error in [HandBrake *]:
Nested wildcard found in DetectFile
Key: DetectFile2=%LocalAppData%\HandBrake\*\logs\*

Error in [HandBrake *]:
Invalid registry path detected.
Key: RegKey1=Software\HandBrake\Recent

4 possible errors detected.
```

###### **Output file (`winapp2-debugged.ini`)**

```ini
[HandBrake *]
LangSecRef=3023
Detect=HKCU\Software\HandBrake
DetectFile1=%LocalAppData%\HandBrake
DetectFile2=%LocalAppData%\HandBrake\*\logs\*
FileKey1=%LocalAppData%\HandBrake\logs|*.txt
RegKey1=Software\HandBrake\Recent
```

**Explanation**
- `hkcu\` becomes `HKCU\`
- `DetectFile1` loses its trailing backslash
- `DetectFile2`'s nested wildcard is reported but left alone
- `RegKey1` has no registry hive. It's reported and left alone

---

## Errors WinappDebug will not repair

### Example 10: Errors that cannot be repaired automatically

**Context**

An entry has an invalid `LangSecRef` value and no detection keys at all. WinappDebug can't know the right category or the right detection path, so there's nothing it can safely repair.

**Intent**

We want to see what happens to report-only errors when saving is on.

**Files**

###### **Input file (`winapp2.ini`)**

```ini
[Winamp *]
LangSecRef=9999
FileKey1=%AppData%\Winamp|*.log
```

**Command**
```
winapp2ool -debug -c
```

**Output**

```
Error in [Winamp *]:
LangSecRef holds an invalid value.
Key: LangSecRef=9999

Error in [Winamp *]:
Entry has no valid detection keys (Detect, DetectFile, DetectOS, SpecialDetect)

2 possible errors detected.
```

###### **Output file (`winapp2-debugged.ini`)**

```ini
[Winamp *]
LangSecRef=9999
FileKey1=%AppData%\Winamp|*.log
```

**Explanation**
- Both errors are reported, and the saved entry is unchanged
- You have to fix report-only errors by hand

---

## Flavors and Default keys

### Example 11: Flavor-aware linting (System Ninja)

**Context**

System Ninja doesn't support wildcards in `DetectFile` keys, so the System Ninja flavor of winapp2.ini can't use them. The same key is fine in the base and CCleaner flavors.

**Intent**

We want to check a file against System Ninja's rules for one run.

**Files**

###### **Input file (`winapp2.ini`)**

```ini
[Brave Session *]
Section=Brave Web Browser
DetectFile=%LocalAppData%\BraveSoftware\Brave-Browser*
FileKey1=%LocalAppData%\BraveSoftware\Brave-Browser*\User Data\*\Sessions|*|REMOVESELF
```

**Command**
```
winapp2ool -sn -debug
```

**Output**

```
Error in [Brave Session *]:
Wildcard (*) found in DetectFile
Key: DetectFile=%LocalAppData%\BraveSoftware\Brave-Browser*

1 possible errors detected.
```

**Explanation**
- `-sn` sets the flavor to System Ninja for this run, which turns on its DetectFile wildcard check
- The same command without `-sn` reports `0 possible errors detected.`
- The check is report-only, so you have to fix the key yourself

---

### Example 12: Flavor-aware linting (BleachBit)

**Context**

BleachBit versions 5 and older didn't implement registry exclusions, so an `ExcludeKey` with the `REG` flag should be reported as an error.

**Intent**

We want to find registry ExcludeKeys before a file goes out to BleachBit users.

**Files**

###### **Input file (`winapp2.ini`)**

```ini
[WinRAR *]
LangSecRef=3021
Detect=HKCU\Software\WinRAR
FileKey1=%AppData%\WinRAR|*.log
RegKey1=HKCU\Software\WinRAR\DialogEditHistory
ExcludeKey1=REG|HKCU\Software\WinRAR\DialogEditHistory\ArcName
```

**Command**
```
winapp2ool -bb -debug
```

**Output**

```
Error in [WinRAR *]:
ExcludeKey contains REG flag in BleachBit flavor
Key: ExcludeKey1=REG|HKCU\Software\WinRAR\DialogEditHistory\ArcName

1 possible errors detected.
```

**Explanation**
- `-bb` turns on the BleachBit check for this run only. Your saved Flavor setting doesn't change
- The same run without `-bb` reports `0 possible errors detected.`
- Like the System Ninja check, this one is report-only

---

### Example 13: Auditing Default key values

**Context**

We're preparing a file for a program that expects every entry to have `Default=False`.

**Intent**

We want every entry to end up with `Default=False`, added where it's missing and corrected where it's wrong.

**Files**

###### **Input file (`winapp2.ini`)**

```ini
[Audacity *]
LangSecRef=3023
Detect=HKCU\Software\Audacity
FileKey1=%AppData%\audacity|*.log

[VLC Media Player *]
LangSecRef=3023
Detect=HKCU\Software\VideoLAN\VLC
FileKey1=%AppData%\vlc|*.log
Default=True
```

**Steps**

The Default Value Audit has no command-line flag, so we turn it on from the menu:

1. Open **WinappDebug** from the main menu
2. Select **Toggle Saving** to turn saving on
3. Select **Toggle Default Value Audit** to turn the audit on (the expected value starts out as `False`)
4. Select **Run**

**Output**

```
Error in [Audacity *]:
No Default Key found

Error in [VLC Media Player *]:
Incorrect value for Default Key found
Key: Default=True

2 possible errors detected.
```

###### **Output file (`winapp2-debugged.ini`)**

```ini
[Audacity *]
LangSecRef=3023
Detect=HKCU\Software\Audacity
Default=False
FileKey1=%AppData%\audacity|*.log

[VLC Media Player *]
LangSecRef=3023
Detect=HKCU\Software\VideoLAN\VLC
Default=False
FileKey1=%AppData%\vlc|*.log
```

**Explanation**
- `[Audacity *]` had no Default key, so it gets `Default=False`
- `[VLC Media Player *]`'s `Default=True` becomes `Default=False`
- For a program that wants everything enabled, **Toggle Expected Default** switches the required value to `True`
- Without the audit, the same run would *remove* `Default=True` instead, and say nothing about `[Audacity *]`. See [Example 1](#example-1-scanning-a-file-for-errors)

---

### Example 14: Preserving Default keys with -keepdefaults

**Context**

FluentCleaner turns entries on unless they say otherwise, so the FluentCleaner flavor's build marks browser-privacy entries `Default=False` and leaves cache and telemetry entries enabled. Those `Default=` keys are much of the point of the flavor. Linting that output under the usual Defaults policy would quietly delete every one of them, and the file would still look fine.

**Intent**

We want to lint the FluentCleaner output for real style errors and leave its `Default=` keys alone.

**Files**

###### **Input file (`winapp2.ini`)**

```ini
[Vivaldi Web Browsing Cookies *]
LangSecRef=3033
DetectFile=%LocalAppData%\Vivaldi\User Data
Default=False
FileKey1=%LocalAppData%\Vivaldi\User Data\*\Network|Cookies;Cookies-journal

[Vivaldi Web Browsing Internet Cache *]
LangSecRef=3033
DetectFile=%LocalAppData%\Vivaldi\User Data
FileKey1=%LocalAppData%\Vivaldi\User Data\*\Cache|*|RECURSE
```

**Commands**

First without the flag, then with it:

```
winapp2ool -debug -c -3f no-flag.ini
winapp2ool -debug -c -keepdefaults -3f with-flag.ini
```

**Output**

The first command reports the `Default` key as an error:

```
Error in [Vivaldi Web Browsing Cookies *]:
Entry has a Default key where there should be none

1 possible errors detected.
```

The second reports nothing:

```
0 possible errors detected.
```

###### **Output file (`no-flag.ini`) without the flag**

```ini
[Vivaldi Web Browsing Cookies *]
LangSecRef=3033
DetectFile=%LocalAppData%\Vivaldi\User Data
FileKey1=%LocalAppData%\Vivaldi\User Data\*\Network|Cookies;Cookies-journal

[Vivaldi Web Browsing Internet Cache *]
LangSecRef=3033
DetectFile=%LocalAppData%\Vivaldi\User Data
FileKey1=%LocalAppData%\Vivaldi\User Data\*\Cache|*|RECURSE
```

###### **Output file (`with-flag.ini`) with `-keepdefaults`**

```ini
[Vivaldi Web Browsing Cookies *]
LangSecRef=3033
DetectFile=%LocalAppData%\Vivaldi\User Data
Default=False
FileKey1=%LocalAppData%\Vivaldi\User Data\*\Network|Cookies;Cookies-journal

[Vivaldi Web Browsing Internet Cache *]
LangSecRef=3033
DetectFile=%LocalAppData%\Vivaldi\User Data
FileKey1=%LocalAppData%\Vivaldi\User Data\*\Cache|*|RECURSE
```

**Explanation**
- Without the flag, `Default=False` is reported as an error and deleted
- With `-keepdefaults`, WinappDebug neither reports nor touches the key, and the run finds no errors

---

## Advanced

### Example 15: Merging redundant FileKeys (Optimizations)

**Context**

[Transmute's Example 9](../transmute/EXAMPLES.md#example-9-correcting-syntax) adds a `FileKey` to a generated entry. The new key has no number, and it points at a path another FileKey already covers. The Optimizations category can fold the two together.

**Intent**

We want the extra FileKey merged into the existing key for the same path, and the rest of the entry's syntax cleaned up.

**Files**

###### **Input file (`winapp2.ini`)**

```ini
[360 Secure Browser Bookmarked Websites *]
Section=.360 Secure Browser Web Browser
DetectFile=%AppData%\360se6\User Data
FileKey1=%AppData%\360se6\User Data\*|bookmarks;BookmarkMergedSurfaceOrdering
FileKey2=%AppData%\360se6\User Data\*\power_bookmarks|*|REMOVESELF
FileKey=%AppData%\360se6\User Data\*|360Bookmarks*
```

**Command**
```
winapp2ool -debug -c -opti
```

###### Note: For menu runs, you can turn Optimizations on in **Scan Settings** instead

**Output**

```
Error in [360 Secure Browser Bookmarked Websites *]:
FileKey parameters are not in alphabetical order
Key:      FileKey1=%AppData%\360se6\User Data\*|bookmarks;BookmarkMergedSurfaceOrdering
Current:  bookmarks;BookmarkMergedSurfaceOrdering
Sorted:   BookmarkMergedSurfaceOrdering;bookmarks

Error in [360 Secure Browser Bookmarked Websites *]:
FileKey entry is incorrectly numbered.
Expected: FileKey3
Found:    FileKey

Error in [360 Secure Browser Bookmarked Websites *]:
FileKey alphabetization
FileKey3=%AppData%\360se6\User Data\*|360Bookmarks* appears to be out of place
Expected position: 1

[360 Secure Browser Bookmarked Websites *] has keys which can be merged
The following keys can be merged into other keys:
FileKey2=%AppData%\360se6\User Data\*|BookmarkMergedSurfaceOrdering;bookmarks
The resulting key list will be reduced to:
FileKey1=%AppData%\360se6\User Data\*|360Bookmarks*;BookmarkMergedSurfaceOrdering;bookmarks
FileKey2=%AppData%\360se6\User Data\*\power_bookmarks|*|REMOVESELF

3 possible errors detected.
```

###### **Output file (`winapp2-debugged.ini`)**

```ini
[360 Secure Browser Bookmarked Websites *]
Section=.360 Secure Browser Web Browser
DetectFile=%AppData%\360se6\User Data
FileKey1=%AppData%\360se6\User Data\*|360Bookmarks*;BookmarkMergedSurfaceOrdering;bookmarks
FileKey2=%AppData%\360se6\User Data\*\power_bookmarks|*|REMOVESELF
```

**Explanation**
- The unnumbered `FileKey` targets `%AppData%\360se6\User Data\*` with no flag, the same as `FileKey1`, so its `360Bookmarks*` parameter is merged into `FileKey1` and the extra key is deleted
- The merge gets its own report block, listing the keys it removes and the key list that results. It doesn't count toward the error total, but the numbering and ordering errors in the same entry count as usual
- The merged parameter list joins the parameters in key order (`360Bookmarks*;BookmarkMergedSurfaceOrdering;bookmarks`)
- [Transmute's Example 9](../transmute/EXAMPLES.md#example-9-correcting-syntax) runs the same cleanup after adding a key to a generated entry

---

### Example 16: Scripted in-place repair

**Context**

The winapp2.ini build lints each flavor as its last stage. Nothing is interactive and nothing goes to a second file: WinappDebug repairs the file where it sits and stamps it with the build date.

**Intent**

We want one non-interactive command that repairs a file in place and records when it was built.

**Files**

###### **Input file (`winapp2.ini`)**

```ini
[Audacity *]
LangSecRef=3023
Detect=HKCU\Software\Audacity
FileKey1=%AppData%\audacity|*.log
FileKey2=%AppData%\audacity\cache|*.tmp
```

**Command**

This is the build's own lint stage for the base file, copied from `build winapp2.ps1`, plus `-writelog` so the run leaves a record:

```
winapp2ool -s -offline -writelog -debug -usedate -c -opti -1f Winapp2.ini -3f Winapp2.ini
```

**Output**

Nothing appears on the console, since `-s` suppresses all of it. The report goes to `winapp2ool.log`, which ends with:

```
Lint complete
Entry count: 1
0 errors detected
2 ms
```

###### **Output file (`winapp2.ini`) first three lines**

```ini
; Version: 260724
; # of entries: 1
;
```

**Explanation**
- `-3f Winapp2.ini` points the save target at the input, so the file is repaired in place. Without it, the result would go to `winapp2-debugged.ini` and the input would stay as it was
- `-usedate` stamps the header's version line with today's date as `yymmdd`
- `-offline` skips the connection check at startup
- `-s` makes the run non-interactive. It also hides the whole report, which is why the command needs `-writelog`
- `-opti` turns on the FileKey merger, as every lint stage of the real build does
- The run finds no errors and still rewrites the file, because WinappDebug writes the header fresh on every save whether or not any entry changed
