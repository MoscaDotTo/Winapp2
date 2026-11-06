# Diff Examples

Worked examples for [Diff](readme.md). The [Common Uses](readme.md#common-uses) table in the readme lists the commands; these show what each one does, with real input and output.

# Table of Contents

1. [Example 1: Default Workflow](#example-1-default-workflow)
2. [Example 2: Full Database Comparison](#example-2-full-database-comparison)
3. [Example 3: Comparing Archived Versions](#example-3-comparing-archived-versions)
4. [Example 4: Verbose Mode](#example-4-verbose-mode)
5. [Example 5: Saving the Log](#example-5-saving-the-log)
6. [Example 6: Diffing a Non-Default Flavor](#example-6-diffing-a-non-default-flavor)
7. [Example 7: Automated Changelogs (CI/CD)](#example-7-automated-changelogs-cicd)

---

###### Note: All outputs below are lifted from real tool runs. Examples 2–5 compare base-flavor archives extracted from this repository (`Non-CCleaner/Winapp2.ini` at commit `f5ea7371` for v220510, and at tags `v251002` / `v251109`). Example 1 is a live trimmed run and depends on the software installed on the machine performing the diff. `[...]` marks lines cut for brevity.

## Example 1: Default Workflow

**Context:** Your local `winapp2.ini` is a trimmed copy from May 2022 (version 220510) and you want to see what has changed since.

**Command:**

```
winapp2ool -diff -base
```

**Output:**

```
Beginning Diff
Diff: version 220510 -> version 260915

 Web Browser additions

   Chromium Web Browser

   Google Chrome Web Browser

   Microsoft Edge Web Browser

   Mozilla Firefox Web Browser

   Opera Web Browser

 5 web browsers added

 Entries removed without replacement:

     Accounts Control * has been removed

     ActiveSync * has been removed

     ART * has been removed
[...]

 - 167 removed without replacement

 Entries renamed without other changes:

     MS Visual Studio Backups * has been renamed to Microsoft Visual Studio Backups *

     MS Visual Studio Installation Packages * has been renamed to Microsoft Visual Studio Installation Packages *
[...]

     RegEdit * has been renamed to Windows Registry Editor *

     Recent Colors History * has been renamed to Windows Shell - Theme Colors History *

 10 entries renamed (name-only changes)

 Entries merged or split into other entries:

     .NET Framework Isolated Storage * has been merged into .NET Framework *

     Active Setup Temp Folder * has been merged into Windows Subsystems *
[...]

 205 entries merged or split into other entries
[...]

 Cross-Entry key movements:

   ASUS * moved 1 key:
       -> FileKey2=%ProgramData%\ASUS\*|*.log|RECURSE moved to [ASUS USB Charger Plus *]
[...]

 35 keys moved between entries (6 source entries)

 Modified entries:
[...]

     Discord * has been modified

     5 keys added:

       Added 5 FileKeys
             FileKey1=%AppData%\BetterDiscord|*.log
             FileKey2=%AppData%\BetterDiscord\temp|*|REMOVESELF
             FileKey6=%AppData%\Discord*\Crashpad\reports|*
             FileKey8=%AppData%\Discord*\module_data\crashlogs|*
             FileKey12=%LocalAppData%\Discord*\packages|*.nupkg

     1 key removed:

       Removed 1 FileKey
             FileKey10=%LocalAppData%\SquirrelTemp|*|REMOVESELF

     2 keys modified:

       Modified 2 FileKeys
[...]

 53 modified entries

 Added entries containing merged content:
[...]

   Chromium Caches * has been added
   This entry contains content originating from 13 removed entries:
     - Application Cache Extras *
     - Blob Storage *
     - Budget Database *
     - Code Cache *
[...]

 Total diff time: 00:00:00.6979621

 Diff Summary
   Net entry count change: 492 → 432 (-60)
    + 5 new browsers added
   Modified entries: 68
    + 317 added keys across 52 entries
    - 39 removed keys without replacement across 26 entries
    ~ 108 updated keys replaced 168 old keys across 41 entries
    ~ 35 keys moved from 6 entries into 8 entries
       + 35 keys moved into 8 added entries
    + 15 entries also received merged content from removed entries (see merged entries below)
       = 14 entries contain 55 keys carried over unchanged from merged sources
   Removed entries: 399
    @ 205 removed entries have been merged into other entries
       @ 31 merged into 15 modified entries
       + 181 merged into 102 added entries
       - 49 split entries lost 1379 keys that no entry kept
    & 27 removed entries have been renamed
       = 10 are name-only changes (no key differences)
       + 21 added keys across 4 entries
       - 10 removed keys across 10 entries
       ~ 26 updated keys replaced 31 old keys across 11 entries
    - 167 entries have been removed without replacement
   Added entries: 339
    @ 102 entries consolidate content from 181 removed entries
       + 73 entries contain 641 novel keys (not from merged sources)
       = 77 entries contain 491 keys carried over unchanged from merged sources
       ~ 82 entries contain 356 keys capturing 509 removed keys
       - 45 entries dropped 395 keys from merged sources
    + 210 novel entries (without merged content)
    & 27 added entries are renamed versions of removed entries and may contain other minor changes

Diff complete
```

**Explanation:**

- `-base` selects the base flavor
- Diff downloaded the latest winapp2.ini and trimmed it for the current system before comparing
- Only entries for software detected on this machine appear in the output
- The header names the version that was current on GitHub at the time of the run, here 260915. The trimmed file keeps its version line
- `Net entry count change: 492 → 432 (-60)`: the new file has fewer *total* entries for this machine
- Many entries were consolidated: 205 of the 399 removed entries were mergers
- All 35 moved keys landed in added entries, eg. `AVG *` gave 19 keys to the new `AVG AntiVirus *`

**Notes:** The exact output and counts depend on what software is installed. Running with `-donttrim` shows the full database diff regardless of installed software; see Example 2. If you've installed something since your local copy was trimmed, its entry can show up as Added even though it's been in the database all along.

---

## Example 2: Full Database Comparison

**Context:** You want to see all changes across the entire database, not just entries relevant to your system.

**Command:**

```
winapp2ool -diff -d -1f winapp2-220510.ini -2f winapp2-251109.ini
```

**Output:**

```
Beginning Diff
Diff: version 220510 -> version 251109

 Web Browser additions

   .360 Secure Browser Web Browser
[...]

 49 web browsers added

 Entries removed without replacement:

     2K Launcher * has been removed
[...]

 - 526 removed without replacement

 Entries renamed without other changes:

     3delite Filesystem Dialogs * has been renamed to 3delite Filesystem Dialogs Library *
[...]

 112 entries renamed (name-only changes)

 Entries merged or split into other entries:
[...]

     Abelssoft GoogleClean * has been merged into Abelssoft GClean *
[...]

     Adobe Reader DC * has been split/merged into 2 entries
       • Adobe Acrobat *
       • Adobe Acrobat Reader *
[...]

 547 entries merged or split into other entries
[...]

 Total diff time: 00:00:03.2057302

 Diff Summary
   Net entry count change: 3257 → 3715 (+458)
    + 49 new browsers added
   Modified entries: 460
    + 486 added keys across 196 entries
    - 344 removed keys without replacement across 163 entries
    ~ 388 updated keys replaced 603 old keys across 220 entries
    ~ 100 keys moved from 41 entries into 24 entries
       ~ 2 keys moved into 1 modified entry
       & 5 keys moved into 1 renamed entry
       + 93 keys moved into 22 added entries
    + 135 entries also received merged content from removed entries (see merged entries below)
       = 118 entries contain 365 keys carried over unchanged from merged sources
   Removed entries: 1227
    @ 547 removed entries have been merged into other entries
       @ 175 merged into 135 modified entries
       + 382 merged into 316 added entries
       - 77 split entries lost 1481 keys that no entry kept
    & 154 removed entries have been renamed
       = 112 are name-only changes (no key differences)
       + 25 added keys across 12 entries
       - 23 removed keys across 21 entries
       ~ 66 updated keys replaced 78 old keys across 30 entries
    - 526 entries have been removed without replacement
   Added entries: 1685
    @ 316 entries consolidate content from 382 removed entries
       + 174 entries contain 1078 novel keys (not from merged sources)
       = 222 entries contain 1246 keys carried over unchanged from merged sources
       ~ 258 entries contain 935 keys capturing 1278 removed keys
       - 84 entries dropped 259 keys from merged sources
    + 1215 novel entries (without merged content)
    & 154 added entries are renamed versions of removed entries and may contain other minor changes

Diff complete
```

**Explanation:**

- This is the complete changelog for three and a half years of the database, May 2022 to November 2025
- Diff can compare any two versions of winapp2.ini, however far apart

---

## Example 3: Comparing Archived Versions

**Context:** You keep an archive of winapp2.ini releases and want a month-over-month changelog between two of them, without downloading anything.

**Command:**

```
winapp2ool -diff -d -1d C:\archive -1f winapp2-251002.ini -2d C:\archive -2f winapp2-251109.ini
```

**Output:**

```
Beginning Diff
Diff: version 251002 -> version 251109

 Web Browser additions

   BriskBard Web Browser

   Ghost Web Browser

   Helium Web Browser

   Lunascape Phoebe Web Browser

   Midori Web Browser

   Mullvad Web Browser

   Norton Neo Web Browser

   Wavebox Web Browser

 8 web browsers added

 Entries removed without replacement:

     Abylon Protection Manager * has been removed
[...]

 Total diff time: 00:00:00.2155657

 Diff Summary
   Net entry count change: 3519 → 3715 (+196)
    + 8 new browsers added
   Modified entries: 81
    + 150 added keys across 55 entries
    - 242 removed keys without replacement across 17 entries
    ~ 53 updated keys replaced 54 old keys across 33 entries
    ~ 18 keys moved from 6 entries into 7 entries
       ~ 1 key moved into 1 modified entry
       + 17 keys moved into 6 added entries
    + 1 entry also received merged content from removed entries (see merged entries below)
       = 1 entry contains 1 key carried over unchanged from merged sources
   Removed entries: 313
    @ 17 removed entries have been merged into other entries
       @ 1 merged into 1 modified entry
       + 17 merged into 34 added entries
       - 4 split entries lost 29 keys that no entry kept
    & 88 removed entries have been renamed
       = 83 are name-only changes (no key differences)
       + 13 added keys across 2 entries
       ~ 5 updated keys replaced 5 old keys across 4 entries
    - 208 entries have been removed without replacement
   Added entries: 509
    @ 34 entries consolidate content from 17 removed entries
       + 29 entries contain 225 novel keys (not from merged sources)
       = 34 entries contain 132 keys carried over unchanged from merged sources
       ~ 23 entries contain 39 keys capturing 39 removed keys
       - 3 entries dropped 5 keys from merged sources
    + 387 novel entries (without merged content)
    & 88 added entries are renamed versions of removed entries and may contain other minor changes

Diff complete
```

**Explanation:**

- `-d` disables downloading; both files must exist locally
- `-1d`/`-1f` set the "old" file's directory and name; `-2d`/`-2f` set the "new" file's. Here, both live in `C:\archive` rather than the working directory
- The rename block has no `- N removed keys` line because no renamed entry lost a key

---

## Example 4: Verbose Mode

**Context:** The default output shows only the changed keys of each entry. You want the full text of the affected entries in context.

**Command:**

```
winapp2ool -diff -d -1f winapp2-251002.ini -2f winapp2-251109.ini -verbose
```

**Output:**

A renamed entry prints both its old and new full text:

```
     360 Secure Browser Autofill Data * has been renamed to 360 Secure Browser Autofill Data & Search Engine Preferences *


       Old entry:
           [360 Secure Browser Autofill Data *]
           Section=.360 Secure Browser Web Browser
           DetectFile=%AppData%\360se6\User Data
           FileKey1=%AppData%\360se6\User Data\*|*Web Data
           FileKey2=%AppData%\360se6\User Data\*\AutoFill*|*|REMOVESELF
           FileKey3=%AppData%\360se6\User Data\AutoFill*|*|REMOVESELF


   Renamed entry:
           [360 Secure Browser Autofill Data & Search Engine Preferences *]
           Section=.360 Secure Browser Web Browser
           DetectFile=%AppData%\360se6\User Data
           FileKey1=%AppData%\360se6\User Data\*|*Web Data
           FileKey2=%AppData%\360se6\User Data\*\AutoFill*|*|REMOVESELF
           FileKey3=%AppData%\360se6\User Data\AutoFill*|*|REMOVESELF
```

A modified entry prints its new full text beneath the header, followed by the key diff:

```
     4Team Sync2 * has been modified
           [4Team Sync2 *]
           LangSecRef=3022
           DetectFile=%AppData%\4Team\Sync2
           FileKey1=%AppData%\4Team\Sync2|*.log;*.err
           FileKey2=%AppData%\4Team\Sync2\Logs|*
           FileKey3=%LocalAppData%\4Team\Sync2Cloud|*.log


     2 keys added:

       Added 2 FileKeys
             FileKey2=%AppData%\4Team\Sync2\Logs|*
             FileKey3=%LocalAppData%\4Team\Sync2Cloud|*.log

     1 key modified:

       Modified 1 FileKey

       FileKey1 has been modified, replacing 1 old key
              + New: FileKey1=%AppData%\4Team\Sync2|*.log;*.err
              - Old: FileKey1=%AppData%\4Team\Sync2|*.log|RECURSE
```

**Explanation:**

- `-verbose` (or the `Toggle verbose mode` menu option) prints the full text of each changed entry alongside its change line: removed entries show their old text, renamed and split/merged entries show both old and new text, and modified entries show their new text
- Verbose output is a lot longer

---

## Example 5: Saving the Log

**Context:** You want to archive the diff output so you can read the Source Entry Key Status Report.

**Command:**

```
winapp2ool -diff -d -1f winapp2-220510.ini -2f winapp2-251109.ini -savelog
```

**Output:** The same console output as Example 2. When the run finishes, Diff writes it to `diff.txt` in the current directory, and the Diff menu header says:

```
diff.txt saved
```

In `diff.txt`, the Source Entry Key Status Report sits between the Added entries section and the Diff Summary:

```
 1215 novel entries added
 SOURCE ENTRY KEY STATUS REPORT:
 ✓ = Captured, ✗ = Dropped

 [Abelssoft GoogleClean *] - 5 captured, 0 dropped
     ✓ [CATEGORY] LangSecRef=3024
     ✓ [DETECTION] DetectFile1=%LocalAppData%\Abelssoft\GClean
     ✓ [DETECTION] DetectFile2=%LocalAppData%\Abelssoft\GoogleClean
     ✓ [DELETION] FileKey1=%LocalAppData%\Abelssoft\GClean|*.log
     ✓ [DELETION] FileKey2=%LocalAppData%\Abelssoft\GoogleClean\log|*|REMOVESELF

 [Abelssoft PCFresh Backups *] - 3 captured, 0 dropped
     ✓ [CATEGORY] LangSecRef=3024
     ✓ [DETECTION] DetectFile=%LocalAppData%\Abelssoft\PCFresh
     ✓ [DELETION] FileKey1=%LocalAppData%\Abelssoft\PCFresh\Backup|*
[...]

 KEY STATUS SUMMARY:                 Total  Captured  Dropped  Capture Rate
 DELETION  (FileKey/RegKey):          2511      1825      686         72.7%
 DETECTION (Detect/DetectFile):       1251       273      978         21.8%
 CATEGORY  (Section/LangSecRef):       382       164      218         42.9%
 OTHER     (Warning etc.):              92        18       74         19.6%
 All keys:                            4236      2280     1956         53.8%
```

**Explanation:**

- `-savelog` turns on log saving, and `diff.txt` in the current directory is the default target
- The saved file holds the full diff output, including the Source Entry Key Status Report and the summary block
- Use **Log Viewer** in the Diff menu to review the most recent run's output without re-running

**Notes:**

- To save to a different file name: `winapp2ool -diff -savelog -3f changelog.txt`
- To save to a different directory: `winapp2ool -diff -savelog -3d C:\logs`

---

## Example 6: Diffing a Non-Default Flavor

**Context:** You maintain a BleachBit installation and want to know what changed in the BleachBit build of winapp2.ini since your copy.

**Command:**

```
winapp2ool -diff -bb -donttrim
```

**Explanation:**

- `-bb` selects the BleachBit flavor for the download
- `-donttrim` disables the trim step

---

## Example 7: Automated Changelogs (CI/CD)

**Context:** You want a build pipeline or scheduled task to generate a changelog with no interactive console session.

**Command:**

```
winapp2ool -s -diff -base -donttrim -savelog -3f changelog.txt
```

**Output:**

None on the console, since silent mode (`-s`) prints nothing.

The full diff log, identical in structure to the saved log in Example 5, is written to `changelog.txt`.

**Explanation:**

- `-s` suppresses all interactive output and prompts
- `-base` specifies the base flavor
- `-savelog -3f changelog.txt` writes the complete changelog to `changelog.txt`
- Diff fetches the latest winapp2.ini from GitHub automatically
- `-donttrim` disables the trim step

**Notes:** To diff two local files instead:

```
winapp2ool -s -diff -d -1f winapp2-old.ini -2f winapp2.ini -savelog -3f changelog.txt
```
