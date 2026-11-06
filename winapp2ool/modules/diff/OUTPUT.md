# Reading Diff's Output

How to read the changelog Diff produces: what each change category means, with real samples, and a full sample of Diff's output. See the [Diff readme](readme.md) for running it.

# Table of Contents

1. [Diff Categories](#diff-categories)
   - [Added Entries](#added-entries)
   - [Modified Entries](#modified-entries)
   - [Removed Entries](#removed-entries)
2. [Sample Output](#sample-output)

---

# Diff Categories

###### Note: The output excerpts in this section are taken verbatim from the saved log of a real Diff between base-flavor database versions 220510 (May 2022) and 251109 (November 2025). To reproduce them, extract `Non-CCleaner/Winapp2.ini` at commit `f5ea7371` and at tag `v251109`, then diff the two locally.

## Added Entries

An entry is **Added** if it exists in the new file and has no entry of the same name in the old file. Added entries render in two separate sections of the output: entries which consolidate content from removed entries appear first (under `Added entries containing merged content:`), followed by entirely novel entries (under `Added entries:`). Each section reports its own total count.

### Novel entries

Entirely new entries with no connection to any removed entry appear as a single line:

```
     .NET Platform * has been added

     360 Secure Browser Autofill Data & Search Engine Preferences * has been added
```

### Entries consolidating removed entries

When a new entry contains keys from one or more removed entries, its header is followed by a `This entry contains content originating from N removed entries:` line listing the source entries. Diff then sorts the keys of the new entry and its sources into these categories:

| Category | Type | Annotation | Conditions |
|:-|:-|:-|:-|
| **Novel** | Added keys | `(novel)` in an `N keys added:` block | key is new content with no equivalent in any source entry |
| **Moved in** | Added keys | `(from [Other *])` in an `N keys added:` block | key [moved](#cross-entry-key-movements) here from an entry that still exists |
| **Carried over** | Carried-over keys | `(from [Source *])` in an `N keys carried over from merged entries:` block | key is carried over verbatim from the source entry |
| **Dropped** | Removed keys | Appears in an `N keys from merged entries not in this entry:` block | the key existed in a source entry merged into this entry alone, and this entry doesn't have it, unchanged or captured. A split source's lost keys are listed under the split. |
| **Captured** | Modified keys | `(from [Source *])` on the old key | the key does not appear verbatim but is replaced by another key that covers it — typically a wildcard path that absorbs one or more specific old paths |

A key that several source entries share is credited to each of them in name order, eg. `(from [Source A *], [Source B *])`. Past three sources, Diff names the first three and counts the rest: `(from [Source A *], [Source B *], [Source C *] and 4 others)`.

**Example 1: Novel and carried keys (Abelssoft PCFresh):**

**Old (v220510):**
```ini
[Abelssoft PCFresh Backups *]
LangSecRef=3024
DetectFile=%LocalAppData%\Abelssoft\PCFresh
FileKey1=%LocalAppData%\Abelssoft\PCFresh\Backup|*.*
```

**New (v251109):**
```ini
[Abelssoft PCFresh *]
LangSecRef=3024
DetectFile=%LocalAppData%\Abelssoft\PCFresh
FileKey1=%LocalAppData%\Abelssoft\PCFresh|*.log
FileKey2=%LocalAppData%\Abelssoft\PCFresh\Backup|*
```

**Diff output:**
```
   Abelssoft PCFresh * has been added
   This entry contains content originating from 1 removed entry:
     - Abelssoft PCFresh Backups *

   1 key added:

       Added 1 FileKey
             FileKey1=%LocalAppData%\Abelssoft\PCFresh|*.log (novel)

   3 keys carried over from merged entries:

       Carried over 1 LangSecRef
       Carried over 1 DetectFile
       Carried over 1 FileKey
             LangSecRef=3024 (from [Abelssoft PCFresh Backups *])
             DetectFile=%LocalAppData%\Abelssoft\PCFresh (from [Abelssoft PCFresh Backups *])
             FileKey2=%LocalAppData%\Abelssoft\PCFresh\Backup|* (from [Abelssoft PCFresh Backups *])
```

**Explanation:**

`FileKey1` is novel; no equivalent exists in `[Abelssoft PCFresh Backups *]`.

`LangSecRef`, `DetectFile`, and `FileKey2` were carried over unchanged, so they're listed in their own group, with attribution to their source entry.

Note that `FileKey2`'s pattern changed from `*.*` to `*`, and that the old key is printed as `|*`. Diff rewrites deprecated values in the old file before comparing, so the key counts as carried over rather than modified. See [Value Normalization](readme.md#value-normalization).

**Example 2: Captured keys (EFSoftware):**

**Old (v220510):**

```ini
[EF Duplicate Files Manager *]
LangSecRef=3024
DetectFile=%AppData%\EFSoftware\DFMCache
FileKey1=%AppData%\EFSoftware|DFMCache

[EF Duplicate MP3 Finder *]
LangSecRef=3024
DetectFile=%AppData%\EFSoftware\MP3Cache
FileKey1=%AppData%\EFSoftware|MP3Cache
```

**New (v251109):**

```ini
[EFSoftware  *]
LangSecRef=3024
DetectFile=%AppData%\EFSoftware
FileKey1=%AppData%\EFSoftware|*Cache
```

**Diff output:**

```
   EFSoftware  * has been added
   This entry contains content originating from 2 removed entries:
     - EF Duplicate Files Manager *
     - EF Duplicate MP3 Finder *

   1 key carried over from merged entries:

       Carried over 1 LangSecRef
             LangSecRef=3024 (from [EF Duplicate Files Manager *], [EF Duplicate MP3 Finder *])

   2 keys capturing content from merged entries:

       Modified 1 Detection criteria
       Modified 1 FileKey

       Detection criteria modified, replacing 2 old keys
              + New: DetectFile=%AppData%\EFSoftware
              - Old: DetectFile=%AppData%\EFSoftware\DFMCache (from [EF Duplicate Files Manager *])
              - Old: DetectFile=%AppData%\EFSoftware\MP3Cache (from [EF Duplicate MP3 Finder *])
       FileKey1 has been modified, replacing 2 old keys
              + New: FileKey1=%AppData%\EFSoftware|*Cache
              - Old: FileKey1=%AppData%\EFSoftware|DFMCache (from [EF Duplicate Files Manager *])
              - Old: FileKey1=%AppData%\EFSoftware|MP3Cache (from [EF Duplicate MP3 Finder *])
```

**Explanation:**

`LangSecRef` is carried over and credited to both source entries.

The `DetectFile` and `FileKey1` in the new entry each cover both old paths with a single pattern; `%AppData%\EFSoftware` is a parent path, and `*Cache` matches both old parameters. Both are reported as captured (modified) keys replacing 2 old keys each.

**Example 3: Dropped keys (Cook'n):**

Three related Cook'n entries were consolidated into a single `Cook'n *` entry, but the `Warning=` key from `Cook'n Dups *` was not carried over or replaced.

**Old (v220510):**

```ini
[Cook'n Cache *]
LangSecRef=3021
Detect=HKLM\Software\Microsoft\Windows\CurrentVersion\Uninstall\Cook'n
FileKey1=%AppData%\Mozilla\eclipse\Cache|*.*

[Cook'n Dups *]
LangSecRef=3021
Detect=HKLM\Software\Microsoft\Windows\CurrentVersion\Uninstall\Cook'n
Warning=This removes identical duplicates of the huge Getting Started Guide.
FileKey1=%LocalAppData%\DVO\Cook'n10App\plugins|Getting Started Guide.rtf|RECURSE

[Cook'n Logs *]
LangSecRef=3021
Detect=HKLM\Software\Microsoft\Windows\CurrentVersion\Uninstall\Cook'n
FileKey1=%Documents%\Cook'n*|*.log|RECURSE
FileKey2=%LocalAppData%\DVO\Cook'n*App|*.log|RECURSE
```

**New (v251109):**

```ini
[Cook'n *]
LangSecRef=3021
Detect=HKLM\Software\Microsoft\Windows\CurrentVersion\Uninstall\Cook'n
FileKey1=%AppData%\Mozilla\eclipse\Cache|*
FileKey2=%LocalAppData%\DVO\Cook'n*App|*.log|RECURSE
FileKey3=%LocalAppData%\DVO\Cook'n10App\plugins|Getting Started Guide.rtf|RECURSE
FileKey4=%UserProfile%\Documents\Cook'n*|*.log|RECURSE
```

**Diff output:**

```
   Cook'n * has been added
   This entry contains content originating from 3 removed entries:
     - Cook'n Cache *
     - Cook'n Dups *
     - Cook'n Logs *

   6 keys carried over from merged entries:

       Carried over 1 LangSecRef
       Carried over 1 Detect
       Carried over 4 FileKeys
             LangSecRef=3021 (from [Cook'n Cache *], [Cook'n Dups *], [Cook'n Logs *])
             Detect=HKLM\Software\Microsoft\Windows\CurrentVersion\Uninstall\Cook'n (from [Cook'n Cache *], [Cook'n Dups *], [Cook'n Logs *])
             FileKey1=%AppData%\Mozilla\eclipse\Cache|* (from [Cook'n Cache *])
             FileKey2=%LocalAppData%\DVO\Cook'n*App|*.log|RECURSE (from [Cook'n Logs *])
             FileKey3=%LocalAppData%\DVO\Cook'n10App\plugins|Getting Started Guide.rtf|RECURSE (from [Cook'n Dups *])
             FileKey4=%UserProfile%\Documents\Cook'n*|*.log|RECURSE (from [Cook'n Logs *])

   1 key from merged entries not in this entry:

       Dropped 1 Warning
             Warning=This removes identical duplicates of the huge Getting Started Guide. (from [Cook'n Dups *])
```

**Explanation:**

Every `FileKey` from all three source entries was carried over, attributed to its originating entry.

Keys that all three source entries shared (`LangSecRef`, `Detect`) are credited to all three.

The `Warning=` key from `Cook'n Dups *` was dropped, and is listed under the dropped heading with attribution to its originating entry.

The change from `%Documents%` to `%UserProfile%\Documents` is a deprecated-path rewrite, so `FileKey4` matches `Cook'n Logs *`'s `FileKey1` despite the textual difference. See [Value Normalization](readme.md#value-normalization).

---

## Modified Entries

An entry is **Modified** if it exists in both files with the same name but has changed. Modifications include added keys, removed keys, and changed key values or parameters. Diff itemizes the specific key-level changes for each modified entry.

Keys annotated `(novel)` are new content with no detected equivalent in the old file. Keys annotated `(from [Entry Name *])` were sourced from a specific entry: the old version of the same entry, a removed entry that was merged in, or an entry the key [moved](#cross-entry-key-movements) from.

Modified output itemizes five kinds of key-level change:

- **Added keys**: present in the new entry with no detected equivalent in the old entry
- **Removed keys**: present in the old entry with no detected equivalent in the new entry or anywhere else
- **Modified keys**: a key whose value changed, shown as `X has been modified, replacing N old keys`; in the simple case one old key was updated to a new value; in more complex cases a single new key uses a wildcard or semicolon-joined pattern to capture multiple old keys
- **Moved keys**: keys that left this entry for another one, listed under `N keys moved to other entries:` as `FileKey2=... moved to [Other *]`. See [Cross-Entry Key Movements](#cross-entry-key-movements)
- **Moved-in keys**: keys this entry took from another entry that still exists, listed under `N keys moved in from other entries:` as `FileKey2=... (from [Other *])`. An entry that also absorbed removed entries lists these among its added keys instead, with the same credit. 
 ###### Note: A moved-in key that the block shows as an update or capture of a merged source's key credits only the source, so the key is also listed under `N keys moved in from other entries:`

Each kind has its own list, under a heading with its total (`13 keys added:`, `2 keys removed:`, `7 keys modified:`), followed by a count line for each key type (`Added 2 DetectFiles`). An entry whose only change is a moved key is still itemized.

Changes to detection keys (`Detect`, `DetectFile`) are grouped and reported together as `Detection criteria`.

Modified entries are displayed in two separate output sections: entries which absorbed content from removed entries appear first (each followed by a note itemizing the removed entries against which the changes are measured), and entries with ordinary modifications appear later under the `Modified entries:` header.

### Simple modification

**Example: Modified keys and detection criteria (Facebook):**

**Old (v220510):**

```ini
[Facebook *]
LangSecRef=3031
Detect=HKCU\Software\Classes\Local Settings\Software\Microsoft\Windows\CurrentVersion\AppModel\SystemAppData\Facebook.Facebook_8xx8rvfyw5nnt
FileKey1=%LocalAppData%\Packages\Facebook.Facebook_*\AC\INet*|*.*|RECURSE
FileKey2=%LocalAppData%\Packages\Facebook.Facebook_*\AC\Microsoft\CryptnetUrlCache\*|*.*|RECURSE
FileKey3=%LocalAppData%\Packages\Facebook.Facebook_*\AC\Temp|*.*|RECURSE
```

**New (v251109):**

```ini
[Facebook *]
LangSecRef=3022
DetectFile=%LocalAppData%\Packages\Facebook.Facebook_*
FileKey1=%LocalAppData%\Packages\Facebook.Facebook_*\AC|*|RECURSE
```

**Diff output:**

```
     Facebook * has been modified

     3 keys modified:

       Modified 1 LangSecRef
       Modified 1 FileKey
       Modified 1 Detection criteria

       LangSecRef has been modified, replacing 1 old key
              + New: LangSecRef=3022
              - Old: LangSecRef=3031
       FileKey1 has been modified, replacing 3 old keys
              + New: FileKey1=%LocalAppData%\Packages\Facebook.Facebook_*\AC|*|RECURSE
              - Old: FileKey1=%LocalAppData%\Packages\Facebook.Facebook_*\AC\INet*|*|RECURSE
              - Old: FileKey2=%LocalAppData%\Packages\Facebook.Facebook_*\AC\Microsoft\CryptnetUrlCache\*|*|RECURSE
              - Old: FileKey3=%LocalAppData%\Packages\Facebook.Facebook_*\AC\Temp|*|RECURSE
       Detection criteria modified
              + New: DetectFile=%LocalAppData%\Packages\Facebook.Facebook_*
              - Old: Detect=HKCU\Software\Classes\Local Settings\Software\Microsoft\Windows\CurrentVersion\AppModel\SystemAppData\Facebook.Facebook_8xx8rvfyw5nnt
```

**Explanation:**

The updated `LangSecRef` is identified as modified.

The updated `FileKey1` is identified as modified, replacing 3 old keys. The new `\AC|*|RECURSE` path captures the three separate subdirectory paths from the old entry.

The replacement of `Detect` with `DetectFile` is paired into a single `Detection criteria modified` notification.

### Modification by merger

When entries from the old file are merged into a modified entry, their absorbed keys appear alongside the entry's own changes. The `(from [Entry Name *])` annotation on each old key identifies the old entry or entries from which it came. Keys labeled with the entry's own name came from its prior version; keys labeled with a different name came from a removed entry that was consolidated. Keys the entry's prior version had and lost are listed as removed; keys that only a merged entry had and that didn't survive are listed separately, under `N keys from merged entries not in this entry:`, as in an added entry's merger block. When a removed entry was split across several entries, its keys aren't listed as dropped here. A key that went to one of the others is counted on a closing line that names where it went, eg. `2 other keys from merged entries went to [Windows Font Cache *]`. A key that no entry kept is listed once, in the removed entry's own section under [Merged](#merged), and counted on a second closing line, eg. `173 other keys that no entry kept are listed under [Browser Metrics *], [Crash Reports *], [Download History *] and 8 others`. Each such block opens with the same `This entry contains content originating from N removed entries:` list as an added entry's, right under its header.

In this example, two version-specific log path entries were merged into a single entry using a wildcard to cover both:

**Old (v220510):**

```ini
[Age of Wonders II *]
Section=Games
DetectFile=%ProgramFiles%\Age of Wonders II
FileKey1=%ProgramFiles%\Age of Wonders II|*aow2Log.txt
FileKey2=%ProgramFiles%\Age of Wonders II\Resource\FX\Spell|*.lnk

[Age of Wonders: Shadow Magic *]
Section=Games
DetectFile=%ProgramFiles%\Age of Wonders II
FileKey1=%ProgramFiles%\Age of Wonders Shadow Magic|*aowsmLog.txt
FileKey2=%ProgramFiles%\Age of Wonders Shadow Magic\Resource\Maps|*.bak
```

**New (v251109):**

```ini
[Age of Wonders II *]
Section=Games
DetectFile=%ProgramFiles%\Age of Wonders II
FileKey1=%ProgramFiles%\Age of Wonders II\Resource\FX\Spell|*.lnk
FileKey2=%ProgramFiles%\Age of Wonders Shadow Magic\Resource\Maps|*.bak
FileKey3=%ProgramFiles%\Age of Wonders*|*Log.txt
```

**Diff output:**

```
     Age of Wonders: Shadow Magic * has been merged into Age of Wonders II *
[...]

     Age of Wonders II * has been modified
     This entry contains content originating from 1 removed entry:
       - Age of Wonders: Shadow Magic *

     1 key carried over from merged entries:

       Carried over 1 FileKey
             FileKey2=%ProgramFiles%\Age of Wonders Shadow Magic\Resource\Maps|*.bak (from [Age of Wonders: Shadow Magic *])

     1 key modified:

       Modified 1 FileKey

       FileKey3 has been modified, replacing 2 old keys
              + New: FileKey3=%ProgramFiles%\Age of Wonders*|*Log.txt
              - Old: FileKey1=%ProgramFiles%\Age of Wonders Shadow Magic|*aowsmLog.txt (from [Age of Wonders: Shadow Magic *])
              - Old: FileKey1=%ProgramFiles%\Age of Wonders II|*aow2Log.txt (from [Age of Wonders II *])
```

**Explanation:**

The merger itself is reported in the merged-entries list

The key-level detail appears in the modified-with-merged-content section.

Shadow Magic's `Resource\Maps` key survives unchanged as `FileKey2`, so it is listed as carried over and credited to the removed entry.

The `FileKey3` wildcard captures both games' log paths into a single key, one from the removed entry and one from this entry's own prior version.

### Renames with key changes

When an entry is renamed and also has other key changes, it does not appear in the Renamed list (which contains only name-only renames). Instead, it appears in a dedicated `Minor changes to renamed entries:` section, displayed as `Old Name * has been renamed to New Name *` alongside its key changes.

In this example, `IDLE Recent Files *` was renamed to `IDLE *` and simultaneously recategorized from `3021` to `3024`:

**Old (v220510):**

```ini
[IDLE Recent Files *]
LangSecRef=3021
Detect=HKCU\Software\Python
FileKey1=%UserProfile%\.idlerc|recent-files.lst
```

**New (v251109):**

```ini
[IDLE *]
LangSecRef=3024
Detect=HKCU\Software\Python
FileKey1=%UserProfile%\.idlerc|recent-files.lst
```

**Diff output:**

```
     IDLE Recent Files * has been renamed to IDLE *

     1 key modified:

       Modified 1 LangSecRef

       LangSecRef has been modified, replacing 1 old key
              + New: LangSecRef=3024
              - Old: LangSecRef=3021
```

**Explanation:**

The `FileKey` and `Detect` content of `[IDLE Recent Files *]` was found unchanged in `[IDLE *]`, so the entry is classified as renamed rather than removed.

`LangSecRef`'s value also changed, so the entry appears in the `Minor changes to renamed entries:` section with its key diffs, rather than in the name-only rename list.

### Cross-Entry Key Movements

A key that leaves an entry present in both files and lands in any other entry has moved. Moves are reported in a dedicated `Cross-Entry key movements:` section, grouped by the entry the keys left, each group headed with its total, eg. `MP3Gain * moved 2 keys:`. The receiving entry can be modified, renamed, or added, so splitting part of an entry out into a new one shows up here. Unlike with a merger, the source entry still exists.

The section is omitted entirely when no movements are detected.

**Diff output:**

```
 Cross-Entry key movements:
[...]

   CleanMyPC Registry Cleaner * moved 1 key:
       -> FileKey2=%WinDir%\$regcmp$|*|REMOVESELF moved to [Windows Subsystems *]
[...]

   MP3Gain * moved 2 keys:
       -> RegKey1=HKCU\Software\VB and VBA Program Settings\MP3GainAnalysis\StartUp|AddFilesPath moved to [MP3Gain StartUp *]
       -> RegKey2=HKCU\Software\VB and VBA Program Settings\MP3GainAnalysis\StartUp|AddFolderPath moved to [MP3Gain StartUp *]
[...]

   Soulseek * moved 5 keys:
       -> FileKey1=%UserProfile%\Documents\Soulseek Chat Logs\Private|*.txt moved to [Soulseek Chat History *]
       -> FileKey2=%LocalAppData%\Soulseek Chat Logs\*|*.log moved to [Soulseek Chat History *]
       -> FileKey3=%LocalAppData%\SoulseekQt\Soulseek Chat Logs|*.log|RECURSE moved to [Soulseek Chat History *]
       -> FileKey5=%ProgramFiles%\Soulseek*\Private_Chat_Logs|*.txt moved to [Soulseek Chat History *]
       -> FileKey6=%ProgramFiles%\Soulseek*\Room_Chat_Logs|*.txt moved to [Soulseek Chat History *]
[...]

   Windows Subsystems * moved 13 keys:
       -> FileKey1=%ProgramData%\Microsoft\PlayReady|*.hds moved to [Microsoft PlayReady *]
       -> FileKey2=%ProgramData%\Microsoft\PlayReady\Cache|* moved to [Microsoft PlayReady *]
       -> FileKey6=%ProgramData%\Microsoft\Windows\DRM|*.log moved to [Microsoft PlayReady *]
       -> FileKey7=%ProgramData%\Microsoft\Windows\DRM\Cache|*|RECURSE moved to [Microsoft PlayReady *]
       -> FileKey8=%ProgramData%\Microsoft\Windows\DRM\PreUpgrade|*.log moved to [Microsoft PlayReady *]
       -> FileKey11=%LocalAppData%\Microsoft\Windows\PRICache|*|RECURSE moved to [Universal Windows Platform *]
       -> FileKey14=%SystemDrive%\spoolerlogs|spooler.xml moved to [Windows Printing *]
       -> FileKey17=%WinDir%\System32\spool|spooler.xml moved to [Windows Printing *]
       -> RegKey1=HKCU\Software\Microsoft\Direct3D\MostRecentApplication moved to [Microsoft DirectX *]
       -> RegKey2=HKLM\Software\Microsoft\Direct3D\MostRecentApplication moved to [Microsoft DirectX *]
       -> RegKey3=HKLM\Software\Microsoft\DirectDraw\MostRecentApplication moved to [Microsoft DirectX *]
       -> RegKey5=HKLM\Software\Wow6432Node\Microsoft\Direct3D\MostRecentApplication moved to [Microsoft DirectX *]
       -> RegKey6=HKLM\Software\Wow6432Node\Microsoft\DirectDraw\MostRecentApplication moved to [Microsoft DirectX *]
[...]

 100 keys moved between entries (41 source entries)
```

**Explanation:**

`[CleanMyPC Registry Cleaner *]` gave its `%WinDir%\$regcmp$` key to `[Windows Subsystems *]`, an entry that already existed.

`[MP3Gain *]` and `[Soulseek *]` were each split: their recent-file and chat-log keys went into new entries of their own. The last two Soulseek keys landed in one wildcard key, `%ProgramFiles%\Soulseek*\*Chat_Logs|*.txt`, rather than identical ones.

`[Windows Subsystems *]` gave keys to four entries. `[Microsoft PlayReady *]` is a renamed `[Media Play Ready Client *]`. The other three are new, and each also absorbed removed entries, so their merger output credits these keys `(from [Windows Subsystems *])`.

Every source entry still exists in the new file, which is what distinguishes a movement from a merger.

Each move also appears in the itemized output of both entries, under `Moved` in the source and `Moved in` in the receiving entry (see [Modified Entries](#modified-entries)):

```
     CleanMyPC Registry Cleaner * has been modified

     1 key moved to other entries:

       Moved 1 FileKey
             FileKey2=%WinDir%\$regcmp$|*|REMOVESELF moved to [Windows Subsystems *]
```

## Removed Entries

An entry is **Removed** if it exists in the old file but not in the new file. Removed entries fall into one of three sub-categories.

### Renamed

The entry no longer exists by its old name, but another entry in the new file contains its content mostly unchanged. Diff reports both the old and new names.

The Renamed list contains only renames with no key changes. Renames with accompanying key changes appear separately in the `Minor changes to renamed entries:` section.

**Sample output:**

```
     3delite Filesystem Dialogs * has been renamed to 3delite Filesystem Dialogs Library *

     4Sync Extras * has been renamed to 4Sync *

     Tracks Eraser Pro * has been renamed to Acesoft Tracks Eraser Pro *
```

The list closes with a count: `N entries renamed (name-only changes)`.

### Merged

The entry no longer exists, but its content was absorbed into another entry in the new file that is substantially different from the old version. Diff reports the name of the absorbing entry. Mergers are further categorized in the summary report:

- **Modified by merger**: the absorbing entry existed in the old file and was modified to include the removed entry's content
- **Added with merger**: the absorbing entry is itself new and contains content from one or more removed entries

**Sample output (merged into an existing entry):**

```
     .NET Application History * has been merged into .NET Framework *
[...]

     Adobe Reader XI * has been merged into Adobe Acrobat Reader *
[...]

     Common Language Runtime * has been merged into .NET Framework *
```

**Sample output (merged into a new entry):**

```
     Abelssoft GoogleClean * has been merged into Abelssoft GClean *
```

In the added-entries output, the entry's header is followed by `This entry contains content originating from 1 removed entry:` and the source's name.

**Example with before/after context:**

Two companion entries (`Clover Bookmarks Backups *` and `Clover Session *`) were consolidated into the main `Clover *` entry. Each contributed a single FileKey, which the new entry combines into a single pattern.

**Old (v220510):**

```ini
[Clover *]
LangSecRef=3024
Detect=HKCU\Software\Clover
FileKey1=%LocalAppData%\Clover\User Data\Default\JumpListIcons*|*tmp
FileKey2=%LocalAppData%\Clover\User Data\temp|*.*|RECURSE

[Clover Bookmarks Backups *]
LangSecRef=3024
Detect=HKCU\Software\Clover
FileKey1=%LocalAppData%\Clover\User Data\Default|Bookmarks.bak

[Clover Session *]
LangSecRef=3024
Detect=HKCU\Software\Clover
FileKey1=%LocalAppData%\Clover\User Data\Default|Current *;Last *
```

**New (v251109):**

```ini
[Clover *]
LangSecRef=3024
Detect=HKCU\Software\Clover
FileKey1=%LocalAppData%\Clover\User Data\Default|Bookmarks.bak;Current *;Last *
FileKey2=%LocalAppData%\Clover\User Data\Default\JumpListIcons*|*tmp
FileKey3=%LocalAppData%\Clover\User Data\temp|*|RECURSE
```

**Diff output:**

```
     Clover Bookmarks Backups * has been merged into Clover *

     Clover Session * has been merged into Clover *
[...]

     Clover * has been modified
     This entry contains content originating from 2 removed entries:
       - Clover Bookmarks Backups *
       - Clover Session *

     1 key modified:

       Modified 1 FileKey

       FileKey1 has been modified, replacing 2 old keys
              + New: FileKey1=%LocalAppData%\Clover\User Data\Default|Bookmarks.bak;Current *;Last *
              - Old: FileKey1=%LocalAppData%\Clover\User Data\Default|Bookmarks.bak (from [Clover Bookmarks Backups *])
              - Old: FileKey1=%LocalAppData%\Clover\User Data\Default|Current *;Last * (from [Clover Session *])
```

The two removed entries each contributed a FileKey, and the new version of `Clover *` combines their values into a single `FileKey1`.

Because `Clover *` is a *modified* entry rather than an added one, its source entries do **not** appear in the [Source Entry Key Status Report](readme.md#source-entry-key-status-report).

**Sample output (split across multiple entries):**

```
     Adobe Reader DC * has been split/merged into 2 entries
       • Adobe Acrobat *
       • Adobe Acrobat Reader *
```

A split/merged result means the old entry's content was distributed across more than one new entry, i.e. different keys matched in different places. The keys of a split entry that none of those entries kept, unchanged or captured, are listed once, here, under `N keys not kept by any of these entries:`. The merger blocks of the entries it went into don't list them as dropped; each ends with a line pointing back here instead (see [Modification by merger](#modification-by-merger)).

**Example with before/after context (large split):** The following shows how `Alternate Services *`, a single legacy entry covering nine different Firefox-based browsers, was split into six browser-specific entries. Each browser that received a dedicated entry now appears as a separate absorbing target, and the keys for the three browsers that got no entry are listed under the split.

**Old (v220510):**

```ini
[Alternate Services *]
LangSecRef=3026
Detect1=HKCU\Software\ArtistScope\ArtisBrowser
Detect2=HKCU\Software\Classes\Local Settings\Software\Microsoft\Windows\CurrentVersion\AppModel\SystemAppData\Mozilla.Firefox_n80bbvh6b1yt2
Detect3=HKCU\Software\LibreWolf
Detect4=HKLM\Software\ComodoGroup\IceDragon
Detect5=HKLM\Software\FlashPeak\SlimBrowser
Detect6=HKLM\Software\Mozilla\Basilisk
Detect7=HKLM\Software\Mozilla\Pale Moon
Detect8=HKLM\Software\Mozilla\SeaMonkey
Detect9=HKLM\Software\Mozilla\Waterfox
DetectFile=%AppData%\Mozilla\Firefox
FileKey1=%AppData%\ArtistScope\ArtisBrowser\Profiles\*|AlternateServices.txt
FileKey2=%AppData%\Comodo\IceDragon\Profiles\*|AlternateServices.txt
FileKey3=%AppData%\FlashPeak\SlimBrowser\Profiles\*|AlternateServices.txt
FileKey4=%AppData%\LibreWolf\Profiles\*|AlternateServices.txt
FileKey5=%AppData%\Moonchild Productions\Basilisk\Profiles\*|AlternateServices.txt
FileKey6=%AppData%\Moonchild Productions\Pale Moon\Profiles\*|AlternateServices.txt
FileKey7=%AppData%\Mozilla\Firefox\Profiles\*|AlternateServices.txt
FileKey8=%AppData%\Mozilla\SeaMonkey\Profiles\*|AlternateServices.txt
FileKey9=%AppData%\Waterfox\Profiles\*|AlternateServices.txt
FileKey10=%LocalAppData%\Packages\Mozilla.Firefox_*\LocalCache\Roaming\Mozilla\Firefox\Profiles\*|AlternateServices.txt
```

**Diff output (220510→251109):**

```
     Alternate Services * has been split/merged into 6 entries
       • ArtisBrowser Caches *
       • LibreWolf Caches *
       • Mozilla Firefox Caches *
       • Pale Moon Caches *
       • SeaMonkey Caches *
       • Waterfox Caches *
   
     14 keys not kept by any of these entries:

       Lost 1 LangSecRef
       Lost 9 Detects
       Lost 1 DetectFile
       Lost 3 FileKeys
             LangSecRef=3026
             Detect1=HKCU\Software\ArtistScope\ArtisBrowser
             Detect2=HKCU\Software\Classes\Local Settings\Software\Microsoft\Windows\CurrentVersion\AppModel\SystemAppData\Mozilla.Firefox_n80bbvh6b1yt2
             Detect3=HKCU\Software\LibreWolf
             Detect4=HKLM\Software\ComodoGroup\IceDragon
             Detect5=HKLM\Software\FlashPeak\SlimBrowser
             Detect6=HKLM\Software\Mozilla\Basilisk
             Detect7=HKLM\Software\Mozilla\Pale Moon
             Detect8=HKLM\Software\Mozilla\SeaMonkey
             Detect9=HKLM\Software\Mozilla\Waterfox
             DetectFile=%AppData%\Mozilla\Firefox
             FileKey2=%AppData%\Comodo\IceDragon\Profiles\*|AlternateServices.txt
             FileKey3=%AppData%\FlashPeak\SlimBrowser\Profiles\*|AlternateServices.txt
             FileKey5=%AppData%\Moonchild Productions\Basilisk\Profiles\*|AlternateServices.txt
```

The new entries detect their browsers with other keys (`LibreWolf Caches *`, for one, uses `DetectFile1=%AppData%\librewolf\Profiles`), so none of the old Detects survive verbatim. They are categorized by `Section=` rather than `LangSecRef=3026`, so the old category key is lost too.

Six of the nine browser FileKeys were matched in six new per-browser entries.

The remaining three (IceDragon, SlimBrowser, Basilisk) were not matched in any new entry.

### Removed Without Replacement

The entry is gone and its content does not appear in any other entry in the new file. These are listed individually under the `Entries removed without replacement:` header, closed by a count.

**Sample output:**

```
 Entries removed without replacement:

     2K Launcher * has been removed

     337 Wallpaper * has been removed

     3D Builder * has been removed
[...]

 - 526 removed without replacement
```

---

# Sample Output

The following is a trimmed extract of the saved log for a real diff run comparing base-flavor winapp2.ini version 220510 (May 2022) against version 251109 (November 2025). `[...]` marks sections cut for brevity.

```
Beginning Diff
Diff: version 220510 -> version 251109

 Web Browser additions

   .360 Secure Browser Web Browser

   AOL Shield Web Browser
[...]

 49 web browsers added

 Entries removed without replacement:

     2K Launcher * has been removed

     337 Wallpaper * has been removed
[...]

 - 526 removed without replacement

 Entries renamed without other changes:

     3delite Filesystem Dialogs * has been renamed to 3delite Filesystem Dialogs Library *

     4Sync Extras * has been renamed to 4Sync *
[...]

 112 entries renamed (name-only changes)

 Entries merged or split into other entries:

     .NET Application History * has been merged into .NET Framework *

     .NET Framework Isolated Storage * has been merged into .NET Framework *

     3D Viewer * has been merged into Microsoft 3D Viewer *
[...]

 547 entries merged or split into other entries

 Minor changes to renamed entries:

     Abelssoft SSD Fresh Backups * has been renamed to Abelssoft SSD Fresh *

     1 key modified:

       Modified 1 FileKey

       FileKey1 has been modified, replacing 1 old key
              + New: FileKey1=%LocalAppData%\Abelssoft\SSD Fresh|backup.*;*.log
              - Old: FileKey1=%LocalAppData%\Abelssoft\SSD Fresh|backup.*
[...]

 Minor changes to 42 renamed entries

 Modified entries incorporating merged content:

     .NET Framework * has been modified
     This entry contains content originating from 5 removed entries:
       - .NET Application History *
       - .NET Framework Isolated Storage *
       - Common Language Runtime *
       - Font Cache *
       - Presentation Foundation *
[...]

 135 modified entries incorporating merged content

 Cross-Entry key movements:

   Adobe Acrobat Distiller * moved 1 key:
       -> FileKey1=%AppData%\Adobe\Acrobat\Distiller*|*.log moved to [Adobe *]
[...]

 100 keys moved between entries (41 source entries)

 Modified entries:

     //N.P.P.D. RUSH// - The Milk of Ultra Violet * has been modified

     1 key modified:

       Modified 1 Detection criteria

       Detection criteria modified
              + New: DetectFile=%LocalAppData%\NPPDRUSH
              - Old: Detect=HKCU\Software\Valve\Steam\Apps\270090
[...]

 325 modified entries

 Added entries containing merged content:

   Abelssoft GClean * has been added
   This entry contains content originating from 1 removed entry:
     - Abelssoft GoogleClean *

   1 key added:

       Added 1 FileKey
             FileKey2=%LocalAppData%\Abelssoft\GClean\Backups|* (novel)

   3 keys carried over from merged entries:

       Carried over 1 LangSecRef
       Carried over 2 FileKeys
             LangSecRef=3024 (from [Abelssoft GoogleClean *])
[...]

 316 added entries consolidating removed content

 Added entries:

     .NET Platform * has been added

     360 Secure Browser Autofill Data & Search Engine Preferences * has been added
[...]

 1215 novel entries added
 SOURCE ENTRY KEY STATUS REPORT:
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
