# Downloader Examples

Worked examples for [Downloader](readme.md). The [Common Uses](readme.md#common-uses) table in the readme lists the commands; these show what each one does, with real input and output.

# Table of Contents

1. [Acquiring winapp2.ini](#acquiring-winapp2ini)
   - [Example 1: Downloading winapp2.ini](#example-1-downloading-winapp2ini)
   - [Example 2: Choosing a flavor without changing your settings](#example-2-choosing-a-flavor-without-changing-your-settings)
   - [Example 3: Downloading into a directory that doesn't exist yet](#example-3-downloading-into-a-directory-that-doesnt-exist-yet)
   - [Example 4: Renaming the download](#example-4-renaming-the-download)
2. [Scripting](#scripting)
   - [Example 5: Silent, scriptable downloads](#example-5-silent-scriptable-downloads)
   - [Example 6: Downloading over a file that already exists](#example-6-downloading-over-a-file-that-already-exists)
3. [Updating Winapp2ool](#updating-winapp2ool)
   - [Example 7: Updating winapp2ool in place](#example-7-updating-winapp2ool-in-place)
   - [Example 8: Downloading winapp2ool.exe without updating](#example-8-downloading-winapp2oolexe-without-updating)
4. [Advanced Downloads](#advanced-downloads)
   - [Example 9: winapp3.ini and Archived entries.ini](#example-9-winapp3ini-and-archived-entriesini)

---

The transcripts below are taken from real runs. The current directory is shown as `C:\Tools\winapp2` throughout, and file sizes reflect winapp2.ini version `251109`.

## Acquiring winapp2.ini

### Example 1: Downloading winapp2.ini

**Context**

You want the current winapp2.ini database in the directory from which winapp2ool is currently running.

**Intent**

Download winapp2.ini, and see which flavor you actually get.

**Command**
```
winapp2ool -download winapp2
```

**Output**
```
Downloading winapp2.ini...
Download Complete.
Downloaded winapp2.ini to C:\Tools\winapp2
```

**Result**
```
C:\Tools\winapp2\
    winapp2.ini          (CCleaner flavor)
    winapp2ool.exe
```

**Explanation**
- No flavor flag was given, so the default flavor (CCleaner) was downloaded 
- The save directory defaults to the current directory
- Winapp2ool does not exit afterwards; it opens its main menu. See [Example 5](#example-5-silent-scriptable-downloads) for silent mode

**Notes**

To get the base (non-CCleaner) winapp2.ini instead:

```
winapp2ool -download winapp2 -base
```
```
Downloading winapp2.ini...
Download Complete.
Downloaded winapp2.ini to C:\Tools\winapp2
```

---

### Example 2: Choosing a flavor without changing your settings

**Context**

Your winapp2ool is configured for the base flavor because that's what you work with day to day. Today you need the CCleaner 7 database for a different machine, and the System Ninja rules for a third.

**Intent**

Fetch two other flavors without touching the saved flavor setting, and without visiting the settings menu twice.

**Commands**
```
winapp2ool -download winapp2 -cc7 -1d "C:\Tools\winapp2\ex2"
winapp2ool -download winapp2 -sn  -1d "C:\Tools\winapp2\ex2"
```

**Output**
```
Downloading winapp2.ini...
Download Complete.
Downloaded winapp2.ini to C:\Tools\winapp2\ex2

Downloading winapp2.rules...
Download Complete.
Downloaded winapp2.rules to C:\Tools\winapp2\ex2
```

**Result**
```
C:\Tools\winapp2\ex2\
    winapp2.ini          (CCleaner 7 flavor)
    winapp2.rules        (System Ninja flavor)
```

**Explanation**
- `-cc7` and `-sn` override the active flavor (base) 
- Command-line runs never write settings, so the saved flavor is still base afterwards

**Notes**

The System Ninja flavor is named `winapp2.rules` which is why both files can land in one directory without the second overwriting the first

---

### Example 3: Downloading into a directory that doesn't exist yet

**Context**

You keep dated snapshots of the database so you can diff releases against each other later.

**Intent**

Download the BleachBit flavor into a dated folder without creating the folder first.

**Command**
```
winapp2ool -download winapp2 -bb -1d "C:\Tools\winapp2\ex3\bleachbit\2026-07"
```

**Output**
```
Downloading winapp2.ini...
Download Complete.
Downloaded winapp2.ini to C:\Tools\winapp2\ex3\bleachbit\2026-07
```

**Result**
```
C:\Tools\winapp2\ex3\bleachbit\2026-07\
    winapp2.ini          (BleachBit flavor)
```

**Explanation**
- Neither `ex3`, `bleachbit`, nor `2026-07` existed beforehand; all three were created
- The BleachBit flavor of winapp2.ini is downloaded and placed in the newly created `C:\Tools\winapp2\ex3\bleachbit\2026-07`
---

### Example 4: Renaming the download

**Context**

You already have a `winapp2.ini` in your working directory that you have customized, and you want the upstream copy alongside it rather than over it.

**Intent**

Download winapp2.ini under a different name.

**Command**
```
winapp2ool -download 1 -1f winapp2-upstream.ini
```

**Output**
```
Downloading winapp2-upstream.ini...
Download Complete.
Downloaded winapp2-upstream.ini to C:\Tools\winapp2
```

**Result**
```
C:\Tools\winapp2\
    winapp2.ini                  (your customized copy)
    winapp2-upstream.ini         (newly downloaded upstream copy)
```

**Explanation**
- `-1f` overrides the default name that the file argument would otherwise choose

**Notes**

- `1` and `winapp2` are interchangeable
`-1f` also accepts a leading subdirectory: `-1f "\archive\winapp3.ini"` saves into `archive\` beneath the save directory, creating it if needed. See [Example 9](#example-9-winapp3ini-and-archived-entriesini)

---

## Scripting

### Example 5: Silent, scriptable downloads

**Context**

You want a scheduled task to refresh winapp2.ini overnight. The interactive behavior, the overwrite prompt, and the main menu that opens when the download finishes, would both hang it.

**Intent**

Download with no console output or requests for input

**Command**
```
winapp2ool -download winapp2 -s
```

**Output**
```
(nothing)
```

**Explanation**
- `-s` suppresses all console output
- The overwrite prompt is skipped and any existing file is overwritten without asking
- Winapp2ool exits instead of opening its menu, with exit code `0`

**Notes**

A bad argument or a failed download exits `1`, so a script can check the exit code:

```powershell
winapp2ool -download winapp2 -s -1d "C:\Tools\winapp2"
if ($LASTEXITCODE -ne 0) { throw "winapp2.ini download failed" }
```

---

### Example 6: Downloading over a file that already exists

**Context**

You are working interactively and re-download the readme into a directory that already has one.

**Intent**

Keep the old copy by saving the new one under a different name.

**Command**
```
winapp2ool -download readme
```

**Output**
```
readme.txt already exists in the target directory.
Enter a new file name, or leave blank to overwrite the existing file: readme-2026-07.md
Downloading readme-2026-07.md...
Download Complete.
Downloaded readme-2026-07.md to C:\Tools\winapp2
```

**Result**
```
C:\Tools\winapp2\
    readme.txt              (original file)
    readme-2026-07.md       (the new download)
```

**Explanation**
- The prompt appears because `readme.txt` already exists
- Entering `readme-2026-07.md` saved to a new file; the original was left alone
- Leaving the prompt blank and pressing enter would have overwritten `readme.txt` instead
- This prompt appears on the command line as well as in the menu, and it blocks. `-s` skips it and always overwrites. See [Example 5](#example-5-silent-scriptable-downloads)

---

## Updating Winapp2ool

### Example 7: Updating winapp2ool in place

**Context**

Your copy of winapp2ool is out of date and you want the current release.

**Intent**

Replace the running executable with the latest build from GitHub.

**Starting state**
```
C:\Tools\winapp2\
    winapp2ool.exe       (428,032 bytes, v1.6.9701.23408)
```

**Command**
```
winapp2ool -download winapp2ool
```

**Output**
```
(nothing)
```

**Result**
```
C:\Tools\winapp2\
    winapp2ool.exe                          (the new release, now running)
    winapp2ool v1.6.9701.23408.exe.bak      (the starting version, closed)
```

**Explanation**
- The save directory was the directory from which winapp2ool was running, so winapp2ool performs an automatic update
- The download's signature verified against a key built into winapp2ool, and its version was newer than the running one
- The running executable was renamed to `winapp2ool v<version>.exe.bak` before the new one took its place
- A new winapp2ool process was started and the old one exited

**Notes**

Self-updating is disabled when winapp2ool is running from the Windows temp folder, or when .NET is older than 4.6. 

With **Beta Participation** enabled, this same command installs the `Branch1` build instead of the `master` release.

---

### Example 8: Downloading winapp2ool.exe without updating

**Context**

You want to put a copy of winapp2ool on a network share for other machines, without disturbing the copy you're running.

**Intent**

Download the executable as an ordinary file.

**Command**
```
winapp2ool -download winapp2ool -1d "C:\Tools\winapp2\tools"
```

**Output**
```
Downloading winapp2ool.exe...
Download Complete.
Downloaded winapp2ool.exe to C:\Tools\winapp2\tools
```

**Result**
```
C:\Tools\winapp2\
    winapp2ool.exe                (still running)
    tools\
        winapp2ool.exe            (the downloaded release)
```

**Explanation**
- The save directory differs from the current directory, so no auto-update
- From the menu, the equivalent is **Change Save Directory** before choosing **Winapp2ool**

---

## Advanced Downloads

### Example 9: winapp3.ini and Archived entries.ini

**Context**

You maintain a machine running software that winapp2.ini no longer covers, and you want to fold the archived entries for it into your local database. Transmute's `-w` and `-a` preset flags look for `winapp3.ini` and `Archived Entries.ini` in the source directory.

**Intent**

Download both supplementary databases into an `archive\` subdirectory, then merge winapp3.ini into a local winapp2.ini.

**Commands**
```
winapp2ool -download winapp3  -1f "\archive\winapp3.ini"
winapp2ool -download archived -1f "\archive\Archived entries.ini"
winapp2ool -transmute -add -2d "C:\Tools\winapp2\archive" -2f winapp3.ini -3f winapp2-extended.ini
```

**Output**
```
Downloading winapp3.ini...
Download Complete.
Downloaded winapp3.ini to C:\Tools\winapp2\archive

Downloading Archived entries.ini...
Download Complete.
Downloaded Archived entries.ini to C:\Tools\winapp2\archive
```

**Result**
```
C:\Tools\winapp2\
    winapp2.ini                          (starting local copy)
    winapp2-extended.ini                 (transmute output)
    archive\
        winapp3.ini                        (downloaded)
        Archived entries.ini               (downloaded)
```

**Explanation**
- The `-1f "\subdir\name"` put both files in `archive\`, which was created by the first download if it didnt exist already 
- The Transmute step writes to `winapp2-extended.ini`, leaving the original winapp2.ini intact
