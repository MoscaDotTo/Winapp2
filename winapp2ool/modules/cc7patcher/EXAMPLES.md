# CC7Patcher Examples

Worked examples for [CC7Patcher](readme.md). The [Common Uses](readme.md#common-uses) table in the readme lists the commands; these show what each one does, with real input and output.

# Table of Contents

1. [Example 1: Patching a ccleaner.ini](#example-1-patching-a-ccleanerini)
2. [Example 2: Trimming before patching](#example-2-trimming-before-patching)
3. [Example 3: The default workflow (download and patch)](#example-3-the-default-workflow-download-and-patch)
4. [Example 4: Review before deploying](#example-4-review-before-deploying)
5. [Example 5: Re-patching in place](#example-5-re-patching-in-place)

---

The outputs below come straight from real CC7Patcher runs. Examples 1, 2, and 5 use this four-entry excerpt of the real CCleaner 7 flavor as their `winapp2.ini`:

###### **Source file (`winapp2.ini`)**

```ini
[Google Chrome Autoplay Preferences *]
DetectFile=%LocalAppData%\Google\Chrome*
FileKey1=%LocalAppData%\Google\Chrome*\User Data\MEIPreload|*|REMOVESELF
ID=Google Chrome Autoplay Preferences *
Author=Winapp2.ini Project
Tags=Google,Chrome,Browser

[Opera GX Autoplay Preferences *]
DetectFile=%AppData%\Opera Software\Opera GX*
FileKey1=%AppData%\Opera Software\Opera GX*\_side_profiles\*\MEIPreload|*|REMOVESELF
FileKey2=%AppData%\Opera Software\Opera GX*\MEIPreload|*|REMOVESELF
ID=Opera GX Autoplay Preferences *
Author=Winapp2.ini Project
Tags=OperaGX,Browser

[Windows System Assessment Tool *]
DetectFile=%WinDir%\Performance\WinSAT
FileKey1=%WinDir%\Performance\WinSAT|*|REMOVESELF
ID=Windows System Assessment Tool *
Author=Winapp2.ini Project
Tags=ccwindows

[Windows Task Scheduler *]
Detect=HKCU\Software\Microsoft\Windows
FileKey1=%WinDir%\System32\Tasks_Migrated|*|REMOVESELF
ID=Windows Task Scheduler *
Author=Winapp2.ini Project
Tags=ccwindows
```

The `ccleaner.ini` they patch is a minimal stand-in for the real one:

###### **Base file (`ccleaner.ini`)**

```ini
; CCleaner - Cleaning rules

[CCleaner Entry 1]
ID=CCleaner.Entry1
Author=CCleaner
Tags=ccapps
```

---

### Example 1: Patching a ccleaner.ini

**Context**

We have a local CCleaner 7 format winapp2.ini and want to install its entries into `ccleaner.ini`. We'd like to leave the original file alone for now, so the output goes to a separate file.

**Intent**

We want to add every entry from winapp2.ini to `ccleaner.ini`, saving the result as `ccleaner-patched.ini`.

**Command**

```
winapp2ool -cc7patcher -nodownload -3f ccleaner-patched.ini
```

###### Note: Without `-3f`, CC7Patcher overwrites `ccleaner.ini`

**Output**

###### **Output file (`ccleaner-patched.ini`)**

```ini
[CCleaner Entry 1]
ID=CCleaner.Entry1
Author=CCleaner
Tags=ccapps

[Google Chrome Autoplay Preferences *]
DetectFile=%LocalAppData%\Google\Chrome*
FileKey1=%LocalAppData%\Google\Chrome*\User Data\MEIPreload|*|REMOVESELF
ID=Google Chrome Autoplay Preferences *
Author=Winapp2.ini Project
Tags=Google,Chrome,Browser

[Opera GX Autoplay Preferences *]
DetectFile=%AppData%\Opera Software\Opera GX*
FileKey1=%AppData%\Opera Software\Opera GX*\_side_profiles\*\MEIPreload|*|REMOVESELF
FileKey2=%AppData%\Opera Software\Opera GX*\MEIPreload|*|REMOVESELF
ID=Opera GX Autoplay Preferences *
Author=Winapp2.ini Project
Tags=OperaGX,Browser

[Windows System Assessment Tool *]
DetectFile=%WinDir%\Performance\WinSAT
FileKey1=%WinDir%\Performance\WinSAT|*|REMOVESELF
ID=Windows System Assessment Tool *
Author=Winapp2.ini Project
Tags=ccwindows

[Windows Task Scheduler *]
Detect=HKCU\Software\Microsoft\Windows
FileKey1=%WinDir%\System32\Tasks_Migrated|*|REMOVESELF
ID=Windows Task Scheduler *
Author=Winapp2.ini Project
Tags=ccwindows
```

**Explanation**

- All four winapp2.ini entries are added exactly as written in the source
- The `[CCleaner Entry 1]` section and every key in it are untouched
- The output is sorted alphabetically by section name
- The `; CCleaner - Cleaning rules` comment is gone. Winapp2ool doesn't keep comments when it rewrites an ini file

---

### Example 2: Trimming before patching

**Context**

Most of winapp2.ini targets software we don't have, and there's no reason to make CCleaner 7 parse all of it at every startup.

**Intent**

We want to patch `ccleaner.ini` with only the entries relevant to this machine. The demo machine has neither Google Chrome nor Opera GX installed.

**Command**

```
winapp2ool -cc7patcher -nodownload -trim -3f ccleaner-trimmed.ini
```

**Output**

The run reports `Trimming complete: 2 entries remain` before patching.

###### **Output file (`ccleaner-trimmed.ini`)**

```ini
[CCleaner Entry 1]
ID=CCleaner.Entry1
Author=CCleaner
Tags=ccapps

[Windows System Assessment Tool *]
DetectFile=%WinDir%\Performance\WinSAT
FileKey1=%WinDir%\Performance\WinSAT|*|REMOVESELF
ID=Windows System Assessment Tool *
Author=Winapp2.ini Project
Tags=ccwindows

[Windows Task Scheduler *]
Detect=HKCU\Software\Microsoft\Windows
FileKey1=%WinDir%\System32\Tasks_Migrated|*|REMOVESELF
ID=Windows Task Scheduler *
Author=Winapp2.ini Project
Tags=ccwindows
```

**Explanation**

- Before patching, Trim checked each entry's detection criteria against the demo machine
- `[Google Chrome Autoplay Preferences *]` and `[Opera GX Autoplay Preferences *]` were dropped because their `DetectFile` paths don't exist there
- The WinSAT folder and the `HKCU\Software\Microsoft\Windows` key both exist, so `[Windows System Assessment Tool *]` and `[Windows Task Scheduler *]` were kept and added to `ccleaner.ini`

---

### Example 3: The default workflow (download and patch)

**Context**

We want the current CCleaner 7 flavor, installed straight into CCleaner's live configuration.

**Intent**

We want to download the flavor and patch the real `ccleaner.ini` in `%ProgramFiles%\Piriform\CCleaner 7`.

**Command**

```
winapp2ool -cc7patcher -2d "%ProgramFiles%\Piriform\CCleaner 7" -3d "%ProgramFiles%\Piriform\CCleaner 7"
```

**Output**

The run adds every entry in the current CCleaner 7 flavor to `ccleaner.ini`. The patched file is far too large to show here, so this is the console output instead:

```
Patching ccleaner.ini with entries from winapp2.ini
Patched file saved to C:\Program Files\Piriform\CCleaner 7\ccleaner.ini
CCleaner 7 patching complete
```

**Explanation**

- `-2d`/`-3d` point the input and output at CCleaner's configuration directory
- Program Files isn't writable without admin rights, so winapp2ool asks for elevation through a UAC prompt before it patches
- After the run, the winapp2.ini entries appear in CCleaner 7's cleaning options alongside its built-in definitions

---

### Example 4: Review before deploying

**Context**

We'd rather not let a tool rewrite CCleaner's live configuration sight unseen, so we want to look at the result first.

**Intent**

We want the patched file in the working directory, where we can review it before it goes anywhere.

**Commands**

```
copy "%ProgramFiles%\Piriform\CCleaner 7\ccleaner.ini" .
winapp2ool -cc7patcher -3f ccleaner-patched.ini
```

Then, after reviewing `ccleaner-patched.ini` in a text editor:

```
copy "%ProgramFiles%\Piriform\CCleaner 7\ccleaner.ini" "%ProgramFiles%\Piriform\CCleaner 7\ccleaner.ini.bak"
copy ccleaner-patched.ini "%ProgramFiles%\Piriform\CCleaner 7\ccleaner.ini"
```

**Explanation**

- The first `copy` brings `ccleaner.ini` from the CCleaner 7 directory into the working directory
- `-3f ccleaner-patched.ini` writes the patched result to its own file, leaving the copied input alone
- The unpatched `ccleaner.ini` in the CCleaner 7 directory is backed up as `ccleaner.ini.bak`
- `ccleaner-patched.ini` then replaces CCleaner 7's `ccleaner.ini`
- The two `copy` commands that write into Program Files need a Command Prompt run as administrator

---

### Example 5: Re-patching in place

**Context**

A new winapp2.ini has been released, and we want to update the `ccleaner.ini` we patched in Example 1.

**Intent**

We want the update without any duplicated entries. CC7Patcher prunes the winapp2 entries the previous patch added, then adds the current set.

**Command**

```
winapp2ool -cc7patcher -nodownload -2f ccleaner-patched.ini -3f ccleaner-repatched.ini
```

**Output**

###### **Output file (`ccleaner-repatched.ini`, excerpt)**

```ini
[Google Chrome Autoplay Preferences *]
DetectFile=%LocalAppData%\Google\Chrome*
FileKey1=%LocalAppData%\Google\Chrome*\User Data\MEIPreload|*|REMOVESELF
ID=Google Chrome Autoplay Preferences *
Author=Winapp2.ini Project
Tags=Google,Chrome,Browser
```

**Explanation**

- This entry was already in `ccleaner-patched.ini`
- The prune pass removes it because it has an `Author=Winapp2.ini Project` key
- The Add pass puts it back
- Our run used the same four-entry winapp2.ini as Example 1, and the whole output file came out byte for byte the same as Example 1's
- If the new winapp2.ini had dropped an entry, the prune would delete the old copy and nothing would add it back
- Entries without that `Author=` key, like `[CCleaner Entry 1]`, are left alone
