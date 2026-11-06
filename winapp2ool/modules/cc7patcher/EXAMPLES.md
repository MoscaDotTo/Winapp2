# CC7Patcher Examples

Worked examples for [CC7Patcher](readme.md). The [Common Uses](readme.md#common-uses) table in the readme lists the commands; these show what each one does, with real input and output.

# Table of Contents

1. [Example 1: Patching a ccleaner.ini](#example-1-patching-a-ccleanerini)
2. [Example 2: Trimming before patching](#example-2-trimming-before-patching)
3. [Example 3: The default workflow (download and patch)](#example-3-the-default-workflow-download-and-patch)
4. [Example 4: Review before deploying](#example-4-review-before-deploying)
5. [Example 5: Re-patching in place](#example-5-re-patching-in-place)

---

The outputs below are taken verbatim from real CC7Patcher runs. The `winapp2.ini` used in Examples 1, 2, and 5 is this four-entry excerpt of the real CCleaner 7 flavor:

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

The `ccleaner.ini` being patched is this minimal illustration:

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

We have a local CCleaner 7 format winapp2.ini and want to install its entries into `ccleaner.ini`. To keep the original file untouched for now, we direct the output to a separate file.

**Intent**

We want to append every entry from winapp2.ini into `ccleaner.ini`, saving the result as `ccleaner-patched.ini`.

**Command**

```
winapp2ool -cc7patcher -nodownload -3f ccleaner-patched.ini
```

###### Note: Without `-3f`, the output would overwrite `ccleaner.ini` in place

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

- All four winapp2.ini entries are appended into the file, exactly as written in the source
- The `[CCleaner Entry 1]` section and every key in it are untouched
- The output is sorted alphabetically by section name 
- The CCleaner 7 header comments are incidentally stripped from `ccleaner.ini`

---

### Example 2: Trimming before patching

**Context**

Most of the 3,700+ entries in winapp2.ini target software we don't have. There is no reason to make CCleaner 7 parse all of them at every startup. 

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

- Trim evaluated each entry in winapp2.ini's detection criteria against the demo machine before patching
- `[Google Chrome Autoplay Preferences *]` and `[Opera GX Autoplay Preferences *]` were dropped because their `DetectFile` paths do not exist on the demo machine
- `[Windows System Assessment Tool *]` and `[Windows Task Scheduler *]` detection criteria was found on the dmeo machine. These entries are retained and then appended into `ccleaner.ini`

---

### Example 3: The default workflow (download and patch)

**Context**

Download the current CCleaner 7 flavor and install it directly into CCleaner's live configuration.

**Intent**

We want to patch the real `ccleaner.ini` in `%ProgramFiles%\Piriform\CCleaner 7`

**Command**

```
winapp2ool -cc7patcher -2d "%ProgramFiles%\Piriform\CCleaner 7" -3d "%ProgramFiles%\Piriform\CCleaner 7"
```

**Output**

The run appends all entries from the current CCleaner 7 flavor into `ccleaner.ini`.

###### Note: The full `ccleaner.ini` produced by this command is too large to display here. The console output has been provided instead. 

```
Patching ccleaner.ini with entries from winapp2.ini
Patched file saved to C:\Program Files\Piriform\CCleaner 7\ccleaner.ini
CCleaner 7 patching complete
```

**Explanation**

- `-2d`/`-3d` point the input and output at CCleaner's configuration directory
- After the run, the winapp2.ini entries appear in CCleaner 7's cleaning options alongside its built-in definitions

---

### Example 4: Review before deploying

**Context**

We are cautious about a tool rewriting CCleaner's live configuration and want to inspect the result before committing to it.

**Intent**

We want to produce the patched file in the working directory so that we can review it. 

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

- `ccleaner.ini` is copied into the into the working directory from the CCleaner 7 directory
- `-3f ccleaner-patched.ini` writes the patched result to a separate file from the input for review 
- CCleaner 7's current unpatched `ccleaner.ini` is renamed to `ccleaner.ini.bak` as a backup 
- The newly patched `ccleaner-patched.ini` is copied into the CCleaner 7 directory and renamed to `ccleaner.ini`
---

### Example 5: Re-patching in place

**Context**

A new winapp2.ini version has been released, and we want to update our already-patched `ccleaner.ini` 

**Intent**

Show that re-patching is idempotent: CC7Patcher prunes the winapp2 entries the previous patch added, then re-adds the current set, so nothing is duplicated.

**Command**

```
winapp2ool -cc7patcher -nodownload -2f ccleaner-patched.ini -3f ccleaner-repatched.ini
```

**Output**

###### **Output file (`ccleaner-repatched.ini`)**

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
- The Add pass put it back 
- If this version of winapp2.ini had dropped an entry, the old copy would have been deleted without being readded
- Any entry not written by the Winapp2.ini Project is left unmodified
