# Contributing to Winapp2.ini

This guide covers how to add and update entries in winapp2.ini under the current build workflow.

---

## How Winapp2.ini is maintained

Winapp2.ini is no longer edited directly. Instead, it is assembled from the source files in the `Assembler/` directory by [winapp2ool](https://github.com/MoscaDotTo/Winapp2/blob/master/winapp2ool/Readme.md) whenever a maintainer runs the build script.

**All contributions must target those source files, not any `Winapp2.ini` file directly.**

---

## Table of contents

1. [Quick reference](#quick-reference)
2. [Adding your first entry](#adding-your-first-entry)
3. [General rules](#general-rules)
4. [Standard entries](#standard-entries)
5. [Browser entries](#browser-entries)
6. [UWP entries](#uwp-entries)
7. [Winapp3.ini](#winapp3ini)
8. [Flavor corrections](#flavor-corrections)
9. [How to submit](#how-to-submit)
   - [Verifying your work](#verifying-your-work)

---

## Quick reference

| What you want to do | Where to edit |
| :- | :- |
| Add or update an entry for a win32-only application (traditional desktop software) | `Assembler/Entries/<letter>.ini` |
| Add or update an entry for an application with a UWP (or UWP and also a win32) version (Windows Store apps, generally) | `Assembler/UWP/AppInfo/<letter>.ini` |
| Add or update support for a Chromium-based web browser | `Assembler/BrowserBuilder/chromium.ini` |
| Add or update support for a Gecko-based web browser | `Assembler/BrowserBuilder/gecko.ini` |
| Fix or extend a generated browser entry | `Assembler/BrowserBuilder/browser_*.ini` |
| Add or update an entry for a power-user or aggressive operation | `Winapp3/Winapp3.ini` |
| Report a flavor-specific issue | Open an issue, or edit the appropriate flavor file in `Assembler/<FlavorName>`. The `<FlavorName>` is simply the name of the tool whose flavor you want to modify: `CCleaner`, `CCleaner7`, `BleachBit`,
  `SystemNinja`, or `Tron` |

---

## Adding your first entry

The path from "I want winapp2.ini to clean app X" to a merged PR.

### 1. Decide what kind of entry you're adding

| If the app is... | You're adding | Edit |
| :- | :- | :- |
| A traditional desktop install (.exe, .msi) | A standard entry | `Assembler/Entries/<letter>.ini` |
| A Microsoft Store app, or has both Store and desktop versions | A UWP entry | `Assembler/UWP/AppInfo/<letter>.ini` |
| A web browser | A browser definition | `Assembler/BrowserBuilder/chromium.ini` or `gecko.ini` |
| An aggressive operation aimed at power users | A Winapp3 entry | `Winapp3/Winapp3.ini` |

When in doubt, write a standard entry. 

### 2. Find what to clean

Install the app, use it normally for a while, and do whatever produces the data you want gone: open files, browse around, sign in, run a few jobs. Identify the location of this data on disk. Be careful to make sure that this data isn't comingled with important data such as application configuration.

### 3. Write the entry

Open the right source file, find the alphabetical spot for your entry, and write it. Follow the [Standard entries](#standard-entries), [UWP
entries](#uwp-entries), or [Browser entries](#browser-entries) section depending on what you're adding.

### 4. Validate

See [Verifying your work](#verifying-your-work).

### 5. Submit

Open a PR against `master`. 

---

## General rules

### Software availability

Entries are generally accepted for any software currently available for download. Direct vendor downloads are preferred; entries for software only available through download aggregators (e.g. Softonic, Uptodown) will be accepted, but entries are periodically audited, and software no longer available from its original vendor may be moved to the archive.

### Categorization

- **Games** and directly gaming-related software (launchers, modding tools, ancillary utilities, etc.) must use `Section=Games`.
- **Web browsers** are primarily generated entries. Their categorization is handled in the `BrowserInfo` section for that browser (see [Browser entries](#browser-entries)).
- **Everything else** should use the most appropriate `LangSecRef`. Prefer `LangSecRef` over `Section` when a standard CCleaner category fits. See the table in [Standard entries](#standard-entries) for valid values.

### Safety

Entries should clean transient, rebuildable data such as caches, logs, temporary files, recently-used lists or similar. Don't target settings, configurations, or anything a user would expect to survive a clean.

---

## Standard entries

Standard entries cover desktop (Win32) applications that are not web browsers and do not have UWP versions. They use standard winapp2.ini syntax and live in `Assembler/Entries/`.

### Where to edit

Each file in `Assembler/Entries/` contains entries whose names begin with a particular letter. Add your entry to the file matching the first letter of the entry name. Numbers and symbols go in `#.ini`.

### Entry name

Entry names must end with ` *]` — a space, an asterisk, and a closing bracket. This trailing ` *` is required.

```ini
[Application Name *]
```

### Style rules

Run WinappDebug on your entry before submitting to catch style and syntax issues. Many common issues can be automatically repaired by WinappDebug. 

| Rule | Detail |
| :- | :- |
| **Alphabetical order** | Entries must be sorted alphabetically within the file. Numbers and symbols sort before letters. |
| **Key order** | Within each entry: categorization → detection → warnings → deletion. Within a group, keys of the same type come before keys of a different type in alphabetical order by key name (Detect before DetectFile), and keys of the same type are sorted numerically by their index. |
| **Key numbering** | Keys of the same type are numbered sequentially starting from 1 with no gaps (`FileKey1`, `FileKey2`, ...). `Detect` and `DetectFile` may omit the `1` in the case that there's only a single key. `FileKey`, `RegKey`, and `ExcludeKey` always have numbers, even when there is only one |
| **No blank lines within an entry** | A single blank line separates entries from each other. No blank lines inside an entry. Winapp2ool will automatically remove blank lines from inside entries. |
| **No trailing whitespace** | Lines must not have trailing spaces or tabs. |
| **Use environment variables** | All filesystem paths must use environment variables. Never use hardcoded drive letters. |

### Valid `LangSecRef` values

LangSecRef values are defined by CCleaner v6.39 and some (but not all) are supported by other tools. Values marked "CCleaner only" are not supported by other tools and will display the number instead of the CCleaner section heading. 

| LangSecRef | CCleaner section | Notes |
| :- | :- | :- |
| 3001 | Internet Explorer | Deprecated |
| 3005 | Microsoft Edge (legacy) | Deprecated |
| 3006 | Edge Chromium | Called `Microsoft Edge` in BleachBit |
| 3021 | Applications | |
| 3022 | Internet | |
| 3023 | Multimedia | |
| 3024 | Utilities | |
| 3025 | Windows | Called `Microsoft Windows` in BleachBit |
| 3026 | Firefox | |
| 3027 | Opera | |
| 3028 | Safari | |
| 3029 | Google Chrome | |
| 3030 | Thunderbird | |
| 3031 | Windows Store | |
| 3032 | CCleaner Browser | CCleaner only |
| 3033 | Vivaldi | |
| 3034 | Brave | |
| 3035 | Opera GX | CCleaner only |
| 3036 | Spotify | CCleaner only |
| 3037 | Avast Secure Browser | CCleaner only |
| 3038 | AVG Secure Browser | CCleaner only |
| 3039 | Arc Browser | CCleaner only |
| 3040 | iTunes | CCleaner only |
| 3042 | WhatsApp | CCleaner only |
| 3043 | Norton Private Browser | CCleaner only |
| 3044 | Avira Secure Browser | CCleaner only |

###### Note: `3041` is not a valid `LangSecRef` value. 

### Key syntax reference

**Categorization** 

Required (exactly one). Defines the UI category the entry appears under.

###### Examples


```ini
LangSecRef=3021
```
```ini
Section=Games
```

**Detection** 

At least one Detect or DetectFile key is required. 

Detection keys control whether or not an entry is displayed to the user. If at least one detection key matches a path, file, or registry key on the current system, the entry is displayed. 

Scope detection keys to the target application as tightly as possible to avoid false positives. You may wish to capture individual versions or all of them, depending on the scope of the entry.

The last path segment of a `DetectFile` accepts wildcards. 

Nested wildcards are not supported by `DetectFile` in any tool except winapp2ool. 

`Detect` (registry) paths do not support wildcards. 

System Ninja does not support wildcards anywhere within a `DetectFile`.

###### Examples

```ini
Detect1=HKCU\Software\Vendor\AppName
Detect2=HKCU\Software\Vendor\AppName\AppVersion
DetectFile1=%LocalAppData%\Vendor\AppName
DetectFile2=%LocalAppData%\Vendor\AppName\AppVersion\App.exe
```

**Warnings** 

Optional. Shows a message when the cleaning routine moves from unchecked to checked in the UI. Many users skip them, so keep warnings short and specific.

Not supported by System Ninja or R-Wipe&Clean

```ini
Warning=This will delete your saved session data.
```

**Deletion** 

At least one required. Use `FileKey` for filesystem paths and `RegKey` for registry paths.

###### FileKey Example 
```ini
FileKey1=%LocalAppData%\App\Temp|*.tmp
FileKey2=%LocalAppData%\App\Cache|*|RECURSE
FileKey3=%LocalAppData%\App\OldVersions|*|REMOVESELF
```

- A `|` separates the path from the pattern.
- Multiple patterns in one key: `FileKey1=%Path%|file1.log;file2.log;*.tmp`
- `RECURSE` deletes matching files in all subdirectories. In CCleaner, also deletes empty subdirectories when provided `*.*` 
- `REMOVESELF` does the same as RECURSE, but always removes empty subdirectories left behind. Implies RECURSE. Will delete the parent folder if it is empty when the clean completes.

###### RegKey Example
```ini
RegKey1=HKCU\Software\App\RecentFiles
RegKey2=HKCU\Software\App\Settings|LastOpenedPath
```

A `|value` suffix deletes only that named value, not the key itself. 

`RegKey` does not support wildcards.

**Exclusions** 

Optional. Preserves specific files or registry keys from deletion. Prefer tightening deletion keys over adding exclusions.

```ini
ExcludeKey1=FILE|%AppData%\App\|important.db
ExcludeKey2=PATH|%AppData%\App\|*.cfg
ExcludeKey3=REG|HKCU\Software\App\Preserve
```

- `FILE` excludes a specific named file from deletion
- `PATH` excludes files matching a wildcard pattern; covers the directory and all subdirectories
- `REG` excludes a registry key and all values and subkeys beneath it 

`ExcludeKey` has no `RECURSE` or `REMOVESELF` equivalent. Subdirectories are covered by use of the `PATH` flag.

### Complete example

```ini
[Example Application *]
LangSecRef=3021
Detect=HKCU\Software\Example\ExampleApp
FileKey1=%LocalAppData%\Example\ExampleApp\Cache|*|REMOVESELF
FileKey2=%LocalAppData%\Example\ExampleApp\Logs|*.log
RegKey1=HKCU\Software\Example\ExampleApp\RecentFiles
```

### Environment variables

All filesystem paths must use environment variables. Never use hardcoded drive letters.

Variables marked with \* check both 64-bit and 32-bit locations on 64-bit systems.

Variables marked "CCleaner only" do not expand to anything in other tools and will not work with them. Avoid using CCleaner-only variables in submissions.

| Variable | Windows Vista–11 path | Notes |
| :- | :- | :- |
| `%AppData%` | `C:\Users\%UserName%\AppData\Roaming` | |
| `%CommonAppData%` | `C:\ProgramData` | CCleaner only |
| `%CommonProgramFiles%`\* | `C:\Program Files\Common Files` | |
| `%Documents%` | `C:\Users\%UserName%\Documents` | CCleaner only |
| `%LocalAppData%` | `C:\Users\%UserName%\AppData\Local` | |
| `%LocalLowAppData%` | `C:\Users\%UserName%\AppData\LocalLow` | CCleaner only |
| `%Music%` | `C:\Users\%UserName%\Music` | CCleaner only |
| `%Pictures%` | `C:\Users\%UserName%\Pictures` | CCleaner only |
| `%ProgramData%` | `C:\ProgramData` | |
| `%ProgramFiles%`\* | `C:\Program Files` | |
| `%Public%` | `C:\Users\%UserName%\Public` | |
| `%SystemDrive%` | `C:` | |
| `%UserProfile%` | `C:\Users\%UserName%` | |
| `%Video%` | `C:\Users\%UserName%\Videos` | CCleaner only |
| `%WinDir%` | `C:\Windows` | |

### Registry variables

| Variable | Registry hive |
| :- | :- |
| `HKCR` | `HKEY_CLASSES_ROOT` |
| `HKCU` | `HKEY_CURRENT_USER` |
| `HKLM` | `HKEY_LOCAL_MACHINE` |
| `HKU` | `HKEY_USERS` |
| `HKCC` | `HKEY_CURRENT_CONFIG` |

### Keys not accepted in contributions

The following keys are valid CCleaner-only syntax but are not used in winapp2.ini and should not appear in submitted entries.

| Key | Notes |
| :- | :- |
| `Default` | Controls whether an entry is enabled by default in CCleaner. All winapp2.ini entries are opt-in. |
| `DetectOS` | Limits an entry to specific Windows versions by kernel number. CCleaner only. |
| `SpecialDetect` | Uses CCleaner's internal application detection patterns. CCleaner only. Deprecated. |

---

## Browser entries

Web browser entries are generated automatically by winapp2ool's BrowserBuilder module. **Do not add browser entries manually to `Assembler/Entries/`.** Instead, edit the files in `Assembler/BrowserBuilder/`.

### How BrowserBuilder works

BrowserBuilder reads two files, `chromium.ini` and `gecko.ini`, and generates a cleaning entry for every browser/category combination described in them. Each browser is a `[BrowserInfo: Browser Name]` section. Each cleaning category (Caches, Cookies, History, etc.) is an `[EntryScaffold: Category Name]` section. BrowserBuilder produces one entry per browser per category.

After generating the base entries, BrowserBuilder applies corrections from the browser flavor files to handle cases where the generated output is wrong or incomplete for specific browsers.

### Adding a new browser

To add a new Chromium-based browser, add a `[BrowserInfo: Browser Name]` section to `chromium.ini`. For a Gecko-based browser, add to `gecko.ini`. Place the section alphabetically by browser name within the file.

**Required keys**

| Key | Purpose |
| :- | :- |
| `Section=` | The CCleaner section name for all generated entries for this browser. Use the format `Browser Name Web Browser` to group entries under their own section. Sections beginning with a number will fail to display correctly in CCleaner. Prepend them with a `.` eg. `.360 Secure Web Browser`|
| `UserDataPath=` | The filesystem path to the directory containing the browser's user profile data. For Chromium browsers this is the `User Data` folder. For Gecko browsers this is the `Profiles` folder. This path is also used to generate a `DetectFile` for every entry. |

**Optional keys**

| Key | Purpose |
| :- | :- |
| `RegistryRoot=` | The registry root key for this browser. Multiple values can be provided on separate lines. If omitted, RegKeys wont be created and EntryScaffolds carrying a `RequiresRegistryRoot` key will skip this browser. |
| `TruncateDetect=` | Provide this key (any value) to strip `\User Data\` from the generated `DetectFile`. Required when the `UserDataPath` contains a wildcard in the direct parent directory name. This value appears boolean but it is not, provide this key if and only if you want the DetectFile truncated.  |
| `Skip=` | Provide this key (any value) to exclude this browser from generation. Used to retire support without losing the configuration. This value appears boolean but it is not, provide this key if and only if you want the entry skipped |

Multiple `UserDataPath=` and `RegistryRoot=` values are supported; list them as separate keys.

**Style rules**

| Rule | Detail |
| :- | :- |
| **Section name format** | Must be `[BrowserInfo: Browser Name]` exactly. |
| **Alphabetical order** | Sections must be sorted alphabetically by browser name within the file. |
| **No trailing whitespace** | Lines must not have trailing spaces or tabs. |

**Chromium example**

```ini
[BrowserInfo: Example Browser]
Section=Example Web Browser
UserDataPath=%LocalAppData%\Example\ExampleBrowser\User Data
RegistryRoot=HKCU\Software\Example\ExampleBrowser
```

```ini
[BrowserInfo: Example Browser 2]
Section=Example Web Browser 2
UserDataPath=%LocalAppData%\RegistryLess\Browser\User Data
```

**TruncateDetect Example**

When a TruncateDetect key is provided, BrowserBuilder drops the last directory in the generated DetectFile to avoid problems with the CCleaner `DetectFile` parser. 

Note: `DetectFile` keys containing a wildcard generated with TruncateDetect are still incompatible with System Ninja due to System Ninja not supporting wildcards in `DetectFile`.
```
[BrowserInfo: Brave]
Section=Brave Web Browser
TruncateDetect=True
UserDataPath=%LocalAppData%\BraveSoftware\Brave-Browser*\User Data
RegistryRoot=HKCU\Software\BraveSoftware\Brave-Browser
RegistryRoot=HKCU\Software\BraveSoftware\Brave-Browser-Beta
RegistryRoot=HKCU\Software\BraveSoftware\Brave-Browser-Nightly
```

###### Generated DetectFile
```
DetectFile=%LocalAppData%\BraveSoftware\Brave-Browser*
```

**Gecko-specific variables**

Gecko's `gecko.ini` supports an additional inferred variable in `EntryScaffold` patterns:

- `%LocalDataPath%`: Gecko profile user data is stored in a folder located within `%AppData%`. Cache data is stored in a folder by the same name in `%LocalAppData%`. The `%LocalAppData%` path is inferred from `%UserDataPath%`.  

This is handled automatically and requires no action from contributors adding a Gecko browser.

### Applying corrections to generated browser entries

Generated entries aren't always right for every browser. The correction files let you fix or extend individual entries without touching the scaffolds.

Section names in correction files must exactly match the generated entry name, which follows the pattern `Browser Name Category *` (e.g. `Brave Caches *`).

BrowserBuilder corrections use the same Flavorize winapp2ool module as the tool-specific flavors. These files live in `/Assembler/BrowserBuilder/`

**When to use each file**

| File | Use when |
| :- | :- |
| 1. `browser_section_removals.ini` | An entire generated entry does not apply to a specific browser.
| 2. `browser_name_removals.ini` | A specific key name should be removed from a browser's entry. Number-sensitive match, key numbers must match. (eg. `FileKey3=` wont remove `FileKey1=`)|
| 3. `browser_value_removals.ini` | A specific key value is wrong for a browser regardless of which numbered key it appears on. Write the key with the target value; the number you give it is irrelevant. (eg. FileKey=some\path|* will remove any FileKey whose value is some\path|*)|
| 4. `browser_section_replacements.ini` | The generated entry for a browser is so different from the default that it needs to be replaced entirely. |
| 5. `browser_key_replacements.ini` | A specific key needs a different value for a browser (exact name and section match). |
| 6. `browser_additions.ini` | A browser needs extra keys beyond what the scaffold generates, or needs an entirely new entry that BrowserBuilder does not produce. |

**Style rules for correction files**

| Rule | Detail |
| :- | :- |
| **Section name must match exactly** | The section name must match the generated entry name precisely, including the ` *]` suffix. |
| **Include a comment** | Add a brief comment above each correction explaining why it is needed. A comment is a line which begins with `;` and appears directly above an entry. |
| **Keys only in the right files** | `browser_section_removals.ini` ignores key content — only the section name matters. `browser_name_removals.ini` ignores key values — only the name matters. Follow each file's rules. |
| **No trailing whitespace** | Lines must not have trailing spaces or tabs. |

**Addition example**

```ini
; Example Browser stores crash reports in a non-standard location not covered by the scaffold
[Example Browser Telemetry *]
FileKey1=%LocalAppData%\Example\ExampleBrowser\CrashReports|*|REMOVESELF
```

**Section removal example**

Section removals are performed on a name-only basis and need not contain any keys. Any provided keys will be ignored. 

```ini
; Example Browser does not support pinned tabs
[Example Browser Pinned Tabs *]
```

### Adding new EntryScaffolds

Adding a new `EntryScaffold` (a new cleaning category applied across all browsers) is possible. New scaffolds must be consistent: if the data exists across all browsers, the scaffold should generate correct entries for all of them, or corrections must be provided for browsers where it does not apply.

Every `EntryScaffold` section name must follow the format `[EntryScaffold: Category Name]`. The category name is appended to the browser name: `[EntryScaffold: Example Category]` produces entries named `Browser Name Example Category *`.

**Keys**

| Key | Purpose |
| :- | :- |
| `FileKeyBase=` | A FileKey template. BrowserBuilder substitutes browser-specific variables and adds the result as a numbered `FileKey` in the generated entry. Multiple values are supported; list them as separate keys. |
| `RegKeyBase=` | A RegKey template. BrowserBuilder substitutes `%RegistryRoot%` and adds the result as a numbered `RegKey`. Multiple values are supported. |
| `RequiresRegistryRoot=` | BrowserBuilder skips this scaffold for any browser without a `RegistryRoot=` key. |

A scaffold with neither `FileKeyBase=` nor `RegKeyBase=` produces no output.

**Variables**

Each template variable resolves to a value from the browser's `BrowserInfo` section.

| Variable | Available in | Expands to |
| :- | :- | :- |
| `%UserDataPath%` | `FileKeyBase` | The browser's user data directory, as declared by `UserDataPath=` in its `BrowserInfo` section. |
| `%BrowserPath%` | `FileKeyBase` | The parent directory of `%UserDataPath%`. For Chromium browsers, also generates a second FileKey with the path mirrored under `%ProgramFiles%`. |
| `%LocalDataPath%` | `FileKeyBase` (Gecko only) | The cache directory inferred from `%UserDataPath%` by substituting `%LocalAppData%` for `%AppData%`. |
| `%RegistryRoot%` | `RegKeyBase` | The browser's registry root path, as declared by `RegistryRoot=` in its `BrowserInfo` section. Expands once per `RegistryRoot=` value, producing one `RegKey` per root. |

**Example**

```ini
[BrowserInfo: Example Browser]
Section=Example Web Browser
UserDataPath=%AppData%\Example Browser\User Data
RegistryRoot=HKCU\Software\ExampleBrowser


[EntryScaffold: Example Category]
FileKeyBase=%UserDataPath%\*\ExampleData|*|REMOVESELF
FileKeyBase=%BrowserPath%\Application|example.log
RegKeyBase=%RegistryRoot%\ExampleKey
```

```ini
[Example Browser Example Category *]
Section=Example Web Browser
FileKey1=%AppData%\Example Browser\User Data\ExampleData|*|REMOVESELF
FileKey2=%AppData%\Example Browser\Application|example.log
RegKey1=HKCU\Software\ExampleBrowser\ExampleKey
```
---

## UWP entries

UWP entries cover Universal Windows Platform applications generally installed from the Microsoft Store. They are generated by winapp2ool's UWPBuilder module and live in `Assembler/UWP/AppInfo/`.

### How UWPBuilder works

UWPBuilder reads `Assembler/UWP/UWP.ini` (the baseline cleaning scaffold applied to every UWP app) alongside all 27 alphabetical AppInfo files. For each application, it generates a complete winapp2.ini entry by combining the scaffold's `FileKeyBase` patterns with app-specific keys, substituting the app's package folder name wherever `%Package%` appears.

```ini
[EntryScaffold: UWP App]
DetectFileBase=%Package%
FileKeyBase=%Package%\AC|*|RECURSE
FileKeyBase=%Package%\Settings|*.log*
FileKeyBase=%Package%\SystemAppData\Helium|*.log*
FileKeyBase=%Package%\TempState|*|REMOVESELF
```

The baseline scaffold above targets the standard UWP storage locations present in nearly every UWP application:
- `\AC\` — Application cache
- `\Settings\*.log*` — Settings logs
- `\SystemAppData\Helium\*.log*` — Helium logs
- `\TempState\` — Temporary state

Likewise, it also creates one `DetectFile` per provided package.  

UWPBuilder adds app-specific keys after the scaffold keys.

### Where to edit

Add your entry to `Assembler/UWP/AppInfo/<letter>.ini`, where `<letter>` is the first letter of the application name. Numbers and symbols go in `#.ini`.

### Finding the package folder name

The package folder name is the value for the `Package=` key. It is the name of the application's folder under `%LocalAppData%\Packages\`. You can find it by:

- Browsing `%LocalAppData%\Packages\` in Explorer and looking for a folder matching the application name.
- Running `Get-AppxPackage | Where-Object {$_.Name -like "*AppNameHere*"}` in PowerShell. The `PackageFamilyName` field gives you the folder name.

For multi-package applications, list each package folder separately as `Package1=`, `Package2=`, etc.

### Style rules

| Rule | Detail |
| :- | :- |
| **Section name format** | Must be `[Application Name *]` the trailing ` *` is required. This is the name that the generated entry will have. |
| **Alphabetical order** | Entries must be sorted alphabetically within the file. |
| **Categorization** | Exactly one of `LangSecRef=` or `Section=` is required. Use `Section=Games` for games and gaming-related apps. |
| **Package key** | At least one `Package=` key is required. Single-package apps use `Package=`. Multi-package apps use `Package1=`, `Package2=`, etc. |
| **No trailing whitespace** | Lines must not have trailing spaces or tabs. |
| **No blank lines within an entry** | A single blank line separates entries from each other. |

### Supported keys

| Key | Required | Purpose |
| :- | :- | :- |
| `Package=` | Yes (at least one) | Package folder name under `%LocalAppData%\Packages\`. |
| `LangSecRef=` or `Section=` | Yes (exactly one) | Categorization. |
| `Detect=`, `Detect1`, `Detect2=`, ... | No | Additional registry detection criteria. Passed through verbatim and renumbered. |
| `DetectFile=` | No | Additional filesystem detection criteria. Useful for hybrid win32+UWP apps. Each package folder is always generated as a `DetectFile` due to the root UWP.ini's scaffold containing `DetectFileBase=%Package%`  |
| `FileKeyBase=` | No | App-specific FileKey templates using `%Package%` or `%PackageN%` variables. Added after the scaffold keys. This is a stylistic choice only, and providing a `FileKey` will also work. |
| `FileKey=` | No | Standard winapp2.ini FileKey targeting win32 or non-package locations. Can also use `%Package%` variables. This is a stylistic choice only, and providing a `FileKeyBase` will also work. |
| `RegKey=` | No | Standard winapp2.ini RegKey. Passed through verbatim and renumbered. |
| `ExcludeKey=` | No | Standard winapp2.ini ExcludeKey. Can use `%Package%` variables. |
| `Skip=` | No | Provide this key (any value) to exclude the app from generation without removing its configuration. This value appears boolean but it is not, provide this key if and only if you want the entry skipped |

### The `%Package%` variable

In `FileKeyBase=`, `FileKey=`, and `ExcludeKey=` values, `%Package%` expands to the full path `%LocalAppData%\Packages\<PackageFolderName>` in the generated output.

For multi-package apps, `%Package%` expands to each of the packages. Use `%Package1%`, `%Package2%`, etc. to target a specific package. 

### Single-package example

###### UWP AppInfo Input
```ini
[Example UWP App *]
LangSecRef=3024
Package=ExampleCorp.ExampleApp_abc123xyz
FileKeyBase=%Package%\LocalCache\Local\Logs|*.log
```

###### UWPBuilder output (after linting)
```ini
[Example UWP App *]
LangSecRef=3024
DetectFile=%LocalAppData%\Packages\ExampleCorp.ExampleApp_abc123xyz
FileKey1=%LocalAppData%\Packages\ExampleCorp.ExampleApp_abc123xyz\AC|*|RECURSE
FileKey2=%LocalAppData%\Packages\ExampleCorp.ExampleApp_abc123xyz\LocalCache\Local\Logs|*.log
FileKey3=%LocalAppData%\Packages\ExampleCorp.ExampleApp_abc123xyz\Settings|*.log*
FileKey4=%LocalAppData%\Packages\ExampleCorp.ExampleApp_abc123xyz\SystemAppData\Helium|*.log*
FileKey5=%LocalAppData%\Packages\ExampleCorp.ExampleApp_abc123xyz\TempState|*|REMOVESELF
```
### Multi-package example

###### UWP AppInfo Input
```ini
[Multi-Package App *]
LangSecRef=3021
Package1=Vendor.AppCore_abc123
Package2=Vendor.AppService_def456
FileKeyBase=%Package%\LocalCache\Local\Logs|*.log
FileKeyBase=%Package2%\LocalState\Logs|*.log
FileKey=%ProgramData%\Vendor\AppService\Logs|*
```

###### UWPBuilder output (after linting)
```ini
[Multi-Package App *]
LangSecRef=3021
DetectFile1=%LocalAppData%\Packages\Vendor.AppCore_abc123
DetectFile2=%LocalAppData%\Packages\Vendor.AppService_def456
FileKey1=%LocalAppData%\Packages\Vendor.AppCore_abc123\AC|*|RECURSE
FileKey2=%LocalAppData%\Packages\Vendor.AppCore_abc123\LocalCache\Local\Logs|*.log
FileKey3=%LocalAppData%\Packages\Vendor.AppCore_abc123\Settings|*.log*
FileKey4=%LocalAppData%\Packages\Vendor.AppCore_abc123\SystemAppData\Helium|*.log*
FileKey5=%LocalAppData%\Packages\Vendor.AppCore_abc123\TempState|*|REMOVESELF
FileKey6=%LocalAppData%\Packages\Vendor.AppService_def456\AC|*|RECURSE
FileKey7=%LocalAppData%\Packages\Vendor.AppService_def456\LocalCache\Local\Logs|*.log
FileKey8=%LocalAppData%\Packages\Vendor.AppService_def456\LocalState\Logs|*.log
FileKey9=%LocalAppData%\Packages\Vendor.AppService_def456\Settings|*.log*
FileKey10=%LocalAppData%\Packages\Vendor.AppService_def456\SystemAppData\Helium|*.log*
FileKey11=%LocalAppData%\Packages\Vendor.AppService_def456\TempState|*|REMOVESELF
FileKey12=%ProgramData%\Vendor\AppService\Logs|*
```

This generates `DetectFile` entries for both packages. The scaffold is applied once per package. `%Package2%` targets only the second package in the app-specific `FileKey`.

### Hybrid win32 + UWP example

Some applications have both a traditional desktop (Win32) installer and a Microsoft Store (UWP) release. Winapp2.ini bundles these into a single entry generated through the UWP process. Win32-targeted paths use `FileKey=` alongside the package-targeted `FileKeyBase=` entries.

###### UWP AppInfo Input
```ini
[Hybrid App *]
LangSecRef=3021
Package=Vendor.HybridApp_ghi789
DetectFile=%LocalAppData%\Vendor\HybridApp
FileKeyBase=%Package%\LocalState\Cache|*|RECURSE
FileKey=%LocalAppData%\Vendor\HybridApp\Cache|*|RECURSE
FileKey=%LocalAppData%\Vendor\HybridApp\Logs|*.log
```

###### UWPBuilder Output 
```ini
[Hybrid App *]
LangSecRef=3021
DetectFile1=%LocalAppData%\Packages\Vendor.HybridApp_ghi789
DetectFile2=%LocalAppData%\Vendor\HybridApp
FileKey1=%LocalAppData%\Packages\Vendor.HybridApp_ghi789\AC|*|RECURSE
FileKey2=%LocalAppData%\Packages\Vendor.HybridApp_ghi789\LocalState\Cache|*|RECURSE
FileKey3=%LocalAppData%\Packages\Vendor.HybridApp_ghi789\Settings|*.log*
FileKey4=%LocalAppData%\Packages\Vendor.HybridApp_ghi789\SystemAppData\Helium|*.log*
FileKey5=%LocalAppData%\Packages\Vendor.HybridApp_ghi789\TempState|*|REMOVESELF
FileKey6=%LocalAppData%\Vendor\HybridApp\Cache|*|RECURSE
FileKey7=%LocalAppData%\Vendor\HybridApp\Logs|*.log
```

---

## Winapp3.ini

`Winapp3/Winapp3.ini` is an extension to winapp2.ini for power users. It uses identical syntax to standard winapp2.ini entries. Entries target operations that are more aggressive, broader in scope, or otherwise carry a higher risk of unintended data loss than entries in the main database. Users should understand what an entry does before enabling it.

Contributions to Winapp3.ini are subject to additional scrutiny. Entries must be clearly scoped, accurately described, and appropriate for a knowing power user. Include a `Warning=` key for entries that could affect application functionality.

Edit `Winapp3/Winapp3.ini` directly. The same style rules as standard entries apply.

---

## Flavor corrections

Flavors are custom builds of winapp2.ini with certain tweaks applied to make it more compatible with particular tools. 

Each flavor is built by applying a set of Transmute operations (structured changes in the form of additions, removals, and substitutions) to the base file. Correction files for each flavor live in `Assembler/<FlavorName>/`. If you find a flavor-specific problem or shortcoming, you can submit changes to these files or open an issue.

Each flavor directory contains the same six file types:

| File | What it does |
| :- | :- |
| `*_section_removals.ini` | Removes entire entries from the flavor output. |
| `*_name_removals.ini` | Removes specific keys from entries by key name (number-sensitive). |
| `*_value_removals.ini` | Removes specific keys from entries by exact value match (number-insensitive). |
| `*_section_replacements.ini` | Replaces entire entries in the flavor output. |
| `*_key_replacements.ini` | Replaces specific keys in entries. |
| `*_additions.ini` | Adds new entries or appends keys to existing entries. |

Corrections are applied in the order listed above. The same style rules apply as for the equivalent browser correction files: section names must exactly match the target entry name. Comments are helpful in the flavor files and will not be compiled into the final winapp2.ini.

---

## How to submit

**Pull request:** The preferred method. Fork the repository, make your changes to the appropriate source files in `Assembler/`, `Winapp3/`, or wherever relevant, and open a pull request against the `master` branch. Describe what changed and why.

**Issue:** If you found a problem but aren't comfortable making the change yourself, an issue works just as well. Include enough detail (application name, affected paths, what is wrong or missing) for maintainers to act on it.

### Verifying your work

Validation requires [downloading and running winapp2ool](https://github.com/MoscaDotTo/Winapp2/raw/refs/heads/master/winapp2ool/bin/Release/winapp2ool.exe) 

Launch Winapp2ool and select the Entry Lab menu by inputting 6, then select[BrowserBuilder](https://github.com/MoscaDotTo/Winapp2/blob/master/winapp2ool/modules/browserbuilder/README.md) by inputting 1. Use the BrowserBuilder to generate the standard entries for assessment if you've created or edited a browser entry.
BrowserBuilder is an entry generation tool that creates winapp2.ini entries for web browsers. See the BrowserBuilder README for usage guidance. Run WinappDebug on the generated output (browsers.ini by default)

Launch Winapp2ool and select the Entry Lab menu by inputting 6, then select [UWPBuilder](https://github.com/MoscaDotTo/Winapp2/blob/master/winapp2ool/modules/uwpbuilder/README.md) by inputting 2. Use the UWPBuilder generate the standard entries for assessment if you've created or edited a UWP AppInfo entry.
UWPBuilder is an entry generation tool that creates winapp2.ini entries for UWP applications. See the UWPBuilder README for usage guidance. Run WinappDebug on the generated output. (uwp.ini by default)

Launch Winapp2ool and input 1 to open [WinappDebug](https://github.com/MoscaDotTo/Winapp2/blob/master/winapp2ool/modules/winappdebug/README.md). Use this tool to evaluate any standard entries you edited or generated. Make any corrections necessary (or allow WinappDebug to make them if possible) such that your entries do not cause any errors. 
WinappDebug is a static analysis tool that catches and repairs a wide variety of style and syntax issues in winapp2.ini. See the WinappDebug README for usage guidance.   

If you can, run the full build:

1. Clone the repo. The build script is designed to run from the `Assembler/` directory and expects to find and place each flavor of winapp2.ini via relative paths. The build script will fail without the proper directory structure.
2. Make your changes to the source files. 
3. Place `winapp2ool.exe` 1.7+ in `Assembler/` or in your PATH variable. 
4. Run `& '.\build winapp2.ps1'` from PowerShell in the `Assember/` directory. If you get an execution policy error, run `Set-ExecutionPolicy -Scope Process RemoteSigned`
5. Open the generated `winapp2.ini` (each flavor is placed in its home location as part of the build script) and check that your entry is there and looks the way you expected. If something is wrong, go back and re-edit it until it is correct.
6. If you can go one step further: drop the built `winapp2.ini` into a cleaning tool, run the clean, and verify it removes what you expected and leaves what you expected. 