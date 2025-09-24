# Winapp2.ini

### What is winapp2.ini? 

**Winapp2.ini** is a massive community-driven database of extended cleaning routines for Microsoft Windows, compatible with CCleaner, BleachBit, System Ninja, and R-Wipe&Clean. 

### Why winapp2.ini

Winapp2.ini is an extension of all of the applications with which it is compatible, which means it is able to update independently of them. This enables greater freedom of user choice to move between applications and particular versions of individual applications. Additionally, being a user extensible text file empowers individual users to further customize this already extensive set of cleaning routines to suit their needs as individuals, and to contribute them back to the community. 

### What are flavors?

Flavors are specific sets of modifications applied to each winapp2.ini update to produce variants which cater more closely to the specific features supported by particular applications. This is an automated process carried out when winapp2.ini is built for each update, so these flavors are always up to date with the latest version of winapp2.ini even if the copy shipped with the application is not. These are intended to function as drop-in replacements to the winapp2.ini shipped with each of these applications. 

### Disclaimer 
Winapp2.ini is provided as-is and without warranty. Understand that its intent is to enable you to delete files, folders, and registry keys off of your system in a way that is programmatic and potentially irreversible. Please exercise caution and take appropriate backups where relevant while using winapp2.ini. It is advised you use winapp2ool to manage your local copy of winapp2.ini, as it can provide bespoke changelogs which should be read carefully to fully understand the scope of changes made between versions.   

---

# Table of Contents

1. [Quick Start](#quick-start)
2. [Files of Interest](#files-of-interest)
3. [Installation & Configuration](#installation--configuration)
   - [CCleaner](#ccleaner)
   - [BleachBit](#bleachbit)
   - [System Ninja](#system-ninja)
   - [Avira System Speedup](#avira-system-speedup)
   - [Tron](#tron)
   - [R-Wipe & Clean](#r-wipe--clean)
4. [Creating Entries](#creating-entries)
   - [Styling](#styling)
   - [Naming](#naming)
   - [Categorization](#categorization)
   - [Detection Criteria](#detection-criteria)
   - [Warnings](#warnings)
   - [Deletion Routines](#deletion-routines)
   - [Unsupported Functions](#unsupported-functions)
   - [Variables](#variables)
5. [Custom Content](#custom-content)

---

# [Quick Start](#quick-start)
1. Download [winapp2ool.exe](https://github.com/MoscaDotTo/Winapp2/raw/master/winapp2ool/bin/Release/winapp2ool.exe) and your preferred flavor of winapp2.ini
2. Follow the installation guide for your cleaner application below
3. Use winapp2ool to keep your copy updated and trimmed for optimal performance

---


# [Files of interest](#files-of-interest)

Name           		                                                                                                            | Purpose       
:-                                                                                                                               | :-
[Winapp2ool](https://github.com/MoscaDotTo/Winapp2/raw/master/winapp2ool/bin/Release/winapp2ool.exe)                             | A robust tool that allows you to manage Winapp2.ini for your system, including automatic downloading and trimming. This tool has its own ReadMe [here](https://github.com/MoscaDotTo/Winapp2/tree/master/winapp2ool).
[Winapp2.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/master/Non-CCleaner/Winapp2.ini)                              | This is the base winapp2.ini file, it has no content removed or changed and includes and may overlap or conflict with CCleaner/BleachBit rules. 
[CCleaner Winapp2.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/master/Winapp2.ini)                                  | The CCleaner flavor of winapp2.ini, designed to reduce overlap with CCleaner rules and better integrate with its UI.
[BleachBit Winapp2.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/master/Non-CCleaner/BleachBit/Winapp2.ini)          | The BleachBit flavor of winapp2.ini, designed to remove unsupported rules and pass the sanity checker.
[System Ninja winapp2.rules](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/master/Non-CCleaner/SystemNinja/Winapp2.rules) | The System Ninja flavor of winapp2.ini, designed to replace unsupported rules with ones compatible with System Ninja. 
[Tron winapp2.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/master/Non-CCleaner/Tron/Winapp2.ini)                    | The Tron flavor of winapp2.ini, designed to capture the downstream changes made by Tron to the CCleaner flavor.  
[Winapp3.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/master/Winapp3/Winapp3.ini)                                   | An extension for an extension; contains entries for use by power users. *You should **not** use this file if you do not know what you are doing. Entries in this file can potentially be very aggressive/dangerous to your file system.*

# [Installation & Configuration](#installation--configuration) 

It is strongly recommended you keep a copy of [winapp2ool.exe](https://github.com/MoscaDotTo/Winapp2/raw/master/winapp2ool/bin/Release/winapp2ool.exe) in the same folder as winapp2.ini for the purpose of keeping it up-to-date irrespective of which application you are using. 

## [CCleaner](#ccleaner)
###### [Download CCleaner](https://www.ccleaner.com/ccleaner)

### Flavor

You should use the [CCleaner flavor](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/master/Winapp2.ini) for ideal integration into the UI and minimized rule overlap, however the base ("Non-CCleaner") [Winapp2.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/master/Non-CCleaner/Winapp2.ini) will also work 

### Installation

Place winapp2.ini in the same folder as ccleaner.exe. By default this is `..\Program Files\CCleaner`

It is advised that you use the Trim function of winapp2ool when updating winapp2.ini to reduce the CCleaner startup time, as the full winapp2.ini file can unnecessarily slow down the CCleaner start up process. 

### Configuration 

CCleaner will display the set of winapp2.ini entries which it detects as valid for your system inside its Applications tab. In modern versions of CCleaner, this tab is found in the Custom Clean section of the application. All winapp2.ini entries are disabled by default in CCleaner, and must be enabled individually or in groups. To enable an entire group of entries, right click on the section header and select "Check all."


###### Note: CCleaner 5.64.7577 is the last version to work on Windows XP and Vista (for non-SSE2 CPUs CCleaner 5.26.5937). Winapp2.ini and Winapp3.ini will continue to work with this version.

## [BleachBit](#bleachbit)
###### [Download BleachBit](https://www.bleachbit.org)

### Flavor

You should use the [BleachBit flavor](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/master/Non-CCleaner/BleachBit/Winapp2.ini). This flavor is designed to improve compatbility with BleachBit by eliminating errors thrown by BleachBit's sanity checker when using the base winapp2.ini. Use of any other flavor will throw a small number of errors and not allow you to run any entries which contain them, but will otherwise function correctly. 

### Installation

1. Ensure that you have disabled "Download and update cleaners from community (Winapp2.ini)" in the BleachBit settings.
2. Place winapp2.ini in `%AppData%\BleachBit\Cleaners`. 

Likewise, BleachBit maintains their own [customized version of winapp2.ini](https://github.com/bleachbit/winapp2.ini) which you can enable the use of from within the application:
1. Open BleachBit.
2. Select the "Edit" tab, and then "Preferences".
3. Check the box that reads "Download and update cleaners from community (Winapp2.ini)".

### Configuration

BleachBit will display the set of winapp2.ini entries which it detects as both having valid syntax and also as being valid for your system in its sidebar. All winapp2.ini entries are disabled by default in BleachBit, and must be enabled individually or in groups. To enable an entire group of entries, select the check box next to the section header. 

###### Note: BleachBit 2.2 is the last version to work on Windows XP. Winapp2.ini and Winapp3.ini will continue to work with this version. 

## [System Ninja](#system-ninja)
###### [Download System Ninja](https://singularlabs.com/software/system-ninja)

### Flavor 
You should use the [System Ninja Flavor](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/master/Non-CCleaner/SystemNinja/Winapp2.rules). This flavor is designed to improve compatibility with System Ninja by replacing keys with unsupported features, such that they become functional in System Ninja. It is not advised you use any other flavor with System Ninja. 

### Installation 
System Ninja ships with a copy Winapp2.ini by default, served from their servers, storing it in your `..\System Ninja\scripts\` directory as `winapp2.rules`

To keep your system ninja winapp2.rules up to date with winapp2.ini using winapp2ool instead: 
1. Open System Ninja
2. Select the "Options" tab
3. *Uncheck* the box that reads "Update cleaning rules and language files automatically"

This will prevent System Ninja from overwriting your local winapp2.rules file on launch

### Configuration 

System Ninja does not provide an interface for individually configuring which winapp2.ini entries run, they are all run and their output is displayed in the Junk Scanner window. The "Type" column in System Ninja displays the name of the entry as it appears in winapp2.ini 

###### Note: System Ninja 3.2.7 is the last version to work on Windows XP and Vista. Winapp2.ini and Winapp3.ini will continue to work with this version.

## [Avira System Speedup](#avira-system-speedup)
###### [Download System Speedup](https://www.avira.com/en/avira-system-speedup-free)

### Installation 
Avira System Speedup ships with a copy Winapp2.ini by default, served by Avira, storing it in your `..\Avira\System Speedup\sdf` directory. You can replace or update this local copy without issue or changing any of the Avira System Speedup settings.

If you are updating your winapp2.ini for Avira System Speedup, it is suggested you use the the base ("Non-CCleaner") [Winapp2.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/master/Non-CCleaner/Winapp2.ini)

### Configuration 

Avira System Speedup scans every winapp2.ini entry and displays the results of that scan in a panel labeled Third Party Applications. You can manually enable/disable items within this menu before activating the clean function. 

###### Note: Cleaning "Third Party Applications" is a paid feature of Avira System Speedup Pro. winapp2.ini is and always will be free, and supported by a variety of free applications.   

## [Tron](#tron)
###### [Tron GitHub](https://github.com/bmrf/tron)

### Flavor 

You should use [Tron Flavor](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/master/Non-CCleaner/Tron/Winapp2.ini)

### Installation 

Tron maintains and ships their own [customized version of Winapp2.ini](https://raw.githubusercontent.com/bmrf/tron/refs/heads/master/resources/stage_1_tempclean/ccleaner/winapp2.ini) by default, storing it in your `..\resources\stage_1_tempclean\ccleaner` directory.
  
Likewise, there exists a [Tron flavor of winapp2.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/master/Non-CCleaner/Tron/Winapp2.ini) on this repo which is kept up to date with the latest winapp2.ini. It is based on the CCleaner flavor and is intended to capture the downstream changes required by the Tron project while also keeping the rest of winapp2.ini up to date 

### Configuration 

Tron ships with its own configuration. You can modify it by opening the copy of CCleaner shipped with Tron.

## [R-Wipe & Clean](#r-wipe--clean)
###### [Download R-Wipe&Clean](https://www.r-wipe.com)

### Installation 

R-Wipe & Clean has unofficial support for Winapp2.ini. The steps below are adapted from [this thread](https://forum.r-tt.com/viewtopic.php?t=11018) 

1. In R-Wipe & Clean, select "Tools" from the menu bar 
2. From the Tools menu, click R-Wipe&Clean Smart
3. In R-Wipe & Clean Smart, select "Settings" from the menu bar 
4. From the settings menu, select "INI Import Settings"
5. In the INI Import Settings window, fill out the following: 
   *  Key Name for Registry Key Detection: `Detect`
   * Key Name for File/Folder Detection: `DetectFile`
   * Parameter for Recursive Subfolder Cleaning: `RECURSE`
   * Parameter for Removing Folder after Cleaning: `REMOVESELF`
6. Press OK
7. In R-Wipe & Clean Smart, select "Advanced" from the menu bar
8. In the Advanced menu, select "Import From .INI"
9. Browse to and select winapp2.ini 
10. Import as either many wipe lists or as a single one, whichever is your preference 
    * Because you will need to repeat the process of importing winapp2.ini as wipe lists in order to apply updates, it is advised you import it as a single wipelist which you can then easily remove after updating 
11. Once winapp2.ini is imported, you should see a message saying `LangSecRef`, `Section`, and `Warning` are unsupported functions, press OK 

### Configuration 

Configure the imported wipe lists individually just as you would R-Wipe & Clean's native wipe lists

---

# [Creating Entries](#creating-entries)

## [Styling](#styling) 

Winapp2.ini entries are organized alphabetically, between both sections in the files and individual key values in those sections. Alphabetically, numbers and symbols have precedence over letters. Entries should be ordered with their keys in the same order the appear in this guide (that is, categorization, detection, and deletion). You can use winapp2ool's WinappDebug module to ensure your style and syntax are correct. 

## [Naming](#naming) 

Entry names are displayed in the application UI to indicate the name of a set of cleaning rules. They appear exactly in the UI exactly as they are written between the `[brackets]`. 

`[Entry Name *]`

 Please include the space between the name and the \* when submitting as this is how we visually indicate that Winapp2.ini entries are not built into CCleaner. The * is not displayed in BleachBit or System Ninja 

## [Categorization](#categorization)

### `LangSecRef` or `Section`

These keys define where the entry will be displayed within the application UI. You must provide one and only one of these keys (ie. not neither and also not both). If neither key is provided, CCleaner appends the entry to the bottom of its Applications tab in a nameless section. BleachBit, however, contains a winapp2.ini sanity checker, and will report an error in the side panel. Entries which fail the BleachBit sanity check are not displayed to the user and cannot be enabled. 
- Any value can be provided for `Section` keys.
- If submitting an entry for a video game or related software, please use `Section=Games`.
- For `LangSecRef`, CCleaner syntax is followed. Tables of valid values are below:

LangSecRef     | CCleaner UI Header                | Notes 
:-             | :-                                | :-
3001           | Internet Explorer                 | No longer used in winapp2.ini 
3005           | Microsoft Edge                    | No longer used in winapp2.ini 
3006           | Edge Chromium                     | Called Microsoft Edge in BleachBit 
3021           | Applications                      | 
3022           | Internet                          |
3023           | Multimedia                        | 
3024           | Utilities                         |
3025           | Windows                           | Called Microsoft Windows in BleachBit
3026           | Firefox                           |
3027           | Opera                             |
3028           | Safari                            |
3029           | Google Chrome                     |
3030           | Thunderbird                       |
3031           | Windows Store                     |
3032           | CCleaner Browser                  | Not available in BleachBit 
3033           | Vivaldi                           | 
3034           | Brave                             |
3035           | Opera GX                          | Not available in BleachBit
3036           | Spotify                           | Not available in BleachBit
3037           | Avast Secure Browser              | Not available in BleachBit
3038           | AVG Secure Browser                | Not available in BleachBit

#### Examples 
`LangSecRef=3026` will cause an entry to be displayed under CCleaner's built-in Firefox section. 
`Section=Games` will cause an entry to be displayed in a new section called Games at the bottom of the CCleaner application tab or in the BleachBit side panel. 

## [Detection Criteria](#detection-criteria)

### Detect or DetectFile

 These keys specify the condition under which the entry should be considered valid for a system.
 If the target of at least one ofthese keys exists, the entry will be shown to the user. Otherwise, it won't. `Detect` keys point to Windows Registry paths and `DetectFile` keys point to Windows Filesystem paths. You should select a detection criteria that is only valid to the scope of a particular application's installation for best results, however you can use any key. You can define as many detection criteria as is necessary.  

 ###### Notes:
 - **System Ninja does not support wildcards anywhere in the DetectFile**
 - Nesting wildcards in DetectFiles is supported by winapp2ool but not by CCleaner.
 - If you place a wildcard in a DetectFile, ensure that it is in the last part of the parameterization. 
 - Wildcards are not supported in the registry by any application. 

#### Examples 

To set a Detection criteria using the registry, simply point a `Detect` key at a valid registry path.
- `Detect=HKLM\Software\Microsoft\Windows` will return as valid on every copy of Windows 
- `Detect=HKCU\Software\GRAHL\PDFAnnotator` will only return as valid on systems which have PDF Annotater installed.   

To set a Detection criteria using the file system, simply point a `DetectFile` key at a valid file or path on the system
- `DetectFile=%WinDir%` will return as valid on every copy of windows 
- `DetectFile=%LocalAppData%\Packages\9E2F88E3.Twitter_*` will return as valid on systems which have Twitter installed from the Microsoft Store 

## [Warnings](#warnings)

### `Warning`

This key provides a notice to the user with important information about the entry. It should be provided when running an entry may have unforeseen consequences for the user. This information is only displayed when an option is elected (which is effectively only once for most users), so it is wise to keep warnings short and concise so users may read and fully understand them. This key is only supported by CCleaner and BleachBit. 

## [Deletion Routines](#deletion-routines) 

### `FileKey`

These keys specify a file system location to be cleaned. They support wildcards and nesting wildcards. You can specify multiple file parameters in a single statement by appending them to the first file parameter as comments. Additionally, this key supports two flags: `RECURSE` and `REMOVESELF`. The `RECURSE` flag executes the deletion pattern across all subfolders of the given folder. `REMOVESELF` does the same, removing any empty subfolders left behind. 

#### Examples 
* `FileKey1=%SystemRoot%\junk|file` will delete `file` from `%SystemRoot%\junk`
* `FileKey1=%SystemRoot%\junk|file1;file2` will delete both `file1` and `file2` from `%SystemRoot%\junk`.
* `FileKey1=%SystemRoot%\junk|junkfile*|RECURSE` will delete any files whose name begin with `junkfile` from `%SystemRoot%\junk` and all of its subdirectories.
* `FileKey1=%SystemRoot%\junk|junkfile*|REMOVESELF` will delete any files whose name begin with `junkfile` from `%SystemRoot%\junk` and all of its subdirectories, also removing any empty folders either found or left behind by this operations. 

### `RegKey`

These keys specify a registry location to be cleaned. They do not support wildcards of any sort. Cleaning is supported for both keys and subkeys. To specify a subkey, use `|` after the key and then specify the sub key name 

#### Examples 
* `RegKey1=RegKey1=HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\CD Burning\StagingInfo` will delete the registry key `HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\CD Burning\StagingInfo`
* `RegKey1=HKCU\Software\Microsoft\Windows\Windows Error Reporting\Debug|StoreLocation` will delete the sub-key `StoreLocation` from the registry key `HKEY_CURRENT_USER\Software\Microsoft\Windows\Windows Error Reporting\Debug`, but will not delete the key itself. 

### `ExcludeKey`

These keys specify a file (or files), directory, or registry path to be excluded from cleaning, even if they would otherwise be deleted by a FileKey or RegKey. These keys support three flags for their parameters: `FILE`, `PATH`, and `REG`. Use the `FILE` flag to specify just a single file for exclusion. Use the `PATH` flag to specify an extension type. Use the `REG` flag to exclude a registry key. 

#### Examples 

* `ExcludeKey1=FILE|%WinDir%\System32\LogFiles\|myfile.txt` excludes `myfile.txt` in the `%WinDir%\System32\LogFiles` directory from being deleted.
* `ExcludeKey1=PATH|C:\Windows\|*.exe` excludes files of type `.exe` in the `C:\Windows` directory from being deleted.
* `ExcludeKey2=PATH|C:\Temp\|*.*` excludes all of the files located in the `C:\Temp` directory and all sub directories from being deleted.
* `ExcludeKey3=PATH|%WinDir%\System32\LogFiles\SCM\|*-*-*-*.*` excludes all of the files whose name matches the pattern  `*-*-*-*.*` in the `%WinDir%\System32\LogFiles\SCM` directory from being deleted.
* `ExcludeKey2=REG|HKCU\Software\Piriform` will prevent registry keys in `HKEY_CURRENT_USER\Software\Piriform` from being deleted. 
  
## [Unsupported functions](#unsupported-functions)

The following are functions that are not used in the official Winapp2.ini file, but are valid syntax. 

### `Default`
This key is used by only by CCleaner to decide whether or not an entry is enabled by default when it is loaded for the first time. `Default=True` causes an entry to be enabled by default, `Default=False` will not. CCleaner assumes `Default=False` when no `Default` key is provied.

### `DetectOS`
This key is used only by CCleaner and is used to limit the set of Windows version for which an entry should be considered valid. `DetectOS` failure supercedes `Detect` and `DetectFile` successes in CCleaner and winapp2ool. Kernel version numbering is used, below is a table for reference. To specify a minimum version number, use `DetectOS=num|`. To specify a maximum version number, use `DetectOS=|num`. To limit an entry to a set of windows versions, use `DetectOS=num1|num2`. To strictly limit an entry to a particular Windows version, use `DetectOS=num|num`.


| Kernel Number  | Windows Version
| :-             | :-
| 5.0            | Windows 2000
| 5.1            | Windows XP
| 5.2            | Windows XP 64-Bit Edition, Windows Server 2003, Windows Home Server
| 6.0            | Windows Vista, Windows Server 2008
| 6.1            | Windows 7, Windows Server 2008 R2, Windows Home Server 2011
| 6.2            | Windows 8, Windows Server 2012
| 6.3            | Windows 8.1, Windows Server 2012 R2
|10.0            | Windows 10, Windows 11, Windows Server 2016, Windows Server 2019

#### Examples 
  * `DetectOS=6.0|` will only run on versions of Windows including and newer than Windows Vista.
  * `DetectOS=|6.1` will only run on versions of Windows including and older than Windows 7.
  * `DetectOS=5.1|6.1` will run only on Windows XP, Windows Vista, and Windows 7.
  * `DetectOS=5.1|5.1` will only run on WindowsXP.

### `SpecialDetect`
This key is used only by CCleaner and provides the same search patterns used by some of the internal CCleaner rules for detection. Like `DetectOS`, this key superecedes successes or failures from `Detect` or `DetectFile` keys. 

#### Examples 

`SpecialDetect=DET_CHROME` will return valid on any system for which CCleaner detects chrome

## [Variables](#variables)

### Environment variables:

These are all the possible variables that can be used for defining filesystem paths in Winapp2.ini.
##### Variables marked with a * natively check both 64bit and 32bit locations on 64bit systems.

| Variable                | Windows Vista-11 Path                  | WindowsXP Path                                                         | Notes
| :-                      | :-                                     | :-                                                                     |:-
| `%AppData%`             | `C:\Users\%UserName%\AppData\Roaming`  | `C:\Documents and Settings\%UserName%\Application Data`                |
| `%CommonAppData%`       | `C:\ProgramData`                       | `C:\Documents and Settings\All Users\Application Data`                 | CCleaner only 
| `%CommonProgramFiles%`* | `C:\Program Files\Common Files`        | `C:\Program Files\Common Files`                                        | 
| `%Documents%`           | `C:\Users\%UserName%\Documents`        | `C:\Documents and Settings\%UserName%\My Documents`                    | CCleaner only
| `%LocalAppData%`        | `C:\Users\%UserName%\AppData\Local`    | `C:\Documents and Settings\%UserName%\Local Settings\Application Data` |
| `%LocalLowAppData%`     | `C:\Users\%UserName%\AppData\LocalLow` | N/A                                                                    | CCleaner only
| `%Music%`               | `C:\Users\%UserName%\Music`            | `C:\Documents and Settings\%UserName%\My Documents\My Music`           | CCleaner only
| `%Pictures%`            | `C:\Users\%UserName%\Pictures`         | `C:\Documents and Settings\%UserName%\My Documents\My Pictures`        | CCleaner only 
| `%ProgramData%`         | `C:\ProgramData`                       | N/A                                                                    |
| `%ProgramFiles%`*       | `C:\Program Files`                     | `C:\Program Files`                                                     | 
| `%Public%`              | `C:\Users\Public`                      | N/A                                                                    |
| `%SystemDrive%`         | `C:`                                   | `C:`                                                                   | 
| `%UserProfile%`         | `C:\Users\%UserName%`                  | `C:\Documents and Settings\%UserName%`                                 | 
| `%Video%`               | `C:\Users\%UserName%\Videos`           | `C:\Documents and Settings\%UserName%\My Documents\My Videos`          | CCleaner only 
| `%WinDir%`              | `C:\Windows`                           | `C:\Windows`                                                           |

### Registry variables

These are all the possible variables that can be used for defining registry paths in Winapp2.ini 

| Variable       | Registry Path 
| :-             | :-        
| HKCR           | HKEY_CLASSES_ROOT
| HKCU           | HKEY_CURRENT_USER
| HKLM           | HKEY_LOCAL_MACHINE 
| HKU            | HKEY_USERS
| HKCC           | HKEY_CURRENT_CONFIG

---

# [Custom content](#custom-content)

Winapp2.ini does not support non-English system configurations or portable software natively. If you have need for these features, we recommend you utilize a "Custom.ini" file, and use Winapp2ool's [Transmute](https://github.com/MoscaDotTo/Winapp2/tree/master/winapp2ool/modules/transmute) feature with the Transmute mode set to `Add` to add your custom configurations while keeping winapp2.ini up to date.

Winapp2ool 1.6 removed the Merge feature and replaced it with Transmute. If you were previously using Custom.ini with Merge, please see [Migrating From Merge](https://github.com/MoscaDotTo/Winapp2/tree/master/winapp2ool/modules/transmute#migrating-from-merge) in the Transmute ReadMe.

