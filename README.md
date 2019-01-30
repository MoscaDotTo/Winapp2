
# Winapp2.ini
A database of extended cleaning routines for popular Windows PC based maintenance software.

### Files of interest on this repo:

Name           | Purpose       
:------------- | :-------------
[winapp2.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/master/Winapp2.ini) | An extended database of cleaning routines for CCleaner. *This is the "main" file, **and the one most users will want***.
[winapp2ool](https://github.com/MoscaDotTo/Winapp2/raw/master/winapp2ool/bin/Release/winapp2ool.exe) | A robust tool that allows you to manage winapp2.ini for your system, including automatic downloading and trimming. This tool has its own ReadMe [here](https://github.com/MoscaDotTo/Winapp2/tree/master/winapp2ool).
[non-CCleaner winapp2.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/master/Non-CCleaner/Winapp2.ini) | If you don't use CCleaner, this is the file you want. It includes entries that were removed from the main file due to having been included in CCleaner's official distribution. *You should **not** use this file with CCleaner.*
[winapp3.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/master/Winapp3/Winapp3.ini) | An extension for an extension; contains entries for use by power users. *You should **not** use this file if you do not know what you are doing. Entries in this file can potentially be very aggressive/dangerous to your file system.*

### How to use winapp2.ini for the following cleaners:

* [CCleaner](https://www.ccleaner.com/ccleaner):
  * Download the latest winapp2.ini from this repo and place it in the same directory as ccleaner.exe.

* [BleachBit](https://www.bleachbit.org):
  * Open BleachBit.
  * Select the "Edit" tab, and then "Preferences".
  * Check the box that reads "Download and update cleaners from community (winapp2.ini)".

* [System Ninja](https://singularlabs.com/software/system-ninja):
  * System Ninja ships with winapp2.ini by default, storing it in your `..\System Ninja\scripts\` directory.

* [Avira System Speedup](https://www.avira.com/en/avira-system-speedup-free):
  * Avira System Speedup ships with winapp2.ini by default, storing it in your `..\Avira\System Speedup\sdf` directory.

* [Tron](https://github.com/bmrf/tron):
  * Tron ships with winapp2.ini by default, storing it in your `..\tron\resources\stage_1_tempclean\ccleaner` directory.

## Creating entries

Winapp2.ini entries are organized alphabetically, between sections in the file and individual key values in those sections. Alphabetically, numbers and symbols have precedence over letters. Entries should be ordered with their keys in the following precedence order.

`[Entry Name *]`
* The name of the entry as it will appear to users.
* Please include the space between the name and the \* when submitting as this is how we visual indicate that winapp2.ini entries are not a part of the host application running them.

`DetectOS`
* This key is used to specify which operating systems the entry is for.
* Not needed unless you want to limit your entry from running on specific versions of windows.
* Kernel version numbering is used, here is a table for reference:

Kernel Number  | Windows Version
:------------- | :-------------  
 5.0 | Windows 2000
 5.1 | Windows XP
 5.2 | Windows XP 64-Bit Edition, Windows Server 2003, Windows Home Server
 6.0 | Windows Vista, Windows Server 2008
 6.1 | Windows 7, Windows Server 2008 R2, Windows Home Server 2011
 6.2 | Windows 8, Windows Server 2012
 6.3 | Windows 8.1, Windows Server 2012 R2
10.0 | Windows 10, Windows Server 2016, Windows Server 2019

* To strictly limit an entry to a particular Windows version, use `DetectOS=num|num`:
  * `DetectOS=5.1|5.1` will only run on WindowsXP.
* To specify a minimum version number, use `DetectOS=num|`:
  * `DetectOS=6.0|` will only run on versions of Windows including and newer than Windows Vista.
* To specify a maximum version number, use `DetectOS=|num`:
  * `DetectOS=|6.1` will only run on versions of Windows including and older than Windows 7.
* DetectOS criteria supersede all other detection criteria.
  * This means that entries with valid `Detect` or `DetectFile` keys will not be runnable if the system does not also pass the `DetectOS` check.

`LangSecRef` or `Section`
* Defines the category of the entry.
* If this key is not provided, CCleaner will appended the entry to the bottom of the Applications tab.
* Any value can be provided for `Section` keys.
* If submitting an entry for a video game, please use `Section=Games`.
* For LangSecRef, CCleaner syntax is followed. A table of valid values is below:

LangSecRef     | Section
:------------- | :-------------
3021 | Applications
3022 | Internet
3023 | Multimedia
3024 | Utilities
3025 | Windows
3026 | Firefox
3027 | Opera
3028 | Safari
3029 | Google Chrome
3030 | Thunderbird
3031 | Windows Store

`Detect` or `DetectFile`
* These keys specify the condition under which the key should be considered valid for a system.
* If the target of these keys exists, the entry will be shown to the user. Otherwise it won't.
* `Detect` keys point to Windows Registry paths.
* `DetectFile` keys point to Windows Filesystem paths. Can point to either a directory or a specific file.
  * Wildcards are supported, but only in the last part of the parameterization.
  * Nesting wildcards in DetectFile is not supported by CCleaner.
  
`Default`
* Indicates the default state of the entry.
* Only used by CCleaner.
* When submitting, every entry must contain `Default=False` as winapp2.ini entries are always disabled by default.

`Warning`
* This key is not required by any application.
* Displays a warning to the user when they select an option.
  * Encouraged when running an entry may have unforeseen consequences for the user. 
  * Only displayed when an option is selected (effectively once for most users), so it is wise to keep warnings short and concise, as users may read or fully understand them.

`FileKey`
* Specifies a File System location(s) to be cleaned.
* Supports wildcards and nesting.
* You can specify multiple files to be cleaned at once by adding them as comments:
  * `FileKey1=%SystemRoot%\junk|file1;file2` will delete both `file1` and `file2` from `%SystemRoot%\junk`.
* The `RECURSE` flag will apply your cleaning parameters to all subdirectories of your given file system path:
  * `FileKey1=%SystemRoot%\junk|junkfile*|RECURSE` will delete any files whose name begin with `junkfile` from `%SystemRoot%\junk` and all of its subdirectories.
* The `REMOVESELF` flag will do the same as `RECURSE`, but also remove any empty directories.

`RegKey`
* Specifies a Windows Registry location to be cleaned.
* Does not support wildcards of any sort.

`ExcludeKey`
* Used to exclude a file, directory or registry path from cleaning, even if otherwise included in a FileKey or RegKey:
  * `ExcludeKey1=FILE|%WinDir%\System32\LogFiles\|myfile.txt` excludes `myfile.txt` in the `%WinDir%\System32\LogFiles` directory from being deleted.
  * `ExcludeKey2=REG|HKCU\Software\Piriform` will prevent keys in `HKEY_CURRENT_USER\Software\Piriform` from being deleted from the Registry.
* Supports wildcards: 
  * `ExcludeKey1=PATH|%WinDir%\System32\LogFiles\SCM\|*-*-*-*.*` excludes all of the files whose name matches the pattern  `*-*-*-*.*` in the `%WinDir%\System32\LogFiles\SCM` directory from being deleted.
  * `ExcludeKey1=PATH|C:\Temp\|*.*` excludes all of the files located in the `C:\Temp` directory and all sub directories from being deleted.
  * `ExcludeKey2=PATH|C:\Windows\|*.exe;*.bat` excludes files of types `.exe` and `.bat` in the `C:\Windows` directory from being deleted.

### Environment variables:

These are all the possible variables that can be used for writing paths in winapp2.ini.
##### Variables marked with a * natively check both 64bit and 32bit locations on 64bit systems.

Variable       | Windows Vista-10 Path | WindowsXP Path
:------------- | :-------------        | :-------------
`%AppData%` | `C:\Users\%UserName%\AppData\Roaming` | `C:\Documents and Settings\%UserName%\Application Data`
`%CommonAppData%` | `C:\ProgramData` | `C:\Documents and Settings\All Users\Application Data`
`%CommonProgramFiles%`* | `C:\Program Files\Common Files` | `C:\Program Files\Common Files`
`%Documents%` | `C:\Users\%UserName%\Documents` | `C:\Documents and Settings\%UserName%\My Documents`
`%LocalAppData%` | `C:\Users\%UserName%\AppData\Local` | `C:\Documents and Settings\%UserName%\Local Settings\Application Data`
`%LoalLowAppData%` | `C:\Users\%UserName%\AppData\LocalLow` | N/A
`%Music` | `C:\Users\%UserName%\Music` | `C:\Documents and Settings\%UserName%\My Documents\My Music`
`%Pictures%` | `C:\Users\%UserName%\Pictures` | `C:\Documents and Settings\%UserName%\My Documents\My Pictures`
`%ProgramFiles%`* | `C:\Program Files` | `C:\Program Files`
`%Public%` | `C:\Users\Public` | N/A
`%SystemDrive%` | `C:` | `C:`
`%UserProfile%` | `C:\Users\%UserName%` | `C:\Documents and Settings\%UserName%`
`%Video%` | `C:\Users\%UserName%\Videos` | `C:\Documents and Settings\%UserName%\My Documents\My Videos`
`%WinDir%` | `C:\Windows` | `C:\Windows`

## Custom.ini

Winapp2.ini does not support non-English system configurations or portable software natively. If you have need for these features, we recommend you utilize a "custom.ini" file, and use winapp2ool to merge it with the main file using the Add&Replace setting to override the existing entries.

## Donations

If you like winapp2.ini and want to help keep development going, please consider donating some money to our great minds behind winapp2.ini

You can donate to Robbie via PayPal (https://paypal.me/moscadotto/) or Venmo (@woadleaves).
