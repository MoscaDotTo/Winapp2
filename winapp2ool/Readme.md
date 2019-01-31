# Winapp2ool 

Winapp2ool is a small but robust application designed to take the busy work out of maintaining, managing, downloading, and deploying winapp2.ini. It is designed for use with a network connection, but many functions will work without one by acting on local files.

If winapp2ool is launched without a network connection or .NET Framework 4.6 (or newer) is not installed, some functions and menu options will be unavailable. A prompt will be displayed to the user to retry their connection or update their .NET Framework in these instances.

By default, each tool in the application assumes that local files it is looking for are in the same folder as the executable.

## Requirements & Notes

##### Minimum:
* Windows Vista SP2
* .NET Framework 4.5
##### Suggested:
* Windows 7 or higher
* .NET Framework 4.6 or higher (for automatically updating the executable)

WindowsXP users should use winapp2oolXP, which disables the ability to download the executable, but retains all other functionality.

Winapp2ool requires administrative permissions to run correctly.

## Commandline Switches

Winapp2ool supports command line switches. There are several top level switches which tell winapp2ool how to work and then there are tool specific switches which will be defined in the respective section for those tools.

### Valid commandline args & their effects:

* `1,2,3,4,5,6`,`debug,trim,merge,diff,ccdebug,download`: Calls the module associated with the menu number in the winapp2ool menu. 1 for WinappDebug, 6 for Downloader. All other commands should follow this one. Use of a `-` before each flag is not required here, but is supported.
* `-s`: Actives "silent mode" muting most or all output from the application. When silent mode is active, programs will silently fail if they hit an error during execution. The application will automatically exit after it has finished its task.
* `-1f`, `-1d`: Defines a new file name or path for winapp2.ini to be used during module execution. You can define subdirectories of the current directory easily with the `-f` flag by simply prepending the directory to the file name. eg: `-1f \subdir\winapp2.ini`. File and directory parameters should always immediately follow their flag.
* `-2f`,`-2d`: As above, defines the target for the "second file" in program execution, which is more specifically addressed in the module documentation below.
* `-3f`,`-3d`: As above, defines the target for the "third file", if applicable.

#### Example:

`winapp2ool.exe -1 -c`: Opens winapp2ool and runs WinappDebug with the autocorrect option enabled.  
`winapp2ool.exe -2 -d -s`: Silently opens winapp2ool, downloads the latest winapp2.ini, and trims it.  
`winapp2ool.exe download winapp2 -s`: Silently opens winapp2ool and downloads the latest winapp2.ini.

### Menu Options

* WinappDebug, Trim, Merge, Diff, CCiniDebug, Downloader
  * Opens selected module
  * Downloader unavailable in offline mode

* Go Online
  * Attempts to reestablish your network connection
  * Only shown in offline mode

* Update
  * Downloads the latest winapp2.ini to the current folder
  * Only shown when an update is available for winapp2.ini

* Update & Trim
  * Downloads the latest winapp2.ini to the current folder and trims it
  * Only shown when an update is available for winapp2.ini

* Show update Diff
  * Diffs your local winapp2.ini against the latest version on GitHub to show what has been changed since you last updated
  * Only shown when an update is available for winapp2.ini

* Update
  * Attempts to automatically update winapp2ool.exe to the latest version on GitHub
  * Only shown when an update is available for winapp2ool.exe
  * Unavailable in offline mode
  * Unavailable for WindowsXP and machines running .NET Framework versions older than 4.6




## WinappDebug

WinappDebug is essentially a basic [lint](https://www.wikiwand.com/en/Lint_%28software%29) for winapp2.ini. It will open winapp2.ini and check it for style and syntax errors. Additionally and optionally, the debugger will attempt to automatically correct many of the errors it reports.

### Valid commandline args & their effects:

* `-1f`,`-1d`: Defines the path for winapp2.ini
* `-2f`,-`2d`: Defines the path for the save file
* `-c`: Enables autocorrecting/saving of corrected errors

### Menu Options

* Exit
  * Returns you to the winapp2ool menu

* Run (default)
  * Runs the debugger using your current settings

* File Chooser (winapp2.ini)
  * Opens the interface to select a new local file path for winapp2.ini

* Toggle Autocorrect
  * Enables or disables saving changes made by the debugger back to a local file
  * Disabled by default
  * When enabled, overwrites the input file by default

* File Chooser (save)
  * Opens the interface to select a new target file for saving changes made by the debugger instead of overwriting the input file
    * Suggests "winapp2-debugged.ini" as the default rename
  * Only shown when autocorrection is enabled first

* Toggle Scan Settings
  * Opens the Scan and Repair menu
  * Allows the enabling and disabling of individual scan and repair routines, helpful for diagnosing application crashes or ignoring certain types of errors if they are known

* Reset Settings
  * Restores the tool to its default state, undoing any changes the user may have made
  * Only shown if settings have been modified

### Detected Errors

**Bold** items are correctable

#### General 
* **Duplicate key names or values**
* **Incorrect/Unnecessary key numbering**
* **Incorrect key alphabetization**
* **Forward slash (/) use where there should be a backslash (\\)**
* **Trailing Semicolons (;)**
* **Invalid %EnvironmentVariable% CamelCasing**
* Invalid %EnvironmentVariable% formatting
* Unsupported %EnvironmentVariable% target
* **Leading or trailing whitespace on key names or values**
* **Invalid CamelCasing on winapp2.ini commands**
* Invalid winapp2.ini commands
* No valid Detect / DetectFile / DetectOS / SpecialDetect provided
* No valid FileKeys or RegKeys provided
* ExcludeKeys specified in the absence of corresponding FileKeys or RegKeys
* **No Default key provided**

#### DetectOS
* More than one DetectOS key provided

#### LangSecRef / Section
* More than one (of either) key provided
* Invalid LangSecRef value provided

#### SpecialDetect
* More than one SpecialDetect key provided
* Invalid SpecialDetect provided

#### Detect
* Invalid Registry path provided

#### DetectFile
* **Trailing backslashes (\\)**
* Nested wildcard provided (supported by Trim, not supported by CCleaner)
* Invalid file system path provided

#### Default
* More than one Default key provided
* **Default value other than "False" provided**

#### Warning
* More than one Warning key provided

#### FileKey
* **Duplicate parameters provided in a single FileKey**
* **Empty FileKey parameters**
* Semicolon provided before Pipe symbol
* Incorrect RECURSE/REMOVESELF spelling/too many Pipe symbols
* Incorrect VirtualStore locations
* **Backslash use before Pipe**
* Missing backslash after %EnvironmentVariable%
* **Experimental: Similar FileKey merger**
  * The debugger can attempt to merge FileKeys it thinks can collapse into a single key. Results may vary.

#### RegKey
* Invalid Registry path provided

#### ExcludeKey
* Missing backslash before pipe symbol
* No valid flag (FILE, PATH, REG) provided




## Trim

Trim is designed to do as its name implies: trim winapp2.ini. It does this by processing the detection keys provided in each entry and confirming their existence on the user's machine. Any entries whose detection criteria are invalid for the current system are pruned from the file, resulting in a much smaller file filled only with entries relevant to the current machine. The performance impact of this is most notable on CCleaner which takes much less time to load without the full winapp2.ini.

### Valid commandline args & their effects:
* `-1f`,`-1d`: Overrides the default path for winapp2.ini
* `-2f`,`-2d`: Overrides the default save path for the trimmed file
* `-d`: Enables downloading the latest winapp2.ini to trim
* `-ncc`: Enables downloading the latest non-CCleaner winapp2.ini to trim

### Menu Options

* Exit
  * Returns you to the winapp2ool menu

* Run (default)
  * Runs the trimmer using your current settings

* Toggle Download
  * Enables/Disables downloading of the latest winapp2.ini from GitHub to use as the input for the trimmer
  * Enabled by default if you have an internet connection
  * Unavailable in offline mode

* Toggle Download (non-CCleaner)
  * Toggles downloading the non-CCleaner winapp2.ini
  * CCleaner users should not use this option
  * Disabled by default
  * Unavailable if downloading is disabled
  * Unavailable in offline mode

* File Chooser (winapp2.ini)
  * Opens the interface to choose a new local file path for winapp2.ini
  * Only shown if not downloading

* File Chooser (save)
  * Opens the interface to select a new save location for the trimmed file
    * Suggests "winapp2-trimmed.ini" as the default rename

* Reset Settings
  * Restores the tool to its default state, undoing any changes the user may have made
  * Only shown if settings have been modified




## Merge

This tool is designed to simply merge two local ini files. It has two modes, both of which specify the behavior that should be applied when merging entries of the same name.

### Valid commandline args & their effects:
* `-1f`,`-1d`: Overrides the default path for winapp2.ini
* `-2f`,`-2d`: Sets the merge file name
* `-r`: Uses "removed entries.ini" as the merge file name
* `-c`: Uses "custom.ini" as the merge file name
* `-w`: Uses "winapp3.ini" as the merge file name
* `-a`: Uses "archived entries.ini" as the merge file name
* `-3f`,`-3d`: Overrides the default save path for the merged file
* `-mm`: Switches the Merge Mode from Add&Replace to Add&Remove

### Menu Options

* Exit
  * Returns you to the winapp2ool menu

* Run (default)
  * Runs the merger using your current settings
    * Will not run if you have not selected a merge file name

* Removed Entries, Custom, Winapp3.ini
  * Uses the chosen file name as the file to be merged

* File Chooser (winapp2.ini)
  * Opens the interface to choose a new local winapp2.ini path

* File Chooser (merge)
  * Opens the interface to choose a new local file to be merged

* File Chooser (save)
  * Opens the interface to choose a new save location for the merged file
    * Suggests "winapp2-merged" as the default rename

* Toggle Merge Mode
  * Toggles the merge mode between its two settings
  * Default is Add & Replace
    * Overwrites any entries whose names appear in both the merge file and the save file
  * Toggles to Add & Remove
    * Removes any entries whose names appear in both the merged file and the save file

* Reset Settings
  * Restores the tool to its default state, undoing any changes the user may have made
  * Only shown if settings have been modified




## Diff

### Valid commandline args & their effects:
* `-1f`,`-1d`: Overrides the default path for the older (local) winapp2.ini file
* `-2f`,`-2d`: Sets the path for the newer (local) winapp2.ini file
* `-3f`,`-3d`: Overrides the default save path for the log
* `-d`: Enables downloading the latest winapp2.ini to Diff against
* `-ncc`: Enables downloading the latest non-CCleaner winapp2.ini to Diff against
* `-savelog`: Enables saving the Diff log to disk

### Menu Options

* Exit
  * Returns you to the winapp2ool menu

* Run (default)
  * Runs the differ with your current settings
    * Will not run if you have not chosen a newer/online file

* File Chooser
  * Opens the interface to choose a new local winapp2.ini path

* Winapp2.ini (online)
  * Enables downloading the latest copy of winapp2.ini from GitHub to use as the comparator file
  * Enabled by default if you have an internet connection  

* Winapp2.ini (non-CCleaner)
  * Enables downloading the latest non-CCleaner variant of winapp2.ini to use as the comparator file
  * CCleaner users should not use this option
  * Unavailable if downloading winapp2.ini is not enabled

* Toggle Log Saving
  * Enables or disables saving the diff output to a file at the end of execution
  * Default log name is diff.txt

* File Chooser (log)
  * Opens the interface to choose a new file path for the log file
  * Only shown if saving a log

* Reset Settings
  * Restores the tool to its default state, undoing any changes the user may have made
  * Only shown if settings have been modified




## CCiniDebug

This tool was born of necessity in the advent of a mass renaming of entries in winapp2.ini that left many orphaned settings in CCleaner.ini. Put simply, this tool is a small debugger for CCleaner.ini that will remove any orphaned entry settings left over by removed winapp2.ini keys.

### Valid commandline args & their effects:
* `-1f`,`-1d`: Overrides the default path for winapp2.ini
* `-2f`,`-2d`: Overrides the default path for ccleaner.ini
* `-3f`,`-3d`: Overrides the default save path for the debugged ccleaner.ini
* `-noprune`: Disable pruning stale winapp2.ini entries from ccleaner.ini
* `-nosort`: Disable sorting the Options section of ccleaner.ini
* `-nosave`: Disable saving the changes made during debug back to disk

### Menu Options

* Exit
  * Returns you to the winapp2ool menu

* Run (default)
  * Runs the debugger using your current settings
    * Will not run if all three settings are disabled

* Toggle Pruning
  * Enables or disables the pruning of stale winapp2.ini entry settings
  * Default is enabled

* Toggle Saving
  * Enables or disables saving the changes made to ccleaner.ini back to disk
  * Default is enabled

* Toggle Sorting
  * Enables or disables sorting the keys in CCleaner.ini alphabetically
  * Default is enabled

* File Chooser (ccleaner.ini)
  * Opens the interface to choose a new ccleaner.ini path

* File Chooser (winapp2.ini)
  * Opens the interface to choose a new winapp2.ini path
  * Only shown if pruning is enabled

* File Chooser (save)
  * Opens the interface to choose a new save file path instead of overwriting ccleaner.ini
    * Suggests "ccleaner-debugged.ini" as the default rename
  * Only shown if saving is enabled

* Reset Settings
  * Restores the tool to its default state, undoing any changes the user may have made
  * Only shown if settings have been modified




## Downloader

### Valid commandline args & their effects:
* `1` or `winapp2`: Downloads winapp2.ini
* `2`: Downloads the non-CCleaner winapp2.ini
* `3`or `winapp2ool`: Downloads winapp2ool.exe
* `4`or `removed`: Downloads removed entries.ini
* `5`or `winapp3`: Downloads winapp3.ini
* `6`or `archived`: Downloads archived entries.ini
* `7`or `java`: Downloads java.ini
* `8`or `readme`: Downloads the winapp2ool readme

### Menu Options

* Exit
  * Returns you to the winapp2ool menu

* Winapp2.ini
  * Downloads the latest winapp2.ini from GitHub

* Non-CCleaner
  * Downloads the latest non-CCleaner winapp2.ini from GitHub

* Winapp2ool
  * Downloads winapp2ool.exe from GitHub
  * Will attempt to replace the currently running executable
  * Unavailable on winapp2oolXP and on machines running .NET Framework versions older than 4.6

* Removed Entries.ini
  * Downloads the file used to create the non-CCleaner variant of winapp2.ini

* Directory
  * Opens the interface to allow you to choose a new download directory

* Advanced
  * Opens the menu that allows you to download winapp3.ini
