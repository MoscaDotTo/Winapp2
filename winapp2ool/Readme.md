# Winapp2ool

Winapp2ool is a small but robust application designed to take the busy work out of maintaining, managing, downloading, and deploying winapp2.ini. 

### What is winapp2.ini?
Winapp2.ini is a community-maintained declarative set of application cleaning rules for system cleaners such as CCleaner, BleachBit, System Ninja, and others. Winapp2.ini has its own readme [here](https://github.com/MoscaDotTo/Winapp2/blob/master/README.md)

### Why winapp2ool?
Winapp2ool was created to help automate otherwise complex or time consuming tasks for both maintainers of winapp2.ini and also end users. It provides tools for generating entries, static analysis, and applying "Flavors" onto winapp2.ini with a series of transmutations to help simplify the task of maintaining the file or creating variants for particular systems or toolsets. Likewise, it provides tools useful to the end user such as Downloading the latest file, Trim and Diff, which tailor the database to the current system and perform complex semantic analysis (like a  ["Diff"](https://en.wikipedia.org/wiki/Diff) with context and the ability to capture abstractions) between any two versions of winapp2.ini (including trimmed ones), always providing a detailed changelog which makes clear what has changed and how. 

# Requirements 

### Minimum

* Windows Vista SP2
* .NET Framework 4.5
* Administrative permissions (see notes)

### Suggested

* Windows 7 or higher
* .NET Framework 4.6 or higher (for automatically updating the executable)
* Network connection for full functionality

# Installation
Download the latest winapp2ool.exe from the [Releases page](https://github.com/MoscaDotTo/Winapp2/releases/)

### Updates
Winapp2ool will prompt you to update it from within the application whenever an update is available. The application will create a backup of itself with the `.bak` extension when doing this. Rename to `.exe` to restore the backed up version

### Beta builds 
Beta builds are occasionally made available to the public while features are in development. To access the beta build, open the winapp2ool settings and enable beta participation. The newest beta build will automatically be downloaded and launched. 

# Notes

### General
.NET Framework 4.5 (or newer) comes pre-installed by default on Windows 8 and newer.   

By default, each tool in the application assumes that local files it is looking for are in the same folder as the executable. 

The parent directory of the current directory is abbreviated in the menu as ".."

Winapp2ool performs queries against protected system areas such as the Program Files and Windows directories and will return invalid results if run without administrative permissions.

Winapp2ool does not perform any automatic backup of ini files before modifying them. Data loss may be possible when misconfigured. 

### Windows XP
WindowsXP users should use winapp2oolXP. 
Winapp2oolXP is no longer maintained and no longer receives updates, but retains the ability to download and trim the latest winapp2.ini for users on that platform. We are unable to provide application support for users of Winapp2oolXP. 

### Limited functionality
All functions not directly related to downloading files will work without a network connection by acting on local files.

If launched without a network connection, a prompt will be displayed to the user to retry their connection. Some menu options are unavailable without a network connection 

Winapp2ool requires .NET 4.6 or newer to provide updates to itself. A warning will be displayed in the main menu when .NET 4.6 or higher is not detected.

Winapp2ool uses the %tmp% directory as a staging ground for its downloads and updates, when launched from this directory, the executable will not update itself. A warning will be displayed in the main menu when launched from %tmp%

# Menu Options

### Quick Start
The main menu presents some options when an update is available for either Winapp2ool or winapp2.ini 
1. Download the latest winapp2.ini: Choose "Update winapp2.ini"
2. Optimize the latest winapp2.ini for your system: Choose "Update & Trim"
3. See the changelog between your local version and the latest: Choose "Show update Diff"
4. Download the latest Winapp2ool: Choose "Update Winapp2ool" 

### Menu


|Option|Effect|Notes
:-|:-|:-
Exit | Exits the application
WinappDebug|Opens WinappDebug  | Enforces thewinapp2.ini style and syntax guidelines 
Trim|Opens Trim  | Reduces winapp2.ini to only the set of entries valid for the system at runtime, greatly speeding application load times for tools like CCleaner 
Transmute|Opens Transmute | Applies "patches" to an ini file with 3 discrete settings: Add, Remove, and Replace, also contains the Flavorize tool
Diff|Opens Diff | Provides context-aware changelogs between any two versions of winapp2.ini 
CCiniDebug|Opens CCiniDebug | Removes outdated winapp2.ini rules from ccleaner.ini 
Browser Builder | Opens Browser Builder | Generates winapp2.ini entries for software built on Gecko or Chromium
Combine | Opens Combine | Joins together every ini file in a folder and its subfolders into a single ini file 
Downloader|Opens Downloader | Downloads files from the winapp2 GitHub. **Unavailable in offline mode**
Settings | Opens the winapp2ool global settings | Contains application-level settings for beta mode, saving settings to disk, and more 
Go Online|Attempts to reestablish your network connection.| **Only available in offline mode**
Update winapp2.ini|Downloads the latest winapp2.ini from GitHub to the current folder| **Only available alongside a winapp2.ini update, unavailable in offline mode**
Update & Trim|Downloads the latest winapp2.ini to the current folder and trims it|**Only available alongside a winapp2.ini update, unavailable in offline mode**
Show winapp2.ini changelog |Diffs your local copy of winapp2.ini against the latest version hosted on GitHub in order to show a changelog| **Only available alongside a winapp2.ini update, unavailable in offline mode**
Update Winapp2ool|Attempts to automatically update winapp2ool.exe to the latest version from GitHub |**Only available alongside a winapp2ool update. Unavailable in offline mode, and on machines with .NET Framework 4.5 or lower installed (ie. Winapp2oolXP)**

# Command-line arguments

Winapp2ool supports command line arguments ("args"). These allow Winapp2ool to be used from a scripting environment (such as a shell script) without having to interact with the UI. There are several top level args which apply settings globally, and then there are tool specific args which will be defined in the respective section for those tools.

The first argument provided should always refer to the module you would like to use, as below. Use of `-` before these args is not required, but it is supported.

### Module args

|Arg|Effect|
|:-|:-|
|`1` or `debug`|Launches WinappDebug
|`2` or `trim`|Launches Trim
|`3` or `transmute`|Launches Transmute
|`4` or `diff`|Launches Diff
|`5` or `ccdebug`|Launches CCiniDebug
|`6` or `browserbuilder`|Launches Browser Builder
|`7` or `combine` | Launches Combine 
|`8` or `download` | Launches Downloader
|`9` or `flavorize` | Launches Flavorizer

### Other top level args

|Arg|Effect|
|:-|:-|
`-s`|Enables "silent mode" - muting almost all output and prompts for input. Exceptions and errors may not be shown when silent mode is enabled
`-1d`, `-2d`, ... `-9d`|Defines a new file name and/or path for the module's respectively numbered file. *
`-1f`, `-2f`, ... `-9f`| Defines a new file name for the module's respectively numbered file **
`-ncc`|Enables "Non-CCleaner" mode and sets the Non-CCleaner ini as the default

##### Notes

\* The "first file" (`-1d` or `-1f`) in all modules is winapp2.ini. The "third file" (`-3d` or `-3f`) is typically the output file, if one exists. Refer to a specific module's documentation for information on its file configuration

\** You can easily define subdirectories by using the `-f` flag for your file and providing the directory before the file name, eg `-1f \subdir\winapp2.ini`

### Examples

Args|Effect|
|:-|:-|
|`winapp2ool.exe -1 -c`|Opens and runs WinappDebug with saving of changes enabled
|`winapp2ool.exe -2 -d -s`|Silently downloads and trims the latest winapp2.ini from GitHub
|`winapp2ool download winapp2 -s`|Silently opens Downloader and downloads the latest winapp2.ini
|`winapp2ool -transmute -remove -bykey -byname -2f key_name_removals.ini` | Sets the Transmute mode to `Remove`, the Removal mode to `By Key`, and the Key Removal mode to `By Name`. Sets the source file name to `key_name_removals.ini` and applies the Transmutation  