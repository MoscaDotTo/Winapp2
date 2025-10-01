# Winapp2ool

Winapp2ool is a domain-specific companion utility for Winapp2.ini, offering a suite of functions tailored to its structure and purpose, while remaining simple to use. It is designed for use as either a CLI application with a simple menu system, or to be called silently from scripting environments to allow for automation.   

### What is winapp2.ini?
Winapp2.ini is a massive, community-driven database of declarative cleaning routines for Microsoft Windows, compatible CCleaner, BleachBit, System Ninja, R-Wipe&Clean, and HDCleaner. Winapp2.ini has its own readme [here](https://github.com/MoscaDotTo/Winapp2/blob/master/README.md).

### Why winapp2ool?
Winapp2ool was created to help automate otherwise complex or time consuming tasks for maintainers, contributors, and end users of Winapp2.ini. It provides several tools:

* WinappDebug: Performs static analysis and corrects style and syntax errors. 
* Trim: Reduces the database to entries relevant to the current system.
* Transmute: Applies individual structured patches to ini files.
* Flavorize: Applies batches of structured patches to ini files. 
* Diff: Generates changelogs between any two Winapp2.ini versions using context-aware abstraction tracking.
* CCiniDebug: Removes stale Winapp2.ini configurations from the CCleaner settings. 
* Browser Builder: Generates web browser entries for entire families of web browsers.
* Combine: Merges folders (including subfolders) of ini files into a single file. 

Winapp2ool is not just a helper utility: it actively facilitates the build process for Winapp2.ini by assembling the [base entries](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/Entries), [generating and correcting web browser entries](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/BrowserBuilder), and generating each of the [Flavors](https://github.com/MoscaDotTo/Winapp2/tree/master?tab=readme-ov-file#what-are-flavors) of Winapp2.ini.

---

# Requirements 

### Minimum

* Windows Vista SP2
* .NET Framework 4.5
* Administrative permissions (see notes)

### Suggested

* Windows 7 or higher
* .NET Framework 4.6 or higher (for automatically updating the executable)
* Network connection for full functionality

---

# Installation
Download the latest winapp2ool.exe from the [Releases page](https://github.com/MoscaDotTo/Winapp2/releases/) or from the [Release directory](https://github.com/MoscaDotTo/Winapp2/tree/master/winapp2ool/bin/Release).

### Updates
Winapp2ool will prompt you to update it from within the application whenever an update is available. The application will create a backup of itself with the `.bak` extension when doing this. Rename to `.exe` to restore the backed up version.

### Beta builds 
Beta builds are occasionally made available to the public while features are in development. To access the beta build, open the winapp2ool settings and enable beta participation. The newest beta build will automatically be downloaded and launched. 

---

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

---

# Quick Start
The main menu presents some options when an update is available for either Winapp2ool or winapp2.ini 
1. Download the latest winapp2.ini: Choose "Update winapp2.ini"
2. Optimize the latest winapp2.ini for your system: Choose "Update & Trim"
3. See the changelog between your local version and the latest: Choose "Show update Diff"
4. Download the latest Winapp2ool: Choose "Update Winapp2ool" 

# Menu Options

Click on the link in the Options column to see the ReadMe for a particular module 

| Option                                                                                                | Effect                                                                                                          |  Notes                                                                                                                                                       |
|:-                                                                                                     | :-                                                                                                              | :-                                                                                                                                                           |
| Exit        																				            | Exits the application 				 																		  |                                                                                                                                                              |
| WinappDebug																			                | Opens WinappDebug     				 																		  | Enforces thewinapp2.ini style and syntax guidelines                                                                                                          |
| Trim        																				            | Opens Trim            				 																		  | Reduces winapp2.ini to only the set of entries valid for the system at runtime, greatly speeding application load times for tools like CCleaner              |
| [Transmute](https://github.com/MoscaDotTo/Winapp2/blob/master/winapp2ool/modules/transmute/readme.md) | Opens Transmute       				 																	 	  | Provides precise control over modifying configuration files through three primary operations with granular conflict resolution                               |
| Diff																						            | Opens Diff            				 																 		  | Provides context-aware changelogs between any two versions of winapp2.ini                                                                                    |
| CCiniDebug                         														            | Opens CCiniDebug	  				 																	     	  | Removes outdated winapp2.ini rules from ccleaner.ini                                                                                                         |
| Browser Builder 																			            | Opens Browser Builder 				 																		  | Generates winapp2.ini entries for software built on Gecko or Chromium                                                                                        |
| Combine 																				                | Opens Combine         				 																		  | Joins together every ini file in a folder and its subfolders into a single ini file                                                                          |
| Downloader																				            | Opens Downloader     			     																		      | Downloads files from the winapp2 GitHub. **Unavailable in offline mode**                                                                                     |
| Settings 																					            | Opens the winapp2ool global settings            															      | Contains application-level settings for beta mode, saving settings to disk, and more                                                                         |
| Go Online																					            | Attempts to reestablish your network connection 														 	      | **Only available in offline mode**                                                                                                                           |
| Update winapp2.ini																		            | Downloads the latest winapp2.ini from GitHub to the current folder 											  | **Only available alongside a winapp2.ini update, unavailable in offline mode**                                                                               |
| Update & Trim																				            | Downloads the latest winapp2.ini to the current folder and trims it 										      | **Only available alongside a winapp2.ini update, unavailable in offline mode**                                                                               |
| Show winapp2.ini changelog 																            | Diffs your local copy of winapp2.ini against the latest version hosted on GitHub in order to show a changelog   | **Only available alongside a winapp2.ini update, unavailable in offline mode**                                                                               |
| Update Winapp2ool																			            | Attempts to automatically update winapp2ool.exe to the latest version from GitHub  							  | **Only available alongside a winapp2ool update. Unavailable in offline mode, and on machines with .NET Framework 4.5 or lower installed (ie. Winapp2oolXP)** |
# Command-line arguments

Winapp2ool supports command line arguments ("args"). These allow Winapp2ool to be used from a scripting environment (such as a shell script) without having to interact with the UI. There are several top level args which apply settings globally, and then there are tool specific args which will be defined in the respective section for those tools.

The first argument provided should always refer to the module you would like to use, as below. Use of `-` before these args is not required, but it is supported.

### Module args

| Arg                     | Effect                   |
| :-                      | :-                       |
| `1` or `debug` 		  | Launches WinappDebug     |
| `2` or `trim`    		  | Launches Trim            |
| `3` or `transmute`	  | Launches Transmute       |
| `4` or `diff`			  | Launches Diff            |
| `5` or `ccdebug` 		  | Launches CCiniDebug      |
| `6` or `browserbuilder` | Launches Browser Builder |
| `7` or `combine` 		  | Launches Combine         |
| `8` or `download` 	  | Launches Downloader      |
| `9` or `flavorize` 	  | Launches Flavorizer      |

### Global args

| Arg                     | Effect                                                                              | Notes                                                                   | 
| :-                      | :-                                                                                  | :-                                                                      |
| `-s`  			      | Enables "silent mode" - muting almost all output and prompts for input.             | Some exceptions and errors may not be shown when silent mode is enabled |
| `-1d`, `-2d`, ... `-9d` | Defines a new file name and/or path for the module's respectively numbered file. \* | Paths with spaces must be provided in quotes, eg. `-1d "C:\New Folder"` | 
| `-1f`, `-2f`, ... `-9f` | Defines a new file name for the module's respectively numbered file **              |                                                                         |
| `-ncc` or `-base` 	  | Sets the Flavor to Non-CCleaner (base) when downloading                             |                                                                         |            
| `-ccleaner` or `-cc`    | Sets the Flavor to CCleaner when downloading                                        | Default                                                                 | 
| `-bleachbit` or `-bb`   | Sets the Flavor to BleachBit when downloading                                       |                                                                         |  
| `-systemninja` or `-sn` | Sets the Flavor to System Ninja when downloading                                    |                                                                         |
| `-tron`                 | Sets the Flavor to Tron when downloading                                            |                                                                         |

##### Notes

\* The "first file" (`-1d` or `-1f`) in all modules is winapp2.ini. The "third file" (`-3d` or `-3f`) is typically the output file, if one exists. Refer to a specific module's documentation for information on its file configuration

\** You can easily define subdirectories by using the `-f` flag for your file and providing the directory before the file name, eg `-1f \subdir\winapp2.ini`

### Examples

| Args                                                                     | Effect                                                                                                                                                                                       |
| :-                                                                       | :-                                                                                                                                                                                           |
| `winapp2ool.exe -1 -c`												   | Opens and runs WinappDebug with saving of changes enabled                                                                                                                                    |
| `winapp2ool.exe -2 -d -s` 											   | Silently downloads and trims the latest winapp2.ini of your selected flavor (default: `ccleaner`) from GitHub                                                                                |
| `winapp2ool.exe -bleachbit -2 -d -s`                                     | Silently downloads and trims the latest BleachBit flavor of winapp2.ini from GitHub                                                                                                          |                                                                           
| `winapp2ool download winapp2 -s`										   | Silently opens Downloader and downloads the latest winapp2.ini                                                                                                                               |
| `winapp2ool -transmute -remove -bykey -byname -2f key_name_removals.ini` | Sets the Transmute mode to `Remove`, the Removal mode to `By Key`, and the Key Removal mode to `By Name`. Sets the source file name to `key_name_removals.ini` and applies the Transmutation |