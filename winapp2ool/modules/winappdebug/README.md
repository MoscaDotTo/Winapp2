# WinappDebug

WinappDebug is a basic [linter](https://www.wikiwand.com/en/Lint_%28software%29) for winapp2.ini. It performs static analysis on winapp2.ini to ensure and enforce correctness of style and syntax across a wide variety of configurable categories. Optionally and additionally, it will save the linted file back to disk. 

## Menu Options

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



### Valid commandline args & their effects:

* `-1f`,`-1d`: Defines the path for winapp2.ini
* `-2f`,-`2d`: Defines the path for the save file
* `-c`: Enables autocorrecting/saving of corrected errors
