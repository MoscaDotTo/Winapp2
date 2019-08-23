# WinappDebug

WinappDebug is a basic [linter](https://www.wikiwand.com/en/Lint_%28software%29) for winapp2.ini. It performs static analysis on winapp2.ini to ensure and enforce correctness of style and syntax across a wide variety of configurable categories. Optionally and additionally, it will save the linted file back to disk.

## Menu Options

Option|Effect|Notes
:-|:-|:-
Exit|Returns you to the winapp2ool menu|
Run|Runs the tool using your current settings|Default option
File Chooser (winapp2.ini)|Opens the interface to select a new local file for winapp2.ini|Default file path is winapp2.ini in the current working directory
Toggle Saving|Toggles the saving of corrected errors back to disk|Disabled by default
File Chooser (save)|Opens the interface to select a different local path to save changes too|Overwrites the input file by default. <br/> Only shown when saving is enabled.
Toggle Scan Settings|Opens the interface to toggle individual scans and/or repairs|See "Scan Settings" section below
Reset Settings|Restores settings to their default states, undoing any changes the user may have made|Only shown if a setting has or may have been changed (ie. a File Chooser was opened)
Log Viewer|Shows a summary of the last analysis|Only shown if errors are detected

## Scan Settings Menu Options

Option|Type of Scan/Repair controlled
:-|:-
Casing|Instances of improper CamelCasing
Alphabetization|Instances of improper alphabetization
Improper Numbering|Instances of numbered keys having incorrect values
Parameters|Instances of FileKeys having errors in their parameters
Flags|Instances of incorrect flag masks in FileKeys and ExcludeKeys
Slashes|Instances of improper uses of slashes
Defaults|Instances of keys lacking a "Default=False" key
Duplicates|Instances of duplicate key names or values in a single entry
Unneeded Numbering|Instances of keys having numbers that they do not need or should not have
Multiples|Instances of singleton keys occurring more than once
Invalid Values|Instances of invalid values for certain keyTypes
Syntax Errors|Instances of entries whose configuration may not run in CCleaner
Path Validity|Instances of invalid file system or registry locations
Semicolons|Instances of improperly used semicolons (;)
Optimizations|Instances where FileKeys may possibly be merged **(experimental, disabled by default)**

### Detected Errors

**Bold** items are correctable
*italicised* items are covered by tests

#### General

* ***Duplicate key names or values***
* ***Incorrect/Unnecessary key numbering***
* ***Incorrect key alphabetization***
* ***Forward slash (/) use where there should be a backslash (\\)***
* ***Multiple backslashes (\\\\)***
* ***Trailing Semicolons (;)***
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
* **Missing = between iniKey Name and Value**
  * ie FileKey1%WinDir%\tmp|\*.\*

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

## Command-line args

|Arg|Effect|
|:-|:-
`-1f` or `-1d`|Defines the path for winapp2.ini
`-2f` or `-2d`|Defines the path for the save file
`-c`|Enables saving
