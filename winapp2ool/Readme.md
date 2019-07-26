
# Winapp2ool 

Winapp2ool is a small but robust application designed to take the busy work out of maintaining, managing, downloading, and deploying winapp2.ini. 

## Requirements & Notes

##### Minimum:
* Windows Vista SP2
* .NET Framework 4.5
##### Suggested:
* Windows 7 or higher
* .NET Framework 4.6 or higher (for automatically updating the executable)

WindowsXP users should use winapp2oolXP, which disables the ability to download the executable, but retains all other functionality.

Winapp2ool requires administrative permissions to run correctly.

It is designed for use with a network connection, but most functions will work without one by acting on local files.

If winapp2ool is launched without a network connection, or .NET Framework 4.6 (or newer) is not installed, some functions and menu options will be unavailable. A prompt will be displayed to the user to retry their connection or update their .NET Framework in these instances.

By default, each tool in the application assumes that local files it is looking for are in the same folder as the executable.

## Menu Options

|Option|Effect
:-|:-
WinappDebug|Opens WinappDebug, a static analysis tool for winapp2.ini
Trim|Opens Trim, a tool for tailoring winapp2.ini to a user's system 
Merge|Opens Merge, a tool for merging and removal operations between multiple ini files
Diff|Opens Diff, a tool for generating Diffs between two versions of an ini file
CCiniDebug|Opens CCiniDebug, a tool for cleaning up ccleaner.ini
Downloader|Opens Downloader, a tool for downloading files from GitHub. **Unavailable in offline mode**
Go Online|Attempts to reestablish your network connection. **Only available in offline mode**
Update|Downloads the latest winapp2.ini from GitHub to the current folder. **Only available alongside an available update**
Update & Trim|Downloads the latest winapp2.ini to the current folder and trims it. **Only available alongside an available update**
Show Update Diff|Diffs your local copy of winapp2.ini against the latest version hosted on GitHub in order to show a changelog. **Only available alongside an available update**
Update|Attempts to automatically update winapp2ool.exe to the latest version from GitHub. ****Only available alongside an available update for winap2ool. Unavailable in offline mode, or on machines with .NET Framework 4.5 or lower installed (ie. Winapp2oolXP)** 



## Command-line arguments 

Winapp2ool supports command line arguments ("args"). There are several top level args which apply settings globally, and then there are tool specific args which will be defined in the respective section for those tools.

The first argument provided should always refer to the module you would like to use, as below. Use of `-` before these args is not required, but it is supported.

|Arg|Effect|
|:-|:-|
|`1` or `debug`|Launches WinappDebug
|`2` or `trim`|Launches Trim
|`3` or `merge`|Launches Merge
|`4` or `diff`|Launches Diff
|`5` or `ccdebug`|Launches CCiniDebug
|`6` or `download`|Launches Downloader

#### Other top level args 
|Arg|Effect|
|:-|:-|
`-s`|Enables "silent mode" - muting almost all output and prompts for input. Exceptions may not be shown when silent mode is enabled
`-1d`, `-2d`, or `-3d`|Defines a new file name and/or path for the module's respectively numbered file. *
`-1f`, `-2f`, or `-3f`| Defines a new file name for the module's respectively numbered file **

##### Notes
\* The "first file" (`-1d` or `-1f`) in all modules is winapp2.ini. Refer to a specific module's documentation for information on its other files.
\** You can easily define subdirectories by using the `-f` flag for your file and providing the directory before the file name, eg `-1f \subdir\winapp2.ini` 
 
#### Example:
Args|Effect|
|:-|:-|
|`winapp2ool.exe -1 -c`|Opens and runs WinappDebug with saving of changes enabled
|`winapp2ool.exe -2 -d -s`|Silently opens Trim and trims the latest winapp2.ini from GitHub
|`winapp2ool download winapp2 -s`|Silently opens Downloader and downloads the latest winapp2.ini


##### Module code documentation below
# Winapp2ool


# Properties
Name|Type|Default Value|Description
:-|:-|:-|:-
RemoteWinappIsNonCC|`Boolean`|`False`|Indicates that winapp2ool is in "Non-CCleaner" mode and should collect the appropriate ini from GitHub
DotNetFrameworkOutOfDate|`Boolean`|`False`|Indicates that the .NET Framework installed on the current machine is below the targeted version (.NET Framework 4.5)
isOffline|`Boolean`|`False`|Indicates that winapp2ool currently has access to the internet
isBeta|`Boolean`|`False` on Release, `True` on Beta|Indicates that this build is beta and should check the beta branch link for updates

## Helper functions

### getWinVer
#### Checks the version of Windows on the current system and returns it as a Double
```vb
Public Function getWinVer() As Double
        Dim osVersion = System.Environment.OSVersion.ToString().Replace("Microsoft Windows NT ", "")
        Dim ver = osVersion.Split(CChar("."))
        Dim out = Val($"{ver(0)}.{ver(1)}")
        Return out
    End Function
 ```

### getFirstDir
#### Returns the first portion of a registry or filepath parameterization
```vb
Public Function getFirstDir(val As String) As String
    Return val.Split(CChar("\"))(0)
End Function
```
Parameter|Type|Description
:-|:-|:-
val|`String`|A Windows filesystem or registry path from which the root should be returned

### enforceFileHasContent
#### Ensures that an iniFile has content and informs the user if it does not. Returns false if there are no sections

Parameter|Type|Description
:-|:-|:-
iFile|`iniFile`|An iniFile to be checked for content