# WinappDebug

### What is this?
WinappDebug is a console tool for detecting errors in your winapp2.ini file, both syntactical and stylistic. 

Detected syntax errors include:
* Improperly numbered or misspelled commands 
* Improperly formatted RECURSE and REMOVESELF
* Improperly formatted %EnvironmentVariables%
* Missing pipes or equalities ( '|' or '=' )
* Improper targeting for registry & file specific commands (eg. RegKey points to a filesystem location or FileKey points to a registry location)
* Nested Wildcards in DetectFiles
* Trailing backslashes in DetectFiles
* VirtualStore directory formatting

Detected style errors include:
* Improper alphabetization ordering 
* Improper CamelCasing 
* Entries that are enabled by default
* Leading and trailing whitespace
* Trailing whitspace and semicolons 
* Duplicate commands on different lines

### How to use

* Place WinappDebug.exe in a folder with Winapp2.ini
* Run WinappDebug.exe

The tool will run and generate as its output a description of any errors it finds, including the line number on which the error was found.
