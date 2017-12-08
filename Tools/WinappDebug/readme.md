#WinappDebug

###What is this?
WinappDebug is a console tool for detecting errors in your winapp2.ini file, both syntactical and stylistic. 

Detected syntax errors include:
* Improperly numbered or misspeled commands 
* Improperly formatted RECURSE and REMOVESELF
* Missing pipes or equalities ( '|' or '=' )
* Missing backslashes after %EnvironmentVariables%

Detected style errors include:
* Improper alphabetization ordering 
* Improper CamelCasing 
* Entries that are enabled by default


### How to use

* Place WinappDebug.exe in a folder with Winapp2.ini
* Run WinappDebug.exe

The tool will run and generate as its output a description of any errors it finds, including the line number on which the error was found.
