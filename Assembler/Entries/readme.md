# Winapp2.ini Base Entries 

### What is this?

This folder contains the set of base entries for winapp2.ini, delineated into 27 different files alphabetically. Each file contains all of the entries in Winapp2.ini which begin with its particular letter (numbers and symbols can be found in `#.ini`), and also do not target web browsers or UWP applications.

These files are assembled by Winapp2ool as part of the Winapp2.ini build process. Contributions to Winapp2.ini should be made against these files and not Winapp2.ini itself. See the [contributor guidelines](https://github.com/MoscaDotTo/Winapp2/blob/master/CONTRIBUTING.md) for how to write and submit an entry.

### What about the EntryBuilder folder?

Some base entries have been migrated into optional shorthand source files in [`Assembler/EntryBuilder/`](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/EntryBuilder). Each entry lives in exactly one of the two folders, and standard winapp2.ini syntax is valid in both. See the [contributor guidelines](https://github.com/MoscaDotTo/Winapp2/blob/master/CONTRIBUTING.md) for details.

### What about web browser entries?

Web browser entries are generated through the Browser Builder process. Contributions to browser entries should be made against the files in [this folder](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/BrowserBuilder).

### What about UWP applications?

UWP Application entries are generated through the UWPBuilder process. Contributions to UWP entries should be made against the files in [this folder](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/UWP/AppInfo).