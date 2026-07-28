# CCleaner 7 Flavor Transmutation Rules 

### What is this? 

This folder contains the set of Transmutation rules required by Winapp2ool to produce the [CCleaner 7 flavor](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Non-CCleaner/CCleaner7/Winapp2.ini) of Winapp2.ini 

### What changes are made to create the CCleaner 7 flavor? 

CCleaner 7 dropped support for reading a separate winapp2.ini and requires additional keys on every cleaning definition, while replacing the `LangSecRef`/`Section` categorization system with a comma-delimited `Tags` string. The rules in this folder are applied to the **CCleaner flavor** of Winapp2.ini (not the base) and perform that format conversion on every entry:

* Each entry's `LangSecRef=` or `Section=` category key is replaced with the corresponding `Tags=` key via global `[*Map:]` rules. Categories without a specific mapping funnel through a wildcard fallback rule into `Tags=ccapps` 
* `ID=<entry name>` and `Author=Winapp2.ini Project` are added to every entry via a global `[*]` addition using the `%EntryName%` token 

Any additional modifications are performed on a per-entry basis and are intended to leverage CCleaner 7's extended feature set. 

The resulting file is consumed by Winapp2ool's CC7Patcher module, which appends the entries into CCleaner 7's `ccleaner.ini` 

Note: rules for categories with no current entries emit a "matched nothing" warning during the build. This is expected for dormant categories and is not an error 

# Files 
| File                                                                                                                                                   | Description                                                                                                                 |
| :-                                                                                                                                                     | :-                                                                                                                          |
| [cc7_additions.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/CCleaner7/cc7_additions.ini)                      | Contains the global `[*]` addition which gives every entry its `ID` and `Author` keys                                       |
| [cc7_key_replacements.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/CCleaner7/cc7_key_replacements.ini)        | Contains the global `[*Map:]` rules which convert each entry's category key into its CCleaner 7 `Tags` key                  |
