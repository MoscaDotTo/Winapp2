# CCleaner 7 Flavor Transmutation Rules

### What is this?

This folder holds the Transmutation rules winapp2ool uses to make the [CCleaner 7 flavor](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2-ccleaner7.ini) of winapp2.ini.

Applied to the CCleaner flavor, they produce `winapp2-ccleaner7.ini`, published as [`Releases/winapp2-ccleaner7.ini`](https://github.com/MoscaDotTo/Winapp2/blob/master/Releases/winapp2-ccleaner7.ini).

### What changes are made to create the CCleaner 7 flavor?

CCleaner 7 no longer reads a separate winapp2.ini. Its cleaning definitions need extra keys, and it replaced the `LangSecRef`/`Section` categories with a comma-separated `Tags` string. These rules run on the **CCleaner flavor** of winapp2.ini (not the base) and convert every entry to that format:

* Global `[*Map:]` rules replace each entry's `LangSecRef=` or `Section=` with the matching `Tags=` key. A wildcard rule at the end catches every category without its own mapping and gives it `Tags=ccapps`
* A global `[*]` addition gives every entry `ID=<entry name>` and `Author=Winapp2.ini Project`, using the `%EntryName%` token

Changes for single entries, to use CCleaner 7's extra features, would go in the other four files, which are empty for now.

winapp2ool's CC7Patcher module downloads the result and adds its entries to CCleaner 7's `ccleaner.ini`.

Rules for categories that have no entries right now warn that they matched nothing during the build. That's expected, and isn't an error.

# Files

| File | Description |
| :- | :- |
| [cc7_additions.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/Sources/Flavors/CCleaner7/cc7_additions.ini) | The global `[*]` addition that gives every entry its `ID` and `Author` keys |
| [cc7_key_replacements.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/Sources/Flavors/CCleaner7/cc7_key_replacements.ini) | The global `[*Map:]` rules that turn each entry's category key into its CCleaner 7 `Tags` key |
| [cc7_name_removals.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/Sources/Flavors/CCleaner7/cc7_name_removals.ini) | Empty |
| [cc7_section_removals.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/Sources/Flavors/CCleaner7/cc7_section_removals.ini) | Empty |
| [cc7_section_replacements.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/Sources/Flavors/CCleaner7/cc7_section_replacements.ini) | Empty |
| [cc7_value_removals.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/Sources/Flavors/CCleaner7/cc7_value_removals.ini) | Empty |
