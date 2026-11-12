# CCleaner Flavor Transmutation Rules

### What is this?

This folder holds the Transmutation rules winapp2ool uses to make the [CCleaner flavor](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2-ccleaner.ini) of winapp2.ini.

They produce `winapp2-ccleaner.ini`, published as [`Releases/winapp2-ccleaner.ini`](https://github.com/MoscaDotTo/Winapp2/blob/master/Releases/winapp2-ccleaner.ini). The CCleaner 7 flavor is built from the CCleaner flavor, and the FluentCleaner build runs these rules as its second pass, so `winapp2-ccleaner7.ini` and `winapp2-fluentcleaner.ini` carry them too.

### What changes are made to create the CCleaner flavor?

Mostly recategorization. CCleaner has its own section for some browsers, and for each of those, a `[*Map:]` rule in `cc_key_replacements.ini` moves the browser's entries out of the section winapp2ool generates for it (`Section=Brave Web Browser`) and into CCleaner's (`LangSecRef=3034`). A rule matches every entry in its browser's generated section, so a new entry for that browser needs no change here. Browsers without a CCleaner section keep their generated one. The flavor also moves Spotify's two entries into CCleaner's own Spotify section (`LangSecRef=3036`).

# Files

| File | Description |
| :- | :- |
| [cc_additions.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/Sources/Flavors/CCleaner/cc_additions.ini) | Entries and keys added to the base winapp2.ini |
| [cc_key_replacements.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/Sources/Flavors/CCleaner/cc_key_replacements.ini) | Keys whose values are replaced, including the browser `[*Map:]` rules |
| [cc_name_removals.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/Sources/Flavors/CCleaner/cc_name_removals.ini) | Keys removed by name |
| [cc_section_removals.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/Sources/Flavors/CCleaner/cc_section_removals.ini) | Entries removed entirely |
| [cc_section_replacements.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/Sources/Flavors/CCleaner/cc_section_replacements.ini) | Entries replaced entirely |
| [cc_value_removals.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/Sources/Flavors/CCleaner/cc_value_removals.ini) | Keys removed by value |
