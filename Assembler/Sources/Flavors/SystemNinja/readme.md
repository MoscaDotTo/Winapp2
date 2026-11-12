# System Ninja Flavor Transmutation Rules

### What is this?

This folder holds the Transmutation rules winapp2ool uses to make the [System Ninja flavor](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2-systemninja.rules) of winapp2.ini.

They produce `winapp2-systemninja.rules`, published as [`Releases/winapp2-systemninja.rules`](https://github.com/MoscaDotTo/Winapp2/blob/master/Releases/winapp2-systemninja.rules).

### What changes are made to create the System Ninja flavor?

System Ninja doesn't support wildcards in `DetectFile` keys. Each `[*Map:]` rule in `sn_key_replacements.ini` replaces one wildcard `DetectFile` with one or more that don't need it, so System Ninja can use those entries. Most of the rules are for browsers, where one rule fixes every generated entry for that browser, and the rest are for a few games and apps.

Work on this flavor is ongoing, and some entries still carry `DetectFile` keys with wildcards and will not work.

The flavor also removes every entry's `LangSecRef` and `Section` key, with a global `[*]` rule in `sn_name_removals.ini`. System Ninja doesn't do anything with these keys.

# Files

| File | Description |
| :- | :- |
| [sn_additions.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/Sources/Flavors/SystemNinja/sn_additions.ini) | Entries and keys added to the base winapp2.ini |
| [sn_key_replacements.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/Sources/Flavors/SystemNinja/sn_key_replacements.ini) | Keys whose values are replaced, including the `DetectFile` `[*Map:]` rules |
| [sn_name_removals.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/Sources/Flavors/SystemNinja/sn_name_removals.ini) | Keys removed by name, including the `[*]` rule that removes categories |
| [sn_section_removals.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/Sources/Flavors/SystemNinja/sn_section_removals.ini) | Entries removed entirely |
| [sn_section_replacements.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/Sources/Flavors/SystemNinja/sn_section_replacements.ini) | Entries replaced entirely |
| [sn_value_removals.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/Sources/Flavors/SystemNinja/sn_value_removals.ini) | Keys removed by value |
