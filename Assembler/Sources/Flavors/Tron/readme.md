# Tron Flavor Transmutation Rules

### What is this?

This folder holds the Transmutation rules winapp2ool uses to make the [Tron flavor](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2-tron.ini) of winapp2.ini.

They produce `winapp2-tron.ini`, published as [`Releases/winapp2-tron.ini`](https://github.com/MoscaDotTo/Winapp2/blob/master/Releases/winapp2-tron.ini). They're applied to the base file.

### What changes are made to create the Tron flavor?

The Tron flavor takes on the changes Tron's maintainers made to their own copy of winapp2.ini. Right now that's only removing Steam's depot cache cleaning from `Steam *`.

# Files

| File | Description |
| :- | :- |
| [tron_additions.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/Sources/Flavors/Tron/tron_additions.ini) | Entries and keys added to the base winapp2.ini |
| [tron_key_replacements.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/Sources/Flavors/Tron/tron_key_replacements.ini) | Keys whose values are replaced |
| [tron_name_removals.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/Sources/Flavors/Tron/tron_name_removals.ini) | Keys removed by name |
| [tron_section_removals.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/Sources/Flavors/Tron/tron_section_removals.ini) | Entries removed entirely |
| [tron_section_replacements.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/Sources/Flavors/Tron/tron_section_replacements.ini) | Entries replaced entirely |
| [tron_value_removals.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/Sources/Flavors/Tron/tron_value_removals.ini) | Keys removed by value |
