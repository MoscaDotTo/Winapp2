# FluentCleaner Flavor Transmutation Rules

### What is this?

This folder holds the Transmutation rules winapp2ool uses to make the **FluentCleaner** flavor of winapp2.ini. The flavor carries the changes FluentCleaner made to [its own downstream copy](https://raw.githubusercontent.com/builtbybel/FluentCleaner/refs/heads/main/Winapp2.ini) of winapp2.ini.

These rules, followed by the rules in `Sources\Flavors\CCleaner`, produce `winapp2-fluentcleaner.ini`, published as [`Releases/winapp2-fluentcleaner.ini`](https://github.com/MoscaDotTo/Winapp2/blob/master/Releases/winapp2-fluentcleaner.ini).

### What changes are made to create the FluentCleaner flavor?

FluentCleaner turns an entry *on* by default when it has no `Default` key. This flavor adds `Default=False` to the entries FluentCleaner ships turned off, and carries the extra warnings FluentCleaner puts on some entries.

### How it is built

FluentCleaner works from the CCleaner flavor, so the build runs two Flavorize passes in order:

1. The rules in this folder flavorize the base winapp2.ini, making the changes only FluentCleaner needs
2. The rules in the CCleaner folder flavorize that output, which gives every entry its CCleaner `LangSecRef`


The WinappDebug pass after them runs with `-keepdefaults`, so it doesn't remove the `Default=` keys we added in the first pass.

# Files

| File | Contents |
| :- | :- |
| [fc_additions.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/Sources/Flavors/FluentCleaner/fc_additions.ini) | The `[*Name:]` rules that add `Default=False` to browser entries, and the `Warning=` keys for single entries |
| [fc_section_removals.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/Sources/Flavors/FluentCleaner/fc_section_removals.ini) | Empty |
| [fc_name_removals.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/Sources/Flavors/FluentCleaner/fc_name_removals.ini) | Empty |
| [fc_value_removals.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/Sources/Flavors/FluentCleaner/fc_value_removals.ini) | Empty |
| [fc_section_replacements.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/Sources/Flavors/FluentCleaner/fc_section_replacements.ini) | Empty |
| [fc_key_replacements.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/Sources/Flavors/FluentCleaner/fc_key_replacements.ini) | Empty |
