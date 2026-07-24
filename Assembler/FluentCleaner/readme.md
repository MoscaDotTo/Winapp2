# FluentCleaner Flavor Transmutation Rules

### What is this?

This folder contains the set of Transmutation rules Winapp2ool uses to produce the **FluentCleaner** flavor of Winapp2.ini. These Transmutation rules produce a winapp2.ini flavor that captures the changes made to [FluentCleaner's downstream copy](https://raw.githubusercontent.com/builtbybel/FluentCleaner/refs/heads/main/Winapp2.ini) of winapp2.ini

### What changes are made to create the FluentCleaner flavor?

In FluentCleaner, winapp2.ini entries are *enabled* by default (ie. when there is no `Default` key). This flavor inserts `Default=False` keys into entries that FluentCleaner ships as disabled by default. Additionally, FluentCleaner carries some additional warnings on some entries and those are present in this flavor. 

### How it is built

FluentCleaner operates off the CCleaner flavor, so the build runs two sequenced Flavorize passes:

1. The rules in this folder flavorize the base winapp2.ini, making the FluentCleaner exclusive changes
2. The rules in the CCleaner folder flavorize pass 1's output with the full CCleaner flavor to LangSecRef-categorize the result.

The subsequent WinappDebug pass is run with `-keepdefaults`, to prevent WinappDebug from removing the `Default=` keys we added as part of the first pass.  

# Files

| File | Contents |
| :- | :- |
| [fc_additions.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/FluentCleaner/fc_additions.ini) | The global mapping rules that add `Default=False` to the browser entries and `Warning=` keys for individual entries |
| [fc_section_removals.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/FluentCleaner/fc_section_removals.ini) | Empty  |
| [fc_name_removals.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/FluentCleaner/fc_name_removals.ini) | Empty |
| [fc_value_removals.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/FluentCleaner/fc_value_removals.ini) | Empty |
| [fc_section_replacements.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/FluentCleaner/fc_section_replacements.ini) | Empty |
| [fc_key_replacements.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/FluentCleaner/fc_key_replacements.ini) | Empty |
