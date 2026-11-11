# Tron Flavor Transmutation Rules 

### What is this? 

This folder contains the set of Transmutation rules required by Winapp2ool to produce the [Tron flavor](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2-tron.ini) of Winapp2.ini 

These rules produce `winapp2-tron.ini`, published as [`Releases/winapp2-tron.ini`](https://github.com/MoscaDotTo/Winapp2/blob/master/Releases/winapp2-tron.ini). They are applied to the CCleaner flavor, not to the base file, so `winapp2-tron.ini` carries the CCleaner rules too.

### What changes are made to create the Tron flavor? 

The Tron flavor is created by adopting the changes made downstream by the Tron maintainers to their copy of Winapp2.ini. Currently, this is only the removal of the Depot Cache cleaning from `Steam`

# Files 
| File                                                                                                                                                 | Description                                                                                                  |
| :-                                                                                                                                                   | :-                                                                                                            |
| [tron_additions.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/Sources/Flavors/Tron/tron_additions.ini)                       | Contains the set of entries and keys which are added to the base Winapp2.ini to create the Tron flavor       |
| [tron_key_replacements.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/Sources/Flavors/Tron/tron_key_replacements.ini)         | Contains the set of key replacement values made to the base Winapp2.ini to create the Tron flavor            |
| [tron_name_removals.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/Sources/Flavors/Tron/tron_name_removals.ini)               | Contains the set of keys to be removed by their name from the base Winapp2.ini to create the Tron flavor     |
| [tron_section_removals.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/Sources/Flavors/Tron/tron_section_removals.ini)         | Contains the set of sections to be removed by their name from the base Winapp2.ini to create the Tron flavor | 
| [tron_section_replacements.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/Sources/Flavors/Tron/tron_section_replacements.ini) | Contains the set of section replacements made to the base Winapp2.ini to create the Tron flavor              | 
| [tron_value_removals.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/Sources/Flavors/Tron/tron_value_removals.ini)             | Contains the set of key values to be removed from the base Winapp2.ini to create the Tron flavor             | 