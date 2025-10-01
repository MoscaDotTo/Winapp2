# System Ninja Flavor Transmutation Rules 

### What is this? 

This folder contains the set of Transmutation rules required by Winapp2ool to produce the [System Ninja flavor](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Non-CCleaner/SystemNinja/Winapp2.rules) of Wianpp2.ini 

### What changes are made to create the System Ninja flavor? 

System Ninja does not support wildcards in `DetectFile` keys. The System Ninja flavor is created by removing `DetectFile` keys containing wildcards and replacing them with ones which do not. This enables System Ninja to properly utilize many Winapp2.ini cleaning rules which it otherwise would not.  

# Files 
| File                                                                                                                                                    | Description                                                                                                          |
| :-                                                                                                                                                      | :-                                                                                                                   |
| [sn_additions.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/SystemNinja/sn_additions.ini)                       | Contains the set of entries and keys which are added to the base Winapp2.ini to create the System Ninja flavor       |
| [sn_key_replacements.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/SystemNinja/sn_key_replacements.ini)         | Contains the set of key replacement values made to the base Winapp2.ini to create the System Ninja flavor            |
| [sn_name_removals.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/SystemNinja/sn_name_removals.ini)               | Contains the set of keys to be removed by their name from the base Winapp2.ini to create the System Ninja flavor     |
| [sn_section_removals.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/SystemNinja/sn_section_removals.ini)         | Contains the set of sections to be removed by their name from the base Winapp2.ini to create the System Ninja flavor | 
| [sn_section_replacements.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/SystemNinja/sn_section_replacements.ini) | Contains the set of section replacements made to the base Winapp2.ini to create the System Ninja flavor              | 
| [sn_value_removals.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/SystemNinja/sn_value_removals.ini)             | Contains the set of key values to be removed from the base Winapp2.ini to create the System Ninja flavor             | 