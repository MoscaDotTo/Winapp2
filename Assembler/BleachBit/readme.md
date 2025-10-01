# BleachBit Flavor Transmutation Rules 

### What is this? 

This folder contains the set of Transmutation rules required by Winapp2ool to produce the [BleachBit flavor](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Non-CCleaner/BleachBit/Winapp2.ini) of Wianpp2.ini 

### What changes are made to create the BleachBit flavor? 

BleachBit does not support `ExcludeKey` keys which target the registry. The BleachBit flavor is created by removing all `ExcludeKeys` from the base Winapp2.ini which contain the `REG` flag. There are no other changes. 

# Files 
| File                                                                                                                                                  | Description                                                                                                       |
| :-                                                                                                                                                    | :-                                                                                                                |
| [bb_additions.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/BleachBit/bb_additions.ini)                       | Contains the set of entries and keys which are added to the base Winapp2.ini to create the BleachBit flavor       |
| [bb_key_replacements.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/BleachBit/bb_key_replacements.ini)         | Contains the set of key replacement values made to the base Winapp2.ini to create the BleachBit flavor            |
| [bb_name_removals.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/BleachBit/bb_name_removals.ini)               | Contains the set of keys to be removed by their name from the base Winapp2.ini to create the BleachBit flavor     |
| [bb_section_removals.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/BleachBit/bb_section_removals.ini)         | Contains the set of sections to be removed by their name from the base Winapp2.ini to create the BleachBit flavor | 
| [bb_section_replacements.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/BleachBit/bb_section_replacements.ini) | Contains the set of section replacements made to the base Winapp2.ini to create the BleachBit flavor              | 
| [bb_value_removals.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/BleachBit/bb_value_removals.ini)             | Contains the set of key values to be removed from the base Winapp2.ini to create the BleachBit flavor             | 