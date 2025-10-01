# Winapp2.ini Non-CCleaner Flavors 

### What is this?

This folder contains the set of Non-CCleaner flavors of Winapp2.ini, including the "base" Winapp2.ini, formerly the singular Non-CCleaner version, from which all flavors are derived. 

These files are assembled by Winapp2ool as part of the build process for Winapp2.ini. 

### Contributions 

You should not open up contributions directly against the Winapp2.ini files found in these folders. Instead, you should modify the files used to build Winapp2.ini. The base entries which make up the bulk of winapp2.ini can be found [here](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/Entries). The entries targeting web browsers can be found [here](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/BrowserBuilder). 

If you are contributing a change specific to a particular flavor, please make changes against the Transmutation rules in the [folder associated](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler) with that flavor.

Links to transmutation rules per-flavor:
* [BleachBit](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/BleachBit)
* [CCleaner](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/CCleaner)
* [System Ninja](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/SystemNinja)
* [Tron](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/Tron)

# Files of interest

| File                | Description |
| :-                                                                                                                               | :-                                                                                                                                                       |
| [Base winapp2.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/master/Non-CCleaner/Winapp2.ini)                         | This is the base winapp2.ini file, it has no content removed or changed, and includes rules which may overlap or conflict with CCleaner/BleachBit rules. |
| [BleachBit Winapp2.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/master/Non-CCleaner/BleachBit/Winapp2.ini)          | The BleachBit flavor of winapp2.ini, designed to remove unsupported rules and pass the sanity checker.                                                   |
| [System Ninja winapp2.rules](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/master/Non-CCleaner/SystemNinja/Winapp2.rules) | The System Ninja flavor of winapp2.ini, designed to replace unsupported rules with ones compatible with System Ninja.                                    |
| [Tron winapp2.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/master/Non-CCleaner/Tron/Winapp2.ini)                    | The Tron flavor of winapp2.ini, designed to capture the downstream changes made by Tron to the CCleaner flavor.                                          |
