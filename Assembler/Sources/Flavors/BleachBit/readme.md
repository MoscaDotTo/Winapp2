# BleachBit Flavor Transmutation Rules

### What is this?

This folder holds the Transmutation rules winapp2ool uses to make the [BleachBit flavor](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2-bleachbit.ini) of winapp2.ini.

They produce `winapp2-bleachbit.ini`, published as [`Releases/winapp2-bleachbit.ini`](https://github.com/MoscaDotTo/Winapp2/blob/master/Releases/winapp2-bleachbit.ini).

The BleachBit flavor is intended only for BleachBit versions older than 6. BleachBit 6+ should use the base file.

### What changes are made to create the BleachBit flavor?

BleachBit 5 and older didn't support `ExcludeKey`s that target the registry, so the BleachBit flavor removes every `ExcludeKey` with the `REG` flag from the base winapp2.ini. Nothing else changes.

`bb_value_removals.ini` names each of those keys, so a new `REG` ExcludeKey in the base needs its own line there. Running WinappDebug with `-bleachbit` on the built flavor reports any that got through.

# Files

| File | Description |
| :- | :- |
| [bb_additions.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/Sources/Flavors/BleachBit/bb_additions.ini) | Entries and keys added to the base winapp2.ini |
| [bb_key_replacements.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/Sources/Flavors/BleachBit/bb_key_replacements.ini) | Keys whose values are replaced |
| [bb_name_removals.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/Sources/Flavors/BleachBit/bb_name_removals.ini) | Keys removed by name |
| [bb_section_removals.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/Sources/Flavors/BleachBit/bb_section_removals.ini) | Entries removed entirely |
| [bb_section_replacements.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/Sources/Flavors/BleachBit/bb_section_replacements.ini) | Entries replaced entirely |
| [bb_value_removals.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/Sources/Flavors/BleachBit/bb_value_removals.ini) | Keys removed by value |
