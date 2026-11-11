# Winapp2.ini Non-CCleaner Flavors 

> [!WARNING]
> **This folder is deprecated and will be removed in a future release.** Every flavor of winapp2.ini is now published in the [`Releases`](https://github.com/MoscaDotTo/Winapp2/tree/master/Releases) folder, under the same file names it has on the [releases page](https://github.com/MoscaDotTo/Winapp2/releases/latest). The files here are copies, kept up to date for a limited time so scripts that download them have time to switch. Please download from `Releases` or from `https://github.com/MoscaDotTo/Winapp2/releases/latest/download/<file>` instead.
>
> | Old path | New file |
> | :- | :- |
> | `Winapp2.ini` (repository root) | [`winapp2-ccleaner.ini`](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2-ccleaner.ini) |
> | `Non-CCleaner/Winapp2.ini` | [`winapp2.ini`](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2.ini) |
> | `Non-CCleaner/BleachBit/Winapp2.ini` | [`winapp2-bleachbit.ini`](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2-bleachbit.ini) |
> | `Non-CCleaner/CCleaner7/Winapp2.ini` | [`winapp2-ccleaner7.ini`](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2-ccleaner7.ini) |
> | `Non-CCleaner/FluentCleaner/Winapp2.ini` | [`winapp2-fluentcleaner.ini`](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2-fluentcleaner.ini) |
> | `Non-CCleaner/Tron/Winapp2.ini` | [`winapp2-tron.ini`](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2-tron.ini) |
> | `Non-CCleaner/SystemNinja/Winapp2.rules` | [`winapp2-systemninja.rules`](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2-systemninja.rules) |
>
> Each `diff.txt` here is the `changelog-<flavor>.txt` of the same flavor in `Releases`.

### What is this?

This folder contains the set of Non-CCleaner flavors of Winapp2.ini, including the "base" Winapp2.ini, formerly the singular Non-CCleaner version, from which all flavors are derived. 

These files are assembled by Winapp2ool as part of the build process for Winapp2.ini. 

### Contributions 

You should not open up contributions directly against the Winapp2.ini files found in these folders. Instead, you should modify the files used to build Winapp2.ini. The base entries which make up the bulk of winapp2.ini can be found [here](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/Sources/EntryBuilder). The entries targeting web browsers can be found [here](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/Sources/BrowserBuilder). 

If you are contributing a change specific to a particular flavor, please make changes against the Transmutation rules in the [folder associated](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler) with that flavor.

Links to transmutation rules per-flavor:
* [BleachBit](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/Sources/Flavors/BleachBit)
* [CCleaner](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/Sources/Flavors/CCleaner)
* [System Ninja](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/Sources/Flavors/SystemNinja)
* [Tron](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/Sources/Flavors/Tron)

# Files of interest

| File                | Description |
| :-                                                                                                                               | :-                                                                                                                                                       |
| [Base winapp2.ini](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2.ini)                                   | This is the base winapp2.ini file, it has no content removed or changed, and includes rules which may overlap or conflict with CCleaner/BleachBit rules. |
| [BleachBit Winapp2.ini](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2-bleachbit.ini)                    | The BleachBit flavor of winapp2.ini, designed to remove unsupported rules and pass the sanity checker.                                                   |
| [System Ninja winapp2.rules](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2-systemninja.rules)           | The System Ninja flavor of winapp2.ini, designed to replace unsupported rules with ones compatible with System Ninja.                                    |
| [Tron winapp2.ini](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2-tron.ini)                              | The Tron flavor of winapp2.ini, designed to capture the downstream changes made by Tron to the base flavor.                                          |
