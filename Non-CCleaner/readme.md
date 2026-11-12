# Winapp2.ini Non-CCleaner Flavors

> [!WARNING]
> **This folder is deprecated and will be removed in a future release.** Every flavor of winapp2.ini is now published in the [`Releases`](https://github.com/MoscaDotTo/Winapp2/tree/master/Releases) folder, under the file name it has on the [releases page](https://github.com/MoscaDotTo/Winapp2/releases/latest). The files here are copies, kept up to date for a while so scripts that download them have time to switch. Download from `Releases`, or from `https://github.com/MoscaDotTo/Winapp2/releases/latest/download/<file>`, instead.
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
> Each `diff.txt` here is the same flavor's `changelog-<flavor>.txt` in `Releases`.

### What is this?

This folder holds the flavors of winapp2.ini other than the CCleaner one, including the base winapp2.ini every flavor is built from. The base file used to be the only non-CCleaner version.

winapp2ool builds these files as part of the winapp2.ini build.

### Contributions

Don't open changes against the winapp2.ini files in these folders. Change the files winapp2.ini is built from instead. The base entries, which are most of winapp2.ini, are in [`Assembler/Sources/EntryBuilder`](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/Sources/EntryBuilder). Browser entries are generated from [`Assembler/Sources/BrowserBuilder`](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/Sources/BrowserBuilder), and UWP entries from [`Assembler/Sources/UWPBuilder`](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/Sources/UWPBuilder).

If your change is for one flavor only, edit that flavor's Transmutation rules:
* [BleachBit](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/Sources/Flavors/BleachBit)
* [CCleaner](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/Sources/Flavors/CCleaner)
* [CCleaner 7](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/Sources/Flavors/CCleaner7)
* [FluentCleaner](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/Sources/Flavors/FluentCleaner)
* [System Ninja](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/Sources/Flavors/SystemNinja)
* [Tron](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/Sources/Flavors/Tron)

# Files of interest

| File | Description |
| :- | :- |
| [Base winapp2.ini](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2.ini) | The base winapp2.ini. Nothing is removed or changed, so it includes rules that may overlap or conflict with CCleaner's or BleachBit's own |
| [BleachBit Winapp2.ini](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2-bleachbit.ini) | The BleachBit flavor, without the rules BleachBit doesn't support, so it passes BleachBit's sanity checker |
| [CCleaner 7 Winapp2.ini](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2-ccleaner7.ini) | The CCleaner 7 flavor, which winapp2ool's CC7Patcher adds to CCleaner 7's `ccleaner.ini` |
| [FluentCleaner Winapp2.ini](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2-fluentcleaner.ini) | The FluentCleaner flavor, carrying the changes FluentCleaner made to its own copy |
| [System Ninja winapp2.rules](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2-systemninja.rules) | The System Ninja flavor, with some rules System Ninja can't read replaced by ones it can |
| [Tron winapp2.ini](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2-tron.ini) | The Tron flavor, carrying the changes Tron made to its own copy |
