# Assembler 

### What is this?

This folder contains the build script for Winapp2.ini and everything required by Winapp2ool to build Winapp2.ini and all of its flavors. 

### Building Winapp2.ini 

To build Winapp2.ini, simply place a copy of Winapp2ool 1.8 or newer in this folder and run `& '.\build winapp2.ps1'` from PowerShell

The script regenerates the committed build artifacts in [Generated](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/Generated): the base entries from the [EntryBuilder](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/Sources/EntryBuilder) sources plus the BrowserBuilder and UWPBuilder generator outputs (the generators read the shared cleaning-pattern catalogs from [Scaffolds](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/Sources/Scaffolds)),  merges them all into the base Winapp2.ini with a single strict Combine pass, and then produces each flavor from it. It publishes each flavor and its changelog to [Releases](https://github.com/MoscaDotTo/Winapp2/tree/master/Releases) at the repository root, named as they are on each GitHub release. Nothing in this folder is published

Running `& '.\build winapp2.ps1' -Verify` performs no build: it regenerates every artifact into a folder and byte-compares against the committed files under `Generated\`, failing the build if they don't match.

### Contributions

Contributions to Winapp2.ini or its flavors should be made to the files in this folder and its children, not to individual Winapp2.ini files. See the [contributor guidelines](https://github.com/MoscaDotTo/Winapp2/blob/master/CONTRIBUTING.md) for how to write and submit changes

# Files and Folders

| Name | Description |
| :- | :- |
| [`build winapp2.ps1`](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/build%20winapp2.ps1) | The build script for Winapp2.ini. Creates the base winapp2.ini and each of the flavors |
| [`Sources\`](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/Sources) | Everything that is written by hand. This is the only place to edit |
| [`Generated\`](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/Generated) | Everything the build writes for itself. Never edit these files |

## Sources\

| Folder | Contents |
| :- | :- |
| [EntryBuilder](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/Sources/EntryBuilder) | The source files for all base entries, comprising the bulk of Winapp2.ini. EntryBuilder's shorthand is optional here |
| [BrowserBuilder](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/Sources/BrowserBuilder) | The generative rules for Chromium and Gecko based browsers, as well as the Transmutation rules to correct them |
| [UWPBuilder](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/Sources/UWPBuilder) | The AppInfo definitions and scaffold UWPBuilder uses to generate entries for Microsoft Store apps |
| [Scaffolds](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/Sources/Scaffolds) | The shared cleaning-pattern catalogs (WebView2, QtWebEngine, Electron) used by EntryBuilder and UWPBuilder |
| [Flavors\BleachBit](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/Sources/Flavors/BleachBit) | The Transmutation rules for the BleachBit flavor |
| [Flavors\CCleaner](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/Sources/Flavors/CCleaner) | The Transmutation rules for the CCleaner flavor |
| [Flavors\CCleaner7](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/Sources/Flavors/CCleaner7) | The Transmutation rules which convert the CCleaner flavor into the CCleaner 7 format |
| [Flavors\FluentCleaner](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/Sources/Flavors/FluentCleaner) | The Transmutation rules for the FluentCleaner flavor |
| [Flavors\SystemNinja](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/Sources/Flavors/SystemNinja) | The Transmutation rules for the System Ninja flavor |
| [Flavors\Tron](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/Sources/Flavors/Tron) | The Transmutation rules for the Tron flavor |

## Generated\

| Folder | Contents |
| :- | :- |
| [EntryBuilder](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/Generated/EntryBuilder) | `#.ini`, `A.ini` ... `Z.ini`, written by EntryBuilder from `Sources\EntryBuilder` |
| BrowserBuilder | `browsers.ini`, written by BrowserBuilder from `Sources\BrowserBuilder` |
| [UWPBuilder](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/Generated/UWPBuilder) | `#.ini`, `A.ini` ... `Z.ini`, written by UWPBuilder from `Sources\UWPBuilder` |

These are the pieces the build's strict Combine joins into the base winapp2.ini. They are committed so every change to winapp2.ini can be reviewed at the entry level.

## Which rules produce which published file

The mapping is not one to one. Tron, CCleaner 7 and FluentCleaner are built on top of the CCleaner rules, so two rule sets reach each of them. Each file is published to `Releases\` and attached to each GitHub release under the name below.

| Rule set | Published as |
| :- | :- |
| none (no flavor pass) | `winapp2.ini` (base) |
| `Flavors\CCleaner` | `winapp2-ccleaner.ini` |
| `Flavors\FluentCleaner`, then `Flavors\CCleaner` | `winapp2-fluentcleaner.ini` |
| `Flavors\CCleaner`, then `Flavors\CCleaner7` | `winapp2-ccleaner7.ini` |
| `Flavors\BleachBit` | `winapp2-bleachbit.ini` |
| `Flavors\CCleaner`, then `Flavors\Tron` | `winapp2-tron.ini` |
| `Flavors\SystemNinja` | `winapp2-systemninja.rules` |
