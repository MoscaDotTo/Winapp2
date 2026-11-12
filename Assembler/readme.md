# Assembler

### What is this?

This folder holds everything winapp2ool needs to build winapp2.ini and its flavors: the hand-written sources, the build script, and the generated pieces the build puts together.

### Building Winapp2.ini

Put a copy of winapp2ool in this folder and run `& '.\build winapp2.ps1'` from PowerShell. A clone of the repository already has one. The script refuses a winapp2ool older than the version the sources need, and tells you which version that is.

The script regenerates the committed build artifacts in [Generated](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/Generated): the base entries from the [EntryBuilder](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/Sources/EntryBuilder) sources, plus BrowserBuilder's and UWPBuilder's output (EntryBuilder and UWPBuilder also read the shared cleaning-pattern catalogs in [Scaffolds](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/Sources/Scaffolds)). One strict Combine pass joins them into the base winapp2.ini, and each flavor is made from that. Every flavor and its changelog goes to [Releases](https://github.com/MoscaDotTo/Winapp2/tree/master/Releases) at the repository root, under the name it has on each GitHub release. Nothing in this folder is published.

`& '.\build winapp2.ps1' -Verify` builds nothing. It regenerates every artifact into a scratch folder, byte-compares the result against the committed files under `Generated\`, and fails if anything differs.

### Contributions

Changes to winapp2.ini or its flavors go in the files under `Sources\`, never in a built winapp2.ini. The [contributor guidelines](https://github.com/MoscaDotTo/Winapp2/blob/master/CONTRIBUTING.md) explain how to write and submit them.

# Files and Folders

| Name | Description |
| :- | :- |
| [`build winapp2.ps1`](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/build%20winapp2.ps1) | The build script. Builds the base winapp2.ini and every flavor |
| [`Sources\`](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/Sources) | Everything written by hand. This is the only place to make changes |
| [`Generated\`](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/Generated) | Everything the build writes for itself. Never edit these files |

## Sources\

| Folder | Contents |
| :- | :- |
| [EntryBuilder](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/Sources/EntryBuilder) | The sources for every base entry, which is most of winapp2.ini. EntryBuilder's shorthand is optional here |
| [BrowserBuilder](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/Sources/BrowserBuilder) | The rules that generate entries for Chromium and Gecko browsers, and the Transmutation rules that correct them |
| [UWPBuilder](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/Sources/UWPBuilder) | The scaffold and AppInfo definitions UWPBuilder turns into entries for Microsoft Store apps |
| [Scaffolds](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/Sources/Scaffolds) | The shared cleaning-pattern catalogs (WebView2, QtWebEngine, Electron) EntryBuilder and UWPBuilder draw on |
| [Flavors\BleachBit](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/Sources/Flavors/BleachBit) | The Transmutation rules for the BleachBit flavor |
| [Flavors\CCleaner](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/Sources/Flavors/CCleaner) | The Transmutation rules for the CCleaner flavor |
| [Flavors\CCleaner7](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/Sources/Flavors/CCleaner7) | The Transmutation rules that convert the CCleaner flavor to CCleaner 7's format |
| [Flavors\FluentCleaner](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/Sources/Flavors/FluentCleaner) | The Transmutation rules for the FluentCleaner flavor |
| [Flavors\SystemNinja](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/Sources/Flavors/SystemNinja) | The Transmutation rules for the System Ninja flavor |
| [Flavors\Tron](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/Sources/Flavors/Tron) | The Transmutation rules for the Tron flavor |

## Generated\

| Folder | Contents |
| :- | :- |
| [EntryBuilder](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/Generated/EntryBuilder) | `#.ini`, `A.ini` ... `Z.ini`, written by EntryBuilder from `Sources\EntryBuilder` |
| BrowserBuilder | `browsers.ini`, written by BrowserBuilder from `Sources\BrowserBuilder` |
| [UWPBuilder](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/Generated/UWPBuilder) | `#.ini`, `A.ini` ... `Z.ini`, written by UWPBuilder from `Sources\UWPBuilder` |

These are the pieces the build's strict Combine joins into the base winapp2.ini. We commit them so every change to winapp2.ini can be reviewed entry by entry.

## Which rules produce which published file

It isn't one to one. CCleaner 7 and FluentCleaner are both built with the CCleaner rules too, so two rule sets reach each of them. Every file goes to `Releases\`, and onto each GitHub release, under the name below.

| Rule set | Published as |
| :- | :- |
| none (no flavor pass) | `winapp2.ini` (base) |
| `Flavors\CCleaner` | `winapp2-ccleaner.ini` |
| `Flavors\FluentCleaner`, then `Flavors\CCleaner` | `winapp2-fluentcleaner.ini` |
| `Flavors\CCleaner`, then `Flavors\CCleaner7` | `winapp2-ccleaner7.ini` |
| `Flavors\BleachBit` | `winapp2-bleachbit.ini` |
| `Flavors\Tron` | `winapp2-tron.ini` |
| `Flavors\SystemNinja` | `winapp2-systemninja.rules` |
