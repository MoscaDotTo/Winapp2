# Winapp2.ini

### What is Winapp2.ini?

**Winapp2.ini** is a community maintained database of declarative cleaning routines for Windows. It maps individual applications and system components to their transient data such as temporary files, caches, logs, recently used lists, and more, offering an  inspectable system that can be audited and safely customized to suit individual needs. It functions as an extension of the applications with which it is compatible, enabling it to update independently of them and granting users greater freedom to choose and modify their tools.

Winapp2.ini works with FluentCleaner, BleachBit, CCleaner (*including* CCleaner 7), System Ninja, Avira System Speedup, R-Wipe&Clean, HDCleaner, and Reg Organizer. It comes in additional variants, called *flavors*, each tailored to a particular tool's features or limitations. 

### Will this help make my computer faster?
**Probably not.** On modern systems there's little speed to gain from this kind of cleaning. Over-cleaning caches can even *slow* things down, by forcing apps to rebuild data they could have reused.

That said, there are still plenty of good reasons to clean:

* Troubleshooting app issues
* Reclaiming disk space
* Minimizing the size of system backups
* Privacy
* Or simply because tidying up feels good sometimes

### What are flavors?

A flavor is winapp2.ini with a special set of changes applied to improve compatability with a specifc tool, and to replace the version shipped with that tool (if applicable). Flavors are built automatically alongside the base winapp2.ini file, so they're up to date with the latest winapp2.ini even if the copy shipped with that application isn't.

##

### Disclaimer
Winapp2.ini is provided as-is and without warranty. Understand that its intent is to enable you to delete files, folders, and registry keys off of your system in a way that is programmatic and potentially irreversible. Please exercise caution and take appropriate backups where relevant while using winapp2.ini. It is advised you use winapp2ool to manage your local copy of winapp2.ini, as it can provide bespoke changelogs which should be read carefully to fully understand the scope of changes made between versions.

---

# Table of Contents

1. [Quick Start](#quick-start)
2. [Files of Interest](#files-of-interest)
3. [Installation & Configuration](INSTALL.md)
4. [Contributing](#contributing)
5. [Custom Content](#custom-content)
6. [How Winapp2.ini is built](#how-winapp2ini-is-built)

---

# [Quick Start](#quick-start)
1. Download [winapp2ool.exe](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2ool.exe)
    - Open the winapp2ool settings and pick your tool's flavor. The default is the base flavor. 
2. Follow the [installation guide](INSTALL.md) for your cleaner application.
3. Use winapp2ool to keep your copy up to date and trimmed.

---

# [Files of interest](#files-of-interest)

The winapp2ool and flavor links below download from the [latest release](https://github.com/MoscaDotTo/Winapp2/releases/latest), where most file names carry their flavor, such as `winapp2-ccleaner.ini`. Rename a flavor to `winapp2.ini` (`winapp2.rules` for System Ninja) if you install it by hand.

| Name           		                                                                                                           | Purpose       
| :-                                                                                                                               | :-
| [Winapp2ool](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2ool.exe)                                      | The application that builds winapp2.ini, and the recommended way to install, update, and trim your local copy. It has [its own readme](winapp2ool/Readme.md).
| [Winapp2.ini](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2.ini)                                        | The base winapp2.ini, with nothing removed or changed. [Latest changelog](https://github.com/MoscaDotTo/Winapp2/blob/master/Releases/changelog-base.txt).
| [CCleaner winapp2.ini](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2-ccleaner.ini)                      | The CCleaner flavor, which fits better into its UI. [Latest changelog](https://github.com/MoscaDotTo/Winapp2/blob/master/Releases/changelog-ccleaner.txt).
| [CCleaner 7 winapp2.ini](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2-ccleaner7.ini)                   | The CCleaner 7 flavor, converted to the entry format CCleaner 7 requires. You don't install this file by hand: winapp2ool downloads it and patches it into `ccleaner.ini` for you (see [CCleaner 7](INSTALL.md#ccleaner-7)). [Latest changelog](https://github.com/MoscaDotTo/Winapp2/blob/master/Releases/changelog-ccleaner7.txt).
| [BleachBit winapp2.ini](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2-bleachbit.ini)                    | The BleachBit flavor, **for BleachBit 5 and older only.** It removes the registry exclusions those versions' sanity checker rejects. BleachBit 6+ users should use the base winapp2.ini. [Latest changelog](https://github.com/MoscaDotTo/Winapp2/blob/master/Releases/changelog-bleachbit.txt).
| [System Ninja winapp2.rules](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2-systemninja.rules)           | The System Ninja flavor, which replaces keys System Ninja can't run with ones it can. [Latest changelog](https://github.com/MoscaDotTo/Winapp2/blob/master/Releases/changelog-systemninja.txt).
| [Tron winapp2.ini](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2-tron.ini)                              | The Tron flavor, which carries the changes Tron makes to its own winapp2.ini. [Latest changelog](https://github.com/MoscaDotTo/Winapp2/blob/master/Releases/changelog-tron.txt).
| [FluentCleaner winapp2.ini](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2-fluentcleaner.ini)            | The FluentCleaner flavor,  which carries the changes FluentCleaner makes to its own winapp2.ini. [Latest changelog](https://github.com/MoscaDotTo/Winapp2/blob/master/Releases/changelog-fluentcleaner.txt).
| [Winapp3.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/master/Winapp3/Winapp3.ini)                                   | An extension for an extension, with entries for power users. *Do **not** use this file if you don't know what you're doing. Some of its entries are very aggressive and can do real damage to your file system.*
| [Assembler](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler)                                                         | The sources from which every published file above is built, and the build script that drives them. Contributions go here. See [How Winapp2.ini is built](#how-winapp2ini-is-built).

---

# [Installation & Configuration](INSTALL.md)

Select your cleaner below to learn how to install and configure winapp2.ini for it.

It is strongly recommended that you keep a copy of [winapp2ool.exe](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2ool.exe) in the same folder as winapp2.ini to help keep it up to date. 

##
### [FluentCleaner](INSTALL.md#fluentcleaner) 
##
### [BleachBit](INSTALL.md#bleachbit) 
##
### [CCleaner Classic](INSTALL.md#ccleaner-classic) 
##
### [CCleaner 7](INSTALL.md#ccleaner-7) 
##
### [Tron](INSTALL.md#tron) 
##
### [System Ninja](INSTALL.md#system-ninja) 
##
### [Reg Organizer](INSTALL.md#reg-organizer)
##
### [Avira System Speedup](INSTALL.md#avira-system-speedup) 
##
### [HDCleaner](INSTALL.md#hdcleaner)
##
### [R-Wipe & Clean](INSTALL.md#r-wipe--clean)

---

# [Contributing](#contributing)

To add, update, or remove an entry, see [our contributor guidelines](CONTRIBUTING.md).

---

# [Custom content](#custom-content)

Winapp2.ini doesn't currently cover non-English system configurations or portable software out of the box. If you need them, keep your own entries in a `custom.ini` file and use winapp2ool's [Transmute](winapp2ool/modules/transmute/readme.md) in `Add` mode to apply them to each new winapp2.ini. For a portable browser, [BrowserBuilder](winapp2ool/modules/browserbuilder/readme.md) can generate the entries from your own install's path.

Winapp2ool 1.6 replaced the Merge feature with Transmute. If you used `custom.ini` with Merge, see [Migrating from Merge](winapp2ool/modules/transmute/readme.md#migrating-from-merge) in the Transmute readme.

--- 

# [How Winapp2.ini is built](#how-winapp2ini-is-built)

Winapp2.ini is generated from the sources in [Assembler](Assembler) by [winapp2ool](winapp2ool/Readme.md). The build publishes them to [Releases](Releases), each under the name it carries on the GitHub release. 

### Build pipeline

Each build runs four stages, driven by [`build winapp2.ps1`](Assembler/build%20winapp2.ps1):

| Stage | What runs | What it does |
| :- | :- | :- |
| Generate | [EntryBuilder](winapp2ool/modules/entrybuilder/readme.md) | Expands the per-letter [base entry sources](Assembler/Sources/EntryBuilder) into artifacts. Sources can use an optional shorthand DSL and shared scaffold catalogs for webviews and [Microsoft Store apps](winapp2ool/modules/entrybuilder/readme.md#store-and-hybrid-apps) |
| | [BrowserBuilder](winapp2ool/modules/browserbuilder/readme.md) | Produces one entry per browser per cleaning category, crossing [browser declarations](Assembler/Sources/BrowserBuilder) with the templates that apply to their engine |
| Merge | [Combine](winapp2ool/modules/combine/readme.md) | Joins the generated artifacts into one file in strict mode, which fails the build if two artifacts use the same entry name |
| Lint | [WinappDebug](winapp2ool/modules/winappdebug/README.md) | Checks the merged file's syntax and style, and saves its own corrections |
| Publish | [Flavorizer](winapp2ool/modules/transmute/Flavorizer/readme.md), [WinappDebug](winapp2ool/modules/winappdebug/README.md), and [Diff](winapp2ool/modules/diff/readme.md) | Applies each flavor's ruleset to make the six other variants and lints each one, then writes a changelog against the previously published version for every output whose content changed, and publishes both to [Releases](Releases) |

### CI/CD

The build artifacts under [Assembler/Generated](Assembler/Generated) are committed to the repository instead of living only for the length of a build. That lets us check the build against itself, and gives CI something to hold still while it varies one input at a time:

| Check | What it does |
| :- | :- |
| **Generated artifact guard** | Fails a change that edits build output without touching the source that produces it, or that puts an entry in the wrong letter's source file, and names the file you should have edited |
| **PR verify** | Builds winapp2.ini from `master`, then again with the pull request merged, and comments either the changelog the pull request would produce or the stage where it broke the build |
| **Artifact drift** | Regenerates every artifact into a scratch folder and compares it byte for byte with what's committed, which catches generator output changing when it shouldn't |
| **Binary parity** | Checks that the committed winapp2ool.exe still matches the source beside it. Otherwise every published file would come from code that isn't in the repository |
| **Update signature** | Checks that copies of winapp2ool already out there will accept the committed update: its signature has to verify, and `version.txt` has to match the exe |
| **Maintainer paths guard** | Fails a pull request from outside the project that touches winapp2ool's source, its committed exe, the build scripts, or the workflows |
| **Daily build and release** | Rebuilds from current sources. When the content changed, it commits the result and cuts a tagged release with every flavor and changelog attached |