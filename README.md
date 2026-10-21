# Winapp2.ini

### What is Winapp2.ini?

**Winapp2.ini** is a community-driven database of declarative cleaning routines for Microsoft Windows. It provides a comprehensive mapping of individual applications and system components to their transient data (temporary files, caches, logs, recently used lists, and more). Beginning in [2010](https://community.ccleaner.com/t/winapp2-ini-additions/40035) and hosted on GitHub since 2016, it is bundled or fetched on demand by most of the cleaning tools that support it, several of which serve it to users who never encounter this repository. winapp2.ini is rebuilt automatically every day, and a new version is published when its content has changed.

Winapp2.ini is compatible with CCleaner (*including* CCleaner 7), BleachBit, System Ninja, Avira System Speedup, R-Wipe&Clean, HDCleaner, FluentCleaner, and Reg Organizer. It is published in seven variants (called *flavors*), each tailored to the feature set of the tool that consumes it, and each with its own changelog.

### Why Winapp2.ini?

Winapp2.ini avoids the risks of overreach common in generic cleaning tools by adopting an exhaustive, declarative approach. Where many tools rely on sweeping file-type patterns applied across entire drives, Winapp2.ini demands explicitly defined target paths and conceptual linkage between those targets and their parent applications. This prioritizes clarity, specificity, and control over generalization, offering users an inspectable system that can be audited and safely customized to suit individual needs.

Winapp2.ini functions as an extension of the applications with which it is compatible, enabling it to update independently of them. This decoupling grants users greater freedom to move between tools and versions without sacrificing functionality.

### Will this help make my computer faster?
**Probably not.** On modern systems, there's little performance incentive for this kind of system hygiene. In fact, over-cleaning caches can potentially *reduce* your performance by forcing apps to rebuild data they could have reused.

That said, there are still plenty of good reasons to clean:

* Troubleshooting app issues
* Reclaiming disk space
* Minimizing the size of system backups
* Enhancing privacy
* Or simply because tidying up feels good sometimes

### What are flavors?

Flavors are the result of specific sets of modifications applied to each Winapp2.ini update to produce variants which cater more closely to the features supported by particular applications. This is an automated process carried out when Winapp2.ini is built for each update, so these flavors are always up to date with the latest version of Winapp2.ini even if the copy shipped with the application is not. Flavors are intended to function as drop-in replacements to the Winapp2.ini shipped with each of these applications.

### Disclaimer
Winapp2.ini is provided as-is and without warranty. Understand that its intent is to enable you to delete files, folders, and registry keys off of your system in a way that is programmatic and potentially irreversible. Please exercise caution and take appropriate backups where relevant while using winapp2.ini. It is advised you use winapp2ool to manage your local copy of winapp2.ini, as it can provide bespoke changelogs which should be read carefully to fully understand the scope of changes made between versions.

---

# Table of Contents

1. [Quick Start](#quick-start)
2. [Files of Interest](#files-of-interest)
3. [Installation & Configuration](INSTALL.md)
4. [How Winapp2.ini is built](#how-winapp2ini-is-built)
5. [Contributing](#contributing)
6. [Custom Content](#custom-content)

---

# [Quick Start](#quick-start)
1. Download [winapp2ool.exe](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2ool.exe)
    - If necessary, open the winapp2ool settings and select your preferred flavor. The default flavor is the CCleaner flavor.
2. Follow the [installation guide](INSTALL.md) for your cleaner application.
3. Use winapp2ool to keep your copy updated and trimmed for optimal performance.

---

# [Files of interest](#files-of-interest)

The winapp2ool and flavor links below download from the [latest release](https://github.com/MoscaDotTo/Winapp2/releases/latest), where most file names carry their flavor, such as `winapp2-ccleaner.ini`. Rename a flavor to `winapp2.ini` (`winapp2.rules` for System Ninja) if you install it by hand.

| Name           		                                                                                                           | Purpose       
| :-                                                                                                                               | :-
| [Winapp2ool](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2ool.exe)                                      | The application that builds winapp2.ini, and the recommended way to install, update, and trim your local copy. This tool has its own ReadMe [here](https://github.com/MoscaDotTo/Winapp2/tree/master/winapp2ool).
| [Winapp2.ini](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2.ini)                                        | This is the base winapp2.ini file, it has no content removed or changed. View the latest change log for the base file [here](https://github.com/MoscaDotTo/Winapp2/blob/master/Non-CCleaner/diff.txt).
| [CCleaner Winapp2.ini](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2-ccleaner.ini)                      | The CCleaner flavor of winapp2.ini, designed to reduce overlap with CCleaner rules and better integrate with its UI. View the latest change log for this flavor [here](https://github.com/MoscaDotTo/Winapp2/blob/master/diff.txt).
| [CCleaner 7 Winapp2.ini](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2-ccleaner7.ini)                   | The CCleaner 7 flavor of winapp2.ini, converted to the entry format CCleaner 7 requires. This file is not installed by hand; winapp2ool downloads it for you and patches it into `ccleaner.ini`. See [CCleaner 7](INSTALL.md#ccleaner-7). View the latest change log for this flavor [here](https://github.com/MoscaDotTo/Winapp2/blob/master/Non-CCleaner/CCleaner7/diff.txt).
| [BleachBit Winapp2.ini](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2-bleachbit.ini)                    | The BleachBit flavor of winapp2.ini, **for BleachBit 5 and older only.** It removes the registry exclusions those versions do not support, which their sanity checker rejects. BleachBit 6 and newer support them, so those users should use the base winapp2.ini above. View the latest change log for this flavor [here](https://github.com/MoscaDotTo/Winapp2/blob/master/Non-CCleaner/BleachBit/diff.txt).
| [System Ninja winapp2.rules](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2-systemninja.rules)           | The System Ninja flavor of winapp2.ini, designed to replace unsupported rules with ones compatible with System Ninja. View the latest change log for this flavor [here](https://github.com/MoscaDotTo/Winapp2/blob/master/Non-CCleaner/SystemNinja/diff.txt).
| [Tron winapp2.ini](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2-tron.ini)                              | The Tron flavor of winapp2.ini, designed to capture the downstream changes made by Tron to the CCleaner flavor. View the latest change log for this flavor [here](https://github.com/MoscaDotTo/Winapp2/blob/master/Non-CCleaner/Tron/diff.txt). 
| [FluentCleaner winapp2.ini](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2-fluentcleaner.ini)            | The FluentCleaner flavor of winapp2.ini, designed to capture the downstream changes made by FluentCleaner to their winapp2.ini. View the latest change log for this flavor [here](https://github.com/MoscaDotTo/Winapp2/blob/master/Non-CCleaner/FluentCleaner/diff.txt).
| [Winapp3.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/master/Winapp3/Winapp3.ini)                                   | An extension for an extension; contains entries for use by power users. *You should **not** use this file if you do not know what you are doing. Entries in this file can potentially be very aggressive/dangerous to your file system.*
| [Assembler](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler)                                                         | The sources from which every published file above is built, and the build script that drives them. Contributions target these files. See [How Winapp2.ini is built](#how-winapp2ini-is-built).

---

# [Installation & Configuration](INSTALL.md)

Click the link for your tool below for information on how to install & configure winapp2.ini for it:

[CCleaner Classic](INSTALL.md#ccleaner-classic) &middot; [CCleaner 7](INSTALL.md#ccleaner-7) &middot; [BleachBit](INSTALL.md#bleachbit) &middot; [System Ninja](INSTALL.md#system-ninja) &middot; [Avira System Speedup](INSTALL.md#avira-system-speedup) &middot; [Tron](INSTALL.md#tron) &middot; [R-Wipe & Clean](INSTALL.md#r-wipe--clean) &middot; [HDCleaner](INSTALL.md#hdcleaner) &middot; [FluentCleaner](INSTALL.md#fluentcleaner) &middot; [Reg Organizer](INSTALL.md#reg-organizer)

It is strongly recommended you keep a copy of [winapp2ool.exe](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2ool.exe) in the same folder as winapp2.ini for the purpose of keeping it up-to-date irrespective of which application you are using.

---

# [How Winapp2.ini is built](#how-winapp2ini-is-built)

No published winapp2.ini file in this repository is edited by hand. Every one of them, and every changelog beside it, is generated from source in [Assembler](Assembler) by [winapp2ool](winapp2ool/Readme.md), a purpose-built console application maintained alongside the database. Its readme covers the tool itself, and each stage below links the readme for the module that runs it.

### Build pipeline

Each build runs four stages, orchestrated by [`build winapp2.ps1`](Assembler/build%20winapp2.ps1)

| Stage | What runs | What it does |
| :- | :- | :- |
| Generate | [EntryBuilder](winapp2ool/modules/entrybuilder/readme.md) | Expands the per-letter [base entry sources](Assembler/EntryBuilder) into artifacts, resolving an optional shorthand DSL and shared scaffold catalogs |
| | [BrowserBuilder](winapp2ool/modules/browserbuilder/readme.md) | Produces one entry per browser per cleaning category, crossing [browser declarations](Assembler/BrowserBuilder) with the templates that apply to their engine |
| | [UWPBuilder](winapp2ool/modules/uwpbuilder/readme.md) | Generates Microsoft Store entries from [package declarations](Assembler/UWP) and a shared scaffold |
| Merge | [Combine](winapp2ool/modules/combine/readme.md) | Joins the generated artifacts into one file using Strict Mode, failing the build if names collide between artifacts |
| Lint | [WinappDebug](winapp2ool/modules/winappdebug/README.md) | Static analysis over the merged file: validates syntax, enforces the style rules, and applies its own corrections |
| Publish | [Flavorizer](winapp2ool/modules/transmute/Flavorizer/readme.md) and [Diff](winapp2ool/modules/diff/readme.md) | Applies each flavor's ruleset to derive the six additional variants, then generates a changelog for every output against the previously published version |

### CI/CD

The build artifacts under [Assembler/Entries](Assembler/Entries) are committed to the repository rather than existing only during a build. That makes the build checkable against itself, and gives CI something to hold fixed while varying one input at a time:

| Check | Question it answers |
| :- | :- |
| **Generated artifact guard** | Did this change edit build output without editing the source that produces it? If so, names the file you should have edited |
| **PR verify** | Builds winapp2.ini twice, from `master` alone and from `master` with the PR merged, then comments the resulting changelog on the PR, or the stage at which the build broke |
| **Artifact drift** | Regenerates every artifact into a scratch directory and byte-compares against what is committed. Catches generator output changing when it was not supposed to |
| **Binary parity** | Catches a committed exe that no longer matches the source code beside it, which would otherwise keep generating every published file from code that isn't in the repo. |
| **Daily build and release** | Rebuilds from current sources, commits the result, and cuts a tagged release with every flavor and changelog attached |

Winapp2.ini is rebuilt automatically every day, and a new version is published only when its content has changed. The version line is restamped on every daily build, so it is excluded from that comparison; otherwise an untouched rebuild would look like a change and publish an empty changelog every day.

---

# [Contributing](#contributing)

Contributions target the source files under [Assembler](Assembler), never the published files the build produces from them. [Our contributor guidelines](CONTRIBUTING.md) cover where each kind of entry lives, the syntax rules for each, and what the automated checks will tell you after you open a pull request. Editing from the GitHub web interface is enough for most contributions; you never need to run the build or generate any output file yourself.

---

# [Custom content](#custom-content)

Winapp2.ini does not support non-English system configurations or portable software natively. If you have need for these features, we recommend you utilize a "Custom.ini" file, and use Winapp2ool's [Transmute](https://github.com/MoscaDotTo/Winapp2/tree/master/winapp2ool/modules/transmute) feature with the Transmute mode set to `Add` to add your custom configurations while keeping winapp2.ini up to date.

Winapp2ool 1.6 removed the Merge feature and replaced it with Transmute. If you were previously using Custom.ini with Merge, please see [Migrating From Merge](https://github.com/MoscaDotTo/Winapp2/tree/master/winapp2ool/modules/transmute#migrating-from-merge) in the Transmute ReadMe.