# Winapp2.ini

### What is Winapp2.ini?

**Winapp2.ini** is a community-maintained database of declarative cleaning routines for Windows. It maps individual applications and system components to their transient data: temporary files, caches, logs, recently used lists, and so on. It started in [2010](https://community.ccleaner.com/t/winapp2-ini-additions/40035) and has lived on GitHub since 2016. Most of the cleaning tools that support it bundle it or fetch it on demand, so plenty of the people using it never see this repository (welcome!). Winapp2.ini is rebuilt every day, and a new version is published whenever its content changes.

Winapp2.ini works with CCleaner (*including* CCleaner 7), BleachBit, System Ninja, Avira System Speedup, R-Wipe&Clean, HDCleaner, FluentCleaner, and Reg Organizer. It comes in seven variants, called *flavors*, each tailored to what its tool supports and each with its own changelog.

### Why Winapp2.ini?

A lot of cleaning tools sweep whole drives for file types they think are junk. Winapp2.ini names its targets instead. Every entry lists the exact paths it cleans and the application they belong to, so you can read what an entry deletes before it runs, and change it if it doesn't suit you.

Winapp2.ini is an extension to the cleaners that read it, so it updates on its own schedule instead of waiting for theirs. That also means you can switch tools, or versions of a tool, without losing your cleaning rules.

### Will this help make my computer faster?
**Probably not.** On modern systems there's little speed to gain from this kind of cleaning. Over-cleaning caches can even *slow* things down, by forcing apps to rebuild data they could have reused.

That said, there are still plenty of good reasons to clean:

* Troubleshooting app issues
* Reclaiming disk space
* Minimizing the size of system backups
* Privacy
* Or simply because tidying up feels good sometimes

### What are flavors?

A flavor is winapp2.ini with one tool's changes applied. The build makes every flavor from each new winapp2.ini, so your flavor is as current as the base file even when the copy your cleaner shipped with isn't. Each one is meant as a drop-in replacement for that copy.

### Disclaimer
Winapp2.ini is provided as-is and without warranty. Its whole purpose is to delete files, folders, and registry keys from your system, automatically and potentially irreversibly, so be careful with it and back up anything that matters. We recommend managing your local copy with winapp2ool, which can show you the changelog between your copy and the latest. Read it before you update.

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
    - If you use something other than CCleaner, open the winapp2ool settings and pick your tool's flavor. CCleaner is the default.
2. Follow the [installation guide](INSTALL.md) for your cleaner application.
3. Use winapp2ool to keep your copy up to date, and trimmed to the entries your machine needs.

---

# [Files of interest](#files-of-interest)

The winapp2ool and flavor links below download from the [latest release](https://github.com/MoscaDotTo/Winapp2/releases/latest), where most file names carry their flavor, such as `winapp2-ccleaner.ini`. Rename a flavor to `winapp2.ini` (`winapp2.rules` for System Ninja) if you install it by hand.

| Name           		                                                                                                           | Purpose       
| :-                                                                                                                               | :-
| [Winapp2ool](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2ool.exe)                                      | The application that builds winapp2.ini, and the recommended way to install, update, and trim your local copy. It has [its own readme](winapp2ool/Readme.md).
| [Winapp2.ini](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2.ini)                                        | The base winapp2.ini, with nothing removed or changed. [Latest changelog](https://github.com/MoscaDotTo/Winapp2/blob/master/Non-CCleaner/diff.txt).
| [CCleaner winapp2.ini](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2-ccleaner.ini)                      | The CCleaner flavor, which fits better into its UI. [Latest changelog](https://github.com/MoscaDotTo/Winapp2/blob/master/diff.txt).
| [CCleaner 7 winapp2.ini](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2-ccleaner7.ini)                   | The CCleaner 7 flavor, converted to the entry format CCleaner 7 requires. You don't install this file by hand: winapp2ool downloads it and patches it into `ccleaner.ini` for you (see [CCleaner 7](INSTALL.md#ccleaner-7)). [Latest changelog](https://github.com/MoscaDotTo/Winapp2/blob/master/Non-CCleaner/CCleaner7/diff.txt).
| [BleachBit winapp2.ini](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2-bleachbit.ini)                    | The BleachBit flavor, **for BleachBit 5 and older only.** It removes the registry exclusions those versions' sanity checker rejects. BleachBit 6 and newer handle them fine, so use the base winapp2.ini with those. [Latest changelog](https://github.com/MoscaDotTo/Winapp2/blob/master/Non-CCleaner/BleachBit/diff.txt).
| [System Ninja winapp2.rules](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2-systemninja.rules)           | The System Ninja flavor, which replaces keys System Ninja can't run with ones it can. [Latest changelog](https://github.com/MoscaDotTo/Winapp2/blob/master/Non-CCleaner/SystemNinja/diff.txt).
| [Tron winapp2.ini](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2-tron.ini)                              | The Tron flavor, which carries the changes Tron makes to its own winapp2.ini. [Latest changelog](https://github.com/MoscaDotTo/Winapp2/blob/master/Non-CCleaner/Tron/diff.txt).
| [FluentCleaner winapp2.ini](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2-fluentcleaner.ini)            | The FluentCleaner flavor, matching the edits FluentCleaner makes to the copy it ships. [Latest changelog](https://github.com/MoscaDotTo/Winapp2/blob/master/Non-CCleaner/FluentCleaner/diff.txt).
| [Winapp3.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/master/Winapp3/Winapp3.ini)                                   | An extension for an extension, with entries for power users. *Do **not** use this file if you don't know what you're doing. Some of its entries are very aggressive and can do real damage to your file system.*
| [Assembler](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler)                                                         | The sources every published file above is built from, and the build script that drives them. Contributions go here. See [How Winapp2.ini is built](#how-winapp2ini-is-built).

---

# [Installation & Configuration](INSTALL.md)

Pick your cleaner to see how to install and configure winapp2.ini for it:

[CCleaner Classic](INSTALL.md#ccleaner-classic) &middot; [CCleaner 7](INSTALL.md#ccleaner-7) &middot; [BleachBit](INSTALL.md#bleachbit) &middot; [System Ninja](INSTALL.md#system-ninja) &middot; [Avira System Speedup](INSTALL.md#avira-system-speedup) &middot; [Tron](INSTALL.md#tron) &middot; [R-Wipe & Clean](INSTALL.md#r-wipe--clean) &middot; [HDCleaner](INSTALL.md#hdcleaner) &middot; [FluentCleaner](INSTALL.md#fluentcleaner) &middot; [Reg Organizer](INSTALL.md#reg-organizer)

Whatever cleaner you use, we strongly recommend keeping a copy of [winapp2ool.exe](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2ool.exe) in the same folder as winapp2.ini. That's how you keep it up to date.

---

# [How Winapp2.ini is built](#how-winapp2ini-is-built)

Nobody edits the published winapp2.ini files in this repository by hand. Every one of them, and every changelog beside it, is generated from the sources in [Assembler](Assembler) by [winapp2ool](winapp2ool/Readme.md), a console application we maintain alongside the database. Its readme covers the tool itself, and each stage below links the readme for the module that runs it.

### Build pipeline

Each build runs four stages, driven by [`build winapp2.ps1`](Assembler/build%20winapp2.ps1):

| Stage | What runs | What it does |
| :- | :- | :- |
| Generate | [EntryBuilder](winapp2ool/modules/entrybuilder/readme.md) | Expands the per-letter [base entry sources](Assembler/EntryBuilder) into artifacts. Sources can use a shorthand DSL and the shared scaffold catalogs |
| | [BrowserBuilder](winapp2ool/modules/browserbuilder/readme.md) | Produces one entry per browser per cleaning category, crossing [browser declarations](Assembler/BrowserBuilder) with the templates that apply to their engine |
| | [UWPBuilder](winapp2ool/modules/uwpbuilder/readme.md) | Generates Microsoft Store entries from [package declarations](Assembler/UWP), with the same scaffold catalogs |
| Merge | [Combine](winapp2ool/modules/combine/readme.md) | Joins the generated artifacts into one file in strict mode, which fails the build if two artifacts use the same entry name |
| Lint | [WinappDebug](winapp2ool/modules/winappdebug/README.md) | Checks the merged file's syntax and style, and saves its own corrections |
| Publish | [Flavorizer](winapp2ool/modules/transmute/Flavorizer/readme.md), [WinappDebug](winapp2ool/modules/winappdebug/README.md), and [Diff](winapp2ool/modules/diff/readme.md) | Applies each flavor's ruleset to make the six other variants and lints each one, then writes a changelog against the previously published version for every output whose content changed |

### CI/CD

The build artifacts under [Assembler/Entries](Assembler/Entries) are committed to the repository instead of living only for the length of a build. That lets us check the build against itself, and gives CI something to hold still while it varies one input at a time:

| Check | What it does |
| :- | :- |
| **Generated artifact guard** | Fails a change that edits build output without touching the source that produces it, and names the file you should have edited |
| **PR verify** | Builds winapp2.ini from `master`, then again with the pull request merged, and comments either the changelog the pull request would produce or the stage where it broke the build |
| **Artifact drift** | Regenerates every artifact into a scratch folder and compares it byte for byte with what's committed, which catches generator output changing when it shouldn't |
| **Binary parity** | Checks that the committed winapp2ool.exe still matches the source beside it. Otherwise every published file would come from code that isn't in the repository |
| **Update signature** | Checks that copies of winapp2ool already out there will accept the committed update: its signature has to verify, and `version.txt` has to match the exe |
| **Maintainer paths guard** | Fails a pull request from outside the project that touches winapp2ool's source, its committed exe, the build scripts, or the workflows |
| **Daily build and release** | Rebuilds from current sources. When the content changed, it commits the result and cuts a tagged release with every flavor and changelog attached |

Every build restamps the version line, so the daily build ignores that line when it decides whether the content changed. Otherwise an untouched rebuild would look like a change and publish an empty changelog every day.

---

# [Contributing](#contributing)

Contributions go to the source files under [Assembler](Assembler), never to the published files the build makes from them. [Our contributor guidelines](CONTRIBUTING.md) cover where each kind of entry lives, the syntax rules for each, and what the automated checks will tell you after you open a pull request. Editing in the GitHub web interface is enough for most contributions. You never need to run the build or generate an output file yourself.

---

# [Custom content](#custom-content)

Winapp2.ini doesn't cover non-English system configurations or portable software out of the box. If you need them, keep your own entries in a `custom.ini` file and use winapp2ool's [Transmute](winapp2ool/modules/transmute/readme.md) in `Add` mode to apply them to each new winapp2.ini. For a portable browser, [BrowserBuilder](winapp2ool/modules/browserbuilder/readme.md) can generate the entries from your own install's path.

Winapp2ool 1.6 replaced the Merge feature with Transmute. If you used `custom.ini` with Merge, see [Migrating from Merge](winapp2ool/modules/transmute/readme.md#migrating-from-merge) in the Transmute readme.