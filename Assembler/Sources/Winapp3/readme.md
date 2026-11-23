# Winapp3 Source Files

### What is this?

This folder holds the sources for every entry in winapp3.ini, split by first letter into 27 files. Changes to Winapp3 entries go here; see the [contributor guidelines](https://github.com/MoscaDotTo/Winapp2/blob/master/CONTRIBUTING.md#winapp3ini).

Winapp3 entries are more aggressive than the ones in winapp2.ini. Some delete files their software or your system needs, and some take a very long time to scan. That is why they live in their own file, which people add to their winapp2.ini only when they want them.

### How it is built

During the build, winapp2ool's [EntryBuilder](https://github.com/MoscaDotTo/Winapp2/blob/master/winapp2ool/modules/entrybuilder/readme.md) module reads all 27 files and writes one file, [`Assembler/Winapp3/winapp3.ini`](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/Winapp3). The build lints it and publishes it as `winapp3.ini` on every release, with its own changelog. Comments stay in the sources, and do not reach the generated file.

Most entries are written in plain winapp2.ini syntax, which passes through EntryBuilder unchanged. EntryBuilder's shorthand works here too.

An entry's `Section` decides its group in winapp3.ini: `Language Files`, `Dangerous Long`, or any other section whose name starts with `Dangerous`, such as `Dangerous Games`.

### Rules

* An entry goes in the file for the first character of its name: `A.ini` through `Z.ini` for a letter, in any case, and `#.ini` for anything else. A pull request that puts an entry in the wrong file fails the [generated artifact guard](https://github.com/MoscaDotTo/Winapp2/blob/master/CONTRIBUTING.md#the-generated-artifact-guard), which names the file it belongs in
* A name in winapp3.ini cannot also be in winapp2.ini or archived.ini. The build fails when one is
* `Archive=` does not work here. To retire a Winapp3 entry, delete it

# Files

| File | Description |
| :- | :- |
| `#.ini` | Entries for applications whose names start with a number or symbol |
| `A.ini` through `Z.ini` | Entries for applications whose names start with that letter |
