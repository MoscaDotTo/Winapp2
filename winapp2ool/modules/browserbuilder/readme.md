# BrowserBuilder

**BrowserBuilder** is a winapp2ool build-chain module for winapp2.ini maintainers that generates winapp2.ini entries for web browsers from a small set of rules. You describe each browser once and each kind of entry once, and BrowserBuilder writes out every combination of the two.

Its main job is building the browser entries in the published winapp2.ini, and most people will never need to run it. The exception is a browser the published file can't find, like a portable copy or one installed somewhere unusual. BrowserBuilder can generate entries for that install from your own path (see [Example 8](EXAMPLES.md#example-8-generating-entries-for-a-portable-browser)).

### What does BrowserBuilder do?

BrowserBuilder reads two rulesets, `chromium.ini` and `gecko.ini`, from one source directory. Each holds `BrowserInfo` sections, which say where a browser keeps its data, and `EntryScaffold` sections, which are entry templates. Every scaffold is applied to every browser in the same ruleset. BrowserBuilder then applies any Flavorizer correction files it finds in that directory, normalizes the result with WinappDebug, and saves it to `browsers.ini`.

### Why BrowserBuilder?

Every browser gets the same coverage because every browser's entries come from the same templates. When we find another place Chromium browsers keep a cache, we add one line to one scaffold and every Chromium browser picks it up. Adding a browser takes a short `BrowserInfo` section and no code changes. The rulesets stay small while the output grows, since each new browser multiplies against every scaffold and each new scaffold against every browser.

The same rules let an end user cover a portable or relocated install that the published winapp2.ini can't.

---

# Table of Contents

**Using BrowserBuilder**

1. [Requirements](#requirements)
2. [Quick Start](#quick-start)
3. [Common Uses](#common-uses)
4. [Menu Options](#menu-options)
5. [Command-Line Arguments](#command-line-arguments)
   - [File Selection](#file-selection)
6. [Source File Format](#source-file-format)
   - [BrowserInfo Sections](#browserinfo-sections)
   - [EntryScaffold Sections](#entryscaffold-sections)
   - [Template Tokens](#template-tokens)
   - [Variables](#variables)
   - [Special Case: Opera GX](#special-case-opera-gx)
7. [Post-Generation Flavorizing](#post-generation-flavorizing)
8. [Output Format](#output-format)
9. [Tips & Best Practices](#tips--best-practices)
10. [Troubleshooting](#troubleshooting)

**Examples**

11. **[Examples](EXAMPLES.md)**: eight worked examples with real input and output

---

# Requirements

- A source directory holding `chromium.ini`, `gecko.ini`, or both, with `BrowserInfo` and `EntryScaffold` sections in them
- Optionally, [flavor correction files](#post-generation-flavorizing) in the same directory

Every input file comes from that one directory, under a fixed name. Choosing the directory is the whole configuration.

---

# Quick Start

### Common Workflow

1. Put `chromium.ini` and/or `gecko.ini`, plus any flavor correction files, in one directory
2. From the winapp2ool main menu, select Entry Lab, then BrowserBuilder
3. Set the source directory to that directory, or launch winapp2ool from it (the default is the current directory)
4. Choose **Run**. BrowserBuilder generates entries for every browser in the rulesets, applies any flavor files beside them, and saves the result to `browsers.ini`

---

# Common Uses

| To | Run | Example |
|:-|:-|:-|
| Generate entries from the rulesets in the current directory | `winapp2ool -browserbuilder` | |
| Generate the project's browser entries from its rulesets and flavor files, saving `browsers.ini` in the current directory | `winapp2ool -browserbuilder -1d "C:\Winapp2\Assembler\Sources\BrowserBuilder"` | |
| Generate from a ruleset folder and save the result under another name in another folder | `winapp2ool -browserbuilder -1d "C:\rulesets" -2d "C:\output" -2f generated.ini` | |
| Generate entries for a portable or relocated browser from your own ruleset | `winapp2ool -browserbuilder -1d C:\myrules -2d C:\myrules` | [Generating Entries for a Portable Browser](EXAMPLES.md#example-8-generating-entries-for-a-portable-browser) |
| Add those generated entries to your winapp2.ini | `winapp2ool -transmute -add -1f winapp2.ini -2d C:\myrules -2f browsers.ini -3f winapp2.ini` | [Generating Entries for a Portable Browser](EXAMPLES.md#example-8-generating-entries-for-a-portable-browser) |

In the menu, set **Choose source directory** and **Choose save target**, then choose **Run**; see [Menu Options](#menu-options). The last row is a [Transmute](../transmute/readme.md) command.

---

# Menu Options

| Option | Effect | Notes |
|:-|:-|:-|
| Run (default) | Generate browser entries from the configured source directory | Needs `chromium.ini` or `gecko.ini` in the source directory |
| Choose source directory | Select the directory holding `chromium.ini`, `gecko.ini`, and the flavor correction files | Default: current directory |
| Choose save target | Select the output file | Default: `browsers.ini` in the current directory |
| Reset Settings | Restore all settings to their defaults | Only shown when settings have been changed |

Below the options, the menu shows the current source directory and save target.

---

# Command-Line Arguments

BrowserBuilder takes two file arguments, the source directory and the output file, and no toggles. The rulesets and flavor files inside the source directory always go by their fixed names.

## File Selection

| Arg | Effect | Default |
|:-|:-|:-|
| `-1d path` | Set the source directory holding `chromium.ini`, `gecko.ini`, and the flavor correction files | Current directory |
| `-2d path` | Set the directory the output is saved to | Current directory |
| `-2f name` | Set the output file name | `browsers.ini` |

---

# Source File Format

BrowserBuilder reads ordinary ini files with two kinds of section, `BrowserInfo` and `EntryScaffold`. Any other section is ignored with a warning, and the build fails (exit code 1).

```ini
[BrowserInfo: Browser Name]
Section=Category Name
UserDataPath=%LocalAppData%\BrowserName\User Data
RegistryRoot=HKCU\Software\BrowserName

[EntryScaffold: Entry Description]
FileKeyBase=<UserDataPath>\*\Cache|*|REMOVESELF
```

Neither the section prefixes nor the key names are case-sensitive, and the space after the colon is optional, so `[browserinfo:Chrome]` works the same as `[BrowserInfo: Chrome]`. A section needs a name after its prefix.

## BrowserInfo Sections

Each `BrowserInfo` section describes one browser. The name in its header starts the name of every entry generated for that browser.

| Key | Effect | Notes |
|:-|:-|:-|
| `Section=<value>` | The `Section=` value for every entry generated for this browser | Required |
| `UserDataPath=<path>` | The browser's user data directory | Required. Repeat it for a browser that can keep its data in more than one place |
| `RegistryRoot=<path>` | A registry root for the browser | Optional, and can repeat. Fills `<RegistryRoot>` in `RegKeyBase` templates |
| `TruncateDetect` | Detect on the parent of `UserDataPath` instead | No value needed. For a `UserDataPath` whose parent folder has a wildcard in its name; see [TruncateDetect](#truncatedetect) |
| `Skip` | Leave this browser out | No value needed. The section is read but produces no entries |

### Example

```ini
[BrowserInfo: Google Chrome]
Section=Google Chrome Web Browser
UserDataPath=%LocalAppData%\Google\Chrome\User Data
RegistryRoot=HKCU\Software\Google\Chrome
```

`Skip` and `TruncateDetect` take effect whenever the key is there, whatever its value. `Skip=False` **still skips** the browser.

Any other key declares one of the browser's [variables](#variables).

## EntryScaffold Sections

Each `EntryScaffold` section is the template for one kind of entry. The name in its header follows the browser name in every entry it generates.

| Key | Effect | Notes |
|:-|:-|:-|
| `FileKeyBase=<template>` | A FileKey template to generate for each browser | Can repeat. Must use exactly one kind of [file token](#template-tokens), or the build fails |
| `RegKeyBase=<template>` | A RegKey template to generate for each browser | Can repeat. Must use `<RegistryRoot>`, or the build fails |
| `RequiresRegistryRoot` | Skip this scaffold for any browser without a `RegistryRoot` | No value needed. Without it, a browser with no root still gets the entry, minus its RegKeys |

### Example

```ini
[EntryScaffold: Web Browsing Cookies]
FileKeyBase=<UserDataPath>\*\Network|Cookies*;Device Bound Sessions*
```

Every browser gets an entry from this scaffold named `<browser name> Web Browsing Cookies *`, with its own path filled into the `FileKey`.

Any other key declares one of the scaffold's [variables](#variables).

## Template Tokens

| Token | Replaced with | Available in |
|:-|:-|:-|
| `<UserDataPath>` | The browser's user data path from `BrowserInfo` | `FileKeyBase` |
| `<UserDataParent>` | The parent directory of `UserDataPath` | `FileKeyBase` |
| `<LocalUserDataPath>` | `UserDataPath` with `%AppData%` swapped for `%LocalAppData%` | `FileKeyBase` (Gecko only) |
| `<RegistryRoot>` | A registry root path from `BrowserInfo` | `RegKeyBase` |

Tokens match in any case, so `<userdatapath>` works too.

Once the tokens are filled in, each value goes through the same `<>` expander the other builders use, with the scaffold's [variables](#variables). Inline lists work there too: `FileKeyBase=<UserDataPath>\*\<Cache,Code Cache>|*|REMOVESELF` produces one FileKey per value, and a `RegKeyBase` can hold a list as well.

### Refused templates

BrowserBuilder checks each template once, before generating anything from it. A template that breaks one of these rules is dropped with a warning naming it, and the build fails (exit code 1). The scaffold's other templates still build.

| Template | Warning | Before these checks |
|:-|:-|:-|
| A `FileKeyBase` with no file token | `[EntryScaffold: X].FileKeyBase has no file token; a FileKeyBase needs one of <UserDataPath>, <UserDataParent> or <LocalUserDataPath>` | Silently skipped for every browser |
| A `FileKeyBase` with two kinds of file token | ``[EntryScaffold: X].FileKeyBase mixes `<UserDataParent>` and `<UserDataPath>`; a FileKeyBase takes one kind of file token`` | One token was filled in (the user data path ahead of the parent, the parent ahead of the local path) and the other shipped as literal text |
| `<LocalUserDataPath>` in `chromium.ini` | `[EntryScaffold: X].FileKeyBase uses <LocalUserDataPath>, which only Gecko browsers have` | Silently skipped for every browser |
| `<RegistryRoot>` in a `FileKeyBase` | `[EntryScaffold: X].FileKeyBase uses <RegistryRoot>, which only a RegKeyBase takes` | Shipped as literal text beside a file token, or skipped without one |
| A file token in a `RegKeyBase` | ``[EntryScaffold: X].RegKeyBase uses the file token `<UserDataPath>`; a RegKeyBase takes only <RegistryRoot>`` | Shipped as literal text, once per root |
| A `RegKeyBase` without `<RegistryRoot>` | `[EntryScaffold: X].RegKeyBase has no <RegistryRoot> token` | Written out unchanged, once per root |
| Any other `<Name>` token, such as a misspelling | ``Unknown token `<Typo>` in [EntryScaffold: X].FileKeyBase; BrowserBuilder recognizes <UserDataPath>, <UserDataParent>, <LocalUserDataPath>, <RegistryRoot> and the scaffold's own variables`` | A misspelled `%...%` token shipped as literal text, except in a `FileKeyBase` with no other token, which was skipped |
| An inline list with no values, such as `<,>` | ``Inline list `<,>` in [EntryScaffold: X].FileKeyBase has no values`` | New with inline lists. Without the check the template disappears for every browser and the run still succeeds. A list with at least one value (`<Cache,>`) is fine, and the empty value is dropped |
| An empty, stray or nested bracket: `<>`, `<UserDataPath>>`, `<<UserDataPath>,z>` | ``Malformed `<...>` token in [EntryScaffold: X].FileKeyBase; a token is one <Name> or one <a,b> list, not empty or nested`` | New with the `<Name>` grammar. Without the check the bracket ships as text, and a nested list ships its other value as the path |
| A variable with no values, or an index past the end of its list | ``Axis <Dirs> has no values, referenced by [EntryScaffold: X].FileKeyBase`` or ``[EntryScaffold: X].FileKeyBase: <Dirs[3]> is out of range (2 values)`` | New with variables |
| A variable whose value holds a token nothing declares | ``[EntryScaffold: X].FileKeyBase still holds a `<...>` token once its variables are filled in; a variable's value can't hold a token nothing declares`` | New with variables |

Each warning ends with ``; dropping `FileKeyBase=...` `` or ``; dropping `RegKeyBase=...` ``, naming the whole key.

### BrowserInfo values

`UserDataPath=` and `RegistryRoot=` values can use the browser's own [variables](#variables) and inline lists. Each value becomes one plain path per combination before anything else reads it, so `RegistryRoot=HKCU\Software\Microsoft\<Edge,Edge Beta>` is exactly the same as two `RegistryRoot` lines.

A value that can't be expanded is skipped with a warning, and the build fails. The browser's other values still build:

```
Unknown variable `<Foo>` in [BrowserInfo: Chrome].UserDataPath; skipping `UserDataPath=%LocalAppData%\<Foo>\User Data`
```

The rules are the same as for templates: no malformed brackets, no empty inline lists, no variables without values, and no variables holding a token nothing declares. The four scaffold tokens are refused here too, since only scaffolds take them.

### Token behavior details

**`<UserDataPath>`** is replaced with the `UserDataPath` value. A browser with several `UserDataPath` keys gets a `DetectFile` and a set of FileKeys for each path. See [Example 3](EXAMPLES.md#example-3-multiple-userdatapaths).

**`<UserDataParent>`** is the parent directory of `UserDataPath`. For Chromium browsers, each `<UserDataParent>` template also produces a *second* FileKey with `%LocalAppData%` or `%AppData%` changed to `%ProgramFiles%`, which covers installs made for all users. The swap only touches text written in the template itself, not text a scaffold variable fills in later. Gecko browsers don't get the `%ProgramFiles%` copy. See [Example 6](EXAMPLES.md#example-6-userdataparent-and-the-programfiles-mirror).

**`<LocalUserDataPath>`** is Gecko only. It's `UserDataPath` with `%AppData%` changed to `%LocalAppData%`, which is where Gecko keeps its local, non-roaming profile data. See [Example 7](EXAMPLES.md#example-7-localuserdatapath-gecko).

**`<RegistryRoot>`** is filled in `RegKeyBase` templates. Each `RegKeyBase` produces one RegKey per `RegistryRoot`. See [Example 5](EXAMPLES.md#example-5-registry-scaffolds-and-requiresregistryroot).

## Variables

Any key a section doesn't reserve declares a variable, named after the key without any trailing digits, whose values are its comma-separated list. `Dirs1=` and `Dirs2=` both declare `Dirs`, so the second replaces the first, with a warning. You use it as `<Name>`, and each value produces its own copy of the line. Variables work the way they do in [EntryBuilder](../entrybuilder/readme.md) and [UWPBuilder](../uwpbuilder/readme.md): names match in any case, one variable can use another, and `<Name[N]>` picks the Nth value, counting from 1.

Each section's variables are its own. A browser's variables fill only its `UserDataPath=` and `RegistryRoot=` values, and a scaffold's fill only its own templates. A scaffold can't use a browser's variable, so a mistyped token in a scaffold is still caught once, before anything is generated.

```ini
[BrowserInfo: Thunderbird]
Section=Thunderbird Email Client
Channels=MZLA,ThunderbirdBeta
UserDataPath=%LocalAppData%\Packages\MozillaThunderbird.<Channels>_h5892qc0xkpca\LocalCache\Roaming\Thunderbird\Profiles

[EntryScaffold: Caches]
Dirs=cache2,thumbnails
FileKeyBase=<UserDataPath>\*\<Dirs>|*|REMOVESELF
```

`Channels` turns the one `UserDataPath` into two, and `Dirs` turns the one template into two FileKeys per path:

```ini
[Thunderbird Caches *]
Section=Thunderbird Email Client
DetectFile1=%LocalAppData%\Packages\MozillaThunderbird.MZLA_h5892qc0xkpca\LocalCache\Roaming\Thunderbird\Profiles
DetectFile2=%LocalAppData%\Packages\MozillaThunderbird.ThunderbirdBeta_h5892qc0xkpca\LocalCache\Roaming\Thunderbird\Profiles
FileKey1=%LocalAppData%\Packages\MozillaThunderbird.MZLA_h5892qc0xkpca\LocalCache\Roaming\Thunderbird\Profiles\*\cache2|*|REMOVESELF
FileKey2=%LocalAppData%\Packages\MozillaThunderbird.MZLA_h5892qc0xkpca\LocalCache\Roaming\Thunderbird\Profiles\*\thumbnails|*|REMOVESELF
FileKey3=%LocalAppData%\Packages\MozillaThunderbird.ThunderbirdBeta_h5892qc0xkpca\LocalCache\Roaming\Thunderbird\Profiles\*\cache2|*|REMOVESELF
FileKey4=%LocalAppData%\Packages\MozillaThunderbird.ThunderbirdBeta_h5892qc0xkpca\LocalCache\Roaming\Thunderbird\Profiles\*\thumbnails|*|REMOVESELF
```

A section's reserved keys can't be variables, and neither can the four token names: a key called `UserDataParent=` is dropped with a warning, and the build fails. A variable nothing uses gets a warning but doesn't fail the build. That's how a misspelled key such as `RegistyRoot=` shows up, since it reads as a variable.

List values are trimmed and empty ones dropped, so `Edge<, Beta>` gives only `EdgeBeta`. Write each value out instead: `<Edge,Edge Beta>`.

## Special Case: Opera GX

Opera GX keeps its main profile directly in its user data directory, with any other profiles under `_side_profiles\*`, rather than in the usual `Default` and `Profile 1` folders. When a browser is named `Opera GX`, in any case, BrowserBuilder turns each `<UserDataPath>` template into two FileKeys: one against the user data directory itself, dropping the template's `\*` profile wildcard if it has one, and one against `_side_profiles\*`. That's why Opera GX's entries look different from every other Chromium browser's.

---

# Post-Generation Flavorizing

Templates can't know everything about every browser. Some browsers lack a feature a scaffold covers, and some write vendor-specific files that no template describes. Before saving, BrowserBuilder runs Flavorizer over the generated entries to remove the ones that shouldn't exist and add the keys that can't be templated.

The flavor files come from the source directory, under fixed names:

| Stage | Operation | File name (fixed) |
|:-|:-|:-|
| 1 | Section removals | `browser_section_removals.ini` |
| 2 | Key name removals | `browser_name_removals.ini` |
| 3 | Key value removals | `browser_value_removals.ini` |
| 4 | Section replacements | `browser_section_replacements.ini` |
| 5 | Key replacements | `browser_key_replacements.ini` |
| 6 | Additions | `browser_additions.ini` |

Every flavor file is optional, and a missing one is skipped without a warning.

For what each stage does, see the [Flavorizer README](../transmute/Flavorizer/readme.md) and the [Transmute examples](../transmute/EXAMPLES.md).

---

# Output Format

The output file (`browsers.ini` by default) is ordinary winapp2.ini. Each generated entry has:

- A section header in the form `[<browser name> <scaffold name> *]`
- A `Section=` key from the browser's `BrowserInfo`
- A `DetectFile=` pointing at the `UserDataPath`, or at its parent if `TruncateDetect` is set. A browser with several paths gets numbered `DetectFile1=`, `DetectFile2=` and so on
- Numbered `FileKey` and `RegKey` keys from the scaffold's templates, with the tokens filled in

Before saving, BrowserBuilder runs the whole output through **WinappDebug** with optimizations on. That sorts the entries, puts each entry's keys in standard winapp2.ini order and renumbers them, alphabetizes FileKey patterns, and merges FileKeys that share a path. If the lint would change what an entry means rather than just how it's written, the [reconciliation guard](../winappdebug/README.md#the-reconciliation-guard) fails the build instead.

The saved file starts with this header:

```ini
; # of entries: #,###
; browsers.ini is generated by the Winapp2ool Browser Builder
; Entries in this file may be incomplete and are not intended to be used directly with any cleaning software
; They are utilized by winapp2ool to create the final winapp2.ini file for distribution
; If you are not maintaining winapp2.ini for distribution, you probably don't need this file!
; Refer to the Winapp2ool documentation for more information: https://github.com/MoscaDotTo/Winapp2/blob/master/winapp2ool/Readme.md
; You can find the complete winapp2.ini file here: https://raw.githubusercontent.com/MoscaDotTo/Winapp2/master/Non-CCleaner/Winapp2.ini
```

---

# Tips & Best Practices

### Always Use Flavor Files

Generated output on its own has extras and gaps: entries for features some browsers don't have, and nothing for files only one vendor writes. The flavor files fix both. BrowserBuilder only finds them in the source directory, under their exact names.

### Skipping Discontinued Browsers

Add `Skip` to a browser's `BrowserInfo` section to stop generating its entries without deleting the section. If the browser comes back, delete the `Skip` line.

### Multiple UserDataPaths

Give a browser one `UserDataPath` key for each place it can keep its profiles, such as a standalone install and a Microsoft Store package. Each path gets its own `DetectFile` and its own set of FileKeys.

### TruncateDetect

Use `TruncateDetect` when the folder above `UserDataPath` has a wildcard in its name, as in `%LocalAppData%\BraveSoftware\Brave-Browser*\User Data`. CCleaner only accepts a wildcard in the last part of a `DetectFile` path, so we detect on the wildcarded folder itself. See [Example 4](EXAMPLES.md#example-4-truncatedetect).

---

# Troubleshooting

| Symptom | Cause |
|:-|:-|
| "No valid generative rulesets found", and the build failed | The source directory has no `chromium.ini` or `gecko.ini`, they have no sections, or the directory doesn't exist. Check the source directory setting |
| Warning: "Invalid section found and ignored: [section]", and the build failed | The section's name doesn't start with `BrowserInfo:` or `EntryScaffold:` followed by a name. Check the prefix's spelling |
| Warning: "Variable 'X=' declared in [section] but never referenced by any \<X> token; possible typo" | A key the section doesn't reserve declares a [variable](#variables), and nothing uses it. Usually a misspelled key, such as `RegistyRoot=` |
| Warning: "[section] declares `X`, which is a token, not a variable", and the build failed | A key is named after one of the four [template tokens](#template-tokens). Those names can't be variables |
| Warning: "No valid UserDataPath key found in [browser]" | The `BrowserInfo` section has no `UserDataPath` key, so the browser's entries get no `DetectFile` and no FileKeys |
| Warning: "No valid Section key found in [browser]" | The `BrowserInfo` section has no `Section` key, so its entries get an empty `Section=` |
| Warning: "Skipping [browser] for [scaffold]: scaffold requires RegistryRoot" | The scaffold has `RequiresRegistryRoot` and the browser has no `RegistryRoot`, so that pair produces no entry |
| A browser is missing from the output entirely | Check its `BrowserInfo` section for `Skip` |
| One entry is missing for one browser | Most likely a `RequiresRegistryRoot` scaffold and a browser with no `RegistryRoot` |
| A scaffold key is missing from every entry, and the build failed | The template was [refused](#refused-templates). The warning names it and says why |
| Warning: "[BrowserInfo: X].UserDataPath has no value; skipping it" | An empty `UserDataPath=` or `RegistryRoot=` line. It's ignored, and the build carries on |
| A browser lost a path or registry root, and the build failed | The `BrowserInfo` value couldn't be expanded and was [skipped](#browserinfo-values). The warning says why |
| The build failed and left a `browsers.prelint.ini` beside the output | The WinappDebug pass would have changed what an entry means. That's a bug in the rulesets or the generator; the log lists each change. See the [reconciliation guard](../winappdebug/README.md#the-reconciliation-guard) |
| Flavor corrections aren't being applied | The flavor file isn't in the source directory, or its name doesn't exactly match the fixed name. A missing file is skipped without a warning |
| Generated entries are incomplete or wrong | Without flavor files the output keeps the templates' gaps. Check that the flavor files are in the source directory |
