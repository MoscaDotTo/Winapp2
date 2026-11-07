# BrowserBuilder

**BrowserBuilder** is a winapp2ool build-chain module for winapp2.ini maintainers that generates winapp2.ini entries for web browsers using a small scripting interface. Rather than writing individual entries for every browser and every entry type by hand, you define browser metadata and entry templates once. BrowserBuilder then cross-applies them to produce a consistent set of entries for every supported browser automatically.

BrowserBuilder is primarily an internal devops tool used to maintain the browser sections of the winapp2.ini distribution. End users with browsers installed in non-standard locations or using portable installations can also use it to generate entries for those scenarios (see [Example 8](EXAMPLES.md#example-8-generating-entries-for-a-portable-browser)).

### What does BrowserBuilder do?

BrowserBuilder reads two DSL files (`chromium.ini` and `gecko.ini`) from a single configurable source directory, each containing **BrowserInfo** sections describing individual browsers and **EntryScaffold** sections describing entry templates. It cross-applies every scaffold against every browser to produce a complete set of generated entries, applies Flavorizer corrections from the same directory to correct over- and under-coverage, normalizes the result with WinappDebug, and saves it to `browsers.ini`.

### Why BrowserBuilder?

- Consistency: every browser receives the same baseline coverage, generated from the same templates. No browser lags behind when a template location is discovered
- Scale: at the time of writing, 62 `BrowserInfo` sections × 35 `EntryScaffold` templates generate over 1,100 entries from two small ruleset files
- Low-cost maintenance: adding support for a new browser means writing one 3-5 line `BrowserInfo` section, no code changes; fixing a scaffold once fixes it for every browser
- Custom coverage: end users can generate entries for portable or non-standard installations that the distributed winapp2.ini cannot cover

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
   - [Special Case: Opera GX](#special-case-opera-gx)
7. [Post-Generation Flavorizing](#post-generation-flavorizing)
8. [Output Format](#output-format)
9. [Tips & Best Practices](#tips--best-practices)
10. [Troubleshooting](#troubleshooting)

**Examples**

11. **[Examples](EXAMPLES.md)**: eight worked examples with captured input and output

---

# Requirements

- A source directory containing at least one of `chromium.ini` or `gecko.ini` with valid `BrowserInfo` and `EntryScaffold` sections
- Optionally, [flavor correction files](#post-generation-flavorizing) in the same directory, resolved by their fixed names

All input files are read from the single source directory. The individual files are not separately configurable; pointing BrowserBuilder at a directory is the entire configuration.

---

# Quick Start

### Common Workflow

1. Place `chromium.ini` and/or `gecko.ini` (and any flavor correction files) in one directory
2. From the winapp2ool main menu, select Entry Lab, and then select BrowserBuilder 
3. Set the source directory to that location (or launch winapp2ool from it; the default is the current directory)
4. Run: BrowserBuilder generates entries for all browsers defined in the rulesets, applies any flavor corrections found next to them, and saves the result to `browsers.ini`

---

# Common Uses

| To | Run | Example |
|:-|:-|:-|
| Generate entries from the rulesets in the current directory | `winapp2ool -browserbuilder` | |
| Generate the project's browser entries from its rulesets and flavor files, saving `browsers.ini` in the current directory | `winapp2ool -browserbuilder -1d "C:\Winapp2\Assembler\BrowserBuilder"` | |
| Generate from a ruleset folder and save the result under another name in another folder | `winapp2ool -browserbuilder -1d "C:\rulesets" -2d "C:\output" -2f generated.ini` | |
| Generate entries for a portable or relocated browser from your own ruleset | `winapp2ool -browserbuilder -1d C:\myrules -2d C:\myrules` | [Generating Entries for a Portable Browser](EXAMPLES.md#example-8-generating-entries-for-a-portable-browser) |
| Add those generated entries to your winapp2.ini | `winapp2ool -transmute -add -1f winapp2.ini -2d C:\myrules -2f browsers.ini -3f winapp2.ini` | [Generating Entries for a Portable Browser](EXAMPLES.md#example-8-generating-entries-for-a-portable-browser) |

In the menu, the same choices are **Choose source directory** and **Choose save target**, set before **Run**; see [Menu Options](#menu-options). The last row is a [Transmute](../transmute/readme.md) command.

---

# Menu Options

| Option | Effect | Notes |
|:-|:-|:-|
| Run (default) | Generate browser entries from the configured source directory | Requires at least one of `chromium.ini` or `gecko.ini` in the source directory |
| Choose source directory | Select the directory containing `chromium.ini`, `gecko.ini`, and flavor correction files | Default: current directory |
| Choose save target | Select the output file path | Default: `browsers.ini` in the current directory |
| Reset Settings | Restore all settings to their defaults | Only shown when settings have been changed |

The menu also displays the current source directory and save target below the options.

---

# Command-Line Arguments

BrowserBuilder accepts two file parameters: the source directory and the output file. There are no toggle arguments. The ruleset and flavor file names within the source directory are fixed and cannot be overridden from the command line.

## File Selection

| Arg | Effect | Default |
|:-|:-|:-|
| `-1d path` | Set the source directory containing `chromium.ini`, `gecko.ini`, and flavor correction files | Current directory |
| `-2d path` | Set the output file path | Current directory |
| `-2f name` | Set the output file name | `browsers.ini` |

---

# Source File Format

BrowserBuilder reads standard ini files containing two types of sections: `BrowserInfo` sections and `EntryScaffold` sections. Sections of any other type are ignored with a warning.

```ini
[BrowserInfo: Browser Name]
Section=Category Name
UserDataPath=%LocalAppData%\BrowserName\User Data
RegistryRoot=HKCU\Software\BrowserName

[EntryScaffold: Entry Description]
FileKeyBase=<UserDataPath>\*\Cache|*|REMOVESELF
```

The section type prefixes are **case-sensitive** and must include the space after the colon: exactly `[BrowserInfo: ...]` and `[EntryScaffold: ...]`. A section named `[browserinfo: ...]` or `[BrowserInfo:...]` is ignored as invalid. Key names *within* sections are case-insensitive.

## BrowserInfo Sections

Each `BrowserInfo` section describes a single browser. The browser name from the section header is prepended to the entry name in every generated entry.

| Key | Effect | Notes |
|:-|:-|:-|
| `Section=<value>` | The `Section=` value for all entries generated for this browser | Required |
| `UserDataPath=<path>` | The browser's user data directory | Required; multiple keys allowed for browsers with more than one possible path |
| `RegistryRoot=<path>` | A registry path root for the browser | Optional; multiple keys allowed; used as `<RegistryRoot>` in `RegKeyBase` templates |
| `TruncateDetect` | Use the parent of `UserDataPath` for the `DetectFile` instead | No value needed; useful when the path's parent directory contains a version wildcard |
| `Skip` | Exclude this browser from generation | No value needed; the section is read but no entries are produced |

### Example

```ini
[BrowserInfo: Google Chrome]
Section=Google Chrome Web Browser
UserDataPath=%LocalAppData%\Google\Chrome\User Data
RegistryRoot=HKCU\Software\Google\Chrome
```

`Skip` and `TruncateDetect` are presence-based: providing the key at all (regardless of value) activates the behavior. `Skip=False` **will** cause the browser to be skipped.

Any other key in a `BrowserInfo` section produces a warning and is otherwise ignored.

## EntryScaffold Sections

Each `EntryScaffold` section is a template for one winapp2.ini entry type. The description from the section header is appended to the browser name in every generated entry.

| Key | Effect | Notes |
|:-|:-|:-|
| `FileKeyBase=<template>` | A FileKey template to generate for each browser | Multiple keys allowed; must contain exactly one kind of [file token](#template-tokens), or the build fails |
| `RegKeyBase=<template>` | A RegKey template to generate for each browser | Multiple keys allowed; must contain `<RegistryRoot>`, or the build fails |
| `RequiresRegistryRoot` | Skip this scaffold for any browser without a `RegistryRoot` | No value needed; prevents generating entries with no working keys |

### Example

```ini
[EntryScaffold: Web Browsing Cookies]
FileKeyBase=<UserDataPath>\*\Network|Cookies*;Device Bound Sessions*
```

For every `BrowserInfo` section, this scaffold produces an entry named `<browser name> Web Browsing Cookies *` with the corresponding `FileKey` values substituted in.

Any other key in an `EntryScaffold` section produces a warning and is otherwise ignored.

## Template Tokens

| Token | Replaced with | Available in |
|:-|:-|:-|
| `<UserDataPath>` | The browser's user data path from `BrowserInfo` | `FileKeyBase` |
| `<UserDataParent>` | The parent directory of `UserDataPath` | `FileKeyBase` |
| `<LocalUserDataPath>` | `UserDataPath` with `%AppData%` swapped for `%LocalAppData%` | `FileKeyBase` (Gecko only) |
| `<RegistryRoot>` | A registry root path from `BrowserInfo` | `RegKeyBase` |

Tokens match in any case, so `<userdatapath>` works too. These four are the whole vocabulary: unlike EntryBuilder and UWPBuilder, BrowserBuilder has no variable declarations.

After the tokens are substituted, each value goes through the same `<>` expander the other builders use, with nothing declared. That makes inline lists work: `FileKeyBase=<UserDataPath>\*\<Cache,Code Cache>|*|REMOVESELF` produces one FileKey per value, and a `RegKeyBase` can hold one too.

### Refused templates

BrowserBuilder checks each template once, before any browser sees it. A template that breaks one of these rules is dropped with a warning naming it, and the build fails (exit code 1). The scaffold's other templates still build.

| Template | Warning | Before these checks |
|:-|:-|:-|
| A `FileKeyBase` with no file token | `[EntryScaffold: X].FileKeyBase has no file token; a FileKeyBase needs one of <UserDataPath>, <UserDataParent> or <LocalUserDataPath>` | Silently skipped for every browser |
| A `FileKeyBase` with two kinds of file token | ``[EntryScaffold: X].FileKeyBase mixes `<UserDataParent>` and `<UserDataPath>`; a FileKeyBase takes one kind of file token`` | One token was substituted (the user data path ahead of the parent, the parent ahead of the local path) and the other shipped as literal text |
| `<LocalUserDataPath>` in `chromium.ini` | `[EntryScaffold: X].FileKeyBase uses <LocalUserDataPath>, which only Gecko browsers have` | Silently skipped for every browser |
| `<RegistryRoot>` in a `FileKeyBase` | `[EntryScaffold: X].FileKeyBase uses <RegistryRoot>, which only a RegKeyBase takes` | Shipped as literal text beside a file token, or skipped without one |
| A file token in a `RegKeyBase` | ``[EntryScaffold: X].RegKeyBase uses the file token `<UserDataPath>`; a RegKeyBase takes only <RegistryRoot>`` | Shipped as literal text, once per root |
| A `RegKeyBase` without `<RegistryRoot>` | `[EntryScaffold: X].RegKeyBase has no <RegistryRoot> token` | Emitted unchanged, once per root |
| Any other `<Name>` token, such as a misspelling | ``Unknown token `<Typo>` in [EntryScaffold: X].FileKeyBase; BrowserBuilder recognizes <UserDataPath>, <UserDataParent>, <LocalUserDataPath> and <RegistryRoot>`` | A misspelled `%...%` token shipped as literal text, except in a `FileKeyBase` with no other token, which was skipped |
| An inline list with no values, such as `<,>` | ``Inline list `<,>` in [EntryScaffold: X].FileKeyBase has no values`` | New with inline lists. Without the check it drops for every browser and the run still succeeds. A list with at least one value (`<Cache,>`) is fine: the empty one is dropped |
| An empty, stray or nested bracket: `<>`, `<UserDataPath>>`, `<<UserDataPath>,z>` | ``Malformed `<...>` token in [EntryScaffold: X].FileKeyBase; a token is one <Name> or one <a,b> list, not empty or nested`` | New with the `<Name>` grammar. Without the check the bracket ships as text, and a nested list ships its other value as the path |

Each warning ends with ``; dropping `FileKeyBase=...` `` or ``; dropping `RegKeyBase=...` ``, naming the whole key.

### BrowserInfo values

`UserDataPath=` and `RegistryRoot=` values are literal paths and take no tokens. A value holding `<` or `>` is skipped with a warning, and the build fails; the browser's other values still build:

```
[BrowserInfo: Chrome].UserDataPath holds `<` or `>`, but BrowserInfo values are literal paths and take no tokens; skipping `UserDataPath=%LocalAppData%\<Foo>\User Data`
```

Without the check such a value would land in each template ahead of the inline-list pass, so `RegistryRoot=HKCU\Software\<a,b>` would fan out into two roots and `<Foo>` in a `UserDataPath` would drop every FileKey.

### Retired spelling

Older versions of winapp2ool spelled these tokens `%UserDataPath%`, `%BrowserPath%`, `%LocalDataPath%` and `%RegistryRoot%`. `%X%` is how winapp2.ini writes environment variables, so a mistyped token looked like a real path. The old spelling is now refused like the templates above, with a warning naming the new one. Two tokens were renamed, not only respelled: `%BrowserPath%` is the folder above the user data folder, not the install folder, so it is now `<UserDataParent>`, and `%LocalDataPath%` is now `<LocalUserDataPath>`. A token with its closing `%` missing is caught too:

```
`%BrowserPath%` in [EntryScaffold: Logs].FileKeyBase is now spelled `<UserDataParent>`; dropping `FileKeyBase=%BrowserPath%\Application|debug.log`
`%BrowserPath` in [EntryScaffold: Logs].FileKeyBase looks like the retired `%BrowserPath%` token, which is now spelled `<UserDataParent>`; dropping `FileKeyBase=%BrowserPath\Application|debug.log`
```

`BrowserInfo` values aren't checked for the retired spellings: they are literal paths, and one holding `<` or `>` is refused [on its own](#browserinfo-values).

### Token behavior details

**`<UserDataPath>`** is substituted directly with the `UserDataPath` value. For browsers with multiple `UserDataPath` keys, one `DetectFile` and one set of FileKeys is generated per path. See [Example 3](EXAMPLES.md#example-3-multiple-userdatapaths).

**`<UserDataParent>`** resolves to the parent directory of `UserDataPath`. For Chromium browsers, each `<UserDataParent>` template generates a *second* FileKey with `%LocalAppData%` or `%AppData%` replaced by `%ProgramFiles%`, covering system-wide (All Users) installations. See [Example 6](EXAMPLES.md#example-6-userdataparent-and-the-programfiles-mirror). Gecko browsers do not receive the `%ProgramFiles%` mirror.

**`<LocalUserDataPath>`** is a Gecko-only token that takes `UserDataPath` and replaces `%AppData%` with `%LocalAppData%`, targeting the local (non-roaming) profile data directory. See [Example 7](EXAMPLES.md#example-7-localuserdatapath-gecko).

**`<RegistryRoot>`** is substituted in `RegKeyBase` templates. One `RegKey` is generated per `RegistryRoot` key per `RegKeyBase` template. See [Example 5](EXAMPLES.md#example-5-registry-scaffolds-and-requiresregistryroot).

## Special Case: Opera GX

Opera GX stores profile data directly in its user data directory and keeps additional profiles under `_side_profiles\*` rather than the usual `Default`/`Profile 1` folders. BrowserBuilder recognizes the browser name `Opera GX` and rewrites every `<UserDataPath>` template into two FileKeys: one against the user data directory itself and one against `_side_profiles\*`. If you diff the generated output, Opera GX entries will structurally differ from every other Chromium browser's; this is intentional.

---

# Post-Generation Flavorizing

After generating entries, BrowserBuilder applies Flavorizer to the output before saving. This corrects coverage issues that arise from the template-based generation approach. For example: removing entries for features a particular browser does not support, or adding vendor-specific keys that cannot be templated.

The flavor files are read from the source directory. The names are not configurable:

| Stage | Operation | File name (fixed) |
|:-|:-|:-|
| 1 | Section removals | `browser_section_removals.ini` |
| 2 | Key name removals | `browser_name_removals.ini` |
| 3 | Key value removals | `browser_value_removals.ini` |
| 4 | Section replacements | `browser_section_replacements.ini` |
| 5 | Key replacements | `browser_key_replacements.ini` |
| 6 | Additions | `browser_additions.ini` |

All flavor files are optional; missing files are silently skipped.

For additional context, see the [Flavorizer README](../transmute/Flavorizer/readme.md) and the [Transmute README's Usage Examples](../transmute/EXAMPLES.md).

---

# Output Format

The output file (`browsers.ini` by default) contains standard winapp2.ini format entries. Each generated entry has:

- A section header in the form `[<browser name> <scaffold name> *]`
- A `Section=` key from the browser's `BrowserInfo`
- `DetectFile=` (or numbered `DetectFile1=`, `DetectFile2=` for multiple paths) pointing to the `UserDataPath` (or its parent, if `TruncateDetect` is set)
- Numbered `FileKey` and `RegKey` entries from the scaffold templates with tokens substituted

Before saving, the entire output is run through **WinappDebug** with full optimizations forced: entries are alphabetized, keys are placed in standard winapp2.ini order, key numbering is normalized, and parameter lists are alphabetized. 

The saved file begins with a generated comment header:

```ini
; Version <yyMMdd>
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

Raw generated output will have coverage issues: entries for features not supported by certain browsers, or missing browser-specific keys and entries. Flavor files are how these are corrected. They must sit in the source directory under their exact fixed names to be picked up.

### Skipping Discontinued Browsers

Use `Skip` in a `BrowserInfo` section to exclude a browser without deleting its data. This preserves the section for potential future re-activation.

### Multiple UserDataPaths

Provide multiple `UserDataPath` keys in a `BrowserInfo` section to support browsers with more than one possible profile location (e.g., a standalone install and a Microsoft Store package). BrowserBuilder generates a `DetectFile` and a full set of FileKeys for each path.

### TruncateDetect

Use `TruncateDetect` when `UserDataPath` contains a wildcard in its direct parent directory name (e.g., `%LocalAppData%\BraveSoftware\Brave-Browser*\User Data`). Some cleaning tools (notably CCleaner) do not support wildcards in `DetectFile` parent paths; truncating moves the detection up to the wildcarded directory itself. See [Example 4](EXAMPLES.md#example-4-truncatedetect).

---

# Troubleshooting

| Symptom | Cause |
|:-|:-|
| "No valid generative rulesets found" | Neither `chromium.ini` nor `gecko.ini` in the source directory contains any parseable sections; check the source directory setting |
| Warning: "Invalid section found and ignored: [section]" | A section's name does not start with exactly `BrowserInfo:` or `EntryScaffold:`; the prefixes are case-sensitive |
| Warning: "Unexpected KeyType in [section]: [key]" | A `BrowserInfo` or `EntryScaffold` section contains a key name the module doesn't recognize; the key is ignored |
| Warning: "No valid UserDataPath key found in [browser]" | The `BrowserInfo` section is missing its `UserDataPath` key; the browser's entries will have no `DetectFile` and no FileKeys |
| Warning: "No valid Section key found in [browser]" | The `BrowserInfo` section is missing its `Section` key; generated entries will carry an empty `Section=` |
| Warning: "Skipping [browser] for [scaffold]: scaffold requires RegistryRoot" | The scaffold declares `RequiresRegistryRoot` and the browser has no `RegistryRoot` key; no entry is generated for that pairing |
| A browser is missing from the output entirely | It may have `Skip` set; check its `BrowserInfo` section |
| One entry is missing for one browser | Likely a `RequiresRegistryRoot` scaffold and a browser without a `RegistryRoot` |
| A scaffold key is missing from every entry, and the build failed | The template was [refused](#refused-templates); the warning names it and the reason |
| A browser lost a path or registry root, and the build failed | The `BrowserInfo` value held `<` or `>` and was [skipped](#browserinfo-values) |
| Flavor corrections aren't being applied | The flavor file's name doesn't exactly match its fixed name, or it isn't in the source directory. Missing files are skipped silently |
| Generated entries are incomplete or wrong | Flavor files correct coverage issues; ensure the appropriate flavor files are present in the source directory |
