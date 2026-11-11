# Flavorizer Examples

Worked examples for [Flavorizer](readme.md). The readme's [Common Uses](readme.md#common-uses) table lists the commands. These show what each one does, with real input and output.

# Table of Contents

1. [Example 1: A Single-Stage Flavor](#example-1-a-single-stage-flavor)
2. [Example 2: A Multi-Stage Flavor](#example-2-a-multi-stage-flavor)
3. [Example 3: Auto-Detecting a Flavor](#example-3-auto-detecting-a-flavor)
4. [Example 4: Global Operations in Flavor Files](#example-4-global-operations-in-flavor-files)

---

###### Note: With winapp2.ini formatting on (the default), the output file starts with the winapp2.ini header comments. The outputs below leave the header out to save space.

## Example 1: A Single-Stage Flavor

**Context**

You keep your own version of winapp2.ini, and there's an entry you never want: here `[Arc Pinned Tabs *]`, which targets data your cleaning setup can't handle safely. Deleting it by hand after every update gets old fast. A section removal file does it in one step, every time.

###### Note: winapp2ool no longer generates this entry. We kept the example because it still shows how the stage works.

**Intent**

We want to remove `[Arc Pinned Tabs *]` from winapp2.ini.

**Files**

###### **Base file (`winapp2.ini`)**
```ini
[And Yet It Moves *]
Section=Games
Detect=HKCU\Software\Valve\Steam\Apps\18700
FileKey1=%AppData%\Broken Rules\And Yet It Moves Steam|console.log

[Arc Pinned Tabs *]
Section=Arc Web Browser
DetectFile=%LocalAppData%\Packages\TheBrowserCompany.Arc_ttt1ap7aakyb4\LocalCache\Local\Arc\User Data
RegKey1=
```

###### **Correction file (`section_removals.ini`)**
```ini
; Arc stores pinned tabs in a JSON file shared with persistent configuration
[Arc Pinned Tabs *]
```

###### Note: You don't need to include any keys when removing whole sections.

**Command**
```
winapp2ool -flavorize -3f section_removals.ini
```

**Output**

###### **Output file (`winapp2-flavorized.ini`)**
```ini
[And Yet It Moves *]
Section=Games
Detect=HKCU\Software\Valve\Steam\Apps\18700
FileKey1=%AppData%\Broken Rules\And Yet It Moves Steam|console.log
```

**Explanation**
- The base file defaults to winapp2.ini and the save target to winapp2-flavorized.ini, so `-3f` is the only argument we need
- Stage 1 removes `[Arc Pinned Tabs *]`. The other five stages have no files and change nothing
- Every section the correction file doesn't name stays as it was

**Notes**

Transmute can make a single change like this by itself (`winapp2ool -transmute -remove -bysection -2f section_removals.ini`), though it saves to `winapp2-transmuted.ini` instead. Flavorizer is worth reaching for once a second stage joins in, as in [Example 2](#example-2-a-multi-stage-flavor).

---

## Example 2: A Multi-Stage Flavor

**Context**

The base winapp2.ini puts generated browser entries under `Section=<browser name> Web Browser`. CCleaner groups browser entries under its own built-in `LangSecRef` categories instead. To adapt the Vivaldi entries for CCleaner, we remove the `Section` key from each one and add `LangSecRef=3033` (CCleaner's Vivaldi category) in its place. Transmute alone needs two commands for this ([Example 8](../EXAMPLES.md#example-8-chaining-operations-together-2)). Flavorizer needs one.

###### Note: A [global operation](../readme.md#global-operations) is the better way to do this now (see [Example 4](#example-4-global-operations-in-flavor-files)). We wrote this example before global operations existed, and kept it because it still shows two stages working together.

**Intent**

We want to remove `Section=Vivaldi Web Browser` from each Vivaldi entry and add `LangSecRef=3033`, in a single run.

**Files**

When we captured this, the base file had 22 generated Vivaldi entries. Three are shown here, and the rest follow the same pattern and get the same treatment.

###### **Base file (`browsers.ini`)**
```ini
[Vivaldi Autofill Data *]
Section=Vivaldi Web Browser
DetectFile=%LocalAppData%\Vivaldi\User Data
FileKey1=%LocalAppData%\Vivaldi\User Data\*|*Web Data
FileKey2=%LocalAppData%\Vivaldi\User Data\*\AutoFill*|*|REMOVESELF
FileKey3=%LocalAppData%\Vivaldi\User Data\AutoFill*|*|REMOVESELF

[Vivaldi Autoplay Preferences *]
Section=Vivaldi Web Browser
DetectFile=%LocalAppData%\Vivaldi\User Data
FileKey1=%LocalAppData%\Vivaldi\User Data\MEIPreload|*|REMOVESELF

[Vivaldi Bookmark Backups *]
Section=Vivaldi Web Browser
DetectFile=%LocalAppData%\Vivaldi\User Data
FileKey1=%LocalAppData%\Vivaldi\User Data\*|Bookmarks.bak

; 19 further Vivaldi entries, all carrying Section=Vivaldi Web Browser
```

###### **Correction file (`value_removals.ini`, Stage 3)**
```ini
[Vivaldi Autofill Data *]
Section=Vivaldi Web Browser

[Vivaldi Autoplay Preferences *]
Section=Vivaldi Web Browser

[Vivaldi Bookmark Backups *]
Section=Vivaldi Web Browser

; one section per remaining Vivaldi entry, each listing the same Section key to remove
```

###### **Correction file (`additions.ini`, Stage 6)**
```ini
[Vivaldi Autofill Data *]
LangSecRef=3033

[Vivaldi Autoplay Preferences *]
LangSecRef=3033

[Vivaldi Bookmark Backups *]
LangSecRef=3033

; one section per remaining Vivaldi entry, each adding the same LangSecRef
```

**Command**
```
winapp2ool -flavorize -1f browsers.ini -2f browsers-cc.ini -5f value_removals.ini -8f additions.ini
```

**Output**

###### **Output file (`browsers-cc.ini`)**
```ini
[Vivaldi Autofill Data *]
LangSecRef=3033
DetectFile=%LocalAppData%\Vivaldi\User Data
FileKey1=%LocalAppData%\Vivaldi\User Data\*|*Web Data
FileKey2=%LocalAppData%\Vivaldi\User Data\*\AutoFill*|*|REMOVESELF
FileKey3=%LocalAppData%\Vivaldi\User Data\AutoFill*|*|REMOVESELF

[Vivaldi Autoplay Preferences *]
LangSecRef=3033
DetectFile=%LocalAppData%\Vivaldi\User Data
FileKey1=%LocalAppData%\Vivaldi\User Data\MEIPreload|*|REMOVESELF

[Vivaldi Bookmark Backups *]
LangSecRef=3033
DetectFile=%LocalAppData%\Vivaldi\User Data
FileKey1=%LocalAppData%\Vivaldi\User Data\*|Bookmarks.bak

; ... 19 further Vivaldi entries, likewise now carrying LangSecRef=3033 ...
```

**Explanation**

| Stage | File | Effect |
|:-|:-|:-|
| 3 (`-5f`) | value_removals.ini | The `Section=Vivaldi Web Browser` key is removed from each listed entry |
| 6 (`-8f`) | additions.ini | `LangSecRef=3033` is added to each listed entry |

- Stage 3 runs before Stage 6, so each entry finishes the run with exactly one category key
- The output matches what the two chained commands in Transmute's Example 8 produce, from one command
- Sections the correction files don't name stay as they were

**Notes**

The winapp2.ini project's CCleaner flavor doesn't keep these per-entry lists anymore. One `[*Map:]` rule per browser does the same conversion for every entry at once, including entries that don't exist yet. See [Example 4](#example-4-global-operations-in-flavor-files).

---

## Example 3: Auto-Detecting a Flavor

**Context**

The winapp2.ini project keeps each flavor's correction files in a directory of their own, named by the auto-detection conventions with a flavor prefix added. The [CCleaner flavor directory](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/Sources/Flavors/CCleaner) has:

```
cc_additions.ini
cc_key_replacements.ini
cc_name_removals.ini
cc_section_removals.ini
cc_section_replacements.ini
cc_value_removals.ini
```

Detection is a substring match, so the `cc_` prefixes don't get in the way.

**Intent**

We want to apply the whole CCleaner flavor without assigning six file slots by hand.

**Command**
```
winapp2ool -flavorize -autodetect -2f winapp2-ccleaner-flavor.ini -9d C:\path\to\Assembler\Sources\Flavors\CCleaner
```

The build pipeline runs the same command from the Assembler directory, where `\CCleaner` resolves, and adds the global flags `-s` (silent) and `-offline` (skip the network check):

```
winapp2ool -s -offline -flavorize -autodetect -2f winapp2-ccleaner-flavor.ini -9d \Sources\Flavors\CCleaner
```

**What gets assigned**

| File found | Slot |
|:-|:-|
| cc_section_removals.ini | Section removal file (Stage 1) |
| cc_name_removals.ini | Key name removal file (Stage 2) |
| cc_value_removals.ini | Key value removal file (Stage 3) |
| cc_section_replacements.ini | Section replacement file (Stage 4) |
| cc_key_replacements.ini | Key replacement file (Stage 5) |
| cc_additions.ini | Additions file (Stage 6) |

**Explanation**
- `-9d` sets the directory auto-detection scans. Auto-detection doesn't fill the base file (slot 1) or the save target (slot 2); those come from their defaults or the command line
- With `-autodetect`, Flavorizer doesn't read slots 3–8 from the command line at all
- Here the base file is the default, `winapp2.ini` in the current directory, and the result goes to `winapp2-ccleaner-flavor.ini`

---

## Example 4: Global Operations in Flavor Files

**Context**

Correction files normally match base sections by name, so covering the 22 Vivaldi entries in Example 2 takes 22 sections in *each* correction file, and every new browser entry BrowserBuilder generates means extending those lists in step. Transmute's [global operations](../readme.md#global-operations) do away with that upkeep, and flavor files get them too. The CCleaner flavor's Stage 5 file has one `[*Map:]` rule per browser. The System Ninja flavor's Stage 5 file has one-to-many `[*Map:]` rules that turn each wildcard `DetectFile` into its hardcoded variants, because System Ninja doesn't support wildcards in detection.

**Files**

###### **Key replacement file (`cc_key_replacements.ini`, Stage 5)**
```ini
[*Map: Brave]
Match=Section=Brave Web Browser
Replace=LangSecRef=3034

[*Map: Chrome]
Match=Section=Google Chrome Web Browser
Replace=LangSecRef=3029

[*Map: Vivaldi]
Match=Section=Vivaldi Web Browser
Replace=LangSecRef=3033
```

###### **Key replacement file (`sn_key_replacements.ini`, Stage 5)**
```ini
[*Map: Brave DetectFile]
Match=DetectFile=%LocalAppData%\BraveSoftware\Brave-Browser*
Replace1=DetectFile1=%LocalAppData%\BraveSoftware\Brave-Browser
Replace2=DetectFile2=%LocalAppData%\BraveSoftware\Brave-Browser-Beta
Replace3=DetectFile3=%LocalAppData%\BraveSoftware\Brave-Browser-Nightly

[*Map: DuckDuckGo DetectFile]
Match=DetectFile=%LocalAppData%\Packages\*DuckDuckGo*Browser*
Replace1=DetectFile1=%LocalAppData%\Packages\63909DuckDuckGoInc.DuckDuckGoPrivateBrowser_qzdjx70tn762j
Replace2=DetectFile2=%LocalAppData%\Packages\DuckDuckGo.DesktopBrowser_ya2fgkz3nks94
```

**Command**
```
winapp2ool -flavorize -1f browsers.ini -2f browsers-cc.ini -7f cc_key_replacements.ini
```

**Output**

Every entry in browsers.ini whose `Section` value matches a rule has that whole key replaced, name included. Three rules give Example 2's output for all 22 Vivaldi entries, and convert every Brave and Chrome entry too, with no per-entry lists:

```ini
[Brave Autofill Data *]
LangSecRef=3034
DetectFile=%LocalAppData%\BraveSoftware\Brave-Browser\User Data
FileKey1=%LocalAppData%\BraveSoftware\Brave-Browser\User Data\*|*Web Data

[Vivaldi Autofill Data *]
LangSecRef=3033
DetectFile=%LocalAppData%\Vivaldi\User Data
FileKey1=%LocalAppData%\Vivaldi\User Data\*|*Web Data

; ... every other entry matching a rule likewise converted ...
```

**Explanation**
- Flavorizer only reads `[*Map:]` rules in the Key Replacement stage (Stage 5), the one stage that runs Transmute's Replace ByKey mode. Any other stage skips them with a warning
- The System Ninja rules are one-to-many. `Replace1` takes the matched key's place and the other `ReplaceN` keys go in right after it, so one rule swaps the wildcard for every hardcoded variant. These rules replaced a remove-and-add pair that used to be split between `sn_value_removals.ini` and `sn_additions.ini`
- A `[*]` section works in any stage whose mode supports it, applying its keys to *every* entry (eg. removing a key value everywhere in Stage 3). Stages 1 and 4, which work on whole sections, refuse `[*]` with a warning
- Global sections run before the named sections in the same file, so per-entry corrections can adjust what the global ones did
- New entries for these browsers are covered automatically, without editing any flavor file
- A rule that matches nothing in the base file warns, which usually means the rule has gone stale
