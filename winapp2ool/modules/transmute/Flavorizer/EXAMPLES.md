# Flavorizer Examples

Worked examples for [Flavorizer](readme.md). The [Common Uses](readme.md#common-uses) table in the readme lists the commands; these show what each one does, with real input and output.

# Table of Contents

1. [Example 1: A Single-Stage Flavor](#example-1-a-single-stage-flavor)
2. [Example 2: A Multi-Stage Flavor](#example-2-a-multi-stage-flavor)
3. [Example 3: Auto-Detecting a Flavor](#example-3-auto-detecting-a-flavor)
4. [Example 4: Global Operations in Flavor Files](#example-4-global-operations-in-flavor-files)

---

###### Note: When winapp2.ini formatting is enabled (the default), the output file begins with the regenerated standard winapp2.ini preamble comments. The example outputs below omit this preamble for brevity.

## Example 1: A Single-Stage Flavor

**Context**

You maintain a local variant of winapp2.ini and there is an entry you never want, in this case `[Arc Pinned Tabs *]`, which targets data your cleaning setup cannot handle safely. Deleting it by hand after every update doesn't scale; a section removal file does it in one step, every time.

###### Note: This entry is no longer generated in the current version of winapp2ool, however this example is left here for the purpose of this readme. 

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

###### Note: There is no need to include keys when removing entire sections.

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
- The base file defaults to winapp2.ini and the save target to winapp2-flavorized.ini; only `-3f` needed to be provided
- Stage 1 removes `[Arc Pinned Tabs *]`; the other five stages have no files and make no changes
- Sections in the base file not named in the correction file remain unchanged

**Notes**

A single transformation like this can equally be done with Transmute directly (`winapp2ool -transmute -remove -bysection -2f section_removals.ini`). Flavorizer becomes more useful as more stages join in. see Example 2.

---

## Example 2: A Multi-Stage Flavor

**Context**

The base winapp2.ini categorizes generated browser entries with `Section=<browser name> Web Browser`. CCleaner instead groups browser entries under its built-in `LangSecRef` categories. Adapting the Vivaldi entries for CCleaner means removing the `Section` key from every Vivaldi entry and adding `LangSecRef=3033` (Vivaldi's CCleaner category) in its place. In the Transmute readme this takes two chained invocations ([Example 8](../EXAMPLES.md#example-8-chaining-operations-together-2)); Flavorizer does it in one.

###### This is actually best achieved with a [global operation](../readme.md#global-operations), but this example was written before global operations were implemented. It is left here for the purposes of the readme. 

**Intent**

We want to remove `Section=Vivaldi Web Browser` from each Vivaldi entry and add `LangSecRef=3033`, in a single run.

**Files**

There are 22 generated Vivaldi entries in the base file; three are shown here. The remaining entries follow the same pattern and receive the same treatment.

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

- Removals run before additions (Stage 3 before Stage 6), so every entry ends the run with exactly one categorization key
- This produces the same output as the two chained Transmute commands in Transmute's Example 8, in a single invocation
- Sections in the base file not named in the correction files remain unchanged

**Notes**

The winapp2.ini project's actual CCleaner flavor no longer maintains these per-entry cohorts: a single `[*Map:]` rule per browser achieves the same conversion for every entry at once, including entries that don't exist yet. See Example 4.

---

## Example 3: Auto-Detecting a Flavor

**Context**

The winapp2.ini project stores each flavor's correction files in a dedicated directory, named with the auto-detection conventions plus a flavor prefix. The [CCleaner flavor directory](https://github.com/MoscaDotTo/Winapp2/tree/master/Assembler/CCleaner) contains:

```
cc_additions.ini
cc_key_replacements.ini
cc_name_removals.ini
cc_section_removals.ini
cc_section_replacements.ini
cc_value_removals.ini
```

Because detection is a substring match, the `cc_` prefixes don't interfere.

**Intent**

We want to apply the complete CCleaner flavor without assigning six file slots by hand.

**Command**
```
winapp2ool -flavorize -autodetect -2f winapp2-ccleaner-flavor.ini -9d C:\path\to\Assembler\CCleaner
```

This mirrors the build pipeline's real invocation, which runs from the Assembler directory and additionally passes the global flags `-s` (silent) and `-offline` (skip the network check):

```
winapp2ool -s -offline -flavorize -autodetect -2f winapp2-ccleaner-flavor.ini -9d \CCleaner
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
- `-9d` sets the directory that auto-detection scans; the base file (slot 1) and save target (slot 2) are not auto-detected and come from their defaults or the command line
- With `-autodetect`, slots 3–8 are not read from the command line at all
- The base file here is the default `winapp2.ini` in the current directory; the flavorized result is saved to `winapp2-ccleaner-flavor.ini`

---

## Example 4: Global Operations in Flavor Files

**Context**

Correction files normally match base sections by name, so covering the 22 Vivaldi entries in Example 2 takes 22 sections in *each* correction file — and every new browser entry generated by BrowserBuilder needs the cohorts extended in lockstep. Transmute's [global operations](../readme.md#global-operations) remove that maintenance burden, and flavor files inherit them: the CCleaner flavor's Stage 5 file carries one `[*Map:]` rule per browser, and the System Ninja flavor's Stage 5 file carries one-to-many `[*Map:]` rules that de-abstract each wildcard `DetectFile` into its hardcoded variants (System Ninja does not support wildcards in detection).

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

Every entry in browsers.ini whose `Section` value matches a rule has the whole key replaced, name included. The Example 2 output for all 22 Vivaldi entries, plus every Brave and Chrome entry, from three rules and no per-entry cohorts:

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
- `[*Map:]` rules are only recognized in the Key Replacement stage (Stage 5), because it is the only stage running Transmute's Replace ByKey mode. In any other stage's file they are skipped with a warning
- The System Ninja rules above are one-to-many: `Replace1` takes the matched key's position and the remaining `ReplaceN` keys are inserted immediately after it, replacing the wildcard with every hardcoded variant in one operation. These rules retired a remove + add lockstep formerly split between `sn_value_removals.ini` and `sn_additions.ini`
- A `[*]` section works in any stage whose underlying mode supports it, applying its keys to *every* entry under that stage's mode (eg. removing a key value everywhere during Stage 3). Stages 1 and 4 (the BySection stages) refuse `[*]` with a warning
- Global sections are processed before the named sections in the same file, so per-entry corrections can refine the result of global ones
- New entries generated for these browsers are covered automatically, no flavor file edits needed
- A rule that matches nothing in the base file emits a warning; that usually means the rule has gone stale
