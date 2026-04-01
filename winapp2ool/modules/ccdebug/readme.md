# CCiniDebug

**CCiniDebug** is a winapp2ool module that performs housekeeping on CCleaner Classic's configuration file, `ccleaner.ini`. When winapp2.ini is used with CCleaner 1-6 (classic), CCleaner stores the enabled/disabled state of each winapp2.ini entry as a key in `ccleaner.ini`'s `[Options]` section. As entries are renamed or removed from winapp2.ini over time, their corresponding configuration keys are left behind as orphans in this section. CCiniDebug finds and removes these orphans, and can also sort the keys in the section alphabetically.

### Should I use CCiniDebug with CCleaner 7?
No. Although CCleaner 7 has a `ccleaner.ini` file, it holds cleaning definitions instead of CCleaner's configuration. Running CCiniDebug on CCleaner 7's `ccleaner.ini` will, at best, do nothing. 

### What does CCiniDebug do?

CCiniDebug can perform up to three operations on `ccleaner.ini`, each independently toggleable:

- **Prune**: Scan the `[Options]` section for winapp2.ini configuration keys that no longer have a matching entry in winapp2.ini and remove them
- **Sort**: Sort the keys in the `[Options]` section alphabetically
- **Save**: Write the result back to disk

---

# Table of Contents

1. [Requirements](#requirements)
2. [Quick Start](#quick-start)
3. [Menu Options](#menu-options)
4. [Operations](#operations)
   - [Pruning](#pruning)
   - [Sorting](#sorting)
   - [Saving](#saving)
5. [Command-Line Arguments](#command-line-arguments)
   - [Toggles](#toggles)
   - [File Selection](#file-selection)
   - [Examples](#examples)
6. [Tips & Best Practices](#tips--best-practices)
7. [Troubleshooting](#troubleshooting)

---

# [Requirements](#requirements)

- A `ccleaner.ini` file to be debugged 
- A `winapp2.ini` file to check against *(only required when Pruning is enabled)*

---

# [Quick Start](#quick-start)

### Common Workflow

1. Locate your `ccleaner.ini` (typically in `%ProgramFiles%\CCleaner\`) and your current `winapp2.ini`
2. Select both files in CCiniDebug
3. Run — pruning, sorting, and saving are all enabled by default

---

# [Menu Options](#menu-options)

| Option | Effect | Notes |
|:-|:-|:-|
| Run (default) | Execute the enabled operations | Requires at least one operation to be enabled |
| Toggle pruning | Enable or disable removal of orphaned winapp2.ini settings | Default: `True` |
| Toggle saving | Enable or disable writing the result to disk | Default: `True` |
| Toggle sorting | Enable or disable alphabetical sorting of `[Options]` | Default: `True` |
| Choose winapp2.ini | Select the reference winapp2.ini file | Only visible when Pruning is enabled |
| Choose ccleaner.ini | Select the ccleaner.ini file to debug | |
| Choose save target | Select the output location | Only visible when Saving is enabled |

At least one operation must be enabled before Run is available.

---

# [Operations](#operations)

## [Pruning](#pruning)

Scans the `[Options]` section of ccleaner.ini for stale winapp2.ini entry settings and removes them.

### How CCleaner stores winapp2.ini settings

When CCleaner 6 loads winapp2.ini, it creates a key in the `[Options]` section for each entry, recording whether it is enabled or disabled:

```ini
[Options]
(App)Some Application *=True
(App)Another Application *=False
```

When entries are later renamed or removed from winapp2.ini, these keys remain in ccleaner.ini indefinitely. Over time, a large number of orphaned keys can accumulate.

### How pruning works

CCiniDebug scans every key in the `[Options]` section and identifies candidates by checking that the key:

1. Starts with `(App)`
2. Contains `*` (the winapp2.ini entry name indicator)

For each candidate, it strips `(App)` from the key name to recover the raw entry name, then checks whether a section of that name exists in the provided winapp2.ini. If no matching section is found, the key is flagged as orphaned and removed. The lookup is case-insensitive.

### What pruning reports

- Each orphaned key found is listed by name
- A count of total orphaned settings removed is shown at the end

### Requirement

Pruning requires a winapp2.ini file to compare against. The winapp2.ini selector in the menu is only available when pruning is enabled.

---

## [Sorting](#sorting)

Sorts the keys within the `[Options]` section of ccleaner.ini alphabetically. Sorting is applied after pruning (if both are enabled), so the output reflects the pruned state.

---

## [Saving](#saving)

Writes the result of all enabled operations back to disk. When saving is disabled, CCiniDebug still performs pruning and sorting in memory and reports what it found, but makes no changes to any file.

The save target defaults to `ccleaner-debugged.ini` in the current directory. Change it with **Choose save target** if you want to overwrite your original file directly or save to a different location.

---

# [Command-Line Arguments](#command-line-arguments)

CCiniDebug supports command-line automation. All three operations are enabled by default; the command-line flags disable them selectively.

### [Toggles](#toggles)

| Arg | Effect |
|:-|:-|
| `-noprune` | Disable pruning of stale winapp2.ini entries |
| `-nosort` | Disable alphabetical sorting of `[Options]` |
| `-nosave` | Disable saving the result to disk |

### [File Selection](#file-selection)

| Arg | Effect | Default |
|:-|:-|:-|
| `-1d path` | Set winapp2.ini directory | Current directory |
| `-1f name` | Set winapp2.ini file name | `winapp2.ini` |
| `-2d path` | Set ccleaner.ini directory | Current directory |
| `-2f name` | Set ccleaner.ini file name | `ccleaner.ini` |
| `-3d path` | Set save target directory | Current directory |
| `-3f name` | Set save target file name | `ccleaner-debugged.ini` |

### [Examples](#examples)

| Command | Effect |
|:-|:-|
| `winapp2ool -ccinidebug` | Prune, sort, and save using default file names in the current directory |
| `winapp2ool -ccinidebug -nosave` | Prune and sort, report results but make no changes to disk |
| `winapp2ool -ccinidebug -noprune -nosort` | Save only — write ccleaner.ini to the save target with no modifications |
| `winapp2ool -ccinidebug -2d "%ProgramFiles%\CCleaner" -3d "%ProgramFiles%\CCleaner" -3f ccleaner.ini` | Debug ccleaner.ini in place from its default CCleaner location |

---

# [Tips & Best Practices](#tips--best-practices)

### Keep winapp2.ini Current

Always use the most recent version of winapp2.ini when pruning.

### Sorting Without Pruning

You can run CCiniDebug with pruning disabled (`-noprune`) to sort ccleaner.ini alphabetically without making any entry-level changes.

### Dry Run

Disable saving (`-nosave` or toggle it off in the menu) to see exactly which entries would be pruned without modifying any file. The output will still list every orphaned key detected.

---

# [Troubleshooting](#troubleshooting)

| Symptom | Cause |
|:-|:-|
| Run is red and prompts an error | All three operations are disabled — enable at least one |
| No orphaned entries detected despite expecting some | Your winapp2.ini may be outdated; the entries may still exist in it. Use the current version. |
| A valid entry was pruned | Its name in ccleaner.ini may not exactly match the section name in winapp2.ini (e.g., extra spaces or punctuation differences) |
| The save target is not being updated | Saving is disabled — toggle it on or check the save target path |
