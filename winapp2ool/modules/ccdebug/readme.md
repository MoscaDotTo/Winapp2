# CCiniDebug

**CCiniDebug** is a winapp2ool module for end users that tidies up CCleaner Classic's configuration file, `ccleaner.ini`. When you use winapp2.ini with CCleaner 1-6 (Classic), CCleaner records whether each winapp2.ini entry is enabled as a key in the `[Options]` section of `ccleaner.ini`. Entries get renamed or removed from winapp2.ini over time, but their keys stay behind. CCiniDebug finds these orphaned keys and removes them, and it can sort the section alphabetically too.

### Should I use CCiniDebug with CCleaner 7?
No. CCleaner 7 has a `ccleaner.ini` too, but it holds cleaning definitions, not CCleaner's settings. Running CCiniDebug on it will, at best, do nothing. To get winapp2.ini into CCleaner 7, use [CC7Patcher](../cc7patcher/readme.md).

### What does CCiniDebug do?

CCiniDebug does up to three things to `ccleaner.ini`, and you can turn each one on or off:

- **Prune**: remove the `[Options]` keys whose winapp2.ini entry no longer exists
- **Sort**: sort the keys in `[Options]` alphabetically
- **Save**: write the result to the save target

---

# Table of Contents

**Using CCiniDebug**

1. [Requirements](#requirements)
2. [Quick Start](#quick-start)
3. [Common Uses](#common-uses)
4. [Menu Options](#menu-options)
5. [Command-Line Arguments](#command-line-arguments)
   - [Toggles](#toggles)
   - [File Selection](#file-selection)
6. [Operations](#operations)
   - [Pruning](#pruning)
   - [Sorting](#sorting)
   - [Saving](#saving)
7. [Tips & Best Practices](#tips--best-practices)
8. [Troubleshooting](#troubleshooting)

---

# Requirements

- The `ccleaner.ini` you want to debug
- A `winapp2.ini` to check it against *(only needed when pruning is on)*

---

# Quick Start

### Common Workflow

1. Find your `ccleaner.ini` (usually in `%ProgramFiles%\CCleaner\`) and a current `winapp2.ini`
2. Use **Choose ccleaner.ini** and **Choose winapp2.ini** to select them
3. Select **Run**. Pruning, sorting and saving are all on by default, and the result is saved as `ccleaner.ini` in the current directory. To keep the original, use **Choose save target** first

---

# Common Uses

| To | Run |
|:-|:-|
| Prune, sort, and save using the default file names in the current directory | `winapp2ool -ccdebug` |
| See which keys would be pruned without changing any file | `winapp2ool -ccdebug -nosave` |
| Write `ccleaner.ini` to the save target with no changes | `winapp2ool -ccdebug -noprune -nosort` |
| Keep the original and save the result as `ccleaner-debugged.ini` | `winapp2ool -ccdebug -3f ccleaner-debugged.ini` |
| Clean up `ccleaner.ini` in place in its default CCleaner location | `winapp2ool -ccdebug -2d "%ProgramFiles%\CCleaner" -3d "%ProgramFiles%\CCleaner"` |

Program Files isn't writable without admin rights, so the last command brings up a UAC prompt if necessary before winapp2ool saves.

In the menu, these are toggles and file pickers you set before **Run**. See [Menu Options](#menu-options).

###### Note: Winapp2ool does not expand environment variables in path arguments. `%ProgramFiles%` works above because the Command Prompt (cmd) expands it before winapp2ool sees it. From PowerShell, use `$env:ProgramFiles` instead.

---

# Menu Options

| Option | Effect | Notes |
|:-|:-|:-|
| Run (default) | Run the enabled operations | Turns red and won't run while every operation is off |
| Toggle pruning | Turn removal of orphaned winapp2.ini settings on or off | Default: `On` |
| Toggle saving | Turn writing the result to disk on or off | Default: `On` |
| Toggle sorting | Turn alphabetical sorting of `[Options]` on or off | Default: `On` |
| Choose winapp2.ini | Select the winapp2.ini to check against | Only shown while pruning is on |
| Choose ccleaner.ini | Select the ccleaner.ini to debug | |
| Choose save target | Select where the result is saved | Only shown while saving is on |
| Reset Settings | Restore all settings to their defaults | Only shown when settings have been changed |

---

# Command-Line Arguments

On the command line, CCiniDebug is `winapp2ool -ccdebug` (`ccdebug`, `5`, and `-5` also work). All three operations are on by default, and each flag below turns one off.

### Toggles

| Arg | Effect |
|:-|:-|
| `-noprune` | Don't prune orphaned winapp2.ini settings |
| `-nosort` | Don't sort `[Options]` |
| `-nosave` | Don't save the result |

### File Selection

| Arg | Effect | Default |
|:-|:-|:-|
| `-1d path` | Set winapp2.ini directory | Current directory |
| `-1f name` | Set winapp2.ini file name | `winapp2.ini` |
| `-2d path` | Set ccleaner.ini directory | Current directory |
| `-2f name` | Set ccleaner.ini file name | `ccleaner.ini` |
| `-3d path` | Set save target directory | Current directory |
| `-3f name` | Set save target file name | `ccleaner.ini` |

---

# Operations

## Pruning

Removes the `[Options]` keys left behind by winapp2.ini entries that no longer exist.

### How CCleaner stores winapp2.ini settings

When CCleaner Classic loads winapp2.ini, it adds a key to `[Options]` for each entry, recording whether that entry is enabled:

```ini
[Options]
(App)Some Application *=True
(App)Another Application *=False
```

Rename or remove an entry in winapp2.ini and its key stays in ccleaner.ini forever. After enough winapp2.ini updates, they pile up.

### How pruning works

CCiniDebug looks at every key in `[Options]` and treats it as a winapp2.ini setting if its name:

1. Starts with `(App)`
2. Contains `*`, which marks a winapp2.ini entry name

It drops the `(App)` prefix and any surrounding spaces to get the entry name, then looks for a section with that name in your winapp2.ini, ignoring case. If there isn't one, the key is an orphan and CCiniDebug removes it.

### What pruning reports

- An `Orphaned: <entry name>` line for each orphan
- The number of orphaned settings at the end, as "removed" once the save succeeds and "found" otherwise

### Requirement

Pruning needs a winapp2.ini to compare against, so the menu only shows **Choose winapp2.ini** while pruning is on.

---

## Sorting

Sorts the keys in the `[Options]` section of ccleaner.ini alphabetically, ignoring case. When pruning is on too, sorting runs after it, so you get the pruned section sorted.

---

## Saving

Writes the result of the enabled operations to the save target. With saving off, CCiniDebug still prunes and sorts in memory and reports what it found, but doesn't change any file.

The save target defaults to `ccleaner.ini` in the current directory, so a run on the `ccleaner.ini` there overwrites it. To keep the original, use **Choose save target**, which offers `ccleaner-debugged.ini` as a one-step rename, or pass `-3f`.

---

# Tips & Best Practices

### Keep winapp2.ini Current

Prune against the winapp2.ini CCleaner actually loads, which should be the newest one. CCiniDebug calls a key orphaned whenever its entry is missing from the file you give it, so pruning against a different version gets some keys wrong.

### Sorting Without Pruning

Turn pruning off (`-noprune`) to sort ccleaner.ini without removing anything.

### Dry Run

Turn saving off (`-nosave`, or **Toggle saving** in the menu) to see which keys would be pruned without touching any file. CCiniDebug still lists every orphan it finds, and the count at the end says how many it found.

---

# Troubleshooting

| Symptom | Cause |
|:-|:-|
| Run is red and says "Please enable at least one option" | All three operations are off. Turn at least one on |
| No orphans found when you expected some | Your winapp2.ini may be out of date and still have those entries. Use the current version |
| A valid entry was pruned | Its name in ccleaner.ini differs from the section name in winapp2.ini (punctuation, or a doubled space inside the name; case doesn't matter), or you pruned against a different winapp2.ini than the one CCleaner loads |
| The save target isn't being updated | Saving is off, or the save target points somewhere else. Turn saving on and check the path |
