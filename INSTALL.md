# Installing Winapp2.ini

This page covers installing and configuring winapp2.ini in each application that supports it. For what winapp2.ini is and how it's built, see the [main readme](README.md).

Whichever application you use, install winapp2.ini with [winapp2ool.exe](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2ool.exe) and keep a copy of it in the same folder as winapp2.ini. That's how you keep the file up to date afterwards.

The download links on this page fetch each file from the [latest release](https://github.com/MoscaDotTo/Winapp2/releases/latest), where most file names carry their flavor, such as `winapp2-ccleaner.ini`. If you install one by hand, rename it to `winapp2.ini` first, or to `winapp2.rules` for System Ninja.

---

## Installing with winapp2ool

Winapp2ool is the recommended way to install winapp2.ini for every application on this page, and the only one that also keeps it current afterwards. The procedure is the same everywhere:

1. Download [winapp2ool.exe](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2ool.exe) and put it in the folder where your application looks for winapp2.ini. Each section below names that folder.
2. Run it, open **Settings**, and press **Change Flavor** until it shows the flavor your application's section names. It cycles through CCleaner, NonCCleaner, BleachBit, SystemNinja, Tron, CCleaner7, and FluentCleaner, and the default is CCleaner. To keep your choice between sessions, turn on **Toggle Saving Settings** and **Toggle Reading settings**.
3. Go back to the main menu and choose **Update Winapp2.ini** to install, or **Update & Trim** to install a copy without the entries that don't apply to your machine.

Both options write into the folder winapp2ool is running from, replacing the copy that's already there. They save `winapp2.ini`, or `winapp2.rules` for System Ninja. They're listed only while winapp2ool sees your local copy as out of date, which includes having no copy at all, so they'll be waiting for you on a first install. Neither is available in offline mode. If the folder is protected, such as one under Program Files, winapp2ool asks for administrator rights when it starts.

CCleaner 7 has no standalone winapp2.ini, so it's installed with [CC7Patcher](#ccleaner-7) instead.

### From the command line

Every install below can be scripted. Winapp2ool expects to live in the folder where winapp2.ini belongs, so if that's where you put it, leave the paths off and it uses the folder it's running from:

```
winapp2ool -download winapp2
```

That fetches whichever flavor winapp2ool is set to and writes `winapp2.ini` (`winapp2.rules` for System Ninja) beside the exe. A flavor flag overrides the setting for one run without changing it: `-cc`, `-base`, `-bb`, `-sn`, `-tron`, `-cc7`, `-fc`.

To download and trim in one step, writing the result straight to `winapp2.ini`:

```
winapp2ool -trim -d -3f winapp2.ini
```

Trim writes over `winapp2.ini` by default, so `-3f winapp2.ini` only spells that out. To keep your existing file, for instance to compare the two, pass `-3f winapp2-trimmed.ini` instead. Without `-d`, Trim trims the local `winapp2.ini` you already have instead of fetching a fresh one. It reads the whole file before writing it back, so trimming in place is safe.

Add `-s` to either command for unattended runs. It suppresses the output, and the keypress winapp2ool otherwise waits for at the end.

The defaults follow the working directory, not the exe's location. If you call winapp2ool by its full path from somewhere else, the file lands wherever you called it from. In a scheduled task, set the "Start in" field to the target folder. Folders you picked in the menu don't carry over, because every module resets its own settings for a command-line run. The flavor setting does carry over.

The sections below give the same commands with explicit paths, for installing without keeping the exe beside winapp2.ini. Those paths are written for `cmd`. In PowerShell, write `$env:ProgramFiles` in place of `%ProgramFiles%`.

---

## Contents

1. [CCleaner Classic](#ccleaner-classic)
2. [CCleaner 7](#ccleaner-7)
3. [BleachBit](#bleachbit)
4. [System Ninja](#system-ninja)
5. [Avira System Speedup](#avira-system-speedup)
6. [Tron](#tron)
7. [R-Wipe & Clean](#r-wipe--clean)
8. [HDCleaner](#hdcleaner)
9. [FluentCleaner](#fluentcleaner)
10. [Reg Organizer](#reg-organizer)

Each application has a preferred [flavor](README.md#what-are-flavors) of winapp2.ini. The table in [Files of interest](README.md#files-of-interest) links every published flavor and its changelog.

---

## CCleaner Classic

###### [Download CCleaner](https://www.filepuma.com/download/ccleaner_6.23.11010-38881/)

### Note: The instructions below apply to CCleaner 6.39 and earlier. CCleaner 7 dropped support for loading a standalone winapp2.ini and is installed differently, see [CCleaner 7](#ccleaner-7)

### Flavor

You should use the [CCleaner flavor](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2-ccleaner.ini), which fits into CCleaner's UI better and overlaps less with its own rules. The base ("Non-CCleaner") [winapp2.ini](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2.ini) works too.

### Installation

Place winapp2.ini in the same folder as ccleaner.exe. By default this is `..\Program Files\CCleaner`

We recommend trimming winapp2.ini whenever you update it, because the full file slows down CCleaner's startup for no benefit. With winapp2ool in that folder and the flavor left at its CCleaner default, **Update & Trim** is the main menu option you want.

The same install from the command line:

```
winapp2ool -trim -d -cc -3d "%ProgramFiles%\CCleaner" -3f winapp2.ini
```

For the untrimmed file:

```
winapp2ool -download winapp2 -cc -1d "%ProgramFiles%\CCleaner"
```

### Configuration

CCleaner shows the winapp2.ini entries it detects on your system in its Applications tab, which modern versions keep under Custom Clean. Every winapp2.ini entry starts out disabled, and you enable them one at a time or in groups. To enable a whole group, right click its section header and select "Check all."

###### Note: CCleaner 5.64.7577 is the last version to work on Windows XP and Vista (for non-SSE2 CPUs CCleaner 5.26.5937). Winapp2.ini and Winapp3.ini will continue to work with this version.

## CCleaner 7

###### [Download CCleaner](https://www.ccleaner.com/ccleaner)

CCleaner 7 no longer loads a separate winapp2.ini. Its cleaning definitions live inside `ccleaner.ini` beside CCleaner's own, in a modified entry format. You can't just drop a file next to `ccleaner.exe`: the entries have to be patched into `ccleaner.ini`. Winapp2ool's CC7Patcher does that, and it's the only installation method we endorse.

### Flavor

You should use the [CCleaner 7 flavor](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2-ccleaner7.ini), which CC7Patcher downloads for you by default, so you don't need to download or place it yourself. If you supply your own local copy instead, it must be in the CCleaner 7 format. No other flavor produces working entries.

### Installation

Back up your `ccleaner.ini` before your first run. Patching rewrites the file in place, which reorders its sections and strips its comments.

1. Download [winapp2ool.exe](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2ool.exe) and run it
2. Select **CC7Patcher** from the main menu
3. Use **Change ccleaner.ini** to point at CCleaner 7's `ccleaner.ini`, by default this is `..\Program Files\Piriform\CCleaner 7`
4. Optionally turn on **Toggle Trim** to install only the entries that apply to your system, which shortens CCleaner's startup
5. Select **Run (default)**

The same install from the command line:

```
winapp2ool -cc7patcher -2d "%ProgramFiles%\Piriform\CCleaner 7" -3d "%ProgramFiles%\Piriform\CCleaner 7"
```

Add `-trim` to trim before patching. CC7Patcher has [its own readme](winapp2ool/modules/cc7patcher/readme.md).

### Updating

Run CC7Patcher again over your existing `ccleaner.ini`. It removes the entries left by the previous patch before installing the current ones, so nothing is duplicated and entries removed from winapp2.ini are cleared out. There is no need to restore a clean `ccleaner.ini` first.

You also need to run it again after every CCleaner 7 update, because the update overwrites `ccleaner.ini` and takes the winapp2.ini entries with it.

### Configuration

CCleaner 7 shows the winapp2.ini entries it detects on your system beside its own cleaning options in Custom Clean. Every winapp2.ini entry starts out disabled, and you enable them one at a time or in groups.

CCleaner 7 puts a `custom-clean.name.` prefix on every winapp2.ini entry, and there's no way to turn it off. The entry list can't be widened either, so hover over an entry to see its full name. Entries are sorted alphabetically, *ignoring* the prefix.

###### Note: CC7Patcher identifies the entries it installed by their `Author=Winapp2.ini Project` key, and removes them on the next run. If you customize a winapp2.ini entry inside `ccleaner.ini`, delete that key from your copy so your changes survive updating.

## BleachBit

###### [Download BleachBit](https://www.bleachbit.org)

### Flavor

**BleachBit 5 or earlier:** You should use the [BleachBit flavor](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2-bleachbit.ini), which removes the registry exclusions BleachBit didn't support before version 6. Any other flavor works, except that BleachBit reports errors for the few entries that use those exclusions and won't run them.

**BleachBit 6 or newer:** You should use the base [winapp2.ini](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2.ini)

### Installation

1. Ensure that you have disabled "Download and update cleaners from community (Winapp2.ini)" in the BleachBit settings.
2. Place winapp2.ini in `%AppData%\BleachBit\Cleaners`.

With winapp2ool in `%AppData%\BleachBit\Cleaners`, set the flavor to NonCCleaner (or BleachBit on version 5 and older) and choose **Update Winapp2.ini**.

The same install from the command line:

```
winapp2ool -download winapp2 -base -1d "%AppData%\BleachBit\Cleaners"
```

On BleachBit 5 or older, use `-bb` instead of `-base`.

BleachBit also maintains its own [customized version of winapp2.ini](https://github.com/bleachbit/winapp2.ini), which you can turn on from inside the application:
1. Open BleachBit.
2. Select the "Edit" tab, and then "Preferences".
3. Check the box that reads "Download and update cleaners from community (Winapp2.ini)".

###### Note: Using BleachBit's winapp2.ini will overwrite any existing local file.

### Configuration 

BleachBit shows the winapp2.ini entries that have valid syntax and are detected on your system in its sidebar. Every winapp2.ini entry starts out disabled, and you enable them one at a time or in groups. To enable a whole group, select the check box next to its section header.

###### Note: BleachBit 2.2 is the last version to work on Windows XP. Winapp2.ini and Winapp3.ini will continue to work with this version.

## System Ninja

###### [Download System Ninja](https://singularlabs.com/software/system-ninja)

### Flavor
You should use the [System Ninja flavor](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2-systemninja.rules), which replaces keys System Ninja can't run with ones it can. We don't recommend any other flavor with System Ninja.

### Installation
System Ninja ships with its own copy of winapp2.ini, served from its own servers, and stores it in your `..\System Ninja\scripts\` folder as `winapp2.rules`.

To keep that `winapp2.rules` up to date with winapp2ool instead:
1. Open System Ninja
2. Select the "Options" tab
3. *Uncheck* the box that reads "Update cleaning rules and language files automatically"

That stops System Ninja from overwriting your `winapp2.rules` when it starts.

Put winapp2ool in the `scripts` folder, set the flavor to SystemNinja, and choose **Update Winapp2.ini**. Winapp2ool updates the `winapp2.rules` it finds there. The same install from the command line:

```
winapp2ool -download winapp2 -sn -1d "C:\Path\To\System Ninja\scripts"
```

### Configuration

System Ninja doesn't let you choose which winapp2.ini entries run. It runs all of them and shows the results in the Junk Scanner window, where the "Type" column holds each entry's name as it appears in winapp2.ini.

###### Note: System Ninja 3.2.7 is the last version to work on Windows XP and Vista. Winapp2.ini and Winapp3.ini will continue to work with this version.

## Avira System Speedup

###### [Download System Speedup](https://www.avira.com/en/avira-system-speedup-free)

### Flavor

You should use the base [winapp2.ini](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2.ini)

### Installation
Avira System Speedup ships with its own copy of winapp2.ini, served by Avira, and stores it in your `..\Avira\System Speedup\sdf` folder. You can replace or update that copy without changing any of Avira System Speedup's settings.

Put winapp2ool in that folder, set the flavor to NonCCleaner, and choose **Update Winapp2.ini**. The same install from the command line:

```
winapp2ool -download winapp2 -base -1d "C:\Path\To\Avira\System Speedup\sdf"
```

### Configuration

Avira System Speedup scans every winapp2.ini entry and shows the results in a panel labeled Third Party Applications. You can turn items on or off there before you clean.

###### Note: Cleaning "Third Party Applications" is a paid feature of Avira System Speedup Pro. Winapp2.ini is and always will be free, and plenty of free applications support it.

## Tron

###### [Tron GitHub](https://github.com/bmrf/tron)

### Flavor

You should use the [Tron flavor](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2-tron.ini), which carries the changes Tron makes to winapp2.ini and is otherwise as current as the rest.

### Installation

Tron maintains and ships its own [customized version of winapp2.ini](https://raw.githubusercontent.com/bmrf/tron/refs/heads/master/resources/stage_1_tempclean/ccleaner/winapp2.ini), stored in your `..\resources\stage_1_tempclean\ccleaner` folder. Overwrite that file to update it.

Put winapp2ool in that folder, set the flavor to Tron, and choose **Update Winapp2.ini**. The same install from the command line:

```
winapp2ool -download winapp2 -tron -1d "C:\Path\To\tron\resources\stage_1_tempclean\ccleaner"
```

### Configuration

Tron ships with its own configuration. You can modify it by opening the copy of CCleaner or modifying the `ccleaner.ini` file shipped with Tron.

## R-Wipe & Clean

###### [Download R-Wipe&Clean](https://www.r-wipe.com)

### Flavor

You should use the base [winapp2.ini](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2.ini)

### Installation

R-Wipe & Clean supports winapp2.ini unofficially.

R-Wipe & Clean imports winapp2.ini once through a prompt, instead of reading it from a fixed location when it starts. Put winapp2ool wherever you find convenient, set the flavor to NonCCleaner, and choose **Update Winapp2.ini** to fetch the file you are about to import. The same from the command line:

```
winapp2ool -download winapp2 -base -1d "C:\Path\To\Your\Winapp2 Folder"
```

The import steps below are adapted from [this thread](https://forum.r-tt.com/viewtopic.php?t=11018)

1. In R-Wipe & Clean, select "Tools" from the menu bar
2. From the Tools menu, click R-Wipe&Clean Smart
3. In R-Wipe & Clean Smart, select "Settings" from the menu bar
4. From the settings menu, select "INI Import Settings"
5. In the INI Import Settings window, fill out the following:
   *  Key Name for Registry Key Detection: `Detect`
   * Key Name for File/Folder Detection: `DetectFile`
   * Parameter for Recursive Subfolder Cleaning: `RECURSE`
   * Parameter for Removing Folder after Cleaning: `REMOVESELF`
6. Press OK
7. In R-Wipe & Clean Smart, select "Advanced" from the menu bar
8. In the Advanced menu, select "Import From .INI"
9. Browse to and select winapp2.ini
10. Import as either many wipe lists or as a single one, whichever is your preference
    * You'll repeat the import every time you update, so we recommend a single wipe list, which is easy to remove before the next import
11. Once winapp2.ini is imported, you should see a message saying `LangSecRef`, `Section`, and `Warning` are unsupported functions, press OK

### Configuration

Configure the imported wipe lists the same way as R-Wipe & Clean's own.

## HDCleaner

###### [Download HDCleaner](https://kurtzimmermann.com/index_en.html)

### Flavor

You should use the base [winapp2.ini](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2.ini)

### Installation

HDCleaner ships with a built-in copy of winapp2.ini. You can override or add to it by placing winapp2.ini in the same folder as `HDCleaner.exe`. By default, that's `..\Program Files\HDCleaner`.

Put winapp2ool in that folder, set the flavor to NonCCleaner, and choose **Update Winapp2.ini**. The same install from the command line:

```
winapp2ool -download winapp2 -base -1d "%ProgramFiles%\HDCleaner"
```

### Configuration

Once winapp2.ini is beside `HDCleaner.exe`, tell HDCleaner how to handle rules that conflict with its own:

1. Open HDCleaner
2. From the side bar, select the bottom most icon (settings)
3. From the settings panel, select "Options"
4. Change the setting "If using Winapp2.ini" to "Ignore duplicate entries in HDCleaner resource file"

Your winapp2.ini now takes priority over HDCleaner's built-in entries.

## FluentCleaner

###### [Download FluentCleaner](https://github.com/builtbybel/FluentCleaner/releases)

### Flavor

You should use the [FluentCleaner flavor](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2-fluentcleaner.ini), which carries the changes FluentCleaner makes to its own winapp2.ini.

### Installation

FluentCleaner ships with a copy of winapp2.ini in the same folder as `FCleaner.exe` or `FluentCleaner.Classic.exe`. You can update winapp2.ini by replacing this file yourself.

Put winapp2ool in that folder, set the flavor to FluentCleaner, and choose **Update Winapp2.ini**. The same install from the command line:

```
winapp2ool -download winapp2 -fc -1d "C:\Path\To\FluentCleaner"
```

### Configuration

With FluentCleaner's own winapp2.ini or the FluentCleaner flavor, almost every entry starts out enabled. Clear an entry's checkbox to disable it.

## Reg Organizer

###### [Download Reg Organizer](https://www.chemtable.com/organizer.htm)

###### ChemTable also provides their own information about winapp2.ini [here](https://www.chemtable.com/blog/en/community-driven-cleanup.htm)

### Flavor

You should use the base [winapp2.ini](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2.ini)

### Installation

Reg Organizer ships with a copy of winapp2.ini maintained by ChemTable, in `..\AppData\Roaming\ChemTable Software\Reg Organizer\CleanupDatabases` by default. Replace it with the one from this repository to update it yourself. Reg Organizer may put its own copy back when it updates.

Put winapp2ool in that folder, set the flavor to NonCCleaner, and choose **Update Winapp2.ini**. The same install from the command line:

```
winapp2ool -download winapp2 -base -1d "%AppData%\ChemTable Software\Reg Organizer\CleanupDatabases"
```

### Configuration

In the left toolbar, open the *Miscellaneous Tools* drop-down and select *Community-Driven Cleanup*. Every entry is scanned by default, so review the results carefully before you choose what to delete.

---

## Keeping your copy up to date

Run winapp2ool from the folder holding your winapp2.ini and choose `Update Winapp2.ini` or `Update & Trim`. If your copy's header has a `; Flavor:` line, winapp2ool updates it in that flavor. Otherwise it uses the flavor in Settings, which only lasts between sessions if you turned on saving and reading settings, so set it again if you didn't. Before you update, `Show winapp2.ini changelog` shows what changed since your copy was published, and `Show trimmed changelog` does the same against a trimmed copy.

You can also update in-place using the winapp2ool commands provided in the sections above.

Applications that ship their own copy of winapp2.ini may overwrite yours when they update themselves.

See the [winapp2ool readme](winapp2ool/Readme.md) for the full set of options, including the command-line arguments for running it unattended.
