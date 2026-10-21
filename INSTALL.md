# Installing Winapp2.ini

This page covers installing and configuring winapp2.ini in each application that supports it. For what winapp2.ini is and how it is built, see the [main readme](README.md).

Whichever application you use, install winapp2.ini with [winapp2ool.exe](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2ool.exe) and keep a copy of it in the same folder as winapp2.ini. That is how you keep the file up to date afterwards.

The download links on this page fetch each file from the [latest release](https://github.com/MoscaDotTo/Winapp2/releases/latest), where most file names carry their flavor, such as `winapp2-ccleaner.ini`. If you install one by hand, rename it to `winapp2.ini` first, or to `winapp2.rules` for System Ninja.

---

## Installing with winapp2ool

Winapp2ool is the recommended way to install winapp2.ini for every application on this page, and the only one that also keeps it current afterwards. The procedure is the same everywhere:

1. Download [winapp2ool.exe](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2ool.exe) and put it in the folder where your application looks for winapp2.ini. Each section below names that folder.
2. Run it, open **Settings**, and use **Change Flavor** until the flavor named in your application's section is shown. It cycles through CCleaner, NonCCleaner, BleachBit, SystemNinja, Tron, CCleaner7, and FluentCleaner. The default is CCleaner. Enable reading and writing settings from/to disk to persist your choice between sessions.
3. Return to the main menu and choose **Update Winapp2.ini** to install, or **Update & Trim** to install a copy with the entries that do not apply to your machine removed.

Both options write `winapp2.ini` into the folder from which winapp2ool is running, overwriting anything already there. They are listed only while winapp2ool sees your local copy as out of date, including having no local copy at all, so they are waiting for you on a first install. Neither is available in offline mode.

CCleaner 7 has no standalone winapp2.ini and is installed with [CC7Patcher](#ccleaner-7). System Ninja expects a file named `winapp2.rules`, which the main menu options do not produce, so it is installed through the Downloader instead, see [its section](#system-ninja).

### From the command line

Every install below can be scripted. Winapp2ool is built around keeping the exe in the folder where winapp2.ini belongs, so if that is what you have done, leave the paths off and let it default to the folder from which it is running:

```
winapp2ool -download winapp2
```

That fetches whichever flavor winapp2ool is set to and writes `winapp2.ini` beside the exe. A flavor flag overrides the setting for one run without changing it: `-cc`, `-base`, `-bb`, `-sn`, `-tron`, `-cc7`, `-fc`.

To download and trim in one step, writing the result straight to `winapp2.ini`:

```
winapp2ool -trim -d -3f winapp2.ini
```

Without `-3f` the result is written as `winapp2-trimmed.ini` and your existing `winapp2.ini` is left alone, which is handy if you want to compare the two. Without `-d` it trims the local `winapp2.ini` you already have instead of fetching a fresh one; the file is read fully before it is written back, so trimming in place is safe.

Add `-s` to either command for unattended runs. It suppresses all output and the keypress winapp2ool otherwise waits for at the end.

Two things to watch. The defaults follow the working directory rather than the location of the exe, so calling winapp2ool by its full path from somewhere else puts the file wherever you called it from; in a scheduled task, set the "Start in" field to the target folder. And if you enabled reading settings from disk and have since changed the Downloader's save directory, that saved directory wins over the working directory for `-download`, so pass `-1d` if you want to be certain. Trim resets its paths on every command line run and is not affected.

The per-application sections below give the same commands with explicit paths, for installing without keeping the exe beside winapp2.ini. Those paths are written for `cmd`; in PowerShell, write `$env:ProgramFiles` in place of `%ProgramFiles%`.

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

You should use the [CCleaner flavor](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2-ccleaner.ini) for ideal integration into the UI and minimized rule overlap, however the base ("Non-CCleaner") [Winapp2.ini](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2.ini) will also work

### Installation

Place winapp2.ini in the same folder as ccleaner.exe. By default this is `..\Program Files\CCleaner`

It is advised that you use the Trim function of winapp2ool when updating winapp2.ini to reduce the CCleaner startup time, as the full winapp2.ini file can unnecessarily slow down the CCleaner start up process. With winapp2ool in that folder and the flavor left at its CCleaner default, **Update & Trim** is the main menu option you want.

The same install from the command line:

```
winapp2ool -trim -d -cc -3d "%ProgramFiles%\CCleaner" -3f winapp2.ini
```

For the untrimmed file:

```
winapp2ool -download winapp2 -cc -1d "%ProgramFiles%\CCleaner"
```

### Configuration

CCleaner will display the set of winapp2.ini entries which it detects as valid for your system inside its Applications tab. In modern versions of CCleaner, this tab is found in the Custom Clean section of the application. All winapp2.ini entries are disabled by default in CCleaner, and must be enabled individually or in groups. To enable an entire group of entries, right click on the section header and select "Check all."

###### Note: CCleaner 5.64.7577 is the last version to work on Windows XP and Vista (for non-SSE2 CPUs CCleaner 5.26.5937). Winapp2.ini and Winapp3.ini will continue to work with this version.

## CCleaner 7

###### [Download CCleaner](https://www.ccleaner.com/ccleaner)

CCleaner 7 no longer loads a separate winapp2.ini. Its cleaning definitions live inside `ccleaner.ini` alongside CCleaner's own, in a modified entry format. Installation is therefore not a matter of dropping a file next to `ccleaner.exe`; the entries have to be patched into `ccleaner.ini`. Winapp2ool's CC7Patcher does this, and it is the only installation method we endorse.

### Flavor

You should use the [CCleaner 7 flavor](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2-ccleaner7.ini), which CC7Patcher downloads for you by default. You do not need to download or place this file yourself. If you choose to supply your own local copy instead, it must be in the CCleaner 7 format; no other flavor will produce functional entries.

### Installation

Back up your `ccleaner.ini` before your first run. Patching rewrites the file in place, which reorders its sections and strips its comments.

1. Download [winapp2ool.exe](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2ool.exe) and run it
2. Select **CC7Patcher** from the main menu
3. Use **Change ccleaner.ini** to point at CCleaner 7's `ccleaner.ini`, by default this is `..\Program Files\Piriform\CCleaner 7`
4. Optionally enable **Toggle Trim** to install only the entries relevant to your system, which reduces CCleaner's startup time
5. Select **Run**

The same install from the command line:

```
winapp2ool -cc7patcher -2d "%ProgramFiles%\Piriform\CCleaner 7" -3d "%ProgramFiles%\Piriform\CCleaner 7"
```

Add `-trim` to trim before patching. CC7Patcher has its own ReadMe [here](https://github.com/MoscaDotTo/Winapp2/tree/master/winapp2ool/modules/cc7patcher).

### Updating

Run CC7Patcher again over your existing `ccleaner.ini`. It removes the entries left by the previous patch before installing the current ones, so nothing is duplicated and entries removed from winapp2.ini are cleared out. There is no need to restore a clean `ccleaner.ini` first.

You will also need to run it again after every CCleaner 7 update, as updating overwrites `ccleaner.ini` and removes the winapp2.ini entries with it.

### Configuration

CCleaner 7 will display the winapp2.ini entries it detects as valid for your system alongside its own cleaning options in the Custom Clean interface. All winapp2.ini entries are disabled by default and must be enabled individually or in groups.

All winapp2.ini entries are identified with a `custom-clean.name.` prefix which cannot be disabled or removed. This is a CCleaner 7 feature and not a winapp2.ini feature. There is no way to increase the width of the entry list, so to see the full name of an entry you must hover your mouse over it. Entries are sorted alphabetically, *ignoring* this prefix.

###### Note: CC7Patcher identifies the entries it installed by their `Author=Winapp2.ini Project` key, and removes them on the next run. If you customize a winapp2.ini entry inside `ccleaner.ini`, delete that key from your copy so your changes survive updating.

## BleachBit

###### [Download BleachBit](https://www.bleachbit.org)

### Flavor

**BleachBit 5 or earlier:** You should use the [BleachBit flavor](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2-bleachbit.ini), which removes registry exclusions not supported by BleachBit before version 6. Using any other flavor will throw a small number of errors and not allow you to run the entries which contain them, but will otherwise function correctly.

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

Alternatively, BleachBit maintains their own [customized version of winapp2.ini](https://github.com/bleachbit/winapp2.ini) which you can enable the use of from within the application:
1. Open BleachBit.
2. Select the "Edit" tab, and then "Preferences".
3. Check the box that reads "Download and update cleaners from community (Winapp2.ini)".

###### Note: Using BleachBit's winapp2.ini will overwrite any existing local file.

### Configuration

BleachBit will display the set of winapp2.ini entries which it detects as both having valid syntax and also as being valid for your system in its sidebar. All winapp2.ini entries are disabled by default in BleachBit, and must be enabled individually or in groups. To enable an entire group of entries, select the check box next to the section header.

###### Note: BleachBit 2.2 is the last version to work on Windows XP. Winapp2.ini and Winapp3.ini will continue to work with this version.

## System Ninja

###### [Download System Ninja](https://singularlabs.com/software/system-ninja)

### Flavor
You should use the [System Ninja Flavor](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2-systemninja.rules). This flavor is designed to improve compatibility with System Ninja by replacing keys with unsupported features, such that they become functional in System Ninja. It is not advised you use any other flavor with System Ninja.

### Installation
System Ninja ships with a copy of Winapp2.ini by default, served from their servers, storing it in your `..\System Ninja\scripts\` directory as `winapp2.rules`

To keep your system ninja winapp2.rules up to date with winapp2.ini using winapp2ool instead:
1. Open System Ninja
2. Select the "Options" tab
3. *Uncheck* the box that reads "Update cleaning rules and language files automatically"

This will prevent System Ninja from overwriting your local winapp2.rules file on launch

Winapp2ool's main menu update options always write `winapp2.ini`, so install this one through the Downloader:

1. Run winapp2ool and select **Downloader**
2. Use **Change Save Directory** to point at `..\System Ninja\scripts\`
3. Select **System Ninja Winapp2.rules**

The same install from the command line:

```
winapp2ool -download winapp2 -sn -1d "C:\Path\To\System Ninja\scripts"
```

### Configuration

System Ninja does not provide an interface for individually configuring which winapp2.ini entries run, they are all run and their output is displayed in the Junk Scanner window. The "Type" column in System Ninja displays the name of the entry as it appears in winapp2.ini

###### Note: System Ninja 3.2.7 is the last version to work on Windows XP and Vista. Winapp2.ini and Winapp3.ini will continue to work with this version.

## Avira System Speedup

###### [Download System Speedup](https://www.avira.com/en/avira-system-speedup-free)

### Flavor

You should use the base [winapp2.ini](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2.ini)

### Installation
Avira System Speedup ships with a copy of Winapp2.ini by default, served by Avira, storing it in your `..\Avira\System Speedup\sdf` directory. You can replace or update this local copy without issue or changing any of the Avira System Speedup settings.

Put winapp2ool in that folder, set the flavor to NonCCleaner, and choose **Update Winapp2.ini**. The same install from the command line:

```
winapp2ool -download winapp2 -base -1d "C:\Path\To\Avira\System Speedup\sdf"
```

### Configuration

Avira System Speedup scans every winapp2.ini entry and displays the results of that scan in a panel labeled Third Party Applications. You can manually enable/disable items within this menu before activating the clean function.

###### Note: Cleaning "Third Party Applications" is a paid feature of Avira System Speedup Pro. winapp2.ini is and always will be free, and supported by a variety of free applications.

## Tron

###### [Tron GitHub](https://github.com/bmrf/tron)

### Flavor

You should use the [Tron Flavor](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2-tron.ini). This flavor is designed to capture the downstream changes Tron makes to winapp2.ini while still otherwise remaining up-to-date with the latest changes.

### Installation

Tron maintains and ships their own [customized version of Winapp2.ini](https://raw.githubusercontent.com/bmrf/tron/refs/heads/master/resources/stage_1_tempclean/ccleaner/winapp2.ini) by default, storing it in your `..\resources\stage_1_tempclean\ccleaner` directory. You can overwrite this file to update winapp2.ini.

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

R-Wipe & Clean has unofficial support for Winapp2.ini.

R-Wipe & Clean imports winapp2.ini once through a prompt rather than reading it from a particular location on start. Put winapp2ool wherever you find convenient, set the flavor to NonCCleaner, and choose **Update Winapp2.ini** to fetch the file you are about to import. The same from the command line:

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
    * Because you will need to repeat the process of importing winapp2.ini as wipe lists in order to apply updates, it is advised you import it as a single wipelist which you can then easily remove after updating
11. Once winapp2.ini is imported, you should see a message saying `LangSecRef`, `Section`, and `Warning` are unsupported functions, press OK

### Configuration

Configure the imported wipe lists individually just as you would R-Wipe & Clean's native wipe lists

## HDCleaner

###### [Download HDCleaner](https://kurtzimmermann.com/index_en.html)

### Flavor

You should use the base [winapp2.ini](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2.ini)

### Installation

HDCleaner ships with an internal copy of winapp2.ini. You can override or supplement this by placing winapp2.ini in the same directory as `HDCleaner.exe`. By default, this is `..\Program Files\HDCleaner`

Put winapp2ool in that folder, set the flavor to NonCCleaner, and choose **Update Winapp2.ini**. The same install from the command line:

```
winapp2ool -download winapp2 -base -1d "%ProgramFiles%\HDCleaner"
```

### Configuration

After placing winapp2.ini in the same directory as `HDCleaner.exe`, you must configure HDCleaner's response to rule conflicts.

1. Open HDCleaner
2. From the side bar, select the bottom most icon (settings)
3. From the settings panel, select "Options"
4. Change the setting "If using Winapp2.ini" to "Ignore duplicate entries in HDCleaner resource file"

This will override the built in HDCleaner entries with your drop-in replacement

## FluentCleaner

###### [Download FluentCleaner](https://github.com/builtbybel/FluentCleaner/releases)

### Flavor

You should use the [FluentCleaner flavor](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2-fluentcleaner.ini). This flavor is designed to catch the downstream changes made by FluentCleaner to their winapp2.ini.

### Installation

FluentCleaner ships with a copy of winapp2.ini in the same folder as `FCleaner.exe` or `FluentCleaner.Classic.exe`. You can update winapp2.ini by replacing this file yourself.

Put winapp2ool in that folder, set the flavor to FluentCleaner, and choose **Update Winapp2.ini**. The same install from the command line:

```
winapp2ool -download winapp2 -fc -1d "C:\Path\To\FluentCleaner"
```

### Configuration

Almost all winapp2.ini entries are enabled by default in FluentCleaner when using their winapp2.ini or the FluentCleaner flavor. Click their checkbox to disable.

## Reg Organizer

###### [Download Reg Organizer](https://www.chemtable.com/organizer.htm)

###### ChemTable also provides their own information about winapp2.ini [here](https://www.chemtable.com/blog/en/community-driven-cleanup.htm)

### Flavor

You should use the base [winapp2.ini](https://github.com/MoscaDotTo/Winapp2/releases/latest/download/winapp2.ini)

### Installation

Reg Organizer ships with a copy of winapp2.ini maintained by ChemTable. By default, this is located in `..\AppData\Roaming\ChemTable Software\Reg Organizer\CleanupDatabases`. Replace the winapp2.ini file found here with the one from this repo to update it yourself. Be aware that Reg Organizer may replace this file when updating.

Put winapp2ool in that folder, set the flavor to NonCCleaner, and choose **Update Winapp2.ini**. The same install from the command line:

```
winapp2ool -download winapp2 -base -1d "%AppData%\ChemTable Software\Reg Organizer\CleanupDatabases"
```

### Configuration

From the left hand toolbar, select the *Miscellaneous Tools* drop down. From the resulting list, select *Community-Driven Cleanup*. All entries are scanned by default, so make sure to carefully review the results before selecting which to delete.

---

## Keeping your copy up to date

Run winapp2ool from the folder holding your winapp2.ini and choose `Update Winapp2.ini` or `Update & Trim`. The flavor you set the first time is only still in effect if you chose to save and read settings from disk; set it again if not. `Show winapp2.ini changelog` prints what changed since your copy was published before you commit to the update, and `Show trimmed changelog` does the same against a trimmed copy.

You can also update in-place using the winapp2ool commands provided in the sections above.

The applications that ship their own copy of winapp2.ini may overwrite yours when they update themselves.

See the [winapp2ool readme](winapp2ool/Readme.md) for the full set of options, including the command-line arguments for running it unattended.
