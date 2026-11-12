# Winapp2ool Global Settings

The **Global Settings** menu holds the settings that apply to all of winapp2ool: whether settings are saved to disk, offline mode, beta builds, the active winapp2.ini flavor, and the log. Open it with **Settings** on the main menu.

---

# Table of Contents

1. [Menu Options](#menu-options)
2. [Settings Persistence](#settings-persistence)
3. [Winapp2.ini Flavor](#winapp2ini-flavor)
4. [Offline Mode](#offline-mode)
5. [Log Management](#log-management)
6. [Beta Participation](#beta-participation)
7. [Troubleshooting](#troubleshooting)

---

# Menu Options

| Option | Effect | Notes |
|:-|:-|:-|
| Toggle Saving Settings | Save your settings to `winapp2ool.ini`, so they carry over to the next session | Default: `False` |
| Toggle Reading settings | Load the modules' settings from `winapp2ool.ini` at startup | Default: `False` |
| Toggle Beta Participation | Switch to or from the beta builds of winapp2ool | Default: `False`. Needs .NET Framework 4.8. Turning it on updates winapp2ool right away |
| Toggle Offline Mode | Turn offline mode on or off for this session | Default: on only when winapp2ool can't connect at startup. Never saved |
| Change Flavor | Move to the next winapp2.ini flavor | See [Winapp2.ini Flavor](#winapp2ini-flavor) |
| View Log | Print winapp2ool's log so far | |
| Save log | Write winapp2ool's log to disk | |
| Change Save Target | Choose where the log is saved | Default: `winapp2ool.log` in the current folder |
| Visit GitHub | Open the winapp2 GitHub page in your default browser | |
| Reset Settings | Put every global setting back to its default | Shown only when a setting has changed |

---

# Settings Persistence

By default, winapp2ool starts from its defaults every time it runs, and anything you change in a module (file paths, toggles and so on) is gone when it exits.

Turning on **Saving Settings** writes your settings to `winapp2ool.ini` in the current folder as you change them. Turning on **Reading Settings** loads every module's settings from that file at startup. Module settings carry over only with both on. The settings on this menu (the two toggles themselves, beta participation, the flavor and the log's save target) load whenever `winapp2ool.ini` exists, even with Reading Settings off.

They're separate so you can read a fixed settings file without a session's changes overwriting it. Command line runs never write the file.

---

# Winapp2.ini Flavor

The active flavor decides which winapp2.ini winapp2ool downloads when it needs the latest copy. **Change Flavor** moves through the flavors in this order:

| Flavor | Description |
|:-|:-|
| CCleaner | For CCleaner Classic (versions 1 to 6) |
| Base | The base winapp2.ini every other flavor is built from (default) |
| BleachBit | For BleachBit 5 and older |
| SystemNinja | For System Ninja |
| Tron | For Tron |
| CCleaner7 | For CCleaner 7 |
| FluentCleaner | For FluentCleaner |

The Settings menu shows the current flavor. Diff and Trim use it when they download winapp2.ini, and so does the Download module's `winapp2` command line argument. WinappDebug uses it to decide whether to run its BleachBit and System Ninja checks.

Some downloads don't use it. The main menu's update options download the flavor your local winapp2.ini names in its header, and use this setting only when the file names none or you don't have one. The Download menu has an option for each flavor, and CC7Patcher always downloads the CCleaner 7 flavor.

---

# Offline Mode

In offline mode winapp2ool makes no network requests, and the options that need a connection aren't available. winapp2ool goes offline by itself when it can't connect at startup, and **Toggle Offline Mode** switches it by hand. Offline mode lasts only for the session: it's never saved, and the next launch checks the connection again. `-offline` on the command line forces it for one run.

While offline, **Go online** on the main menu retries the connection.

---

# Log Management

winapp2ool keeps a log while it runs, with diagnostic output and a record of what it did. **View Log** prints it and **Save log** writes it to the save target, which is `winapp2ool.log` in the current folder unless you change it with **Change Save Target**. You can also type `printlog` or `savelog` at the main menu.

A silent (`-s`) run that fails saves the log by itself. Add `-writelog` to the command line to save it at the end of any run.

---

# Beta Participation

Turning on beta participation moves winapp2ool to the beta track and updates it to the latest beta right away. Beta builds can have experimental or unfinished features. It needs .NET Framework 4.8.

Turning it off doesn't roll you back. winapp2ool only ever updates to a newer version, so you stay on the beta build until a release newer than it comes out.

---

# Troubleshooting

| Symptom | Cause |
|:-|:-|
| Settings are not saved between sessions | Module settings need both **Saving Settings** and **Reading Settings** on. Saving alone writes the file, and reading alone loads a file saved earlier |
| Saving Settings is back on after I turned it off | Turning saving off is itself never saved, so a `winapp2ool.ini` saved with it on turns it back on at the next launch. Delete `winapp2ool.ini` to stop saving for good |
| Beta option is refused | Your .NET Framework is older than 4.8 |
| Changing the flavor has no effect | The flavor only matters when winapp2ool downloads winapp2.ini. A module working on a local file isn't affected, and the main menu's update options go by your local file's flavor |
| winapp2ool is stuck in offline mode | The connection check failed at startup. Use **Go online** on the main menu to retry |
