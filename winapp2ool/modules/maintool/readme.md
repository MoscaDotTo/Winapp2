# Winapp2ool Global Settings

The **Global Settings** menu provides high-level controls for winapp2ool's application-wide behavior: persistence of settings to disk, offline mode, beta participation, the active winapp2.ini flavor, and log management. It is accessible from the main menu by selecting **Settings**.

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
| Toggle Saving Settings | Enable or disable writing winapp2ool's settings to `winapp2ool.ini` on exit | Default: `False` |
| Toggle Reading Settings | Enable or disable loading settings from `winapp2ool.ini` at startup | Default: `False` |
| Toggle Beta Participation | Opt in or out of beta builds of winapp2ool | Default: `False`; requires .NET Framework 4.6+; triggers an immediate update |
| Toggle Offline Mode | Force winapp2ool into offline mode | Default: `False` |
| Change Flavor | Cycle the active winapp2.ini flavor to the next in sequence | See [Winapp2.ini Flavor](#winapp2ini-flavor) |
| View Log | Print winapp2ool's current internal log to the console | |
| Save Log | Write winapp2ool's internal log to disk | |
| Change Save Target | Select a new file path for the log save target | Default: `winapp2ool.log` in the current directory |
| Visit GitHub | Open the winapp2 GitHub page in the default web browser | |
| Reset Settings | Restore all global settings to their defaults | Only shown when settings have been changed |

---

# Settings Persistence

By default, winapp2ool starts with factory defaults every time it runs. Changes to module settings (file paths, toggles, etc.) made during a session are lost when winapp2ool exits.

Enabling **Saving Settings** causes winapp2ool to write all current settings to `winapp2ool.ini` when it exits. Enabling **Reading Settings** causes winapp2ool to load that file at startup and restore the saved state. Both toggles must be enabled for settings to persist across sessions.

The two are kept separate so you can read a fixed settings file without having runtime changes overwrite it.

---

# Winapp2.ini Flavor

The active flavor determines which version of winapp2.ini is used when modules download a remote copy. **Change Flavor** cycles through the available flavors in order:

| Flavor | Description |
|:-|:-|
| CCleaner | The CCleaner-compatible variant (default) |
| NonCCleaner | The base (non-CCleaner) variant |
| BleachBit | The BleachBit-compatible variant |
| SystemNinja | The System Ninja-compatible variant |
| Tron | The Tron-compatible variant |
| CCleaner7 | The CCleaner 7-compatible variant |
| FluentCleaner | The FluentCleaner-compatible variant |

The current flavor is shown on the Settings menu. Modules that download winapp2.ini (Diff, Trim, CC7Patcher when downloading) use this setting to determine which file to fetch.

---

# Offline Mode

Offline mode disables all network operations. Modules and menu options that require an internet connection become unavailable. winapp2ool enters offline mode automatically if it cannot establish a network connection at startup; you can also force it on manually via **Toggle Offline Mode**.

Use **Go online** from the main menu (visible when in offline mode) to retry the network connection.

---

# Log Management

winapp2ool maintains an internal log throughout its runtime, recording both diagnostic output and a record of operations performed. The log grows over the course of a session.

- **View Log** prints the current log to the console
- **Save Log** writes the log to disk at the configured save target
- **Change Save Target** changes where the log is saved

The log save target defaults to `winapp2ool.log` in the current directory. The global log is also accessible at any time from the main menu by typing `printlog` (view) or `savelog` (save).

---

# Beta Participation

Enabling beta participation switches winapp2ool to the beta release track and immediately triggers a self-update to download the latest beta build. Beta builds may contain experimental features or unstable code. Requires .NET Framework 4.6 or later.

---

# Troubleshooting

| Symptom | Cause |
|:-|:-|
| Settings are not saved between sessions | Both **Saving Settings** and **Reading Settings** must be enabled; saving alone writes the file, reading alone loads a previously saved file |
| Beta option is unavailable | Your .NET Framework version is below 4.6 |
| Flavor change has no effect | The flavor only applies to modules that download a remote winapp2.ini; modules using a local file are not affected |
| winapp2ool is stuck in offline mode | The network check failed at startup; use **Go online** from the main menu to retry |
