# Entry Lab

**Entry Lab** is a menu grouping the winapp2ool modules that generate winapp2.ini entries from templates. It is reached as the Entry Lab option on the main winapp2ool menu.

# Menu Options

| Option            | Effect                                                                 |
| :-                | :-                                                                     |
| 0. Exit           | Returns to the main menu                                               |
| 1. [BrowserBuilder](../browserbuilder/readme.md) | Generates winapp2.ini entries for web browsers (Chromium and Gecko families) |
| 2. [UWPBuilder](../uwpbuilder/readme.md)         | Generates winapp2.ini entries for Universal Windows Platform (Microsoft Store) apps |
| 3. [EntryBuilder](../entrybuilder/readme.md)     | Generates winapp2.ini entries for win32 apps from a shorthand DSL      |

Pressing Enter without a selection opens BrowserBuilder.

# Command-line arguments

Entry Lab itself has no command-line argument. The generators it groups are launched directly as modules:

| Arg                      | Effect                 |
| :-                       | :-                     |
| `6` or `browserbuilder`  | Launches BrowserBuilder |
| `10` or `uwpbuilder`     | Launches UWPBuilder     |
| `11` or `entrybuilder`   | Launches EntryBuilder   |

Refer to each generator's readme for its module-specific arguments, and to the [main readme](../../Readme.md) for global arguments.
