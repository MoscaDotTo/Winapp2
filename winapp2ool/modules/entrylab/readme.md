# Entry Lab

**Entry Lab** is a winapp2ool menu for the two generators that build winapp2.ini entries from templates. They're mostly for winapp2.ini maintainers, since the build runs both, but they're handy if you write entries for your own use too. Open it from the Entry Lab option on the main menu.

# Menu Options

| Option            | Effect                                                                 |
| :-                | :-                                                                     |
| 0. Exit           | Returns to the main menu                                               |
| 1. [BrowserBuilder](../browserbuilder/readme.md) | Generates winapp2.ini entries for web browsers (Chromium and Gecko families) |
| 2. [EntryBuilder](../entrybuilder/readme.md)     | Generates winapp2.ini entries for win32 and [Microsoft Store](../entrybuilder/readme.md#store-and-hybrid-apps) apps from a shorthand DSL |

Pressing Enter without a selection opens BrowserBuilder.

# Command-line arguments

Entry Lab has no command-line argument of its own. To run a generator from the command line, name it directly:

| Arg                      | Effect                 |
| :-                       | :-                     |
| `6` or `browserbuilder`  | Launches BrowserBuilder |
| `11` or `entrybuilder`   | Launches EntryBuilder   |

Each generator's readme lists its own arguments, and the [main readme](../../Readme.md) covers the global ones.
