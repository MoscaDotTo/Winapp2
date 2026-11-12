# BrowserBuilder

### What is this?

This folder holds the files winapp2ool uses to generate entries for web browsers: one ruleset for Chromium and one for Gecko, plus the Transmutation rules that correct or extend what they generate. The build writes the result to `Assembler/Generated/BrowserBuilder/browsers.ini`.

### Why are web browser entries generated now?

There are a lot of browsers, and they change fast. BrowserBuilder takes a little information about each one (where it keeps its data, its registry key, which section it goes in) and applies it to a set of "entry scaffolds", so every browser on an engine gets the same cleaning rules. A fix to a scaffold reaches every browser on the next build.

### Why are there Transmutation rules?

Generation gives every browser the same baseline, but browsers aren't all the same. One lacks a feature, another keeps some data somewhere else, a third needs something extra. The Transmutation rules add to, remove from, or change the generated entries for those browsers.

# Files

| File | Description |
| :- | :- |
| [chromium.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/Sources/BrowserBuilder/chromium.ini) | The Chromium browser definitions and the scaffolds applied to them |
| [gecko.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/Sources/BrowserBuilder/gecko.ini) | The Gecko browser definitions and the scaffolds applied to them |
| [browser_additions.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/Sources/BrowserBuilder/browser_additions.ini) | Entries and keys added to the generated browser entries |
| [browser_key_replacements.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/Sources/BrowserBuilder/browser_key_replacements.ini) | Keys whose values are replaced in the generated entries |
| [browser_name_removals.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/Sources/BrowserBuilder/browser_name_removals.ini) | Keys removed by name from the generated entries |
| [browser_section_removals.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/Sources/BrowserBuilder/browser_section_removals.ini) | Generated entries removed entirely |
| [browser_section_replacements.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/Sources/BrowserBuilder/browser_section_replacements.ini) | Generated entries replaced entirely |
| [browser_value_removals.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/Sources/BrowserBuilder/browser_value_removals.ini) | Keys removed by value from the generated entries |
