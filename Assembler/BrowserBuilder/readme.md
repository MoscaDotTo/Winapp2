# BrowserBuilder 

### What is this?

This folder contains the set of files needed by Winapp2ool to generate entries for web browsers, including the separate rulesets for both Chromium and Gecko and the Transmutation rules required to correct or augment them.  

### Why are web browser entries generated now?

Web browsers are a particularly complex and fast evolving ecosystem. In order to provide consistent support across a larger number of web browsers, the Browser Builder system takes in some general information about the browsers and applies this information to a set of "Entry Scaffolds." This results in a consistent set of cleaning rules applied across all browsers for a particular engine and allows changes to be made across all browsers with less cognitive effort and no need to manually edit dozens of different browser configurations. 

### Why are there Transmutation rules? 

The generation process is not perfect, it is meant only to establish the most consistent possible baseline for the set of all browser entries and reduce the effort required to maintain them. Variations in features and implementations between browsers sometimes necessitate that generated content be added to, subtracted from, or modified in order to better fit its purpose. 

# Files 
| File                                                                                                                                                                 | Description                                                                                        |
| :-                                                                                                                                                                   | :-                                                                                                 |
| [chromium.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/BrowserBuilder/chromium.ini)                                         | Contains the generative rulesets and browser information for Chromium browsers                     | 
| [gecko.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/BrowserBuilder/gecko.ini)                                               | Contains the generative rulesets and browser information for Gecko browsers                        |
| [browser_additions.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/BrowserBuilder/browser_additions.ini)                       | Contains the set of entries and keys which are added to the set of generated browser entries       |
| [browser_key_replacements.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/BrowserBuilder/browser_key_replacements.ini)         | Contains the set of key replacement values made to the set of generated browser entries            |
| [browser_name_removals.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/BrowserBuilder/browser_name_removals.ini)               | Contains the set of keys to be removed by their name from the set of generated browser entries     |
| [browser_section_removals.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/BrowserBuilder/browser_section_removals.ini)         | Contains the set of sections to be removed by their name from the set of generated browser entries | 
| [browser_section_replacements.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/BrowserBuilder/browser_section_replacements.ini) | Contains the set of section replacements made to the set of generated browser entries              | 
| [browser_value_removals.ini](https://raw.githubusercontent.com/MoscaDotTo/Winapp2/refs/heads/master/Assembler/BrowserBuilder/browser_value_removals.ini)             | Contains the set of key values to be removed from the set of generated browser entries             | 