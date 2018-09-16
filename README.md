# Winapp2
A database of extended cleaning routines for popular Windows PC based maintenance software.


Winapp2 comes in 3 different versions. The first version is Winapp2 itself, which is specifically for CCleaner users. It is optimized to remove Winapp2 entries that are already included in CCleaner by default. The second version is the same, except it contains all the removed entries that were removed due to being added in CCleaner. This are for non-CCleaner users. The 3rd version is Winapp3, a much more advanced version that contains dangerous entries that could damage software or even nuke your system. Use this file with extreme cautious.

We also have a tool specifically for these files, which is called Winapp2ool. This tool allows you to combine custom.ini files, detect errors in Winapp2, and trim your file to remove unneeded entries.


### How to use Winapp2 for the following cleaners:

- For [CCleaner users](https://www.ccleaner.com/ccleaner), download the latest copy from here. After that, move the Winapp2 copy to your main directory (where ever the CCleaner.exe file is). Optionally, if CCleaner is taking too long to load, please download Winapp2ool, run it, and run the Trim feature to trim your Winapp2 file. This will speed up CCleaner load times with Winapp2.

- For [BleachBit users](https://www.bleachbit.org), open up BleachBit, click on "Edit" tab, then click on "Preferences". Make sure "Download and update cleaners from community (winapp2.ini)" is checked.

- For [System Ninja users](https://singularlabs.com/software/system-ninja), System Ninja has Winapp2 built into their cleaner and you should already have a copy installed in your System Ninja/scripts folder.

- For [Avira System Speedup users](https://www.avira.com/en/avira-system-speedup-free), Avira System Speedup has Winapp2 built into their cleaner and you should already have a copy in your Avira/System Speedup/sdf folder.

- For [Tron users](https://github.com/bmrf/tron), Tron has Winapp2 built into their cleaner and you should already have a copy in your tron/resources/stage_1_tempclean/ccleaner folder.


### How to make your own entries:

Winapp2 entries follow ABC ordering and we currently support all of the following code:

**[Name \*]** Name of the entry. Please include the space between the name and the \*.

**DetectOS=** Specifies which OS this entry is for. We use kernel version numbering, so for example, 5.1 refers to Windows XP and 6.0 refers to Windows Vista. If you are writing an entry for all Windows operating systems, then you do not need to include this. If you are writing an entry that only exists on Windows XP, you would put DetectOS=5.1|5.1. If you are writing an entry that is for Vista and above, you would put DetectOS=6.0|. If you are writing an entry that is for Windows 7 and below, you would put DetectOS=|6.1. It is very important the vertical bar is put in the right position, otherwise it will detect the wrong operating systems.

**LangSecRef=** or **Section=** This refers to what section this belongs to. Each entry must be sectioned properly. We use numbering to refer to a specific section. You can also use Section="Name" to refer to a custom entry. We currently only use Section= for Games and for Winapp3. The following numbers refer to the specific section:

LangSecRef=3021 = Applications  
LangSecRef=3022 = Internet  
LangSecRef=3023 = Multimedia  
LangSecRef=3024 = Utilities  
LangSecRef=3025 = Windows  
LangSecRef=3026 = Firefox  
LangSecRef=3027 = Opera  
LangSecRef=3028 = Safari  
LangSecRef=3029 = Google Chrome  
LangSecRef=3030 = Thunderbird  
LangSecRef=3031 = Windows Store  
Section=Games

**Detect=** This refers to the program itself. This is needed in order to clean any program. Usually these refer to a registry key, however, if one doesn't exist, you will need to use **DetectFile=** instead a point to a specific file path or file (usually the main directory or the main .exe file).

**DetectFile=** As mentioned above, this is for when a registry key doesn't exist and must use a file path or a specific file.

**Default=** This refers to if the entry should be enabled by default. CCleaner is the only program that requires this to exist. This should always be **Default=False**, even if the entry is guaranteed to be safe.

**Warning=** If an entry may break a program or does something weird to one's system, it is a good idea to put a warning in an entry. Please leave warnings short and to the point.

**FileKey=** This refers to the junk files that need to be cleaned. This can include a path or a specific file. Each file or path must be in its own FileKey. So, if you are trying to clean 2 file paths, you would put 1 path in FileKey1= and the other in FileKey2=. You may also use wildcards in FileKeys to help shrink the amount of FileKeys created. A good time to use these would be when cleaning junk files with the same extension, such as .log. Instead of making multiple FileKeys for cleaning multiple log files, you can make one entry with using a wildcard, for example: Path|\*.log This will tell Winapp2 to clean any file that has .log at the end. If your goal is to delete any file within a folder, then you would specify this with just adding a period, for example: path|\*.\* A | must be used at the end of each path for every FileKey in order for your entry to work properly. Feel free to look throughout Winapp2 if you ever need a better example. Alternatively, you can use RECURSE and REMOVESELF in your FileKeys. **RECURSE** tells your entry to clean the files within the path specified, as well as in sub-folders. **REMOVESELF** does the same as RECURSE, except it also removes the folders along with it, as well.

**RegKey=** This is for cleaning registry entries. The process is relatively the same as FileKey, except you would specify each RegKey as a registry path. We support all registry paths in Windows. We do not support wildcards in RegKey, so each RegKey has to be the whole path.


### Environment variables:

These are all the possible variables that can be used for writing paths in Winapp2.

**%AppData%**  
*Vista-10:* C:\\Users\\{username}\\AppData\\Roaming  
*XP:* C:\\Documents and Settings\\{username}\\Application Data

**%CommonAppData%**  
*Vista-10:* C:\\ProgramData  
*XP:* C:\\Documents and Settings\\All Users\\Application Data

**%CommonProgramFiles%**  
*XP-10:* C:\\Program Files\\Common Files  
*XP-10:* C:\\Program Files (x86)\\Common Files  
(This will work for both 32-bit and 64-bit folders as CCleaner will detect it on its own.)

**%Documents%**  
*Vista-10:* C:\\Users\\{username}\\Documents  
*XP:* C:\\Documents and Settings\\{username}\\My Documents

**%LocalAppData%**  
*Vista-10:* C:\\Users\\{username}\\AppData\\Local  
*XP:* C:\\Documents and Settings\\{username}\\Local Settings\\Application Data

**%LocalLowAppData%**  
*Vista-10:* C:\\Users\\{username}\\AppData\\LocalLow

**%Music%**  
*Vista-10:* C:\\Users\\{username}\\Music  
*XP:* C:\\Documents and Settings\\{username}\\My Documents\\My Music

**%Pictures%**  
*Vista-10:* C:\\Users\\{username}\\Pictures  
*XP:* C:\\Documents and Settings\\{username}\\My Documents\\My Pictures

**%ProgramFiles%**  
*XP-10:* C:\\Program Files  
*XP-10:* C:\\Program Files (x86)  
(This will work for both 32-bit and 64-bit folders as CCleaner will detect it on its own.)

**%Public%**  
*Vista-10:* C:\\Users\\Public

**%SystemDrive%**  
*XP-10:* C:

**%UserProfile%**  
*Vista-10:* C:\\Users\\{username}  
*XP:* C:\\Documents and Settings\\{username}

**%Video%**  
*Vista-10:* C:\\Users\\{username}\\Videos  
*XP:* C:\\Documents and Settings\\{username}\\My Documents\\My Videos

**%WinDir%**  
*XP-10:* C:\\Windows


### Excluding a file or path:

In a situation where you are writing an entry, but you do not want to remove a certain file because it may damage your software, ExcludeKeys can be used. You may use ExcludeKeys to exclude a file, path, or a registry key. Wildcards are also acceptable in ExcludeKeys. Examples are below.

*Example 1:* ExcludeKey1=PATH|%WinDir%\\System32\\LogFiles\\SCM\\|\*-\*-\*-\*.\*

This will exclude all of the log files with the pattern \*-\*-\*-\*.\* in the \\system32\\LogFiles\\SCM\\ folder of the user's Windows directory.

*Example 2:* ExcludeKey2=FILE|%WinDir%\\System32\\LogFiles\\|myfile.txt

This will exclude the myfile.txt file located in the \\system32\\LogFiles folder of the user's Windows directory.

*Example 3:* ExcludeKey3=REG|HKCU\\Software\\Piriform

This will exclude the key located at HKCU\\software\\piriform (and any subentries)

*Example 4:* ExcludeKey4=PATH|C:\\Temp\\|\*.\*

This will exclude the files located in the C:\\temp folder and all subfolders.

*Example 5:* ExcludeKey5=PATH|C:\\Windows\\|\*.exe;\*.bat

This will exclude files of types .exe and .bat in the C:\\Windows folder.


### Currently not supported features:

- We do not support wildcards in RegKeys.

- We do not support non-English systems, although it should still be possible to run Winapp2 on non-English systems. If you need a cleaner in a non-English format, we suggest writing a custom.ini file.

- We do not support portable software. This is because these can be placed anywhere on a system. We suggest writing a custom.ini file for portable software.


### Donations:

If you like Winapp2 and want to help keep development going, please consider donating some money to our great minds behind Winapp2.

Donate to Robert/MoscaDotTo: http://www.winapp2.com/ go to the link and click on Donate
