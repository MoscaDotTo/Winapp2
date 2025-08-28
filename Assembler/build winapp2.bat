:: building winapp2.ini with this script requires winapp2ool v1.6 or newer
@echo off
echo Building winapp2.ini, please wait
echo Combining base entries
winapp2ool -s -combine -1d \Entries -3f base-entries-combined.ini
echo Generating browser entries
winapp2ool -s -browserbuilder -1d \BrowserBuilder -2d \BrowserBuilder -3f browsers.ini -4d \BrowserBuilder -5d \BrowserBuilder -6d \BrowserBuilder -7d \BrowserBuilder -8d \BrowserBuilder -9d \BrowserBuilder
echo Joining base entries with browser entries
winapp2ool -s -transmute -add -1f base-entries-combined.ini -2f browsers.ini -3f winapp2.ini
echo Performing static analysis and saving corrections
winapp2ool -s -debug -usedate -c -1f winapp2.ini -3f winapp2.ini
move winapp2.ini ..
del base-entries-combined.ini
del browsers.ini