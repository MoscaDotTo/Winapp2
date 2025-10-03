:: building winapp2.ini with this script requires winapp2ool v1.6 or newer
@echo off
echo Creating backups of existing winapp2.ini files 
copy ..\Winapp2.ini winapp2-cc.old
copy ..\Non-CCleaner\Winapp2.ini winapp2.old
copy ..\Non-CCleaner\BleachBit\Winapp2.ini winapp2-bb.old
copy ..\Non-CCleaner\SystemNinja\Winapp2.rules winapp2-sn.old
copy ..\Non-CCleaner\Tron\Winapp2.ini winapp2-tron.old
echo Building winapp2.ini, please wait
echo Combining base entries
winapp2ool -s -combine -1d \Entries -3f base-entries-combined.ini
echo Generating browser entries
winapp2ool -s -browserbuilder -1d \BrowserBuilder -2d \BrowserBuilder -3f browsers.ini -4d \BrowserBuilder -5d \BrowserBuilder -6d \BrowserBuilder -7d \BrowserBuilder -8d \BrowserBuilder -9d \BrowserBuilder
echo Joining base entries with browser entries
winapp2ool -s -transmute -add -1f base-entries-combined.ini -2f browsers.ini -3f winapp2.ini
echo Performing static analysis and saving corrections
winapp2ool -s -debug -usedate -c -1f Winapp2.ini -3f Winapp2.ini
echo Creating changelog
winapp2ool -s -diff -d -savelog -1f winapp2.old -2f Winapp2.ini -3f diff.txt
echo Winapp2.ini built 
copy Winapp2.ini ..\Non-CCleaner
copy diff.txt ..\Non-CCleaner
echo Creating CCleaner flavor
winapp2ool -s -flavorize -autodetect -2f winapp2-ccleaner-flavor.ini -9d \CCleaner
echo Creating BleachBit flavor 
winapp2ool -s -flavorize -autodetect -2f winapp2-bleachbit-flavor.ini -9d \BleachBit
echo Creating Tron flavor 
winapp2ool -s -flavorize -autodetect -1f winapp2-ccleaner-flavor.ini -2f winapp2-tron-flavor.ini -9d \Tron
echo Creating System Ninja flavor 
winapp2ool -s -flavorize -autodetect -2f winapp2-systemninja-flavor.ini -9d \SystemNinja
echo Performing static analysis and saving corrections
winapp2ool -s -debug -usedate -c -1f winapp2-ccleaner-flavor.ini 
echo Creating changelog
winapp2ool -s -diff -d -savelog -1f winapp2-cc.old -2f Winapp2.ini -3f diff.txt
copy Winapp2.ini ..
copy diff.txt ..
echo CCleaner flavor built 
winapp2ool -s -debug -usedate -c -1f winapp2-bleachbit-flavor.ini
echo Creating changelog
winapp2ool -s -diff -d -savelog -1f winapp2-bb.old -2f Winapp2.ini -3f diff.txt
copy Winapp2.ini ..\Non-CCleaner\BleachBit
copy diff.txt ..\Non-CCleaner\BleachBit
echo BleachBit flavor built 
winapp2ool -s -debug -usedate -c -1f winapp2-tron-flavor.ini
echo Creating changelog
winapp2ool -s -diff -d -savelog -1f winapp2-tron.old -2f Winapp2.ini -3f diff.txt
copy Winapp2.ini ..\Non-CCleaner\Tron
copy diff.txt ..\Non-CCleaner\Tron
echo Tron flavor built 
winapp2ool -s -debug -usedate -c -1f winapp2-systemninja-flavor.ini -3f Winapp2.rules
echo Creating changelog
winapp2ool -s -diff -d -savelog -1f winapp2-sn.old -2f Winapp2.rules -3f diff.txt
copy Winapp2.rules ..\Non-CCleaner\SystemNinja
copy diff.txt ..\Non-CCleaner\SystemNinja
echo System Ninja flavor built
echo cleaning up
del winapp2*.old
del winapp2*.ini
del base-entries-combined.ini
del browsers.ini
del winapp2.rules
del diff.txt
echo Winapp2.ini successfully built