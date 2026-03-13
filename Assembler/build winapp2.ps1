<#
.SYNOPSIS
    Builds winapp2.ini and its various flavors using winapp2ool.

.DESCRIPTION
    This script requires winapp2ool v1.7 or newer.
    Creates backups, combines base entries, generates and corrects browser entries, 
    generates UWP entries, combines these into the base flavor, generates tool flavors,
    and creates a diff for each generated flavor.

.NOTES
    Author: Hazel Ward
    Version 20260313
    Copyright 2026
#>

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$script:Winapp2oolPath = $null

function Write-Step {
    param([string]$Message)
    # Adding a clear newline before each major step
    Write-Host "`n>> $Message" -ForegroundColor Cyan
}

function Write-ErrorMsg {
    param([string]$Message)
    Write-Host "ERROR: $Message" -ForegroundColor Red
}

function Find-Winapp2ool {
    # Check current directory first
    $localPath = Join-Path $PSScriptRoot "winapp2ool.exe"
    if (Test-Path $localPath) {
        Write-Host "Found winapp2ool.exe in current directory" -ForegroundColor Green
        return $localPath
    }
    
    # Check PATH
    $pathCommand = Get-Command "winapp2ool.exe" -ErrorAction SilentlyContinue
    if ($pathCommand) {
        Write-Host "Found winapp2ool.exe in PATH" -ForegroundColor Green
        return $pathCommand.Source
    }
    
    # Prompt user for path
    Write-Host "`nwinapp2ool.exe not found in current directory or PATH" -ForegroundColor Yellow
    Write-Host "This script requires winapp2ool v1.7 or newer`n" -ForegroundColor Yellow
    
    do {
        $userPath = Read-Host "Please enter the full path to winapp2ool.exe (or 'q' to quit)"
        
        if ($userPath -eq 'q') {
            return $null
        }
        
        $userPath = $userPath.Trim('"').Trim("'")
        
        if (Test-Path $userPath) {
            Write-Host "Found winapp2ool.exe at specified path" -ForegroundColor Green
            return $userPath
        }
        
        Write-Host "File not found at: $userPath" -ForegroundColor Red
    } while ($true)
}

function Invoke-Winapp2ool {
    param(
        [string[]]$Arguments,
        [string]$ErrorMessage = "Winapp2ool command failed"
    )
    
    try {
        # Dropping -NoNewWindow and using -WindowStyle Hidden
        # This isolates the tool in an invisible window, which can prevent
        # errors from being obviously observed. Revert to -NoNewWindow
        # if you need/want to see the console output 
        $process = Start-Process -FilePath $script:Winapp2oolPath `
                                 -ArgumentList $Arguments `
                                 -WindowStyle Hidden `
                                 -Wait `
                                 -PassThru
        
        $exitCode = $process.ExitCode

        if ($exitCode -ne 0) {
            Write-ErrorMsg "$ErrorMessage (Exit Code: $exitCode)"
            return $false
        }
        return $true
    }
    catch {
        Write-ErrorMsg "$ErrorMessage - $_"
        return $false
    }
}

function Backup-Files {
    Write-Step "Creating backups of existing winapp2.ini files"
    
    $backups = @(
        @{Source = "..\Winapp2.ini"; Dest = "winapp2-cc.old"}
        @{Source = "..\Non-CCleaner\Winapp2.ini"; Dest = "winapp2.old"}
        @{Source = "..\Non-CCleaner\BleachBit\Winapp2.ini"; Dest = "winapp2-bb.old"}
        @{Source = "..\Non-CCleaner\CCleaner7\Winapp2.ini"; Dest = "winapp2-cc7.old"}
        @{Source = "..\Non-CCleaner\SystemNinja\Winapp2.rules"; Dest = "winapp2-sn.old"}
        @{Source = "..\Non-CCleaner\Tron\Winapp2.ini"; Dest = "winapp2-tron.old"}
    )
    
    foreach ($backup in $backups) {
        if (Test-Path $backup.Source) {
            Copy-Item $backup.Source $backup.Dest -Force -ErrorAction SilentlyContinue
        }
    }
}

function Build-MainFile {
    Write-Step "Building winapp2.ini, please wait"
    
    Write-Step "Combining base entries"
    if (-not (Invoke-Winapp2ool -Arguments '-s', '-combine', '-1d', '\Entries', '-3f', 'base-entries-combined.ini' `
        -ErrorMessage "Failed to combine base entries")) { return $false }
    
    Write-Step "Generating browser entries"
    if (-not (Invoke-Winapp2ool -Arguments '-s', '-browserbuilder', '-1d', '\BrowserBuilder', '-2f', 'browsers.ini' `
        -ErrorMessage "Failed to generate browser entries")) { return $false }

    Write-Step "Generating UWP entries"
    if (-not (Invoke-Winapp2ool -Arguments '-s', '-uwpbuilder', '-1d', '\UWP', '-2f', 'uwp.ini' `
        -ErrorMessage "Failed to generate UWP entries")) { return $false }

    Write-Step "Joining base entries with browser entries"
    if (-not (Invoke-Winapp2ool -Arguments '-s', '-transmute', '-add', '-1f', 'base-entries-combined.ini', `
        '-2f', 'browsers.ini', '-3f', 'Winapp2.ini' `
        -ErrorMessage "Failed to join entries")) { return $false }

    Write-Step "Joining UWP entries"
    if (-not (Invoke-Winapp2ool -Arguments '-s', '-transmute', '-add', '-1f', 'Winapp2.ini', `
        '-2f', 'uwp.ini', '-3f', 'Winapp2.ini' `
        -ErrorMessage "Failed to join UWP entries")) { return $false }
    
    Write-Step "Performing static analysis and saving corrections"
    if (-not (Invoke-Winapp2ool -Arguments '-s', '-debug', '-usedate', '-c', '-1f', 'Winapp2.ini', '-3f', 'Winapp2.ini' `
        -ErrorMessage "Failed static analysis")) { return $false }
    
    Write-Step "Creating changelog"
    if (-not (Invoke-Winapp2ool -Arguments '-s', '-diff', '-d', '-savelog', '-1f', 'winapp2.old', `
        '-2f', 'Winapp2.ini', '-3f', 'diff.txt' `
        -ErrorMessage "Failed to create changelog")) { return $false }
    
    Copy-Item 'Winapp2.ini' '..\Non-CCleaner' -Force
    Copy-Item 'diff.txt' '..\Non-CCleaner' -Force
    Write-Host "Winapp2.ini built" -ForegroundColor Green
    
    return $true
}

function Build-CCCleanerFlavor {
    Write-Step "Creating CCleaner flavor"
    
    if (-not (Invoke-Winapp2ool -Arguments '-s', '-flavorize', '-autodetect', '-2f', 'winapp2-ccleaner-flavor.ini', `
        '-9d', '\CCleaner' -ErrorMessage "Failed to create CCleaner flavor")) { return $false }
    
    Write-Step "Creating BleachBit flavor"
    if (-not (Invoke-Winapp2ool -Arguments '-s', '-flavorize', '-autodetect', '-2f', 'winapp2-bleachbit-flavor.ini', `
        '-9d', '\BleachBit' -ErrorMessage "Failed to create BleachBit flavor")) { return $false }
    
    Write-Step "Creating Tron flavor"
    if (-not (Invoke-Winapp2ool -Arguments '-s', '-flavorize', '-autodetect', '-1f', 'winapp2-ccleaner-flavor.ini', `
        '-2f', 'winapp2-tron-flavor.ini', '-9d', '\Tron' `
        -ErrorMessage "Failed to create Tron flavor")) { return $false }
    
    Write-Step "Creating System Ninja flavor"
    if (-not (Invoke-Winapp2ool -Arguments '-s', '-flavorize', '-autodetect', '-2f', 'winapp2-systemninja-flavor.ini', `
        '-9d', '\SystemNinja' -ErrorMessage "Failed to create System Ninja flavor")) { return $false }
    
    Write-Step "Performing static analysis and saving corrections"
    if (-not (Invoke-Winapp2ool -Arguments '-s', '-debug', '-usedate', '-c', '-1f', 'winapp2-ccleaner-flavor.ini' `
        -ErrorMessage "Failed CCleaner flavor analysis")) { return $false }
    
    Write-Step "Creating changelog"
    if (-not (Invoke-Winapp2ool -Arguments '-s', '-diff', '-d', '-savelog', '-1f', 'winapp2-cc.old', `
        '-2f', 'Winapp2.ini', '-3f', 'diff.txt' `
        -ErrorMessage "Failed to create CCleaner changelog")) { return $false }
    
    Copy-Item 'Winapp2.ini' '..' -Force
    Copy-Item 'Winapp2.ini' 'winapp2-cc7base.ini' -Force
    Copy-Item 'diff.txt' '..' -Force
    Write-Host "CCleaner flavor built" -ForegroundColor Green
    
    return $true
}

function Build-CCleaner7Flavor {
    Write-Step "Creating CCleaner7 flavor"
    
    if (-not (Invoke-Winapp2ool -Arguments '-s', '-flavorize', '-cc7ify', '-2f', 'winapp2-cc7.ini' `
        -ErrorMessage "Failed to create CCleaner7 flavor")) { return $false }
    
    Write-Step "Creating changelog"
    if (-not (Invoke-Winapp2ool -Arguments '-s', '-diff', '-d', '-savelog', '-1f', 'winapp2-cc7.old', `
        '-2f', 'winapp2-cc7.ini', '-3f', 'diff.txt' `
        -ErrorMessage "Failed to create CCleaner7 changelog")) { return $false }
    
    Copy-Item 'winapp2-cc7.ini' '..\Non-CCleaner\CCleaner7\Winapp2.ini' -Force
    Copy-Item 'diff.txt' '..\Non-CCleaner\CCleaner7' -Force
    Write-Host "CCleaner7 flavor built" -ForegroundColor Green
    
    return $true
}

function Build-BleachBitFlavor {
    if (-not (Invoke-Winapp2ool -Arguments '-s', '-debug', '-usedate', '-c', '-1f', 'winapp2-bleachbit-flavor.ini' `
        -ErrorMessage "Failed BleachBit flavor analysis")) { return $false }
    
    Write-Step "Creating changelog"
    if (-not (Invoke-Winapp2ool -Arguments '-s', '-diff', '-d', '-savelog', '-1f', 'winapp2-bb.old', `
        '-2f', 'Winapp2.ini', '-3f', 'diff.txt' `
        -ErrorMessage "Failed to create BleachBit changelog")) { return $false }
    
    Copy-Item 'Winapp2.ini' '..\Non-CCleaner\BleachBit' -Force
    Copy-Item 'diff.txt' '..\Non-CCleaner\BleachBit' -Force
    Write-Host "BleachBit flavor built" -ForegroundColor Green
    
    return $true
}

function Build-TronFlavor {
    if (-not (Invoke-Winapp2ool -Arguments '-s', '-debug', '-usedate', '-c', '-1f', 'winapp2-tron-flavor.ini' `
        -ErrorMessage "Failed Tron flavor analysis")) { return $false }
    
    Write-Step "Creating changelog"
    if (-not (Invoke-Winapp2ool -Arguments '-s', '-diff', '-d', '-savelog', '-1f', 'winapp2-tron.old', `
        '-2f', 'Winapp2.ini', '-3f', 'diff.txt' `
        -ErrorMessage "Failed to create Tron changelog")) { return $false }
    
    Copy-Item 'Winapp2.ini' '..\Non-CCleaner\Tron' -Force
    Copy-Item 'diff.txt' '..\Non-CCleaner\Tron' -Force
    Write-Host "Tron flavor built" -ForegroundColor Green
    
    return $true
}

function Build-SystemNinjaFlavor {
    if (-not (Invoke-Winapp2ool -Arguments '-s', '-debug', '-usedate', '-c', '-1f', 'winapp2-systemninja-flavor.ini', `
        '-3f', 'Winapp2.rules' -ErrorMessage "Failed System Ninja flavor analysis")) { return $false }
    
    Write-Step "Creating changelog"
    if (-not (Invoke-Winapp2ool -Arguments '-s', '-diff', '-d', '-savelog', '-1f', 'winapp2-sn.old', `
        '-2f', 'Winapp2.rules', '-3f', 'diff.txt' `
        -ErrorMessage "Failed to create System Ninja changelog")) { return $false }
    
    Copy-Item 'Winapp2.rules' '..\Non-CCleaner\SystemNinja' -Force
    Copy-Item 'diff.txt' '..\Non-CCleaner\SystemNinja' -Force
    Write-Host "System Ninja flavor built" -ForegroundColor Green
    
    return $true
}

function Remove-TemporaryFiles {
    Write-Step "Cleaning up"
    
    $filesToRemove = @(
        'winapp2*.old',
        'winapp2*.ini',
        'base-entries-combined.ini',
        'browsers.ini',
        'uwp.ini',
        'winapp2.rules',
        'diff.txt'
    )
    
    foreach ($pattern in $filesToRemove) {
        Remove-Item $pattern -Force -ErrorAction SilentlyContinue
    }
}

# Main execution
try {
    $script:Winapp2oolPath = Find-Winapp2ool
    if (-not $script:Winapp2oolPath) {
        Write-ErrorMsg "Cannot continue without winapp2ool.exe"
        exit 1
    }
    
    Backup-Files
    
    if (-not (Build-MainFile)) { exit 1 }
    if (-not (Build-CCCleanerFlavor)) { exit 1 }
    if (-not (Build-CCleaner7Flavor)) { exit 1 }
    if (-not (Build-BleachBitFlavor)) { exit 1 }
    if (-not (Build-TronFlavor)) { exit 1 }
    if (-not (Build-SystemNinjaFlavor)) { exit 1 }
    
    Remove-TemporaryFiles
    
    Write-Host "`nWinapp2.ini successfully built" -ForegroundColor Green
    exit 0
}
catch {
    Write-ErrorMsg "Unexpected error: $_"
    exit 1
}