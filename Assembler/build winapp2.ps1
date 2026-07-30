<#
.SYNOPSIS
    Builds winapp2.ini and its various flavors using winapp2ool.

.DESCRIPTION
    This script requires winapp2ool v1.7 or newer
    Creates backups, regenerates the committed build artifacts under Entries\
    (base entries from EntryBuilder sources, browser entries into Entries\Browsers,
    UWP entries into Entries\UWP), merges them with a single strict-mode Combine
    pass, generates tool flavors, and creates a diff for each generated flavor.
    Generator stages read the shared scaffold catalogs from the Scaffolds directory.

    With -Verify, no build is performed: all artifacts are instead regenerated into
    a scratch directory and byte-compared against the committed Entries\ artifacts,
    exiting nonzero on any drift. This is the characteristic test used to prove the
    committed artifacts match their sources

.PARAMETER Verify
    Regenerate all artifacts into Verify\ and byte-compare them against the
    committed artifacts under Entries\ instead of building. Exits 0 when every
    artifact matches, 1 on any drift (the scratch output is retained for
    inspection on failure).

.PARAMETER GenerateTo
    Regenerate the three generator artifacts into the given root and stop -- no
    combine, no flavors, no comparison. The root is winapp2ool-relative (leading
    backslash, e.g. '\Parity\Source'). Paired with -ToolPath this is how CI runs
    two different winapp2ool binaries over the same sources and diffs the results.

.PARAMETER ToolPath
    Use this winapp2ool.exe instead of the one beside this script. Directory args
    resolve against the working directory, not the exe's location, so the binary
    may live anywhere as long as the script is run from Assembler\.

.NOTES
    Author: Hazel Ward
    Version 20260730
    Copyright 2026
#>

[CmdletBinding()]
param(
    [switch]$Verify,
    [string]$GenerateTo,
    [string]$ToolPath
)

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
    # An explicit -ToolPath wins over everything, including the prompt: an unattended
    # caller that named a binary wants that binary or a failure, never a question.
    if ($ToolPath) {
        if (-not (Test-Path $ToolPath -PathType Leaf)) {
            Write-ErrorMsg "-ToolPath is not a file: $ToolPath"
            return $null
        }
        $resolved = (Resolve-Path $ToolPath).Path
        Write-Host "Using winapp2ool at $resolved (-ToolPath)" -ForegroundColor Green
        return $resolved
    }

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
        [string]$ErrorMessage = "Winapp2ool command failed",
        [string[]]$RequiredFiles = @(),
        [string[]]$RequiredDirs = @()
    )

    foreach ($f in $RequiredFiles) {
        if (-not (Test-Path $f -PathType Leaf)) {
            Write-ErrorMsg "$ErrorMessage - required file not found: $f"
            return $false
        }
    }
    foreach ($d in $RequiredDirs) {
        if (-not (Test-Path $d -PathType Container)) {
            Write-ErrorMsg "$ErrorMessage - required directory not found: $d"
            return $false
        }
    }

    try {
        # Use -NoNewWindow to see the console output
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

function Get-ComparableContent {
    # Every lint pass restamps "; Version: <today>" via -usedate, so a rebuild whose sources did
    # not move still differs from the published file on exactly that line. Drop it and what
    # remains is the content we actually publish.
    param([string]$Path)

    return @(Get-Content -LiteralPath $Path | Where-Object { $_ -notmatch '^;\s*Version:' })
}

function Test-ContentChanged {
    <#
    .SYNOPSIS
        Answers whether a freshly built artifact differs from the one it would replace.

    .DESCRIPTION
        The gate for republishing. This is deliberately a TEXTUAL comparison and not Diff's
        verdict: Diff normalizes deprecated paths on both sides, cannot see comments outside the
        preamble, and ignores ordering, so it can report "no changes" for a file that genuinely
        changed. A false negative here would silently drop a real release, so the gate uses the
        one comparison that has no false negatives and Diff supplies the human-readable detail.

        A missing baseline counts as changed, the first build must always publish.
    #>
    param([string]$Built, [string]$Previous)

    if (-not (Test-Path $Previous)) { return $true }

    $new = Get-ComparableContent $Built
    $old = Get-ComparableContent $Previous

    if ($new.Count -ne $old.Count) { return $true }

    for ($i = 0; $i -lt $new.Count; $i++) {
        if ($new[$i] -cne $old[$i]) { return $true }
    }

    return $false
}

function Invoke-Changelog {
    <#
    .SYNOPSIS
        Runs Diff over the old and new artifacts, writing the changelog and reporting what it found.

    .DESCRIPTION
        -4f asks Diff for its outcome as key=value lines alongside the prose changelog, which is
        the only way this script can see inside a run it launched with a hidden window. The counts
        are echoed into the build log so an unattended build says what changed, not merely that
        something did.
    #>
    param([string]$Old, [string]$New, [string]$LogName = 'diff.txt')

    $summary = 'diffsummary.txt'
    Remove-Item $summary -Force -ErrorAction SilentlyContinue

    if (-not (Invoke-Winapp2ool -Arguments '-s', '-offline', '-diff', '-savelog', '-1f', $Old, `
        '-2f', $New, '-3f', $LogName, '-4f', $summary `
        -ErrorMessage "Failed to create changelog for $New")) { return $false }

    if (Test-Path $summary) {
        $s = ConvertFrom-StringData (Get-Content -LiteralPath $summary -Raw)
        Write-Host ("  changes: {0} added, {1} removed, {2} modified entries; {3} added, {4} removed, {5} updated, {6} moved keys" -f `
            $s.addedentries, $s.removedentries, $s.modifiedentries, `
            $s.addedkeys, $s.removedkeys, $s.updatedkeys, $s.movedkeys) -ForegroundColor Gray

        # The textual gate already said the file changed. Diff disagreeing means the delta is
        # something Diff normalizes away. never worth suppressing the release.
        if ($s.haschanges -eq 'false') {
            Write-Host "  note: the file changed but Diff reports no semantic changes (normalized paths, comments, or ordering)" -ForegroundColor Yellow
        }
    }

    return $true
}

function Backup-Files {
    Write-Step "Creating backups of existing winapp2.ini files"

    $backups = @(
        @{Source = "..\Winapp2.ini"; Dest = "winapp2-cc.old"}
        @{Source = "..\Non-CCleaner\Winapp2.ini"; Dest = "winapp2.old"}
        @{Source = "..\Non-CCleaner\BleachBit\Winapp2.ini"; Dest = "winapp2-bb.old"}
        @{Source = "..\Non-CCleaner\CCleaner7\Winapp2.ini"; Dest = "winapp2-cc7.old"}
        @{Source = "..\Non-CCleaner\FluentCleaner\Winapp2.ini"; Dest = "winapp2-fc.old"}
        @{Source = "..\Non-CCleaner\SystemNinja\Winapp2.rules"; Dest = "winapp2-sn.old"}
        @{Source = "..\Non-CCleaner\Tron\Winapp2.ini"; Dest = "winapp2-tron.old"}
    )

    foreach ($backup in $backups) {
        if (Test-Path $backup.Source) {
            Copy-Item $backup.Source $backup.Dest -Force -ErrorAction SilentlyContinue
        }
    }
}

function Build-Artifacts {
    param([string]$OutputRoot)

    Write-Step "Generating browser entries"
    if (-not (Invoke-Winapp2ool -Arguments '-s', '-offline', '-browserbuilder', '-1d', '\BrowserBuilder', `
        '-2d', "$OutputRoot\Browsers", '-2f', 'browsers.ini' `
        -ErrorMessage "Failed to generate browser entries" -RequiredDirs @(Join-Path $PSScriptRoot 'BrowserBuilder'))) { return $false }

    Write-Step "Generating UWP entries"
    if (-not (Invoke-Winapp2ool -Arguments '-s', '-offline', '-uwpbuilder', '-1d', '\UWP', `
        '-2d', "$OutputRoot\UWP", '-2f', 'uwp.ini', `
        '-3d', '\Scaffolds' `
        -ErrorMessage "Failed to generate UWP entries" `
        -RequiredDirs @((Join-Path $PSScriptRoot 'UWP'), (Join-Path $PSScriptRoot 'Scaffolds')))) { return $false }

    Write-Step "Generating base entries from EntryBuilder sources"
    if (-not (Invoke-Winapp2ool -Arguments '-s', '-offline', '-entrybuilder', '-split', '-1d', '\EntryBuilder', `
        '-2d', $OutputRoot, `
        '-3d', '\Scaffolds' `
        -ErrorMessage "Failed to generate EntryBuilder entries" `
        -RequiredDirs @((Join-Path $PSScriptRoot 'EntryBuilder'), (Join-Path $PSScriptRoot 'Scaffolds')))) { return $false }

    return $true
}

function Build-MainFile {
    Write-Step "Building winapp2.ini, please wait"

    if (-not (Build-Artifacts -OutputRoot '\Entries')) { return $false }

    Write-Step "Merging entry artifacts"
    if (-not (Invoke-Winapp2ool -Arguments '-s', '-offline', '-combine', '-strict', '-1d', '\Entries', '-3f', 'Winapp2.ini' `
        -ErrorMessage "Failed to merge entry artifacts" -RequiredDirs @(Join-Path $PSScriptRoot 'Entries'))) { return $false }

    Write-Step "Performing static analysis and saving corrections"
    if (-not (Invoke-Winapp2ool -Arguments '-s', '-offline', '-debug', '-usedate', '-c', '-opti', '-1f', 'Winapp2.ini', '-3f', 'Winapp2.ini' `
        -ErrorMessage "Failed static analysis")) { return $false }

    if (-not (Test-ContentChanged 'Winapp2.ini' 'winapp2.old')) {
        Write-Host "Base winapp2.ini is unchanged since the last build - not republishing" -ForegroundColor Yellow
        return $true
    }

    Write-Step "Creating changelog"
    $haveDiff = Test-Path 'winapp2.old'
    if ($haveDiff) {
        if (-not (Invoke-Changelog -Old 'winapp2.old' -New 'Winapp2.ini')) { return $false }
    } else {
        Write-Host "No previous build found, skipping changelog" -ForegroundColor Yellow
    }

    Copy-Item 'Winapp2.ini' '..\Non-CCleaner' -Force
    if ($haveDiff) { Copy-Item 'diff.txt' '..\Non-CCleaner' -Force }
    Write-Host "Winapp2.ini built" -ForegroundColor Green

    return $true
}

function Build-CCCleanerFlavor {
    Write-Step "Creating CCleaner flavor"

    if (-not (Invoke-Winapp2ool -Arguments '-s', '-offline', '-flavorize', '-autodetect', '-2f', 'winapp2-ccleaner-flavor.ini', `
        '-9d', '\CCleaner' -ErrorMessage "Failed to create CCleaner flavor" -RequiredDirs @(Join-Path $PSScriptRoot 'CCleaner'))) { return $false }

    Write-Step "Creating BleachBit flavor"
    if (-not (Invoke-Winapp2ool -Arguments '-s', '-offline', '-flavorize', '-autodetect', '-2f', 'winapp2-bleachbit-flavor.ini', `
        '-9d', '\BleachBit' -ErrorMessage "Failed to create BleachBit flavor" -RequiredDirs @(Join-Path $PSScriptRoot 'BleachBit'))) { return $false }

    Write-Step "Creating Tron flavor"
    if (-not (Invoke-Winapp2ool -Arguments '-s', '-offline', '-flavorize', '-autodetect', '-1f', 'winapp2-ccleaner-flavor.ini', `
        '-2f', 'winapp2-tron-flavor.ini', '-9d', '\Tron' `
        -ErrorMessage "Failed to create Tron flavor" -RequiredDirs @(Join-Path $PSScriptRoot 'Tron'))) { return $false }

    Write-Step "Creating System Ninja flavor"
    if (-not (Invoke-Winapp2ool -Arguments '-s', '-offline', '-flavorize', '-autodetect', '-2f', 'winapp2-systemninja-flavor.ini', `
        '-9d', '\SystemNinja' -ErrorMessage "Failed to create System Ninja flavor" -RequiredDirs @(Join-Path $PSScriptRoot 'SystemNinja'))) { return $false }

    Write-Step "Performing static analysis and saving corrections"
    if (-not (Invoke-Winapp2ool -Arguments '-s', '-offline', '-debug', '-usedate', '-c', '-opti', '-1f', 'winapp2-ccleaner-flavor.ini', `
        '-3f', 'winapp2-ccleaner-flavor.ini' `
        -ErrorMessage "Failed CCleaner flavor analysis")) { return $false }

    if (-not (Test-ContentChanged 'winapp2-ccleaner-flavor.ini' 'winapp2-cc.old')) {
        Write-Host "CCleaner flavor is unchanged since the last build - not republishing" -ForegroundColor Yellow
        return $true
    }

    Write-Step "Creating changelog"
    $haveDiff = Test-Path 'winapp2-cc.old'
    if ($haveDiff) {
        if (-not (Invoke-Changelog -Old 'winapp2-cc.old' -New 'winapp2-ccleaner-flavor.ini')) { return $false }
    } else {
        Write-Host "No previous build found, skipping changelog" -ForegroundColor Yellow
    }

    Copy-Item 'winapp2-ccleaner-flavor.ini' '..\Winapp2.ini' -Force
    if ($haveDiff) { Copy-Item 'diff.txt' '..' -Force }
    Write-Host "CCleaner flavor built" -ForegroundColor Green

    return $true
}

function Build-CCleaner7Flavor {
    Write-Step "Creating CCleaner7 flavor"

    if (-not (Invoke-Winapp2ool -Arguments '-s', '-offline', '-flavorize', '-autodetect', '-1f', 'winapp2-ccleaner-flavor.ini', `
        '-2f', 'winapp2-cc7.ini', '-9d', '\CCleaner7' `
        -ErrorMessage "Failed to create CCleaner7 flavor" -RequiredDirs @(Join-Path $PSScriptRoot 'CCleaner7'))) { return $false }

    if (-not (Test-ContentChanged 'winapp2-cc7.ini' 'winapp2-cc7.old')) {
        Write-Host "CCleaner7 flavor is unchanged since the last build - not republishing" -ForegroundColor Yellow
        return $true
    }

    Write-Step "Creating changelog"
    $haveDiff = Test-Path 'winapp2-cc7.old'
    if ($haveDiff) {
        if (-not (Invoke-Changelog -Old 'winapp2-cc7.old' -New 'winapp2-cc7.ini')) { return $false }
    } else {
        Write-Host "No previous build found, skipping changelog" -ForegroundColor Yellow
    }

    Copy-Item 'winapp2-cc7.ini' '..\Non-CCleaner\CCleaner7\Winapp2.ini' -Force
    if ($haveDiff) { Copy-Item 'diff.txt' '..\Non-CCleaner\CCleaner7' -Force }
    Write-Host "CCleaner7 flavor built" -ForegroundColor Green

    return $true
}

function Build-FluentCleanerFlavor {
    Write-Step "Creating FluentCleaner flavor"

    # FluentCleaner is designed around the CCleaner flavor, so we'll make some
    # changes to the base file that are pulled from the downstream 
    # FluentCleaner winapp2.ini and then apply the CCleaner flavor on top of that 
    if (-not (Invoke-Winapp2ool -Arguments '-s', '-offline', '-flavorize', '-autodetect', '-1f', 'Winapp2.ini', `
        '-2f', 'winapp2-fluentcleaner-pre.ini', '-9d', '\FluentCleaner' `
        -ErrorMessage "Failed to create FluentCleaner pre-pass" -RequiredDirs @(Join-Path $PSScriptRoot 'FluentCleaner'))) { return $false }

    if (-not (Invoke-Winapp2ool -Arguments '-s', '-offline', '-flavorize', '-autodetect', '-1f', 'winapp2-fluentcleaner-pre.ini', `
        '-2f', 'winapp2-fluentcleaner.ini', '-9d', '\CCleaner' `
        -ErrorMessage "Failed to categorize FluentCleaner flavor" -RequiredDirs @(Join-Path $PSScriptRoot 'CCleaner'))) { return $false }

    # -keepdefaults preserves FluentCleaner's Default=False keys through the linter 
    Write-Step "Performing static analysis and saving corrections"
    if (-not (Invoke-Winapp2ool -Arguments '-s', '-offline', '-debug', '-usedate', '-c', '-opti', '-keepdefaults', '-1f', 'winapp2-fluentcleaner.ini', `
        '-3f', 'winapp2-fluentcleaner.ini' `
        -ErrorMessage "Failed FluentCleaner flavor analysis")) { return $false }

    if (-not (Test-ContentChanged 'winapp2-fluentcleaner.ini' 'winapp2-fc.old')) {
        Write-Host "FluentCleaner flavor is unchanged since the last build - not republishing" -ForegroundColor Yellow
        return $true
    }

    Write-Step "Creating changelog"
    $haveDiff = Test-Path 'winapp2-fc.old'
    if ($haveDiff) {
        if (-not (Invoke-Changelog -Old 'winapp2-fc.old' -New 'winapp2-fluentcleaner.ini')) { return $false }
    } else {
        Write-Host "No previous build found, skipping changelog" -ForegroundColor Yellow
    }

    $publishDir = '..\Non-CCleaner\FluentCleaner'
    if (-not (Test-Path $publishDir)) { New-Item -ItemType Directory -Force -Path $publishDir | Out-Null }
    Copy-Item 'winapp2-fluentcleaner.ini' "$publishDir\Winapp2.ini" -Force
    if ($haveDiff) { Copy-Item 'diff.txt' $publishDir -Force }
    Write-Host "FluentCleaner flavor built" -ForegroundColor Green

    return $true
}

function Build-BleachBitFlavor {
    if (-not (Invoke-Winapp2ool -Arguments '-s', '-offline', '-debug', '-usedate', '-c', '-opti', '-1f', 'winapp2-bleachbit-flavor.ini', `
        '-3f', 'winapp2-bleachbit-flavor.ini' `
        -ErrorMessage "Failed BleachBit flavor analysis")) { return $false }

    if (-not (Test-ContentChanged 'winapp2-bleachbit-flavor.ini' 'winapp2-bb.old')) {
        Write-Host "BleachBit flavor is unchanged since the last build - not republishing" -ForegroundColor Yellow
        return $true
    }

    Write-Step "Creating changelog"
    $haveDiff = Test-Path 'winapp2-bb.old'
    if ($haveDiff) {
        if (-not (Invoke-Changelog -Old 'winapp2-bb.old' -New 'winapp2-bleachbit-flavor.ini')) { return $false }
    } else {
        Write-Host "No previous build found, skipping changelog" -ForegroundColor Yellow
    }

    Copy-Item 'winapp2-bleachbit-flavor.ini' '..\Non-CCleaner\BleachBit\Winapp2.ini' -Force
    if ($haveDiff) { Copy-Item 'diff.txt' '..\Non-CCleaner\BleachBit' -Force }
    Write-Host "BleachBit flavor built" -ForegroundColor Green

    return $true
}

function Build-TronFlavor {
    if (-not (Invoke-Winapp2ool -Arguments '-s', '-offline', '-debug', '-usedate', '-c', '-opti', '-1f', 'winapp2-tron-flavor.ini', `
        '-3f', 'winapp2-tron-flavor.ini' `
        -ErrorMessage "Failed Tron flavor analysis")) { return $false }

    if (-not (Test-ContentChanged 'winapp2-tron-flavor.ini' 'winapp2-tron.old')) {
        Write-Host "Tron flavor is unchanged since the last build - not republishing" -ForegroundColor Yellow
        return $true
    }

    Write-Step "Creating changelog"
    $haveDiff = Test-Path 'winapp2-tron.old'
    if ($haveDiff) {
        if (-not (Invoke-Changelog -Old 'winapp2-tron.old' -New 'winapp2-tron-flavor.ini')) { return $false }
    } else {
        Write-Host "No previous build found, skipping changelog" -ForegroundColor Yellow
    }

    Copy-Item 'winapp2-tron-flavor.ini' '..\Non-CCleaner\Tron\Winapp2.ini' -Force
    if ($haveDiff) { Copy-Item 'diff.txt' '..\Non-CCleaner\Tron' -Force }
    Write-Host "Tron flavor built" -ForegroundColor Green

    return $true
}

function Build-SystemNinjaFlavor {
    if (-not (Invoke-Winapp2ool -Arguments '-s', '-offline', '-debug', '-usedate', '-c', '-opti', '-1f', 'winapp2-systemninja-flavor.ini', `
        '-3f', 'Winapp2.rules' -ErrorMessage "Failed System Ninja flavor analysis")) { return $false }

    if (-not (Test-ContentChanged 'Winapp2.rules' 'winapp2-sn.old')) {
        Write-Host "System Ninja flavor is unchanged since the last build - not republishing" -ForegroundColor Yellow
        return $true
    }

    Write-Step "Creating changelog"
    $haveDiff = Test-Path 'winapp2-sn.old'
    if ($haveDiff) {
        if (-not (Invoke-Changelog -Old 'winapp2-sn.old' -New 'Winapp2.rules')) { return $false }
    } else {
        Write-Host "No previous build found, skipping changelog" -ForegroundColor Yellow
    }

    Copy-Item 'Winapp2.rules' '..\Non-CCleaner\SystemNinja' -Force
    if ($haveDiff) { Copy-Item 'diff.txt' '..\Non-CCleaner\SystemNinja' -Force }
    Write-Host "System Ninja flavor built" -ForegroundColor Green

    return $true
}

function Test-Artifacts {
    Write-Step "Regenerating artifacts into scratch for characteristic testing"

    $scratch = Join-Path $PSScriptRoot 'Verify'
    if (Test-Path $scratch) { Remove-Item $scratch -Recurse -Force }

    if (-not (Build-Artifacts -OutputRoot '\Verify')) { return $false }

    Write-Step "Comparing regenerated artifacts against committed Entries"

    $entriesDir = Join-Path $PSScriptRoot 'Entries'
    $drift = @()

    foreach ($file in Get-ChildItem $entriesDir -Filter '*.ini' -Recurse -File) {
        $rel = $file.FullName.Substring($entriesDir.Length + 1)
        $counterpart = Join-Path $scratch $rel
        if (-not (Test-Path $counterpart)) {
            $drift += "missing from regeneration: $rel"
            continue
        }
        if ((Get-FileHash $file.FullName).Hash -ne (Get-FileHash $counterpart).Hash) {
            $drift += "content drift: $rel"
        }
    }

    foreach ($file in Get-ChildItem $scratch -Filter '*.ini' -Recurse -File) {
        $rel = $file.FullName.Substring($scratch.Length + 1)
        if (-not (Test-Path (Join-Path $entriesDir $rel))) {
            $drift += "not present in committed artifacts: $rel"
        }
    }

    if ($drift.Count -gt 0) {
        Write-ErrorMsg "Artifact drift detected - committed Entries do not match their sources"
        $drift | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }
        Write-Host "Regenerated artifacts retained in $scratch for inspection" -ForegroundColor Yellow
        return $false
    }

    Remove-Item $scratch -Recurse -Force
    Write-Host "All committed artifacts match their sources" -ForegroundColor Green
    return $true
}

function Remove-TemporaryFiles {
    Write-Step "Cleaning up"

    $filesToRemove = @(
        'winapp2*.old',
        'winapp2*.ini',
        'winapp2.rules',
        'diff.txt',
        'diffsummary.txt'
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

    # Generate-and-stop. Deliberately does NOT compare against Entries\: the caller is
    # diffing two generated trees against each other, so the committed artifacts are
    # not part of the question and their staleness must not fail the run.
    if ($GenerateTo) {
        $target = Join-Path $PSScriptRoot $GenerateTo.TrimStart('\')
        if (Test-Path $target) { Remove-Item $target -Recurse -Force }
        if (-not (Build-Artifacts -OutputRoot $GenerateTo)) { exit 1 }
        Write-Host "`nArtifacts generated into $target" -ForegroundColor Green
        exit 0
    }

    if ($Verify) {
        if (-not (Test-Artifacts)) { exit 1 }
        Write-Host "`nCharacteristic test passed" -ForegroundColor Green
        exit 0
    }

    Backup-Files

    if (-not (Build-MainFile)) { exit 1 }
    if (-not (Build-CCCleanerFlavor)) { exit 1 }
    if (-not (Build-CCleaner7Flavor)) { exit 1 }
    if (-not (Build-FluentCleanerFlavor)) { exit 1 }
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
