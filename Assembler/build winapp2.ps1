<#
.SYNOPSIS
    Builds winapp2.ini and its various flavors using winapp2ool.

.DESCRIPTION
    This script requires winapp2ool v1.8.9821.32740 or newer, and refuses to run an older one:
    Creates backups, regenerates the committed build artifacts under Generated\
    (base entries, Store apps included, from Sources\EntryBuilder into Generated\EntryBuilder,
    and browser entries into Generated\BrowserBuilder), merges them with a single strict-mode
    Combine pass, generates tool flavors, and creates a diff for each generated flavor.
    Each flavor and its changelog are published to ..\Releases under their release asset names,
    then copied to their old paths for the deprecation window ($script:DeprecatedPublishPaths).
    The published files, their changelogs and their old paths are rows of
    ..\.github\scripts\flavors.csv, loaded through release-assets.ps1 beside it.
    Generator stages read the shared scaffold catalogs from the Sources\Scaffolds directory.

    Entries marked Archive= in the sources go to Archived\ instead, one archived.ini per
    builder. A second strict Combine joins those into archived.ini, which is linted with
    -archive, checked against winapp2.ini for shared names, and published with its own
    changelog like a flavor.

    With -Verify, no build is performed: all artifacts are instead regenerated into
    scratch directories and byte-compared against the committed Generated\ and Archived\ artifacts,
    exiting nonzero on any drift. This is the characteristic test used to prove the
    committed artifacts match their sources

.PARAMETER Verify
    Regenerate all artifacts into Verify\Generated\ and Verify\Archived\ and byte-compare
    each against its committed tree (Generated\, Archived\) instead of building. Exits 0
    when every artifact matches, 1 on any drift (the scratch output is retained for
    inspection on failure).

.PARAMETER GenerateTo
    Regenerate the two generators' artifacts (EntryBuilder and BrowserBuilder) into
    <root>\Generated\ and <root>\Archived\ and stop -- no combine, no flavors, no comparison. The root is winapp2ool-relative
    (leading backslash, e.g. '\Parity\Source'), and '\' regenerates the committed trees.
    Paired with -ToolPath this is how CI runs two different winapp2ool binaries over the
    same sources and diffs the results. Only the builders' folders under the two trees
    are cleared first; anything else is left alone. A root inside Generated\, Archived\
    or Sources\ is refused, since the build reads those trees.

.PARAMETER ToolPath
    Use this winapp2ool.exe instead of the one beside this script. Directory args
    resolve against the working directory, not the exe's location, so the binary
    may live anywhere as long as the script is run from Assembler\.

.NOTES
    Author: Hazel Ward
    Version 20261121
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
$script:MinimumToolVersion = [version]'1.8.9821.40534'

function Write-Step {
    param([string]$Message)
    # Adding a clear newline before each major step
    Write-Host "`n>> $Message" -ForegroundColor Cyan
}

function Write-ErrorMsg {
    param([string]$Message)
    Write-Host "ERROR: $Message" -ForegroundColor Red
}

function Test-Winapp2oolVersion {
    # Older builds exit 0 on a failed read or write, which every stage here trusts as success
    param([string]$Path)

    $raw = (Get-Item -LiteralPath $Path).VersionInfo.FileVersion
    $version = $null
    if (-not [version]::TryParse("$raw", [ref]$version)) {
        Write-ErrorMsg "Could not read the version of $Path"
        return $false
    }

    if ($version -lt $script:MinimumToolVersion) {
        Write-ErrorMsg "$Path is winapp2ool v$version, but this script requires v$($script:MinimumToolVersion) or newer"
        return $false
    }

    return $true
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
    Write-Host "This script requires winapp2ool v$($script:MinimumToolVersion) or newer`n" -ForegroundColor Yellow

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

# ---- Deprecation window: the old published paths --------------------------------------------
# Until the window closes, each flavor is copied to the OldData/OldChangelog paths in flavors.csv
# after it is published to ..\Releases. Nothing reads these copies: the change gate and the
# changelog compare against ..\Releases only. Closing the window deletes this function, its call
# in the main block, the block in Publish-Flavor that reads the table, and the two CSV columns.
function Set-DeprecatedPublishPaths {
    # flavors.csv holds repo paths with forward slashes, and an empty cell means no old copy
    $script:DeprecatedPublishPaths = @{}
    foreach ($row in $PublishedFiles) {
        $changelog = if ($row.OldChangelog) { '..\' + $row.OldChangelog.Replace('/', '\') } else { $null }
        $script:DeprecatedPublishPaths[$row.Asset] = @{ Data = '..\' + $row.OldData.Replace('/', '\'); Changelog = $changelog }
    }
}
# ---- End of the deprecation window -----------------------------------------------------------

function Publish-Flavor {
    <#
    .SYNOPSIS
        Publishes a built flavor to ..\Releases under its release asset name, with its changelog
        when there is one, then copies the same bytes to the flavor's old paths.
    #>
    param([string]$Built, [string]$DataFile, [string]$Changelog, [bool]$HaveDiff)

    $releases = '..\Releases'
    if (-not (Test-Path $releases)) { New-Item -ItemType Directory -Force -Path $releases | Out-Null }

    Copy-Item $Built "$releases\$DataFile" -Force
    if ($HaveDiff) { Copy-Item 'diff.txt' "$releases\$Changelog" -Force }

    # Deprecation window
    $old = $script:DeprecatedPublishPaths[$DataFile]
    $oldDir = Split-Path $old.Data
    if (-not (Test-Path $oldDir)) { New-Item -ItemType Directory -Force -Path $oldDir | Out-Null }
    Copy-Item "$releases\$DataFile" $old.Data -Force
    if ($HaveDiff -and $old.Changelog) { Copy-Item "$releases\$Changelog" $old.Changelog -Force }
}

function Backup-Files {
    Write-Step "Creating backups of existing winapp2.ini files"

    foreach ($row in $PublishedFiles) {
        $source = "..\Releases\$($row.Asset)"
        $baseline = "$($row.Asset).old"

        # A baseline left over from an interrupted run would stand in for a missing published file
        Remove-Item -LiteralPath $baseline -Force -ErrorAction SilentlyContinue

        if (Test-Path -LiteralPath $source) {
            Copy-Item -LiteralPath $source $baseline -Force -ErrorAction SilentlyContinue
        }
    }
}

function Complete-Flavor {
    <#
    .SYNOPSIS
        Publishes a built file under its flavors.csv row when it differs from the last published
        one, with a changelog against that baseline (<Asset>.old, from Backup-Files).
    #>
    param([string]$Key, [string]$Built)

    $row = @($PublishedFiles | Where-Object { $_.Key -eq $Key })
    if ($row.Count -ne 1) {
        Write-ErrorMsg "flavors.csv has $($row.Count) rows for $Key, expected one"
        return $false
    }
    $row = $row[0]
    $baseline = "$($row.Asset).old"

    if (-not (Test-ContentChanged $Built $baseline)) {
        Write-Host "$($row.Label) is unchanged since the last build - not republishing" -ForegroundColor Yellow
        return $true
    }

    Write-Step "Creating changelog"
    $haveDiff = Test-Path -LiteralPath $baseline
    if ($haveDiff) {
        if (-not (Invoke-Changelog -Old $baseline -New $Built)) { return $false }
    } else {
        Write-Host "No previous build found, skipping changelog" -ForegroundColor Yellow
    }

    Publish-Flavor -Built $Built -DataFile $row.Asset -Changelog $row.Changelog -HaveDiff $haveDiff
    Write-Host "$($row.Label) built" -ForegroundColor Green

    return $true
}

function Build-Artifacts {
    # Each builder writes its live output under <Root>\Generated and its archived.ini under
    # <Root>\Archived. Root is winapp2ool-relative, and '' is the committed trees.
    param([string]$Root = '')

    $OutputRoot = "$Root\Generated"
    $ArchiveRoot = "$Root\Archived"

    Write-Step "Generating browser entries"
    if (-not (Invoke-Winapp2ool -Arguments '-s', '-offline', '-browserbuilder', '-1d', '\Sources\BrowserBuilder', `
        '-2d', "$OutputRoot\BrowserBuilder", '-2f', 'browsers.ini', '-3d', "$ArchiveRoot\BrowserBuilder" `
        -ErrorMessage "Failed to generate browser entries" -RequiredDirs @(Join-Path $PSScriptRoot 'Sources\BrowserBuilder'))) { return $false }

    Write-Step "Generating base entries from EntryBuilder sources"
    if (-not (Invoke-Winapp2ool -Arguments '-s', '-offline', '-entrybuilder', '-split', '-1d', '\Sources\EntryBuilder', `
        '-2d', "$OutputRoot\EntryBuilder", `
        '-3d', '\Sources\Scaffolds', '-4d', "$ArchiveRoot\EntryBuilder" `
        -ErrorMessage "Failed to generate EntryBuilder entries" `
        -RequiredDirs @((Join-Path $PSScriptRoot 'Sources\EntryBuilder'), (Join-Path $PSScriptRoot 'Sources\Scaffolds')))) { return $false }

    return $true
}

function Build-MainFile {
    Write-Step "Building winapp2.ini, please wait"

    if (-not (Build-Artifacts)) { return $false }

    Write-Step "Merging entry artifacts"
    if (-not (Invoke-Winapp2ool -Arguments '-s', '-offline', '-combine', '-strict', '-1d', '\Generated', '-3f', 'Winapp2.ini' `
        -ErrorMessage "Failed to merge entry artifacts" -RequiredDirs @(Join-Path $PSScriptRoot 'Generated'))) { return $false }

    Write-Step "Performing static analysis and saving corrections"
    if (-not (Invoke-Winapp2ool -Arguments '-s', '-offline', '-base', '-debug', '-usedate', '-c', '-opti', '-1f', 'Winapp2.ini', '-3f', 'Winapp2.ini' `
        -ErrorMessage "Failed static analysis")) { return $false }

    return (Complete-Flavor -Key Base -Built 'Winapp2.ini')
}

function Get-SectionNames {
    param([string]$Path)

    return @(Get-Content -LiteralPath $Path | Where-Object { $_ -match '^\[(.+)\]\s*$' } | ForEach-Object { $Matches[1] })
}

function Test-ArchiveNames {
    <#
    .SYNOPSIS
        Fails when an entry name is both live and archived.

    .DESCRIPTION
        EntryBuilder catches a live and an archived entry of one name inside its own sources,
        but nothing else sees across builders: an archived EntryBuilder entry named like a
        generated browser entry would otherwise ship in both files.
    #>
    param([string]$Live, [string]$Archive)

    $liveNames = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($name in Get-SectionNames $Live) { [void]$liveNames.Add($name) }

    $both = @(Get-SectionNames $Archive | Where-Object { $liveNames.Contains($_) })
    if ($both.Count -eq 0) { return $true }

    Write-ErrorMsg "$($both.Count) entry name(s) are in both $Live and $Archive. An entry is either live or archived:"
    $both | ForEach-Object { Write-Host "  [$_]" -ForegroundColor Red }
    return $false
}

function Build-Archive {
    Write-Step "Building archived.ini"

    Remove-Item 'archived.ini' -Force -ErrorAction SilentlyContinue

    # Build-MainFile's generators already wrote one archived.ini per builder under Archived\
    if (-not (Invoke-Winapp2ool -Arguments '-s', '-offline', '-combine', '-strict', '-1d', '\Archived', '-3f', 'archived.ini' `
        -ErrorMessage "Failed to merge the archive artifacts" -RequiredDirs @(Join-Path $PSScriptRoot 'Archived'))) { return $false }

    # Combine exits 0 when it finds nothing to combine
    if (-not (Test-Path 'archived.ini')) {
        Write-ErrorMsg "Merging the archive artifacts produced no archived.ini"
        return $false
    }

    Write-Step "Performing static analysis on the archive and saving corrections"
    if (-not (Invoke-Winapp2ool -Arguments '-s', '-offline', '-debug', '-archive', '-usedate', '-c', '-opti', '-1f', 'archived.ini', '-3f', 'archived.ini' `
        -ErrorMessage "Failed archive analysis")) { return $false }

    Write-Step "Checking that no entry is both live and archived"
    if (-not (Test-ArchiveNames -Live 'Winapp2.ini' -Archive 'archived.ini')) { return $false }

    return (Complete-Flavor -Key Archive -Built 'archived.ini')
}

function Build-CCCleanerFlavor {
    Write-Step "Creating CCleaner flavor"

    if (-not (Invoke-Winapp2ool -Arguments '-s', '-offline', '-cc', '-flavorize', '-autodetect', '-2f', 'winapp2-ccleaner-flavor.ini', `
        '-9d', '\Sources\Flavors\CCleaner' -ErrorMessage "Failed to create CCleaner flavor" -RequiredDirs @(Join-Path $PSScriptRoot 'Sources\Flavors\CCleaner'))) { return $false }

    Write-Step "Creating BleachBit flavor"
    if (-not (Invoke-Winapp2ool -Arguments '-s', '-offline', '-bb', '-flavorize', '-autodetect', '-2f', 'winapp2-bleachbit-flavor.ini', `
        '-9d', '\Sources\Flavors\BleachBit' -ErrorMessage "Failed to create BleachBit flavor" -RequiredDirs @(Join-Path $PSScriptRoot 'Sources\Flavors\BleachBit'))) { return $false }

    Write-Step "Creating System Ninja flavor"
    if (-not (Invoke-Winapp2ool -Arguments '-s', '-offline', '-sn', '-flavorize', '-autodetect', '-2f', 'winapp2-systemninja-flavor.ini', `
        '-9d', '\Sources\Flavors\SystemNinja' -ErrorMessage "Failed to create System Ninja flavor" -RequiredDirs @(Join-Path $PSScriptRoot 'Sources\Flavors\SystemNinja'))) { return $false }

    Write-Step "Performing static analysis and saving corrections"
    if (-not (Invoke-Winapp2ool -Arguments '-s', '-offline', '-cc', '-debug', '-usedate', '-c', '-opti', '-1f', 'winapp2-ccleaner-flavor.ini', `
        '-3f', 'winapp2-ccleaner-flavor.ini' `
        -ErrorMessage "Failed CCleaner flavor analysis")) { return $false }

    return (Complete-Flavor -Key CCleaner -Built 'winapp2-ccleaner-flavor.ini')
}

function Build-CCleaner7Flavor {
    Write-Step "Creating CCleaner7 flavor"

    if (-not (Invoke-Winapp2ool -Arguments '-s', '-offline', '-cc7', '-flavorize', '-autodetect', '-1f', 'winapp2-ccleaner-flavor.ini', `
        '-2f', 'winapp2-cc7.ini', '-9d', '\Sources\Flavors\CCleaner7' `
        -ErrorMessage "Failed to create CCleaner7 flavor" -RequiredDirs @(Join-Path $PSScriptRoot 'Sources\Flavors\CCleaner7'))) { return $false }

    return (Complete-Flavor -Key CCleaner7 -Built 'winapp2-cc7.ini')
}

function Build-FluentCleanerFlavor {
    Write-Step "Creating FluentCleaner flavor"

    # FluentCleaner is designed around the CCleaner flavor, so we'll make some
    # changes to the base file that are pulled from the downstream 
    # FluentCleaner winapp2.ini and then apply the CCleaner flavor on top of that 
    if (-not (Invoke-Winapp2ool -Arguments '-s', '-offline', '-fc', '-flavorize', '-autodetect', '-1f', 'Winapp2.ini', `
        '-2f', 'winapp2-fluentcleaner-pre.ini', '-9d', '\Sources\Flavors\FluentCleaner' `
        -ErrorMessage "Failed to create FluentCleaner pre-pass" -RequiredDirs @(Join-Path $PSScriptRoot 'Sources\Flavors\FluentCleaner'))) { return $false }

    if (-not (Invoke-Winapp2ool -Arguments '-s', '-offline', '-fc', '-flavorize', '-autodetect', '-1f', 'winapp2-fluentcleaner-pre.ini', `
        '-2f', 'winapp2-fluentcleaner.ini', '-9d', '\Sources\Flavors\CCleaner' `
        -ErrorMessage "Failed to categorize FluentCleaner flavor" -RequiredDirs @(Join-Path $PSScriptRoot 'Sources\Flavors\CCleaner'))) { return $false }

    # -keepdefaults preserves FluentCleaner's Default=False keys through the linter 
    Write-Step "Performing static analysis and saving corrections"
    if (-not (Invoke-Winapp2ool -Arguments '-s', '-offline', '-fc', '-debug', '-usedate', '-c', '-opti', '-keepdefaults', '-1f', 'winapp2-fluentcleaner.ini', `
        '-3f', 'winapp2-fluentcleaner.ini' `
        -ErrorMessage "Failed FluentCleaner flavor analysis")) { return $false }

    return (Complete-Flavor -Key FluentCleaner -Built 'winapp2-fluentcleaner.ini')
}

function Build-BleachBitFlavor {
    if (-not (Invoke-Winapp2ool -Arguments '-s', '-offline', '-bb', '-debug', '-usedate', '-c', '-opti', '-1f', 'winapp2-bleachbit-flavor.ini', `
        '-3f', 'winapp2-bleachbit-flavor.ini' `
        -ErrorMessage "Failed BleachBit flavor analysis")) { return $false }

    return (Complete-Flavor -Key BleachBit -Built 'winapp2-bleachbit-flavor.ini')
}

function Build-TronFlavor {
    # Tron uses BleachBit 6+, so its flavor starts from the base file, not the CCleaner flavor
    Write-Step "Creating Tron flavor"
    if (-not (Invoke-Winapp2ool -Arguments '-s', '-offline', '-tron', '-flavorize', '-autodetect', '-2f', 'winapp2-tron-flavor.ini', `
        '-9d', '\Sources\Flavors\Tron' -ErrorMessage "Failed to create Tron flavor" -RequiredDirs @(Join-Path $PSScriptRoot 'Sources\Flavors\Tron'))) { return $false }

    Write-Step "Performing static analysis and saving corrections"
    if (-not (Invoke-Winapp2ool -Arguments '-s', '-offline', '-tron', '-debug', '-usedate', '-c', '-opti', '-1f', 'winapp2-tron-flavor.ini', `
        '-3f', 'winapp2-tron-flavor.ini' `
        -ErrorMessage "Failed Tron flavor analysis")) { return $false }

    return (Complete-Flavor -Key Tron -Built 'winapp2-tron-flavor.ini')
}

function Build-SystemNinjaFlavor {
    if (-not (Invoke-Winapp2ool -Arguments '-s', '-offline', '-sn', '-debug', '-usedate', '-c', '-opti', '-1f', 'winapp2-systemninja-flavor.ini', `
        '-3f', 'Winapp2.rules' -ErrorMessage "Failed System Ninja flavor analysis")) { return $false }

    return (Complete-Flavor -Key SystemNinja -Built 'Winapp2.rules')
}

function Compare-ArtifactTree {
    # Returns one line per difference between a committed tree and its regeneration
    param([string]$Committed, [string]$Scratch)

    $name = Split-Path $Committed -Leaf
    $drift = @()

    foreach ($file in Get-ChildItem $Committed -Filter '*.ini' -Recurse -File) {
        $rel = $file.FullName.Substring($Committed.Length + 1)
        $counterpart = Join-Path $Scratch $rel
        if (-not (Test-Path $counterpart)) {
            $drift += "missing from regeneration: $name\$rel"
            continue
        }
        if ((Get-FileHash $file.FullName).Hash -ne (Get-FileHash $counterpart).Hash) {
            $drift += "content drift: $name\$rel"
        }
    }

    foreach ($file in Get-ChildItem $Scratch -Filter '*.ini' -Recurse -File) {
        $rel = $file.FullName.Substring($Scratch.Length + 1)
        if (-not (Test-Path (Join-Path $Committed $rel))) {
            $drift += "not present in committed artifacts: $name\$rel"
        }
    }

    return $drift
}

function New-Artifacts {
    # Clears the builders' folders under <Root>\Generated and <Root>\Archived and regenerates
    # them. A tracked file beside those folders survives (-GenerateTo '\Generated' used to
    # delete Generated\readme.md). Refuses a root inside Generated, Archived or Sources: the
    # full build's Combine reads the first two recursively, and the third is its input.
    param([string]$Root)

    $rel = $Root.Trim('\')
    $Root = if ($rel) { "\$rel" } else { '' }
    $target = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot $rel)).TrimEnd('\')

    foreach ($tree in 'Generated', 'Archived', 'Sources') {
        $treePath = Join-Path $PSScriptRoot $tree
        if ($target -eq $treePath -or $target.StartsWith("$treePath\", [StringComparison]::OrdinalIgnoreCase)) {
            Write-ErrorMsg "Can't generate into $target`: it's inside $tree\, which the build reads"
            return $false
        }
    }

    foreach ($tree in 'Generated', 'Archived') {
        foreach ($builder in 'BrowserBuilder', 'EntryBuilder') {
            $folder = Join-Path $target "$tree\$builder"
            if (Test-Path $folder) { Remove-Item $folder -Recurse -Force }
        }
    }

    return (Build-Artifacts -Root $Root)
}

function Test-Artifacts {
    Write-Step "Regenerating artifacts into scratch for characteristic testing"

    $scratch = Join-Path $PSScriptRoot 'Verify'
    if (Test-Path $scratch) { Remove-Item $scratch -Recurse -Force }

    if (-not (New-Artifacts -Root '\Verify')) { return $false }

    Write-Step "Comparing regenerated artifacts against committed artifacts under Generated and Archived"

    $drift = @()
    foreach ($tree in 'Generated', 'Archived') {
        $drift += @(Compare-ArtifactTree -Committed (Join-Path $PSScriptRoot $tree) -Scratch (Join-Path $scratch $tree))
    }

    if ($drift.Count -gt 0) {
        Write-ErrorMsg "Artifact drift detected - committed artifacts under Generated or Archived do not match their sources"
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
        '*.ini.old',
        '*.rules.old',
        'winapp2*.ini',
        'winapp2.rules',
        'archived*.ini',
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
    if (-not (Test-Winapp2oolVersion $script:Winapp2oolPath)) { exit 1 }

    # Generate-and-stop. Deliberately does NOT compare against Generated\: the caller is
    # diffing two generated trees against each other, so the committed artifacts are
    # not part of the question and their staleness must not fail the run.
    if ($GenerateTo) {
        if (-not (New-Artifacts -Root $GenerateTo)) { exit 1 }
        Write-Host "`nArtifacts generated into Generated and Archived under $GenerateTo" -ForegroundColor Green
        exit 0
    }

    if ($Verify) {
        if (-not (Test-Artifacts)) { exit 1 }
        Write-Host "`nCharacteristic test passed" -ForegroundColor Green
        exit 0
    }

    $releaseAssetList = Join-Path $PSScriptRoot '..\.github\scripts\release-assets.ps1'
    if (-not (Test-Path -LiteralPath $releaseAssetList)) {
        Write-ErrorMsg "Cannot find the published-file list: $releaseAssetList"
        exit 1
    }
    try {
        . $releaseAssetList
    }
    catch {
        Write-ErrorMsg "Cannot load the published-file list: $_"
        exit 1
    }
    Set-DeprecatedPublishPaths

    Backup-Files

    if (-not (Build-MainFile)) { exit 1 }
    if (-not (Build-Archive)) { exit 1 }
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
