<#
.SYNOPSIS
    The one list of what a winapp2.ini GitHub release carries, and the code that stages it.

.DESCRIPTION
    Dot-source this file from any folder; it finds flavors.csv beside itself. Both release
    paths use it, so their asset lists can't drift apart:
      * build-on-merge.yml stages the bundle right after the daily build commits;
      * release.yml's workflow_dispatch recovery path stages it from master's committed files.
    PR verify reads the same rows for its report.

    flavors.csv has one row per published file, in release order: its key (the WinappFlavor
    member name, or Archive), a label for log messages, its release asset, its changelog
    asset, and its deprecated copies as repo paths with forward slashes (an empty cell means
    none). winapp2ool's WinappFlavorInfo names the flavors too, and a unit test holds the two
    together. The release data files live in Releases\ under their asset names. They are
    named one by one, never globbed, so nothing else that lands in Releases\ becomes an asset
    or a gate input by accident. The release also carries winapp2ool.exe.

    Defines:
      $PublishedFiles     the rows of flavors.csv
      $ReleaseDataFiles   every flavor file and archived.ini ("published" files: the release gate's input)
      $ReleaseChangelogs  their changelogs
      $ReleaseFiles       both lists, as repo-relative paths (Releases/<name>)
      $DeprecatedFiles    every deprecated copy, as a repo-relative path
      Get-ReleaseVersion  the highest "; Version:" stamp across the given files
      New-ReleaseBundle   copies the assets plus version.txt and commit.txt into a folder
#>

$FlavorManifest = Join-Path $PSScriptRoot 'flavors.csv'
if (-not (Test-Path -LiteralPath $FlavorManifest)) { throw "flavor manifest missing: $FlavorManifest" }
$PublishedFiles = @(Import-Csv -LiteralPath $FlavorManifest)
if ($PublishedFiles.Count -eq 0) { throw "flavor manifest has no rows: $FlavorManifest" }

$ReleaseDataFiles = @($PublishedFiles | ForEach-Object { $_.Asset })
$ReleaseChangelogs = @($PublishedFiles | ForEach-Object { $_.Changelog })
$ReleaseFiles = @($ReleaseDataFiles + $ReleaseChangelogs | ForEach-Object { "Releases/$_" })
$DeprecatedFiles = @($PublishedFiles | ForEach-Object { $_.OldData; $_.OldChangelog } | Where-Object { $_ })

function Get-ReleaseVersion {
    <#
    .SYNOPSIS
        Returns the highest "; Version: <date>" stamp across the given files, as a string.
        Throws when none of them has one.
    #>
    param([Parameter(Mandatory)][string[]]$Paths)

    $versions = @(foreach ($p in $Paths) {
        $m = Select-String -LiteralPath $p -Pattern '^;\s*Version:\s*(\d+)' | Select-Object -First 1
        if ($m) { [int]$m.Matches.Groups[1].Value }
    })
    if ($versions.Count -eq 0) {
        throw "could not parse a version line from any of: $($Paths -join ', ')"
    }
    return ($versions | Sort-Object -Descending | Select-Object -First 1).ToString()
}

function New-ReleaseBundle {
    <#
    .SYNOPSIS
        Stages the release assets in Destination: each release data file under its
        own name, Assembler/winapp2ool.exe as winapp2ool.exe, and the version.txt and commit.txt
        that release.yml reads. A missing asset throws: 1.8 clients download every flavor from
        the latest release, so a release without one would 404 that flavor for all of them.
    #>
    param(
        [Parameter(Mandatory)][string]$Destination,
        [Parameter(Mandatory)][string]$Version,
        [Parameter(Mandatory)][string]$Sha
    )

    $assets = @($ReleaseFiles | ForEach-Object { @{ Src = $_; Name = Split-Path $_ -Leaf } })
    $assets += @{ Src = 'Assembler/winapp2ool.exe'; Name = 'winapp2ool.exe' }

    $missing = @($assets | Where-Object { -not (Test-Path -LiteralPath $_.Src) } | ForEach-Object { $_.Src })
    if ($missing.Count -gt 0) {
        throw "release assets missing, not releasing: $($missing -join ', ')"
    }

    New-Item -ItemType Directory -Force -Path $Destination | Out-Null
    foreach ($a in $assets) {
        Copy-Item -LiteralPath $a.Src (Join-Path $Destination $a.Name) -Force
    }
    Set-Content -Path (Join-Path $Destination 'version.txt') -Value $Version -NoNewline
    Set-Content -Path (Join-Path $Destination 'commit.txt') -Value $Sha -NoNewline
}
