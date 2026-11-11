<#
.SYNOPSIS
    The one list of what a winapp2.ini GitHub release carries, and the code that stages it.

.DESCRIPTION
    Dot-source this file from the repository root. Both release paths use it, so their asset
    lists can't drift apart:
      * build-on-merge.yml stages the bundle right after the daily build commits;
      * release.yml's workflow_dispatch recovery path stages it from master's committed files.

    The 14 release data files live in Releases\ under the names they carry on the release
    (winapp2ool's WinappFlavorInfo names the same assets). They are named one by one, never
    globbed, so nothing else that lands in Releases\ becomes an asset or a gate input by accident.
    winapp2ool.exe is the fifteenth asset.

    Defines:
      $ReleaseDataFiles   the seven flavor files ("published" files: the release gate's input)
      $ReleaseChangelogs  the seven changelogs
      $ReleaseFiles       all 14, as repo-relative paths (Releases/<name>)
      Get-ReleaseVersion  the highest "; Version:" stamp across the given files
      New-ReleaseBundle   copies the 15 assets plus version.txt and commit.txt into a folder
#>

$ReleaseDataFiles = @(
    'winapp2.ini'
    'winapp2-ccleaner.ini'
    'winapp2-ccleaner7.ini'
    'winapp2-bleachbit.ini'
    'winapp2-fluentcleaner.ini'
    'winapp2-tron.ini'
    'winapp2-systemninja.rules'
)

$ReleaseChangelogs = @(
    'changelog-base.txt'
    'changelog-ccleaner.txt'
    'changelog-ccleaner7.txt'
    'changelog-bleachbit.txt'
    'changelog-fluentcleaner.txt'
    'changelog-tron.txt'
    'changelog-systemninja.txt'
)

$ReleaseFiles = @($ReleaseDataFiles + $ReleaseChangelogs | ForEach-Object { "Releases/$_" })

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
        Stages the release assets in Destination: each of the 14 release data files under its
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
