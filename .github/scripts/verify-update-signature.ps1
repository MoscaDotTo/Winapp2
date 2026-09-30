<#
.SYNOPSIS
    Checks that the published winapp2ool.exe carries a signature its clients will accept.

.DESCRIPTION
    Since 1.8 the self-updater installs winapp2ool\bin\Release\winapp2ool.exe only if
    winapp2ool.exe.sig verifies against a key in TrustedUpdateKeys
    (winapp2ool\helpers\updater\UpdateSignature.vb). A rebuild that isn't re-signed leaves a
    stale .sig, and every updating client then refuses the update. Nothing breaks visibly; users
    just stop receiving updates. This makes that a red X instead.

    Also checks that version.txt names the exe's own file version. The updater decides whether
    to download from version.txt, then refuses any exe that isn't newer, so a version.txt ahead
    of the exe makes clients download an update they will throw away.

    Reads only committed files and executes nothing from them, so it is safe on fork PRs.

.PARAMETER RepoRoot
    The repository root. Defaults to the git top level of the current directory.

.NOTES
    Exit 0 = the signature verifies and the versions agree; 1 = they don't.
    Windows only: FileVersionInfo reads the exe's version resource.
#>
[CmdletBinding()]
param(
    [string]$RepoRoot = (git rev-parse --show-toplevel)
)

$ErrorActionPreference = 'Stop'

$exe = Join-Path $RepoRoot 'winapp2ool\bin\Release\winapp2ool.exe'
$sig = "$exe.sig"
$keySource = Join-Path $RepoRoot 'winapp2ool\helpers\updater\UpdateSignature.vb'
$versionFile = Join-Path $RepoRoot 'winapp2ool\version.txt'

$failures = New-Object System.Collections.Generic.List[string]

foreach ($required in $exe, $sig, $keySource, $versionFile) {
    if (-not (Test-Path -LiteralPath $required)) { $failures.Add("missing: $required") }
}

if ($failures.Count -eq 0) {

    # The key list is the initializer of TrustedUpdateKeys; each key is 64 bytes, so 88 base64 characters
    $source = Get-Content -LiteralPath $keySource -Raw
    $initializer = [regex]::Match($source, '(?s)TrustedUpdateKeys\s+As[^\r\n]*?=\s*New\s+String\(\)\s*\{(.*?)\}').Groups[1].Value
    $keys = [regex]::Matches($initializer, '"([A-Za-z0-9+/]{86}==)"') | ForEach-Object { $_.Groups[1].Value }

    if (-not $keys) {
        $failures.Add('TrustedUpdateKeys lists no keys, so clients built from this source refuse every update')
    }

    $data = [System.IO.File]::ReadAllBytes($exe)
    $signature = $null
    try {
        $signature = [Convert]::FromBase64String((Get-Content -LiteralPath $sig -Raw).Trim())
    } catch {
        $failures.Add("winapp2ool.exe.sig isn't valid base64")
    }

    if ($signature -and $keys) {
        $verifiedBy = $null
        for ($i = 0; $i -lt $keys.Count -and -not $verifiedBy; $i++) {
            $point = [Convert]::FromBase64String($keys[$i])
            $parameters = [System.Security.Cryptography.ECParameters]@{
                Curve = [System.Security.Cryptography.ECCurve+NamedCurves]::nistP256
                Q     = [System.Security.Cryptography.ECPoint]@{ X = $point[0..31]; Y = $point[32..63] }
            }
            $verifier = [System.Security.Cryptography.ECDsa]::Create($parameters)
            try {
                if ($verifier.VerifyData($data, $signature, [System.Security.Cryptography.HashAlgorithmName]::SHA256)) {
                    $verifiedBy = $i
                }
            } finally {
                $verifier.Dispose()
            }
        }

        if ($null -eq $verifiedBy) {
            $failures.Add("winapp2ool.exe.sig doesn't verify against any of the $($keys.Count) trusted keys. Rebuild in Release (which signs) or run Sign-Winapp2oolRelease.ps1, then commit the exe and .sig together")
        } else {
            Write-Host "Signature verifies against trusted key $verifiedBy of $($keys.Count)."
        }
    }

    $exeVersion = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($exe).FileVersion
    $publishedVersion = (Get-Content -LiteralPath $versionFile -TotalCount 1).Trim()
    if ($exeVersion -ne $publishedVersion) {
        $failures.Add("version.txt says $publishedVersion but winapp2ool.exe is $exeVersion. Commit the version.txt the Release build wrote alongside the exe")
    } else {
        Write-Host "version.txt matches the exe ($exeVersion)."
    }
}

if ($failures.Count -gt 0) {
    foreach ($f in $failures) {
        if ($env:GITHUB_ACTIONS) { Write-Host "::error::$f" } else { Write-Host "error: $f" }
    }
    exit 1
}

exit 0
