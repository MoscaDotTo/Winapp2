#Requires -Version 7.0
<#
.SYNOPSIS
    Signs a winapp2ool.exe build so the self-updater will accept it.

.DESCRIPTION
    Signs the exe's bytes with ECDSA P-256 over SHA-256 and writes the 64-byte IEEE P1363 signature,
    base64 encoded, to <exe>.sig beside it: ASCII, no byte order mark, no trailing newline.
    The script then checks the signature it wrote against the public key and reports whether that key
    is one of the TrustedUpdateKeys in helpers\updater\UpdateSignature.vb.

    Sign the exact file that gets committed and published. Rebuilding afterwards invalidates the signature.

    Needs PowerShell 7, because Windows PowerShell's .NET Framework can't import PKCS#8.

.PARAMETER KeyFile
    The encrypted private key written by New-UpdateSigningKey.ps1. Omit it to use the key saved by
    Save-SigningPassword.ps1.

.PARAMETER ExePath
    The file to sign. Defaults to winapp2ool\bin\Release\winapp2ool.exe in this repository.

.PARAMETER Password
    The key password. Omit it to use the password saved by Save-SigningPassword.ps1 for this key,
    or to be prompted when none is saved.

.PARAMETER FromBuild
    Run from the Release build. Never prompts, and deletes the .sig left from the previous build first.
    With no key saved it then warns and succeeds, so machines without the key still build. With a key
    saved, any failure fails the build.

.EXAMPLE
    pwsh -File Sign-Winapp2oolRelease.ps1

.EXAMPLE
    pwsh -File Sign-Winapp2oolRelease.ps1 -KeyFile D:\keys\winapp2ool-backup.p8
#>
[CmdletBinding()]
param(
    [string] $KeyFile,

    [string] $ExePath = (Join-Path $PSScriptRoot '..\bin\Release\winapp2ool.exe'),

    [securestring] $Password,

    [switch] $FromBuild
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'UpdateSigningCommon.ps1')

$exeFullPath = [System.IO.Path]::GetFullPath($PSCmdlet.GetUnresolvedProviderPathFromPSPath($ExePath))
$sigPath = "$exeFullPath.sig"

if (-not (Test-Path -LiteralPath $exeFullPath -PathType Leaf)) {
    throw "'$exeFullPath' was not found."
}

# The build just rewrote the exe, so a .sig from an earlier build no longer matches it
if ($FromBuild -and (Test-Path -LiteralPath $sigPath)) {
    Remove-Item -LiteralPath $sigPath
}

$config = Get-SigningConfig

if (-not $KeyFile) {
    if (-not $config) {
        if ($FromBuild) {
            Write-Warning "No signing key is saved, so $exeFullPath was not signed. Run Save-SigningPassword.ps1 to sign Release builds."
            exit 0
        }
        throw 'No -KeyFile given and no key saved. Pass -KeyFile, or run Save-SigningPassword.ps1 first.'
    }
    $KeyFile = $config.KeyFile
}

$keyPath = [System.IO.Path]::GetFullPath($PSCmdlet.GetUnresolvedProviderPathFromPSPath($KeyFile))

if (-not (Test-Path -LiteralPath $keyPath -PathType Leaf)) {
    throw "Key file '$keyPath' was not found."
}

if (-not $Password -and $config -and $keyPath.Equals([System.IO.Path]::GetFullPath($config.KeyFile), [StringComparison]::OrdinalIgnoreCase)) {
    try {
        $Password = ConvertTo-SecureString -String $config.Password
    }
    catch {
        throw "Could not read the saved password in $SigningConfigPath. It only opens for the Windows account that saved it. Run Save-SigningPassword.ps1 again."
    }
}

if (-not $Password) {
    if ($FromBuild) {
        throw "No saved password for $keyPath, and a build can't prompt for one. Run Save-SigningPassword.ps1."
    }
    $Password = Read-Host -AsSecureString -Prompt 'Key password'
}

$key = Import-SigningKey -KeyPath $keyPath -Password $Password
try {
    $data = [System.IO.File]::ReadAllBytes($exeFullPath)
    $signature = $key.SignData($data, [System.Security.Cryptography.HashAlgorithmName]::SHA256)
    if ($signature.Length -ne 64) {
        throw "Expected a 64-byte signature, got $($signature.Length) bytes."
    }

    $publicParameters = $key.ExportParameters($false)
    $publicKey = Get-RawPublicKey $key
}
finally {
    $key.Dispose()
}

[System.IO.File]::WriteAllText($sigPath, [Convert]::ToBase64String($signature), [System.Text.Encoding]::ASCII)

# Check what landed on disk with the public half alone, the way the updater will
$verifier = [System.Security.Cryptography.ECDsa]::Create($publicParameters)
try {
    $writtenSignature = [Convert]::FromBase64String([System.IO.File]::ReadAllText($sigPath).Trim())
    $verified = $verifier.VerifyData([System.IO.File]::ReadAllBytes($exeFullPath), $writtenSignature, [System.Security.Cryptography.HashAlgorithmName]::SHA256)
}
finally {
    $verifier.Dispose()
}

if (-not $verified) {
    Remove-Item -LiteralPath $sigPath
    throw 'The signature did not verify after writing, so it was deleted.'
}

Write-Host "Signed $exeFullPath"
Write-Host "Wrote $sigPath and verified it."
Write-Host ''
Write-Host 'Public key:'
Write-Host "    $publicKey"

$signatureSource = Join-Path $PSScriptRoot '..\helpers\updater\UpdateSignature.vb'
if ((Test-Path -LiteralPath $signatureSource) -and (Select-String -LiteralPath $signatureSource -SimpleMatch -Quiet -Pattern "`"$publicKey`"")) {
    Write-Host 'This key is listed in TrustedUpdateKeys.'
}
else {
    Write-Warning 'This key is NOT listed in TrustedUpdateKeys in helpers\updater\UpdateSignature.vb. Clients built from this source will reject the signature.'
}
