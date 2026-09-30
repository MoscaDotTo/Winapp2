#Requires -Version 7.0
<#
.SYNOPSIS
    Saves the signing key's location and password so Release builds can sign winapp2ool.exe without asking.

.DESCRIPTION
    Asks for the key password once, checks that it opens the key, and saves the key path and the password
    to %LOCALAPPDATA%\winapp2ool\update-signing.json. The password is encrypted with DPAPI, so only your
    Windows account on this machine can read it back. A copy of that file is useless anywhere else.

    Sign-Winapp2oolRelease.ps1, and through it every Release build, then uses the saved key and password.
    Run this again to change them, or delete the json file to go back to being prompted.

    Keep the backup key off this list: it should only ever be opened with its password typed in.

.PARAMETER KeyFile
    The encrypted private key written by New-UpdateSigningKey.ps1.

.PARAMETER Password
    The key password. Omit it to be prompted, which is the normal way to run this script.

.EXAMPLE
    pwsh -File Save-SigningPassword.ps1 -KeyFile D:\keys\winapp2ool-primary.p8
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string] $KeyFile,

    [securestring] $Password
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'UpdateSigningCommon.ps1')

if (-not $IsWindows) {
    throw 'Saving the password needs DPAPI, which only exists on Windows.'
}

$keyPath = [System.IO.Path]::GetFullPath($PSCmdlet.GetUnresolvedProviderPathFromPSPath($KeyFile))

if (-not (Test-Path -LiteralPath $keyPath -PathType Leaf)) {
    throw "Key file '$keyPath' was not found."
}

if (-not $Password) {
    $Password = Read-Host -AsSecureString -Prompt 'Key password'
}

$key = Import-SigningKey -KeyPath $keyPath -Password $Password
try {
    $publicKey = Get-RawPublicKey $key
}
finally {
    $key.Dispose()
}

$config = [ordered]@{
    KeyFile  = $keyPath
    Password = ConvertFrom-SecureString -SecureString $Password
}

New-Item -ItemType Directory -Force -Path (Split-Path $SigningConfigPath) | Out-Null
$config | ConvertTo-Json | Set-Content -LiteralPath $SigningConfigPath -Encoding utf8NoBOM

Write-Host "Saved the key path and password to $SigningConfigPath"
Write-Host 'Release builds will now sign winapp2ool.exe with this key:'
Write-Host "    $publicKey"
