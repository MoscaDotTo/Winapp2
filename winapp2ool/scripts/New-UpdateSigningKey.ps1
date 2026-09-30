#Requires -Version 7.0
<#
.SYNOPSIS
    Generates a P-256 signing key for winapp2ool self-updates.

.DESCRIPTION
    Writes the private key as password-encrypted PKCS#8 (PBES2, AES-256-CBC, PBKDF2 with SHA-256)
    and prints the matching public key in the form winapp2ool embeds: base64 of the raw point X||Y.
    Paste that public key into TrustedUpdateKeys in helpers\updater\UpdateSignature.vb.

    The key file must live outside every git working tree, and an existing file is never overwritten.
    Keep the key file and its password in separate places, and keep a backup of both.

    Needs PowerShell 7, because Windows PowerShell's .NET Framework can't export PKCS#8.

.PARAMETER OutFile
    Where to write the encrypted private key. Its folder must already exist.

.PARAMETER Password
    The key password. Omit it to be prompted twice, which is the normal way to run this script.

.EXAMPLE
    pwsh -File New-UpdateSigningKey.ps1 -OutFile D:\keys\winapp2ool-primary.p8
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string] $OutFile,

    [securestring] $Password
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'UpdateSigningCommon.ps1')

$outPath = [System.IO.Path]::GetFullPath($PSCmdlet.GetUnresolvedProviderPathFromPSPath($OutFile))
$outDir = [System.IO.Path]::GetDirectoryName($outPath)

if (-not (Test-Path -LiteralPath $outDir -PathType Container)) {
    throw "The folder '$outDir' doesn't exist. Create it first."
}

if (Test-Path -LiteralPath $outPath) {
    throw "'$outPath' already exists. Refusing to overwrite a key file."
}

Assert-OutsideGitRepository -Path $outPath -Folder $outDir

if (-not $Password) {
    $Password = Read-Host -AsSecureString -Prompt 'Key password (12 characters or more)'
    $confirm = Read-Host -AsSecureString -Prompt 'Repeat the password'

    $first = ConvertTo-CharArray $Password
    $second = ConvertTo-CharArray $confirm
    try {
        $same = $first.Length -eq $second.Length
        for ($i = 0; $same -and $i -lt $first.Length; $i++) {
            $same = $first[$i] -ceq $second[$i]
        }
    }
    finally {
        [Array]::Clear($first, 0, $first.Length)
        [Array]::Clear($second, 0, $second.Length)
    }

    if (-not $same) {
        throw 'The passwords did not match. No key was written.'
    }
}

if ($Password.Length -lt 12) {
    throw 'The password must be at least 12 characters. No key was written.'
}

$key = [System.Security.Cryptography.ECDsa]::Create([System.Security.Cryptography.ECCurve+NamedCurves]::nistP256)
$passwordChars = ConvertTo-CharArray $Password
try {
    $pbe = [System.Security.Cryptography.PbeParameters]::new(
        [System.Security.Cryptography.PbeEncryptionAlgorithm]::Aes256Cbc,
        [System.Security.Cryptography.HashAlgorithmName]::SHA256,
        $KeyIterationCount)
    $encrypted = $key.ExportEncryptedPkcs8PrivateKey([char[]] $passwordChars, $pbe)

    # CreateNew fails rather than replacing a file that appeared since the check above
    $stream = [System.IO.File]::Open($outPath, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
    try {
        $stream.Write($encrypted, 0, $encrypted.Length)
    }
    finally {
        $stream.Dispose()
    }

    $publicKey = Get-RawPublicKey $key
}
finally {
    [Array]::Clear($passwordChars, 0, $passwordChars.Length)
    $key.Dispose()
}

Write-Host "Wrote the encrypted private key to $outPath"
Write-Host ''
Write-Host 'Public key:'
Write-Host "    $publicKey"
Write-Host ''
Write-Host 'Add it to TrustedUpdateKeys in winapp2ool\helpers\updater\UpdateSignature.vb as its own quoted string.'
