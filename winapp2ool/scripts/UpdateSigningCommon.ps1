#Requires -Version 7.0
# Helpers shared by the update signing scripts. Dot-source it; it does nothing on its own.

Set-StrictMode -Version Latest

# PBKDF2-SHA256 rounds protecting the private key file
$KeyIterationCount = 600000

$P256Oid = '1.2.840.10045.3.1.7'

# Written by Save-SigningPassword.ps1. The password inside is DPAPI-encrypted to the current Windows user
$SigningConfigPath = Join-Path $env:LOCALAPPDATA 'winapp2ool\update-signing.json'

function Get-SigningConfig {
    <#
    .SYNOPSIS
        Returns the saved key path and DPAPI-protected password, or $null when none is saved.
    #>
    if (-not (Test-Path -LiteralPath $SigningConfigPath -PathType Leaf)) {
        return $null
    }

    $config = Get-Content -LiteralPath $SigningConfigPath -Raw | ConvertFrom-Json
    if (-not $config.KeyFile -or -not $config.Password) {
        throw "$SigningConfigPath is incomplete. Run Save-SigningPassword.ps1 again."
    }

    return $config
}

function Import-SigningKey {
    <#
    .SYNOPSIS
        Decrypts an encrypted PKCS#8 P-256 key file. The caller disposes the returned key.
    #>
    param(
        [Parameter(Mandatory)] [string] $KeyPath,
        [Parameter(Mandatory)] [securestring] $Password
    )

    $key = [System.Security.Cryptography.ECDsa]::Create()
    $passwordChars = ConvertTo-CharArray $Password
    try {
        $keyBytes = [System.IO.File]::ReadAllBytes($KeyPath)
        $bytesRead = 0
        try {
            $key.ImportEncryptedPkcs8PrivateKey([char[]] $passwordChars, [byte[]] $keyBytes, [ref] $bytesRead)
        }
        catch [System.Security.Cryptography.CryptographicException] {
            throw 'Could not decrypt the key file. The password is wrong or the file is not an encrypted PKCS#8 key.'
        }

        if ($bytesRead -ne $keyBytes.Length) {
            throw 'The key file has unexpected data after the key.'
        }

        if (-not (Test-IsP256 $key)) {
            throw 'The key file does not hold a P-256 key.'
        }
    }
    catch {
        $key.Dispose()
        throw
    }
    finally {
        [Array]::Clear($passwordChars, 0, $passwordChars.Length)
    }

    return $key
}

function ConvertTo-CharArray {
    <#
    .SYNOPSIS
        Copies a SecureString into a char array the caller clears when done, so the password never becomes an immutable string.
    #>
    param([Parameter(Mandatory)] [securestring] $SecureString)

    $pointer = [System.Runtime.InteropServices.Marshal]::SecureStringToGlobalAllocUnicode($SecureString)
    try {
        $chars = [char[]]::new($SecureString.Length)
        [System.Runtime.InteropServices.Marshal]::Copy($pointer, $chars, 0, $chars.Length)
        return , $chars
    }
    finally {
        [System.Runtime.InteropServices.Marshal]::ZeroFreeGlobalAllocUnicode($pointer)
    }
}

function Get-RawPublicKey {
    <#
    .SYNOPSIS
        Returns a P-256 key's public point as winapp2ool stores it: base64 of X followed by Y.
    #>
    param([Parameter(Mandatory)] [System.Security.Cryptography.ECDsa] $Key)

    $q = $Key.ExportParameters($false).Q
    return [Convert]::ToBase64String([byte[]] ($q.X + $q.Y))
}

function Test-IsP256 {
    param([Parameter(Mandatory)] [System.Security.Cryptography.ECDsa] $Key)

    $curve = $Key.ExportParameters($false).Curve
    if ($Key.KeySize -ne 256 -or -not $curve.IsNamed) {
        return $false
    }

    return ($curve.Oid.Value -eq $P256Oid) -or ($curve.Oid.FriendlyName -in @('nistP256', 'ECDSA_P256'))
}

function Get-GitTopLevel {
    <#
    .SYNOPSIS
        Returns the root of the git working tree holding a folder, or $null when there is none.
    #>
    param([Parameter(Mandatory)] [string] $Folder)

    $topLevel = & git -C $Folder rev-parse --show-toplevel 2>$null
    if ($LASTEXITCODE -ne 0 -or -not $topLevel) {
        return $null
    }

    return [System.IO.Path]::GetFullPath(($topLevel | Select-Object -First 1).Trim())
}

function Test-PathIsUnder {
    param(
        [Parameter(Mandatory)] [string] $Path,
        [Parameter(Mandatory)] [string] $Root
    )

    $full = [System.IO.Path]::GetFullPath($Path).TrimEnd('\', '/')
    $rootFull = [System.IO.Path]::GetFullPath($Root).TrimEnd('\', '/')

    return $full.Equals($rootFull, [StringComparison]::OrdinalIgnoreCase) -or
        $full.StartsWith($rootFull + [System.IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)
}

function Assert-OutsideGitRepository {
    <#
    .SYNOPSIS
        Throws if a path is inside the winapp2 repository or inside any other git working tree.
    #>
    param(
        [Parameter(Mandatory)] [string] $Path,
        [Parameter(Mandatory)] [string] $Folder
    )

    if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
        throw 'git was not found on PATH, so the script cannot confirm the key file stays out of the repository.'
    }

    $repoRoot = Get-GitTopLevel -Folder $PSScriptRoot
    if (-not $repoRoot) {
        throw "Could not find the git repository holding $PSScriptRoot."
    }

    if (Test-PathIsUnder -Path $Path -Root $repoRoot) {
        throw "Refusing to write a private key inside the repository ($repoRoot). Choose a folder outside it."
    }

    $enclosing = Get-GitTopLevel -Folder $Folder
    if ($enclosing) {
        throw "Refusing to write a private key inside the git working tree at $enclosing. Choose a folder outside any repository."
    }

    # --show-toplevel fails inside a .git folder, so ask about that case directly
    $insideGitDir = & git -C $Folder rev-parse --is-inside-git-dir 2>$null
    if ($LASTEXITCODE -eq 0 -and ($insideGitDir | Select-Object -First 1) -eq 'true') {
        throw "Refusing to write a private key inside a .git folder. Choose a folder outside any repository."
    }
}
