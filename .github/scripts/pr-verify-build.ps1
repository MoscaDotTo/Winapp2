<#
.SYNOPSIS
    PR verify for winapp2.ini: build master, merge the PR, build again, and write a
    report saying whether the PR breaks the build and what it changes in the output.

.DESCRIPTION
    Runs Assembler\build winapp2.ps1 twice around a merge. We save the seven flavor
    files the baseline build published to Releases\, hard reset, merge the PR head, put
    those seven back in Releases\ as the previous version, and build again. The build
    script's Backup-Files stage copies each Releases\ file to a .old file and its
    changelog stage diffs .old against the new build, so the changelogs the second build
    writes to Releases\ are exactly what the PR changes. Nothing reads the deprecated
    copies the build also writes outside Releases\, so we ignore them.

.PARAMETER BaseSha
    The PR base commit on master (github.event.pull_request.base.sha).

.PARAMETER HeadSha
    The PR head commit (github.event.pull_request.head.sha) to merge in.

.PARAMETER PrNumber
    The PR number, written verbatim into the report bundle.

.PARAMETER ReportDir
    The folder to write the report into (created if missing). It gets status.txt,
    report.md, pr-number.txt, winapp2ool.log (when there is one), the build transcripts,
    each flavor's changelog, and a -verbose diff per flavor that prints every added,
    removed and modified entry in full.

.NOTES
    Exits 0 on every completed verification, a broken PR build included. The outcome
    goes in status.txt, so the comment job always runs. A nonzero exit means the
    harness itself failed.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$BaseSha,
    [Parameter(Mandatory)][string]$HeadSha,
    [Parameter(Mandatory)][string]$PrNumber,
    [Parameter(Mandatory)][string]$ReportDir
)

$ErrorActionPreference = 'Stop'
# A merge conflict or a failed build is a verdict about the PR, not a harness error, and
# we branch on $LASTEXITCODE for both. PS 7.3+ would otherwise throw on those exits and
# skip our merge_conflict and build_failed paths.
$PSNativeCommandUseErrorActionPreference = $false

# GitHub comment bodies cap at 65536 chars
$CommentCharBudget = 60000

$RepoRoot = (& git rev-parse --show-toplevel).Trim()
if ($LASTEXITCODE -ne 0 -or -not $RepoRoot) { throw 'not inside a git work tree' }
$Assembler = Join-Path $RepoRoot 'Assembler'
$BuildScript = Join-Path $Assembler 'build winapp2.ps1'
$Exe = Join-Path $Assembler 'winapp2ool.exe'

if (-not (Test-Path $BuildScript)) { throw "build script missing: $BuildScript" }
if (-not (Test-Path $Exe)) { throw "winapp2ool.exe missing from Assembler; a Release build stages it there, and it should be committed" }

New-Item -ItemType Directory -Force -Path $ReportDir | Out-Null
Set-Content -Path (Join-Path $ReportDir 'pr-number.txt') -Value $PrNumber -NoNewline

# The seven published flavor files, relative to the repo root, paired with the flavor
# label and the changelog the build script publishes beside each one. Order matches
# build winapp2.ps1's Backup-Files / changelog stages.
$Outputs = @(
    @{ Flavor = 'Base';          Output = 'Releases\winapp2.ini';               Diff = 'Releases\changelog-base.txt' }
    @{ Flavor = 'CCleaner';      Output = 'Releases\winapp2-ccleaner.ini';      Diff = 'Releases\changelog-ccleaner.txt' }
    @{ Flavor = 'BleachBit';     Output = 'Releases\winapp2-bleachbit.ini';     Diff = 'Releases\changelog-bleachbit.txt' }
    @{ Flavor = 'CCleaner7';     Output = 'Releases\winapp2-ccleaner7.ini';     Diff = 'Releases\changelog-ccleaner7.txt' }
    @{ Flavor = 'FluentCleaner'; Output = 'Releases\winapp2-fluentcleaner.ini'; Diff = 'Releases\changelog-fluentcleaner.txt' }
    @{ Flavor = 'Tron';          Output = 'Releases\winapp2-tron.ini';          Diff = 'Releases\changelog-tron.txt' }
    @{ Flavor = 'SystemNinja';   Output = 'Releases\winapp2-systemninja.rules'; Diff = 'Releases\changelog-systemninja.txt' }
)

function Invoke-Git {
    # git's output goes to the host, never the pipeline, so the exit code is the only
    # thing we return. Otherwise a caller testing `-ne 0` would be comparing git's
    # stdout lines.
    param([Parameter(ValueFromRemainingArguments)][string[]]$GitArgs)
    & git @GitArgs 2>&1 | Out-Host
    return $LASTEXITCODE
}

function Invoke-Build {
    # A child process, so the build script's `exit N` lands in $LASTEXITCODE.
    param([string]$TranscriptPath)
    Push-Location $Assembler
    try {
        # Out-Host keeps the passthrough copy off the pipeline (see Invoke-Git).
        & pwsh -NoProfile -File $BuildScript *>&1 | Tee-Object -FilePath $TranscriptPath | Out-Host
        return $LASTEXITCODE
    } finally {
        Pop-Location
    }
}

function Complete-Verification {
    # ExitCode stays 0 for every verdict about the PR, a broken PR build included, so the
    # comment job treats the run as a completed verification. Pass 1 only when the harness
    # failed rather than the PR, so the run goes red.
    param([string]$Status, [string]$ReportMarkdown, [int]$ExitCode = 0)
    Set-Content -Path (Join-Path $ReportDir 'status.txt') -Value $Status -NoNewline
    Set-Content -Path (Join-Path $ReportDir 'report.md') -Value $ReportMarkdown
    $log = Join-Path $Assembler 'winapp2ool.log'
    if (Test-Path $log) { Copy-Item $log (Join-Path $ReportDir 'winapp2ool.log') -Force }
    Write-Host "PR-verify status: $Status"
    exit $ExitCode
}

function Invoke-VerboseDiff {
    # The build's changelog names added and removed entries without their contents, which
    # leaves a reviewer nothing to audit. Rerun Diff with -verbose over the same pair.
    # Full paths go through -Nd because -Nf treats its value as a bare file name.
    param([string]$Old, [string]$New, [string]$OutFile)
    $Old, $New, $OutFile = $Old, $New, $OutFile | ForEach-Object { [System.IO.Path]::GetFullPath($_) }
    Remove-Item $OutFile -Force -ErrorAction SilentlyContinue
    $diffArgs = '-s', '-offline', '-diff', '-verbose', '-savelog',
                '-1d', "`"$Old`"", '-2d', "`"$New`"", '-3d', "`"$OutFile`""
    # Start-Process rather than &: winapp2ool exits early when its stdout is redirected
    $p = Start-Process -FilePath $Exe -ArgumentList $diffArgs -WorkingDirectory $Assembler `
                       -WindowStyle Hidden -Wait -PassThru
    return ($p.ExitCode -eq 0 -and (Test-Path $OutFile))
}

function Get-Tail {
    param([string]$Path, [int]$Lines = 40)
    if (-not (Test-Path $Path)) { return '(no output captured)' }
    return (Get-Content $Path -Tail $Lines) -join "`n"
}

# --- 1. Baseline build on master --------------------------------------------------
Write-Host ">> Checking out base $BaseSha"
if ((Invoke-Git checkout --force --detach $BaseSha) -ne 0) { throw "cannot checkout base $BaseSha" }
if ((Invoke-Git reset --hard $BaseSha) -ne 0) { throw 'cannot reset to base' }
Invoke-Git clean -fdx | Out-Null

$baselineTranscript = Join-Path $ReportDir 'baseline-build.log'
Write-Host ">> Baseline build (master)"
if ((Invoke-Build -TranscriptPath $baselineTranscript) -ne 0) {
    $body = @"
**The baseline build of master failed.**

This isn't caused by this PR. Master itself doesn't build on the runner right now, and
a maintainer needs to look into it before this PR can be verified.

Failing tail of the baseline build:

``````
$(Get-Tail $baselineTranscript)
``````
"@
    Complete-Verification -Status 'baseline_failed' -ReportMarkdown $body
}

# --- 2. Preserve the seven baseline outputs --------------------------------------
$baselineSave = Join-Path $ReportDir '_baseline'
New-Item -ItemType Directory -Force -Path $baselineSave | Out-Null
foreach ($o in $Outputs) {
    $src = Join-Path $RepoRoot $o.Output
    if (Test-Path $src) {
        $dst = Join-Path $baselineSave $o.Output
        New-Item -ItemType Directory -Force -Path (Split-Path $dst) | Out-Null
        Copy-Item $src $dst -Force
    }
}

# --- 3. Clean, then merge the PR head --------------------------------------------
Write-Host ">> Resetting tree and merging head $HeadSha"
if ((Invoke-Git reset --hard $BaseSha) -ne 0) { throw 'cannot reset before merge' }
Invoke-Git clean -fdx | Out-Null
# `merge --no-ff` writes a commit, and a fresh runner checkout has no committer identity,
# so git refuses before it ever looks at the trees. We pass one with -c instead of setting
# the repo's config, so the script leaves a maintainer's local clone alone.
$mergeOutput = (& git -c 'user.name=winapp2 pr-verify' `
                      -c 'user.email=pr-verify@users.noreply.github.com' `
                      merge --no-edit --no-ff $HeadSha 2>&1 | Out-String)
$mergeExit = $LASTEXITCODE
Write-Host $mergeOutput
if ($mergeExit -ne 0) {
    # Only a content conflict leaves a merge in progress. Every other failure (missing
    # ref, unrelated histories, dirty tree, no identity) is a harness fault, and must
    # not be reported to the contributor as "your branch conflicts".
    & git rev-parse --verify --quiet MERGE_HEAD *> $null
    $conflicted = ($LASTEXITCODE -eq 0)

    if ($conflicted) {
        Invoke-Git merge --abort | Out-Null
        $body = @"
**This PR doesn't merge cleanly into master.**

The verify build was skipped. Rebase or merge the latest master into your branch to
resolve the conflicts, then push again.
"@
        Complete-Verification -Status 'merge_conflict' -ReportMarkdown $body
    }

    $body = @"
**PR verify couldn't merge this PR, and it isn't a conflict.**

The merge failed before git compared the branches, so this is a CI problem rather
than something wrong with the PR. A maintainer needs to look at the run.

``````
$($mergeOutput.Trim())
``````
"@
    Complete-Verification -Status 'merge_failed' -ReportMarkdown $body -ExitCode 1
}

# --- 4. Restore baselines so build #2 diffs against master ------------------------
foreach ($o in $Outputs) {
    $saved = Join-Path $baselineSave $o.Output
    if (Test-Path $saved) {
        $dst = Join-Path $RepoRoot $o.Output
        New-Item -ItemType Directory -Force -Path (Split-Path $dst) | Out-Null
        Copy-Item $saved $dst -Force
    }
}

# --- 5. Post-merge build ----------------------------------------------------------
$mergedTranscript = Join-Path $ReportDir 'merged-build.log'
Write-Host ">> Post-merge build (PR applied)"
if ((Invoke-Build -TranscriptPath $mergedTranscript) -ne 0) {
    # The last ">>" step line in the transcript names the stage that was running.
    $lastStep = (Select-String -Path $mergedTranscript -Pattern '^>>\s*(.+)$' |
                 Select-Object -Last 1).Matches.Groups[1].Value
    if (-not $lastStep) { $lastStep = '(unknown stage)' }
    $body = @"
**This PR breaks the winapp2.ini build.**

Failing stage: **$lastStep**

Tail of the build output:

``````
$(Get-Tail $mergedTranscript)
``````

Tail of winapp2ool.log (full log attached to the workflow run):

``````
$(Get-Tail (Join-Path $Assembler 'winapp2ool.log'))
``````
"@
    Complete-Verification -Status 'build_failed' -ReportMarkdown $body
}

# --- 6. Success: collect diffs ----------------------------------------------------
$diffsDir = Join-Path $ReportDir 'diffs'
New-Item -ItemType Directory -Force -Path $diffsDir | Out-Null

$baseDiffText = $null
$baseVerboseText = $null
$flavorLines = New-Object System.Collections.Generic.List[string]
foreach ($o in $Outputs) {
    $diffPath = Join-Path $RepoRoot $o.Diff
    if (-not (Test-Path $diffPath)) { continue }
    Copy-Item $diffPath (Join-Path $diffsDir ("{0}-diff.txt" -f $o.Flavor)) -Force
    $content = Get-Content $diffPath -Raw
    if ($o.Flavor -eq 'Base') { $baseDiffText = $content }
    $lineCount = (Get-Content $diffPath | Measure-Object -Line).Lines

    # A missing verbose diff costs the reviewer detail, not the verdict, so it never fails the run
    $verboseNote = 'no verbose diff'
    $old = Join-Path $baselineSave $o.Output
    $new = Join-Path $RepoRoot $o.Output
    $verbosePath = Join-Path $diffsDir ("{0}-diff-verbose.txt" -f $o.Flavor)
    if ((Test-Path $old) -and (Test-Path $new)) {
        if (Invoke-VerboseDiff -Old $old -New $new -OutFile $verbosePath) {
            $verboseNote = '{0} lines verbose' -f (Get-Content $verbosePath | Measure-Object -Line).Lines
            if ($o.Flavor -eq 'Base') { $baseVerboseText = Get-Content $verbosePath -Raw }
        } else {
            $verboseNote = 'verbose diff failed'
        }
    }
    $flavorLines.Add(("- **{0}**: {1} lines, {2} (see run artifacts)" -f $o.Flavor, $lineCount, $verboseNote))
}

# Show the verbose diff when it fits, since it's the one that lets a reviewer read added entries
$shownKind = 'diff'
if ($baseVerboseText -and $baseVerboseText.Length -le $CommentCharBudget) {
    $baseDiffText = $baseVerboseText
    $shownKind = 'verbose diff'
}

if (-not $baseDiffText) { $baseDiffText = '(no base diff was produced)' }
$truncNote = ''
if ($baseDiffText.Length -gt $CommentCharBudget) {
    $baseDiffText = $baseDiffText.Substring(0, $CommentCharBudget)
    $truncNote = "`n`n_…base diff truncated. The full changelogs, and verbose ones showing every added and removed entry in full, are attached to the workflow run._"
}

$body = @"
**This PR builds cleanly.** Below is the base winapp2.ini changelog it produces.

Per-flavor changelogs (full text attached to the workflow run, with a verbose version that prints each changed entry in full):

$($flavorLines -join "`n")

<details><summary>Base winapp2.ini $shownKind</summary>

``````
$baseDiffText
``````
$truncNote
</details>
"@
Complete-Verification -Status 'success' -ReportMarkdown $body
