<#
.SYNOPSIS
    PR-verify mechanism for winapp2.ini: build master, merge the PR, build again,
    and emit a report bundle describing whether the PR breaks the build and what it
    changes in the output.

.DESCRIPTION
    Drives Assembler\build winapp2.ps1 twice around a merge.
    Preserve the seven committed build outputs after the baseline build, hard
    reset, merge the PR head, restore those seven as the "previous" baseline, then
    build again. The build script's own Backup-Files stage copies each committed
    output to a .old file and its changelog stage diffs .old vs the freshly built
    file, so the second build's diff.txt is exactly what the PR changes.

.PARAMETER BaseSha
    The PR base commit (github.event.pull_request.base.sha) -- master.

.PARAMETER HeadSha
    The PR head commit (github.event.pull_request.head.sha) to merge in.

.PARAMETER PrNumber
    The PR number, written verbatim into the report bundle.

.PARAMETER ReportDir
    Directory into which to write the report (created if absent). Receives
    status.txt, report.md, pr-number.txt, winapp2ool.log (when present), the
    build transcripts, and the collected diff.txt files.

.NOTES
    Always exits 0 on a completed verification (including a broken PR build)
    the outcome lives in status.txt so the downstream comment job always fires.
    A nonzero exit means the harness failed.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$BaseSha,
    [Parameter(Mandatory)][string]$HeadSha,
    [Parameter(Mandatory)][string]$PrNumber,
    [Parameter(Mandatory)][string]$ReportDir
)

$ErrorActionPreference = 'Stop'
# We drive git and the build as native commands and branch on their exit codes
# ($LASTEXITCODE) -- a nonzero git merge (conflict) or a failed build is expected
# control flow, not a harness error. Keep PS 7.3+ from promoting those exits to
# thrown terminating errors, which would bypass our merge_conflict/build_failed paths.
$PSNativeCommandUseErrorActionPreference = $false

# GitHub comment bodies cap at 65536 chars
$CommentCharBudget = 60000

$RepoRoot = (& git rev-parse --show-toplevel).Trim()
if ($LASTEXITCODE -ne 0 -or -not $RepoRoot) { throw 'not inside a git work tree' }
$Assembler = Join-Path $RepoRoot 'Assembler'
$BuildScript = Join-Path $Assembler 'build winapp2.ps1'
$Exe = Join-Path $Assembler 'winapp2ool.exe'

if (-not (Test-Path $BuildScript)) { throw "build script missing: $BuildScript" }
if (-not (Test-Path $Exe)) { throw "winapp2ool.exe missing from Assembler -- the committed Release build should stage it there" }

New-Item -ItemType Directory -Force -Path $ReportDir | Out-Null
Set-Content -Path (Join-Path $ReportDir 'pr-number.txt') -Value $PrNumber -NoNewline

# The seven committed build outputs, relative to the repo root, paired with the flavor
# label and the location the build script drops each flavor's diff.txt. Order matches
# build winapp2.ps1's Backup-Files / changelog stages.
$Outputs = @(
    @{ Flavor = 'Base';        Output = 'Non-CCleaner\Winapp2.ini';               Diff = 'Non-CCleaner\diff.txt' }
    @{ Flavor = 'CCleaner';    Output = 'Winapp2.ini';                            Diff = 'diff.txt' }
    @{ Flavor = 'BleachBit';   Output = 'Non-CCleaner\BleachBit\Winapp2.ini';     Diff = 'Non-CCleaner\BleachBit\diff.txt' }
    @{ Flavor = 'CCleaner7';   Output = 'Non-CCleaner\CCleaner7\Winapp2.ini';     Diff = 'Non-CCleaner\CCleaner7\diff.txt' }
    @{ Flavor = 'FluentCleaner'; Output = 'Non-CCleaner\FluentCleaner\Winapp2.ini'; Diff = 'Non-CCleaner\FluentCleaner\diff.txt' }
    @{ Flavor = 'Tron';        Output = 'Non-CCleaner\Tron\Winapp2.ini';          Diff = 'Non-CCleaner\Tron\diff.txt' }
    @{ Flavor = 'SystemNinja'; Output = 'Non-CCleaner\SystemNinja\Winapp2.rules'; Diff = 'Non-CCleaner\SystemNinja\diff.txt' }
)

function Invoke-Git {
    # Print git's output to the host (never into the pipeline) so this function's
    # only return value is the exit code -- otherwise callers testing `-ne 0` would
    # be comparing git's stdout lines, not its exit status.
    param([Parameter(ValueFromRemainingArguments)][string[]]$GitArgs)
    & git @GitArgs 2>&1 | Out-Host
    return $LASTEXITCODE
}

function Invoke-Build {
    # Run the committed build script as a child process from Assembler so its
    # `exit N` propagates cleanly; capture the console transcript for the report.
    param([string]$TranscriptPath)
    Push-Location $Assembler
    try {
        # Tee to the transcript for the report; Out-Host keeps the passthrough copy
        # off the pipeline so only the exit code is returned (see Invoke-Git).
        & pwsh -NoProfile -File $BuildScript *>&1 | Tee-Object -FilePath $TranscriptPath | Out-Host
        return $LASTEXITCODE
    } finally {
        Pop-Location
    }
}

function Complete-Verification {
    # Write status + report body, copy the winapp2ool log if present, and stop.
    param([string]$Status, [string]$ReportMarkdown)
    Set-Content -Path (Join-Path $ReportDir 'status.txt') -Value $Status -NoNewline
    Set-Content -Path (Join-Path $ReportDir 'report.md') -Value $ReportMarkdown
    $log = Join-Path $Assembler 'winapp2ool.log'
    if (Test-Path $log) { Copy-Item $log (Join-Path $ReportDir 'winapp2ool.log') -Force }
    Write-Host "PR-verify status: $Status"
    exit 0
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
🛠 **Baseline build of master failed.**

This is **not** caused by this PR: master itself does not currently build on the
runner. A maintainer needs to look into it before this PR can be verified.

Failing tail of the baseline build:

``````
$(Get-Tail $baselineTranscript)
``````
"@
    Complete-Verification -Status 'baseline_failed' -ReportMarkdown $body
}

# --- 2. Preserve the six baseline outputs ----------------------------------------
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
if ((Invoke-Git merge --no-edit --no-ff $HeadSha) -ne 0) {
    Invoke-Git merge --abort | Out-Null
    $body = @"
 **This PR does not merge cleanly into master.**

The verify build was skipped. Rebase or merge the latest master into your branch to
resolve the conflicts, then push again.
"@
    Complete-Verification -Status 'merge_conflict' -ReportMarkdown $body
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
$flavorLines = New-Object System.Collections.Generic.List[string]
foreach ($o in $Outputs) {
    $diffPath = Join-Path $RepoRoot $o.Diff
    if (-not (Test-Path $diffPath)) { continue }
    Copy-Item $diffPath (Join-Path $diffsDir ("{0}-diff.txt" -f $o.Flavor)) -Force
    $content = Get-Content $diffPath -Raw
    if ($o.Flavor -eq 'Base') { $baseDiffText = $content }
    $lineCount = (Get-Content $diffPath | Measure-Object -Line).Lines
    $flavorLines.Add(("- **{0}**: {1} lines (see run artifacts)" -f $o.Flavor, $lineCount))
}

if (-not $baseDiffText) { $baseDiffText = '(no base diff was produced)' }
$truncNote = ''
if ($baseDiffText.Length -gt $CommentCharBudget) {
    $baseDiffText = $baseDiffText.Substring(0, $CommentCharBudget)
    $truncNote = "`n`n_…base diff truncated: full diff.txt files are attached to the workflow run._"
}

$body = @"
**This PR builds cleanly.** Below is the base winapp2.ini changelog it produces.

Per-flavor changelogs (full text attached to the workflow run):

$($flavorLines -join "`n")

<details><summary>Base winapp2.ini diff</summary>

``````
$baseDiffText
``````
$truncNote
</details>
"@
Complete-Verification -Status 'success' -ReportMarkdown $body
