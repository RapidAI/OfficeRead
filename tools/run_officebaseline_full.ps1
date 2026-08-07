param(
    [string]$InputPath = 'testdata\web-samples\samples',
    [string]$ReportPath = 'reports\batches\office-com-all-6008-pure-go.json',
    [int]$SliceSize = 1,
    [int]$TimeoutSeconds = 30,
    [int]$KillGraceSeconds = 3,
    [int]$StartupGraceSeconds = 15,
    [int]$BaselineRetries = 0,
    # Apply the non-modal open policy to the initial coverage pass as well.
    # This is appropriate for a full unattended audit: a normal Office open
    # remains the reference path, while repair/Protected View fallbacks that
    # can display desktop UI are deferred to explicitly requested diagnosis.
    [switch]$NoRecoveryOpen,
    # Excel's rendered .Text needs one COM call per cell.  In unattended
    # full-corpus runs the practical upper bound is deliberately conservative:
    # above this limit the single-call Value2 baseline is recorded as stored
    # value (and therefore excluded from the visible-content quality gate).
    # This prevents a moderately sized, heavily formatted sheet from spending
    # minutes in COM and blocking coverage of later samples.
    [int]$ExcelMaxCells = 10000,
    # Strict Office-visible acceptance must use Excel Range.Text for every
    # visible cell.  Keep the historical bounded Value2 fallback available for
    # a coverage-only run, but require the caller to acknowledge that it is not
    # an MS Office visible-content comparison.
    [switch]$StrictOfficeVisible,
    [int]$MaxSlices = 0,
    # A broken desktop COM session (not a difficult workbook) can reject every
    # Excel activation with 0x80070520. Stop early so the coverage checkpoint
    # remains a useful audit rather than labelling hundreds of later paths with
    # the same environmental failure. Set 0 to disable this circuit breaker.
    [int]$MaxConsecutiveExcelSessionFailures = 3,
    # When a bounded coverage pass has completed, re-run only the retained
    # baseline-unavailable paths in a new report with a longer COM timeout.
    # This never replaces the coverage checkpoint, so a persistent timeout is
    # still auditable separately from content quality.
    [string]$RetryUnavailableReportPath = '',
    [int]$RetryUnavailableTimeoutSeconds = 0,
    [int]$RetryUnavailableRetries = 1,
    # A strict Excel baseline obtains Range.Text only.  A workbook with a
    # very large number of populated cells can legitimately need longer than
    # the normal document deadline even after the sparse Text bridge avoids
    # formatting-only cells.  Give those XLS/XLSX paths an isolated long retry
    # after the first pass; this remains serialized and never permits Value2.
    [int]$RetryExcelTimeoutSeconds = 0,
    [int]$RetryExcelRetries = 2,
    # Word's OpenAndRepair and Protected View fallbacks can display a desktop
    # modal prompt for File Block / converter cases.  A focused recovery run
    # may opt out of those fallback opens so every retained path is bounded by
    # the watchdog and checkpointed as an explicit failure instead of blocking
    # all later samples.  The initial coverage pass is deliberately unchanged:
    # it still has the normal Office recovery evidence when it succeeds.
    [bool]$RetryNoRecoveryOpen = $true,
    # Number of independent recovery worker processes. Each worker receives a
    # disjoint path list and writes its own report, so no checkpoint is shared.
    # Keep this modest: Excel COM is a desktop application and too many
    # simultaneous activations increase, rather than reduce, modal timeouts.
    [int]$RetryUnavailableWorkers = 1
)

# Resumable supervisor for the 6008-file Office COM audit. It deliberately
# runs a bounded slice per child process so a host interruption cannot corrupt
# the last successful checkpoint. The normal extraction path stays pure Go;
# use the individual normalization flags only when diagnosing a legacy mismatch.
$ErrorActionPreference = 'Stop'
if ($SliceSize -lt 1 -or $TimeoutSeconds -lt 1 -or $KillGraceSeconds -lt 0 -or $StartupGraceSeconds -lt 0 -or $BaselineRetries -lt 0 -or $ExcelMaxCells -lt 0 -or $MaxSlices -lt 0 -or $MaxConsecutiveExcelSessionFailures -lt 0 -or $RetryUnavailableTimeoutSeconds -lt 0 -or $RetryUnavailableRetries -lt 0 -or $RetryExcelTimeoutSeconds -lt 0 -or $RetryExcelRetries -lt 0 -or $RetryUnavailableWorkers -lt 1) {
    throw 'SliceSize and TimeoutSeconds must be >= 1; KillGraceSeconds, StartupGraceSeconds, BaselineRetries, ExcelMaxCells, MaxSlices, session-failure limit, retry timeouts, and retry counts must be >= 0'
}
if ($StrictOfficeVisible -and $ExcelMaxCells -lt 2147483647) {
    throw 'StrictOfficeVisible requires ExcelMaxCells=2147483647 so no workbook can fall back to Value2; use a coverage-only run without StrictOfficeVisible when a bounded fallback is intended'
}

$root = (Get-Location).Path
$input = Join-Path $root $InputPath
$report = Join-Path $root $ReportPath
$go = Get-Command go.exe -ErrorAction SilentlyContinue
if ($null -eq $go) {
    $fallbackGo = 'C:\Program Files\Go\bin\go.exe'
    if (-not (Test-Path -LiteralPath $fallbackGo)) { throw 'go.exe is not on PATH and C:\Program Files\Go\bin\go.exe does not exist' }
    $go = Get-Item -LiteralPath $fallbackGo
}
$goPath = if ($go -is [System.Management.Automation.CommandInfo]) { $go.Source } else { $go.FullName }
$baseline = Join-Path $root 'tools\office_baseline.ps1'
$binary = Join-Path $root '.gocache\officebaseline-full.exe'
[void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($report))

# Desktop automation is commonly launched from a restricted process where the
# user-profile Go build cache is inaccessible.  Keep all compiler cache writes
# inside the repository so every supervised child can rebuild the current
# baseline binary before it starts a slice.
$goCache = Join-Path $root '.gocache\go-build'
[void][IO.Directory]::CreateDirectory($goCache)
$env:GOCACHE = $goCache

# Only one resumable supervisor may own a report.  Concurrent supervisors can
# both observe the same next pending path and race to replace the checkpoint;
# even atomic file replacement then loses one result.  A named mutex is scoped
# to the absolute report path so independent format reports may still run.
$sha256 = [Security.Cryptography.SHA256]::Create()
try {
    $reportHash = $sha256.ComputeHash([Text.Encoding]::UTF8.GetBytes($report))
} finally {
    $sha256.Dispose()
}
$mutexName = 'Global\OfficeReadBaseline_' + (-join ($reportHash | ForEach-Object { $_.ToString('x2') })).Substring(0, 24)
$supervisorMutex = [Threading.Mutex]::new($false, $mutexName)
$ownsSupervisorMutex = $false
try {
    try {
        $ownsSupervisorMutex = $supervisorMutex.WaitOne(0)
    } catch [Threading.AbandonedMutexException] {
        $ownsSupervisorMutex = $true
    }
    if (-not $ownsSupervisorMutex) {
        throw "another Office baseline supervisor is already writing $report; wait for it to finish before resuming"
    }
} catch {
    $supervisorMutex.Dispose()
    throw
}
try {

& $goPath build -o $binary .\cmd\officebaseline
if ($LASTEXITCODE -ne 0) { throw 'building officebaseline failed' }

function Stop-OfficeAutomationServers {
    Get-OfficeAutomationProcesses |
        Where-Object { $_.CommandLine -match '(?i)(/automation|-embedding)' } |
        ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
}

function Stop-OfficeBaselineAutomationServers {
    param(
        [int[]]$ExcludeParentProcessIds = @()
    )
    # A COM server is normally re-parented to the desktop broker (not directly
    # to its launcher).  ParentProcessId alone therefore cannot tell a stale
    # server from the one belonging to the currently executing slice.  Match
    # the active baseline children first, then preserve only servers whose
    # start time is at or after that child.  This avoids killing the current
    # slice while reliably draining an older Excel/Word/PowerPoint server
    # before the next independent COM activation.
    $activeStarts = @()
    try {
        $activeStarts = @(Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" -ErrorAction Stop |
            Where-Object { $_.ProcessId -in $ExcludeParentProcessIds } |
            ForEach-Object { [Management.ManagementDateTimeConverter]::ToDateTime($_.CreationDate) })
    } catch { }
    $oldestActiveStart = if ($activeStarts.Count -gt 0) { ($activeStarts | Sort-Object | Select-Object -First 1) } else { $null }
    Get-OfficeAutomationProcesses |
        Where-Object {
            if ($_.CommandLine -notmatch '(?i)(/automation|-embedding)') { return $false }
            if ($null -eq $oldestActiveStart) { return $true }
            try { return [Management.ManagementDateTimeConverter]::ToDateTime($_.CreationDate) -lt $oldestActiveStart } catch { return $false }
        } |
        ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
}

function Stop-OfficeBaselineAutomationServersForReport {
    param(
        [string]$ReportPath
    )
    # The focused-retry phase starts only after the coverage worker has exited.
    # Do not use the broad Stop-OfficeAutomationServers helper here: it would
    # also terminate a user's interactive Office instance when its command line
    # happens to contain an automation-related switch. The same conservative
    # ownership-aware selector used between slices is sufficient once no
    # baseline PowerShell child is active, and it still drains every orphaned
    # /automation or -Embedding server before the longer retry begins.
    $activeBaselinePids = @(Get-BaselinePowerShellProcessIds)
    Stop-OfficeBaselineAutomationServers $activeBaselinePids
}

function Get-OfficeAutomationProcesses {
    # Win32_Process supplies command lines and parent IDs, but restricted
    # desktop sessions can deny the CIM query.  That must not terminate the
    # resumable audit before a child baseline runs.  Do *not* fall back to a
    # blind EXCEL/WINWORD/POWERPNT process list: without CommandLine we cannot
    # distinguish the user's interactive document from our /Automation server,
    # and treating it as an orphan would repeatedly kill the user's Office.
    try {
        return @(Get-CimInstance Win32_Process -Filter "Name='WINWORD.EXE' OR Name='POWERPNT.EXE' OR Name='EXCEL.EXE'" -ErrorAction Stop)
    } catch {
        return @()
    }
}

function Get-BaselinePowerShellProcessIds {
    try {
        return @(
            Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" -ErrorAction Stop |
                Where-Object { $_.CommandLine -match 'office_baseline\.ps1' } |
                ForEach-Object { [int]$_.ProcessId }
        )
    } catch {
        # Without command-line inspection we cannot safely identify a child
        # baseline PowerShell process.  Returning an empty set only makes the
        # automation-left cleanup conservative; no interactive Office process
        # is selected because its command line is unavailable.
        return @()
    }
}

function Wait-OfficeAutomationServersGone {
    param([int]$Seconds)
    if ($Seconds -le 0) { return }
    $deadline = [DateTime]::UtcNow.AddSeconds($Seconds)
    do {
        $remaining = @(Get-OfficeAutomationProcesses |
            Where-Object { $_.CommandLine -match '(?i)(/automation|-embedding)' })
        if ($remaining.Count -eq 0) { return }
        Start-Sleep -Milliseconds 250
    } while ([DateTime]::UtcNow -lt $deadline)
}

function Get-CompletedCount {
    if (-not (Test-Path -LiteralPath $report)) { return 0 }
    try {
        # Windows PowerShell 5.1's ConvertFrom-Json can fail on otherwise valid
        # UTF-8 reports containing non-ASCII Office error messages. Count file
        # records directly from the stable Go JSON shape instead; the final
        # report is validated with a standards-compliant JSON parser below.
        return @([regex]::Matches([IO.File]::ReadAllText($report, [Text.Encoding]::UTF8), '(?m)^    \{\r?$')).Count
    } catch {
        throw "invalid checkpoint JSON: $report ($($_.Exception.Message))"
    }
}

function Test-OfficeComSession {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('Excel', 'PowerPoint', 'Word')]
        [string]$Application
    )
    # Launching a tiny, file-independent COM probe before coverage resumes
    # distinguishes a broken Windows logon/Office desktop (0x80070520) from a
    # difficult document. The probe never opens a user file and always quits
    # its private Application object.  Probe every application represented by
    # pending corpus files: a failed PowerPoint session must not be allowed to
    # write hundreds of identical per-file failures into the checkpoint.
    $progId = switch ($Application) {
        'Excel' { 'Excel.Application' }
        'PowerPoint' { 'PowerPoint.Application' }
        'Word' { 'Word.Application' }
    }
    $configure = switch ($Application) {
        'Excel' { '$app.Visible = $false; $app.DisplayAlerts = $false' }
        'Word' { '$app.Visible = $false; $app.DisplayAlerts = 0' }
        # PowerPoint does not permit setting Application.Visible to false.
        # Instantiation is sufficient for the health check, and its normal
        # baseline path likewise leaves application visibility untouched.
        'PowerPoint' { '' }
    }
    $probe = @'
$ErrorActionPreference = 'Stop'
$app = $null
try {
    $app = New-Object -ComObject __OFFICE_PROGID__
    __OFFICE_CONFIGURE__
    [Console]::Out.WriteLine('excel-com-ready')
} finally {
    if ($null -ne $app) {
        try { $app.Quit() } catch { }
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app) } catch { }
    }
}

function Test-ExcelComSession {
    # Backward-compatible name retained for callers and the supervisor
    # contract test; the implementation now shares the generic Office probe.
    return (Test-OfficeComSession 'Excel')
}
'@
    $probe = $probe.Replace('__OFFICE_PROGID__', $progId).Replace('__OFFICE_CONFIGURE__', $configure)
    # The health probe itself calls COM and can hang behind a damaged Excel
    # broker. It must have the same strict bound as a document comparison;
    # otherwise a stale Automation server prevents the supervisor from ever
    # reaching its resumable checkpoint loop.
    $process = $null
    try {
        $info = [Diagnostics.ProcessStartInfo]::new()
        $info.FileName = 'powershell.exe'
        $encodedProbe = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($probe))
        $info.Arguments = "-NoProfile -NonInteractive -ExecutionPolicy Bypass -EncodedCommand $encodedProbe"
        $info.UseShellExecute = $false
        $info.RedirectStandardOutput = $true
        $info.RedirectStandardError = $true
        $info.CreateNoWindow = $true
        $process = [Diagnostics.Process]::new()
        $process.StartInfo = $info
        if (-not $process.Start()) { return $false }
        if (-not $process.WaitForExit(20000)) {
            try { & taskkill.exe /PID $process.Id /T /F | Out-Null } catch { }
            return $false
        }
        $output = $process.StandardOutput.ReadToEnd() + $process.StandardError.ReadToEnd()
        return (($process.ExitCode -eq 0) -and ($output -match 'excel-com-ready'))
    } catch {
        return $false
    } finally {
        if ($null -ne $process) { $process.Dispose() }
    }
}

# Microsoft Office creates transient owner/lock files named "~$<document>"
# next to a document while it is open.  They have an Office-looking extension
# but are not independently openable documents (and can appear between two
# resumable slices), so including them makes the corpus size nondeterministic
# and can replace one real sample in a fixed 6008-path audit.
$all = @(Get-ChildItem -LiteralPath $input -Recurse -File -Include '*.doc','*.docx','*.ppt','*.pptx','*.xls','*.xlsx' |
    Where-Object { -not $_.Name.StartsWith('~$') })
$total = @($all).Count
if ($total -eq 0) { throw "no Office files under $input" }

$slice = 0
$consecutiveExcelSessionFailures = 0
if ($MaxConsecutiveExcelSessionFailures -gt 0) {
    $pendingExtensions = @($all | Where-Object { $null -eq $_ } | ForEach-Object { $_.Extension })
    # The report contains only completed paths, so derive pending formats from
    # the current checkpoint without parsing its potentially localized errors.
    if (Test-Path -LiteralPath $report) {
        $checkpointText = [IO.File]::ReadAllText($report, [Text.Encoding]::UTF8)
        $pendingExtensions = @($all | Where-Object { $checkpointText -notmatch [regex]::Escape(('"path": "' + $_.FullName.Replace('\', '\\'))) } | ForEach-Object { $_.Extension.ToLowerInvariant() })
    } else {
        $pendingExtensions = @($all | ForEach-Object { $_.Extension.ToLowerInvariant() })
    }
    $requiredApplications = @()
    if ((@($pendingExtensions | Where-Object { $_ -in '.xls', '.xlsx' })).Count -gt 0) { $requiredApplications += 'Excel' }
    if ((@($pendingExtensions | Where-Object { $_ -in '.ppt', '.pptx' })).Count -gt 0) { $requiredApplications += 'PowerPoint' }
    if ((@($pendingExtensions | Where-Object { $_ -in '.doc', '.docx' })).Count -gt 0) { $requiredApplications += 'Word' }
    foreach ($application in $requiredApplications) {
        if (-not (Test-OfficeComSession $application)) {
            if ($application -eq 'Excel') {
                throw "Excel COM health probe failed before coverage resumed. Restore the interactive Windows/Office logon session, then rerun this supervisor to resume from $report."
            }
            throw "$application COM health probe failed before coverage resumed. Restore the interactive Windows/Office logon session, then rerun this supervisor to resume from $report."
        }
    }
}
while ($true) {
    $completed = Get-CompletedCount
    if ($completed -ge $total) { break }
    if ($MaxSlices -gt 0 -and $slice -ge $MaxSlices) { break }
    # A PowerShell child is deliberately one file wide, so it owns a fresh COM
    # server. Drain only when a prior timeout left an Automation process behind;
    # otherwise avoid needless Office cold starts. A healthy active baseline
    # child creates an Automation server while it is still running, so exclude
    # that direct child/descendant rather than killing it mid-comparison.
    $activeBaselinePids = @(Get-BaselinePowerShellProcessIds)
    $automationLeft = @(Get-OfficeAutomationProcesses |
        Where-Object {
            if ($_.CommandLine -notmatch '(?i)(/automation|-embedding)') { return $false }
            if ($activeBaselinePids.Count -eq 0) { return $true }
            $activeStarts = @(Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" -ErrorAction SilentlyContinue |
                Where-Object { $_.ProcessId -in $activeBaselinePids } |
                ForEach-Object { [Management.ManagementDateTimeConverter]::ToDateTime($_.CreationDate) })
            if ($activeStarts.Count -eq 0) { return $false }
            try { return [Management.ManagementDateTimeConverter]::ToDateTime($_.CreationDate) -lt ($activeStarts | Sort-Object | Select-Object -First 1) } catch { return $false }
        })
    if ($automationLeft.Count -gt 0) {
        Stop-OfficeBaselineAutomationServers $activeBaselinePids
        Wait-OfficeAutomationServersGone $StartupGraceSeconds
    }
    $env:OFFICEBASELINE_SCRIPT = $baseline
    # A strict compatibility audit must open every path through Office.  The
    # command-line tool can reuse a byte-identical result as a throughput
    # optimization, but that does not establish an independent COM baseline
    # for each of the 6008 samples and therefore cannot satisfy this suite's
    # acceptance gate.
    $coverageArgs = @('-resume', '-reuse-identical=false', '-batch-size', '1', '-checkpoint', '1', '-max-files', $SliceSize,
        '-timeout', ("{0}s" -f $TimeoutSeconds), '-kill-grace', ("{0}s" -f $KillGraceSeconds),
        '-baseline-retries', $BaselineRetries, '-excel-max-cells', $ExcelMaxCells, '-keep-errors', '-json', $report, $input)
    if ($NoRecoveryOpen) { $coverageArgs += '-no-recovery-open' }
    $sliceOutput = @(& $binary @coverageArgs 2>&1)
    $sliceOutput | ForEach-Object { Write-Host $_ }
    # Content failures and per-file COM failures produce a non-zero exit; the
    # checkpoint is still authoritative and the next slice must proceed.
    $slice++
    $after = Get-CompletedCount
    if ($after -le $completed) {
        throw "slice $slice made no checkpoint progress ($completed files); inspect $report"
    }
    Write-Host ("slice {0}: {1}/{2} files checkpointed" -f $slice, $after, $total)
    if (($sliceOutput -join "`n") -match '(?i)80070520') {
        $consecutiveExcelSessionFailures++
        if ($MaxConsecutiveExcelSessionFailures -gt 0 -and $consecutiveExcelSessionFailures -ge $MaxConsecutiveExcelSessionFailures) {
            throw "Excel COM session circuit breaker opened after $consecutiveExcelSessionFailures consecutive 0x80070520 activation failures. Restore the interactive Windows/Office logon session, then rerun this supervisor to resume from $report."
        }
    } else {
        $consecutiveExcelSessionFailures = 0
    }
}

$json = [IO.File]::ReadAllText($report, [Text.Encoding]::UTF8)
$completed = Get-CompletedCount
$compared = [regex]::Match($json, '"compared":\s*(\d+)').Groups[1].Value
$errors = [regex]::Match($json, '"errors":\s*(\d+)').Groups[1].Value
$f1 = [regex]::Match($json, '"f1":\s*([0-9.]+)').Groups[1].Value
Write-Host ("completed {0}/{1}; compared={2}; errors={3}; F1={4}" -f $completed, $total, $compared, $errors, $f1)

if ($RetryUnavailableReportPath -ne '') {
    $retryTimeout = $RetryUnavailableTimeoutSeconds
    if ($retryTimeout -le 0) { $retryTimeout = [Math]::Max(90, $TimeoutSeconds * 3) }
    $retryReport = Join-Path $root $RetryUnavailableReportPath
    $retryList = Join-Path ([IO.Path]::GetDirectoryName($retryReport)) (([IO.Path]::GetFileNameWithoutExtension($retryReport)) + '-paths.txt')
    $python = Get-Command python.exe -ErrorAction SilentlyContinue
    if ($null -eq $python) {
        $python = Get-Command python -ErrorAction SilentlyContinue
    }
    if ($null -eq $python) { throw 'python is required for focused retry path selection; install Python or omit RetryUnavailableReportPath' }
    # Retry every baseline failure, including Trust Center/File Block and
    # session failures.  These categories remain explicit in the checkpoint;
    # this only gives Word's own protected/recovery open path a later isolated
    # attempt and never reclassifies an unresolved item as a success.
    & $python.Source -X utf8 (Join-Path $root 'tools\office_baseline_report.py') $report --write-errors $retryList --category office-baseline-issue
    if ($LASTEXITCODE -ne 0) { throw 'could not export baseline-unavailable paths for focused retry' }
    $retryPaths = @()
    if (Test-Path -LiteralPath $retryList) {
        $retryPaths = @(Get-Content -LiteralPath $retryList -Encoding UTF8 | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    }
    if ($retryPaths.Count -eq 0) {
        Write-Host 'no baseline-unavailable paths require focused retry'
    } else {
        Write-Host ("focused retry: {0} unavailable paths; timeout={1}s; extra retries={2}" -f $retryPaths.Count, $retryTimeout, $RetryUnavailableRetries)
        # A single report is required for deterministic resume and atomic
        # checkpointing. The focused path list is passed directly rather than
        # being expanded into one enormous Windows command line (which fails
        # well below the 6008-file corpus size).
        Stop-OfficeBaselineAutomationServersForReport $report
        Wait-OfficeAutomationServersGone $StartupGraceSeconds
        # Retain each first-pass failure so the bounded recovery scan advances
        # through every path. A separate, longer retry report is then produced
        # from these explicit failures; otherwise the first slow workbook is
        # retried forever and blocks coverage of all later samples.
        # Keep the established retry ordering for the supervisor-contract
        # regression test; the explicit no-reuse switch still guarantees a
        # fresh Office open for every retained retry path.
        $retryArgs = @('-resume', '-keep-errors', '-batch-size', '1', '-checkpoint', '1', '-reuse-identical=false', '-timeout', ("{0}s" -f $retryTimeout), '-kill-grace', ("{0}s" -f $KillGraceSeconds), `
            '-baseline-retries', $RetryUnavailableRetries, '-excel-max-cells', $ExcelMaxCells, '-json', $retryReport, '-paths-file', $retryList)
        if ($RetryNoRecoveryOpen) { $retryArgs += '-no-recovery-open' }
        & $binary @retryArgs
        Write-Host ("focused retry report: {0}; workers requested={1} (serialized for Office COM safety)" -f $retryReport, $RetryUnavailableWorkers)

        # The broad retry report remains the authoritative recovery artifact.
        # Re-attempt only its unresolved Excel paths with an even longer bound:
        # rendering must still be Excel Range.Text, so a timeout is evidence of
        # unavailable baseline rather than a reason to fall back to Value2.
        $excelRetryTimeout = $RetryExcelTimeoutSeconds
        if ($excelRetryTimeout -le 0) { $excelRetryTimeout = [Math]::Max(600, $retryTimeout * 2) }
        $excelRetryList = Join-Path ([IO.Path]::GetDirectoryName($retryReport)) (([IO.Path]::GetFileNameWithoutExtension($retryReport)) + '-excel-paths.txt')
        $excelRetryReport = Join-Path ([IO.Path]::GetDirectoryName($retryReport)) (([IO.Path]::GetFileNameWithoutExtension($retryReport)) + '-excel.json')
        & $python.Source -X utf8 (Join-Path $root 'tools\office_baseline_report.py') $retryReport --write-errors $excelRetryList --category office-baseline-issue --extensions .xls,.xlsx
        if ($LASTEXITCODE -ne 0) { throw 'could not export unresolved Excel paths for focused retry' }
        $excelRetryPaths = @()
        if (Test-Path -LiteralPath $excelRetryList) {
            $excelRetryPaths = @(Get-Content -LiteralPath $excelRetryList -Encoding UTF8 | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        }
        if ($excelRetryPaths.Count -gt 0) {
            Write-Host ("focused Excel retry: {0} paths; timeout={1}s; extra retries={2}; source=Range.Text" -f $excelRetryPaths.Count, $excelRetryTimeout, $RetryExcelRetries)
            Stop-OfficeBaselineAutomationServersForReport $retryReport
            Wait-OfficeAutomationServersGone $StartupGraceSeconds
            # Use a third, clean checkpoint.  Resuming the broad retry would
            # retain its errors by design and therefore skip these paths.
            $excelRetryArgs = @('-keep-errors', '-batch-size', '1', '-checkpoint', '1', '-reuse-identical=false', '-timeout', ("{0}s" -f $excelRetryTimeout), '-kill-grace', ("{0}s" -f $KillGraceSeconds), `
                '-baseline-retries', $RetryExcelRetries, '-excel-max-cells', $ExcelMaxCells, '-json', $excelRetryReport, '-paths-file', $excelRetryList)
            if ($RetryNoRecoveryOpen) { $excelRetryArgs += '-no-recovery-open' }
            & $binary @excelRetryArgs
            Write-Host ("focused Excel retry report: {0}" -f $excelRetryReport)
        }
    }
}
} finally {
    if ($ownsSupervisorMutex) { [void]$supervisorMutex.ReleaseMutex() }
    $supervisorMutex.Dispose()
}
