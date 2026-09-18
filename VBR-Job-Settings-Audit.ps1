<#
.SYNOPSIS
    M365admintools.com - Author Charles Arconi - updated 7/7/2026
    Hub-and-spoke job settings audit for Veeam Backup and Replication v13,
    Veeam Backup for Azure, and Veeam Data Cloud for M365. Produces a single
    HTML dashboard listing every job's settings and flagging any deviation
    from the documented Job Design Standard.

.DESCRIPTION
    This script pulls current job configuration from three platforms in a
    Forgent Power style hub-and-spoke Veeam deployment and checks each job
    against the standard documented in the operations runbook, Section 8.2
    (Job Design Standard) and Section 1.5 (VBA Backup Policy).

    Platform 1, on-premises VBR. Reads all backup jobs and backup copy jobs
    from the local (or specified) VBR server using the Veeam.Backup.PowerShell
    module. Checked directly against live configuration.

    Platform 2, Veeam Backup for Azure (VBA). Reads policies from the VBA
    appliance REST API. Checked directly against live configuration.

    Platform 3, Veeam Data Cloud for M365. Veeam does not currently publish a
    customer-facing REST API for pulling M365 policy settings out of the Data
    Cloud console, so this section reads from a small JSON file you maintain
    by hand after checking the console. Run -SeedDataCloudTemplate once to
    generate a starter file, then update it whenever a Data Cloud policy
    changes. This section is clearly marked "manual" in the report so it is
    never mistaken for a live check.

.PARAMETER VbrServer
    VBR server to connect to. Default is localhost. Because MGMCVEEAM is a
    workgroup server, run this script directly on the VBR console rather
    than remotely unless you have already worked through the TrustedHosts
    considerations in runbook Section 7.1.

.PARAMETER VbrCredential
    Optional PSCredential for Connect-VBRServer when VbrServer is not
    localhost. Not needed for a local run.

.PARAMETER VbaApplianceUrl
    Base HTTPS URL of the Veeam Backup for Azure appliance REST API, for
    example https://10.10.16.10:4443. Omit to skip the VBA section.

.PARAMETER VbaCredential
    PSCredential for a VBA local or SSO user with rights to read policies.
    Required if VbaApplianceUrl is supplied.

.PARAMETER DataCloudInputPath
    Path to the Data Cloud M365 policy JSON file. Defaults to
    .\DataCloud-M365-Policies.json in the script folder.

.PARAMETER SeedDataCloudTemplate
    Writes a starter Data Cloud JSON file to DataCloudInputPath, pre-filled
    with the two policies already on record (PwrQ-1-Year, Gold-6-week), and
    exits without running the audit.

.PARAMETER OutputPath
    Path for the generated HTML report. Defaults to a timestamped file in
    the script folder.

.PARAMETER Diagnose
    Pulls one on-prem job, runs Get-VBRJobOptions and Get-VBRJobScheduleOptions
    against it, and dumps every property name with Get-Member. Run this once
    against your specific VBR 13.x build before trusting the flagged report.
    Veeam's PowerShell object model has changed property names across
    releases, and this script cannot be tested against a live server ahead
    of time. Every property read below is wrapped in a try/catch, so a
    mismatch shows "n/a" in the report rather than stopping the script, but
    -Diagnose is the fast way to confirm the property names are current
    before you rely on the flags for real.

.NOTES
    Requires the Veeam.Backup.PowerShell module (ships with VBR console) for
    the on-prem section. Requires PowerShell 5.1 or later for the VBA REST
    calls (Invoke-RestMethod).

.EXAMPLE
    .\VBR-Job-Settings-Audit.ps1 -Diagnose

.EXAMPLE
    .\VBR-Job-Settings-Audit.ps1 -VbaApplianceUrl https://10.10.16.10:4443 -VbaCredential (Get-Credential)

.EXAMPLE
    .\VBR-Job-Settings-Audit.ps1 -SeedDataCloudTemplate
#>

[CmdletBinding()]
param(
    [string]$VbrServer = 'localhost',
    [System.Management.Automation.PSCredential]$VbrCredential,

    [string]$VbaApplianceUrl,
    [System.Management.Automation.PSCredential]$VbaCredential,

    [string]$DataCloudInputPath = (Join-Path $PSScriptRoot 'DataCloud-M365-Policies.json'),
    [switch]$SeedDataCloudTemplate,

    [string]$OutputPath = (Join-Path $PSScriptRoot ("VBR-Job-Audit-{0}.html" -f (Get-Date -Format 'yyyyMMdd-HHmm'))),

    [switch]$Diagnose,

    # Veeam emits "This cmdlet is no longer supported for Backup Copy jobs"
    # during collection in environments that have backup copy jobs. Adding
    # -WarningAction to the individual Get-VBRJob call did not suppress it,
    # and the exact cmdlet raising it has not been pinned down, so the
    # script sets $WarningPreference for the collection phase instead.
    # Pass -ShowVeeamWarnings to see Veeam's warnings again, which is worth
    # doing if the report ever looks wrong in a way this might explain.
    [switch]$ShowVeeamWarnings
)

if (-not $ShowVeeamWarnings) { $WarningPreference = 'SilentlyContinue' }
$script:AnyDayBasedRetention = $false

# ---------------------------------------------------------------------------
# Documented standard, sourced from the Forgent Power Operations Runbook
# Section 8.2 (on-prem job design) and Section 1.5 (VBA backup policy).
# Adjust these values here if the documented standard changes, rather than
# hunting through the check logic further down.
# ---------------------------------------------------------------------------
$Standard = [ordered]@{
    OnPremDailyRetentionCycles   = 14
    OnPremWeeklyGfsCount         = 4
    OnPremMonthlyGfsCount        = 3
    JobStaggerMinutesMin         = 10
    JobStaggerMinutesMax         = 15
    BackupCopyModeExpected       = 'Immediate'
    BackupCopyEncryptionRequired = $true
    OneJobPerHyperVHost          = $true
    VbaDailyRetentionDays        = 14
    VbaWeeklyRetentionMonths     = 1
    VbaMonthlyRetentionMonths    = 12
    VbaImmutabilityRequired      = $true
    VbaAppAwareRequired          = $true
    # Extend this list if your environment uses other offsite naming
    # conventions. Matching is a simple wildcard contains check.
    # VeeamBackupCommerce confirmed as Commerce's offsite Azure target (its
    # copy jobs mirror the WACO/STATES/IEAG pattern exactly, just without
    # "Azure Blob" in the name). Rename the repository to match the other
    # sites' convention if you'd rather close this gap that way instead.
    CloudRepoNameHints           = @('Data Cloud', 'DataCloud', 'VDC', 'Cloud Connect', 'Azure Blob', 'Blob Storage', 'VeeamBackupCommerce')
}

# ---------------------------------------------------------------------------
# Helper: safe property read. Returns 'n/a' instead of throwing if a
# property path does not exist on this build's object model.
# ---------------------------------------------------------------------------
function Get-SafeValue {
    param([scriptblock]$Expr)
    try {
        $result = & $Expr
        if ($null -eq $result) { return 'n/a' }
        return $result
    } catch {
        return 'n/a'
    }
}

# Tries each scriptblock in order, returns the first one that succeeds and
# returns a non-null value. Used where a property path is uncertain across
# VBR builds and there are a few plausible alternates worth trying before
# giving up and showing 'n/a'.
function Get-FirstSafeValue {
    param([scriptblock[]]$Candidates)
    foreach ($c in $Candidates) {
        $v = Get-SafeValue $c
        if ($v -ne 'n/a') { return $v }
    }
    return 'n/a'
}

# ---------------------------------------------------------------------------
# Seed a starter Data Cloud policy file and exit if requested
# ---------------------------------------------------------------------------
if ($SeedDataCloudTemplate) {
    $seed = @(
        [ordered]@{
            Repository   = 'vdcm365amerokx9eln1xidak'
            PolicyName   = 'PwrQ-1-Year'
            Location     = 'Azure (US) East US'
            Status       = 'ACTIVE'
            Tenant       = 'alliedpowerandcontrolcom.onmicrosoft.com'
            RetentionSummary = 'Confirm current retention in the Data Cloud console'
            LastVerified = (Get-Date -Format 'yyyy-MM-dd')
        },
        [ordered]@{
            Repository   = 'vdcm365amer9cymqi7t08hr8'
            PolicyName   = 'Gold-6-week'
            Location     = 'Azure (US) East US'
            Status       = 'INACTIVE'
            Tenant       = 'alliedpowerandcontrolcom.onmicrosoft.com'
            RetentionSummary = 'Legacy 6 week policy, retained for its retention window, not in active use'
            LastVerified = (Get-Date -Format 'yyyy-MM-dd')
        }
    )
    $seed | ConvertTo-Json -Depth 4 | Out-File -FilePath $DataCloudInputPath -Encoding UTF8
    Write-Host "Data Cloud template written to $DataCloudInputPath. Edit it, then re-run without -SeedDataCloudTemplate." -ForegroundColor Green
    return
}

# ---------------------------------------------------------------------------
# Connect to VBR and load the module
# ---------------------------------------------------------------------------
function Connect-Vbr {
    param([string]$ServerName, [System.Management.Automation.PSCredential]$Cred)

    if (-not (Get-Module -Name Veeam.Backup.PowerShell)) {
        try {
            Import-Module Veeam.Backup.PowerShell -ErrorAction Stop -DisableNameChecking
        } catch {
            throw "Could not load Veeam.Backup.PowerShell. Run this script on the VBR console server, or from a machine with the Veeam PowerShell module installed. $($_.Exception.Message)"
        }
    }

    if ($ServerName -ne 'localhost') {
        if ($Cred) {
            Connect-VBRServer -Server $ServerName -Credential $Cred -ErrorAction Stop
        } else {
            Connect-VBRServer -Server $ServerName -ErrorAction Stop
        }
    }
}

# ---------------------------------------------------------------------------
# Diagnose mode: dump property names for one job so paths can be confirmed
# ---------------------------------------------------------------------------
if ($Diagnose) {
    Connect-Vbr -ServerName $VbrServer -Cred $VbrCredential
    $sample = Get-VBRJob | Select-Object -First 1
    if (-not $sample) {
        Write-Host "No jobs found to diagnose against." -ForegroundColor Yellow
        return
    }
    Write-Host "`nDiagnosing against job: $($sample.Name)" -ForegroundColor Cyan

    Write-Host "`n--- Get-VBRJobOptions members ---" -ForegroundColor Cyan
    Get-VBRJobOptions -Job $sample | Get-Member -MemberType Property

    Write-Host "`n--- Get-VBRJobOptions.BackupStorageOptions members ---" -ForegroundColor Cyan
    try {
        (Get-VBRJobOptions -Job $sample).BackupStorageOptions | Get-Member -MemberType Property
    } catch { Write-Host "BackupStorageOptions not present on this object." -ForegroundColor Yellow }

    Write-Host "`n--- Get-VBRJobScheduleOptions members ---" -ForegroundColor Cyan
    Get-VBRJobScheduleOptions -Job $sample | Get-Member -MemberType Property

    Write-Host "`n--- Get-VBRJobObject members (first object) ---" -ForegroundColor Cyan
    $obj = Get-VBRJobObject -Job $sample | Select-Object -First 1
    if ($obj) { $obj | Get-Member -MemberType Property }

    $bc = Get-VBRBackupCopyJob | Select-Object -First 1
    if ($bc) {
        Write-Host "`n--- Get-VBRBackupCopyJob members (sample: $($bc.Name)) ---" -ForegroundColor Cyan
        $bc | Get-Member -MemberType Property
    } else {
        Write-Host "`nNo backup copy jobs found to diagnose." -ForegroundColor Yellow
    }

    Write-Host "`nCompare these property names against the ones used in the Get-OnPremAudit function, and edit that function if anything has moved." -ForegroundColor Green
    return
}

# ---------------------------------------------------------------------------
# Platform 1: on-prem VBR audit
# ---------------------------------------------------------------------------
function Get-OnPremAudit {
    param($Standard)

    Connect-Vbr -ServerName $VbrServer -Cred $VbrCredential

    $rows = [System.Collections.Generic.List[object]]::new()
    $globalFlags = [System.Collections.Generic.List[string]]::new()

    # --- Primary backup jobs -------------------------------------------------
    # Get-VBRBackupCopyJob results also show up in Get-VBRJob with a JobType
    # that matches the loose 'Backup' pattern (copy jobs are internally typed
    # as a kind of backup job too), which double counted every copy job as a
    # primary job. Build the copy job ID set first and exclude it explicitly,
    # rather than relying on JobType string matching alone.
    $copyJobsForDedupe = Get-VBRBackupCopyJob -ErrorAction SilentlyContinue
    $copyJobIds = @($copyJobsForDedupe | ForEach-Object { Get-SafeValue { $_.Id } } | Where-Object { $_ -ne 'n/a' })

    $primaryJobs = Get-VBRJob -ErrorAction SilentlyContinue -WarningAction SilentlyContinue |
        Where-Object { $_.JobType -match 'Backup' } |
        Where-Object {
            $thisId = Get-SafeValue { $_.Id }
            $thisId -eq 'n/a' -or ($copyJobIds -notcontains $thisId)
        }

    foreach ($job in $primaryJobs) {

        $flags = [System.Collections.Generic.List[string]]::new()

        $jobOpts   = Get-SafeValue { Get-VBRJobOptions -Job $job }
        $schedOpts = Get-SafeValue { Get-VBRJobScheduleOptions -Job $job }
        $objs      = Get-SafeValue { Get-VBRJobObject -Job $job }

        # RetainCycles sits at Veeam's default of 7 on jobs that actually use
        # day-based retention, so reading it blindly reports a vestigial
        # value. Check RetentionType first and read the matching field.
        $retainCycles = 'n/a'
        $retentionType = 'n/a'
        if ($jobOpts -ne 'n/a') {
            $retentionType = Get-SafeValue { $jobOpts.BackupStorageOptions.RetentionType }
            if ($retentionType -eq 'Days') {
                $retainCycles = Get-SafeValue { $jobOpts.BackupStorageOptions.RetainDaysToKeep }
            } else {
                $retainCycles = Get-SafeValue { $jobOpts.BackupStorageOptions.RetainCycles }
            }
        }

        # Display unit matters: the documented standard is 14 restore points,
        # so a job set to 14 days is a near match but not an exact one.
        $retentionUnit = if ($retentionType -eq 'Days') { 'days' }
                         elseif ($retentionType -eq 'n/a') { '' }
                         else { 'restore points' }
        $retentionDisplay = if ($retainCycles -eq 'n/a') { 'n/a' }
                            else { (("$retainCycles $retentionUnit").Trim()) }
        if ($retentionType -eq 'Days') { $script:AnyDayBasedRetention = $true }

        $startTime = Get-FirstSafeValue @(
            { $schedOpts.StartDateTimeLocal },
            { $schedOpts.OptionsDaily.Periods },
            { $schedOpts.NextRun }
        )

        $repoName = Get-SafeValue { $job.GetBackupTargetRepository().Name }

        # Distinct protected hosts in this job's scope. Location returns a
        # "host\objectname" style path (or just "host" for a container entry),
        # so take the token before the backslash rather than the whole
        # string, or every VM on one host counts as a separate host.
        $hostNames = @()
        if ($objs -ne 'n/a') {
            $hostNames = $objs | ForEach-Object {
                $loc = Get-SafeValue { $_.Location }
                if ($loc -ne 'n/a') { ($loc -split '\\')[0] } else { 'n/a' }
            } | Where-Object { $_ -ne 'n/a' } | Select-Object -Unique
        }

        if ($Standard.OneJobPerHyperVHost -and $hostNames.Count -gt 1) {
            $flags.Add("Scope spans $($hostNames.Count) hosts; standard is one job per Hyper-V host")
        }

        if ($retainCycles -ne 'n/a' -and [int]::TryParse($retainCycles, [ref]$null) -and [int]$retainCycles -ne $Standard.OnPremDailyRetentionCycles) {
            $flags.Add("Retention is $retainCycles restore points, standard is $($Standard.OnPremDailyRetentionCycles)")
        }

        # Repository/host collocation check
        if ($repoName -ne 'n/a' -and $hostNames.Count -gt 0) {
            $repoHost = Get-SafeValue {
                (Get-VBRBackupRepository -Name $repoName).Host.Name
            }
            if ($repoHost -ne 'n/a' -and ($hostNames -contains $repoHost)) {
                $flags.Add("Repository '$repoName' appears to be hosted on a protected Hyper-V host ($repoHost)")
            }
        }

        $isEnabled = Get-SafeValue { $job.IsScheduleEnabled }

        # A disabled primary job means nothing in its scope is being
        # protected on a schedule. Worth surfacing explicitly rather than
        # leaving it to be spotted in the Enabled column.
        if ($isEnabled -eq $false) {
            $flags.Add('Job schedule is disabled; workloads in this job are not being backed up on a schedule')
        }

        $rows.Add([pscustomobject]@{
            Category   = 'On-Prem Backup Job'
            Name       = $job.Name
            Scope      = if ($hostNames.Count -gt 0) { $hostNames -join ', ' } else { 'n/a' }
            Repository = $repoName
            Retention  = $retentionDisplay
            StartTime  = $startTime
            Enabled    = $isEnabled
            Flags      = $flags
        })
    }

    # Stagger check across primary jobs sharing the same repository.
    # StartDateTimeLocal carries whatever date the schedule was originally
    # configured on, which differs per job, so comparing full datetimes
    # produces gaps of months instead of the real time-of-day collision.
    # Compare TimeOfDay only.
    # Disabled jobs are excluded because a job that never runs cannot
    # contend for repository I/O with the one next to it.
    $rows | Where-Object {
            $_.Category -eq 'On-Prem Backup Job' -and
            $_.Repository -ne 'n/a' -and
            $_.StartTime -ne 'n/a' -and
            $_.Enabled -ne $false
        } |
        Group-Object Repository | ForEach-Object {
            $repoGroupName = $_.Name
            $withTimeOfDay = $_.Group | ForEach-Object {
                try {
                    [pscustomobject]@{ Name = $_.Name; TimeOfDay = ([datetime]$_.StartTime).TimeOfDay }
                } catch { $null }
            } | Where-Object { $_ }
            $sorted = @($withTimeOfDay | Sort-Object TimeOfDay)
            for ($i = 1; $i -lt $sorted.Count; $i++) {
                $gap = ($sorted[$i].TimeOfDay - $sorted[$i-1].TimeOfDay).TotalMinutes
                if ($gap -ge 0 -and $gap -lt $Standard.JobStaggerMinutesMin) {
                    $globalFlags.Add("Jobs '$($sorted[$i-1].Name)' and '$($sorted[$i].Name)' on repository '$repoGroupName' both start around $($sorted[$i-1].TimeOfDay.ToString('hh\:mm')), only $([math]::Round($gap,1)) minutes apart; standard is $($Standard.JobStaggerMinutesMin) to $($Standard.JobStaggerMinutesMax) minutes")
                }
            }
        }

    # --- Backup copy jobs -----------------------------------------------------
    # Reuse the set already pulled for dedupe above instead of querying twice.
    $copyJobs = $copyJobsForDedupe

    foreach ($bc in $copyJobs) {
        $flags = [System.Collections.Generic.List[string]]::new()

        $mode = Get-FirstSafeValue @(
            { $bc.Mode },
            { $bc.ScheduleOptions.Type },
            { $bc.Options.JobScheduleOptions.Type }
        )

        $targetRepo = Get-FirstSafeValue @(
            { $bc.TargetRepository.Name },
            { $bc.GetTargetRepository().Name }
        )

        # SourceRepository and BackupJob are both arrays on this object model,
        # not single objects, so .Name needs an index first or it silently
        # returns nothing. Confirmed via -Diagnose against a real copy job.
        $sourceRepo = Get-FirstSafeValue @(
            { $bc.SourceRepository[0].Name },
            { $bc.BackupJob[0].Name }
        )

        # Get-VBRJobOptions is explicitly unsupported for backup copy jobs on
        # this build (confirmed by the "no longer supported" warning on every
        # run) and calling it was producing that warning without returning
        # anything useful. StorageOptions is the real property container for
        # a copy job's storage settings; the exact nested encryption field
        # name is not yet confirmed, this is the most likely name based on
        # the confirmed sibling property on the primary job's
        # BackupStorageOptions. If this still shows 'n/a' in the report, run
        # $bc.StorageOptions | Get-Member against a real copy job and this
        # candidate list needs the correct name.
        $encEnabled = Get-FirstSafeValue @(
            { $bc.StorageOptions.StorageEncryptionEnabled },
            { $bc.StorageOptions.EncryptionEnabled }
        )

        if ($mode -ne 'n/a' -and $mode -ne $Standard.BackupCopyModeExpected) {
            $flags.Add("Copy mode is '$mode', standard is '$($Standard.BackupCopyModeExpected)'")
        }

        if ($Standard.BackupCopyEncryptionRequired -and $encEnabled -eq $false) {
            $flags.Add('Encryption is disabled on this backup copy job')
        }

        if ($targetRepo -ne 'n/a') {
            $looksCloud = $false
            foreach ($hint in $Standard.CloudRepoNameHints) {
                if ($targetRepo -like "*$hint*") { $looksCloud = $true; break }
            }
            if (-not $looksCloud) {
                $flags.Add("Target repository '$targetRepo' does not match expected offsite/Data Cloud naming; confirm this job actually leaves the site")
            }
        }

        $retentionDisplay = Get-FirstSafeValue @(
            { "$($bc.RetentionNumber) $($bc.RetentionType)" }
        )

        $rows.Add([pscustomobject]@{
            Category   = 'Backup Copy Job'
            Name       = $bc.Name
            Scope      = $sourceRepo
            Repository = $targetRepo
            Retention  = $retentionDisplay
            StartTime  = $mode
            Enabled    = Get-SafeValue { $bc.JobEnabled }
            Flags      = $flags
        })
    }

    # Missing offsite copy check: any source repository with a primary job
    # but no backup copy job that appears linked to it.
    #
    # Primary jobs whose own target already looks like a cloud/offsite
    # repository are excluded from the missing-copy test, because those
    # write straight to the offsite target in one hop and were never meant
    # to have a separate copy job. Flagging them as "no backup copy job"
    # was the wrong framing. The real question about that design is the
    # opposite one (no local restore point), raised as a separate note
    # below rather than as a missing-copy flag.
    $sourceRepos = @()
    $singleHopRepos = @()
    foreach ($r in ($rows | Where-Object { $_.Category -eq 'On-Prem Backup Job' -and $_.Repository -ne 'n/a' })) {
        $alreadyCloud = $false
        foreach ($hint in $Standard.CloudRepoNameHints) {
            if ($r.Repository -like "*$hint*") { $alreadyCloud = $true; break }
        }
        if ($alreadyCloud) { $singleHopRepos += $r.Repository }
        else { $sourceRepos += $r.Repository }
    }
    $sourceRepos = $sourceRepos | Select-Object -Unique
    $singleHopRepos = $singleHopRepos | Select-Object -Unique

    foreach ($shr in $singleHopRepos) {
        $shJobs = ($rows | Where-Object { $_.Category -eq 'On-Prem Backup Job' -and $_.Repository -eq $shr } | Select-Object -ExpandProperty Name) -join ', '
        $globalFlags.Add("Repository '$shr' looks like an offsite/object storage target and is being used as the PRIMARY backup target ($shJobs). This is a single-hop design with no local restore point, so restores come back over the WAN. Confirm this is intentional rather than a missed local tier.")
    }

    $sourceJobNames   = @($rows | Where-Object { $_.Category -eq 'On-Prem Backup Job' } | Select-Object -ExpandProperty Name -Unique)
    $copiedIdentifiers = @($rows | Where-Object { $_.Category -eq 'Backup Copy Job' } | Select-Object -ExpandProperty Scope -Unique)

    if ($copiedIdentifiers -contains 'n/a' -or $copiedIdentifiers.Count -eq 0) {
        $globalFlags.Add("Backup copy job source could not be resolved on this run (property path still unconfirmed), so the missing-offsite-copy check was skipped rather than risk false flags. Run -Diagnose against a copy job and share the Get-VBRBackupCopyJob member list to fix this.")
    } else {
        foreach ($sr in $sourceRepos) {
            if ($sr -ne 'n/a' -and ($copiedIdentifiers -notcontains $sr) -and ($sourceJobNames -notcontains $sr)) {
                $siteJobsOnThisRepo = $rows | Where-Object { $_.Category -eq 'On-Prem Backup Job' -and $_.Repository -eq $sr } | Select-Object -ExpandProperty Name
                $matchedByJobName = $false
                foreach ($jn in $siteJobsOnThisRepo) {
                    if ($copiedIdentifiers -contains $jn) { $matchedByJobName = $true; break }
                }
                if (-not $matchedByJobName) {
                    $globalFlags.Add("Repository '$sr' has a primary backup job but no backup copy job found linked to it offsite")
                }
            }
        }
    }

    # The documented standard is expressed in restore points. If jobs are
    # actually configured with day-based retention, the numeric comparison
    # above is close but not exact, so say so once rather than per job.
    if ($script:AnyDayBasedRetention) {
        $globalFlags.Add("One or more jobs use day-based retention (RetentionType = Days) while the documented standard in runbook Section 8.2 is expressed as $($Standard.OnPremDailyRetentionCycles) daily restore points. The numbers are compared directly, so a job showing '$($Standard.OnPremDailyRetentionCycles) days' passes the check but is not literally the documented setting. Decide which unit the standard should use and align the runbook or the jobs.")
    }

    [pscustomobject]@{
        Rows        = $rows
        GlobalFlags = $globalFlags
    }
}

# ---------------------------------------------------------------------------
# Platform 2: Veeam Backup for Azure (VBA) audit, via REST API
# ---------------------------------------------------------------------------
function Get-VbaAudit {
    param($ApplianceUrl, $Cred, $Standard)

    if (-not $ApplianceUrl) { return $null }

    $rows = [System.Collections.Generic.List[object]]::new()
    $globalFlags = [System.Collections.Generic.List[string]]::new()

    try {
        $tokenBody = @{
            grant_type = 'Password'
            username   = $Cred.UserName
            password   = $Cred.GetNetworkCredential().Password
        }
        $tokenResp = Invoke-RestMethod -Method Post -Uri "$ApplianceUrl/api/oauth2/token" -Body $tokenBody -ContentType 'application/x-www-form-urlencoded' -ErrorAction Stop
        $accessToken = $tokenResp.access_token
    } catch {
        $globalFlags.Add("Could not authenticate to the VBA REST API at $ApplianceUrl. $($_.Exception.Message). VBA policies were not included in this report.")
        return [pscustomobject]@{ Rows = $rows; GlobalFlags = $globalFlags }
    }

    $headers = @{ Authorization = "Bearer $accessToken" }

    try {
        $policies = Invoke-RestMethod -Method Get -Uri "$ApplianceUrl/api/v8/policies" -Headers $headers -ErrorAction Stop
    } catch {
        $globalFlags.Add("Authenticated to VBA but could not read /api/v8/policies. $($_.Exception.Message). If this appliance is on an older or newer API version, adjust the version segment in the URL and re-run.")
        return [pscustomobject]@{ Rows = $rows; GlobalFlags = $globalFlags }
    }

    foreach ($p in $policies) {
        $flags = [System.Collections.Generic.List[string]]::new()

        $name        = Get-SafeValue { $p.name }
        $appAware    = Get-SafeValue { $p.snapshotSettings.applicationAwareSnapshot.isEnabled }
        $dailyDays   = Get-SafeValue { $p.backupSettings.dailyBackup.retentionSettings.retentionDays }
        $dailyImm    = Get-SafeValue { $p.backupSettings.dailyBackup.immutabilitySettings.isEnabled }
        $weeklyImm   = Get-SafeValue { $p.backupSettings.weeklyBackup.immutabilitySettings.isEnabled }
        $monthlyImm  = Get-SafeValue { $p.backupSettings.monthlyBackup.immutabilitySettings.isEnabled }

        if ($Standard.VbaAppAwareRequired -and $appAware -eq $false) {
            $flags.Add('Application-aware processing is disabled on this policy')
        }
        if ($dailyDays -ne 'n/a' -and [int]::TryParse($dailyDays, [ref]$null) -and [int]$dailyDays -ne $Standard.VbaDailyRetentionDays) {
            $flags.Add("Daily retention is $dailyDays days, standard is $($Standard.VbaDailyRetentionDays)")
        }
        if ($Standard.VbaImmutabilityRequired) {
            foreach ($tier in @(@{n='daily';v=$dailyImm}, @{n='weekly';v=$weeklyImm}, @{n='monthly';v=$monthlyImm})) {
                if ($tier.v -eq $false) {
                    $flags.Add("Immutability is disabled on the $($tier.n) tier")
                }
            }
        }

        $rows.Add([pscustomobject]@{
            Category   = 'VBA Azure Policy'
            Name       = $name
            Scope      = Get-SafeValue { ($p.virtualMachinesSettings.includedVirtualMachines | ForEach-Object { $_.name }) -join ', ' }
            Repository = Get-SafeValue { $p.backupSettings.dailyBackup.backupRepositoryName }
            Retention  = "Daily $dailyDays d / weekly-monthly per policy"
            StartTime  = Get-SafeValue { $p.snapshotSettings.dailySnapshot.dailyScheduleSettings.time }
            Enabled    = Get-SafeValue { $p.isEnabled }
            Flags      = $flags
        })
    }

    [pscustomobject]@{ Rows = $rows; GlobalFlags = $globalFlags }
}

# ---------------------------------------------------------------------------
# Platform 3: Veeam Data Cloud for M365, manual input file
# ---------------------------------------------------------------------------
function Get-DataCloudAudit {
    param($InputPath)

    $rows = [System.Collections.Generic.List[object]]::new()
    $globalFlags = [System.Collections.Generic.List[string]]::new()

    if (-not (Test-Path $InputPath)) {
        $globalFlags.Add("No Data Cloud input file found at $InputPath. Run with -SeedDataCloudTemplate to create one, fill it in from the Data Cloud console, then re-run.")
        return [pscustomobject]@{ Rows = $rows; GlobalFlags = $globalFlags }
    }

    $policies = Get-Content $InputPath -Raw | ConvertFrom-Json

    foreach ($p in $policies) {
        $flags = [System.Collections.Generic.List[string]]::new()
        if ($p.Status -eq 'INACTIVE') {
            $flags.Add('Policy is marked inactive; confirm it is still needed or remove it during cleanup')
        }
        $rows.Add([pscustomobject]@{
            Category   = 'Data Cloud M365 Policy (manual)'
            Name       = $p.PolicyName
            Scope      = $p.Tenant
            Repository = $p.Repository
            Retention  = $p.RetentionSummary
            StartTime  = "Verified $($p.LastVerified)"
            Enabled    = $p.Status
            Flags      = $flags
        })
    }

    [pscustomobject]@{ Rows = $rows; GlobalFlags = $globalFlags }
}

# ---------------------------------------------------------------------------
# HTML report builder
# ---------------------------------------------------------------------------
function ConvertTo-HtmlReport {
    param($OnPrem, $Vba, $DataCloud, $Standard, $OutputPath)

    $allRows = @()
    if ($OnPrem)    { $allRows += $OnPrem.Rows }
    if ($Vba)       { $allRows += $Vba.Rows }
    if ($DataCloud) { $allRows += $DataCloud.Rows }

    $totalJobs   = $allRows.Count
    $flaggedJobs = ($allRows | Where-Object { $_.Flags.Count -gt 0 }).Count
    $totalFlags  = ($allRows | ForEach-Object { $_.Flags.Count } | Measure-Object -Sum).Sum
    $globalFlags = @()
    if ($OnPrem)    { $globalFlags += $OnPrem.GlobalFlags }
    if ($Vba)       { $globalFlags += $Vba.GlobalFlags }
    if ($DataCloud) { $globalFlags += $DataCloud.GlobalFlags }

    function Get-RowHtml {
        param($row)
        $badgeClass = if ($row.Flags.Count -gt 0) { 'flag' } else { 'pass' }
        $badgeText  = if ($row.Flags.Count -gt 0) { "$($row.Flags.Count) flag(s)" } else { 'OK' }
        $flagList   = if ($row.Flags.Count -gt 0) {
            '<ul class="flaglist">' + (($row.Flags | ForEach-Object { "<li>$_</li>" }) -join '') + '</ul>'
        } else { '<span class="muted">None</span>' }

        @"
<tr>
  <td>$($row.Name)</td>
  <td>$($row.Scope)</td>
  <td>$($row.Repository)</td>
  <td>$($row.Retention)</td>
  <td>$($row.StartTime)</td>
  <td>$($row.Enabled)</td>
  <td><span class="badge $badgeClass">$badgeText</span></td>
  <td>$flagList</td>
</tr>
"@
    }

    function Get-SectionHtml {
        param($title, $rows)
        if (-not $rows -or $rows.Count -eq 0) {
            return "<h2>$title</h2><p class='muted'>No data collected for this platform in this run.</p>"
        }
        $rowsHtml = ($rows | ForEach-Object { Get-RowHtml $_ }) -join "`n"
        @"
<h2>$title</h2>
<table>
  <thead>
    <tr><th>Name</th><th>Scope</th><th>Repository</th><th>Retention</th><th>Schedule / Mode</th><th>Enabled</th><th>Status</th><th>Flags</th></tr>
  </thead>
  <tbody>
    $rowsHtml
  </tbody>
</table>
"@
    }

    $onPremRows    = $allRows | Where-Object { $_.Category -in @('On-Prem Backup Job', 'Backup Copy Job') }
    $vbaRows       = $allRows | Where-Object { $_.Category -eq 'VBA Azure Policy' }
    $dataCloudRows = $allRows | Where-Object { $_.Category -eq 'Data Cloud M365 Policy (manual)' }

    $globalFlagsHtml = if ($globalFlags.Count -gt 0) {
        '<ul class="flaglist">' + (($globalFlags | ForEach-Object { "<li>$_</li>" }) -join '') + '</ul>'
    } else { '<p class="muted">None</p>' }

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>VBR Job Settings Audit</title>
<style>
  body { font-family: Calibri, Segoe UI, Arial, sans-serif; margin: 0; padding: 0; background: #f4f6f8; color: #1c2733; }
  header { background: #12294d; color: #fff; padding: 24px 32px; }
  header h1 { margin: 0; font-size: 22px; }
  header p { margin: 4px 0 0; color: #c7d2e0; font-size: 13px; }
  main { padding: 24px 32px 60px; }
  .summary { display: flex; gap: 16px; margin-bottom: 28px; flex-wrap: wrap; }
  .card { background: #fff; border: 1px solid #dbe1e8; border-radius: 6px; padding: 16px 20px; min-width: 160px; }
  .card .num { font-size: 26px; font-weight: bold; color: #12294d; }
  .card .label { font-size: 12px; color: #5a6b7d; text-transform: uppercase; letter-spacing: 0.03em; }
  h2 { color: #12294d; border-bottom: 2px solid #12294d; padding-bottom: 6px; margin-top: 36px; }
  table { width: 100%; border-collapse: collapse; background: #fff; margin-top: 10px; }
  th, td { text-align: left; padding: 8px 10px; border-bottom: 1px solid #e4e9ee; font-size: 13px; vertical-align: top; }
  th { background: #eaeff5; color: #12294d; }
  tr:hover { background: #f7f9fb; }
  .badge { display: inline-block; padding: 2px 8px; border-radius: 10px; font-size: 12px; font-weight: bold; }
  .badge.pass { background: #dff3e3; color: #1e7a35; }
  .badge.flag { background: #fbe6d5; color: #a8500f; }
  .flaglist { margin: 0; padding-left: 18px; }
  .flaglist li { margin-bottom: 4px; }
  .muted { color: #7c8a99; font-size: 13px; }
  .standard-box { background: #fff; border: 1px solid #dbe1e8; border-radius: 6px; padding: 16px 20px; margin-top: 10px; font-size: 13px; }
  .standard-box code { background: #eaeff5; padding: 1px 5px; border-radius: 3px; }
</style>
</head>
<body>
<header>
  <h1>Veeam Job Settings Audit and Consistency Report</h1>
  <p>Generated $(Get-Date -Format 'yyyy-MM-dd HH:mm') against VBR server $VbrServer</p>
</header>
<main>
  <div class="summary">
    <div class="card"><div class="num">$totalJobs</div><div class="label">Jobs and policies reviewed</div></div>
    <div class="card"><div class="num">$flaggedJobs</div><div class="label">Items with flags</div></div>
    <div class="card"><div class="num">$totalFlags</div><div class="label">Total flags</div></div>
    <div class="card"><div class="num">$($globalFlags.Count)</div><div class="label">Environment level flags</div></div>
  </div>

  <h2>Environment level flags</h2>
  $globalFlagsHtml

  $(Get-SectionHtml -title "On-Prem VBR: Backup and Backup Copy Jobs" -rows $onPremRows)

  $(Get-SectionHtml -title "Veeam Backup for Azure (VBA) Policies" -rows $vbaRows)

  $(Get-SectionHtml -title "Veeam Data Cloud for M365 Policies (manual entry)" -rows $dataCloudRows)

  <h2>Standard referenced for this audit</h2>
  <div class="standard-box">
    <p>On-prem: <code>$($Standard.OnPremDailyRetentionCycles) daily restore points</code>, one job per Hyper-V host, jobs on the same repository staggered <code>$($Standard.JobStaggerMinutesMin)-$($Standard.JobStaggerMinutesMax) minutes</code> apart, one backup copy job per site in <code>$($Standard.BackupCopyModeExpected)</code> mode with encryption enabled.</p>
    <p>VBA: application-aware processing enabled, daily <code>$($Standard.VbaDailyRetentionDays) days</code> immutable on the hot tier, weekly <code>$($Standard.VbaWeeklyRetentionMonths) month</code> immutable on the cool tier, monthly <code>$($Standard.VbaMonthlyRetentionMonths) months</code> immutable on the archive tier.</p>
    <p>Source: Forgent Power Operations Runbook, Section 8.2 and Section 1.5.</p>
  </div>
</main>
</body>
</html>
"@

    $html | Out-File -FilePath $OutputPath -Encoding UTF8
    return $OutputPath
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
Write-Host "Collecting on-prem VBR job settings..." -ForegroundColor Cyan
$onPremResult = Get-OnPremAudit -Standard $Standard

$vbaResult = $null
if ($VbaApplianceUrl) {
    Write-Host "Collecting VBA policy settings..." -ForegroundColor Cyan
    if (-not $VbaCredential) {
        Write-Host "VbaApplianceUrl was supplied without VbaCredential; skipping VBA section." -ForegroundColor Yellow
    } else {
        $vbaResult = Get-VbaAudit -ApplianceUrl $VbaApplianceUrl -Cred $VbaCredential -Standard $Standard
    }
} else {
    Write-Host "VbaApplianceUrl not supplied; skipping VBA section." -ForegroundColor Yellow
}

Write-Host "Reading Data Cloud M365 policy file..." -ForegroundColor Cyan
$dataCloudResult = Get-DataCloudAudit -InputPath $DataCloudInputPath

Write-Host "Building HTML report..." -ForegroundColor Cyan
$reportPath = ConvertTo-HtmlReport -OnPrem $onPremResult -Vba $vbaResult -DataCloud $dataCloudResult -Standard $Standard -OutputPath $OutputPath

Write-Host "`nReport written to $reportPath" -ForegroundColor Green
