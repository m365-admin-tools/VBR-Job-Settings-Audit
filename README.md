# VBR-Job-Settings-Audit
Veeam Job Settings Audit (PowerShell)

Please visit https://m365admintools.com for more tools and information.

Audits every Veeam job against a documented job design standard and produces a single HTML dashboard showing each job's settings with any deviation flagged and explained in plain language.

This is not a generic best-practice checker. The standard lives in one editable block at the top of the script, so it compares your environment against the settings your own build sheet or runbook says it should have.

It covers three platforms in one report:

- **Veeam Backup & Replication.** Backup jobs and backup copy jobs read live from the VBR PowerShell module.
- **Veeam Backup for Azure.** Policies read live from the appliance REST API.
- **Veeam Data Cloud for Microsoft 365.** Read from a JSON file you maintain by hand, because there is no customer-facing API for pulling those policy settings. This section is labelled "manual" in the report so it is never mistaken for a live check.

Built for MSPs and backup administrators who have a documented standard and need to prove, on demand, which jobs still match it.

<!-- Add a screenshot of the HTML dashboard here, then uncomment:
![Audit dashboard](docs/images/audit-dashboard.png)
-->

## Requirements

| Item | Requirement |
|---|---|
| PowerShell | 5.1 or later |
| Veeam Backup & Replication | v12 or newer. Requires the `Veeam.Backup.PowerShell` module, which ships with the VBR console |
| Where to run it | On the VBR server or a machine with the VBR console installed. For a workgroup VBR server, run it locally rather than remotely |
| Veeam Backup for Azure | Optional. Needs the appliance REST API URL and an account with rights to read policies |
| Veeam Data Cloud for M365 | Optional. Needs a JSON file you maintain. Generate a starter file with `-SeedDataCloudTemplate` |

## Quick start

```powershell
# Confirm the property names on your VBR build before trusting the flags
.\VBR-Job-Settings-Audit.ps1 -Diagnose

# On-premises jobs only
.\VBR-Job-Settings-Audit.ps1

# Include Veeam Backup for Azure policies
.\VBR-Job-Settings-Audit.ps1 -VbaApplianceUrl https://10.10.16.10:4443 -VbaCredential (Get-Credential)

# Create the starter Data Cloud policy file, then edit it
.\VBR-Job-Settings-Audit.ps1 -SeedDataCloudTemplate
```

Run `-Diagnose` first. Veeam's PowerShell object model has changed property names between releases. Diagnose mode pulls one job, dumps every property name from `Get-VBRJobOptions` and `Get-VBRJobScheduleOptions`, and lets you confirm the paths the script reads are still current on your build.

## Parameters

| Parameter | Type | Default | Description |
|---|---|---|---|
| `-VbrServer` | string | `localhost` | VBR server to connect to |
| `-VbrCredential` | PSCredential | Current user | Credentials for a remote VBR connection. Not needed locally |
| `-VbaApplianceUrl` | string | None | Base HTTPS URL of the Veeam Backup for Azure REST API, for example `https://10.10.16.10:4443`. Omit to skip the section |
| `-VbaCredential` | PSCredential | None | Account with rights to read VBA policies. Required when the URL is supplied |
| `-DataCloudInputPath` | string | `.\DataCloud-M365-Policies.json` | Path to the Data Cloud policy JSON file |
| `-SeedDataCloudTemplate` | switch | Off | Writes a starter JSON file and exits without auditing |
| `-OutputPath` | string | Timestamped file in the script folder | Path for the HTML report |
| `-Diagnose` | switch | Off | Dumps the Veeam job option property names and exits |
| `-ShowVeeamWarnings` | switch | Off | Restores Veeam's own warning output, which is suppressed during collection by default |

## Defining your standard

Every threshold is in one ordered hashtable named `$Standard` near the top of the script. Edit it there rather than in the check logic.

| Setting | Default | Meaning |
|---|---|---|
| `OnPremDailyRetentionCycles` | 14 | Expected restore point count on on-premises backup jobs |
| `OnPremWeeklyGfsCount` | 4 | Expected weekly GFS count |
| `OnPremMonthlyGfsCount` | 3 | Expected monthly GFS count |
| `JobStaggerMinutesMin` / `Max` | 10 / 15 | Expected gap between jobs sharing a repository |
| `BackupCopyModeExpected` | `Immediate` | Expected backup copy mode |
| `BackupCopyEncryptionRequired` | true | Backup copy jobs must have encryption enabled |
| `OneJobPerHyperVHost` | true | A job should not span multiple Hyper-V hosts |
| `VbaDailyRetentionDays` | 14 | Expected daily retention on Azure policies |
| `VbaWeeklyRetentionMonths` | 1 | Expected weekly retention |
| `VbaMonthlyRetentionMonths` | 12 | Expected monthly retention |
| `VbaImmutabilityRequired` | true | Immutability must be enabled on every tier |
| `VbaAppAwareRequired` | true | Application-aware processing must be enabled |
| `CloudRepoNameHints` | list of strings | Wildcard matches used to recognize an offsite or cloud target by repository name |

`CloudRepoNameHints` is the one to review first. The script decides whether a backup copy job actually leaves the site by matching the target repository name against this list, so add your own naming conventions or the check will flag jobs that are in fact correct.

## What it checks

**On-premises backup jobs**

- Job scope spans more than one Hyper-V host
- Retention does not match the standard restore point count
- Target repository appears to sit on a Hyper-V host that the same job protects
- Job schedule is disabled, so its workloads are not being backed up on a schedule

**On-premises backup copy jobs**

- Copy mode does not match the standard
- Encryption is disabled
- Target repository name does not match any known offsite or cloud naming pattern, so the job may not be leaving the site

**Veeam Backup for Azure policies**

- Application-aware processing is disabled
- Daily retention does not match the standard
- Immutability is disabled on the daily, weekly, or monthly tier

**Veeam Data Cloud for M365 policies**

- Policy is marked inactive, so it should be confirmed as still needed or removed

Each job appears in the report with its scope, repository, retention, start time, and schedule state, plus an OK badge or a count of flags with each deviation written out.

## Output

A single self-contained HTML file, timestamped, for example `VBR-Job-Audit-20260917-1422.html`. It contains:

- Totals for jobs audited, jobs flagged, and flags raised
- One section per platform, with the Data Cloud section clearly marked as manual entry
- Global flags for anything that prevented collection, such as a failed VBA authentication
- A closing block restating the standard the audit was run against, so the report is readable a year later without the script

## The Data Cloud JSON file

Run `-SeedDataCloudTemplate` once to generate a starter file, then edit it. Each entry holds the repository identifier, policy name, location, status, tenant, a retention summary, and the date you last verified it in the console. Update it whenever a Data Cloud policy changes. The `LastVerified` date is carried into the report so the reader knows how current the manual section is.

## What it changes

Nothing. The script reads job configuration, reads the VBA REST API, reads a local JSON file, and writes one HTML file. No job is created, modified, enabled, disabled, or started.

## Known limitations

- **`n/a` means "not found on this build", not "compliant".** Every property read is wrapped so a renamed property shows `n/a` rather than stopping the script. A report full of `n/a` values means the property paths need updating for your build. Run `-Diagnose` and correct them.
- **Day-based retention.** Jobs using day-based rather than cycle-based retention are noted in the report, because the two are not directly comparable to a restore point count.
- **VBA API version.** The policy call uses the `/api/v8/` path. On an older or newer appliance, change the version segment and re-run. The script reports the failure rather than silently omitting the section.
- **Repository name matching is a wildcard contains check.** It is a naming convention test, not proof that data is leaving the site. Confirm the target manually for anything it flags.
- **Remote connections to a workgroup VBR server** need WinRM TrustedHosts configuration. Running the script locally on the VBR console avoids the issue.
- **Configuration only.** The audit reports what the jobs are set to do, not whether they succeeded, and not whether a restore works.

## Troubleshooting

**`Could not load Veeam.Backup.PowerShell. Run this script on the VBR console server, or from a machine with the Veeam PowerShell module installed.`**

Install the VBR console on this machine, or run the script on the VBR server. Confirm with:

```powershell
Get-Module -ListAvailable Veeam.Backup.PowerShell
```

**`Could not authenticate to the VBA REST API at <url>`**

The appliance rejected the credentials or was unreachable. Check that port 4443 is open from this machine, that the account is a VBA local or SSO user with policy read rights, and that the URL has no trailing slash. This appears as a global flag in the report and the rest of the audit still completes.

**`Authenticated to VBA but could not read /api/v8/policies`**

The appliance is on a different API version. Change the version segment in the script and re-run.

**Most columns show `n/a`**

The property paths do not match your VBR build. Run with `-Diagnose`, compare the dumped property names against the ones used in `Get-OnPremAudit`, and update them.

**`This cmdlet is no longer supported for Backup Copy jobs`**

A Veeam warning raised during collection. It is suppressed by default. Use `-ShowVeeamWarnings` to see it and anything else Veeam emits, which is worth doing if a report looks wrong.

## Related

- Free Microsoft 365, Active Directory, and Veeam tools at [m365admintools.com](https://m365admintools.com)

## Author

Charles Arconi, [m365admintools.com](https://m365admintools.com)

Not affiliated with, endorsed by, or supported by Veeam Software. Veeam is a trademark of Veeam Software Group GmbH.

## License

MIT. See [LICENSE](LICENSE).
