#Requires -Version 5.1
<#
.SYNOPSIS
    Collects the Windows devices in an Entra/Intune group and emits enriched device records
    (Intune managed-device ID, serial number, enrollment profile, etc.) as JSON.

.DESCRIPTION
    Reads the (transitive) device members of an Entra group, resolves each to its Intune managed
    device in batches (Graph $batch, 20 per call), and outputs an enriched JSON array to STDOUT.
    That output is designed to be piped straight into Rename-IntuneDevice.ps1 -Stdin, so the rename
    step needs no further Graph reads.

    All logs go to STDERR; only the JSON result goes to STDOUT, so the two chain cleanly in n8n
    (Execute Command) or a shell pipe:

        pwsh -File Get-IntuneGroupDevices.ps1 -GroupId <guid> | pwsh -File Rename-IntuneDevice.ps1 -Stdin

    Requires IntuneGraphCommon.ps1 in the same folder (dot-sourced).

.PARAMETER GroupId
    Entra group object ID (GUID) whose device members to collect.

.PARAMETER Transitive
    Include devices from nested groups (default true). Use -Transitive:$false for direct members only.

.PARAMETER IncludeUnenrolled
    Also emit group devices that have no Intune managed-device record (IntuneDeviceId = null,
    Status = NotEnrolledInIntune). By default those are logged and omitted from the output.

.PARAMETER OutputPath
    Optional file to also write. Extension decides the format: .csv writes a CSV, anything else
    (or .json) writes the JSON array. STDOUT always receives the JSON array regardless.

.PARAMETER AuthMode
    ClientSecret (default, app-only) or Interactive (browser sign-in as your admin). See README.

.PARAMETER TenantId / ClientId / ClientSecret
    Credentials. Fall back to RENAMEDEVICE_TENANT_ID / _CLIENT_ID / _CLIENT_SECRET env vars.
    For -AuthMode Interactive, ClientId is optional (defaults to the Graph public client).

.PARAMETER GraphBaseUri
    Graph endpoint base. Default 'https://graph.microsoft.com/beta'.

.PARAMETER MaxRetries
    Max retry attempts for transient Graph failures (429/5xx/network). Default 5.

.PARAMETER LogDirectory
    Directory for the run log. Default: <script folder>\logs.

.EXAMPLE
    .\Get-IntuneGroupDevices.ps1 -GroupId <guid> -AuthMode Interactive -TenantId <tenant> -OutputPath .\devices.json

.EXAMPLE
    pwsh -File .\Get-IntuneGroupDevices.ps1 -GroupId <guid> | pwsh -File .\Rename-IntuneDevice.ps1 -Stdin -Rename -Force

.NOTES
    Required Microsoft Graph permissions (Application for ClientSecret; delegated for Interactive):
      * GroupMember.Read.All                       (read the group's members)
      * Device.Read.All                            (read the Entra device objects)
      * DeviceManagementManagedDevices.Read.All    (resolve Intune managed devices)

    Exit codes: 0 = at least one device resolved; 1 = group had members but none resolved to Intune;
                2 = fatal (bad arguments, auth failure, group not found).
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string] $GroupId,

    [bool] $Transitive = $true,

    [switch] $IncludeUnenrolled,

    [string] $OutputPath,

    [ValidateSet('ClientSecret', 'Interactive')]
    [string] $AuthMode = 'ClientSecret',

    [string] $TenantId,
    [string] $ClientId,
    [string] $ClientSecret,

    [string] $GraphBaseUri = 'https://graph.microsoft.com/beta',

    [ValidateRange(0, 10)]
    [int] $MaxRetries = 5,

    [string] $LogDirectory
)

Set-StrictMode -Version 1.0
$ErrorActionPreference = 'Stop'

$ExitSuccess = 0; $ExitNone = 1; $ExitFatal = 2

$scriptRoot = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
. (Join-Path $scriptRoot 'IntuneGraphCommon.ps1')

$runStamp = (Get-Date).ToString('yyyyMMdd_HHmmss')
if (-not $LogDirectory) { $LogDirectory = Join-Path $scriptRoot 'logs' }
if (-not (Test-Path -LiteralPath $LogDirectory)) { New-Item -ItemType Directory -Path $LogDirectory -Force | Out-Null }
$script:GraphLogFile     = Join-Path $LogDirectory ("GetIntuneGroupDevices_{0}.log" -f $runStamp)
$script:GraphLogToStderr = $true   # keep STDOUT clean for the JSON result

try {
    Write-Log -Level INFO -Message '=========================================================='
    Write-Log -Level INFO -Message 'Get-IntuneGroupDevices starting.'
    Write-Log -Level INFO -Message ("PowerShell {0} ({1})" -f $PSVersionTable.PSVersion, $PSVersionTable.PSEdition)

    $guidRef = [ref]([guid]::Empty)
    if (-not [guid]::TryParse($GroupId, $guidRef)) {
        Write-Log -Level ERROR -Message ("GroupId '{0}' is not a valid GUID." -f $GroupId)
        exit $ExitFatal
    }

    Connect-Graph -AuthMode $AuthMode -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret `
        -Scopes @('GroupMember.Read.All', 'Device.Read.All', 'DeviceManagementManagedDevices.Read.All') `
        -MaxRetries $MaxRetries -GraphBaseUri $GraphBaseUri

    # 1) Group device members (transitive by default).
    $memberSeg = if ($Transitive) { 'transitiveMembers' } else { 'members' }
    $membersUri = "$script:GraphBaseUri/groups/$GroupId/$memberSeg/microsoft.graph.device" + '?$select=id,deviceId,displayName,operatingSystem&$top=999'
    Write-Log -Level INFO -Message ("Reading {0} device members of group {1}..." -f $memberSeg, $GroupId)

    $entraDevices = @(Get-GraphCollection -Uri $membersUri)
    Write-Log -Level INFO -Message ("Group has {0} device member(s)." -f $entraDevices.Count)
    if ($entraDevices.Count -eq 0) {
        Write-Log -Level WARN -Message 'No device members found. Emitting empty array.'
        Write-Output '[]'
        exit $ExitNone
    }

    # 2) Resolve each Entra device (deviceId = azureADDeviceId) to its Intune managed device, in batches.
    $mdSelect = 'id,deviceName,serialNumber,operatingSystem,azureADDeviceId,joinType,enrollmentProfileName,lastSyncDateTime'
    $requests = New-Object System.Collections.Generic.List[object]
    $indexMap = @{}
    $n = 0
    foreach ($d in $entraDevices) {
        $aad = ("$($d.deviceId)").Trim()
        if ([string]::IsNullOrWhiteSpace($aad)) { continue }
        $rid = [string]$n
        $url = "/deviceManagement/managedDevices?`$filter=azureADDeviceId eq '$aad'&`$select=$mdSelect"
        $requests.Add(@{ id = $rid; method = 'GET'; url = $url })
        $indexMap[$rid] = $d
        $n++
    }

    Write-Log -Level INFO -Message ("Resolving {0} device(s) to Intune managed devices ({1} batch call(s))..." -f $requests.Count, [Math]::Ceiling($requests.Count / 20))
    $batchResults = Invoke-GraphBatch -Requests $requests.ToArray()

    # 3) Build enriched records.
    $records = New-Object System.Collections.Generic.List[object]
    $resolved = 0; $unenrolled = 0; $failed = 0
    foreach ($rid in $indexMap.Keys) {
        $entra = $indexMap[$rid]
        $res = $batchResults[$rid]
        $md = $null
        if ($res -and $res.status -eq 200 -and $res.body -and $res.body.value) {
            $vals = @($res.body.value)
            if ($vals.Count -gt 1) {
                Write-Log -Level WARN -Message ("Device '{0}' has {1} Intune records; using the most recently synced." -f $entra.displayName, $vals.Count)
                $md = $vals | Sort-Object { $_.lastSyncDateTime } -Descending | Select-Object -First 1
            } else { $md = $vals[0] }
        }
        elseif ($res -and $res.status -ne 200 -and $res.status -ne 404) {
            $failed++
            Write-Log -Level WARN -Message ("Lookup for '{0}' returned status {1}." -f $entra.displayName, $res.status)
        }

        if ($null -eq $md) {
            $unenrolled++
            if ($IncludeUnenrolled) {
                $records.Add([pscustomobject][ordered]@{
                    IntuneDeviceId = $null; AzureADDeviceId = $entra.deviceId; EntraObjectId = $entra.id
                    DeviceName = $entra.displayName; SerialNumber = $null; OperatingSystem = $entra.operatingSystem
                    JoinType = $null; EnrollmentProfileName = $null; Status = 'NotEnrolledInIntune'
                })
            }
            continue
        }

        $resolved++
        $records.Add([pscustomobject][ordered]@{
            IntuneDeviceId        = $md.id
            AzureADDeviceId       = $md.azureADDeviceId
            EntraObjectId         = $entra.id
            DeviceName            = $md.deviceName
            SerialNumber          = ("$($md.serialNumber)").Trim()
            OperatingSystem       = $md.operatingSystem
            JoinType              = $md.joinType
            EnrollmentProfileName = $md.enrollmentProfileName
            Status                = 'Resolved'
        })
    }

    Write-Log -Level INFO -Message ("Resolved {0}, not enrolled {1}, lookup errors {2}." -f $resolved, $unenrolled, $failed)

    # 4) Optional file output.
    if ($OutputPath) {
        try {
            $dir = Split-Path -Parent $OutputPath
            if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
            if ($OutputPath.ToLowerInvariant().EndsWith('.csv')) {
                $records | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding UTF8
            } else {
                $fileJson = if ($records.Count -eq 0) { '[]' } else { ConvertTo-Json -InputObject ([object[]]$records) -Depth 6 }
                Set-Content -LiteralPath $OutputPath -Value $fileJson -Encoding UTF8
            }
            Write-Log -Level SUCCESS -Message ("Wrote {0} record(s) to {1}" -f $records.Count, $OutputPath)
        } catch { Write-Log -Level WARN -Message ("Failed to write output file: {0}" -f $_.Exception.Message) }
    }

    # 5) Emit JSON array to STDOUT (always an array, even for a single record).
    $stdoutJson = if ($records.Count -eq 0) { '[]' } else { ConvertTo-Json -InputObject ([object[]]$records) -Depth 6 }
    Write-Output $stdoutJson

    if ($resolved -eq 0) { exit $ExitNone }
    exit $ExitSuccess
}
catch {
    Write-Log -Level ERROR -Message ("FATAL: {0}" -f $_.Exception.Message)
    Write-Log -Level DEBUG -Message ($_.ScriptStackTrace)
    exit $ExitFatal
}
finally { Clear-GraphState }
